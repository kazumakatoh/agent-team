/**
 * MoneyForward → 資金繰り表 データ同期
 *
 * エンドポイント:
 *   /reports/transition_pl?type=monthly  → PL月次推移（売上・経費）
 *   /reports/transition_bs?type=monthly  → BS月次推移（残高）
 *   /journals                            → 仕訳（売掛回収・買掛支払・借入返済）
 */

function sync2025() { syncFromMF(2025); }
function sync2026() { syncFromMF(2026); }

// ==============================
// 期間計算
// ==============================

function fiscalDates_(year) {
  return { startYear: year - 1, endYear: year };
}

function monthKeys_(year) {
  var fy = fiscalDates_(year);
  var keys = [];
  for (var m = 3; m <= 12; m++) keys.push(fy.startYear + '-' + padZero_(m));
  keys.push(year + '-01');
  keys.push(year + '-02');
  return keys;
}

function padZero_(n) { return n < 10 ? '0' + n : '' + n; }

// ==============================
// 推移表データ取得・パース
// ==============================

/**
 * 推移表PL/BSを取得
 * columns: ["3","4","5",..."1","2","settlement_balance","total"] or similar
 * rows: nested { name, type, values[], rows[] }
 */
function fetchTransition_(type, fiscalYear) {
  var endpoint = (type === 'pl') ? '/reports/transition_pl' : '/reports/transition_bs';
  return mfApiGet_(endpoint, {
    type: 'monthly',
    fiscal_year: fiscalYear
  });
}

/**
 * 推移表レスポンスから指定勘定科目の月次データを抽出
 * @param {Object} data - 推移表レスポンス
 * @param {string} accountName - 勘定科目名
 * @param {Array} monthKeys - ["2024-03", "2024-04", ...]
 * @return {Object} {"2024-03": 金額, ...}
 */
function extractTransitionAccount_(data, accountName, monthKeys) {
  var columns = data.columns || [];
  var result = {};

  // columnsは ["3","4","5",...,"1","2","settlement_balance","total"] 形式
  // monthKeysとの対応を構築
  var colToYM = {};
  var year = data.fiscal_year || 0;
  monthKeys.forEach(function(ym) {
    var month = parseInt(ym.split('-')[1]);
    var colIdx = columns.indexOf(String(month));
    if (colIdx !== -1) colToYM[colIdx] = ym;
  });

  // 再帰的にrowsを探索
  function search(rows) {
    if (!rows) return;
    rows.forEach(function(row) {
      if (row.name === accountName && row.type === 'account' && row.values) {
        row.values.forEach(function(val, idx) {
          if (colToYM[idx] !== undefined) {
            result[colToYM[idx]] = val || 0;
          }
        });
      }
      if (row.rows) search(row.rows);
    });
  }
  search(data.rows || []);
  return result;
}

/**
 * 推移表から複数勘定科目の合計を月次で取得
 */
function extractTransitionSum_(data, accountNames, monthKeys) {
  var totals = {};
  monthKeys.forEach(function(k) { totals[k] = 0; });
  accountNames.forEach(function(name) {
    var vals = extractTransitionAccount_(data, name, monthKeys);
    monthKeys.forEach(function(k) {
      totals[k] += (vals[k] || 0);
    });
  });
  return totals;
}

// ==============================
// 仕訳データ取得（売掛回収・買掛支払・借入返済用）
// ==============================

function fetchJournals_(year) {
  var fy = fiscalDates_(year);
  var startDate = fy.startYear + '-03-01';
  var endDate = fy.endYear + '-02-28';
  var allJournals = [];
  var page = 1;

  while (true) {
    var data = mfApiGet_('/journals', {
      start_date: startDate,
      end_date: endDate,
      page: page,
      per_page: 10000
    });
    var journals = data.journals || [];
    allJournals = allJournals.concat(journals);
    var totalPages = (data.metadata && data.metadata.total_pages) || 1;
    if (page >= totalPages) break;
    page++;
    if (page > 100) break;
  }
  Logger.log('仕訳取得: ' + allJournals.length + '件');
  return allJournals;
}

function getJournalYM_(j) {
  return (j.transaction_date || '').substring(0, 7);
}

/**
 * 仕訳から指定勘定科目の月次金額を集計
 */
function sumJournalByAccount_(journals, accountNames, side) {
  var monthly = {};
  journals.forEach(function(j) {
    var ym = getJournalYM_(j);
    if (!ym) return;
    (j.branches || []).forEach(function(b) {
      var entry = (side === 'debit') ? b.debitor : b.creditor;
      if (entry && accountNames.indexOf(entry.account_name) !== -1 && entry.value > 0) {
        monthly[ym] = (monthly[ym] || 0) + entry.value;
      }
    });
  });
  return monthly;
}

function calcARCollection_(journals) {
  return sumJournalByAccount_(journals, ['売掛金'], 'credit');
}

function calcCashSales_(journals) {
  var monthly = {};
  journals.forEach(function(j) {
    var ym = getJournalYM_(j);
    if (!ym) return;
    var branches = j.branches || [];
    var hasAR = branches.some(function(b) {
      return (b.debitor && b.debitor.account_name === '売掛金') ||
             (b.creditor && b.creditor.account_name === '売掛金');
    });
    if (!hasAR) {
      branches.forEach(function(b) {
        if (b.creditor && b.creditor.account_name === '売上高' && b.creditor.value > 0) {
          monthly[ym] = (monthly[ym] || 0) + b.creditor.value;
        }
      });
    }
  });
  return monthly;
}

function calcAPPayment_(journals) {
  var payment = sumJournalByAccount_(journals, ['未払金'], 'debit');
  var discount = sumJournalByAccount_(journals, ['仕入値引'], 'credit');
  var result = {};
  Object.keys(payment).concat(Object.keys(discount)).forEach(function(ym) {
    result[ym] = (payment[ym] || 0) - (discount[ym] || 0);
  });
  return result;
}

function calcCashPurchase_(journals) {
  var monthly = {};
  journals.forEach(function(j) {
    var ym = getJournalYM_(j);
    if (!ym) return;
    var branches = j.branches || [];
    var hasPayable = branches.some(function(b) {
      return (b.creditor && b.creditor.account_name === '未払金');
    });
    if (!hasPayable) {
      branches.forEach(function(b) {
        if (b.debitor && b.debitor.account_name === '仕入高' && b.debitor.value > 0) {
          monthly[ym] = (monthly[ym] || 0) + b.debitor.value;
        }
      });
    }
  });
  return monthly;
}

function calcLoanRepayment_(journals) {
  return {
    short: sumJournalByAccount_(journals, ['短期借入金'], 'debit'),
    long:  sumJournalByAccount_(journals, ['長期借入金'], 'debit')
  };
}

// ==============================
// メイン同期
// ==============================

function syncFromMF(year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('資金繰り表_' + year);
  if (!sheet) throw new Error('資金繰り表_' + year + ' シートが見つかりません');

  var keys = monthKeys_(year);
  var fiscalYear = year - 1; // 第7期(2025) → fiscal_year=2024
  ss.toast('推移表・仕訳データ取得中...', '同期開始');

  // --- 推移表（PL/BS）取得 ---
  var plData, bsData;
  try {
    plData = fetchTransition_('pl', fiscalYear);
    Logger.log('PL推移表取得成功');
  } catch (e) {
    Logger.log('PL推移表エラー: ' + e.message);
    plData = null;
  }
  try {
    bsData = fetchTransition_('bs', fiscalYear);
    Logger.log('BS推移表取得成功');
  } catch (e) {
    Logger.log('BS推移表エラー: ' + e.message);
    bsData = null;
  }

  // --- 仕訳データ取得 ---
  var journals = fetchJournals_(year);

  // --- 千円変換 ---
  function toK(val) { return Math.round((val || 0) / 1000); }
  function monthlyToK(data) { return keys.map(function(k) { return toK(data[k]); }); }

  // === PL推移表から取得する項目 ===
  if (plData) {
    // 行3: 売上（今期）
    var sales = extractTransitionAccount_(plData, '売上高', keys);
    sheet.getRange('C3:N3').setValues([monthlyToK(sales)]);

    // 行13: 人件費
    var personnel = extractTransitionSum_(plData, ['給料手当', '賞与', '法定福利費', '役員報酬'], keys);
    sheet.getRange('C13:N13').setValues([monthlyToK(personnel)]);

    // 行15: 諸経費（販管費系の合計 - 個別に取得するより推移表の合計を使う）
    var expenses = extractTransitionSum_(plData, [
      '広告宣伝費', '支払手数料', '通信費', '旅費交通費', '消耗品費',
      '地代家賃', '水道光熱費', '保険料', '租税公課', '外注費',
      '荷造運賃', '雑費', '減価償却費', '支払報酬', '新聞図書費',
      '接待交際費', '会議費', '福利厚生費', '業務委託費'
    ], keys);
    sheet.getRange('C15:N15').setValues([monthlyToK(expenses)]);

    // 行18: 経常外収入
    var nonOpIncome = extractTransitionSum_(plData, ['受取利息', '雑収入', '受取配当金'], keys);
    sheet.getRange('C18:N18').setValues([monthlyToK(nonOpIncome)]);

    // 行19: 経常外支出
    var nonOpExpense = extractTransitionSum_(plData, ['支払利息', '雑損失'], keys);
    sheet.getRange('C19:N19').setValues([monthlyToK(nonOpExpense)]);
  }

  // === BS推移表から取得する項目 ===
  if (bsData) {
    // 行14: 商品棚卸高（月次の在庫増減）
    var inventory = extractTransitionAccount_(bsData, '商品', keys);
    var invChange = [];
    for (var i = 0; i < keys.length; i++) {
      if (i === 0) {
        // 3月の前月(2月末)データは推移表にないので初月は0
        invChange.push(0);
      } else {
        invChange.push(toK((inventory[keys[i]] || 0) - (inventory[keys[i-1]] || 0)));
      }
    }
    sheet.getRange('C14:N14').setValues([invChange]);
  }

  // === 仕訳から取得する項目 ===
  var cashSales    = calcCashSales_(journals);
  var arCollection = calcARCollection_(journals);
  var cashPurchase = calcCashPurchase_(journals);
  var apPayment    = calcAPPayment_(journals);
  var loanRepay    = calcLoanRepayment_(journals);

  sheet.getRange('C8:N8').setValues([monthlyToK(cashSales)]);
  sheet.getRange('C9:N9').setValues([monthlyToK(arCollection)]);
  sheet.getRange('C11:N11').setValues([monthlyToK(cashPurchase)]);
  sheet.getRange('C12:N12').setValues([monthlyToK(apPayment)]);
  sheet.getRange('C26:N26').setValues([monthlyToK(loanRepay.short)]);
  sheet.getRange('C27:N27').setValues([monthlyToK(loanRepay.long)]);

  // ※ 行5（前月繰越金3月）、行23,24（融資収入）は手入力

  SpreadsheetApp.flush();
  Logger.log('同期完了: 資金繰り表_' + year);
  ss.toast('MFデータの同期が完了しました！', '資金繰り表_' + year);
}
