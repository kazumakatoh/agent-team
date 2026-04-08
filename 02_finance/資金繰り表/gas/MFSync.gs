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
  // 資金繰り表_2025 → MF fiscal_year=2025（2025年度=2025/3〜2026/2）
  return { startYear: year, endYear: year + 1, fiscalYear: year };
}

function monthKeys_(year) {
  var fy = fiscalDates_(year);
  var keys = [];
  for (var m = 3; m <= 12; m++) keys.push(fy.startYear + '-' + padZero_(m));
  keys.push((fy.startYear + 1) + '-01');
  keys.push((fy.startYear + 1) + '-02');
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
      if (row.name === accountName && row.values) {
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
  var endDate = (fy.startYear + 1) + '-02-28';
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
  var fy = fiscalDates_(year);
  var fiscalYear = fy.fiscalYear;
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
    // 行3: 売上（今期）※ MF上は「売上高合計」（financial_statement_item）
    var sales = extractTransitionAccount_(plData, '売上高合計', keys);
    sheet.getRange('C3:N3').setValues([monthlyToK(sales)]);

    // 行4: 売上（前期）
    try {
      var prevFY = fiscalYear - 1;
      var prevPL = fetchTransition_('pl', prevFY);
      var prevKeys = monthKeys_(year - 1);
      var prevSales = extractTransitionAccount_(prevPL, '売上高合計', prevKeys);
      sheet.getRange('C4:N4').setValues([monthlyToK(prevSales)]);
    } catch (e) {
      Logger.log('前期売上取得エラー: ' + e.message);
    }

    // 行13: 人件費
    var personnel = extractTransitionSum_(plData, ['給料手当', '賞与', '法定福利費', '役員報酬'], keys);
    sheet.getRange('C13:N13').setValues([monthlyToK(personnel)]);

    // 行15: 諸経費 = 販売費及び一般管理費合計 − 人件費
    var sgaTotal = extractTransitionAccount_(plData, '販売費及び一般管理費合計', keys);
    var expenseValues = keys.map(function(k) {
      return toK((sgaTotal[k] || 0) - (personnel[k] || 0));
    });
    sheet.getRange('C15:N15').setValues([expenseValues]);

    // 行18: 経常外収入 = 営業外収益合計
    var nonOpIncome = extractTransitionAccount_(plData, '営業外収益合計', keys);
    sheet.getRange('C18:N18').setValues([monthlyToK(nonOpIncome)]);

    // 行19: 経常外支出 = 営業外費用合計
    var nonOpExpense = extractTransitionAccount_(plData, '営業外費用合計', keys);
    sheet.getRange('C19:N19').setValues([monthlyToK(nonOpExpense)]);
  }

  // === BS推移表から取得する項目 ===
  if (bsData) {
    // 行5: 前月繰越金 = 前月末の（現金及び預金合計 + 売掛金）
    // BS推移表のcolumnsは当月末残高を表す
    // 3月の前月繰越金 = 前年度2月末残高 → 前年度BS推移表が必要
    var cashDeposits = extractTransitionAccount_(bsData, '現金及び預金合計', keys);
    var arBalance = extractTransitionAccount_(bsData, '売掛金', keys);

    // 前年度BSから2月末残高を取得（3月の前月繰越金用）
    var prevCash = 0, prevAR = 0;
    try {
      var prevBS = fetchTransition_('bs', fiscalYear - 1);
      var prevKeys = monthKeys_(year - 1);
      var prevCashData = extractTransitionAccount_(prevBS, '現金及び預金合計', prevKeys);
      var prevARData = extractTransitionAccount_(prevBS, '売掛金', prevKeys);
      var lastPrevKey = prevKeys[prevKeys.length - 1]; // 前年度2月
      prevCash = prevCashData[lastPrevKey] || 0;
      prevAR = prevARData[lastPrevKey] || 0;
    } catch (e) {
      Logger.log('前年度BS取得エラー: ' + e.message);
    }

    // 3月 = 前年度2月末残高、4月以降 = 前月末残高
    var carryForward = [];
    for (var i = 0; i < keys.length; i++) {
      if (i === 0) {
        carryForward.push(toK(prevCash + prevAR));
      } else {
        var prevKey = keys[i - 1];
        carryForward.push(toK((cashDeposits[prevKey] || 0) + (arBalance[prevKey] || 0)));
      }
    }
    // 3月の前月繰越金を上書き（数式ではなく値で）
    sheet.getRange('C5').setValue(carryForward[0]);
    // 4月以降は数式が入っているのでそのまま（翌月繰越金から自動計算）
    // ただし数式がない場合のフォールバックとしてログ出力
    Logger.log('前月繰越金: ' + JSON.stringify(carryForward));

    // 行14: 商品棚卸高（月次の在庫増減）
    var inventory = extractTransitionAccount_(bsData, '商品', keys);
    var invChange = [];
    for (var i = 0; i < keys.length; i++) {
      if (i === 0) {
        invChange.push(0); // 3月の前月データなし
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
