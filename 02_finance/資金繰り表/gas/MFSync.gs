/**
 * MoneyForward → 資金繰り表 データ同期
 *
 * エンドポイント:
 *   /reports/transition_pl?type=monthly  → PL月次推移（売上・経費）
 *   /reports/transition_bs?type=monthly  → BS月次推移（残高）
 *   /journals                            → 仕訳（売掛回収・借入返済）
 */

function sync2025() { syncFromMF(2025); }
function sync2026() { syncFromMF(2026); }

// ==============================
// 期間計算
// ==============================

function fiscalDates_(year) {
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

function fetchTransition_(type, fiscalYear) {
  var endpoint = (type === 'pl') ? '/reports/transition_pl' : '/reports/transition_bs';
  return mfApiGet_(endpoint, { type: 'monthly', fiscal_year: fiscalYear });
}

/**
 * 推移表レスポンスから指定名の月次データを抽出（account/financial_statement_item両対応）
 */
function extractTransitionAccount_(data, accountName, monthKeys) {
  var columns = data.columns || [];
  var result = {};

  var colToYM = {};
  monthKeys.forEach(function(ym) {
    var month = parseInt(ym.split('-')[1]);
    var colIdx = columns.indexOf(String(month));
    if (colIdx !== -1) colToYM[colIdx] = ym;
  });

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

function extractTransitionSum_(data, accountNames, monthKeys) {
  var totals = {};
  monthKeys.forEach(function(k) { totals[k] = 0; });
  accountNames.forEach(function(name) {
    var vals = extractTransitionAccount_(data, name, monthKeys);
    monthKeys.forEach(function(k) { totals[k] += (vals[k] || 0); });
  });
  return totals;
}

// ==============================
// 仕訳データ取得
// ==============================

function fetchJournals_(year) {
  var fy = fiscalDates_(year);
  var startDate = fy.startYear + '-03-01';
  var endDate = (fy.startYear + 1) + '-02-28';
  var allJournals = [];
  var page = 1;
  while (true) {
    var data = mfApiGet_('/journals', {
      start_date: startDate, end_date: endDate, page: page, per_page: 10000
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

// 売掛金回収: 貸方が売掛金
function calcARCollection_(journals) {
  return sumJournalByAccount_(journals, ['売掛金'], 'credit');
}

// 現金仕入: 借方が仕入高 かつ 貸方が普通預金（未払金を介さない仕入）
function calcCashPurchase_(journals) {
  var monthly = {};
  journals.forEach(function(j) {
    var ym = getJournalYM_(j);
    if (!ym) return;
    var branches = j.branches || [];
    // 貸方に普通預金がある仕訳のみ
    var hasBankCredit = branches.some(function(b) {
      return (b.creditor && b.creditor.account_name === '普通預金');
    });
    if (hasBankCredit) {
      branches.forEach(function(b) {
        if (b.debitor && (b.debitor.account_name === '仕入高' || b.debitor.account_name === '仕入（国内）' || b.debitor.account_name === '仕入（輸入）') && b.debitor.value > 0) {
          monthly[ym] = (monthly[ym] || 0) + b.debitor.value;
        }
      });
    }
  });
  return monthly;
}

// 買掛金支払: 借方が未払金 かつ 貸方が普通預金 の仕訳 − 貸方が仕入値引 かつ 借方が普通預金 の仕訳
function calcAPPayment_(journals) {
  var monthlyPayment = {};
  var monthlyDiscount = {};
  journals.forEach(function(j) {
    var ym = getJournalYM_(j);
    if (!ym) return;
    var branches = j.branches || [];
    // 未払金×普通預金の仕訳（借方:未払金, 貸方:普通預金）
    var hasBankCredit = branches.some(function(b) {
      return (b.creditor && b.creditor.account_name === '普通預金');
    });
    var hasBankDebit = branches.some(function(b) {
      return (b.debitor && b.debitor.account_name === '普通預金');
    });
    if (hasBankCredit) {
      branches.forEach(function(b) {
        if (b.debitor && b.debitor.account_name === '未払金' && b.debitor.value > 0) {
          monthlyPayment[ym] = (monthlyPayment[ym] || 0) + b.debitor.value;
        }
      });
    }
    // 仕入値引×普通預金の仕訳（借方:普通預金, 貸方:仕入値引）
    if (hasBankDebit) {
      branches.forEach(function(b) {
        if (b.creditor && (b.creditor.account_name === '仕入値引・返品' || b.creditor.account_name === '仕入値引') && b.creditor.value > 0) {
          monthlyDiscount[ym] = (monthlyDiscount[ym] || 0) + b.creditor.value;
        }
      });
    }
  });
  var result = {};
  Object.keys(monthlyPayment).concat(Object.keys(monthlyDiscount)).forEach(function(ym) {
    result[ym] = (monthlyPayment[ym] || 0) - (monthlyDiscount[ym] || 0);
  });
  return result;
}

// 借入金返済
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

  // --- 推移表取得 ---
  var plData = null, bsData = null;
  try { plData = fetchTransition_('pl', fiscalYear); Logger.log('PL推移表取得成功'); } catch (e) { Logger.log('PL推移表エラー: ' + e.message); }
  try { bsData = fetchTransition_('bs', fiscalYear); Logger.log('BS推移表取得成功'); } catch (e) { Logger.log('BS推移表エラー: ' + e.message); }

  // --- 仕訳取得 ---
  var journals = fetchJournals_(year);

  // --- 千円変換 ---
  function toK(val) { return Math.round((val || 0) / 1000); }
  function monthlyToK(data) { return keys.map(function(k) { return toK(data[k]); }); }

  // ========================================
  // PL推移表から取得
  // ========================================
  if (plData) {
    // 行3: 今期売上 = 売上高合計
    var sales = extractTransitionAccount_(plData, '売上高合計', keys);
    sheet.getRange('C3:N3').setValues([monthlyToK(sales)]);

    // 行4: 前期売上 = 前年度の売上高合計
    try {
      var prevFY = fiscalYear - 1;
      Logger.log('前期売上: fiscal_year=' + prevFY);
      var prevPL = fetchTransition_('pl', prevFY);
      Logger.log('前期PL取得成功: columns=' + JSON.stringify(prevPL.columns));
      var prevKeys = monthKeys_(year - 1);
      Logger.log('前期monthKeys: ' + JSON.stringify(prevKeys));
      var prevSales = extractTransitionAccount_(prevPL, '売上高合計', prevKeys);
      Logger.log('前期売上データ: ' + JSON.stringify(prevSales));
      sheet.getRange('C4:N4').setValues([monthlyToK(prevSales)]);
    } catch (e) { Logger.log('前期売上取得エラー: ' + e.message + '\n' + e.stack); }

    // 行13: 人件費 = 役員賞与 + 役員報酬 + 法定福利費
    var personnel = extractTransitionSum_(plData, ['役員賞与', '役員報酬', '法定福利費'], keys);
    sheet.getRange('C13:N13').setValues([monthlyToK(personnel)]);

    // 行14: 商品棚卸高 = 期末商品棚卸高 − 期首商品棚卸高
    var endInventory = extractTransitionAccount_(plData, '期末商品棚卸高', keys);
    var beginInventory = extractTransitionAccount_(plData, '期首商品棚卸高', keys);
    var inventoryChange = keys.map(function(k) {
      return toK((endInventory[k] || 0) - (beginInventory[k] || 0));
    });
    sheet.getRange('C14:N14').setValues([inventoryChange]);

    // 行18: 経常外収入 = 営業外収益合計
    var nonOpIncome = extractTransitionAccount_(plData, '営業外収益合計', keys);
    sheet.getRange('C18:N18').setValues([monthlyToK(nonOpIncome)]);

    // 行19: 経常外支出 = 営業外費用合計
    var nonOpExpense = extractTransitionAccount_(plData, '営業外費用合計', keys);
    sheet.getRange('C19:N19').setValues([monthlyToK(nonOpExpense)]);
  }

  // ========================================
  // BS推移表から取得
  // ========================================
  if (bsData) {
    // 行5: 前月繰越金 = 前月末の（現金及び預金合計 + 売掛金）
    var cashDeposits = extractTransitionAccount_(bsData, '現金及び預金合計', keys);
    var arBalance = extractTransitionAccount_(bsData, '売掛金', keys);

    // 3月の前月繰越金 = 前年度2月末残高
    var prevCash = 0, prevAR = 0;
    try {
      var prevBS = fetchTransition_('bs', fiscalYear - 1);
      var prevKeys = monthKeys_(year - 1);
      var prevCashData = extractTransitionAccount_(prevBS, '現金及び預金合計', prevKeys);
      var prevARData = extractTransitionAccount_(prevBS, '売掛金', prevKeys);
      var lastPrevKey = prevKeys[prevKeys.length - 1];
      prevCash = prevCashData[lastPrevKey] || 0;
      prevAR = prevARData[lastPrevKey] || 0;
    } catch (e) { Logger.log('前年度BS取得エラー: ' + e.message); }

    var carryForward = [];
    for (var i = 0; i < keys.length; i++) {
      if (i === 0) {
        carryForward.push(toK(prevCash + prevAR));
      } else {
        var prevKey = keys[i - 1];
        carryForward.push(toK((cashDeposits[prevKey] || 0) + (arBalance[prevKey] || 0)));
      }
    }
    sheet.getRange('C5:N5').setValues([carryForward]);
  }

  // ========================================
  // 仕訳から取得
  // ========================================
  var arCollection = calcARCollection_(journals);
  var cashPurchase = calcCashPurchase_(journals);
  var apPayment    = calcAPPayment_(journals);
  var loanRepay    = calcLoanRepayment_(journals);

  var cashPurchaseK = monthlyToK(cashPurchase);
  var apPaymentK    = monthlyToK(apPayment);

  // 行8: 現金売上 → 0（Amazon物販は全て売掛金回収）
  sheet.getRange('C8:N8').setValues([keys.map(function() { return 0; })]);

  // 行9: 売掛金回収
  sheet.getRange('C9:N9').setValues([monthlyToK(arCollection)]);

  // 行11: 現金仕入（仕入×普通預金の仕訳）
  sheet.getRange('C11:N11').setValues([cashPurchaseK]);

  // 行12: 買掛金支払（未払金×普通預金 − 仕入値引×普通預金）
  sheet.getRange('C12:N12').setValues([apPaymentK]);

  // 行15: 諸経費 = 販売費及び一般管理費合計 − 現金仕入 − 買掛金支払 − 人件費
  if (plData) {
    var sgaTotal = extractTransitionAccount_(plData, '販売費及び一般管理費合計', keys);
    var personnel = extractTransitionSum_(plData, ['役員賞与', '役員報酬', '法定福利費'], keys);
    var expenseValues = keys.map(function(k, i) {
      return toK(sgaTotal[k] || 0) - cashPurchaseK[i] - apPaymentK[i] - toK(personnel[k] || 0);
    });
    sheet.getRange('C15:N15').setValues([expenseValues]);
  }

  // 行26: 借入金返済（短期）
  sheet.getRange('C26:N26').setValues([monthlyToK(loanRepay.short)]);

  // 行27: 借入金返済（長期）
  sheet.getRange('C27:N27').setValues([monthlyToK(loanRepay.long)]);

  // ※ 行23,24（融資収入）は手入力

  SpreadsheetApp.flush();
  Logger.log('同期完了: 資金繰り表_' + year);
  ss.toast('MFデータの同期が完了しました！', '資金繰り表_' + year);
}
