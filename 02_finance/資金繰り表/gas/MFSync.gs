/**
 * MoneyForward → 資金繰り表 データ同期
 *
 * MF APIの /journals から仕訳データを取得。
 * 仕訳構造: journal.branches[] → { debitor: {account_name, value}, creditor: {account_name, value} }
 */

function sync2025() { syncFromMF(2025); }
function sync2026() { syncFromMF(2026); }

// ==============================
// 期間計算
// ==============================

function fiscalDates_(year) {
  return {
    startDate: (year - 1) + '-03-01',
    endDate:   year + '-02-28',
    startYear: year - 1
  };
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
// 仕訳データ取得
// ==============================

function fetchJournals_(year) {
  var fy = fiscalDates_(year);
  var allJournals = [];
  var page = 1;

  while (true) {
    var data = mfApiGet_('/journals', {
      start_date: fy.startDate,
      end_date: fy.endDate,
      page: page,
      per_page: 500
    });
    var journals = data.journals || [];
    allJournals = allJournals.concat(journals);
    if (journals.length < 500) break;
    if (data.total_pages && page >= data.total_pages) break;
    page++;
    if (page > 200) break;
  }
  Logger.log('仕訳取得完了: ' + allJournals.length + '件');
  return allJournals;
}

// ==============================
// 仕訳の日付を取得
// ==============================

function getJournalYM_(j) {
  // MF APIの日付フィールド（複数候補）
  var d = j.date || j.journal_date || j.recognized_at || j.posted_at || '';
  return String(d).substring(0, 7);
}

// ==============================
// 仕訳集計ユーティリティ
// ==============================

/**
 * 仕訳から指定勘定科目の月次発生額を集計
 * MF API構造: journal.branches[] → { debitor: {account_name, value}, creditor: {account_name, value} }
 *
 * @param {Array} journals
 * @param {Array} accountNames - 対象勘定科目名
 * @param {string} side - 'debit' or 'credit'
 * @return {Object} {"2024-03": 金額, ...}
 */
function sumByAccount_(journals, accountNames, side) {
  var monthly = {};
  journals.forEach(function(j) {
    var ym = getJournalYM_(j);
    if (!ym) return;
    (j.branches || []).forEach(function(b) {
      var entry = (side === 'debit') ? b.debitor : b.creditor;
      if (entry && accountNames.indexOf(entry.account_name) !== -1) {
        var amount = entry.value || 0;
        if (amount > 0) {
          monthly[ym] = (monthly[ym] || 0) + amount;
        }
      }
    });
  });
  return monthly;
}

/**
 * 売掛金回収: 貸方が売掛金
 */
function calcARCollection_(journals) {
  return sumByAccount_(journals, ['売掛金'], 'credit');
}

/**
 * 現金売上: 貸方が売上高で、同じ仕訳に売掛金がないもの
 */
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

/**
 * 買掛金支払: 借方の未払金 − 貸方の仕入値引き
 */
function calcAPPayment_(journals) {
  var payment = sumByAccount_(journals, ['未払金'], 'debit');
  var discount = sumByAccount_(journals, ['仕入値引'], 'credit');
  var result = {};
  var allYMs = Object.keys(payment).concat(Object.keys(discount));
  allYMs.forEach(function(ym) {
    result[ym] = (payment[ym] || 0) - (discount[ym] || 0);
  });
  return result;
}

/**
 * 現金仕入: 借方が仕入高で、同じ仕訳に未払金がないもの
 */
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

/**
 * 借入金返済
 */
function calcLoanRepayment_(journals) {
  return {
    short: sumByAccount_(journals, ['短期借入金'], 'debit'),
    long:  sumByAccount_(journals, ['長期借入金'], 'debit')
  };
}

// ==============================
// メイン同期処理
// ==============================

function syncFromMF(year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('資金繰り表_' + year);
  if (!sheet) throw new Error('資金繰り表_' + year + ' シートが見つかりません');

  var keys = monthKeys_(year);
  ss.toast('MFデータ取得中...', '同期開始');

  var journals = fetchJournals_(year);

  // 集計
  var sales        = sumByAccount_(journals, ['売上高'], 'credit');
  var cashSales    = calcCashSales_(journals);
  var arCollection = calcARCollection_(journals);
  var cashPurchase = calcCashPurchase_(journals);
  var apPayment    = calcAPPayment_(journals);
  var personnel    = sumByAccount_(journals, ['給料手当', '賞与', '法定福利費', '役員報酬'], 'debit');
  var loanRepay    = calcLoanRepayment_(journals);

  var expenseAccounts = [
    '広告宣伝費', '支払手数料', '通信費', '旅費交通費',
    '消耗品費', '地代家賃', '水道光熱費', '保険料',
    '租税公課', '外注費', '荷造運賃', '雑費',
    '減価償却費', '支払報酬', '新聞図書費', '接待交際費',
    '会議費', '研修費', '福利厚生費', '車両費',
    '修繕費', '事務用品費', 'リース料', '諸会費',
    '支払手数料（非課税）', '業務委託費'
  ];
  var miscExpense = sumByAccount_(journals, expenseAccounts, 'debit');

  // 商品棚卸高（借方=増加, 貸方=減少）
  var invDebit  = sumByAccount_(journals, ['商品'], 'debit');
  var invCredit = sumByAccount_(journals, ['商品'], 'credit');
  var inventoryChange = {};
  keys.forEach(function(k) { inventoryChange[k] = (invDebit[k] || 0) - (invCredit[k] || 0); });

  var nonOpIncome  = sumByAccount_(journals, ['受取利息', '雑収入', '受取配当金'], 'credit');
  var nonOpExpense = sumByAccount_(journals, ['支払利息', '雑損失'], 'debit');

  // 千円単位変換
  function toK(val) { return Math.round((val || 0) / 1000); }
  function monthlyToK(data) { return keys.map(function(k) { return toK(data[k]); }); }

  // スプシ書き込み（C〜N列 = 3月〜2月）
  sheet.getRange('C3:N3').setValues([monthlyToK(sales)]);
  sheet.getRange('C8:N8').setValues([monthlyToK(cashSales)]);
  sheet.getRange('C9:N9').setValues([monthlyToK(arCollection)]);
  sheet.getRange('C11:N11').setValues([monthlyToK(cashPurchase)]);
  sheet.getRange('C12:N12').setValues([monthlyToK(apPayment)]);
  sheet.getRange('C13:N13').setValues([monthlyToK(personnel)]);
  sheet.getRange('C14:N14').setValues([monthlyToK(inventoryChange)]);
  sheet.getRange('C15:N15').setValues([monthlyToK(miscExpense)]);
  sheet.getRange('C18:N18').setValues([monthlyToK(nonOpIncome)]);
  sheet.getRange('C19:N19').setValues([monthlyToK(nonOpExpense)]);
  sheet.getRange('C26:N26').setValues([monthlyToK(loanRepay.short)]);
  sheet.getRange('C27:N27').setValues([monthlyToK(loanRepay.long)]);

  // デバッグ: 集計結果をログ
  Logger.log('売上: ' + JSON.stringify(sales));
  Logger.log('売掛金回収: ' + JSON.stringify(arCollection));
  Logger.log('買掛金支払: ' + JSON.stringify(apPayment));

  SpreadsheetApp.flush();
  ss.toast('MFデータの同期が完了しました！', '資金繰り表_' + year);
}
