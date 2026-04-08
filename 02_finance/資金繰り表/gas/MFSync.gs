/**
 * MoneyForward → 資金繰り表 データ同期
 *
 * MF APIの /journals エンドポイントから仕訳データを取得し、
 * 各勘定科目の月次集計を行って資金繰り表に反映する。
 *
 * 使い方:
 *   syncFromMF(2025)  → 第7期（資金繰り表_2025）にMFデータを反映
 *   syncFromMF(2026)  → 第8期（資金繰り表_2026）にMFデータを反映
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
  for (var m = 3; m <= 12; m++) {
    keys.push(fy.startYear + '-' + padZero_(m));
  }
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
    if (journals.length === 0) break;
    allJournals = allJournals.concat(journals);
    page++;
    if (page > 200) break;
  }
  Logger.log('仕訳取得完了: ' + allJournals.length + '件');
  return allJournals;
}

// ==============================
// 仕訳集計ユーティリティ
// ==============================

/**
 * 仕訳から指定勘定科目の月次発生額を集計
 * @param {Array} journals - 仕訳データ
 * @param {Array} accountNames - 対象勘定科目名
 * @param {string} side - 'debit'(借方) or 'credit'(貸方)
 * @return {Object} {"2024-03": 金額, ...}
 */
function sumByAccount_(journals, accountNames, side) {
  var monthly = {};
  journals.forEach(function(j) {
    var ym = (j.date || '').substring(0, 7);
    (j.details || []).forEach(function(d) {
      if (accountNames.indexOf(d.account_item_name) !== -1) {
        var amount = (side === 'debit') ? (d.debit_amount || 0) : (d.credit_amount || 0);
        if (amount > 0) {
          monthly[ym] = (monthly[ym] || 0) + amount;
        }
      }
    });
  });
  return monthly;
}

/**
 * 売掛金回収: 貸方が売掛金の仕訳（普通預金に入金された分）
 */
function calcARCollection_(journals) {
  return sumByAccount_(journals, ['売掛金'], 'credit');
}

/**
 * 現金売上: 売上高の仕訳で、同じ仕訳に売掛金がないもの
 */
function calcCashSales_(journals) {
  var monthly = {};
  journals.forEach(function(j) {
    var ym = (j.date || '').substring(0, 7);
    var details = j.details || [];
    var hasSafeReceivable = details.some(function(d) {
      return d.account_item_name === '売掛金';
    });
    if (!hasSafeReceivable) {
      details.forEach(function(d) {
        if (d.account_item_name === '売上高' && (d.credit_amount || 0) > 0) {
          monthly[ym] = (monthly[ym] || 0) + d.credit_amount;
        }
      });
    }
  });
  return monthly;
}

/**
 * 買掛金支払: 未払金の支払（借方）− 仕入値引き
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
 * 現金仕入: 仕入高の仕訳で、同じ仕訳に未払金がないもの
 */
function calcCashPurchase_(journals) {
  var monthly = {};
  journals.forEach(function(j) {
    var ym = (j.date || '').substring(0, 7);
    var details = j.details || [];
    var hasPayable = details.some(function(d) {
      return d.account_item_name === '未払金';
    });
    if (!hasPayable) {
      details.forEach(function(d) {
        if (d.account_item_name === '仕入高' && (d.debit_amount || 0) > 0) {
          monthly[ym] = (monthly[ym] || 0) + d.debit_amount;
        }
      });
    }
  });
  return monthly;
}

/**
 * 借入金返済（元本）
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
  if (!sheet) {
    throw new Error('資金繰り表_' + year + ' シートが見つかりません');
  }

  var keys = monthKeys_(year);
  SpreadsheetApp.getActiveSpreadsheet().toast('MFデータ取得中...', '同期開始');

  // --- 仕訳データ取得 ---
  var journals = fetchJournals_(year);

  // --- 集計 ---
  var sales        = sumByAccount_(journals, ['売上高'], 'credit');
  var cashSales    = calcCashSales_(journals);
  var arCollection = calcARCollection_(journals);
  var cashPurchase = calcCashPurchase_(journals);
  var apPayment    = calcAPPayment_(journals);
  var personnel    = sumByAccount_(journals, ['給料手当', '賞与', '法定福利費', '役員報酬'], 'debit');
  var loanRepay    = calcLoanRepayment_(journals);

  // 諸経費: 販管費系の勘定科目（人件費・仕入以外）
  var expenseAccounts = [
    '広告宣伝費', '支払手数料', '通信費', '旅費交通費',
    '消耗品費', '地代家賃', '水道光熱費', '保険料',
    '租税公課', '外注費', '荷造運賃', '雑費',
    '減価償却費', '支払報酬', '新聞図書費', '接待交際費',
    '会議費', '研修費', '福利厚生費', '車両費',
    '修繕費', '事務用品費', 'リース料', '諸会費'
  ];
  var miscExpense = sumByAccount_(journals, expenseAccounts, 'debit');

  // 商品棚卸高: 商品勘定の増減（借方=増加, 貸方=減少）
  var invDebit  = sumByAccount_(journals, ['商品'], 'debit');
  var invCredit = sumByAccount_(journals, ['商品'], 'credit');
  var inventoryChange = {};
  keys.forEach(function(k) {
    inventoryChange[k] = (invDebit[k] || 0) - (invCredit[k] || 0);
  });

  // 経常外収支
  var nonOpIncome  = sumByAccount_(journals, ['受取利息', '雑収入', '受取配当金'], 'credit');
  var nonOpExpense = sumByAccount_(journals, ['支払利息', '雑損失'], 'debit');

  // --- 千円単位変換 ---
  function toK(val) { return Math.round((val || 0) / 1000); }
  function monthlyToK(data) {
    return keys.map(function(k) { return toK(data[k]); });
  }

  // --- スプシに書き込み ---
  // C〜N列 = 3月〜2月

  // 行3: 売上（今期）
  sheet.getRange('C3:N3').setValues([monthlyToK(sales)]);

  // 行5: 前月繰越金 → 3月は手入力のまま（自動計算が難しいため）

  // 行8: 現金売上
  sheet.getRange('C8:N8').setValues([monthlyToK(cashSales)]);

  // 行9: 売掛金回収
  sheet.getRange('C9:N9').setValues([monthlyToK(arCollection)]);

  // 行11: 現金仕入
  sheet.getRange('C11:N11').setValues([monthlyToK(cashPurchase)]);

  // 行12: 買掛金支払（未払金 − 仕入値引き）
  sheet.getRange('C12:N12').setValues([monthlyToK(apPayment)]);

  // 行13: 人件費
  sheet.getRange('C13:N13').setValues([monthlyToK(personnel)]);

  // 行14: 商品棚卸高
  sheet.getRange('C14:N14').setValues([monthlyToK(inventoryChange)]);

  // 行15: 諸経費
  sheet.getRange('C15:N15').setValues([monthlyToK(miscExpense)]);

  // 行18: 経常外収入
  sheet.getRange('C18:N18').setValues([monthlyToK(nonOpIncome)]);

  // 行19: 経常外支出
  sheet.getRange('C19:N19').setValues([monthlyToK(nonOpExpense)]);

  // 行26: 借入金返済（短期）
  sheet.getRange('C26:N26').setValues([monthlyToK(loanRepay.short)]);

  // 行27: 借入金返済（長期）
  sheet.getRange('C27:N27').setValues([monthlyToK(loanRepay.long)]);

  // ※ 行5（前月繰越金の3月）、行23,24（公庫・信金の借入）は手入力

  SpreadsheetApp.flush();
  Logger.log('同期完了: 資金繰り表_' + year);
  SpreadsheetApp.getActiveSpreadsheet().toast('MFデータの同期が完了しました！', '資金繰り表_' + year);
}
