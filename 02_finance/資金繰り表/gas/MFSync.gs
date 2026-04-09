/**
 * MoneyForward → 資金繰り表 データ同期
 *
 * 設計方針:
 *   全ての経常収支・財務収支を「普通預金」の入出金ベースで集計。
 *   翌月繰越金 = MFの普通預金当月末残高と一致する設計。
 *
 * データソース:
 *   /reports/transition_pl  → 今期売上・前期売上（発生ベース）
 *   /reports/transition_bs  → 前月繰越金（普通預金残高）、整合性チェック
 *   /journals               → 経常収支・経常外収支・財務収支（全て普通預金入出金ベース）
 */

function sync2025() { syncFromMF(2025); }
function sync2026() { syncFromMF(2026); }

// ==============================
// 期間計算
// ==============================

function fiscalDates_(year) {
  return { startYear: year, fiscalYear: year };
}

function monthKeys_(year) {
  var keys = [];
  for (var m = 3; m <= 12; m++) keys.push(year + '-' + padZero_(m));
  keys.push((year + 1) + '-01');
  keys.push((year + 1) + '-02');
  return keys;
}

function padZero_(n) { return n < 10 ? '0' + n : '' + n; }

// ==============================
// 推移表取得・パース
// ==============================

function fetchTransition_(type, fiscalYear) {
  var endpoint = (type === 'pl') ? '/reports/transition_pl' : '/reports/transition_bs';
  return mfApiGet_(endpoint, { type: 'monthly', fiscal_year: fiscalYear });
}

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
          if (colToYM[idx] !== undefined) result[colToYM[idx]] = val || 0;
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
// 仕訳取得
// ==============================

function fetchJournals_(year) {
  var startDate = year + '-03-01';
  var endDate = (year + 1) + '-02-28';
  var allJournals = [];
  var page = 1;
  while (true) {
    var data = mfApiGet_('/journals', {
      start_date: startDate, end_date: endDate, page: page, per_page: 10000
    });
    allJournals = allJournals.concat(data.journals || []);
    var totalPages = (data.metadata && data.metadata.total_pages) || 1;
    if (page >= totalPages) break;
    page++;
    if (page > 100) break;
  }
  Logger.log('仕訳取得: ' + allJournals.length + '件');
  return allJournals;
}

// ==============================
// 普通預金ベースの仕訳分類
// ==============================

/**
 * 全仕訳を普通預金の入出金で分類
 * 返り値: { 入金カテゴリ: {ym: 金額}, 出金カテゴリ: {ym: 金額} }
 */
function classifyBankTransactions_(journals) {
  var result = {
    // 入金（普通預金 借方）
    cashSales: {},       // 売上高
    arCollection: {},    // 売掛金
    nonOpIncome: {},     // 受取利息・雑収入・受取配当金
    loanIncome: {},      // 短期/長期借入金
    otherIncome: {},     // その他入金
    // 出金（普通預金 貸方）
    cashPurchase: {},    // 仕入高（未払金/買掛金を介さない）
    apPayment: {},       // 未払金・買掛金
    personnel: {},       // 役員報酬・給料手当・法定福利費
    nonOpExpense: {},    // 支払利息・雑損失・為替差損
    loanRepayShort: {},  // 短期借入金
    loanRepayLong: {},   // 長期借入金
    miscExpense: {},     // 諸経費（上記以外の販管費）
    otherExpense: {}     // その他出金
  };

  // 出金で除外する勘定科目（諸経費に含めない）
  var excludeFromMisc = [
    '仕入高', '仕入（国内）', '仕入（輸入）',
    '未払金', '買掛金',
    '役員報酬', '給料手当', '法定福利費', '役員賞与',
    '支払利息', '雑損失', '為替差損',
    '短期借入金', '長期借入金',
    '普通預金', '現金', '当座預金', '定期預金',
    '売掛金', '売上高',
    '仮払消費税', '仮受消費税', '仮払金', '仮受金',
    '法人税、住民税及び事業税', '未払法人税等',
    '預り金', '預け金', '貸付金', '借入金',
    '商品', '減価償却費', '繰延資産償却'
  ];

  var nonOpIncomeAccounts = ['受取利息', '雑収入', '受取配当金'];
  var nonOpExpenseAccounts = ['支払利息', '雑損失', '為替差損'];
  var personnelAccounts = ['役員報酬', '給料手当', '法定福利費', '役員賞与'];
  var purchaseAccounts = ['仕入高', '仕入（国内）', '仕入（輸入）'];
  var loanAccounts = ['短期借入金', '長期借入金'];

  function addTo(obj, ym, amount) {
    obj[ym] = (obj[ym] || 0) + amount;
  }

  journals.forEach(function(j) {
    var ym = (j.transaction_date || '').substring(0, 7);
    if (!ym) return;
    var branches = j.branches || [];

    branches.forEach(function(b) {
      // === 普通預金が借方（入金） ===
      if (b.debitor && b.debitor.account_name === '普通預金' && b.creditor && b.creditor.value > 0) {
        var cr = b.creditor;
        var amount = b.debitor.value;

        if (cr.account_name === '売上高') {
          addTo(result.cashSales, ym, amount);
        } else if (cr.account_name === '売掛金') {
          addTo(result.arCollection, ym, amount);
        } else if (nonOpIncomeAccounts.indexOf(cr.account_name) !== -1) {
          addTo(result.nonOpIncome, ym, amount);
        } else if (loanAccounts.indexOf(cr.account_name) !== -1) {
          addTo(result.loanIncome, ym, amount);
        } else {
          addTo(result.otherIncome, ym, amount);
        }
      }

      // === 普通預金が貸方（出金） ===
      if (b.creditor && b.creditor.account_name === '普通預金' && b.debitor && b.debitor.value > 0) {
        var dr = b.debitor;
        var amount = b.creditor.value;

        if (purchaseAccounts.indexOf(dr.account_name) !== -1) {
          addTo(result.cashPurchase, ym, amount);
        } else if (dr.account_name === '未払金' || dr.account_name === '買掛金') {
          addTo(result.apPayment, ym, amount);
        } else if (personnelAccounts.indexOf(dr.account_name) !== -1) {
          addTo(result.personnel, ym, amount);
        } else if (nonOpExpenseAccounts.indexOf(dr.account_name) !== -1) {
          addTo(result.nonOpExpense, ym, amount);
        } else if (dr.account_name === '短期借入金') {
          addTo(result.loanRepayShort, ym, amount);
        } else if (dr.account_name === '長期借入金') {
          addTo(result.loanRepayLong, ym, amount);
        } else if (excludeFromMisc.indexOf(dr.account_name) === -1) {
          addTo(result.miscExpense, ym, amount);
        } else {
          addTo(result.otherExpense, ym, amount);
        }
      }
    });
  });

  return result;
}

// ==============================
// メイン同期
// ==============================

function syncFromMF(year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('資金繰り表_' + year);
  if (!sheet) throw new Error('資金繰り表_' + year + ' シートが見つかりません');

  var keys = monthKeys_(year);
  var fiscalYear = year;
  ss.toast('データ取得中...', '同期開始');

  // --- 推移表取得 ---
  var plData = null, bsData = null;
  try { plData = fetchTransition_('pl', fiscalYear); } catch (e) { Logger.log('PL推移表エラー: ' + e.message); }
  try { bsData = fetchTransition_('bs', fiscalYear); } catch (e) { Logger.log('BS推移表エラー: ' + e.message); }

  // --- 仕訳取得 & 普通預金ベース分類 ---
  var journals = fetchJournals_(year);
  var bank = classifyBankTransactions_(journals);

  // --- 千円変換 ---
  function toK(val) { return Math.round((val || 0) / 1000); }
  function monthlyToK(data) { return keys.map(function(k) { return toK(data[k]); }); }

  // ========================================
  // PL推移表（発生ベース）
  // ========================================
  if (plData) {
    // 行3: 今期売上 = 売上高合計（発生ベース）
    var sales = extractTransitionAccount_(plData, '売上高合計', keys);
    sheet.getRange('C3:N3').setValues([monthlyToK(sales)]);

    // 行4: 前期売上 = 前年度PL推移表の売上高合計
    try {
      var prevPL = fetchTransition_('pl', fiscalYear - 1);
      var prevKeys = monthKeys_(year - 1);
      var prevSales = extractTransitionAccount_(prevPL, '売上高合計', prevKeys);
      sheet.getRange('C4:N4').setValues([monthlyToK(prevSales)]);
    } catch (e) { Logger.log('前期売上エラー: ' + e.message); }
  }

  // ========================================
  // BS推移表（普通預金残高）
  // ========================================
  if (bsData) {
    // 行5: 前月繰越金 = 前月末の普通預金残高
    var bankBalance = extractTransitionAccount_(bsData, '普通預金', keys);

    // 3月の前月繰越金 = 前年度2月末の普通預金残高
    var prevBankBal = 0;
    try {
      var prevBS = fetchTransition_('bs', fiscalYear - 1);
      var prevKeys = monthKeys_(year - 1);
      var prevBankData = extractTransitionAccount_(prevBS, '普通預金', prevKeys);
      prevBankBal = prevBankData[prevKeys[prevKeys.length - 1]] || 0;
    } catch (e) { Logger.log('前年度BS取得エラー: ' + e.message); }

    var carryForward = [];
    for (var i = 0; i < keys.length; i++) {
      if (i === 0) {
        carryForward.push(toK(prevBankBal));
      } else {
        carryForward.push(toK(bankBalance[keys[i - 1]] || 0));
      }
    }
    sheet.getRange('C5:N5').setValues([carryForward]);
  }

  // ========================================
  // 仕訳ベース（普通預金入出金）
  // ========================================

  // --- 経常収入 ---
  sheet.getRange('C8:N8').setValues([monthlyToK(bank.cashSales)]);      // 行8: 現金売上
  sheet.getRange('C9:N9').setValues([monthlyToK(bank.arCollection)]);   // 行9: 売掛金回収

  // --- 経常支出 ---
  sheet.getRange('C11:N11').setValues([monthlyToK(bank.cashPurchase)]); // 行11: 現金仕入
  sheet.getRange('C12:N12').setValues([monthlyToK(bank.apPayment)]);    // 行12: 買掛金支払
  sheet.getRange('C13:N13').setValues([monthlyToK(bank.personnel)]);    // 行13: 人件費
  // 行14: 商品棚卸高 → 集計対象外（0を書き込み）
  sheet.getRange('C14:N14').setValues([keys.map(function() { return 0; })]);
  sheet.getRange('C15:N15').setValues([monthlyToK(bank.miscExpense)]);  // 行15: 諸経費

  // --- 経常外収支 ---
  sheet.getRange('C18:N18').setValues([monthlyToK(bank.nonOpIncome)]);  // 行18: 経常外収入
  sheet.getRange('C19:N19').setValues([monthlyToK(bank.nonOpExpense)]); // 行19: 経常外支出

  // --- 財務収支 ---
  // 行22の「収入」小計は数式。行23,24に内訳を書き込み
  // 融資は全て loanIncome にまとまっているので行22の位置に
  // → 公庫/信金の区分は仕訳だけでは判別困難なので、収入合計を行22に
  // ただしシートの構造上、行23(公庫)・行24(信金)は手入力のまま
  // → loanIncomeをログに出力して手入力の参考にする
  Logger.log('財務収入（借入入金）: ' + JSON.stringify(bank.loanIncome));

  sheet.getRange('C26:N26').setValues([monthlyToK(bank.loanRepayShort)]); // 行26: 短期返済
  sheet.getRange('C27:N27').setValues([monthlyToK(bank.loanRepayLong)]);  // 行27: 長期返済

  // ========================================
  // 整合性チェック: 翌月繰越金 vs MF普通預金残高
  // ========================================
  SpreadsheetApp.flush();

  if (bsData) {
    var bankBalance = extractTransitionAccount_(bsData, '普通預金', keys);
    var checkResults = [];
    var hasError = false;

    for (var i = 0; i < keys.length; i++) {
      var col = String.fromCharCode(67 + i); // C=67
      var calcCarry = sheet.getRange(col + '29').getValue(); // 翌月繰越金（数式計算値）
      var mfBalance = toK(bankBalance[keys[i]] || 0);        // MF普通預金残高
      var diff = calcCarry - mfBalance;

      if (Math.abs(diff) > 1) { // 千円単位の丸め誤差を許容
        checkResults.push(keys[i].substring(5) + '月: 繰越=' + calcCarry + ' MF=' + mfBalance + ' 差=' + diff);
        hasError = true;
      }
    }

    if (hasError) {
      Logger.log('⚠️ 整合性チェック不一致:\n' + checkResults.join('\n'));
      ss.toast('⚠️ 一部の月で繰越金とMF残高に差異あり（実行ログ参照）', '整合性チェック');
    } else {
      Logger.log('✅ 整合性チェックOK: 全月の翌月繰越金がMF普通預金残高と一致');
      ss.toast('✅ 整合性チェックOK', '同期完了');
    }
  }

  Logger.log('同期完了: 資金繰り表_' + year);
}
