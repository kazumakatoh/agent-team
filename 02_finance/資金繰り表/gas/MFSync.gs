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
// 普通預金ベースの仕訳分類（仕訳単位）
// ==============================

/**
 * 全仕訳を普通預金の入出金で分類
 * 複合仕訳対応: 仕訳全体で普通預金の入出金を集計し、相手科目で分類
 */
function classifyBankTransactions_(journals) {
  var result = {
    cashSales: {},       // 入金: 売上高
    arCollection: {},    // 入金: 売掛金
    nonOpIncome: {},     // 入金: 受取利息・雑収入等
    loanIncome: {},      // 入金: 借入金
    otherIncome: {},     // 入金: その他
    cashPurchase: {},    // 出金: 仕入（直接）
    apPayment: {},       // 出金: 未払金・買掛金
    personnel: {},       // 出金: 人件費
    nonOpExpense: {},    // 出金: 支払利息等
    loanRepayShort: {},  // 出金: 短期借入金
    loanRepayLong: {},   // 出金: 長期借入金
    fixedAssetCF: {},    // 出金: 固定資産購入売却
    otherInvestCF: {},   // 出金: その他投資
    miscExpense: {},     // 出金: 諸経費
    otherExpense: {}     // 出金: その他
  };

  function addTo(obj, ym, amount) {
    obj[ym] = (obj[ym] || 0) + amount;
  }

  // 分類定義
  var incomeClassify = {
    '売上高': 'cashSales',
    '売上（国内）': 'cashSales',
    '売上（海外）': 'cashSales',
    '売掛金': 'arCollection',
    '受取利息': 'nonOpIncome',
    '短期借入金': 'loanIncome',
    '長期借入金': 'loanIncome'
  };

  var expenseClassify = {
    '仕入高': 'cashPurchase',
    '仕入（国内）': 'cashPurchase',
    '仕入（輸入）': 'cashPurchase',
    '未払金': 'apPayment',
    '買掛金': 'apPayment',
    '役員報酬': 'personnel',
    '給料手当': 'personnel',
    '給料賃金': 'personnel',
    '法定福利費': 'personnel',
    '役員賞与': 'personnel',
    '支払利息': 'nonOpExpense',
    '短期借入金': 'loanRepayShort',
    '長期借入金': 'loanRepayLong',
    // 投資CF
    // 投資CF - 固定資産
    '工具器具備品': 'fixedAssetCF',
    '建物': 'fixedAssetCF',
    '車両運搬具': 'fixedAssetCF',
    '土地': 'fixedAssetCF',
    '建物附属設備': 'fixedAssetCF',
    '附属設備': 'fixedAssetCF',
    'ソフトウェア': 'fixedAssetCF',
    // 投資CF - その他
    '出資金': 'otherInvestCF',
    '敷金・保証金': 'otherInvestCF',
    '開発費': 'otherInvestCF',
    '繰延資産': 'otherInvestCF',
    '投資有価証券': 'otherInvestCF',
    '長期前払費用': 'otherInvestCF'
  };

  // 諸経費に含めない科目（分類済み or 非現金 or 振替系）
  var excludeFromMisc = [
    '普通預金', '当座預金', '定期預金', '現金',
    '売掛金', '売上高', '商品',
    '仮払消費税', '仮受消費税', '仮払金', '仮受金',
    '預り金', '預け金', '立替金',
    '法人税、住民税及び事業税', '未払法人税等',
    '減価償却費', '繰延資産償却',
    '貸付金', '役員貸付金'
  ];

  journals.forEach(function(j) {
    var ym = (j.transaction_date || '').substring(0, 7);
    if (!ym) return;
    // 開始仕訳・決算整理仕訳を除外
    if (j.entered_by === 'JOURNAL_TYPE_OPENING') return;
    var branches = j.branches || [];

    // 仕訳全体で普通預金の入出金と相手科目を集計
    var bankInflow = 0;   // 普通預金 借方合計（入金）
    var bankOutflow = 0;  // 普通預金 貸方合計（出金）
    var debitAccounts = [];  // 普通預金以外の借方
    var creditAccounts = []; // 普通預金以外の貸方

    branches.forEach(function(b) {
      if (b.debitor) {
        if (b.debitor.account_name === '普通預金') {
          bankInflow += b.debitor.value || 0;
        } else {
          debitAccounts.push({ name: b.debitor.account_name, sub: b.debitor.sub_account_name || '', value: b.debitor.value || 0 });
        }
      }
      if (b.creditor) {
        if (b.creditor.account_name === '普通預金') {
          bankOutflow += b.creditor.value || 0;
        } else {
          creditAccounts.push({ name: b.creditor.account_name, sub: b.creditor.sub_account_name || '', value: b.creditor.value || 0 });
        }
      }
    });

    // === 入金の分類（銀行入金額を貸方科目の比率で按分） ===
    if (bankInflow > 0 && creditAccounts.length > 0) {
      var totalCredit = creditAccounts.reduce(function(s, c) { return s + c.value; }, 0);
      if (totalCredit > 0) {
        creditAccounts.forEach(function(cr) {
          var cashAmount = Math.round(bankInflow * cr.value / totalCredit);
          var cat = incomeClassify[cr.name];
          if (cat) {
            addTo(result[cat], ym, cashAmount);
          } else {
            addTo(result.otherIncome, ym, cashAmount);
          }
        });
      }
    }

    // === 出金の分類（銀行出金額を借方科目の比率で按分） ===
    if (bankOutflow > 0 && debitAccounts.length > 0) {
      var totalDebit = debitAccounts.reduce(function(s, d) { return s + d.value; }, 0);
      if (totalDebit > 0) {
        debitAccounts.forEach(function(dr) {
          var cashAmount = Math.round(bankOutflow * dr.value / totalDebit);
          // 未払費用の補助科目で人件費を判別
          var cat;
          if (dr.name === '未払費用' && /加藤/.test(dr.sub)) {
            cat = 'personnel';
          } else {
            cat = expenseClassify[dr.name];
          }
          if (cat) {
            addTo(result[cat], ym, cashAmount);
          } else if (excludeFromMisc.indexOf(dr.name) === -1) {
            addTo(result.miscExpense, ym, cashAmount);
          } else {
            addTo(result.otherExpense, ym, cashAmount);
          }
        });
      }
    }
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

    // 行4: 前期売上 → 手入力（自動取得しない）
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
  // 仕訳ベース（普通預金入出金）+ 逆算方式
  // ========================================

  // --- 翌月繰越金（= MF普通預金当月末残高）を先に確定 ---
  var endBalance = [];  // 各月末の普通預金残高（千円）
  if (bsData) {
    var bankBalance = extractTransitionAccount_(bsData, '普通預金', keys);
    endBalance = keys.map(function(k) { return toK(bankBalance[k] || 0); });
  }

  // --- 各項目を千円で準備 ---
  var cashSalesK    = monthlyToK(bank.cashSales);
  var arCollectionK = monthlyToK(bank.arCollection);
  var otherIncomeK  = monthlyToK(bank.otherIncome);
  var cashPurchaseK = monthlyToK(bank.cashPurchase);
  var apPaymentK    = monthlyToK(bank.apPayment);
  // 人件費は仕訳ベース（銀行実出金額を按分）
  var personnelK    = monthlyToK(bank.personnel);
  var nonOpIncomeK  = monthlyToK(bank.nonOpIncome);
  var nonOpExpenseK = monthlyToK(bank.nonOpExpense);
  var fixedAssetK   = monthlyToK(bank.fixedAssetCF);
  var otherInvestK  = monthlyToK(bank.otherInvestCF);
  var investCFK     = keys.map(function(k, i) { return fixedAssetK[i] + otherInvestK[i]; });
  var loanIncomeK   = monthlyToK(bank.loanIncome);
  var loanRepShortK = monthlyToK(bank.loanRepayShort);
  var loanRepLongK  = monthlyToK(bank.loanRepayLong);

  // --- 前月繰越金（千円） ---
  var carryForwardK = [];
  if (bsData) {
    var bankBal = extractTransitionAccount_(bsData, '普通預金', keys);
    var prevBankBal = 0;
    try {
      var prevBS = fetchTransition_('bs', fiscalYear - 1);
      var prevKeys = monthKeys_(year - 1);
      var prevBankData = extractTransitionAccount_(prevBS, '普通預金', prevKeys);
      prevBankBal = prevBankData[prevKeys[prevKeys.length - 1]] || 0;
    } catch (e) { Logger.log('前年度BS取得エラー: ' + e.message); }

    for (var i = 0; i < keys.length; i++) {
      if (i === 0) {
        carryForwardK.push(toK(prevBankBal));
      } else {
        carryForwardK.push(toK(bankBal[keys[i - 1]] || 0));
      }
    }
  }

  // --- 諸経費を逆算 ---
  // 諸経費 = 前月繰越金 + 収入合計 - (現金仕入+買掛金+人件費) + 経常外純額 - 投資CF + 財務純額 - 翌月繰越金
  var miscExpenseK = keys.map(function(k, i) {
    var income = cashSalesK[i] + arCollectionK[i] + otherIncomeK[i];
    var knownExpense = cashPurchaseK[i] + apPaymentK[i] + personnelK[i];
    var nonOp = nonOpIncomeK[i] - nonOpExpenseK[i];
    var invest = -investCFK[i];
    var finance = loanIncomeK[i] - loanRepShortK[i] - loanRepLongK[i];
    var misc = carryForwardK[i] + income - knownExpense + nonOp + invest + finance - endBalance[i];
    // マイナス諸経費のデバッグ
    if (misc < 0) {
      Logger.log('⚠️ ' + keys[i].substring(5) + '月 諸経費マイナス: ' + misc);
      Logger.log('  繰越=' + carryForwardK[i] + ' 収入=' + income + ' 既知支出=' + knownExpense + ' 経常外=' + nonOp + ' 投資=' + invest + ' 財務=' + finance + ' 期末=' + endBalance[i]);
    }
    return misc;
  });

  // --- スプシに書き込み（新レイアウト: 投資CFセクション追加） ---
  sheet.getRange('C5:N5').setValues([carryForwardK]);                     // 行5: 前月繰越金
  sheet.getRange('C8:N8').setValues([cashSalesK]);                        // 行8: 現金売上
  sheet.getRange('C9:N9').setValues([arCollectionK]);                     // 行9: 売掛金回収
  sheet.getRange('C11:N11').setValues([cashPurchaseK]);                   // 行11: 現金仕入
  sheet.getRange('C12:N12').setValues([apPaymentK]);                      // 行12: 買掛金支払
  sheet.getRange('C13:N13').setValues([personnelK]);                      // 行13: 人件費
  sheet.getRange('C15:N15').setValues([miscExpenseK]);                    // 行15: 諸経費（逆算）
  sheet.getRange('C18:N18').setValues([nonOpIncomeK]);                    // 行18: 経常外収入
  sheet.getRange('C19:N19').setValues([nonOpExpenseK]);                   // 行19: 経常外支出
  // --- 投資キャッシュフロー ---
  sheet.getRange('C22:N22').setValues([fixedAssetK]);                     // 行22: 固定資産購入売却
  sheet.getRange('C23:N23').setValues([otherInvestK]);                    // 行23: その他投資
  // --- 財務収支 ---
  sheet.getRange('C27:N27').setValues([loanIncomeK]);                     // 行27: 財務収入（公庫に一旦入れる）
  sheet.getRange('C30:N30').setValues([loanRepShortK]);                   // 行30: 短期返済
  sheet.getRange('C31:N31').setValues([loanRepLongK]);                    // 行31: 長期返済
  Logger.log('財務収入（借入入金）: ' + JSON.stringify(bank.loanIncome));
  Logger.log('その他収入（未分類入金）: ' + JSON.stringify(bank.otherIncome));

  // --- 翌月繰越金を値で上書き ---
  sheet.getRange('C33:N33').setValues([endBalance]);

  // --- 整合性チェック ---
  SpreadsheetApp.flush();
  Logger.log('=== 3月分類詳細（千円） ===');
  Logger.log('前月繰越金: ' + carryForwardK[0]);
  Logger.log('現金売上: ' + cashSalesK[0] + ' / 売掛金回収: ' + arCollectionK[0]);
  Logger.log('現金仕入: ' + cashPurchaseK[0] + ' / 買掛金支払: ' + apPaymentK[0] + ' / 人件費: ' + personnelK[0]);
  // 人件費デバッグ: 3月の人件費関連仕訳を表示
  journals.forEach(function(j) {
    var jym = (j.transaction_date || '').substring(0, 7);
    if (jym !== keys[0]) return;
    (j.branches || []).forEach(function(b) {
      if (b.debitor && ['役員報酬','給料手当','給料賃金','法定福利費','役員賞与'].indexOf(b.debitor.account_name) !== -1) {
        Logger.log('  人件費仕訳: ' + b.debitor.account_name + ' ' + b.debitor.value + ' / 貸方=' + (b.creditor ? b.creditor.account_name + ' ' + b.creditor.value : 'null'));
      }
    });
  });
  Logger.log('諸経費（逆算）: ' + miscExpenseK[0]);
  Logger.log('投資CF: ' + investCFK[0]);
  Logger.log('経常外: +' + nonOpIncomeK[0] + ' -' + nonOpExpenseK[0]);
  Logger.log('財務: +' + loanIncomeK[0] + ' -' + loanRepShortK[0] + ' -' + loanRepLongK[0]);
  Logger.log('翌月繰越金（MF残高）: ' + endBalance[0]);
  Logger.log('検算: ' + carryForwardK[0] + ' + ' + (cashSalesK[0]+arCollectionK[0]) + ' - ' + (cashPurchaseK[0]+apPaymentK[0]+personnelK[0]+miscExpenseK[0]) + ' + ... = ' + endBalance[0]);

  Logger.log('✅ 同期完了: 資金繰り表_' + year + '（諸経費は逆算方式）');
  ss.toast('同期完了！翌月繰越金=MF普通預金残高', '資金繰り表_' + year);
}
