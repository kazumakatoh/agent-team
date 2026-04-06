/**
 * キャッシュフロー管理システム - マネーフォワードAPIクライアント
 *
 * MFクラウド会計API (OpenAPI 3.0.3) に基づくデータ取得モジュール
 *
 * 利用エンドポイント:
 *  - GET /api/v3/journals        - 仕訳一覧（入出金明細の取得）
 *  - GET /api/v3/accounts        - 勘定科目一覧（口座IDの特定）
 *  - GET /api/v3/sub_accounts    - 補助科目一覧（銀行口座の特定）
 *  - GET /api/v3/reports/trial_balance_bs - 残高試算表（口座残高）
 *  - GET /api/v3/offices         - 事業者情報
 *
 * 仕訳レスポンス構造:
 *  journal.branches[] → { debitor: { account_name, sub_account_name, value }, creditor: { ... }, remark }
 */

// ==============================
// 口座入出金明細の取得（仕訳ベース）
// ==============================

/**
 * 指定口座の入出金明細をMFの仕訳データから取得する
 *
 * 仕訳の各branch（借方/貸方）から、該当口座のaccount_id or sub_account_idに
 * マッチするものを抽出し、借方=入金、貸方=出金として扱う。
 *
 * @param {string} accountKey - 口座キー（CF005/CF003/SEIBU）
 * @param {string} dateFrom - 開始日（YYYY-MM-DD）
 * @param {string} dateTo - 終了日（YYYY-MM-DD）
 * @return {Array<Object>} 入出金データ配列
 */
function fetchWalletTransactions(accountKey, dateFrom, dateTo) {
  const walletMap = getWalletMap_();
  const walletInfo = walletMap[accountKey];

  if (!walletInfo) {
    throw new Error(`口座 ${accountKey} のIDが未設定です。MF連携を再実行してください。`);
  }

  // 仕訳を一括取得（全仕訳取得 → ローカルで口座フィルタ）
  const allTxns = [];
  let page = 1;
  const perPage = 500;

  Logger.log(`[${accountKey}] 口座マッチング情報: accountId=${walletInfo.accountId}, subAccountId=${walletInfo.subAccountId}, name=${walletInfo.name}`);

  while (true) {
    const data = mfApiRequest_('/journals', {
      start_date: dateFrom,
      end_date: dateTo,
      page: page,
      per_page: perPage
    });

    const journals = data.journals || [];
    Logger.log(`[${accountKey}] page ${page}: ${journals.length}件の仕訳取得`);
    if (journals.length === 0) break;

    journals.forEach(journal => {
      const branches = journal.branches || [];

      branches.forEach(branch => {
        const remark = branch.remark || '';
        const debitor = branch.debitor || {};
        const creditor = branch.creditor || {};

        // 借方（debitor）に該当口座がある → 入金（預金が増える）
        if (isMatchingAccount_(debitor, walletInfo)) {
          allTxns.push({
            date: new Date(journal.transaction_date),
            content: buildDescription_(remark, creditor),
            deposit: debitor.value || 0,
            withdrawal: 0,
            balance: 0,
            source: CF_CONFIG.SOURCE.MF
          });
        }

        // 貸方（creditor）に該当口座がある → 出金（預金が減る）
        if (isMatchingAccount_(creditor, walletInfo)) {
          allTxns.push({
            date: new Date(journal.transaction_date),
            content: buildDescription_(remark, debitor),
            deposit: 0,
            withdrawal: creditor.value || 0,
            balance: 0,
            source: CF_CONFIG.SOURCE.MF
          });
        }
      });
    });

    if (journals.length < perPage) break;
    page++;
    Utilities.sleep(300);
  }

  // 日付昇順ソート
  allTxns.sort((a, b) => a.date - b.date);

  Logger.log(`${CF_CONFIG.ACCOUNTS[accountKey].shortName}: ${allTxns.length}件取得 (${dateFrom} 〜 ${dateTo})`);
  return allTxns;
}

/**
 * 仕訳の借方/貸方が指定口座にマッチするか判定
 * @param {Object} side - debitor or creditor オブジェクト
 * @param {Object} walletInfo - { accountId, subAccountId, name }
 * @return {boolean}
 */
function isMatchingAccount_(side, walletInfo) {
  if (!side || !side.account_id) return false;

  // sub_account_idでマッチ（最も正確）
  if (walletInfo.subAccountId && side.sub_account_id) {
    const match = String(side.sub_account_id) === String(walletInfo.subAccountId);
    if (match) Logger.log(`  ✓ マッチ: ${side.sub_account_name || side.account_name} (subId一致)`);
    return match;
  }

  // account_id + sub_account_nameでマッチ
  if (walletInfo.accountId && String(side.account_id) === String(walletInfo.accountId)) {
    // sub_account_nameが設定されていれば名前でも確認
    if (walletInfo.subAccountName && side.sub_account_name) {
      return side.sub_account_name.includes(walletInfo.subAccountName)
          || walletInfo.subAccountName.includes(side.sub_account_name);
    }
    // sub_accountが不要な場合（accountレベルでマッチ）
    if (!walletInfo.subAccountId && !walletInfo.subAccountName) {
      return true;
    }
  }

  // 名前ベースのフォールバック
  const sideName = (side.sub_account_name || side.account_name || '');
  return sideName === walletInfo.name;
}

/**
 * 摘要（remark）と相手科目から説明文を生成
 */
function buildDescription_(remark, counterSide) {
  if (remark) return remark;
  if (counterSide) {
    const parts = [];
    if (counterSide.account_name) parts.push(counterSide.account_name);
    if (counterSide.sub_account_name) parts.push(counterSide.sub_account_name);
    return parts.join(' ');
  }
  return '';
}

/**
 * 全3口座の入出金明細を一括取得
 */
function fetchAllWalletTransactions(dateFrom, dateTo) {
  const result = {};
  Object.keys(CF_CONFIG.ACCOUNTS).forEach(key => {
    try {
      result[key] = fetchWalletTransactions(key, dateFrom, dateTo);
    } catch (e) {
      Logger.log(`⚠️ ${key} の取得に失敗: ${e.message}`);
      result[key] = [];
    }
  });
  return result;
}

// ==============================
// 仕訳データの取得（月次集計用）
// ==============================

/**
 * 仕訳データを取得する（カテゴリ分類に使用）
 * APIレスポンス: { journals: [{ branches: [{ debitor, creditor, remark }], transaction_date, ... }] }
 */
function fetchDeals(dateFrom, dateTo) {
  const allDeals = [];
  let page = 1;
  const perPage = 500;

  while (true) {
    const data = mfApiRequest_('/journals', {
      start_date: dateFrom,
      end_date: dateTo,
      page: page,
      per_page: perPage
    });

    const journals = data.journals || [];
    if (journals.length === 0) break;

    journals.forEach(journal => {
      const branches = journal.branches || [];
      branches.forEach(branch => {
        const debitor = branch.debitor || {};
        const creditor = branch.creditor || {};

        // 借方の情報
        if (debitor.account_name) {
          allDeals.push({
            date: new Date(journal.transaction_date),
            side: 'debit',
            accountItemName: debitor.account_name || '',
            subAccountName: debitor.sub_account_name || '',
            amount: debitor.value || 0,
            description: branch.remark || '',
            accountGroup: '', // accounts APIから別途取得可能
          });
        }

        // 貸方の情報
        if (creditor.account_name) {
          allDeals.push({
            date: new Date(journal.transaction_date),
            side: 'credit',
            accountItemName: creditor.account_name || '',
            subAccountName: creditor.sub_account_name || '',
            amount: creditor.value || 0,
            description: branch.remark || '',
            accountGroup: '',
          });
        }
      });
    });

    if (journals.length < perPage) break;
    page++;
    Utilities.sleep(300);
  }

  Logger.log(`仕訳データ: ${allDeals.length}件取得 (${dateFrom} 〜 ${dateTo})`);
  return allDeals;
}

// ==============================
// 口座残高の取得（残高試算表BS）
// ==============================

/**
 * 残高試算表（BS）から各口座の期末残高を取得する
 * GET /api/v3/reports/trial_balance_bs
 */
function fetchCurrentBalances() {
  const walletMap = getWalletMap_();
  const result = {};

  try {
    const data = mfApiRequest_('/reports/trial_balance_bs');
    const rows = data.rows || [];

    // 「現金及び預金合計」配下の各口座を探索
    const balanceMap = extractBalancesFromRows_(rows);

    for (const [accountKey, walletInfo] of Object.entries(walletMap)) {
      const balance = balanceMap[walletInfo.name] || balanceMap[walletInfo.subAccountName] || 0;
      result[accountKey] = {
        balance: balance,
        name: walletInfo.name,
        lastSyncedAt: new Date().toISOString()
      };
    }
  } catch (e) {
    Logger.log(`残高試算表取得エラー: ${e.message}`);
    // フォールバック: 残高0で返す
    for (const [accountKey, walletInfo] of Object.entries(walletMap)) {
      result[accountKey] = {
        balance: 0,
        name: walletInfo.name,
        lastSyncedAt: new Date().toISOString()
      };
    }
  }

  return result;
}

/**
 * 試算表のrows構造から口座名→期末残高のマップを再帰的に構築
 * columns: [opening_balance, debit_amount, credit_amount, closing_balance, ratio]
 * closing_balance = values[3]
 */
function extractBalancesFromRows_(rows) {
  const balanceMap = {};
  if (!rows) return balanceMap;

  rows.forEach(row => {
    if (row.type === 'account' || row.type === 'sub_account') {
      // closing_balance（期末残高）は values[3]
      const closingBalance = (row.values && row.values.length >= 4) ? row.values[3] : 0;
      if (row.name) {
        balanceMap[row.name] = closingBalance || 0;
      }
    }
    // 子行を再帰探索
    if (row.rows && row.rows.length > 0) {
      const childMap = extractBalancesFromRows_(row.rows);
      Object.assign(balanceMap, childMap);
    }
  });

  return balanceMap;
}

// ==============================
// カテゴリ自動分類
// ==============================

/**
 * 仕訳データの勘定科目・摘要からカテゴリを自動判定する
 */
function categorizeTransaction(accountItemName, description, type) {
  const categories = type === 'income'
    ? CF_CONFIG.INCOME_CATEGORIES
    : CF_CONFIG.EXPENSE_CATEGORIES;

  // 1. 勘定科目名で照合（優先）
  for (const cat of categories) {
    if (cat.accountItems && cat.accountItems.length > 0) {
      for (const ai of cat.accountItems) {
        if (accountItemName.includes(ai)) return cat.key;
      }
    }
  }

  // 2. 摘要キーワードで照合
  const searchText = `${accountItemName} ${description}`;
  for (const cat of categories) {
    if (cat.keywords && cat.keywords.length > 0) {
      for (const kw of cat.keywords) {
        if (searchText.includes(kw)) return cat.key;
      }
    }
  }

  return 'other';
}

/**
 * 月次の入出金をカテゴリ別に集計する
 */
function getMonthlyBreakdown(year, month) {
  const dateFrom = `${year}-${String(month).padStart(2, '0')}-01`;
  const lastDay = new Date(year, month, 0).getDate();
  const dateTo = `${year}-${String(month).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;

  const deals = fetchDeals(dateFrom, dateTo);

  const income = {};
  CF_CONFIG.INCOME_CATEGORIES.forEach(c => { income[c.key] = 0; });
  const expense = {};
  CF_CONFIG.EXPENSE_CATEGORIES.forEach(c => { expense[c.key] = 0; });

  let totalIncome = 0;
  let totalExpense = 0;

  deals.forEach(deal => {
    // 借方＝費用/資産増、貸方＝収益/負債増 として簡易分類
    const type = deal.side === 'debit' ? 'expense' : 'income';
    const category = categorizeTransaction(deal.accountItemName, deal.description, type);

    if (type === 'income') {
      income[category] = (income[category] || 0) + deal.amount;
      totalIncome += deal.amount;
    } else {
      expense[category] = (expense[category] || 0) + deal.amount;
      totalExpense += deal.amount;
    }
  });

  return {
    income, expense,
    totals: { income: totalIncome, expense: totalExpense, diff: totalIncome - totalExpense }
  };
}

// ==============================
// MF連携状態の確認
// ==============================

/**
 * MF連携状態を表示する
 */
function showMfStatus() {
  const ui = SpreadsheetApp.getUi();

  if (!isMfConnected()) {
    ui.alert('❌ マネーフォワード未連携\n\nメニューから「MF連携開始」を実行してください。');
    return;
  }

  try {
    // 事業者情報を取得
    const officeData = mfApiRequest_('/offices');
    const officeName = officeData.name || '不明';
    const periods = officeData.accounting_periods || [];
    const currentPeriod = periods.length > 0 ? periods[0] : {};

    let msg = `✅ マネーフォワード連携中\n\n`;
    msg += `【事業者】${officeName}\n`;
    msg += `【会計年度】${currentPeriod.fiscal_year || '?'}年度 (${currentPeriod.start_date || '?'} 〜 ${currentPeriod.end_date || '?'})\n\n`;

    // 口座残高を取得
    const balances = fetchCurrentBalances();
    msg += '【口座残高（試算表ベース）】\n';

    let totalBalance = 0;
    for (const [key, account] of Object.entries(CF_CONFIG.ACCOUNTS)) {
      const bal = balances[key];
      if (bal && bal.balance !== 0) {
        msg += `${account.shortName}: ¥${Number(bal.balance).toLocaleString()}\n`;
        totalBalance += bal.balance;
      } else {
        msg += `${account.shortName}: (データなし)\n`;
      }
    }

    msg += `\n【合計】¥${totalBalance.toLocaleString()}`;
    ui.alert(msg);
  } catch (e) {
    ui.alert('⚠️ MF接続エラー\n\n' + e.message);
  }
}
