/**
 * キャッシュフロー管理システム - マネーフォワードAPIクライアント
 *
 * MFクラウド会計から以下のデータを取得:
 *  - 口座の入出金明細（wallet_txns）→ Dailyシート
 *  - 仕訳データ（deals）→ 月次集計のカテゴリ分類
 *  - 口座残高（walletables）→ 現残高シート
 */

// ==============================
// 口座入出金明細の取得
// ==============================

/**
 * 指定口座の入出金明細をMFから取得する
 * @param {string} accountKey - 口座キー（CF005/CF003/SEIBU）
 * @param {string} dateFrom - 開始日（YYYY-MM-DD）
 * @param {string} dateTo - 終了日（YYYY-MM-DD）
 * @return {Array<Object>} 入出金データ配列
 */
function fetchWalletTransactions(accountKey, dateFrom, dateTo) {

  const walletMap = getWalletMap_();
  const walletInfo = walletMap[accountKey];

  if (!walletInfo) {
    throw new Error(`口座 ${accountKey} のwallet IDが未設定です。MF連携を再実行してください。`);
  }

  // 仕訳データ（journals）から該当口座の入出金を抽出
  const allTxns = [];
  let page = 1;
  const perPage = 100;

  while (true) {
    const data = mfApiRequest_(
      '/journals',
      {
        start_date: dateFrom,
        end_date: dateTo,
        page: page,
        per_page: perPage
      }
    );

    const journals = data.journals || [];
    if (journals.length === 0) break;

    journals.forEach(journal => {
      // 仕訳の借方・貸方から該当口座を抽出
      const entries = journal.entries || journal.journal_entries || [];
      entries.forEach(entry => {
        const acctId = String(entry.account_id || entry.sub_account_id || '');
        const acctName = entry.account_name || '';

        // 口座IDまたは口座名でマッチ
        if (acctId === walletInfo.id || acctName.includes(walletInfo.name)) {
          const isDebit = entry.side === 'debit';
          allTxns.push({
            date: new Date(journal.date || journal.journal_date),
            content: journal.description || journal.note || entry.description || '',
            deposit: isDebit ? (entry.amount || 0) : 0,
            withdrawal: !isDebit ? (entry.amount || 0) : 0,
            balance: 0,
            source: CF_CONFIG.SOURCE.MF
          });
        }
      });
    });

    if (journals.length < perPage) break;
    page++;
    Utilities.sleep(500);
  }

  // 日付昇順ソート
  allTxns.sort((a, b) => a.date - b.date);

  Logger.log(`${CF_CONFIG.ACCOUNTS[accountKey].shortName}: ${allTxns.length}件取得 (${dateFrom} 〜 ${dateTo})`);
  return allTxns;
}

/**
 * 全3口座の入出金明細を一括取得
 * @param {string} dateFrom - 開始日
 * @param {string} dateTo - 終了日
 * @return {Object} { CF005: [...], CF003: [...], SEIBU: [...] }
 */
function fetchAllWalletTransactions(dateFrom, dateTo) {
  const result = {};
  const accountKeys = Object.keys(CF_CONFIG.ACCOUNTS);

  accountKeys.forEach(key => {
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
 * @param {string} dateFrom - 開始日（YYYY-MM-DD）
 * @param {string} dateTo - 終了日（YYYY-MM-DD）
 * @return {Array<Object>} 仕訳データ配列
 */
function fetchDeals(dateFrom, dateTo) {

  const allDeals = [];
  let page = 1;
  const perPage = 100;

  while (true) {
    const data = mfApiRequest_(
      '/journals',
      {
        start_date: dateFrom,
        end_date: dateTo,
        page: page,
        per_page: perPage
      }
    );

    const deals = data.journals || [];
    if (deals.length === 0) break;

    deals.forEach(deal => {
      const details = deal.entries || deal.journal_entries || deal.details || [];
      details.forEach(detail => {
        allDeals.push({
          date: new Date(deal.date || deal.journal_date || deal.issue_date),
          type: deal.type,  // 'income' or 'expense'
          accountItemName: detail.account_item_name || '',
          taxName: detail.tax_name || '',
          amount: detail.amount || 0,
          description: detail.description || deal.ref_number || '',
          partnerName: (deal.partner || {}).name || '',
          walletableId: detail.walletable_id ? String(detail.walletable_id) : '',
          walletableName: detail.walletable_name || ''
        });
      });
    });

    if (deals.length < perPage) break;
    page++;
    Utilities.sleep(500);
  }

  Logger.log(`仕訳データ: ${allDeals.length}件取得 (${dateFrom} 〜 ${dateTo})`);
  return allDeals;
}

// ==============================
// 口座残高の取得
// ==============================

/**
 * 全口座の現在残高を取得する
 * @return {Object} { CF005: { balance, name }, CF003: {...}, SEIBU: {...} }
 */
function fetchCurrentBalances() {

  const walletMap = getWalletMap_();
  const result = {};

  // 勘定科目一覧から口座情報を取得
  const data = mfApiRequest_('/accounts');
  const accounts = data.accounts || [];

  for (const [accountKey, walletInfo] of Object.entries(walletMap)) {
    result[accountKey] = {
      balance: 0, // 残高はDailyシートから計算
      name: walletInfo.name,
      lastSyncedAt: new Date().toISOString()
    };
  }

  return result;
}

// ==============================
// カテゴリ自動分類
// ==============================

/**
 * 仕訳データの勘定科目・摘要からカテゴリを自動判定する
 * @param {string} accountItemName - 勘定科目名
 * @param {string} description - 摘要
 * @param {string} type - 'income' or 'expense'
 * @return {string} カテゴリキー
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

  // 3. マッチしない場合は「その他」
  return 'other';
}

/**
 * 月次の入出金をカテゴリ別に集計する
 * @param {number} year - 年
 * @param {number} month - 月
 * @return {Object} { income: { amazon: 0, ... }, expense: { saison: 0, ... }, totals: { income, expense, diff } }
 */
function getMonthlyBreakdown(year, month) {
  const dateFrom = `${year}-${String(month).padStart(2, '0')}-01`;
  const lastDay = new Date(year, month, 0).getDate();
  const dateTo = `${year}-${String(month).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;

  const deals = fetchDeals(dateFrom, dateTo);

  // カテゴリ別集計の初期化
  const income = {};
  CF_CONFIG.INCOME_CATEGORIES.forEach(c => { income[c.key] = 0; });
  const expense = {};
  CF_CONFIG.EXPENSE_CATEGORIES.forEach(c => { expense[c.key] = 0; });

  let totalIncome = 0;
  let totalExpense = 0;

  deals.forEach(deal => {
    const category = categorizeTransaction(deal.accountItemName, deal.description, deal.type);
    if (deal.type === 'income') {
      income[category] = (income[category] || 0) + deal.amount;
      totalIncome += deal.amount;
    } else {
      expense[category] = (expense[category] || 0) + deal.amount;
      totalExpense += deal.amount;
    }
  });

  return {
    income,
    expense,
    totals: {
      income: totalIncome,
      expense: totalExpense,
      diff: totalIncome - totalExpense
    }
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
    const balances = fetchCurrentBalances();
    const walletMap = getWalletMap_();

    let msg = '✅ マネーフォワード連携中\n\n【口座残高】\n';

    for (const [key, account] of Object.entries(CF_CONFIG.ACCOUNTS)) {
      const bal = balances[key];
      const wallet = walletMap[key];
      if (bal) {
        msg += `${account.shortName}: ¥${Number(bal.balance).toLocaleString()}\n`;
      } else if (wallet) {
        msg += `${account.shortName}: (残高取得待ち)\n`;
      } else {
        msg += `${account.shortName}: ⚠️ 未紐付け\n`;
      }
    }

    const total = Object.values(balances).reduce((sum, b) => sum + (b.balance || 0), 0);
    msg += `\n【3口座合計】¥${total.toLocaleString()}`;

    ui.alert(msg);
  } catch (e) {
    ui.alert('⚠️ MF接続エラー\n\n' + e.message);
  }
}
