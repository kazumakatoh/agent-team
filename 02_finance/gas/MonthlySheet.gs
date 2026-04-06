/**
 * キャッシュフロー管理システム - 月次集計モジュール
 *
 * ■ 口座別月次シート（CF005 / CF003 / 西武信金）
 *   月初残高・入金・出金・差・累計差・月末残高をカテゴリ別に集計
 *
 * ■ 月別シート（3口座合算サマリー）
 *   入金合計・出金合計・差・各口座残高・合計残高
 */

// ==============================
// 口座別月次シートの更新
// ==============================

/**
 * 口座別月次シートを更新する
 * @param {string} accountKey - 口座キー（CF005/CF003/SEIBU）
 * @param {number} year - 対象年
 */
function updateAccountMonthlySheet(accountKey, year) {
  const ss = getCfSpreadsheet();
  const account = CF_CONFIG.ACCOUNTS[accountKey];
  const sheet = ss.getSheetByName(account.sheetName);
  if (!sheet) throw new Error(`シート「${account.sheetName}」が見つかりません。`);

  // ヘッダー構成
  const rowLabels = [
    { key: 'opening',  label: '月初預金残高', type: 'header' },
    { key: 'income',   label: '入金', type: 'total' },
    { key: 'expense',  label: '出金', type: 'total' },
    { key: 'diff',     label: '差', type: 'calc' },
    { key: 'cumDiff',  label: '累計差', type: 'calc' },
    { key: 'closing',  label: '月末預金残高', type: 'header' },
    { key: 'sep1',     label: '', type: 'separator' }
  ];

  // 入金カテゴリ行を追加
  CF_CONFIG.INCOME_CATEGORIES.forEach(cat => {
    rowLabels.push({ key: `inc_${cat.key}`, label: cat.label, type: 'income_detail' });
  });

  rowLabels.push({ key: 'sep2', label: '', type: 'separator' });

  // 出金カテゴリ行を追加
  CF_CONFIG.EXPENSE_CATEGORIES.forEach(cat => {
    rowLabels.push({ key: `exp_${cat.key}`, label: cat.label, type: 'expense_detail' });
  });

  // 列: A=ラベル, B=カテゴリ, C〜N=月（1月〜12月）
  sheet.clearContents();

  // ヘッダー行
  sheet.getRange(1, 1).setValue(year + '年');
  sheet.getRange(1, 1).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');

  // 月の列ヘッダー
  for (let m = 1; m <= 12; m++) {
    sheet.getRange(1, m + 1).setValue(m + '月');
    sheet.getRange(1, m + 1).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    sheet.getRange(1, m + 1).setHorizontalAlignment('center');
  }

  // 行ラベルを書き込み
  rowLabels.forEach((row, i) => {
    const r = i + 2;
    sheet.getRange(r, 1).setValue(row.label);

    // スタイル
    if (row.type === 'header') {
      sheet.getRange(r, 1, 1, 13).setBackground('#e8eaf6').setFontWeight('bold');
    } else if (row.type === 'total') {
      sheet.getRange(r, 1, 1, 13).setFontWeight('bold');
    } else if (row.type === 'calc') {
      sheet.getRange(r, 1, 1, 13).setBackground('#f3f3f3');
    }
  });

  // 月別データを取得して書き込む
  // 前月繰越残高をDailyシートから取得（前月繰越行の残高）
  let prevClosing = getCarryForwardBalance_(accountKey);
  let cumDiff = 0;

  for (let m = 1; m <= 12; m++) {
    const col = m + 1;
    const dateFrom = `${year}-${String(m).padStart(2, '0')}-01`;
    const lastDay = new Date(year, m, 0).getDate();
    const dateTo = `${year}-${String(m).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;

    // Dailyシートから当該口座・当該月のデータを集計
    const monthData = aggregateDailyForMonth_(accountKey, year, m);

    // 月初残高
    const opening = prevClosing;
    const totalIncome = monthData.totalIncome;
    const totalExpense = monthData.totalExpense;
    const diff = totalIncome - totalExpense;
    cumDiff += diff;
    const closing = opening + diff;

    let rowIdx = 2; // データ開始行

    // 月初預金残高
    sheet.getRange(rowIdx, col).setValue(opening).setNumberFormat('#,##0');
    rowIdx++;

    // 入金合計
    sheet.getRange(rowIdx, col).setValue(totalIncome).setNumberFormat('#,##0');
    rowIdx++;

    // 出金合計
    sheet.getRange(rowIdx, col).setValue(totalExpense).setNumberFormat('#,##0');
    rowIdx++;

    // 差
    sheet.getRange(rowIdx, col).setValue(diff).setNumberFormat('#,##0');
    if (diff < 0) sheet.getRange(rowIdx, col).setFontColor('#d32f2f');
    else sheet.getRange(rowIdx, col).setFontColor('#000000');
    rowIdx++;

    // 累計差
    sheet.getRange(rowIdx, col).setValue(cumDiff).setNumberFormat('#,##0');
    rowIdx++;

    // 月末預金残高
    sheet.getRange(rowIdx, col).setValue(closing).setNumberFormat('#,##0');
    sheet.getRange(rowIdx, col).setFontWeight('bold');
    rowIdx++;

    rowIdx++; // separator

    // 入金カテゴリ別
    CF_CONFIG.INCOME_CATEGORIES.forEach(cat => {
      const val = monthData.incomeByCategory[cat.key] || 0;
      if (val > 0) sheet.getRange(rowIdx, col).setValue(val).setNumberFormat('#,##0');
      rowIdx++;
    });

    rowIdx++; // separator

    // 出金カテゴリ別
    CF_CONFIG.EXPENSE_CATEGORIES.forEach(cat => {
      const val = monthData.expenseByCategory[cat.key] || 0;
      if (val > 0) sheet.getRange(rowIdx, col).setValue(val).setNumberFormat('#,##0');
      rowIdx++;
    });

    prevClosing = closing;
  }

  // 列幅調整
  sheet.setColumnWidth(1, 140);
  for (let m = 1; m <= 12; m++) {
    sheet.setColumnWidth(m + 1, 110);
  }

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);

  Logger.log(`${account.sheetName}シート更新完了 (${year}年)`);
}

/**
 * Dailyシートから指定口座・月のデータを集計する
 * @param {string} accountKey
 * @param {number} year
 * @param {number} month
 * @return {Object}
 */
function aggregateDailyForMonth_(accountKey, year, month) {
  const ss = getCfSpreadsheet();
  const sheet = ss.getSheetByName(CF_CONFIG.SHEETS.DAILY);

  const result = {
    totalIncome: 0,
    totalExpense: 0,
    incomeByCategory: {},
    expenseByCategory: {}
  };

  // カテゴリ初期化
  CF_CONFIG.INCOME_CATEGORIES.forEach(c => { result.incomeByCategory[c.key] = 0; });
  CF_CONFIG.EXPENSE_CATEGORIES.forEach(c => { result.expenseByCategory[c.key] = 0; });

  if (!sheet || sheet.getLastRow() <= CF_CONFIG.DAILY_HEADER_ROWS) return result;

  const cols = CF_CONFIG.ACCOUNTS[accountKey].daily;
  const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;
  const numRows = sheet.getLastRow() - headerRows;

  const dates = sheet.getRange(headerRows + 1, cols.DATE, numRows, 1).getValues();
  const contents = sheet.getRange(headerRows + 1, cols.CONTENT, numRows, 1).getValues();
  const deposits = sheet.getRange(headerRows + 1, cols.DEPOSIT, numRows, 1).getValues();
  const withdrawals = sheet.getRange(headerRows + 1, cols.WITHDRAWAL, numRows, 1).getValues();

  for (let i = 0; i < numRows; i++) {
    const d = dates[i][0];
    if (!(d instanceof Date)) continue;

    if (d.getFullYear() !== year || d.getMonth() + 1 !== month) continue;

    const deposit = Number(deposits[i][0]) || 0;
    const withdrawal = Number(withdrawals[i][0]) || 0;
    const content = String(contents[i][0]);

    if (deposit > 0) {
      result.totalIncome += deposit;
      const cat = categorizeByKeyword_(content, 'income');
      result.incomeByCategory[cat] += deposit;
    }

    if (withdrawal > 0) {
      result.totalExpense += withdrawal;
      const cat = categorizeByKeyword_(content, 'expense');
      result.expenseByCategory[cat] += withdrawal;
    }
  }

  return result;
}

/**
 * 摘要キーワードでカテゴリを判定する
 */
function categorizeByKeyword_(content, type) {
  const categories = type === 'income'
    ? CF_CONFIG.INCOME_CATEGORIES
    : CF_CONFIG.EXPENSE_CATEGORIES;

  for (const cat of categories) {
    for (const kw of cat.keywords) {
      if (content.includes(kw)) return cat.key;
    }
  }
  return 'other';
}

// ==============================
// 月別シート（3口座合算サマリー）
// ==============================

/**
 * 月別シート（3口座合算）を更新する
 * @param {number} year - 対象年
 */
function updateConsolidatedMonthlySheet(year) {
  const ss = getCfSpreadsheet();
  const sheet = ss.getSheetByName(CF_CONFIG.SHEETS.MONTHLY);
  if (!sheet) throw new Error(`シート「${CF_CONFIG.SHEETS.MONTHLY}」が見つかりません。`);

  // ヘッダー
  const headers = ['', '入金', '出金', '差',
    'PayPay銀行\n法人口座', 'PayPay銀行\n個人口座', '西武信用金庫', '残高'];

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1a73e8').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center')
    .setWrap(true);

  // 月別データ
  for (let m = 1; m <= 12; m++) {
    const row = m + 1;
    const label = `${m}月`;

    // 各口座のDailyデータを集計
    let totalIncome = 0;
    let totalExpense = 0;
    const accountBalances = {};

    Object.keys(CF_CONFIG.ACCOUNTS).forEach(key => {
      const monthData = aggregateDailyForMonth_(key, year, m);
      totalIncome += monthData.totalIncome;
      totalExpense += monthData.totalExpense;

      // 月末残高を取得
      accountBalances[key] = getMonthEndBalance_(key, year, m);
    });

    const diff = totalIncome - totalExpense;
    const totalBalance = Object.values(accountBalances).reduce((sum, b) => sum + b, 0);

    sheet.getRange(row, 1).setValue(label).setFontWeight('bold');
    sheet.getRange(row, 2).setValue(totalIncome).setNumberFormat('#,##0');
    sheet.getRange(row, 3).setValue(totalExpense).setNumberFormat('#,##0');
    sheet.getRange(row, 4).setValue(diff).setNumberFormat('#,##0');
    sheet.getRange(row, 5).setValue(accountBalances.CF005 || 0).setNumberFormat('#,##0');
    sheet.getRange(row, 6).setValue(accountBalances.CF003 || 0).setNumberFormat('#,##0');
    sheet.getRange(row, 7).setValue(accountBalances.SEIBU || 0).setNumberFormat('#,##0');
    sheet.getRange(row, 8).setValue(totalBalance).setNumberFormat('#,##0');

    // 差がマイナスなら赤色
    if (diff < 0) sheet.getRange(row, 4).setFontColor('#d32f2f');

    // 交互色
    if (m % 2 === 0) sheet.getRange(row, 1, 1, headers.length).setBackground('#f5f5f5');
  }

  // 列幅
  sheet.setColumnWidth(1, 60);
  for (let i = 2; i <= headers.length; i++) sheet.setColumnWidth(i, 110);

  sheet.setFrozenRows(1);

  Logger.log(`月別シート更新完了 (${year}年)`);
}

/**
 * 指定口座の月末残高をDailyシートから取得
 */
function getMonthEndBalance_(accountKey, year, month) {
  const ss = getCfSpreadsheet();
  const sheet = ss.getSheetByName(CF_CONFIG.SHEETS.DAILY);
  if (!sheet) return 0;

  const cols = CF_CONFIG.ACCOUNTS[accountKey].daily;
  const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) return 0;

  const numRows = lastRow - headerRows;
  const dates = sheet.getRange(headerRows + 1, cols.DATE, numRows, 1).getValues();
  const balances = sheet.getRange(headerRows + 1, cols.BALANCE, numRows, 1).getValues();

  // 当該月の最後の残高を返す
  let lastBalance = 0;
  for (let i = 0; i < numRows; i++) {
    const d = dates[i][0];
    if (!(d instanceof Date)) continue;
    if (d.getFullYear() === year && d.getMonth() + 1 === month) {
      const bal = Number(balances[i][0]) || 0;
      if (bal !== 0) lastBalance = bal;
    }
  }

  return lastBalance;
}

// ==============================
// 全シート一括更新
// ==============================

/**
 * 全月次シートを一括更新する
 */
function updateAllMonthlySheets() {
  const ui = SpreadsheetApp.getUi();
  const today = new Date();
  const defaultYear = today.getFullYear();

  const result = ui.prompt(
    '月次集計の更新',
    `集計する年を入力してください（デフォルト: ${defaultYear}）`,
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) return;

  const year = parseInt(result.getResponseText()) || defaultYear;

  // 口座別シートを更新
  Object.keys(CF_CONFIG.ACCOUNTS).forEach(key => {
    updateAccountMonthlySheet(key, year);
  });

  // 月別合算シートを更新
  updateConsolidatedMonthlySheet(year);

  ui.alert(`✅ ${year}年の月次集計を更新しました。`);
}

/**
 * Dailyシートから前月繰越残高を取得する
 * 最初の行（入出金なし、残高のみ）を前月繰越として返す
 * @param {string} accountKey
 * @return {number}
 */
function getCarryForwardBalance_(accountKey) {
  const ss = getCfSpreadsheet();
  const sheet = ss.getSheetByName(CF_CONFIG.SHEETS.DAILY);
  if (!sheet) return 0;

  const cols = CF_CONFIG.ACCOUNTS[accountKey].daily;
  const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) return 0;

  // 最初のデータ行を確認
  const firstBalance = sheet.getRange(headerRows + 1, cols.BALANCE).getValue();
  const firstDeposit = sheet.getRange(headerRows + 1, cols.DEPOSIT).getValue();
  const firstWithdrawal = sheet.getRange(headerRows + 1, cols.WITHDRAWAL).getValue();

  // 入出金がなく残高だけある = 前月繰越行
  if ((!firstDeposit || firstDeposit === 0) && (!firstWithdrawal || firstWithdrawal === 0) && firstBalance > 0) {
    return Number(firstBalance);
  }

  return 0;
}
