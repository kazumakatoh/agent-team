/**
 * キャッシュフロー管理システム - 月別集計モジュール
 *
 * 「月別」シートに3口座の月次サマリーを集約表示する。
 * CF005/CF003/西武信金の個別シートは廃止。
 *
 * ■ レイアウト（各月4行）
 *   A列: 月, B列: 金融機関, C列: 月初残高, D列: 入金, E列: 出金, F列: 差, G列: 月末残高
 *
 *   1月  全体         (3口座合計)
 *        PayPay 005   (口座別)
 *        PayPay 003   (口座別)
 *        西武信用金庫  (口座別)
 *   2月  全体
 *        ...
 */

// ==============================
// 月別シートの更新
// ==============================

/**
 * 月別シートを更新する
 * @param {number} year - 対象年
 */
function updateMonthlySheet(year) {
  const ss = getCfSpreadsheet();
  const sheet = ss.getSheetByName(CF_CONFIG.SHEETS.MONTHLY);
  if (!sheet) throw new Error(`シート「${CF_CONFIG.SHEETS.MONTHLY}」が見つかりません。`);

  // ヘッダー
  const headers = ['', '金融機関', '月初残高', '入金', '出金', '差', '月末残高'];

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1a73e8').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');

  const accountKeys = Object.keys(CF_CONFIG.ACCOUNTS);
  const accountLabels = {
    CF005: CF_CONFIG.ACCOUNTS.CF005.shortName,
    CF003: CF_CONFIG.ACCOUNTS.CF003.shortName,
    SEIBU: CF_CONFIG.ACCOUNTS.SEIBU.shortName
  };

  // 各口座の前月繰越残高をDailyシートから取得
  const prevBalances = {};
  accountKeys.forEach(key => {
    prevBalances[key] = getCarryForwardBalance_(key);
  });

  let currentRow = 2;

  for (let m = 1; m <= 12; m++) {
    // 各口座の月次データを集計
    const accountMonthData = {};
    let totalIncome = 0;
    let totalExpense = 0;

    accountKeys.forEach(key => {
      const data = aggregateDailyForMonth_(key, year, m);
      accountMonthData[key] = data;
      totalIncome += data.totalIncome;
      totalExpense += data.totalExpense;
    });

    const totalDiff = totalIncome - totalExpense;

    // 各口座の月初・月末残高を計算
    const monthOpening = {};
    const monthClosing = {};

    accountKeys.forEach(key => {
      monthOpening[key] = prevBalances[key];
      const data = accountMonthData[key];
      monthClosing[key] = monthOpening[key] + data.totalIncome - data.totalExpense;
    });

    const totalOpening = accountKeys.reduce((s, k) => s + monthOpening[k], 0);
    const totalClosing = accountKeys.reduce((s, k) => s + monthClosing[k], 0);

    // --- 全体行 ---
    sheet.getRange(currentRow, 1).setValue(`${m}月`);
    sheet.getRange(currentRow, 2).setValue('全体');
    sheet.getRange(currentRow, 3).setValue(totalOpening).setNumberFormat('#,##0');
    sheet.getRange(currentRow, 4).setValue(totalIncome).setNumberFormat('#,##0');
    sheet.getRange(currentRow, 5).setValue(totalExpense).setNumberFormat('#,##0');
    sheet.getRange(currentRow, 6).setValue(totalDiff).setNumberFormat('#,##0');
    sheet.getRange(currentRow, 7).setValue(totalClosing).setNumberFormat('#,##0');

    // 全体行のスタイル
    sheet.getRange(currentRow, 1, 1, headers.length)
      .setBackground('#e8eaf6').setFontWeight('bold');
    if (totalDiff < 0) sheet.getRange(currentRow, 6).setFontColor('#d32f2f');

    currentRow++;

    // --- 口座別行 ---
    accountKeys.forEach(key => {
      const data = accountMonthData[key];
      const diff = data.totalIncome - data.totalExpense;

      sheet.getRange(currentRow, 2).setValue(accountLabels[key]);
      sheet.getRange(currentRow, 3).setValue(monthOpening[key]).setNumberFormat('#,##0');
      sheet.getRange(currentRow, 4).setValue(data.totalIncome).setNumberFormat('#,##0');
      sheet.getRange(currentRow, 5).setValue(data.totalExpense).setNumberFormat('#,##0');
      sheet.getRange(currentRow, 6).setValue(diff).setNumberFormat('#,##0');
      sheet.getRange(currentRow, 7).setValue(monthClosing[key]).setNumberFormat('#,##0');

      if (diff < 0) sheet.getRange(currentRow, 6).setFontColor('#d32f2f');

      currentRow++;
    });

    // 次月の月初 = 今月の月末
    accountKeys.forEach(key => {
      prevBalances[key] = monthClosing[key];
    });
  }

  // 列幅
  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 120);
  for (let i = 3; i <= headers.length; i++) sheet.setColumnWidth(i, 110);

  sheet.setFrozenRows(1);

  Logger.log(`月別シート更新完了 (${year}年)`);
}

// ==============================
// Dailyシートからのデータ集計
// ==============================

/**
 * Dailyシートから指定口座・月のデータを集計する
 */
function aggregateDailyForMonth_(accountKey, year, month) {
  const ss = getCfSpreadsheet();
  const sheetName = CF_CONFIG.ACCOUNTS[accountKey].dailySheet;
  const sheet = ss.getSheetByName(sheetName);

  const result = { totalIncome: 0, totalExpense: 0 };

  const C = CF_CONFIG.DAILY_COLS;
  const headerRows = CF_CONFIG.DAILY_HEADER_ROWS;
  if (!sheet || sheet.getLastRow() <= headerRows) return result;

  const numRows = sheet.getLastRow() - headerRows;

  const dates = sheet.getRange(headerRows + 1, C.DATE, numRows, 1).getValues();
  const deposits = sheet.getRange(headerRows + 1, C.DEPOSIT, numRows, 1).getValues();
  const withdrawals = sheet.getRange(headerRows + 1, C.WITHDRAWAL, numRows, 1).getValues();

  for (let i = 0; i < numRows; i++) {
    const d = dates[i][0];
    if (!(d instanceof Date)) continue;
    if (d.getFullYear() !== year || d.getMonth() + 1 !== month) continue;

    result.totalIncome += Number(deposits[i][0]) || 0;
    result.totalExpense += Number(withdrawals[i][0]) || 0;
  }

  return result;
}

// ヘルパー関数 getCarryForwardBalance_, getLatestBalance_ は DailySheet.gs に定義済み

// ==============================
// 全シート一括更新
// ==============================

/**
 * 月次集計を更新する
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

  updateMonthlySheet(year);

  ui.alert(`✅ ${year}年の月次集計を更新しました。`);
}
