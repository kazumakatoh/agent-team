/**
 * キャッシュフロー管理システム - 月別集計モジュール
 *
 * 「月別」シートに3口座の月次サマリーを蓄積表示する。
 * 年度をまたいでも下に追加していく（過去データは消えない）。
 *
 * ■ レイアウト（各月4行）
 *   A列: 年月(2025.04形式), B列: 金融機関, C列: 月初残高, D列: 入金, E列: 出金, F列: 差, G列: 月末残高
 *
 *   2025.04  全体         (3口座合計)
 *            PayPay 005   (口座別)
 *            PayPay 003   (口座別)
 *            西武信用金庫  (口座別)
 *   2025.05  全体
 *            ...
 */

// ==============================
// 月別シートの更新
// ==============================

/**
 * 月別シートを更新する
 * 指定した開始年月〜終了年月の範囲を更新（既存データは維持）
 */
function updateMonthlySheet(startYear, startMonth, endYear, endMonth) {
  const ss = getCfSpreadsheet();
  let sheet = ss.getSheetByName(CF_CONFIG.SHEETS.MONTHLY);

  if (!sheet) {
    sheet = ss.insertSheet(CF_CONFIG.SHEETS.MONTHLY);
  }

  // ヘッダーがなければ作成
  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() !== '年月') {
    const headers = ['年月', '金融機関', '月初残高', '入金', '出金', '差', '月末残高'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#1a73e8').setFontColor('#ffffff')
      .setFontWeight('bold').setHorizontalAlignment('center');
    sheet.setColumnWidth(1, 65);
    sheet.setColumnWidth(2, 120);
    for (let i = 3; i <= 7; i++) sheet.setColumnWidth(i, 110);
    sheet.setFrozenRows(1);
  }

  const accountKeys = Object.keys(CF_CONFIG.ACCOUNTS);
  const accountLabels = {
    CF005: CF_CONFIG.ACCOUNTS.CF005.shortName,
    CF003: CF_CONFIG.ACCOUNTS.CF003.shortName,
    SEIBU: CF_CONFIG.ACCOUNTS.SEIBU.shortName
  };

  // 各口座の前月繰越残高を取得
  const prevBalances = {};
  accountKeys.forEach(key => {
    prevBalances[key] = getCarryForwardBalance_(key);
  });

  // 開始月より前の月次データがあれば、そこから前月残高を引き継ぐ
  const firstMonthLabel = `${startYear}.${String(startMonth).padStart(2, '0')}`;
  const existingPrev = getPrevMonthClosing_(sheet, firstMonthLabel, accountKeys);
  if (existingPrev) {
    accountKeys.forEach(key => {
      if (existingPrev[key] !== undefined) prevBalances[key] = existingPrev[key];
    });
  }

  // 月ごとに処理
  let y = startYear, m = startMonth;
  while (y < endYear || (y === endYear && m <= endMonth)) {
    const monthLabel = `${y}.${String(m).padStart(2, '0')}`;

    // この月の行がすでにあるか確認
    let targetRow = findMonthlyRow_(sheet, monthLabel);

    if (targetRow === 0) {
      // なければ末尾に追加（4行分）
      targetRow = sheet.getLastRow() + 1;
      // 4行分の空行を確保
    }

    // 各口座の月次データを集計
    const accountMonthData = {};
    let totalIncome = 0;
    let totalExpense = 0;

    accountKeys.forEach(key => {
      const data = aggregateDailyForMonth_(key, y, m);
      accountMonthData[key] = data;
      totalIncome += data.totalIncome;
      totalExpense += data.totalExpense;
    });

    // 各口座の月初・月末残高
    const monthOpening = {};
    const monthClosing = {};

    accountKeys.forEach(key => {
      monthOpening[key] = prevBalances[key];
      monthClosing[key] = monthOpening[key] + accountMonthData[key].totalIncome - accountMonthData[key].totalExpense;
    });

    const totalOpening = accountKeys.reduce((s, k) => s + monthOpening[k], 0);
    const totalClosing = accountKeys.reduce((s, k) => s + monthClosing[k], 0);
    const totalDiff = totalIncome - totalExpense;

    // --- 全体行 ---
    sheet.getRange(targetRow, 1).setNumberFormat('@').setValue(monthLabel);
    sheet.getRange(targetRow, 2).setValue('全体');
    sheet.getRange(targetRow, 3).setValue(totalOpening).setNumberFormat('#,##0');
    sheet.getRange(targetRow, 4).setValue(totalIncome).setNumberFormat('#,##0');
    sheet.getRange(targetRow, 5).setValue(totalExpense).setNumberFormat('#,##0');
    sheet.getRange(targetRow, 6).setValue(totalDiff).setNumberFormat('#,##0');
    sheet.getRange(targetRow, 7).setValue(totalClosing).setNumberFormat('#,##0');
    sheet.getRange(targetRow, 1, 1, 7).setBackground('#e8eaf6').setFontWeight('bold');
    if (totalDiff < 0) sheet.getRange(targetRow, 6).setFontColor('#d32f2f');
    else sheet.getRange(targetRow, 6).setFontColor('#000000');

    // --- 口座別行 ---
    accountKeys.forEach((key, idx) => {
      const row = targetRow + 1 + idx;
      const data = accountMonthData[key];
      const diff = data.totalIncome - data.totalExpense;

      sheet.getRange(row, 2).setValue(accountLabels[key]);
      sheet.getRange(row, 3).setValue(monthOpening[key]).setNumberFormat('#,##0');
      sheet.getRange(row, 4).setValue(data.totalIncome).setNumberFormat('#,##0');
      sheet.getRange(row, 5).setValue(data.totalExpense).setNumberFormat('#,##0');
      sheet.getRange(row, 6).setValue(diff).setNumberFormat('#,##0');
      sheet.getRange(row, 7).setValue(monthClosing[key]).setNumberFormat('#,##0');
      if (diff < 0) sheet.getRange(row, 6).setFontColor('#d32f2f');
      else sheet.getRange(row, 6).setFontColor('#000000');
    });

    // 次月の月初 = 今月の月末
    accountKeys.forEach(key => {
      prevBalances[key] = monthClosing[key];
    });

    // 次月へ
    m++;
    if (m > 12) { m = 1; y++; }
  }

  Logger.log(`月別シート更新完了 (${startYear}.${startMonth} 〜 ${endYear}.${endMonth})`);
}

/**
 * 月別シートから指定月の開始行を検索
 * @return {number} 行番号（0=見つからない）
 */
function findMonthlyRow_(sheet, monthLabel) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;

  const colA = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < colA.length; i++) {
    if (String(colA[i][0]) === monthLabel) {
      return i + 2;
    }
  }
  return 0;
}

/**
 * 指定月の前月の月末残高を取得
 */
function getPrevMonthClosing_(sheet, monthLabel, accountKeys) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;

  // monthLabelの前月を計算
  const match = monthLabel.match(/^(\d{4})\.(\d{2})$/);
  if (!match) return null;
  let py = parseInt(match[1]);
  let pm = parseInt(match[2]) - 1;
  if (pm === 0) { pm = 12; py--; }
  const prevLabel = `${py}.${String(pm).padStart(2, '0')}`;

  const prevRow = findMonthlyRow_(sheet, prevLabel);
  if (prevRow === 0) return null;

  // 口座別行の月末残高（G列）を取得
  const result = {};
  accountKeys.forEach((key, idx) => {
    const bal = sheet.getRange(prevRow + 1 + idx, 7).getValue();
    result[key] = Number(bal) || 0;
  });

  return result;
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
// メニューから呼ばれる更新関数
// ==============================

/**
 * 月次集計を更新する（ダイアログで期間指定）
 */
function updateAllMonthlySheets() {
  const ui = SpreadsheetApp.getUi();

  const result = ui.prompt(
    '月次集計の更新',
    '更新する期間を入力してください\n\n' +
    '例: 2026/01-2026/04 （範囲指定）\n' +
    '例: 2026 （年全体）\n' +
    '例: ALL （全期間 2025.03〜現在）',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) return;

  const input = result.getResponseText().trim();
  let startYear, startMonth, endYear, endMonth;

  if (input.toUpperCase() === 'ALL') {
    startYear = 2025; startMonth = 3;
    const today = new Date();
    endYear = today.getFullYear();
    endMonth = today.getMonth() + 1;
  } else if (input.includes('-')) {
    // 範囲指定: 2026/01-2026/04
    const parts = input.split('-');
    const start = parts[0].trim().split(/[\/\.]/);
    const end = parts[1].trim().split(/[\/\.]/);
    startYear = parseInt(start[0]); startMonth = parseInt(start[1]);
    endYear = parseInt(end[0]); endMonth = parseInt(end[1]);
  } else {
    // 年指定: 2026
    const year = parseInt(input);
    if (isNaN(year)) {
      ui.alert('⚠️ 入力形式が正しくありません');
      return;
    }
    startYear = year; startMonth = 1;
    endYear = year; endMonth = 12;
  }

  updateMonthlySheet(startYear, startMonth, endYear, endMonth);

  ui.alert(`✅ ${startYear}.${String(startMonth).padStart(2,'0')} 〜 ${endYear}.${String(endMonth).padStart(2,'0')} の月次集計を更新しました。`);
}
