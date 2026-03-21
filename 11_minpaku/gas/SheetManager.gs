/**
 * 民泊自動化システム - スプレッドシート管理モジュール
 * 予約データの読み書き、シート初期化を担当
 */

/**
 * スプレッドシートを取得する
 * @return {Spreadsheet}
 */
function getSpreadsheet() {
  if (CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * 予約データをスプレッドシートに書き込む（重複チェックあり）
 * @param {Array} reservations - parseEmail で取得した予約データ配列
 * @return {number} 追加件数
 */
function writeReservations(reservations) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.RESERVATIONS);
  if (!sheet) throw new Error(`シート「${CONFIG.SHEETS.RESERVATIONS}」が見つかりません。Setupを実行してください。`);

  // 既存データを取得（EmailID重複防止 & 予約ID更新用）
  const lastRow = sheet.getLastRow();
  const existingEmailIds = new Set();
  const reservationIdToRow = {}; // 予約ID → 行番号
  const C = CONFIG.RESERVATION_COLS;
  const NUM_COLS = 19;
  if (lastRow > 1) {
    const allData = sheet.getRange(2, 1, lastRow - 1, NUM_COLS).getValues();
    allData.forEach((row, i) => {
      if (row[C.EMAIL_ID - 1]) existingEmailIds.add(String(row[C.EMAIL_ID - 1]));
      if (row[C.ID - 1]) reservationIdToRow[String(row[C.ID - 1])] = i + 2; // 行番号(1始まり)
    });
  }

  let added = 0;

  reservations.forEach(r => {
    if (existingEmailIds.has(r.emailId)) {
      Logger.log(`スキップ（重複）: ${r.reservationId}`);
      return;
    }

    const nights    = r.nights  || 0;
    const guests    = r.guests  || 1;
    const usageDays = r.usageDays   != null ? r.usageDays   : nights + 1;
    const totalGuests = r.totalGuests != null ? r.totalGuests : usageDays * guests;

    const rowData = new Array(NUM_COLS).fill('');
    rowData[C.ID            - 1] = r.reservationId   || '';
    rowData[C.PLATFORM      - 1] = r.platform         || '';
    rowData[C.BOOKED_DATE   - 1] = r.bookedDate       || new Date();
    rowData[C.CHECKIN       - 1] = r.checkin          || '';
    rowData[C.CHECKOUT      - 1] = r.checkout         || '';
    rowData[C.NIGHTS        - 1] = nights;
    rowData[C.GUESTS        - 1] = guests;
    rowData[C.USAGE_DAYS    - 1] = usageDays;
    rowData[C.TOTAL_GUESTS  - 1] = totalGuests;
    rowData[C.GUEST_NAME    - 1] = r.guestName        || '';
    rowData[C.REVENUE       - 1] = r.revenue          || 0;
    rowData[C.ACCOMMODATION - 1] = r.accommodationFee || 0;
    rowData[C.CLEANING_FEE  - 1] = r.cleaningFee      || 0;
    rowData[C.OTA_FEE       - 1] = r.otaFee           || 0;
    rowData[C.TRANSFER_FEE  - 1] = r.transferFee      || 0;
    rowData[C.PAYOUT        - 1] = r.payoutAmount      || 0;
    rowData[C.STATUS        - 1] = r.status           || '予約';
    rowData[C.NOTES         - 1] = r.notes            || '';
    rowData[C.EMAIL_ID      - 1] = r.emailId          || '';

    // 同じ予約IDが既存行にある場合は更新（変更メール対応）
    const existingRow = reservationIdToRow[String(r.reservationId)];
    if (existingRow && r.status === '変更') {
      sheet.getRange(existingRow, 1, 1, NUM_COLS).setValues([rowData]);
      Logger.log(`予約更新（変更）: ${r.reservationId}`);
    } else {
      sheet.appendRow(rowData);
      Logger.log(`予約追加: ${r.reservationId} (${r.platform})`);
    }

    existingEmailIds.add(r.emailId);
    added++;
  });

  // 日付列のフォーマット
  if (added > 0) {
    formatReservationSheet_(sheet);
  }

  return added;
}

/**
 * 予約IDでステータスを更新する（キャンセル処理用）
 * @param {string} reservationId - 予約ID
 * @param {string} newStatus - 新しいステータス（例: 'キャンセル'）
 * @return {boolean} 更新成功した場合true
 */
function updateReservationStatus(reservationId, newStatus) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.RESERVATIONS);
  if (!sheet || sheet.getLastRow() <= 1) return false;

  const C    = CONFIG.RESERVATION_COLS;
  const data = sheet.getRange(2, C.ID, sheet.getLastRow() - 1, 1).getValues();

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(reservationId)) {
      sheet.getRange(i + 2, C.STATUS).setValue(newStatus);
      Logger.log(`ステータス更新: ${reservationId} → ${newStatus}`);
      return true;
    }
  }

  Logger.log(`予約ID未発見: ${reservationId}`);
  return false;
}

/**
 * 予約リストから月別集計データを取得する
 * @param {number} year  - 集計対象年（例: 2025）
 * @param {number} month - 集計対象月（例: 12）
 * @return {Object} 月別集計データ
 */
function getMonthlyReservationData(year, month) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.RESERVATIONS);
  if (!sheet || sheet.getLastRow() <= 1) return buildEmptyMonthData_(year, month);

  const C    = CONFIG.RESERVATION_COLS;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 19).getValues();

  let revenue      = 0;
  let otaFee       = 0;
  let transferFee  = 0;
  let payout       = 0;
  let cleaningFee  = 0;
  let guests       = 0;
  let usageDays    = 0;
  let bookingCount = 0;

  data.forEach(row => {
    const checkin  = row[C.CHECKIN  - 1] ? new Date(row[C.CHECKIN  - 1]) : null;
    const checkout = row[C.CHECKOUT - 1] ? new Date(row[C.CHECKOUT - 1]) : null;
    const status   = row[C.STATUS   - 1];

    if (!checkin || !checkout || status === 'キャンセル') return;

    // チェックインが当月に含まれる予約をカウント
    const checkinMonth = checkin.getMonth() + 1;
    const checkinYear  = checkin.getFullYear();

    if (checkinYear !== year || checkinMonth !== month) return;

    bookingCount++;
    revenue     += Number(row[C.REVENUE      - 1]) || 0;
    otaFee      += Number(row[C.OTA_FEE      - 1]) || 0;
    transferFee += Number(row[C.TRANSFER_FEE - 1]) || 0;
    payout      += Number(row[C.PAYOUT       - 1]) || 0;
    cleaningFee += Number(row[C.CLEANING_FEE - 1]) || 0;
    guests      += Number(row[C.GUESTS       - 1]) || 0;
    usageDays   += Number(row[C.NIGHTS       - 1]) || 0;
  });

  return {
    year,
    month,
    bookingCount,
    usageDays,
    guests,
    revenue,
    otaFee,
    transferFee,
    payout,
    cleaningFee  // ゲスト負担分（参考値。費用は経費入力シートで管理）
  };
}

/**
 * 経費入力シートから指定月の経費データを取得する
 * @param {number} year
 * @param {number} month
 * @return {Object}
 */
function getMonthlyCostData(year, month) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.COSTS);
  if (!sheet || sheet.getLastRow() <= 1) {
    return { cleaning: 0, supplies: 0, utilities: 0, rent: 0, other: 0 };
  }

  const yearMonth = `${year}-${String(month).padStart(2, '0')}`;
  const C = CONFIG.COST_COLS;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();

  const row = data.find(r => String(r[C.YEAR_MONTH - 1]).startsWith(yearMonth));
  if (!row) {
    return { cleaning: 0, supplies: 0, utilities: 0, rent: 0, other: 0 };
  }

  return {
    cleaning:  Number(row[C.CLEANING   - 1]) || 0,
    supplies:  Number(row[C.SUPPLIES   - 1]) || 0,
    utilities: Number(row[C.UTILITIES  - 1]) || 0,
    rent:      Number(row[C.RENT       - 1]) || 0,
    other:     Number(row[C.OTHER      - 1]) || 0
  };
}

/**
 * 月別集計シートを更新する
 * @param {number} fiscalYear - 事業年度（開始年。例: 2025なら2025年4月〜2026年3月）
 */
function updateMonthlySheet(fiscalYear) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.MONTHLY);
  if (!sheet) throw new Error(`シート「${CONFIG.SHEETS.MONTHLY}」が見つかりません。`);

  // ヘッダー設定
  const headers = [
    '年月', '問い合わせ数', '稼働日数', '利用可能日数',
    '利用件数', '利用人数', '売上', '手数料', '清掃費',
    '備品・消耗品費', '水光熱費', '家賃', 'その他経費', '総経費', '利益',
    'ROI(%)', 'ADR(円)', 'RevPAR(円)', '稼働率(%)'
  ];

  // 既存の問い合わせ数を保持（手動入力値を上書きしないため）
  const existingInquiries = {};
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const existing = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    existing.forEach(row => {
      if (row[0]) existingInquiries[String(row[0])] = Number(row[1]) || 0;
    });
  }

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sheet, headers.length);

  // 4月〜翌3月の12ヶ月
  const months = getFiscalYearMonths_(fiscalYear);
  let totalUsageDays = 0;
  const rows = [];

  months.forEach(({ year, month }) => {
    const resData  = getMonthlyReservationData(year, month);
    const costData = getMonthlyCostData(year, month);
    const daysInMonth = new Date(year, month, 0).getDate();

    // 清掃費（経費）は経費入力シートの値のみ（予約リストの清掃費はゲスト負担の売上）
    const totalCleaning = costData.cleaning;
    const totalCosts    = totalCleaning + costData.supplies + costData.utilities + costData.rent + costData.other;
    const netRevenue    = resData.payout || (resData.revenue - resData.otaFee - resData.transferFee);
    const profit        = netRevenue - totalCosts;

    const kpis = KPICalculator.calcMonthlyKPIs({
      revenue:    resData.revenue,
      commission: resData.otaFee,
      usageDays:  resData.usageDays,
      daysInMonth,
      totalCosts
    });

    totalUsageDays += resData.usageDays;

    // 問い合わせ数：既存の手動入力値を復元（なければ0）
    const yearMonthKey = `${year}/${String(month).padStart(2, '0')}`;
    const inquiries = existingInquiries[yearMonthKey] || 0;

    rows.push([
      yearMonthKey,
      inquiries,              // 問い合わせ数（手動入力値を保持）
      resData.usageDays,
      daysInMonth,
      resData.bookingCount,
      resData.guests,
      resData.revenue,
      resData.otaFee,
      totalCleaning,
      costData.supplies,
      costData.utilities,
      costData.rent,
      costData.other,
      totalCosts,
      profit,
      kpis.roi,
      kpis.adr,
      kpis.revpar,
      kpis.occupancyRate
    ]);
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    formatMonthlySheet_(sheet, rows.length, headers.length);
  }

  // 合計行
  appendMonthlyTotalRow_(sheet, rows.length, headers.length, totalUsageDays, fiscalYear);

  Logger.log(`月別集計シート更新完了 (${fiscalYear}年度)`);
}

// ==============================
// プライベート関数
// ==============================

function getFiscalYearMonths_(fiscalYear) {
  const months = [];
  // 4月〜12月 (fiscalYear)
  for (let m = 4; m <= 12; m++) {
    months.push({ year: fiscalYear, month: m });
  }
  // 1月〜3月 (fiscalYear+1)
  for (let m = 1; m <= 3; m++) {
    months.push({ year: fiscalYear + 1, month: m });
  }
  return months;
}

function buildEmptyMonthData_(year, month) {
  return {
    year, month,
    bookingCount: 0, usageDays: 0, guests: 0,
    revenue: 0, commission: 0, cleaningFee: 0
  };
}

function styleHeaderRow_(sheet, numCols) {
  const headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange.setBackground('#1a73e8')
             .setFontColor('#ffffff')
             .setFontWeight('bold')
             .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
}

function formatReservationSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const C = CONFIG.RESERVATION_COLS;

  // 日付フォーマット
  [C.BOOKED_DATE, C.CHECKIN, C.CHECKOUT].forEach(col => {
    sheet.getRange(2, col, lastRow - 1, 1)
         .setNumberFormat('yyyy/MM/dd');
  });

  // 金額フォーマット
  [C.REVENUE, C.ACCOMMODATION, C.CLEANING_FEE, C.OTA_FEE, C.TRANSFER_FEE, C.PAYOUT].forEach(col => {
    sheet.getRange(2, col, lastRow - 1, 1)
         .setNumberFormat('¥#,##0');
  });

  // 交互に色付け
  for (let i = 2; i <= lastRow; i++) {
    const bg = i % 2 === 0 ? '#f8f9fa' : '#ffffff';
    sheet.getRange(i, 1, 1, 19).setBackground(bg);
  }
}

function formatMonthlySheet_(sheet, numDataRows, numCols) {
  if (numDataRows < 1) return;

  // 金額列のフォーマット
  const moneyStartCol = 7; // G列（売上）から
  const moneyEndCol   = 15; // O列（利益）まで
  sheet.getRange(2, moneyStartCol, numDataRows, moneyEndCol - moneyStartCol + 1)
       .setNumberFormat('¥#,##0');

  // KPI列
  sheet.getRange(2, 16, numDataRows, 1).setNumberFormat('0.0"%"'); // ROI
  sheet.getRange(2, 17, numDataRows, 1).setNumberFormat('¥#,##0'); // ADR
  sheet.getRange(2, 18, numDataRows, 1).setNumberFormat('¥#,##0'); // RevPAR
  sheet.getRange(2, 19, numDataRows, 1).setNumberFormat('0.0"%"'); // 稼働率

  // 交互に色付け
  for (let i = 2; i <= numDataRows + 1; i++) {
    const bg = i % 2 === 0 ? '#e8f0fe' : '#ffffff';
    sheet.getRange(i, 1, 1, numCols).setBackground(bg);
  }
}

function appendMonthlyTotalRow_(sheet, numDataRows, numCols, totalUsageDays, fiscalYear) {
  const totalRow = numDataRows + 2;
  sheet.getRange(totalRow, 1).setValue(`${fiscalYear}年度 合計`);
  sheet.getRange(totalRow, 1, 1, numCols)
       .setBackground('#fce8e6')
       .setFontWeight('bold');

  // SUM式 for 数値列
  const sumCols = [3, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]; // 稼働日数〜利益
  sumCols.forEach(col => {
    const colLetter = columnToLetter_(col);
    sheet.getRange(totalRow, col)
         .setFormula(`=SUM(${colLetter}2:${colLetter}${numDataRows + 1})`);
  });

  // 年間稼働率（180日制限に対する割合）
  const annualOccupancy = (totalUsageDays / CONFIG.PROPERTY.MAX_ANNUAL_DAYS * 100).toFixed(1);
  sheet.getRange(totalRow, numCols).setValue(Number(annualOccupancy));
  sheet.getRange(totalRow, numCols).setNote(`年間${CONFIG.PROPERTY.MAX_ANNUAL_DAYS}日上限に対する稼働率`);
}

function columnToLetter_(column) {
  let letter = '';
  while (column > 0) {
    const mod = (column - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    column = Math.floor((column - 1) / 26);
  }
  return letter;
}
