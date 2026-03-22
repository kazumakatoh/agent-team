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

  let revenue          = 0;
  let accommodationFee = 0;
  let cleaningFee      = 0;  // ゲスト負担の清掃料（売上側）
  let otaFee           = 0;
  let transferFee      = 0;
  let payout           = 0;
  let guests           = 0;
  let usageDays        = 0;
  let bookingCount     = 0;  // 利用件数（キャンセル除く全予約）

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
    revenue          += Number(row[C.REVENUE       - 1]) || 0;
    accommodationFee += Number(row[C.ACCOMMODATION - 1]) || 0;
    cleaningFee      += Number(row[C.CLEANING_FEE  - 1]) || 0;
    otaFee           += Number(row[C.OTA_FEE       - 1]) || 0;
    transferFee      += Number(row[C.TRANSFER_FEE  - 1]) || 0;
    payout           += Number(row[C.PAYOUT        - 1]) || 0;
    guests           += Number(row[C.GUESTS        - 1]) || 0;
    // 稼働日数 = 利用日数（USAGE_DAYS列）の合計
    usageDays        += Number(row[C.USAGE_DAYS    - 1]) || 0;
  });

  return {
    year,
    month,
    bookingCount,   // 問い合わせ数 & 利用件数（同値）
    usageDays,      // 稼働日数（利用日数の合計）
    guests,         // 利用人数（人数の合計）
    revenue,        // 売上（OTA表示の合計額）
    accommodationFee,  // 宿泊料
    cleaningFee,       // 清掃料（ゲスト負担・売上側）
    otaFee,         // OTA手数料
    transferFee,    // 振込手数料
    payout          // 入金金額
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
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();

  const row = data.find(r => String(r[C.YEAR_MONTH - 1]).startsWith(yearMonth));
  if (!row) {
    return { agencyFee: 0, cleaning: 0, linen: 0, supplies: 0, utilities: 0, rent: 0, other: 0 };
  }

  return {
    agencyFee: Number(row[C.AGENCY_FEE - 1]) || 0,
    cleaning:  Number(row[C.CLEANING   - 1]) || 0,
    linen:     Number(row[C.LINEN      - 1]) || 0,
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

  // 新しい27列ヘッダー
  // COL  1: 年月
  // COL  2: 問い合わせ数    COL  3: 稼働日数       COL  4: 利用可能日数
  // COL  5: 利用件数        COL  6: 利用人数
  // COL  7: 売上            COL  8: 宿泊料         COL  9: 清掃料
  // COL 10: OTA手数料       COL 11: 振込手数料     COL 12: 入金金額
  // COL 13: 代行手数料      COL 14: 清掃費         COL 15: リネン費
  // COL 16: 備品・消耗品費  COL 17: 水光熱費       COL 18: 家賃          COL 19: その他経費
  // COL 20: 総売上          COL 21: 流動費         COL 22: 固定費
  // COL 23: 利益            COL 24: 利益率(%)
  // COL 25: ADR(円)         COL 26: RevPAR(円)     COL 27: 稼働率(%)
  const headers = [
    '年月', '問い合わせ数', '稼働日数', '利用可能日数', '利用件数', '利用人数',
    '売上', '宿泊料', '清掃料', 'OTA手数料', '振込手数料', '入金金額',
    '代行手数料', '清掃費', 'リネン費', '備品・消耗品費', '水光熱費', '家賃', 'その他経費',
    '総売上', '流動費', '固定費', '利益', '利益率(%)',
    'ADR(円)', 'RevPAR(円)', '稼働率(%)'
  ];

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sheet, headers.length);

  const months = getFiscalYearMonths_(fiscalYear);
  let totalUsageDays = 0;
  const rows = [];

  months.forEach(({ year, month }) => {
    const resData     = getMonthlyReservationData(year, month);
    const costData    = getMonthlyCostData(year, month);
    const daysInMonth = new Date(year, month, 0).getDate();

    // 集計値
    const grossRevenue   = resData.revenue + resData.accommodationFee + resData.cleaningFee;
    const variableCosts  = resData.otaFee + resData.transferFee
                         + costData.agencyFee + costData.cleaning + costData.linen + costData.supplies;
    const fixedCosts     = costData.utilities + costData.rent + costData.other;
    const profit         = grossRevenue - variableCosts - fixedCosts;
    const profitRate     = grossRevenue > 0 ? Math.round(profit / grossRevenue * 1000) / 10 : 0;

    const adr           = KPICalculator.calcADR(resData.revenue, resData.usageDays);
    const revpar        = KPICalculator.calcRevPAR(resData.revenue, daysInMonth);
    const occupancyRate = KPICalculator.calcOccupancyRate(resData.usageDays, daysInMonth);

    totalUsageDays += resData.usageDays;

    const yearMonthKey = `${year}/${String(month).padStart(2, '0')}`;

    rows.push([
      yearMonthKey,
      resData.bookingCount,      // 問い合わせ数 = 予約件数
      resData.usageDays,         // 稼働日数 = 利用日数合計
      daysInMonth,               // 利用可能日数
      resData.bookingCount,      // 利用件数
      resData.guests,            // 利用人数
      resData.revenue,           // 売上
      resData.accommodationFee,  // 宿泊料
      resData.cleaningFee,       // 清掃料（ゲスト負担）
      resData.otaFee,            // OTA手数料
      resData.transferFee,       // 振込手数料
      resData.payout,            // 入金金額
      costData.agencyFee,        // 代行手数料
      costData.cleaning,         // 清掃費（業者への支払い）
      costData.linen,            // リネン費
      costData.supplies,         // 備品・消耗品費
      costData.utilities,        // 水光熱費
      costData.rent,             // 家賃
      costData.other,            // その他経費
      grossRevenue,              // 総売上
      variableCosts,             // 流動費
      fixedCosts,                // 固定費
      profit,                    // 利益
      profitRate,                // 利益率(%)
      adr,                       // ADR
      revpar,                    // RevPAR
      occupancyRate              // 稼働率(%)
    ]);
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    formatMonthlySheet_(sheet, rows.length, headers.length);
  }

  appendMonthlyTotalRow_(sheet, rows.length, headers.length, totalUsageDays, fiscalYear);

  Logger.log(`月別集計シート更新完了 (${fiscalYear}年度)`);
}

/**
 * 年間集計シートを更新する
 * @param {number} fiscalYear
 */
function updateAnnualSheet(fiscalYear) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ANNUAL);
  if (!sheet) throw new Error(`シート「${CONFIG.SHEETS.ANNUAL}」が見つかりません。`);

  const headers = [
    '事業年度', '稼働日数', '年間上限日数', '利用可能日数', '利用件数', '利用人数',
    '売上', '宿泊料', '清掃料', 'OTA手数料', '振込手数料', '入金金額',
    '代行手数料', '清掃費', 'リネン費', '備品・消耗品費', '水光熱費', '家賃', 'その他経費',
    '総売上', '流動費', '固定費', '利益', '利益率(%)',
    'ADR(円)', 'RevPAR(円)', '稼働率(%)', '法定稼働率(%)'
  ];

  // 既存の年度データを保持
  const existing = {};
  if (sheet.getLastRow() > 1) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
    rows.forEach(r => { if (r[0]) existing[String(r[0])] = r; });
  }

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sheet, headers.length);
  sheet.getRange(1, 1, 1, headers.length)
       .setBackground('#7b1fa2').setFontColor('#ffffff');

  // 集計
  const months = getFiscalYearMonths_(fiscalYear);
  let totals = {
    usageDays: 0, bookingCount: 0, guests: 0,
    revenue: 0, accommodationFee: 0, cleaningFee: 0,
    otaFee: 0, transferFee: 0, payout: 0,
    agencyFee: 0, cleaning: 0, linen: 0, supplies: 0, utilities: 0, rent: 0, other: 0,
    daysInYear: 0
  };

  months.forEach(({ year, month }) => {
    const res  = getMonthlyReservationData(year, month);
    const cost = getMonthlyCostData(year, month);
    const dim  = new Date(year, month, 0).getDate();
    totals.usageDays      += res.usageDays;
    totals.bookingCount   += res.bookingCount;
    totals.guests         += res.guests;
    totals.revenue        += res.revenue;
    totals.accommodationFee += res.accommodationFee;
    totals.cleaningFee    += res.cleaningFee;
    totals.otaFee         += res.otaFee;
    totals.transferFee    += res.transferFee;
    totals.payout         += res.payout;
    totals.agencyFee      += cost.agencyFee;
    totals.cleaning       += cost.cleaning;
    totals.linen          += cost.linen;
    totals.supplies       += cost.supplies;
    totals.utilities      += cost.utilities;
    totals.rent           += cost.rent;
    totals.other          += cost.other;
    totals.daysInYear     += dim;
  });

  const grossRevenue  = totals.revenue + totals.accommodationFee + totals.cleaningFee;
  const variableCosts = totals.otaFee + totals.transferFee + totals.agencyFee
                      + totals.cleaning + totals.linen + totals.supplies;
  const fixedCosts    = totals.utilities + totals.rent + totals.other;
  const profit        = grossRevenue - variableCosts - fixedCosts;
  const profitRate    = grossRevenue > 0 ? Math.round(profit / grossRevenue * 1000) / 10 : 0;
  const adr           = KPICalculator.calcADR(totals.revenue, totals.usageDays);
  const revpar        = KPICalculator.calcRevPAR(totals.revenue, totals.daysInYear);
  const occupancy     = KPICalculator.calcOccupancyRate(totals.usageDays, totals.daysInYear);
  const legalOccupancy = KPICalculator.calcOccupancyRate(totals.usageDays, CONFIG.PROPERTY.MAX_ANNUAL_DAYS);

  const row = [
    `${fiscalYear}年度`,
    totals.usageDays, CONFIG.PROPERTY.MAX_ANNUAL_DAYS, totals.daysInYear,
    totals.bookingCount, totals.guests,
    totals.revenue, totals.accommodationFee, totals.cleaningFee,
    totals.otaFee, totals.transferFee, totals.payout,
    totals.agencyFee, totals.cleaning, totals.linen, totals.supplies,
    totals.utilities, totals.rent, totals.other,
    grossRevenue, variableCosts, fixedCosts, profit, profitRate,
    adr, revpar, occupancy, legalOccupancy
  ];

  // 既存行があれば上書き、なければ追加
  const existingKey = `${fiscalYear}年度`;
  let targetRow = 2;
  if (sheet.getLastRow() >= 2) {
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    const idx = ids.findIndex(r => String(r[0]) === existingKey);
    if (idx >= 0) targetRow = idx + 2;
    else targetRow = sheet.getLastRow() + 1;
  }
  sheet.getRange(targetRow, 1, 1, headers.length).setValues([row]);

  // 書式
  sheet.getRange(targetRow, 7, 1, 17).setNumberFormat('¥#,##0');  // 売上〜利益
  sheet.getRange(targetRow, 24, 1, 1).setNumberFormat('0.0"%"'); // 利益率
  sheet.getRange(targetRow, 25, 1, 2).setNumberFormat('¥#,##0'); // ADR, RevPAR
  sheet.getRange(targetRow, 27, 1, 2).setNumberFormat('0.0"%"'); // 稼働率, 法定稼働率

  Logger.log(`年間集計シート更新完了 (${fiscalYear}年度)`);
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

  // 金額列: COL7〜COL23（売上〜利益）
  sheet.getRange(2, 7, numDataRows, 17).setNumberFormat('¥#,##0');

  // 利益率(%) COL24
  sheet.getRange(2, 24, numDataRows, 1).setNumberFormat('0.0"%"');

  // ADR COL25, RevPAR COL26
  sheet.getRange(2, 25, numDataRows, 2).setNumberFormat('¥#,##0');

  // 稼働率(%) COL27
  sheet.getRange(2, 27, numDataRows, 1).setNumberFormat('0.0"%"');

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

  // SUM式（稼働日数〜その他経費, 総売上〜利益）
  const sumCols = [3, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23];
  sumCols.forEach(col => {
    const colLetter = columnToLetter_(col);
    sheet.getRange(totalRow, col)
         .setFormula(`=SUM(${colLetter}2:${colLetter}${numDataRows + 1})`);
  });

  // 年間稼働率（180日制限に対する割合）
  const annualOccupancy = (totalUsageDays / CONFIG.PROPERTY.MAX_ANNUAL_DAYS * 100).toFixed(1);
  sheet.getRange(totalRow, 27).setValue(Number(annualOccupancy));
  sheet.getRange(totalRow, 27).setNote(`年間${CONFIG.PROPERTY.MAX_ANNUAL_DAYS}日上限に対する稼働率`);
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
