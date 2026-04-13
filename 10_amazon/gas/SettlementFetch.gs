/**
 * Amazon Dashboard - Settlement Report（経費明細）取得モジュール
 *
 * Settlement Report から手数料・返品・広告費等の確定経費を取得
 * 14日サイクルで確定するため、最新のレポートを取得して D2 経費明細に書き込む
 */

/**
 * Settlement Report の一覧を取得
 * @param {number} [maxResults=10] - 取得件数
 * @returns {Array} レポート一覧
 */
function getSettlementReportList(maxResults) {
  const params = {
    reportTypes: 'GET_V2_SETTLEMENT_REPORT_DATA_FLAT_FILE_V2',
    marketplaceIds: MARKETPLACE_ID_JP,
    pageSize: maxResults || 10,
  };
  const result = callSpApi('GET', '/reports/2021-06-30/reports', params);
  return result.reports || [];
}

/**
 * Settlement Report を取得して D2 経費明細に書き込む
 */
function fetchSettlementReports() {
  Logger.log('===== Settlement Report 取得開始 =====');

  setupSettlementHeaders();

  // 最新のSettlement Reportを取得
  const reports = getSettlementReportList(5);
  Logger.log('利用可能なレポート数: ' + reports.length);

  if (reports.length === 0) {
    Logger.log('❌ Settlement Report が見つかりません');
    return;
  }

  // 既にD2に取り込み済みの期間を確認
  const processedPeriods = getProcessedSettlementPeriods();

  let newReports = 0;

  for (const report of reports) {
    if (report.processingStatus !== 'DONE') {
      Logger.log('スキップ（未完了）: ' + report.reportId);
      continue;
    }

    // レポートの期間を識別キーにする
    const periodKey = report.dataStartTime + '_' + report.dataEndTime;
    if (processedPeriods.includes(periodKey)) {
      Logger.log('スキップ（取り込み済み）: ' + periodKey);
      continue;
    }

    Logger.log('--- レポート取得: ' + report.reportId + ' ---');
    Logger.log('期間: ' + report.dataStartTime + ' 〜 ' + report.dataEndTime);

    try {
      const content = downloadReportDocument(report.reportDocumentId);
      const rows = parseSettlementReport(content, report.dataStartTime, report.dataEndTime);

      if (rows.length > 0) {
        appendRows(SHEET_NAMES.D2_SETTLEMENT, rows);
        Logger.log('✅ D2 経費明細: ' + rows.length + ' 行書き込み');
        newReports++;
      }
    } catch (e) {
      Logger.log('⚠️ レポート取得エラー: ' + e.message);
    }
  }

  Logger.log('===== Settlement Report 取得完了: ' + newReports + ' 件の新規レポート =====');
}

/**
 * Settlement Report の内容をパースして D2 用の行データに変換
 */
function parseSettlementReport(content, startTime, endTime) {
  const tsvRows = parseTsv(content);

  if (tsvRows.length <= 1) {
    Logger.log('レポートデータなし');
    return [];
  }

  // ヘッダーからカラムインデックスを取得
  const headers = tsvRows[0];
  const colIndex = {};
  headers.forEach((h, i) => { colIndex[h.trim().toLowerCase().replace(/-/g, '_')] = i; });

  Logger.log('Settlement カラム数: ' + headers.length);
  Logger.log('Settlement ヘッダー例: ' + headers.slice(0, 10).join(', '));

  // カラム名の候補（Amazon のレポートはバージョンによって名前が違う場合あり）
  const dateCol = findCol(colIndex, ['posted_date', 'posted_date_time', 'date_time']);
  const typeCol = findCol(colIndex, ['transaction_type', 'type']);
  const asinCol = findCol(colIndex, ['sku', 'asin']);
  const descCol = findCol(colIndex, ['amount_description', 'description', 'fee_type']);
  const amountCol = findCol(colIndex, ['amount', 'total']);
  const qtyCol = findCol(colIndex, ['quantity_purchased', 'quantity']);

  Logger.log('カラム検出: date=' + dateCol + ', type=' + typeCol + ', asin=' + asinCol + ', amount=' + amountCol);

  const startDate = startTime ? startTime.substring(0, 10) : '';
  const endDate = endTime ? endTime.substring(0, 10) : '';

  const rows = [];
  for (let i = 1; i < tsvRows.length; i++) {
    const row = tsvRows[i];
    if (row.length < 3) continue;

    const postedDate = dateCol !== -1 ? String(row[dateCol]).trim().substring(0, 10) : '';
    const txType = typeCol !== -1 ? String(row[typeCol]).trim() : '';
    const asin = asinCol !== -1 ? String(row[asinCol]).trim() : '';
    const desc = descCol !== -1 ? String(row[descCol]).trim() : '';
    const amount = amountCol !== -1 ? parseFloat(row[amountCol]) || 0 : 0;
    const qty = qtyCol !== -1 ? parseInt(row[qtyCol]) || 0 : 0;

    rows.push([
      startDate,      // 決済期間開始
      endDate,        // 決済期間終了
      postedDate,     // 日付
      asin,           // ASIN/SKU
      txType,         // トランザクション種別
      desc,           // 明細種別
      amount,         // 金額
      qty,            // 数量
    ]);
  }

  return rows;
}

/**
 * カラム名の候補から最初に見つかったインデックスを返す
 */
function findCol(colIndex, candidates) {
  for (const name of candidates) {
    if (colIndex[name] !== undefined) return colIndex[name];
  }
  return -1;
}

/**
 * 既に取り込み済みの Settlement 期間を取得
 */
function getProcessedSettlementPeriods() {
  const sheet = getOrCreateSheet(SHEET_NAMES.D2_SETTLEMENT);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const periods = new Set();
  data.forEach(row => {
    if (row[0] && row[1]) {
      periods.add(row[0] + '_' + row[1]);
    }
  });
  return Array.from(periods);
}

/**
 * テスト: Settlement Report 一覧を表示
 */
function testSettlementList() {
  const reports = getSettlementReportList(10);
  Logger.log('Settlement Report 一覧: ' + reports.length + ' 件');
  reports.forEach(r => {
    Logger.log('  ' + r.reportId + ' | ' + r.processingStatus + ' | ' + (r.dataStartTime || '') + ' 〜 ' + (r.dataEndTime || ''));
  });
}
