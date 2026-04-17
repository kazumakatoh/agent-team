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

    // レポートの期間を識別キー（YYYY-MM-DD 正規化）で照合
    const periodKey = normalizeDateKey(report.dataStartTime) + '_' + normalizeDateKey(report.dataEndTime);
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

        // D2F の暫定データを「確定」にマーク（重複集計防止）
        try {
          markFinanceEventsAsConfirmed(
            String(report.dataStartTime).substring(0, 10),
            String(report.dataEndTime).substring(0, 10)
          );
        } catch (mErr) {
          Logger.log('⚠️ D2F 確定マークエラー: ' + mErr.message);
        }
      }
    } catch (e) {
      Logger.log('⚠️ レポート取得エラー: ' + e.message);
    }
  }

  Logger.log('===== Settlement Report 取得完了: ' + newReports + ' 件の新規レポート =====');

  // 新規取込があれば月次集計も更新
  if (newReports > 0) {
    buildSettlementSummary();
  }
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
 *
 * ⚠️ D2 には日付部分 (YYYY-MM-DD) のみ保存されているため、
 * 比較時も Report の dataStartTime/EndTime を YYYY-MM-DD に正規化して返す。
 */
function getProcessedSettlementPeriods() {
  const sheet = getOrCreateSheet(SHEET_NAMES.D2_SETTLEMENT);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const periods = new Set();
  data.forEach(row => {
    const start = normalizeDateKey(row[0]);
    const end = normalizeDateKey(row[1]);
    if (start && end) {
      periods.add(start + '_' + end);
    }
  });
  return Array.from(periods);
}

/**
 * 任意の日付値（Date/文字列/ISO）を 'YYYY-MM-DD' 文字列に正規化
 */
function normalizeDateKey(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  return String(value).substring(0, 10);
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

/**
 * 診断: D2 経費明細の重複状況を調査
 *
 * 出力内容:
 *   【A】 決済期間ごとの行数
 *   【B】 重複行（決済期間開始+終了+日付+ASIN+種別+明細+金額+数量 が同一）
 *   【C】 重複除外後の想定行数
 */
function debugSettlementDuplicates() {
  const sheet = getOrCreateSheet(SHEET_NAMES.D2_SETTLEMENT);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) { Logger.log('D2 空'); return; }

  Logger.log('D2 総行数: ' + (lastRow - 1));

  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();

  // 【A】 決済期間ごとの行数
  const periodCounts = {};
  for (const row of data) {
    const start = String(row[0] || '').substring(0, 10);
    const end = String(row[1] || '').substring(0, 10);
    const key = start + ' 〜 ' + end;
    periodCounts[key] = (periodCounts[key] || 0) + 1;
  }

  Logger.log('');
  Logger.log('【A】 決済期間ごとの行数:');
  Object.entries(periodCounts)
    .sort((a, b) => b[1] - a[1])
    .forEach(([k, v]) => Logger.log('  ' + k + ': ' + v + ' 行'));

  // 【B】 完全重複チェック（全8列が同一）
  const seen = new Set();
  let duplicates = 0;
  const sampleDups = [];
  for (const row of data) {
    const key = row.map(v => {
      if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
      return String(v);
    }).join('|');
    if (seen.has(key)) {
      duplicates++;
      if (sampleDups.length < 5) sampleDups.push(key);
    } else {
      seen.add(key);
    }
  }

  Logger.log('');
  Logger.log('【B】 完全重複（全8列同一）:');
  Logger.log('  重複行数: ' + duplicates);
  Logger.log('  ユニーク行数: ' + seen.size);
  if (sampleDups.length > 0) {
    Logger.log('  重複サンプル（先頭5件）:');
    sampleDups.forEach((k, i) => Logger.log('    ' + (i + 1) + ') ' + k.substring(0, 150)));
  }

  // 【C】 重複除外後
  Logger.log('');
  Logger.log('【C】 重複除外後の想定行数: ' + seen.size);
  Logger.log('  削減見込み: ' + duplicates + ' 行（' + ((duplicates / (lastRow - 1) * 100).toFixed(1)) + '%）');
}

/**
 * D2 経費明細の完全重複行を削除
 *
 * ヘッダーは残し、2行目以降で全列が同一の重複を1件だけ残す。
 * 実行前に debugSettlementDuplicates で重複状況を確認すること。
 */
function deduplicateSettlement() {
  const t0 = Date.now();
  const sheet = getOrCreateSheet(SHEET_NAMES.D2_SETTLEMENT);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) { Logger.log('D2 空'); return; }

  Logger.log('===== D2 経費明細 重複除外 開始 =====');
  Logger.log('処理前: ' + (lastRow - 1) + ' 行');

  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const seen = new Set();
  const unique = [];

  for (const row of data) {
    const key = row.map(v => {
      if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
      return String(v);
    }).join('|');
    if (!seen.has(key)) {
      seen.add(key);
      unique.push(row);
    }
  }

  const removed = data.length - unique.length;
  Logger.log('重複除外: ' + removed + ' 行削除');
  Logger.log('残存: ' + unique.length + ' 行');

  // シートをクリアして書き戻し
  sheet.getRange(2, 1, lastRow - 1, 8).clearContent();
  if (unique.length > 0) {
    sheet.getRange(2, 1, unique.length, 8).setValues(unique);
  }

  // D2S 月次集計を再構築
  Logger.log('D2S 月次集計を再構築...');
  buildSettlementSummary();

  Logger.log('===== 完了（' + (Date.now() - t0) + 'ms）=====');
}

/**
 * 経費明細(D2)から月次集計を作成（D2_SETTLEMENT_SUMMARY）
 * ASIN × 年月 × (commission, other) の形式
 *
 * ダッシュボードはこの集計シートから読み込むため高速化される
 */
function buildSettlementSummary() {
  const t0 = Date.now();
  Logger.log('===== 経費月次集計 開始 =====');

  const srcSheet = getOrCreateSheet(SHEET_NAMES.D2_SETTLEMENT);
  const lastRow = srcSheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log('経費明細データなし');
    return;
  }

  // 単一読み込み: 列3-7（日付、ASIN、トランザクション種別、明細種別、金額）
  const t1 = Date.now();
  const data = srcSheet.getRange(2, 3, lastRow - 1, 5).getValues();
  Logger.log('読み込み: ' + (Date.now()-t1) + 'ms (' + data.length + '行)');

  // 集計（Utilities.formatDateを避けて高速化）
  const t2 = Date.now();
  const summary = {};

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rawDate = row[0];
    let yearMonth = '';

    if (rawDate instanceof Date) {
      const y = rawDate.getFullYear();
      const m = rawDate.getMonth() + 1;
      yearMonth = y + '-' + (m < 10 ? '0' + m : String(m));
    } else if (rawDate) {
      const s = String(rawDate);
      if (s.length >= 7) yearMonth = s.substring(0, 7);
    }
    if (!yearMonth) continue;

    const asin = String(row[1] || '').trim();
    const itemType = String(row[3]).trim();  // 明細種別
    const amount = parseFloat(row[4]) || 0;

    if (itemType === 'Principal' || itemType === 'Tax') continue;

    const expense = -amount;
    const key = asin + '_' + yearMonth;

    if (!summary[key]) {
      summary[key] = { asin: asin, yearMonth: yearMonth, commission: 0, other: 0 };
    }

    if (itemType === 'Commission') {
      summary[key].commission += expense;
    } else {
      summary[key].other += expense;
    }
  }
  Logger.log('集計: ' + (Date.now()-t2) + 'ms');

  // 書き込み
  const t3 = Date.now();
  const rows = Object.values(summary).map(s => [s.asin, s.yearMonth, s.commission, s.other]);
  rows.sort((a, b) => (a[1] + a[0]).localeCompare(b[1] + b[0]));

  const dstSheet = getOrCreateSheet(SHEET_NAMES.D2S_SETTLEMENT_SUMMARY);
  dstSheet.clear();
  dstSheet.getRange(1, 1, 1, 4).setValues([['ASIN', '年月', '販売手数料', 'その他経費']])
    .setFontWeight('bold').setBackground('#e8f0fe');
  dstSheet.setFrozenRows(1);

  if (rows.length > 0) {
    dstSheet.getRange(2, 2, rows.length, 1).setNumberFormat('@'); // 年月列をテキスト形式に
    dstSheet.getRange(2, 1, rows.length, 4).setValues(rows);
    dstSheet.getRange(2, 3, rows.length, 2).setNumberFormat('#,##0');
  }
  Logger.log('書き込み: ' + (Date.now()-t3) + 'ms (' + rows.length + '件)');

  Logger.log('===== 経費月次集計 完了（' + (Date.now()-t0) + 'ms）=====');
}
