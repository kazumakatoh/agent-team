/**
 * Amazon Dashboard - レポートベース高速データ取得
 *
 * Orders API（1件ずつ）の代わりに SP-API Reports を使用
 * 60日分のデータを1回のAPI呼び出しで取得可能
 */

/**
 * レポートベースで指定期間の注文データを取得→D1シートに書き込み
 *
 * @param {string} startDate - 開始日 (YYYY-MM-DD)
 * @param {string} endDate - 終了日 (YYYY-MM-DD)
 */
function fetchOrdersByReport(startDate, endDate) {
  Logger.log('===== レポートベース注文取得: ' + startDate + ' 〜 ' + endDate + ' =====');

  setupDailyDataHeaders();

  // レポート作成→ダウンロード
  const reportContent = fetchReport(
    'GET_FLAT_FILE_ALL_ORDERS_DATA_BY_ORDER_DATE_GENERAL',
    startDate + 'T00:00:00+09:00',
    endDate + 'T23:59:59+09:00'
  );

  // TSVをパース
  const rows = parseTsv(reportContent);
  Logger.log('レポート行数: ' + rows.length);

  if (rows.length === 0) {
    Logger.log('データなし');
    return;
  }

  // ヘッダーからカラムインデックスを取得
  const headers = rows[0];
  const colIndex = {};
  headers.forEach((h, i) => { colIndex[h.trim().toLowerCase()] = i; });

  // 必要なカラムの存在確認
  const asinCol = colIndex['asin'];
  const dateCol = colIndex['purchase-date'] !== undefined ? colIndex['purchase-date'] : colIndex['purchase date'];
  const qtyCol = colIndex['quantity'];
  const priceCol = colIndex['item-price'] !== undefined ? colIndex['item-price'] : colIndex['item price'];
  const statusCol = colIndex['order-status'] !== undefined ? colIndex['order-status'] : colIndex['order status'];

  Logger.log('カラム検出: ASIN=' + asinCol + ', date=' + dateCol + ', qty=' + qtyCol + ', price=' + priceCol);

  // ASIN × 日付 で集計
  const summary = {};
  const productMaster = getProductMasterMap();
  const newAsins = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[asinCol]) continue;

    // 注文ステータスがキャンセル済みならスキップ
    if (statusCol !== undefined) {
      const status = String(row[statusCol]).trim();
      if (status === 'Cancelled') continue;
    }

    const asin = String(row[asinCol]).trim();
    const purchaseDate = String(row[dateCol]).trim().substring(0, 10); // YYYY-MM-DD
    const qty = parseInt(row[qtyCol]) || 0;
    const price = parseFloat(row[priceCol]) || 0;

    const key = purchaseDate + '_' + asin;
    if (!summary[key]) {
      summary[key] = {
        date: purchaseDate,
        asin: asin,
        sales: 0,
        cvCount: 0,
        units: 0,
      };
    }
    summary[key].sales += price;
    summary[key].cvCount += 1;
    summary[key].units += qty;

    // 新規ASIN検出
    if (!productMaster[asin] && !newAsins.includes(asin)) {
      newAsins.push(asin);
    }
  }

  // D1シートに書き込み
  const dataRows = [];
  for (const data of Object.values(summary)) {
    const master = productMaster[data.asin] || {};
    dataRows.push([
      data.date,
      data.asin,
      master.name || '',
      master.category || '',
      Math.round(data.sales),
      data.cvCount,
      data.units,
      '', '', '', '', '',  // セッション・PV等
      '', '', '',          // FBA手数料・返品
      '', '', '', '',      // 広告データ
      '', '',              // 仕入
      '暫定',
    ]);
  }

  if (dataRows.length > 0) {
    appendRows(SHEET_NAMES.D1_DAILY, dataRows);
    Logger.log('✅ D1 日次データ: ' + dataRows.length + ' 行書き込み完了');
  }

  // 新規ASIN追加
  if (newAsins.length > 0) {
    const masterRows = newAsins.map(asin => [
      asin, '', '', 'アクティブ', '自動検出'
    ]);
    appendRows(SHEET_NAMES.M1_PRODUCT_MASTER, masterRows);
    Logger.log('✅ 新規ASIN: ' + newAsins.length + ' 件追加');
  }

  Logger.log('集計日数: ' + new Set(Object.values(summary).map(s => s.date)).size + ' 日');
  Logger.log('集計ASIN数: ' + new Set(Object.values(summary).map(s => s.asin)).size + ' 件');
  Logger.log('===== 完了 =====');
}

/**
 * TSV文字列をパース
 */
function parseTsv(content) {
  return content.split('\n')
    .filter(line => line.trim())
    .map(line => line.split('\t'));
}

/**
 * 過去60日分を一括取得（レポートベース・高速版）
 * 1回の実行で完了
 */
function bulkFetchByReport() {
  // 過去1年分を月別に取得（30日制限対応）
  const periods = [
    ['2025-04-01', '2025-04-30'],
    ['2025-05-01', '2025-05-31'],
    ['2025-06-01', '2025-06-30'],
    ['2025-07-01', '2025-07-31'],
    ['2025-08-01', '2025-08-31'],
    ['2025-09-01', '2025-09-30'],
    ['2025-10-01', '2025-10-31'],
    ['2025-11-01', '2025-11-30'],
    ['2025-12-01', '2025-12-31'],
    ['2026-01-01', '2026-01-31'],
    ['2026-02-01', '2026-02-28'],
    ['2026-03-01', '2026-03-31'],
    ['2026-04-01', '2026-04-14'],
  ];


  const propKey = 'BULK_PERIOD_INDEX';
  const props = PropertiesService.getScriptProperties();
  let idx = parseInt(props.getProperty(propKey) || '0');

  if (idx >= periods.length) {
    Logger.log('✅ 全期間完了。resetBulkPeriod() でリセット可能');
    return;
  }

  // 1回の実行で2期間ずつ処理（6分制限対策）
  for (let i = 0; i < 2 && idx < periods.length; i++, idx++) {
    const p = periods[idx];
    Logger.log('=== ' + p[0] + ' 〜 ' + p[1] + ' ===');
    fetchOrdersByReport(p[0], p[1]);
  }

  props.setProperty(propKey, String(idx));
  Logger.log('進捗: ' + idx + '/' + periods.length + ' 期間完了');

  if (idx < periods.length) {
    Logger.log('📌 残り' + (periods.length - idx) + '期間。もう一度実行してください。');
  }
}

function resetBulkPeriod() {
  PropertiesService.getScriptProperties().deleteProperty('BULK_PERIOD_INDEX');
  Logger.log('✅ リセット完了');
}


/**
 * 昨日のデータをレポートベースで取得（日次用・高速版）
 */
function dailyFetchByReport() {
  Logger.log('===== 日次データ取得（レポートベース） =====');
  const yesterday = getYesterday();
  fetchOrdersByReport(yesterday, yesterday);
}

/**
 * テスト: 直近7日間をレポートベースで取得
 */
function testReportFetch7Days() {
  const endDate = new Date();
  const startDate = new Date();
  startDate.setDate(startDate.getDate() - 7);

  const start = Utilities.formatDate(startDate, 'Asia/Tokyo', 'yyyy-MM-dd');
  const end = Utilities.formatDate(endDate, 'Asia/Tokyo', 'yyyy-MM-dd');

  fetchOrdersByReport(start, end);
}
/**
 * Sales & Traffic Report を取得して日次データのPV・セッション等を更新
 */
function fetchTrafficReport(startDate, endDate) {
  Logger.log('===== Traffic Report 取得: ' + startDate + ' 〜 ' + endDate + ' =====');

  const body = {
    reportType: 'GET_SALES_AND_TRAFFIC_REPORT',
    marketplaceIds: [MARKETPLACE_ID_JP],
    dataStartTime: startDate,
    dataEndTime: endDate,
    reportOptions: {
      dateGranularity: 'DAY',
      asinGranularity: 'CHILD',
    },
  };

  const result = callSpApi('POST', '/reports/2021-06-30/reports', null, body);
  const reportId = result.reportId;

  for (let i = 0; i < 30; i++) {
    Utilities.sleep(10000);
    const status = getReportStatus(reportId);
    if (status.processingStatus === 'DONE') {
      const content = downloadReportDocument(status.reportDocumentId);
      const data = JSON.parse(content);

      if (data.salesAndTrafficByAsin) {
        writeTrafficToDaily(data.salesAndTrafficByAsin, startDate);
      }
      return;
    }
    if (status.processingStatus === 'FATAL') throw new Error('レポート生成失敗');
  }
}

function writeTrafficToDaily(asinData, dateStr) {
  const sheet = getOrCreateSheet(SHEET_NAMES.D1_DAILY);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const dailyData = sheet.getRange(2, 1, lastRow - 1, 12).getValues();

  // ASINマップ作成
  const trafficMap = {};
  for (const entry of asinData) {
    const asin = entry.childAsin || entry.parentAsin || '';
    const traffic = entry.trafficByAsin || {};
    trafficMap[asin] = {
      sessions: traffic.sessions || 0,
      pv: (traffic.browserPageViews || 0) + (traffic.mobileAppPageViews || 0),
      cvr: traffic.unitSessionPercentage || 0,
      buybox: traffic.buyBoxPercentage || 0,
    };
  }

  let updated = 0;
  for (let i = 0; i < dailyData.length; i++) {
    const date = dailyData[i][0];
    const asin = dailyData[i][1];
    if (!asin) continue;

    const rowDate = (date instanceof Date)
      ? Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd')
      : String(date).substring(0, 10);

    if (rowDate === dateStr && trafficMap[asin]) {
      const t = trafficMap[asin];
      dailyData[i][7] = t.sessions;
      dailyData[i][8] = t.pv;
      dailyData[i][9] = t.pv > 0 ? (t.sessions / t.pv * 100).toFixed(1) : 0;
      dailyData[i][10] = t.cvr;
      dailyData[i][11] = t.buybox;
      updated++;
    }
  }

  if (updated > 0) {
    sheet.getRange(2, 8, lastRow - 1, 5).setValues(
      dailyData.map(row => [row[7], row[8], row[9], row[10], row[11]])
    );
    Logger.log('✅ ' + updated + ' 行のトラフィックデータを更新（' + dateStr + '）');
  } else {
    Logger.log('⚠️ ' + dateStr + ' のマッチなし');
  }
}

/**
 * 過去N日分の Traffic Report を1日ずつ取得（推奨：N=7）
 * 1日ずつ呼ぶことで dateStr マッチングが正しく動く。
 */
function bulkFetchTraffic(days) {
  const n = days || 7;
  const today = new Date();
  for (let d = 1; d <= n; d++) {
    const target = new Date(today);
    target.setDate(target.getDate() - d);
    const dateStr = Utilities.formatDate(target, 'Asia/Tokyo', 'yyyy-MM-dd');
    Logger.log('--- Traffic ' + dateStr + ' ---');
    try {
      fetchTrafficReport(dateStr, dateStr);
    } catch (e) {
      Logger.log('エラー: ' + e.message);
    }
  }
}

function dailyFetchTraffic() {
  const yesterday = getYesterday();
  fetchTrafficReport(yesterday, yesterday);
}


function updateDailyWithTrafficJson(asinData) {
  Logger.log('ASIN別データ: ' + asinData.length + ' 件');

  if (asinData.length > 0) {
    Logger.log('1件目のキー: ' + Object.keys(asinData[0]).join(', '));
  }

  // このレポートは期間全体の合計値（日別ではない）
  // → D1の日次データではなく、サマリーとしてログに出力
  // → 日別ASIN別データが必要な場合は別アプローチが必要

  // まずは取得できたデータをログで確認
  let totalSessions = 0;
  let totalPv = 0;
  const asinSummary = {};

  for (const entry of asinData) {
    const asin = entry.childAsin || entry.parentAsin || '';
    const traffic = entry.trafficByAsin || {};
    const sessions = traffic.sessions || 0;
    const pv = (traffic.browserPageViews || 0) + (traffic.mobileAppPageViews || 0);
    const unitSessionPct = traffic.unitSessionPercentage || 0;
    const buybox = traffic.buyBoxPercentage || 0;

    totalSessions += sessions;
    totalPv += pv;

    if (asin && sessions > 0) {
      asinSummary[asin] = {
        sessions: sessions,
        pv: pv,
        cvr: unitSessionPct,
        buybox: buybox,
      };
    }
  }

  Logger.log('合計セッション: ' + totalSessions);
  Logger.log('合計PV: ' + totalPv);
  Logger.log('セッションありASIN数: ' + Object.keys(asinSummary).length);

  // 上位10件を表示
  const sorted = Object.entries(asinSummary).sort((a, b) => b[1].sessions - a[1].sessions);
  sorted.slice(0, 10).forEach(([asin, d]) => {
    Logger.log('  ' + asin + ': sessions=' + d.sessions + ' pv=' + d.pv + ' cvr=' + d.cvr + '%');
  });

  Logger.log('⚠️ このレポートは期間合計値です。日別ASIN別データは別途対応が必要です。');
}



/**
 * Traffic Report (JSON形式) をパースして日次データを更新
 */
function parseTrafficJson(content) {
  const data = JSON.parse(content);
  const trafficByDate = data.salesAndTrafficByAsin || data.salesAndTrafficByDate || [];

  Logger.log('Traffic JSON エントリ数: ' + trafficByDate.length);

  // ASIN別データがあるか確認
  if (data.salesAndTrafficByAsin) {
    updateDailyWithTraffic(data.salesAndTrafficByAsin, 'asin');
  } else {
    Logger.log('⚠️ ASIN別データなし。日付別データのみ。');
  }
}

/**
 * Traffic Report (TSV形式) をパースして日次データを更新
 */
function parseTrafficTsv(content) {
  const rows = parseTsv(content);
  Logger.log('Traffic TSV 行数: ' + rows.length);

  if (rows.length <= 1) {
    Logger.log('データなし');
    return;
  }

  const headers = rows[0];
  const colIndex = {};
  headers.forEach((h, i) => { colIndex[h.trim().toLowerCase().replace(/[\s-]/g, '_')] = i; });
  Logger.log('Traffic カラム: ' + headers.join(', '));

  // D1の既存データを更新
  const sheet = getOrCreateSheet(SHEET_NAMES.D1_DAILY);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const dailyData = sheet.getRange(2, 1, lastRow - 1, 22).getValues();
  let updated = 0;

  // TSVデータからASIN×日付のマップを作成
  const trafficMap = {};
  const dateCol = findCol(colIndex, ['date', 'date_range']);
  const asinCol = findCol(colIndex, ['(child)_asin', 'child_asin', 'asin']);
  const sessionsCol = findCol(colIndex, ['sessions', 'sessions___total']);
  const pvCol = findCol(colIndex, ['page_views', 'page_views___total']);
  const buyboxCol = findCol(colIndex, ['buy_box_percentage', 'featured_offer_(buy_box)_percentage']);
  const unitSessionCol = findCol(colIndex, ['unit_session_percentage', 'unit_session_percentage___total']);

  Logger.log('Traffic カラム検出: date=' + dateCol + ', asin=' + asinCol + ', sessions=' + sessionsCol + ', pv=' + pvCol);

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const date = dateCol !== -1 ? String(row[dateCol]).trim().substring(0, 10) : '';
    const asin = asinCol !== -1 ? String(row[asinCol]).trim() : '';
    if (date && asin) {
      trafficMap[date + '_' + asin] = {
        sessions: sessionsCol !== -1 ? parseInt(row[sessionsCol]) || 0 : 0,
        pv: pvCol !== -1 ? parseInt(row[pvCol]) || 0 : 0,
        buybox: buyboxCol !== -1 ? parseFloat(row[buyboxCol]) || 0 : 0,
        unitSessionPct: unitSessionCol !== -1 ? parseFloat(row[unitSessionCol]) || 0 : 0,
      };
    }
  }

  Logger.log('Traffic マップ件数: ' + Object.keys(trafficMap).length);

  // D1日次データを更新
  for (let i = 0; i < dailyData.length; i++) {
    const date = dailyData[i][0];
    const asin = dailyData[i][1];
    if (!date || !asin) continue;

    const dateStr = (date instanceof Date)
      ? Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd')
      : String(date).substring(0, 10);

    const key = dateStr + '_' + asin;
    const traffic = trafficMap[key];
    if (traffic) {
      dailyData[i][7] = traffic.sessions;  // H列: セッション数
      dailyData[i][8] = traffic.pv;        // I列: PV
      dailyData[i][9] = traffic.pv > 0 ? (traffic.sessions / traffic.pv * 100).toFixed(1) : 0;  // J列: CTR
      dailyData[i][10] = traffic.unitSessionPct;  // K列: CVR
      dailyData[i][11] = traffic.buybox;   // L列: BuyBox率
      updated++;
    }
  }

  if (updated > 0) {
    sheet.getRange(2, 8, lastRow - 1, 5).setValues(
      dailyData.map(row => [row[7], row[8], row[9], row[10], row[11]])
    );
    Logger.log('✅ ' + updated + ' 行のトラフィックデータを更新');
  } else {
    Logger.log('⚠️ マッチする行なし');
  }
}

/**
 * D1 にトラフィックが入っているか診断
 *   - 直近7日分について、ASIN×日付ベースで「セッション>0 行 / 総行数」を集計
 *   - PV/CTR/CVR/BuyBox 列の最新値を5件サンプリング表示
 *
 * GAS エディタから手動実行して動作確認に使う。
 */
function diagnoseTrafficCoverage() {
  Logger.log('===== Traffic カバレッジ診断 =====');
  const sheet = getOrCreateSheet(SHEET_NAMES.D1_DAILY);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) { Logger.log('D1 が空'); return; }

  const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
  const today = new Date();
  const fmt = d => Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  const start = fmt(new Date(today.getFullYear(), today.getMonth(), today.getDate() - 7));

  const byDate = {};
  for (const row of data) {
    const date = row[0] instanceof Date ? fmt(row[0]) : String(row[0]).substring(0, 10);
    if (!date || date < start) continue;
    if (!byDate[date]) byDate[date] = { rows: 0, withTraffic: 0, sessions: 0, pv: 0, buybox: 0, buyboxN: 0 };
    const x = byDate[date];
    x.rows++;
    const sessions = parseFloat(row[7]) || 0;
    const pv = parseFloat(row[8]) || 0;
    const buybox = parseFloat(row[11]) || 0;
    if (sessions > 0 || pv > 0) x.withTraffic++;
    x.sessions += sessions; x.pv += pv;
    if (buybox > 0) { x.buybox += buybox; x.buyboxN++; }
  }

  const dates = Object.keys(byDate).sort();
  Logger.log('日付 | 行数 | トラフィックあり | セッション計 | PV計 | BuyBox平均');
  for (const d of dates) {
    const x = byDate[d];
    const cov = x.rows > 0 ? Math.round(x.withTraffic / x.rows * 100) + '%' : '-';
    const bb = x.buyboxN > 0 ? (x.buybox / x.buyboxN).toFixed(1) + '%' : '-';
    Logger.log(`${d} | ${x.rows} | ${x.withTraffic}(${cov}) | ${x.sessions} | ${x.pv} | ${bb}`);
  }
  Logger.log('===== 終了 =====');
}
/**
 * FBA在庫データを取得
 */
function fetchInventory() {
  Logger.log('===== FBA在庫データ取得 =====');

  const reportContent = fetchReport(
    'GET_FBA_MYI_UNSUPPRESSED_INVENTORY_DATA'
  );

  Logger.log('レポート先頭300文字: ' + reportContent.substring(0, 300));

  const rows = parseTsv(reportContent);
  Logger.log('在庫レポート行数: ' + rows.length);

  if (rows.length <= 1) {
    Logger.log('データなし');
    return;
  }

  const headers = rows[0];
  const colIndex = {};
  headers.forEach((h, i) => { colIndex[h.trim().toLowerCase().replace(/[\s-]/g, '_')] = i; });
  Logger.log('在庫カラム: ' + headers.join(', '));

  const asinCol = findCol(colIndex, ['asin', 'product_name']);
  const skuCol = findCol(colIndex, ['sku', 'seller_sku']);
  const qtyCol = findCol(colIndex, ['afn_fulfillable_quantity', 'quantity_available']);
  const nameCol = findCol(colIndex, ['product_name', 'product']);
  const conditionCol = findCol(colIndex, ['condition', 'item_condition']);

  Logger.log('カラム検出: asin=' + asinCol + ', sku=' + skuCol + ', qty=' + qtyCol);

  // 在庫データを整理してログに出力
  const inventory = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const asin = asinCol !== -1 ? String(row[asinCol]).trim() : '';
    const sku = skuCol !== -1 ? String(row[skuCol]).trim() : '';
    const qty = qtyCol !== -1 ? parseInt(row[qtyCol]) || 0 : 0;
    const name = nameCol !== -1 ? String(row[nameCol]).trim() : '';

    if (asin || sku) {
      inventory.push({ asin, sku, qty, name });
    }
  }

  Logger.log('在庫商品数: ' + inventory.length);
  Logger.log('在庫あり: ' + inventory.filter(i => i.qty > 0).length + ' 件');
  Logger.log('在庫なし: ' + inventory.filter(i => i.qty === 0).length + ' 件');

  // 上位10件表示
  inventory.sort((a, b) => b.qty - a.qty);
  inventory.slice(0, 10).forEach(i => {
    Logger.log('  ' + i.asin + ' (' + i.sku + '): ' + i.qty + '個');
  });

  Logger.log('===== 完了 =====');
}
