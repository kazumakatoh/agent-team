/**
 * Amazon Dashboard - 在庫管理 + 在庫切れアラート
 *
 * SP-API の GET_FBA_MYI_UNSUPPRESSED_INVENTORY_DATA レポートを取得し、
 * D6 在庫シートに保存。直近N日の日販平均と照合して「残り何日分」を算出。
 * 残り日数 < STOCK_DAYS_THRESHOLD の商品は LINE 緊急アラートへ。
 *
 * ## D6 シート構造（9列）
 *   取得日時 | ASIN | SKU | 商品名 | 在庫数 | 7日平均日販 | 残り日数 | ステータス | 備考
 *
 * ## ステータス
 *   緊急（3日分未満）/ 警告（7日分未満）/ 注意（14日分未満）/ OK
 *
 * ## トリガー: 毎日 AM10:00 (fetchInventoryAndAlert)
 */

const D6_INVENTORY = '在庫';
const D6_INVENTORY_HEADERS = ['取得日時', 'ASIN', 'SKU', '商品名', '在庫数', '7日平均日販', '残り日数', 'ステータス', '備考'];

const STOCK_CRITICAL_DAYS = 3;   // 3日分未満 → 緊急
const STOCK_WARNING_DAYS = 7;    // 7日分未満 → 警告
const STOCK_CAUTION_DAYS = 14;   // 14日分未満 → 注意
const SALES_LOOKBACK_DAYS = 7;   // 日販平均の参照期間

/**
 * メイン: 在庫取得 → D6 更新 → 低在庫アラート
 */
function fetchInventoryAndAlert() {
  const t0 = Date.now();
  Logger.log('===== 在庫取得・アラート 開始 =====');

  setupInventorySheet();

  // 在庫データ取得
  const inventory = fetchInventoryData();
  if (inventory.length === 0) { Logger.log('在庫データなし'); return; }

  // 直近7日の日販平均を計算
  const salesAvg = getRecentDailySalesByAsin(SALES_LOOKBACK_DAYS);

  // 既存 D6 を全削除して最新スナップショットに置換（履歴不要ならこれが一番シンプル）
  const sheet = getOrCreateSheetCompact(D6_INVENTORY, D6_INVENTORY_HEADERS.length, 500);
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, D6_INVENTORY_HEADERS.length).clearContent();

  const now = new Date();
  const nowStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');

  const productMaster = getProductMasterMap();
  const rows = [];
  const alerts = [];
  for (const inv of inventory) {
    const asin = inv.asin;
    if (!asin) continue;
    const avgDaily = salesAvg[asin] || 0;
    const daysLeft = avgDaily > 0 ? (inv.qty / avgDaily) : (inv.qty > 0 ? 999 : 0);
    const status = statusFromDaysLeft(daysLeft, inv.qty);
    const name = (productMaster[asin] || {}).name || inv.name || '';

    rows.push([
      nowStr, asin, inv.sku, name,
      inv.qty, avgDaily.toFixed(1), daysLeft.toFixed(1),
      status, '',
    ]);

    // アラート対象（緊急・警告 かつ 売れている商品のみ）
    if ((status === '緊急' || status === '警告') && avgDaily > 0) {
      alerts.push({
        key: 'STOCK_' + asin + '_' + Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd'),
        type: status === '緊急' ? '🚨在庫切れ間近' : '⚠️在庫警告',
        line: asin + ' (' + (name || inv.sku).substring(0, 15) + '): 残り' +
              inv.qty + '個 / 日販' + avgDaily.toFixed(1) + ' = 約' + daysLeft.toFixed(1) + '日分',
      });
    }
  }

  // D6 書き込み
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, D6_INVENTORY_HEADERS.length).setValues(rows);
    // 残り日数昇順で並べ替え（Googleシート側の並びが一目で分かりやすい）
    sheet.getRange(2, 1, rows.length, D6_INVENTORY_HEADERS.length).sort({ column: 7, ascending: true });
  }

  Logger.log('✅ 在庫 ' + rows.length + ' 件 / アラート ' + alerts.length + ' 件 (' + (Date.now() - t0) + 'ms)');

  // LINE 通知（LineAlert.gs のユーティリティを流用、重複抑止）
  if (alerts.length > 0) notifyStockAlerts(alerts);
}

function statusFromDaysLeft(daysLeft, qty) {
  if (qty <= 0) return '在庫切れ';
  if (daysLeft < STOCK_CRITICAL_DAYS) return '緊急';
  if (daysLeft < STOCK_WARNING_DAYS) return '警告';
  if (daysLeft < STOCK_CAUTION_DAYS) return '注意';
  return 'OK';
}

function notifyStockAlerts(alerts) {
  const sent = getSentAlertMap();
  const now = Date.now();
  const fresh = alerts.filter(a => {
    const last = sent[a.key];
    return !last || (now - last) > ALERT_DEDUP_HOURS * 3600 * 1000;
  });
  if (fresh.length === 0) return;
  pushLineAlert(formatAlertMessage(fresh));
  fresh.forEach(a => { sent[a.key] = now; });
  saveSentAlertMap(sent);
}

/**
 * FBA Inventory API (/fba/inventory/v1/summaries) 経由で在庫を取得
 * レポート型（GET_FBA_MYI_UNSUPPRESSED_INVENTORY_DATA）より
 * タイムアウト耐性が高く、ページングもできる。
 */
function fetchInventoryData() {
  const all = [];
  let nextToken = null;
  let page = 0;
  const maxPages = 20;  // 35商品想定なので1ページで済むはずだが安全マージン

  do {
    const params = {
      granularityType: 'Marketplace',
      granularityId: MARKETPLACE_ID_JP,
      marketplaceIds: MARKETPLACE_ID_JP,
      details: 'true',
    };
    if (nextToken) params.nextToken = nextToken;

    const res = callSpApi('GET', '/fba/inventory/v1/summaries', params);
    const payload = res.payload || {};
    const summaries = payload.inventorySummaries || [];

    for (const s of summaries) {
      const detail = s.inventoryDetails || {};
      all.push({
        asin: s.asin || '',
        sku: s.sellerSku || '',
        qty: detail.fulfillableQuantity || s.totalQuantity || 0,
        name: s.productName || '',
      });
    }

    nextToken = payload.nextToken || null;
    page++;
    if (nextToken) Utilities.sleep(500);
  } while (nextToken && page < maxPages);

  Logger.log('在庫API取得: ' + all.length + ' 件（' + page + ' ページ）');
  return all;
}

/**
 * （旧実装）レポート経由の取得。FATAL エラーが多いため非推奨。
 * APIが不安定な時用のフォールバック。
 */
function fetchInventoryDataByReport() {
  const reportContent = fetchReport('GET_FBA_MYI_UNSUPPRESSED_INVENTORY_DATA');
  const rows = parseTsv(reportContent);
  if (rows.length <= 1) return [];

  const headers = rows[0];
  const colIndex = {};
  headers.forEach((h, i) => { colIndex[h.trim().toLowerCase().replace(/[\s-]/g, '_')] = i; });

  const asinCol = findCol(colIndex, ['asin']);
  const skuCol = findCol(colIndex, ['sku', 'seller_sku']);
  const qtyCol = findCol(colIndex, ['afn_fulfillable_quantity', 'quantity_available']);
  const nameCol = findCol(colIndex, ['product_name', 'product']);

  const inventory = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const asin = asinCol !== -1 ? String(row[asinCol]).trim() : '';
    const sku = skuCol !== -1 ? String(row[skuCol]).trim() : '';
    const qty = qtyCol !== -1 ? parseInt(row[qtyCol]) || 0 : 0;
    const name = nameCol !== -1 ? String(row[nameCol]).trim() : '';
    if (asin || sku) inventory.push({ asin, sku, qty, name });
  }
  return inventory;
}

/**
 * D1 から直近N日の ASIN別 日販平均を算出
 * (ASIN → 平均日販)
 */
function getRecentDailySalesByAsin(days) {
  const dailyData = getDailyDataAll();
  const today = new Date();
  const fmt = d => Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  const start = fmt(new Date(today.getFullYear(), today.getMonth(), today.getDate() - days));
  const end = fmt(new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1));

  const totals = {};
  const daysSeen = {};
  for (const d of dailyData) {
    if (d.date < start || d.date > end) continue;
    if (!d.asin) continue;
    totals[d.asin] = (totals[d.asin] || 0) + d.units;
    if (!daysSeen[d.asin]) daysSeen[d.asin] = new Set();
    daysSeen[d.asin].add(d.date);
  }

  const avg = {};
  for (const asin of Object.keys(totals)) {
    const n = Math.max(1, daysSeen[asin].size);
    avg[asin] = totals[asin] / n;
  }
  return avg;
}

function setupInventorySheet() {
  const sheet = getOrCreateSheetCompact(D6_INVENTORY, D6_INVENTORY_HEADERS.length, 500);
  const existing = sheet.getRange(1, 1, 1, D6_INVENTORY_HEADERS.length).getValues()[0];
  if (existing[0] !== D6_INVENTORY_HEADERS[0]) {
    sheet.getRange(1, 1, 1, D6_INVENTORY_HEADERS.length).setValues([D6_INVENTORY_HEADERS])
      .setFontWeight('bold').setBackground('#e8f0fe');
    sheet.setFrozenRows(1);
  }
}

// findCol() は SettlementFetch.gs で定義済み（共有）

/**
 * テスト: アラート送信せずに在庫取得のみ確認
 */
function testFetchInventoryOnly() {
  setupInventorySheet();
  const inv = fetchInventoryData();
  Logger.log('取得件数: ' + inv.length);
  const sold = inv.filter(i => i.qty > 0).sort((a, b) => b.qty - a.qty);
  Logger.log('在庫あり（qty > 0）: ' + sold.length + ' 件');
  sold.slice(0, 20).forEach(i => {
    Logger.log('  ' + i.asin + ' / ' + i.sku + ' / qty=' + i.qty);
  });
}

/**
 * デバッグ: SP-API の生レスポンスを確認
 * fulfillableQuantity が入っているか構造を調べる
 */
function debugInventoryRaw() {
  const params = {
    granularityType: 'Marketplace',
    granularityId: MARKETPLACE_ID_JP,
    marketplaceIds: MARKETPLACE_ID_JP,
    details: 'true',
  };
  const res = callSpApi('GET', '/fba/inventory/v1/summaries', params);
  const summaries = (res.payload && res.payload.inventorySummaries) || [];
  Logger.log('総件数: ' + summaries.length);
  Logger.log('--- 先頭3件の構造 ---');
  summaries.slice(0, 3).forEach((s, i) => {
    Logger.log('[' + i + '] ' + JSON.stringify(s, null, 2));
  });
}
