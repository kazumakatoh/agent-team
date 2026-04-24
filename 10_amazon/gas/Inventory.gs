/**
 * Amazon Dashboard - 在庫管理 + 在庫切れアラート
 *
 * SP-API の /fba/inventory/v1/summaries を取得し、D6 在庫シートに保存。
 * 直近7日の日販平均と照合して「残り何日分」を算出。
 * 残り20日を切ったら LINE 緊急アラートへ（マイルストーン方式）。
 *
 * ## D6 シート構造（9列）
 *   取得日時 | ASIN | SKU | 商品名 | 在庫数 | 7日平均日販 | 残り日数 | ステータス | 備考
 *
 * ## ステータス
 *   在庫切れ / 緊急（<7日）/ 警告（<15日）/ 注意（<20日）/ OK
 *
 * ## 除外ルール
 *   - 死にSKU（在庫0 かつ 日販0）
 *   - 商品マスター(M1)のカテゴリが「カタログ削除」
 *
 * ## アラート設計（マイルストーン方式）
 *   20日 / 15日 / 7日 / 0日 の境界を下回った時に1回だけ通知。
 *   0日通知後は再通知しない（見切った商品を想定）。
 *   在庫回復（>20日）で状態リセット → 再度在庫減少時に通知再開。
 *
 * ## トリガー: 毎日 AM10:00 (fetchInventoryAndAlert)
 */

const D6_INVENTORY = '在庫';
const D6_INVENTORY_HEADERS = ['取得日時', 'ASIN', 'SKU', '商品名', '在庫数', '7日平均日販', '残り日数', 'ステータス', '備考'];

const SALES_LOOKBACK_DAYS = 7;           // 日販平均の参照期間
const STOCK_THRESHOLDS = [20, 15, 7, 0]; // アラート発火の境界日数（多い順）
const STOCK_RESET_DAYS = 20;             // この日数を超えたら state リセット（= 再通知の準備完了）
const STOCK_ALERT_STATE_KEY = 'STOCK_ALERT_STATE';
const EXCLUDED_CATEGORY = 'カタログ削除';

/**
 * メイン: 在庫取得 → D6 更新 → マイルストーン通知
 */
function fetchInventoryAndAlert() {
  const t0 = Date.now();
  Logger.log('===== 在庫取得・アラート 開始 =====');

  setupInventorySheet();

  const inventory = fetchInventoryData();
  if (inventory.length === 0) { Logger.log('在庫データなし'); return; }

  const salesAvg = getRecentDailySalesByAsin(SALES_LOOKBACK_DAYS);

  // D6 を全削除して最新スナップショットで置換
  const sheet = getOrCreateSheetCompact(D6_INVENTORY, D6_INVENTORY_HEADERS.length, 500);
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, D6_INVENTORY_HEADERS.length).clearContent();

  const now = new Date();
  const nowStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');

  const productMaster = getProductMasterMap();
  const alertState = getStockAlertState();
  const rows = [];
  const alerts = [];
  let skippedDeleted = 0;

  for (const inv of inventory) {
    const asin = inv.asin;
    if (!asin) continue;
    const master = productMaster[asin] || {};
    const category = master.category || '';

    // カタログ削除カテゴリは管理対象外（シートにも出さずアラートも送らない）
    if (category === EXCLUDED_CATEGORY) { skippedDeleted++; continue; }

    const avgDaily = salesAvg[asin] || 0;

    // 死にSKU（在庫0 かつ 日販0）除外
    if (inv.qty === 0 && avgDaily === 0) continue;

    const daysLeft = avgDaily > 0 ? (inv.qty / avgDaily) : (inv.qty > 0 ? 999 : 0);
    const status = statusFromDaysLeft(daysLeft, inv.qty);
    const name = master.name || inv.name || '';

    rows.push([
      nowStr, asin, inv.sku, name,
      inv.qty, avgDaily.toFixed(1), daysLeft.toFixed(1),
      status, '',
    ]);

    // マイルストーン判定（売れている商品のみ対象）
    if (avgDaily > 0) {
      const alert = evaluateStockMilestone(asin, daysLeft, alertState, now);
      if (alert) {
        const displayName = (name || inv.sku || '').substring(0, 30);
        const daysText = daysLeft < 1 ? '在庫切れ' : 'あと' + Math.floor(daysLeft) + '日';
        alerts.push({
          key: 'STOCK_' + asin + '_' + alert.threshold,
          type: alert.threshold === 0 ? '在庫切れ' : '在庫少なめ',
          line: displayName + '：残り' + inv.qty + '個／' + daysText,
        });
      }
    }
  }

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, D6_INVENTORY_HEADERS.length).setValues(rows);
    sheet.getRange(2, 1, rows.length, D6_INVENTORY_HEADERS.length).sort({ column: 7, ascending: true });
  }

  saveStockAlertState(alertState);

  Logger.log('✅ 在庫 ' + rows.length + ' 件 / アラート ' + alerts.length + ' 件 / カタログ削除除外 ' +
             skippedDeleted + ' 件 (' + (Date.now() - t0) + 'ms)');

  if (alerts.length > 0) notifyStockAlerts(alerts);
}

/**
 * マイルストーン判定: 残り日数に応じた通知を発火すべきか判断
 *  - 20, 15, 7, 0 の各境界を「初めて下回った」ときのみ発火
 *  - daysLeft > STOCK_RESET_DAYS (20) に回復したら state をクリア（再通知準備）
 *
 * @returns {Object|null} { threshold: number } | null
 */
function evaluateStockMilestone(asin, daysLeft, state, now) {
  // 回復判定: 20日超に戻ったら送信履歴をリセット
  if (daysLeft > STOCK_RESET_DAYS) {
    if (state[asin]) delete state[asin];
    return null;
  }

  // 該当する最小の閾値を特定（小さいほど深刻）
  let hit = null;
  for (const t of STOCK_THRESHOLDS) {
    if (daysLeft <= t) hit = t;  // 0, 7, 15, 20 の順で上書きされ最終的に最小値が残る
  }
  if (hit === null) return null;

  // 既に同じ or より深刻な閾値を送信済みなら通知しない
  const sent = new Set(state[asin] || []);
  if (sent.has(hit)) return null;
  // 急速な在庫減少で中間閾値をスキップした場合も、過去閾値は「送信済み」扱いで以降抑止
  STOCK_THRESHOLDS.forEach(t => { if (t >= hit) sent.add(t); });
  state[asin] = [...sent];
  return { threshold: hit };
}

function statusFromDaysLeft(daysLeft, qty) {
  if (qty <= 0) return '在庫切れ';
  if (daysLeft < 7) return '緊急';
  if (daysLeft < 15) return '警告';
  if (daysLeft < 20) return '注意';
  return 'OK';
}

function getStockAlertState() {
  const raw = PropertiesService.getScriptProperties().getProperty(STOCK_ALERT_STATE_KEY);
  if (!raw) return {};
  try { return JSON.parse(raw); } catch (e) { return {}; }
}

function saveStockAlertState(state) {
  PropertiesService.getScriptProperties().setProperty(STOCK_ALERT_STATE_KEY, JSON.stringify(state));
}

/**
 * テスト用: アラート送信状態を手動リセット
 * 動作確認で通知を再発火させたいときに使う
 */
function resetStockAlertState() {
  PropertiesService.getScriptProperties().deleteProperty(STOCK_ALERT_STATE_KEY);
  Logger.log('✅ 在庫アラート状態をリセットしました');
}

function notifyStockAlerts(alerts) {
  // マイルストーンで既に一意化されているので LineAlert の重複抑止は不要だが、
  // 念のため同じキー（ASIN+threshold）の連続発火を防ぐ
  const sent = getSentAlertMap();
  const now = Date.now();
  const fresh = alerts.filter(a => {
    const last = sent[a.key];
    return !last || (now - last) > ALERT_DEDUP_HOURS * 3600 * 1000;
  });
  if (fresh.length === 0) return;
  pushLineAlert(formatAlertMessage(fresh, 'Amazon在庫アラート'));
  fresh.forEach(a => { sent[a.key] = now; });
  saveSentAlertMap(sent);
}

/**
 * FBA Inventory API (/fba/inventory/v1/summaries) 経由で在庫を取得
 * レポート型（GET_FBA_MYI_UNSUPPRESSED_INVENTORY_DATA）より
 * タイムアウト耐性が高く、ページングもできる。
 *
 * ## 返却値の数量ロジック（Seller Central UI 準拠）
 *   qty     : 在庫あり + 予約済み内のFC転送（= 事実上出荷待機中の在庫）
 *            = fulfillableQuantity + reservedQuantity.pendingTransshipmentQuantity
 *   inbound : 納品（working + shipped + receiving の合計）
 *            = inboundWorkingQuantity + inboundShippedQuantity + inboundReceivingQuantity
 *
 * 「FC処理中」「お客様の注文」「販売不可」「調査中」は qty に含めない。
 */
function fetchInventoryData() {
  const all = [];
  let nextToken = null;
  let page = 0;
  const maxPages = 50;  // 登録商品が数百ある可能性も考慮

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
      const reserved = detail.reservedQuantity || {};
      const fulfillable = detail.fulfillableQuantity || 0;
      const transshipment = reserved.pendingTransshipmentQuantity || 0;
      const inboundWorking = detail.inboundWorkingQuantity || 0;
      const inboundShipped = detail.inboundShippedQuantity || 0;
      const inboundReceiving = detail.inboundReceivingQuantity || 0;

      all.push({
        asin: s.asin || '',
        sku: s.sellerSku || '',
        // 在庫数（FBA在庫）= 在庫あり + FC転送
        qty: fulfillable + transshipment,
        // 納品中 = 納品 working+shipped+receiving
        inbound: inboundWorking + inboundShipped + inboundReceiving,
        name: s.productName || '',
      });
    }

    // nextToken は payload ではなくルート直下の pagination にある
    nextToken = (res.pagination && res.pagination.nextToken) || null;
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

/**
 * D1 から直近N日の ASIN別 日販平均を算出（発注計画向け・0日込み）
 *
 * getRecentDailySalesByAsin() との違い:
 *   - 分母が「全N日」（売れなかった日も0として算入）
 *   - 発注量を決める際は需要が0の日も平均に含めた値が使いやすい
 *   - 例: 過去7日で14個売れた → 14/7 = 2.0/日
 *
 * @returns {Object} ASIN → avg/day
 */
function getSalesAvgByAsin(days) {
  const dailyData = getDailyDataAll();
  const today = new Date();
  const fmt = d => Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  const start = fmt(new Date(today.getFullYear(), today.getMonth(), today.getDate() - days));
  const end = fmt(new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1));

  const totals = {};
  for (const d of dailyData) {
    if (d.date < start || d.date > end) continue;
    if (!d.asin) continue;
    totals[d.asin] = (totals[d.asin] || 0) + d.units;
  }

  const avg = {};
  for (const asin of Object.keys(totals)) {
    avg[asin] = totals[asin] / days;
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
 * デバッグ: 50件全ての SKU / ASIN / 商品名を一覧表示
 * 現在の商品が含まれているか確認用
 */
function debugInventoryListAll() {
  const inv = fetchInventoryData();
  Logger.log('取得件数: ' + inv.length);
  Logger.log('--- 全件リスト ---');
  inv.forEach((i, idx) => {
    Logger.log('[' + idx + '] ASIN=' + i.asin + ' / SKU=' + i.sku + ' / qty=' + i.qty +
               ' / ' + (i.name || '').substring(0, 40));
  });
}

/**
 * デバッグ: 在庫ゼロ以外を再要求（startDateTimeで最近更新された分のみ）
 */
function debugInventoryRecent() {
  const since = new Date();
  since.setDate(since.getDate() - 90);  // 直近90日に更新された在庫
  const params = {
    granularityType: 'Marketplace',
    granularityId: MARKETPLACE_ID_JP,
    marketplaceIds: MARKETPLACE_ID_JP,
    details: 'true',
    startDateTime: since.toISOString(),
  };
  const res = callSpApi('GET', '/fba/inventory/v1/summaries', params);
  const summaries = (res.payload && res.payload.inventorySummaries) || [];
  Logger.log('90日以内更新あり: ' + summaries.length + ' 件');
  summaries.slice(0, 10).forEach((s, i) => {
    const d = s.inventoryDetails || {};
    Logger.log('[' + i + '] ASIN=' + s.asin + ' / SKU=' + s.sellerSku +
               ' / total=' + s.totalQuantity + ' / fulfillable=' + d.fulfillableQuantity +
               ' / updated=' + s.lastUpdatedTime);
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

// ==========================================================
//  長期在庫リスク（271日以上・AIS 追加保管手数料対象）
// ==========================================================

/**
 * GET_FBA_INVENTORY_PLANNING_DATA レポートから長期在庫（271日以上）の情報を取得
 *
 * @returns {Object} {
 *   totalSkus: 長期在庫を抱えるSKU数,
 *   totalUnits: 対象数量合計,
 *   estimatedFee: 推定AIS追加手数料（円）,
 *   items: [{ sku, asin, name, qty271, qty365, qtyTotal, fee }, ...] （手数料降順）
 * }
 */
function fetchLongTermInventoryRisk() {
  const t0 = Date.now();
  Logger.log('===== 長期在庫リスク取得 開始 =====');

  let content;
  try {
    content = fetchReport('GET_FBA_INVENTORY_PLANNING_DATA');
  } catch (e) {
    Logger.log('❌ GET_FBA_INVENTORY_PLANNING_DATA 取得失敗: ' + e.message);
    return { totalSkus: 0, totalUnits: 0, estimatedFee: 0, items: [] };
  }

  const rows = parseTsv(content);
  if (rows.length <= 1) {
    Logger.log('⚠️ データなし');
    return { totalSkus: 0, totalUnits: 0, estimatedFee: 0, items: [] };
  }

  const headers = rows[0];
  const colIndex = {};
  headers.forEach((h, i) => { colIndex[h.trim().toLowerCase().replace(/[\s-]/g, '_')] = i; });

  const skuCol = findCol(colIndex, ['sku', 'seller_sku']);
  const asinCol = findCol(colIndex, ['asin']);
  const nameCol = findCol(colIndex, ['product_name', 'product']);
  // 271〜365日 / 365日超 の数量
  const age271Col = findCol(colIndex, ['inv_age_271_to_365_days', 'aged_271_365_days', 'inv_age_271_365_days']);
  const age365Col = findCol(colIndex, ['inv_age_365_plus_days', 'aged_365_plus_days']);
  // AIS（Aged Inventory Surcharge）推定手数料（JP は 30日刻みで分割）
  const fee271_300Col = findCol(colIndex, ['estimated_ais_271_300_days', 'estimated_ais_271_to_300_days']);
  const fee301_330Col = findCol(colIndex, ['estimated_ais_301_330_days', 'estimated_ais_301_to_330_days']);
  const fee331_365Col = findCol(colIndex, ['estimated_ais_331_365_days', 'estimated_ais_331_to_365_days', 'estimated_ais_271_to_365_days']);
  const fee365Col = findCol(colIndex, ['estimated_ais_365_plus_days']);
  // 通常月次保管料（AIS抜き）
  const storageCostCol = findCol(colIndex, ['estimated_storage_cost_next_month']);

  Logger.log('カラム: sku=' + skuCol + ', asin=' + asinCol + ', age271=' + age271Col +
             ', age365=' + age365Col + ', fee271-300=' + fee271_300Col +
             ', fee301-330=' + fee301_330Col + ', fee331-365=' + fee331_365Col +
             ', fee365+=' + fee365Col + ', storage=' + storageCostCol);

  const items = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const qty271 = age271Col !== -1 ? (parseInt(row[age271Col]) || 0) : 0;
    const qty365 = age365Col !== -1 ? (parseInt(row[age365Col]) || 0) : 0;
    const fee271_300 = fee271_300Col !== -1 ? (parseFloat(row[fee271_300Col]) || 0) : 0;
    const fee301_330 = fee301_330Col !== -1 ? (parseFloat(row[fee301_330Col]) || 0) : 0;
    const fee331_365 = fee331_365Col !== -1 ? (parseFloat(row[fee331_365Col]) || 0) : 0;
    const fee365 = fee365Col !== -1 ? (parseFloat(row[fee365Col]) || 0) : 0;
    const storageCost = storageCostCol !== -1 ? (parseFloat(row[storageCostCol]) || 0) : 0;
    const qtyTotal = qty271 + qty365;
    const feeTotal = fee271_300 + fee301_330 + fee331_365 + fee365;

    if (qtyTotal > 0 || feeTotal > 0) {
      items.push({
        sku: skuCol !== -1 ? String(row[skuCol] || '').trim() : '',
        asin: asinCol !== -1 ? String(row[asinCol] || '').trim() : '',
        name: nameCol !== -1 ? String(row[nameCol] || '').trim() : '',
        qty271: qty271,
        qty365: qty365,
        qtyTotal: qtyTotal,
        fee: feeTotal,
        storageCost: storageCost,
      });
    }
  }

  // 手数料降順でソート
  items.sort((a, b) => b.fee - a.fee);

  const totalUnits = items.reduce((s, x) => s + x.qtyTotal, 0);
  const estimatedFee = items.reduce((s, x) => s + x.fee, 0);

  Logger.log('✅ 対象SKU: ' + items.length + ' 件 / 対象数量: ' + totalUnits +
             ' / 推定月次手数料: ¥' + Math.round(estimatedFee).toLocaleString() +
             ' (' + (Date.now() - t0) + 'ms)');

  return {
    totalSkus: items.length,
    totalUnits: totalUnits,
    estimatedFee: estimatedFee,
    items: items,
  };
}

/**
 * テスト: 長期在庫リスクを取得してログ出力のみ
 */
function testLongTermInventoryRisk() {
  const risk = fetchLongTermInventoryRisk();
  Logger.log('===== 結果サマリ =====');
  Logger.log('対象SKU: ' + risk.totalSkus + ' / 対象数量: ' + risk.totalUnits +
             ' / 推定手数料: ¥' + Math.round(risk.estimatedFee).toLocaleString());
  Logger.log('--- TOP10（手数料順） ---');
  risk.items.slice(0, 10).forEach((x, i) => {
    Logger.log((i + 1) + '. ' + x.sku + ' / ' + x.asin + ' / ' +
               (x.name || '').substring(0, 30) +
               ' / 271-365: ' + x.qty271 + ' / 365+: ' + x.qty365 +
               ' / 推定: ¥' + Math.round(x.fee).toLocaleString());
  });
}
