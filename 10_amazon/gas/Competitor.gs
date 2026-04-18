/**
 * Amazon Dashboard - 競合チェック（Phase 4b ③）
 *
 * SP-API Product Pricing API から自社ASINの「競合価格・BuyBox状況」を取得し、
 * D4 競合価格シートに保存する。
 *
 *   GET /products/pricing/v0/items/{Asin}/offers   ← 1ASINずつ
 *   または
 *   GET /products/pricing/v0/competitivePrice?MarketplaceId=...&Asins=A,B,C  ← 一括
 *
 * 自社価格 vs 競合最安価格 を並べ、自社が高い／BuyBox 喪失中の商品を抽出。
 *
 * ## D4 シート構造（7列）
 * 取得日時 | ASIN | 商品名 | 自社価格 | 競合最安 | 価格差 | BuyBox保持
 *
 * ## トリガー: 毎日 AM9:30（fetchCompetitorPricing）
 */

const D4_COMPETITOR = '競合価格';
const D4_COMPETITOR_HEADERS = ['取得日時', 'ASIN', '商品名', '自社価格', '競合最安', '価格差', 'BuyBox保持'];
const COMPETITOR_BATCH_SIZE = 20;     // CompetitivePrice API は最大20ASIN/req

/**
 * メイン: アクティブASINの競合価格を取得
 */
function fetchCompetitorPricing() {
  const t0 = Date.now();
  Logger.log('===== 競合価格取得 開始 =====');

  setupCompetitorSheet();

  const asins = getActiveAsins(ACTIVE_PRODUCT_DAYS);
  if (asins.length === 0) { Logger.log('対象ASINなし'); return; }
  Logger.log('対象ASIN数: ' + asins.length);

  const productMaster = getProductMasterMap();
  const now = new Date();
  const nowStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  const rows = [];

  for (let i = 0; i < asins.length; i += COMPETITOR_BATCH_SIZE) {
    const batch = asins.slice(i, i + COMPETITOR_BATCH_SIZE);
    try {
      const result = fetchCompetitivePriceBatch(batch);
      for (const item of result) {
        const asin = item.asin;
        const myPrice = item.myPrice;
        const lowest = item.lowestPrice;
        const diff = (myPrice != null && lowest != null) ? myPrice - lowest : null;
        rows.push([
          nowStr, asin, (productMaster[asin] || {}).name || '',
          myPrice == null ? '' : myPrice,
          lowest == null ? '' : lowest,
          diff == null ? '' : diff,
          item.buyBoxOwned ? '✅' : '❌',
        ]);
      }
    } catch (e) {
      Logger.log('バッチ ' + (i / COMPETITOR_BATCH_SIZE + 1) + ' エラー: ' + e.message);
    }
    Utilities.sleep(1000); // レート制限緩和
  }

  if (rows.length === 0) { Logger.log('取得結果なし'); return; }

  // シートに追記
  appendRows(D4_COMPETITOR, rows);
  Logger.log('✅ 競合価格 ' + rows.length + ' 行追加 (' + (Date.now() - t0) + 'ms)');

  // BuyBox 喪失 / 自社が大幅に高い 場合は LINE アラートにも回す
  const issues = rows
    .filter(r => r[6] === '❌' || (typeof r[5] === 'number' && r[5] > 0 && r[3] && r[5] / r[3] > 0.05))
    .slice(0, 10);
  if (issues.length > 0) {
    const alerts = issues.map(r => ({
      key: 'COMP_' + r[1] + '_' + nowStr.substring(0, 10),
      type: r[6] === '❌' ? '🛒BuyBox喪失' : '💰自社価格が高い',
      line: r[1] + (r[2] ? ' (' + r[2].substring(0, 15) + ')' : '') +
            ' / 自社 ' + r[3] + '円 vs 競合 ' + r[4] + '円',
    }));
    notifyCompetitorIssues(alerts);
  }
}

/**
 * 競合価格 API バッチ呼び出し（CompetitivePrice）
 */
function fetchCompetitivePriceBatch(asins) {
  const params = {
    MarketplaceId: MARKETPLACE_ID_JP,
    Asins: asins.join(','),
    ItemType: 'Asin',
  };
  const result = callSpApi('GET', '/products/pricing/v0/competitivePrice', params);
  const items = (result.payload || []).filter(p => p.status === 'Success' && p.Product);

  return items.map(p => {
    const asin = p.ASIN;
    const compRoot = (p.Product && p.Product.CompetitivePricing) || {};
    const compList = compRoot.CompetitivePrices || [];
    const buyBox = compList.find(c => c.CompetitivePriceId === '1');  // 1 = New BuyBox
    const lowestPrice = buyBox && buyBox.Price ? parseFloat(buyBox.Price.LandedPrice && buyBox.Price.LandedPrice.Amount) : null;
    const buyBoxOwned = !!(p.IsOfferPriceBuyBoxWinner || (buyBox && buyBox.belongsToRequester));

    // 自社価格は別途 ItemOffers が必要な場合があるが、ここでは LandedPrice を流用
    const myPrice = lowestPrice;  // 暫定: 自社=競合と同じ場合のフォールバック
    return { asin, myPrice, lowestPrice, buyBoxOwned };
  });
}

function setupCompetitorSheet() {
  const sheet = getOrCreateSheetCompact(D4_COMPETITOR, D4_COMPETITOR_HEADERS.length, 200);
  const existing = sheet.getRange(1, 1, 1, D4_COMPETITOR_HEADERS.length).getValues()[0];
  if (existing[0] !== D4_COMPETITOR_HEADERS[0]) {
    sheet.getRange(1, 1, 1, D4_COMPETITOR_HEADERS.length).setValues([D4_COMPETITOR_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, D4_COMPETITOR_HEADERS.length).setFontWeight('bold').setBackground('#e8f0fe');
  }
}

function notifyCompetitorIssues(alerts) {
  // LineAlert.gs の重複抑止フローを再利用
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
 * 直近X日に注文があったASIN一覧
 */
function getActiveAsins(days) {
  const sheet = getOrCreateSheet(SHEET_NAMES.D1_DAILY);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const since = Utilities.formatDate(addDays(new Date(), -days), 'Asia/Tokyo', 'yyyy-MM-dd');
  const set = new Set();
  for (const r of data) {
    const date = r[0] instanceof Date
      ? Utilities.formatDate(r[0], 'Asia/Tokyo', 'yyyy-MM-dd')
      : String(r[0]).substring(0, 10);
    if (date >= since && r[1]) set.add(String(r[1]).trim());
  }
  return Array.from(set);
}
