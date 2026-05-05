/**
 * Amazon Dashboard - 日次データ取得モジュール
 *
 * 毎日6:00にトリガーで実行
 * Orders API から前日の注文データを取得し、D1シートに書き込む
 */

/**
 * メイン: 日次データ取得＆書き込み
 * トリガーから呼び出される
 */
function dailyFetch() {
  Logger.log('===== 日次データ取得 開始 =====');

  // 前日の日付範囲
  const yesterday = getYesterday();
  const startOfDay = yesterday + 'T00:00:00+09:00';
  const endOfDay = yesterday + 'T23:59:59+09:00';

  Logger.log('対象日: ' + yesterday);

  // ヘッダーが未設定なら設定
  setupDailyDataHeaders();

  // 注文データ取得
  const orderSummary = fetchAndWriteOrders(yesterday, startOfDay, endOfDay);

  Logger.log('===== 日次データ取得 完了 =====');
  Logger.log('取得ASIN数: ' + Object.keys(orderSummary).length);
}

/**
 * 前日の日付文字列を取得 (YYYY-MM-DD)
 */
function getYesterday() {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
}

/**
 * 今日の日付文字列を取得 (YYYY-MM-DD)
 */
function getToday() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
}

/**
 * 注文データを取得してASIN別に集計し、D1シートに書き込む
 */
function fetchAndWriteOrders(dateStr, startOfDay, endOfDay) {
  Logger.log('注文データ取得中...');

  // Orders API で注文一覧を取得
  const orders = getOrders(startOfDay, endOfDay);
  Logger.log('注文件数: ' + orders.length);

  if (orders.length === 0) {
    Logger.log('注文なし。スキップ。');
    return {};
  }

  // 注文ごとにアイテム詳細を取得してASIN別に集計
  const asinSummary = {};
  const productMaster = getProductMasterMap();
  const newAsins = [];

  for (const order of orders) {
    try {
      // 注文アイテムを取得
      const items = getOrderItems(order.AmazonOrderId);

      for (const item of items) {
        const asin = item.ASIN;
        if (!asinSummary[asin]) {
          asinSummary[asin] = {
            sales: 0,
            cvCount: 0,    // 注文件数（CV）
            units: 0,      // 注文点数
          };
        }

        const price = parseFloat(item.ItemPrice?.Amount || 0);
        const qty = parseInt(item.QuantityOrdered || 0);

        asinSummary[asin].sales += price;
        asinSummary[asin].units += qty;

        // CV は注文単位でカウント（同一注文内の複数アイテムは1CVとしない）
        // ここでは簡易的にアイテム単位でカウント
        asinSummary[asin].cvCount += 1;

        // 商品マスターに未登録のASINを検出
        if (!productMaster[asin] && !newAsins.includes(asin)) {
          newAsins.push(asin);
        }
      }
    } catch (e) {
      Logger.log('⚠️ 注文 ' + order.AmazonOrderId + ' のアイテム取得失敗: ' + e.message);
    }

    // レート制限対策
    Utilities.sleep(500);
  }

  // 仕入単価マップを読み込み（M1 商品マスター から取得）
  const priceMap = {};
  for (const [asin, m] of Object.entries(productMaster)) {
    if (m.purchasePrice > 0) priceMap[asin] = m.purchasePrice;
  }

  // D1シートに書き込み
  const rows = [];
  for (const [asin, data] of Object.entries(asinSummary)) {
    const master = productMaster[asin] || {};
    const unitPrice = priceMap[asin] || 0;
    const cogs = unitPrice * data.units;

    rows.push([
      dateStr,                          // 日付
      asin,                             // ASIN
      master.name || '',                // 商品名
      master.category || '',            // カテゴリ
      Math.round(data.sales),           // 売上金額
      data.cvCount,                     // CV(注文件数)
      data.units,                       // 注文点数
      '', '', '', '', '',               // セッション・PV等（Traffic Reportで後日更新）
      '', '', '',                       // FBA手数料・返品（Settlement Reportで後日更新）
      '', '', '', '',                   // 広告データ（Ads APIで後日更新）
      unitPrice || '',                  // 仕入単価
      cogs || '',                       // 仕入原価合計
      '暫定',                           // ステータス
    ]);
  }

  if (rows.length > 0) {
    appendRows(SHEET_NAMES.D1_DAILY, rows);
    Logger.log('✅ D1 日次データ: ' + rows.length + ' 行書き込み完了');
  }

  // 新規ASINを商品マスターに追加（11列構成）
  if (newAsins.length > 0) {
    const masterRows = newAsins.map(asin => [
      asin, '', '', 'アクティブ',
      '', '', '', '', '', '',
      '自動検出 ' + dateStr,
    ]);
    appendRows(SHEET_NAMES.M1_PRODUCT_MASTER, masterRows);
    applyProductMasterFormulas(getOrCreateSheet(SHEET_NAMES.M1_PRODUCT_MASTER));
    Logger.log('✅ M1 商品マスター: 新規ASIN ' + newAsins.length + ' 件追加');
  }

  return asinSummary;
}

/**
 * 注文のアイテム詳細を取得
 */
function getOrderItems(orderId) {
  const result = callSpApi('GET', '/orders/v0/orders/' + orderId + '/orderItems');
  return result.payload ? result.payload.OrderItems : [];
}

/**
 * テスト: 日次データ取得を実行
 */
function testDailyFetch() {
  dailyFetch();
}

/**
 * テスト: 指定日の注文を取得して表示（書き込みなし）
 */
function testOrdersPreview() {
  const yesterday = getYesterday();
  Logger.log('対象日: ' + yesterday);

  const startOfDay = yesterday + 'T00:00:00+09:00';
  const endOfDay = yesterday + 'T23:59:59+09:00';

  const orders = getOrders(startOfDay, endOfDay);
  Logger.log('注文件数: ' + orders.length);

  // 先頭5件の詳細を表示
  const limit = Math.min(orders.length, 5);
  for (let i = 0; i < limit; i++) {
    const order = orders[i];
    Logger.log('---');
    Logger.log('注文ID: ' + order.AmazonOrderId);
    Logger.log('注文日: ' + order.PurchaseDate);
    Logger.log('合計: ' + (order.OrderTotal ? order.OrderTotal.Amount + ' ' + order.OrderTotal.CurrencyCode : '不明'));
    Logger.log('ステータス: ' + order.OrderStatus);
  }
}
