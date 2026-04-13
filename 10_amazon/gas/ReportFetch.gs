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
  const endDate = new Date();
  const startDate = new Date();
  startDate.setDate(startDate.getDate() - 60);

  const start = Utilities.formatDate(startDate, 'Asia/Tokyo', 'yyyy-MM-dd');
  const end = Utilities.formatDate(endDate, 'Asia/Tokyo', 'yyyy-MM-dd');

  fetchOrdersByReport(start, end);
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
