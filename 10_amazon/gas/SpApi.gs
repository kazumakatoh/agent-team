/**
 * Amazon Dashboard - SP-API データ取得モジュール
 *
 * Phase 1 で実装する各種データ取得関数
 */

/**
 * SP-API にリクエストを送信する共通関数
 *
 * @param {string} method - HTTP メソッド (GET/POST/PUT)
 * @param {string} path - APIパス (例: /orders/v0/orders)
 * @param {Object} [queryParams] - クエリパラメータ
 * @param {Object} [body] - リクエストボディ
 * @returns {Object} レスポンスデータ
 */
function callSpApi(method, path, queryParams, body, retryAttempt) {
  retryAttempt = retryAttempt || 0;
  const token = getSpApiAccessToken();
  let url = SP_API_ENDPOINT + path;

  // クエリパラメータを追加
  if (queryParams) {
    const params = Object.entries(queryParams)
      .filter(([_, v]) => v !== undefined && v !== null)
      .map(([k, v]) => encodeURIComponent(k) + '=' + encodeURIComponent(v))
      .join('&');
    if (params) url += '?' + params;
  }

  const options = {
    method: method.toLowerCase(),
    headers: {
      'x-amz-access-token': token,
      'Content-Type': 'application/json',
    },
    muteHttpExceptions: true,
  };

  if (body && (method === 'POST' || method === 'PUT')) {
    options.payload = JSON.stringify(body);
  }

  const response = UrlFetchApp.fetch(url, options);
  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();
  let responseBody;
  try { responseBody = JSON.parse(responseText); } catch (e) { responseBody = { raw: responseText }; }

  // QuotaExceeded（帯域・スロットリング）の文言検知
  const lowerBody = (responseText || '').toLowerCase();
  const isQuotaErr = lowerBody.indexOf('quotaexceeded') >= 0
                     || lowerBody.indexOf('bandwidth') >= 0
                     || lowerBody.indexOf('帯域') >= 0;

  // 429 / 503 / Quota 系は指数バックオフで最大4回リトライ
  if ((statusCode === 429 || statusCode === 503 || isQuotaErr) && retryAttempt < 4) {
    const waitMs = 2000 * Math.pow(2, retryAttempt); // 2s → 4s → 8s → 16s
    Logger.log('⏳ SP-API throttled (status=' + statusCode + ', quota=' + isQuotaErr +
               '). ' + waitMs + 'ms 待機後リトライ ' + (retryAttempt + 1) + '/4 …');
    Utilities.sleep(waitMs);
    return callSpApi(method, path, queryParams, body, retryAttempt + 1);
  }

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('SP-API エラー: HTTP ' + statusCode + ' - ' + JSON.stringify(responseBody));
  }

  return responseBody;
}

// ===== レポート関連 =====

/**
 * SP-API レポートを作成リクエスト
 *
 * @param {string} reportType - レポート種別
 * @param {string} [startDate] - 開始日 (ISO 8601)
 * @param {string} [endDate] - 終了日 (ISO 8601)
 * @returns {string} reportId
 */
function createReport(reportType, startDate, endDate) {
  const body = {
    reportType: reportType,
    marketplaceIds: [MARKETPLACE_ID_JP],
  };

  if (startDate) body.dataStartTime = startDate;
  if (endDate) body.dataEndTime = endDate;

  const result = callSpApi('POST', '/reports/2021-06-30/reports', null, body);
  Logger.log('レポート作成リクエスト送信: ' + reportType + ' → reportId: ' + result.reportId);
  return result.reportId;
}

/**
 * レポートのステータスを確認
 *
 * @param {string} reportId
 * @returns {Object} レポート情報 (processingStatus, reportDocumentId 等)
 */
function getReportStatus(reportId) {
  return callSpApi('GET', '/reports/2021-06-30/reports/' + reportId);
}

/**
 * レポートドキュメントをダウンロード
 *
 * @param {string} reportDocumentId
 * @returns {string} レポートの内容（TSV/CSV形式の文字列）
 */
function downloadReportDocument(reportDocumentId) {
  const docInfo = callSpApi('GET', '/reports/2021-06-30/documents/' + reportDocumentId);
  const response = UrlFetchApp.fetch(docInfo.url);

  // gzip圧縮されている場合は解凍
  if (docInfo.compressionAlgorithm === 'GZIP') {
    const blob = response.getBlob().setContentType('application/x-gzip');
    const decompressed = Utilities.ungzip(blob);
    return decompressed.getDataAsString('UTF-8');
  }

  return response.getContentText();
}



/**
 * レポートを作成→完了まで待機→ダウンロード
 * GASの6分制限に注意: 長時間かかる場合は分割実行が必要
 *
 * @param {string} reportType
 * @param {string} [startDate]
 * @param {string} [endDate]
 * @returns {string} レポート内容
 */
function fetchReport(reportType, startDate, endDate) {
  const reportId = createReport(reportType, startDate, endDate);

  // ポーリング（最大5分）
  const maxAttempts = 30;
  for (let i = 0; i < maxAttempts; i++) {
    Utilities.sleep(10000); // 10秒待機
    const status = getReportStatus(reportId);
    Logger.log('レポートステータス: ' + status.processingStatus + ' (試行 ' + (i + 1) + '/' + maxAttempts + ')');

    if (status.processingStatus === 'DONE') {
      return downloadReportDocument(status.reportDocumentId);
    }
    if (status.processingStatus === 'FATAL' || status.processingStatus === 'CANCELLED') {
      throw new Error('レポート生成失敗: ' + status.processingStatus);
    }
  }

  // タイムアウト: reportIdを保存して後で再取得
  PropertiesService.getScriptProperties().setProperty('PENDING_REPORT_' + reportType, reportId);
  throw new Error('レポート生成タイムアウト。reportId を保存しました: ' + reportId);
}

// ===== 注文データ =====

/**
 * 指定期間の注文一覧を取得
 *
 * @param {string} createdAfter - 開始日時 (ISO 8601)
 * @param {string} [createdBefore] - 終了日時 (ISO 8601)
 * @returns {Array} 注文リスト
 */
function getOrders(createdAfter, createdBefore) {
  const params = {
    MarketplaceIds: MARKETPLACE_ID_JP,
    CreatedAfter: createdAfter,
    OrderStatuses: 'Shipped,Unshipped',
  };
  if (createdBefore) params.CreatedBefore = createdBefore;

  const result = callSpApi('GET', '/orders/v0/orders', params);
  return result.payload ? result.payload.Orders : [];
}

// ===== テスト関数 =====

/**
 * SP-API 接続テスト: 直近の注文を取得
 */
function testGetOrders() {
  try {
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    const orders = getOrders(yesterday.toISOString());
    Logger.log('✅ 直近の注文: ' + orders.length + ' 件');
    if (orders.length > 0) {
      Logger.log('  最新注文ID: ' + orders[0].AmazonOrderId);
    }
  } catch (e) {
    Logger.log('❌ エラー: ' + e.message);
  }
}

/**
 * SP-API 接続テスト: Sales and Traffic レポート
 */
function testSalesAndTrafficReport() {
  try {
    const endDate = new Date();
    const startDate = new Date();
    startDate.setDate(startDate.getDate() - 7);

    const reportContent = fetchReport(
      'GET_SALES_AND_TRAFFIC_REPORT',
      startDate.toISOString().split('T')[0],
      endDate.toISOString().split('T')[0]
    );
    Logger.log('✅ Sales & Traffic レポート取得成功');
    Logger.log('先頭500文字: ' + reportContent.substring(0, 500));
  } catch (e) {
    Logger.log('❌ エラー: ' + e.message);
  }
}
