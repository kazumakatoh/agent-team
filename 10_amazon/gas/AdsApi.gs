/**
 * Amazon Dashboard - Amazon Ads API 共通モジュール（Phase 3）
 *
 * AdsAuth.gs と対をなす。AdsAuth はトークン発行手順（ユーザー操作）用で、
 * こちらは日々の API 呼び出しで使う：
 *   - access_token をキャッシュして再利用（有効期限内は再取得しない）
 *   - 認証ヘッダ（Bearer / ClientId / Scope=ProfileId）をまとめて付与する callAdsApi()
 *
 * v3 Reporting API の非同期レポート（作成→ポーリング→DL）は AdsReport.gs 側に置く。
 */

// ===== access_token キャッシュ =====
// LWA の access_token は 3600 秒で失効。余裕を見て 50 分で再取得する。
const ADS_TOKEN_TTL_MS = 50 * 60 * 1000;

/**
 * access_token をキャッシュ付きで取得
 * 有効期限内なら Script Properties のキャッシュを返す。
 * 失効していれば refresh_token から再取得して保存。
 */
function getAdsAccessTokenCached() {
  const props = PropertiesService.getScriptProperties();
  const cached = props.getProperty('ADS_ACCESS_TOKEN');
  const expireAt = parseInt(props.getProperty('ADS_ACCESS_TOKEN_EXPIRE_AT') || '0');

  if (cached && expireAt && Date.now() < expireAt) {
    return cached;
  }

  const token = getAdsAccessToken();  // AdsAuth.gs の refresh_token→access_token 交換
  props.setProperty('ADS_ACCESS_TOKEN', token);
  props.setProperty('ADS_ACCESS_TOKEN_EXPIRE_AT', String(Date.now() + ADS_TOKEN_TTL_MS));
  return token;
}

/**
 * Ads API 共通ヘッダを生成
 * @param {string} [contentType] - POST 用の Content-Type（v3 reporting は vnd 形式が必要）
 */
function getAdsHeaders(contentType) {
  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('ADS_CLIENT_ID');
  const profileId = props.getProperty('ADS_PROFILE_ID');

  if (!clientId) throw new Error('ADS_CLIENT_ID 未設定');
  if (!profileId) throw new Error('ADS_PROFILE_ID 未設定。testAdsProfiles() を先に実行してください');

  const headers = {
    'Authorization': 'Bearer ' + getAdsAccessTokenCached(),
    'Amazon-Advertising-API-ClientId': clientId,
    'Amazon-Advertising-API-Scope': profileId,
  };
  if (contentType) headers['Content-Type'] = contentType;
  return headers;
}

/**
 * Amazon Ads API 呼び出し共通ラッパー
 *
 * @param {string} method  - HTTP method (GET/POST/PUT)
 * @param {string} path    - '/reporting/reports' 等のパス
 * @param {Object} [body]  - リクエストボディ（オブジェクト）
 * @param {Object} [opts]  - { contentType, accept, retry }
 * @returns {Object|string} JSON レスポンス（JSONでなければ文字列）
 */
function callAdsApi(method, path, body, opts) {
  opts = opts || {};
  const url = ADS_API_ENDPOINT + path;
  const headers = getAdsHeaders(opts.contentType);
  if (opts.accept) headers['Accept'] = opts.accept;

  const options = {
    method: method.toLowerCase(),
    headers: headers,
    muteHttpExceptions: true,
  };
  if (body && (method === 'POST' || method === 'PUT')) {
    options.payload = JSON.stringify(body);
  }

  const res = UrlFetchApp.fetch(url, options);
  const status = res.getResponseCode();
  const text = res.getContentText();

  // レート制限 → 2秒待ってリトライ（1回のみ）
  if (status === 429 && !opts.retry) {
    Utilities.sleep(2000);
    return callAdsApi(method, path, body, Object.assign({}, opts, { retry: true }));
  }

  // 401 → access_token を一度破棄してリトライ（1回のみ）
  if (status === 401 && !opts.retry) {
    PropertiesService.getScriptProperties().deleteProperty('ADS_ACCESS_TOKEN');
    return callAdsApi(method, path, body, Object.assign({}, opts, { retry: true }));
  }

  if (status < 200 || status >= 300) {
    throw new Error('Ads API エラー ' + status + ' ' + method + ' ' + path + ': ' + text);
  }

  try {
    return JSON.parse(text);
  } catch (e) {
    return text;
  }
}
