/**
 * MoneyForward クラウド会計 OAuth2認証
 *
 * 初回セットアップ:
 *   1. MF会計 > API連携 > アプリ管理 で新規アプリ作成
 *   2. 下記の CLIENT_ID / CLIENT_SECRET を設定
 *   3. リダイレクトURIに getRedirectUri() の結果を登録
 *   4. authorize() を実行 → ブラウザで認証
 *
 * OAuth2ライブラリの追加が必要:
 *   Apps Script エディタ > ライブラリ > 「+」
 *   スクリプトID: 1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF
 *   バージョン: 最新を選択
 */

// ==============================
// 設定（※要変更）
// ==============================
const MF_CLIENT_ID     = PropertiesService.getScriptProperties().getProperty('MF_CLIENT_ID') || '';
const MF_CLIENT_SECRET = PropertiesService.getScriptProperties().getProperty('MF_CLIENT_SECRET') || '';

const MF_API_BASE = 'https://accounting.moneyforward.com/api/v3';
const MF_AUTH_URL = 'https://accounting.moneyforward.com/oauth/authorize';
const MF_TOKEN_URL = 'https://accounting.moneyforward.com/oauth/token';

// ==============================
// OAuth2サービス
// ==============================

function getMFService_() {
  return OAuth2.createService('moneyforward')
    .setAuthorizationBaseUrl(MF_AUTH_URL)
    .setTokenUrl(MF_TOKEN_URL)
    .setClientId(MF_CLIENT_ID)
    .setClientSecret(MF_CLIENT_SECRET)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('office:read account:read journal:read report:read')
    .setParam('access_type', 'offline');
}

/**
 * 認証を開始（この関数を実行してログに表示されるURLを開く）
 */
function authorize() {
  const service = getMFService_();
  if (service.hasAccess()) {
    Logger.log('既に認証済みです。');
  } else {
    const authUrl = service.getAuthorizationUrl();
    Logger.log('以下のURLをブラウザで開いて認証してください:');
    Logger.log(authUrl);
  }
}

/**
 * OAuth2コールバック
 */
function authCallback(request) {
  const service = getMFService_();
  const authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('認証成功！このタブを閉じてください。');
  } else {
    return HtmlService.createHtmlOutput('認証失敗。もう一度お試しください。');
  }
}

/**
 * リダイレクトURIを取得（MFアプリ設定に登録する）
 */
function getRedirectUri() {
  Logger.log('リダイレクトURI:');
  Logger.log(OAuth2.getRedirectUri());
}

/**
 * 認証をリセット
 */
function resetAuth() {
  getMFService_().reset();
  Logger.log('認証をリセットしました。authorize() を再実行してください。');
}

/**
 * MF APIにGETリクエスト
 */
function mfApiGet_(endpoint, params) {
  const service = getMFService_();
  if (!service.hasAccess()) {
    throw new Error('MF未認証です。authorize() を実行してください。');
  }

  var url = MF_API_BASE + endpoint;
  if (params) {
    var qs = Object.keys(params).map(function(k) {
      return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
    }).join('&');
    url += '?' + qs;
  }

  var response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + service.getAccessToken() },
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  if (code !== 200) {
    throw new Error('MF API error (' + code + '): ' + response.getContentText());
  }

  return JSON.parse(response.getContentText());
}

/**
 * 事業所IDを取得・保存
 */
function getOfficeId_() {
  var props = PropertiesService.getScriptProperties();
  var officeId = props.getProperty('MF_OFFICE_ID');
  if (officeId) return officeId;

  var data = mfApiGet_('/offices');
  if (data.offices && data.offices.length > 0) {
    officeId = data.offices[0].id;
    props.setProperty('MF_OFFICE_ID', officeId);
    Logger.log('事業所ID: ' + officeId + ' を保存しました');
    return officeId;
  }
  throw new Error('事業所が見つかりません');
}

/**
 * client_id / client_secret をスクリプトプロパティに保存
 * 初回セットアップ用（手動実行）
 */
function setMFCredentials() {
  var ui = SpreadsheetApp.getUi();
  var clientId = ui.prompt('MF Client ID を入力').getResponseText();
  var clientSecret = ui.prompt('MF Client Secret を入力').getResponseText();

  if (clientId && clientSecret) {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('MF_CLIENT_ID', clientId);
    props.setProperty('MF_CLIENT_SECRET', clientSecret);
    Logger.log('認証情報を保存しました。authorize() を実行してください。');
  }
}
