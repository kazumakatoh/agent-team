/**
 * MoneyForward クラウド会計 認証管理
 *
 * OAuth2ライブラリを使わず、スクリプトプロパティでトークンを管理。
 *
 * セットアップ:
 *   1. MF会計 > API連携 > アプリ管理 で新規アプリ作成
 *   2. リダイレクトURIに https://script.google.com/macros/d/{SCRIPT_ID}/usercallback を登録
 *   3. スプシの MF連携 > 「MF認証情報を設定」で Client ID / Secret を入力
 *   4. MF連携 > 「MF認証を実行」→ ログに表示されるURLを開いて認証
 *   5. 表示される認可コードを MF連携 > 「認可コードを入力」で入力
 */

var MF_API_BASE  = 'https://accounting.moneyforward.com/api/v3';
var MF_AUTH_URL  = 'https://accounting.moneyforward.com/oauth/authorize';
var MF_TOKEN_URL = 'https://accounting.moneyforward.com/oauth/token';

/**
 * client_id / client_secret をスクリプトプロパティに保存
 */
function setMFCredentials() {
  var ui = SpreadsheetApp.getUi();
  var r1 = ui.prompt('MF Client ID を入力');
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var r2 = ui.prompt('MF Client Secret を入力');
  if (r2.getSelectedButton() !== ui.Button.OK) return;

  var props = PropertiesService.getScriptProperties();
  props.setProperty('MF_CLIENT_ID', r1.getResponseText().trim());
  props.setProperty('MF_CLIENT_SECRET', r2.getResponseText().trim());
  ui.alert('認証情報を保存しました。次に「MF認証を実行」してください。');
}

/**
 * 認証URLを表示
 */
function authorize() {
  var props = PropertiesService.getScriptProperties();
  var clientId = props.getProperty('MF_CLIENT_ID');
  if (!clientId) {
    SpreadsheetApp.getUi().alert('先に「MF認証情報を設定」でClient IDを設定してください。');
    return;
  }

  var redirectUri = 'urn:ietf:wg:oauth:2.0:oob';
  var authUrl = MF_AUTH_URL
    + '?client_id=' + encodeURIComponent(clientId)
    + '&redirect_uri=' + encodeURIComponent(redirectUri)
    + '&response_type=code'
    + '&scope=' + encodeURIComponent('office:read account:read journal:read report:read');

  var html = HtmlService.createHtmlOutput(
    '<p>以下のURLをブラウザで開いて認証してください：</p>'
    + '<p><a href="' + authUrl + '" target="_blank">' + authUrl + '</a></p>'
    + '<p>認証後に表示される認可コードをコピーし、<br>MF連携 > 「認可コードを入力」で貼り付けてください。</p>'
  ).setWidth(600).setHeight(200);

  SpreadsheetApp.getUi().showModalDialog(html, 'MF認証');
}

/**
 * 認可コードを入力してトークンを取得
 */
function inputAuthCode() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt('MF認証', '認可コードを入力してください：', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;

  var code = result.getResponseText().trim();
  var props = PropertiesService.getScriptProperties();
  var clientId = props.getProperty('MF_CLIENT_ID');
  var clientSecret = props.getProperty('MF_CLIENT_SECRET');

  var response = UrlFetchApp.fetch(MF_TOKEN_URL, {
    method: 'post',
    payload: {
      grant_type: 'authorization_code',
      client_id: clientId,
      client_secret: clientSecret,
      redirect_uri: 'urn:ietf:wg:oauth:2.0:oob',
      code: code
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    ui.alert('トークン取得に失敗しました: ' + response.getContentText());
    return;
  }

  var tokenData = JSON.parse(response.getContentText());
  props.setProperty('MF_ACCESS_TOKEN', tokenData.access_token);
  if (tokenData.refresh_token) {
    props.setProperty('MF_REFRESH_TOKEN', tokenData.refresh_token);
  }

  // 事業所IDも取得・保存
  fetchAndSaveOfficeId_();

  ui.alert('認証成功！MFデータの同期が可能になりました。');
}

/**
 * アクセストークンをリフレッシュ
 */
function refreshToken_() {
  var props = PropertiesService.getScriptProperties();
  var refreshToken = props.getProperty('MF_REFRESH_TOKEN');
  if (!refreshToken) {
    throw new Error('リフレッシュトークンがありません。「MF認証を実行」からやり直してください。');
  }

  var response = UrlFetchApp.fetch(MF_TOKEN_URL, {
    method: 'post',
    payload: {
      grant_type: 'refresh_token',
      client_id: props.getProperty('MF_CLIENT_ID'),
      client_secret: props.getProperty('MF_CLIENT_SECRET'),
      refresh_token: refreshToken
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('トークンリフレッシュ失敗: ' + response.getContentText());
  }

  var tokenData = JSON.parse(response.getContentText());
  props.setProperty('MF_ACCESS_TOKEN', tokenData.access_token);
  if (tokenData.refresh_token) {
    props.setProperty('MF_REFRESH_TOKEN', tokenData.refresh_token);
  }
  return tokenData.access_token;
}

/**
 * MF APIにGETリクエスト
 */
function mfApiGet_(endpoint, params) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('MF_ACCESS_TOKEN');
  if (!token) {
    throw new Error('MF未認証です。MF連携メニューから認証してください。');
  }

  var url = MF_API_BASE + endpoint;
  if (params) {
    var qs = Object.keys(params).map(function(k) {
      return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
    }).join('&');
    url += '?' + qs;
  }

  var response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });

  // 401 → トークンリフレッシュして再試行
  if (response.getResponseCode() === 401) {
    token = refreshToken_();
    response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
  }

  if (response.getResponseCode() !== 200) {
    throw new Error('MF API error (' + response.getResponseCode() + '): ' + response.getContentText());
  }

  return JSON.parse(response.getContentText());
}

/**
 * 事業所IDを取得して保存
 */
function fetchAndSaveOfficeId_() {
  var data = mfApiGet_('/offices');
  if (data.offices && data.offices.length > 0) {
    var officeId = data.offices[0].id;
    PropertiesService.getScriptProperties().setProperty('MF_OFFICE_ID', officeId);
    Logger.log('事業所ID: ' + officeId);
    return officeId;
  }
  throw new Error('事業所が見つかりません');
}

/**
 * 事業所IDを取得
 */
function getOfficeId_() {
  var officeId = PropertiesService.getScriptProperties().getProperty('MF_OFFICE_ID');
  if (officeId) return officeId;
  return fetchAndSaveOfficeId_();
}

/**
 * 認証をリセット
 */
function resetAuth() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('MF_ACCESS_TOKEN');
  props.deleteProperty('MF_REFRESH_TOKEN');
  props.deleteProperty('MF_OFFICE_ID');
  SpreadsheetApp.getUi().alert('認証をリセットしました。');
}
