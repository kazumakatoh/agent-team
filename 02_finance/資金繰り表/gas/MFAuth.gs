/**
 * MoneyForward クラウド会計 認証管理
 *
 * MF API (api-accounting.moneyforward.com) とOAuth2連携。
 * トークンはスクリプトプロパティで管理。
 */

var MF_API_BASE  = 'https://api-accounting.moneyforward.com/api/v3';
var MF_AUTH_URL  = 'https://api.biz.moneyforward.com/authorize';
var MF_TOKEN_URL = 'https://api.biz.moneyforward.com/token';
var MF_REDIRECT_URI = 'https://script.google.com/a/level1.biz/macros/s/AKfycbwdMoCURYjFU0rKhUmHLyR37sX_EapagMk0po1l6o3V0kj-nn0/exec';

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

function authorize() {
  var props = PropertiesService.getScriptProperties();
  var clientId = props.getProperty('MF_CLIENT_ID');
  if (!clientId) {
    SpreadsheetApp.getUi().alert('先に「MF認証情報を設定」でClient IDを設定してください。');
    return;
  }
  var authUrl = MF_AUTH_URL
    + '?client_id=' + encodeURIComponent(clientId)
    + '&redirect_uri=' + encodeURIComponent(MF_REDIRECT_URI)
    + '&response_type=code'
    + '&scope=' + encodeURIComponent('mfc/accounting/offices.read mfc/accounting/journal.read mfc/accounting/accounts.read mfc/accounting/report.read');
  var html = HtmlService.createHtmlOutput(
    '<p>以下のリンクをクリックしてMF認証してください：</p>'
    + '<p><a href="' + authUrl + '" target="_blank" style="font-size:14px;">MF認証を開始する</a></p>'
    + '<p>認証後、リダイレクト先URLの code= の値をコピーし、<br>MF連携 > ③認可コードを入力 で貼り付けてください。</p>'
  ).setWidth(550).setHeight(180);
  SpreadsheetApp.getUi().showModalDialog(html, 'MF認証');
}

function inputAuthCode() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt('MF認証', '認可コードを入力してください：', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;
  var code = result.getResponseText().trim();
  var props = PropertiesService.getScriptProperties();
  var response = UrlFetchApp.fetch(MF_TOKEN_URL, {
    method: 'post',
    payload: {
      grant_type: 'authorization_code',
      client_id: props.getProperty('MF_CLIENT_ID'),
      client_secret: props.getProperty('MF_CLIENT_SECRET'),
      redirect_uri: MF_REDIRECT_URI,
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
  if (tokenData.refresh_token) props.setProperty('MF_REFRESH_TOKEN', tokenData.refresh_token);
  ui.alert('認証成功！MFデータの同期が可能になりました。');
}

function doGet(e) {
  var code = e.parameter.code;
  if (!code) {
    return HtmlService.createHtmlOutput('<p>認証はスプレッドシートのMF連携メニューから実行してください。</p>');
  }
  var props = PropertiesService.getScriptProperties();
  var response = UrlFetchApp.fetch(MF_TOKEN_URL, {
    method: 'post',
    payload: {
      grant_type: 'authorization_code',
      client_id: props.getProperty('MF_CLIENT_ID'),
      client_secret: props.getProperty('MF_CLIENT_SECRET'),
      redirect_uri: MF_REDIRECT_URI,
      code: code
    },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() === 200) {
    var tokenData = JSON.parse(response.getContentText());
    props.setProperty('MF_ACCESS_TOKEN', tokenData.access_token);
    if (tokenData.refresh_token) props.setProperty('MF_REFRESH_TOKEN', tokenData.refresh_token);
    return HtmlService.createHtmlOutput('<h2>MF認証成功！</h2><p>このタブを閉じてスプレッドシートに戻ってください。</p>');
  }
  return HtmlService.createHtmlOutput('<h2>認証失敗</h2><p>' + response.getContentText() + '</p>');
}

function refreshToken_() {
  var props = PropertiesService.getScriptProperties();
  var refreshToken = props.getProperty('MF_REFRESH_TOKEN');
  if (!refreshToken) {
    throw new Error('リフレッシュトークンがありません。MF連携メニューから再認証してください。');
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
  if (tokenData.refresh_token) props.setProperty('MF_REFRESH_TOKEN', tokenData.refresh_token);
  return tokenData.access_token;
}

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

function resetAuth() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('MF_ACCESS_TOKEN');
  props.deleteProperty('MF_REFRESH_TOKEN');
  SpreadsheetApp.getUi().alert('認証をリセットしました。');
}
