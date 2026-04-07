/**
 * MoneyForward クラウド会計 認証管理
 *
 * セットアップ:
 *   1. MF会計 > アプリポータル > APIキー管理 でアプリのClient ID / Secretを確認
 *   2. MFアプリのリダイレクトURIに、MF連携 > 「リダイレクトURIを確認」で表示されるURLを登録
 *   3. スプシの MF連携 > 「① MF認証情報を設定」で Client ID / Secret を入力
 *   4. MF連携 > 「② MF認証を実行」→ 表示されるURLを開いて認証
 *   5. リダイレクトされたURLのcodeパラメータをコピーし、③で入力
 */

var MF_API_BASE  = 'https://accounting.moneyforward.com/api/v3';
var MF_AUTH_URL  = 'https://accounting.moneyforward.com/oauth/authorize';
var MF_TOKEN_URL = 'https://accounting.moneyforward.com/oauth/token';

/**
 * リダイレクトURIを取得（MFアプリに登録する）
 */
function getRedirectUri() {
  var url = ScriptApp.getService().getUrl();
  if (!url) {
    SpreadsheetApp.getUi().alert(
      'ウェブアプリのデプロイが必要です。\n\n'
      + '手順：\n'
      + '1. Apps Script エディタ > デプロイ > 新しいデプロイ\n'
      + '2. 種類：ウェブアプリ\n'
      + '3. アクセスできるユーザー：自分のみ\n'
      + '4. デプロイ後、再度この関数を実行してください'
    );
    return;
  }
  SpreadsheetApp.getUi().alert(
    'このURLをMFアプリのリダイレクトURIに登録してください：\n\n' + url
  );
  Logger.log('リダイレクトURI: ' + url);
}

/**
 * ウェブアプリのコールバック（MF認証後のリダイレクト先）
 */
function doGet(e) {
  var code = e.parameter.code;
  if (code) {
    // 認可コードを受け取った → トークン交換
    var props = PropertiesService.getScriptProperties();
    var redirectUri = ScriptApp.getService().getUrl();

    var response = UrlFetchApp.fetch(MF_TOKEN_URL, {
      method: 'post',
      payload: {
        grant_type: 'authorization_code',
        client_id: props.getProperty('MF_CLIENT_ID'),
        client_secret: props.getProperty('MF_CLIENT_SECRET'),
        redirect_uri: redirectUri,
        code: code
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      var tokenData = JSON.parse(response.getContentText());
      props.setProperty('MF_ACCESS_TOKEN', tokenData.access_token);
      if (tokenData.refresh_token) {
        props.setProperty('MF_REFRESH_TOKEN', tokenData.refresh_token);
      }
      // 事業所IDも取得
      try { fetchAndSaveOfficeId_(); } catch(err) {}

      return HtmlService.createHtmlOutput(
        '<h2>✅ MF認証成功</h2>'
        + '<p>このタブを閉じて、スプレッドシートに戻ってください。</p>'
        + '<p>MF連携 > 「資金繰り表_2025 を同期」でデータを取得できます。</p>'
      );
    } else {
      return HtmlService.createHtmlOutput(
        '<h2>❌ 認証失敗</h2>'
        + '<p>エラー: ' + response.getContentText() + '</p>'
        + '<p>MF連携 > 「認証をリセット」してやり直してください。</p>'
      );
    }
  }

  // codeパラメータなし → 手動入力用画面
  return HtmlService.createHtmlOutput(
    '<h2>資金繰り表 MF連携</h2>'
    + '<p>認証はスプレッドシートの「MF連携」メニューから実行してください。</p>'
  );
}

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
  ui.alert('認証情報を保存しました。\n\n次のステップ：\n1. 「リダイレクトURIを確認」でURLを取得\n2. MFアプリにそのURLを登録\n3. 「MF認証を実行」で認証');
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

  var redirectUri = ScriptApp.getService().getUrl();
  if (!redirectUri) {
    SpreadsheetApp.getUi().alert(
      'ウェブアプリのデプロイが必要です。\n\n'
      + 'Apps Script > デプロイ > 新しいデプロイ > ウェブアプリ で作成してください。'
    );
    return;
  }

  var authUrl = MF_AUTH_URL
    + '?client_id=' + encodeURIComponent(clientId)
    + '&redirect_uri=' + encodeURIComponent(redirectUri)
    + '&response_type=code'
    + '&scope=' + encodeURIComponent('office:read account:read journal:read report:read');

  var html = HtmlService.createHtmlOutput(
    '<p>以下のリンクをクリックしてMF認証してください：</p>'
    + '<p><a href="' + authUrl + '" target="_blank" style="font-size:14px;">MF認証を開始する</a></p>'
    + '<p>認証後、自動的にトークンが保存されます。</p>'
  ).setWidth(500).setHeight(150);

  SpreadsheetApp.getUi().showModalDialog(html, 'MF認証');
}

/**
 * アクセストークンをリフレッシュ
 */
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
