/**
 * キャッシュフロー管理システム - マネーフォワード OAuth2認証モジュール
 *
 * OAuth2ライブラリ（apps-script-oauth2）を使用してMFクラウド会計と連携する。
 *
 * ■ 前提
 *   - GASエディタ → ライブラリ → 以下のスクリプトIDを追加:
 *     1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF
 *     （OAuth2 for Apps Script）
 *
 * ■ MFアプリポータルでの設定
 *   - リダイレクトURI: https://script.google.com/macros/d/{SCRIPT_ID}/usercallback
 */

/**
 * スクリプトプロパティからMFクレデンシャルを取得する
 * @return {{ clientId: string, clientSecret: string }}
 */
function getMfCredentials_() {
  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('MF_CLIENT_ID');
  const clientSecret = props.getProperty('MF_CLIENT_SECRET');

  if (!clientId || !clientSecret) {
    throw new Error(
      'MF APIのクレデンシャルが未設定です。\n\n' +
      'GASエディタ → プロジェクトの設定 → スクリプトプロパティ に以下を追加:\n' +
      '  MF_CLIENT_ID: (Client ID)\n' +
      '  MF_CLIENT_SECRET: (Client Secret)'
    );
  }

  return { clientId, clientSecret };
}

/**
 * MF OAuth2サービスを構築する
 * @return {OAuth2.Service}
 */
function getMfOAuth2Service_() {
  const config = CF_CONFIG.MF_API;
  const creds = getMfCredentials_();

  return OAuth2.createService('moneyforward')
    .setAuthorizationBaseUrl(config.AUTH_URL)
    .setTokenUrl(config.TOKEN_URL)
    .setClientId(creds.clientId)
    .setClientSecret(creds.clientSecret)
    .setScope(config.SCOPE)
    .setCallbackFunction('mfAuthCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setCache(CacheService.getUserCache());
}

/**
 * OAuth2コールバック（リダイレクトURI先で呼ばれる）
 * @param {Object} request
 * @return {HtmlOutput}
 */
function mfAuthCallback(request) {
  const service = getMfOAuth2Service_();
  const authorized = service.handleCallback(request);

  if (authorized) {
    // 認証成功 → 事業所IDを自動取得して保存
    try {
      fetchAndSaveOfficeId_();
    } catch (e) {
      Logger.log('事業所ID取得エラー（無視）: ' + e.message);
    }

    try {
      fetchAndSaveWalletIds_();
    } catch (e) {
      Logger.log('口座マッピングエラー（無視）: ' + e.message);
    }

    return HtmlService.createHtmlOutput(
      '<h2>✅ マネーフォワード連携成功！</h2>' +
      '<p>スプレッドシートに戻って「💰 CF管理」メニューから操作してください。</p>' +
      '<p>このタブは閉じて構いません。</p>'
    );
  } else {
    return HtmlService.createHtmlOutput(
      '<h2>❌ 認証に失敗しました</h2>' +
      '<p>もう一度やり直してください。</p>'
    );
  }
}

/**
 * MF連携を開始する（認証URLをダイアログ表示）
 */
function startMfAuth() {
  const service = getMfOAuth2Service_();

  if (service.hasAccess()) {
    SpreadsheetApp.getUi().alert('✅ マネーフォワードは既に連携済みです。\n\n再連携する場合は「MF連携解除」を実行してください。');
    return;
  }

  const authUrl = service.getAuthorizationUrl();

  const html = HtmlService.createHtmlOutput(`
    <html>
    <head>
      <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; padding: 20px; text-align: center; }
        a { display: inline-block; margin-top: 16px; padding: 12px 32px; background: #1a73e8;
            color: white; text-decoration: none; border-radius: 6px; font-size: 14px; }
        a:hover { background: #1557b0; }
        .note { font-size: 12px; color: #666; margin-top: 16px; }
      </style>
    </head>
    <body>
      <h3>🔗 マネーフォワード連携</h3>
      <p>下のボタンをクリックして、MFクラウド会計へのアクセスを許可してください。</p>
      <a href="${authUrl}" target="_blank">マネーフォワードに接続</a>
      <p class="note">※ 認証後、このダイアログは閉じてください</p>
    </body>
    </html>
  `).setWidth(420).setHeight(260);

  SpreadsheetApp.getUi().showModalDialog(html, 'マネーフォワード連携');
}

/**
 * MF連携を解除する
 */
function disconnectMf() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'MF連携解除',
    'マネーフォワードとの連携を解除しますか？\n\n既に取り込んだデータは削除されません。',
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) return;

  const service = getMfOAuth2Service_();
  service.reset();

  // 保存済みの設定をクリア
  const props = PropertiesService.getUserProperties();
  props.deleteProperty('MF_OFFICE_ID');
  props.deleteProperty('MF_WALLET_MAP');

  ui.alert('✅ マネーフォワード連携を解除しました。');
}

/**
 * MF連携状態を確認
 * @return {boolean}
 */
function isMfConnected() {
  const service = getMfOAuth2Service_();
  return service.hasAccess();
}

/**
 * MF APIにリクエストを送る（認証ヘッダー付き）
 * @param {string} endpoint - APIエンドポイント（例: /offices）
 * @param {Object} [params] - クエリパラメータ
 * @return {Object} レスポンスJSON
 */
function mfApiRequest_(endpoint, params) {
  const service = getMfOAuth2Service_();
  if (!service.hasAccess()) {
    throw new Error('マネーフォワード未連携です。メニューから「MF連携開始」を実行してください。');
  }

  let url = CF_CONFIG.MF_API.BASE_URL + endpoint;

  if (params) {
    const queryString = Object.entries(params)
      .filter(([, v]) => v !== undefined && v !== null && v !== '')
      .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
      .join('&');
    if (queryString) url += '?' + queryString;
  }

  const response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + service.getAccessToken(),
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  if (code === 401) {
    // トークン期限切れ → リフレッシュ試行
    service.refresh();
    return mfApiRequest_(endpoint, params);
  }
  if (code === 429) {
    // レート制限 → 5秒待ってリトライ
    Utilities.sleep(5000);
    return mfApiRequest_(endpoint, params);
  }
  if (code >= 400) {
    throw new Error(`MF API エラー (${code}): ${response.getContentText()}`);
  }

  return JSON.parse(response.getContentText());
}

/**
 * 事業所IDを取得して保存する
 * MFクラウド会計APIはトークンに紐づく事業所データを直接返す
 */
function fetchAndSaveOfficeId_() {
  const data = mfApiRequest_('/offices');

  // APIはトークンに紐づく事業所を直接返す（配列ではない）
  // accounting_periodsがあれば事業所データが取得できている
  if (data.accounting_periods || data.id || data.display_name) {
    const officeId = String(data.id || 'default');
    PropertiesService.getUserProperties().setProperty('MF_OFFICE_ID', officeId);
    Logger.log('事業所データ取得成功');
  } else if (data.offices && data.offices.length > 0) {
    // リスト形式で返る場合のフォールバック
    const officeId = String(data.offices[0].id);
    PropertiesService.getUserProperties().setProperty('MF_OFFICE_ID', officeId);
    Logger.log('事業所ID保存: ' + officeId);
  } else {
    throw new Error('MFクラウド会計に事業所が見つかりません。レスポンス: ' + JSON.stringify(data).substring(0, 300));
  }
}

/**
 * 口座一覧を取得し、3口座のwallet IDを紐付けて保存する
 *
 * MFの勘定科目構造:
 *   accounts[] → { id, name, category, sub_accounts[] → { id, name, account_id } }
 *
 * 銀行口座はcategory=CASH_AND_DEPOSITSの勘定科目（普通預金等）の
 * sub_accounts（補助科目）として登録されている。
 * 例: 普通預金 → PayPay銀行 ビジネス営業部、PayPay銀行 はやぶさ支店
 */
function fetchAndSaveWalletIds_() {
  const data = mfApiRequest_('/accounts');
  const accounts = data.accounts || [];

  // CASH_AND_DEPOSITS カテゴリの勘定科目と補助科目を抽出
  const bankAccounts = [];
  accounts.forEach(acct => {
    if (acct.category !== 'CASH_AND_DEPOSITS') return;

    const subs = acct.sub_accounts || [];
    subs.forEach(sub => {
      bankAccounts.push({
        accountId: acct.id,           // 勘定科目ID（普通預金等）
        accountName: acct.name,       // 勘定科目名
        subAccountId: sub.id,         // 補助科目ID（各銀行口座）
        subAccountName: sub.name,     // 補助科目名（PayPay銀行 ビジネス営業部等）
        name: sub.name               // マッチング用
      });
    });

    // 補助科目がない場合は勘定科目自体を候補にする
    if (subs.length === 0) {
      bankAccounts.push({
        accountId: acct.id,
        accountName: acct.name,
        subAccountId: '',
        subAccountName: '',
        name: acct.name
      });
    }
  });

  // 口座名でマッチング（部分一致）
  const walletMap = {};
  const matchPatterns = {
    CF005: ['ビジネス営業', 'PayPay.*005', 'PayPay.*ビジネス'],
    CF003: ['はやぶさ', 'PayPay.*003'],
    SEIBU: ['西武信用金庫', '西武信金', '阿佐ヶ谷']
  };

  bankAccounts.forEach(b => {
    const searchName = b.name || '';

    for (const [accountKey, patterns] of Object.entries(matchPatterns)) {
      if (walletMap[accountKey]) continue;
      for (const pattern of patterns) {
        if (new RegExp(pattern, 'i').test(searchName)) {
          walletMap[accountKey] = {
            accountId: String(b.accountId),
            subAccountId: String(b.subAccountId || ''),
            subAccountName: b.subAccountName || '',
            name: searchName
          };
          break;
        }
      }
    }
  });

  PropertiesService.getUserProperties().setProperty('MF_WALLET_MAP', JSON.stringify(walletMap));

  const matched = Object.keys(walletMap).length;
  Logger.log(`口座マッピング完了: ${matched}/3 口座`);
  if (matched < 3) {
    Logger.log('⚠️ 一部の口座がマッチしませんでした。設定シートで手動紐付けが必要です。');
  }
}

/**
 * 保存済みの事業所IDを取得
 * @return {string}
 */
function getOfficeId_() {
  const officeId = PropertiesService.getUserProperties().getProperty('MF_OFFICE_ID');
  if (!officeId) throw new Error('事業所IDが未設定です。MF連携を実行してください。');
  return officeId;
}

/**
 * 保存済みのwallet IDマップを取得
 * @return {Object} { CF005: { id, type, name }, CF003: {...}, SEIBU: {...} }
 */
function getWalletMap_() {
  const json = PropertiesService.getUserProperties().getProperty('MF_WALLET_MAP');
  if (!json) throw new Error('口座マッピングが未設定です。MF連携を実行してください。');
  return JSON.parse(json);
}

/**
 * リダイレクトURIを表示する（MFアプリポータル設定用）
 */
function showRedirectUri() {
  const scriptId = ScriptApp.getScriptId();
  const redirectUri = `https://script.google.com/macros/d/${scriptId}/usercallback`;

  SpreadsheetApp.getUi().alert(
    'リダイレクトURI',
    'MFアプリポータルの「リダイレクトURI」に以下を設定してください：\n\n' + redirectUri,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
