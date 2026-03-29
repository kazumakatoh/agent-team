/**
 * 財務レポート自動化システム - 初期設定・OAuth認証・トリガー管理
 *
 * ■ セットアップ順序（必ずこの順番で）
 *   STEP 1. Google スプレッドシートを新規作成
 *   STEP 2. GASにコードをデプロイ → Web App URLを取得
 *   STEP 3. MFアプリポータルでアプリ登録（Web App URLをリダイレクトURIに設定）
 *   STEP 4. Config.gs に CLIENT_ID / CLIENT_SECRET / REDIRECT_URI / SPREADSHEET_ID を設定
 *   STEP 5. スプレッドシートを開いてメニューから「🔑 MF会計 認証を開始」を実行
 *   STEP 6. 「🏢 部門一覧を出力」でID確認 → Config.gs の DEPARTMENTS に設定
 *   STEP 7. 「⚙️ 日次トリガーを設定」で自動実行を有効化
 */

// ==============================
// OAuth 2.0 コールバックハンドラ
// ==============================

/**
 * GAS Web App として動作する OAuth2 コールバックエントリーポイント
 * MF会計の認証後、ここにリダイレクトされてトークンを自動取得する
 *
 * このURLがMFアプリポータルのリダイレクトURIになります。
 * 形式: https://script.google.com/macros/s/{DEPLOYMENT_ID}/exec
 */
function doGet(e) {
  const code  = e && e.parameter && e.parameter.code;
  const state = e && e.parameter && e.parameter.state;
  const error = e && e.parameter && e.parameter.error;

  // エラー応答
  if (error) {
    return HtmlService.createHtmlOutput(`
      <html><body style="font-family:sans-serif;padding:30px;">
        <h2 style="color:#c62828;">❌ 認証エラー</h2>
        <p>${error}</p>
        <p>このページを閉じてスプレッドシートに戻り、再度お試しください。</p>
      </body></html>
    `);
  }

  if (!code) {
    return HtmlService.createHtmlOutput(`
      <html><body style="font-family:sans-serif;padding:30px;">
        <h2 style="color:#c62828;">❌ 認証コードが受け取れませんでした</h2>
        <p>このページを閉じてスプレッドシートに戻り、再度「MF会計 認証を開始」を実行してください。</p>
      </body></html>
    `);
  }

  // state 検証（CSRF対策）
  const props     = PropertiesService.getScriptProperties();
  const savedState = props.getProperty('MF_OAUTH_STATE');
  if (savedState && state !== savedState) {
    return HtmlService.createHtmlOutput(`
      <html><body style="font-family:sans-serif;padding:30px;">
        <h2 style="color:#c62828;">❌ セキュリティエラー（state不一致）</h2>
        <p>このページを閉じてスプレッドシートに戻り、再度お試しください。</p>
      </body></html>
    `);
  }

  // 認証コード → トークン交換
  try {
    MFApiClient.exchangeCodeForTokens(code);
    props.deleteProperty('MF_OAUTH_STATE');

    return HtmlService.createHtmlOutput(`
      <html>
      <head><base target="_top">
      <style>
        body { font-family: 'Helvetica Neue', sans-serif; padding: 30px; color: #333; }
        h2   { color: #1b5e20; }
        .box { background: #e8f5e9; padding: 16px; border-radius: 8px; margin-top: 16px; }
        ol   { line-height: 2; }
      </style>
      </head>
      <body>
        <h2>✅ MF会計 認証が完了しました</h2>
        <div class="box">
          <strong>次のステップ（スプレッドシートに戻って実行）：</strong>
          <ol>
            <li>「📊 財務レポート」メニュー → 「🏢 部門一覧を出力」</li>
            <li>表示された部門IDを Config.gs の DEPARTMENTS に設定</li>
            <li>「📋 勘定科目一覧を出力」で科目名を確認</li>
            <li>「🔬 API生データを出力」でレスポンス構造を確認</li>
            <li>「⚙️ 日次トリガーを設定」で自動更新を有効化</li>
            <li>「🔄 当期PLを更新」で動作確認</li>
          </ol>
        </div>
        <p style="margin-top:20px;color:#666;">このページを閉じてスプレッドシートに戻ってください。</p>
      </body>
      </html>
    `);
  } catch (err) {
    return HtmlService.createHtmlOutput(`
      <html><body style="font-family:sans-serif;padding:30px;">
        <h2 style="color:#c62828;">❌ トークン取得エラー</h2>
        <p>${err.message}</p>
        <p>Config.gs の CLIENT_ID / CLIENT_SECRET が正しいか確認してください。</p>
      </body></html>
    `);
  }
}

// ==============================
// OAuth 2.0 認証開始
// ==============================

/**
 * 認証URLを生成してダイアログに表示する
 * ユーザーがURLをクリック → MFで認証 → doGet() に自動コールバック
 */
function authorize() {
  const ui = SpreadsheetApp.getUi();

  // 設定チェック
  const missing = [];
  if (!CONFIG.MF_API.CLIENT_ID)    missing.push('MF_API.CLIENT_ID');
  if (!CONFIG.MF_API.CLIENT_SECRET) missing.push('MF_API.CLIENT_SECRET');
  if (!CONFIG.MF_API.REDIRECT_URI)  missing.push('MF_API.REDIRECT_URI（GAS Web App URL）');
  if (!CONFIG.SPREADSHEET_ID)       missing.push('SPREADSHEET_ID');

  if (missing.length > 0) {
    ui.alert(
      '⚠️ Config.gs の設定が不足しています\n\n' +
      '未設定の項目:\n' + missing.map(m => '  ・' + m).join('\n') + '\n\n' +
      '導入手順.md を参照して設定してください。'
    );
    return;
  }

  const state = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty('MF_OAUTH_STATE', state);

  const authUrl = CONFIG.MF_API.AUTH_URL +
    `?client_id=${encodeURIComponent(CONFIG.MF_API.CLIENT_ID)}` +
    `&redirect_uri=${encodeURIComponent(CONFIG.MF_API.REDIRECT_URI)}` +
    `&response_type=code` +
    `&scope=${encodeURIComponent(CONFIG.MF_API.SCOPE)}` +
    `&state=${state}`;

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: 'Helvetica Neue', sans-serif; padding: 20px; font-size: 14px; color: #333; }
        h3   { color: #1a237e; margin-top: 0; }
        .btn { display: inline-block; margin-top: 12px; padding: 12px 24px; background: #1a237e;
               color: white; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 15px; }
        .url { background: #f5f5f5; padding: 10px; border-radius: 4px; word-break: break-all;
               font-size: 11px; margin-top: 12px; }
        ol   { padding-left: 20px; line-height: 2; }
        .note { background: #fff3e0; padding: 10px; border-radius: 4px; font-size: 12px; margin-top: 12px; }
      </style>
    </head>
    <body>
      <h3>🔑 MF会計 認証</h3>
      <ol>
        <li>「MF会計で認証する」ボタンをクリック</li>
        <li>MF会計にログインして「許可する」をクリック</li>
        <li>認証完了後、自動でトークンが保存されます</li>
      </ol>
      <a class="btn" href="${authUrl}" target="_blank">MF会計で認証する</a>
      <div class="note">
        ⚡ 認証後は自動でトークンが保存されます。<br>
        このダイアログを閉じて、スプレッドシートに戻ってください。
      </div>
      <div class="url">認証URL: ${authUrl}</div>
    </body>
    </html>
  `).setWidth(520).setHeight(360);

  ui.showModalDialog(html, '🔑 MF会計 認証');
}

/**
 * 認証情報をリセットする（再認証が必要な場合）
 */
function clearAuth() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('確認', '認証情報をすべてクリアします。再認証が必要になります。よろしいですか？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  MFApiClient.clearTokens();
  ui.alert('✅ 認証情報をクリアしました。\n「MF会計 認証を開始」から再認証してください。');
}

/**
 * 現在のGAS Web App URLをダイアログに表示する（MFポータル登録用）
 * STEP 2でこのURLをMFアプリポータルのリダイレクトURIに登録します
 */
function showWebAppUrl() {
  const ui = SpreadsheetApp.getUi();
  try {
    const url = ScriptApp.getService().getUrl();
    if (!url) {
      ui.alert(
        '⚠️ Web App URLが取得できません\n\n' +
        'GASをWeb Appとしてデプロイしていない可能性があります。\n\n' +
        '導入手順.md の「STEP 2: GASをデプロイする」を実行してください。'
      );
      return;
    }
    const html = HtmlService.createHtmlOutput(`
      <html><head><base target="_top">
      <style>body{font-family:sans-serif;padding:20px;}
        .url{background:#e3f2fd;padding:14px;border-radius:6px;word-break:break-all;font-size:13px;font-weight:bold;}
        p{font-size:13px;color:#555;}
      </style></head>
      <body>
        <h3 style="color:#1a237e;margin-top:0;">Web App URL（リダイレクトURI）</h3>
        <div class="url">${url}</div>
        <p style="margin-top:14px;">
          ① このURLをコピーして MFアプリポータルの「リダイレクトURI」に登録してください。<br>
          ② Config.gs の <code>REDIRECT_URI</code> にも同じURLを貼り付けてください。
        </p>
      </body></html>
    `).setWidth(560).setHeight(240);
    ui.showModalDialog(html, '📋 Web App URL（リダイレクトURI）');
  } catch (e) {
    ui.alert(`❌ エラー: ${e.message}`);
  }
}

// ==============================
// トリガー管理
// ==============================

/**
 * 毎日 AM6:00 に dailyPLUpdate() を自動実行するトリガーを設定する
 */
function setupTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['dailyPLUpdate', 'monthlyPLUpdate'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('dailyPLUpdate')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .inTimezone('Asia/Tokyo')
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ トリガー設定完了\n\n' +
    '毎日 AM6:00 に当期PLレポートを自動更新します。\n' +
    '（当月を含む過去月のみ更新。未来月の予測値は保持されます）'
  );
}

/**
 * 自動更新トリガーを削除する
 */
function removeTriggers() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('確認', '日次自動更新トリガーを削除します。よろしいですか？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  let count = 0;
  ScriptApp.getProjectTriggers()
    .filter(t => ['dailyPLUpdate', 'monthlyPLUpdate'].includes(t.getHandlerFunction()))
    .forEach(t => { ScriptApp.deleteTrigger(t); count++; });

  ui.alert(`✅ ${count}件のトリガーを削除しました。`);
}
