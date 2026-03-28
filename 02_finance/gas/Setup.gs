/**
 * 財務レポート自動化システム - 初期設定・認証・トリガー管理
 *
 * ■ 初回セットアップ手順
 *   1. Config.gs に CLIENT_ID / CLIENT_SECRET / SPREADSHEET_ID を設定
 *   2. authorize() を実行 → 表示されたURLにアクセス
 *   3. MF会計で認証 → 表示された認証コードをコピー
 *   4. exchangeAuthCode() を実行 → コードを貼り付け
 *   5. exportDepartments() で部門IDを確認
 *   6. Config.gs の DEPARTMENTS に部門IDを設定
 *   7. setupTriggers() でトリガーを設定
 */

// ==============================
// OAuth 2.0 認証フロー
// ==============================

/**
 * STEP 1: 認証URLを生成して表示する
 * 表示されたURLにブラウザでアクセスし、MF会計でログイン・許可してください
 */
function authorize() {
  if (!CONFIG.MF_API.CLIENT_ID) {
    SpreadsheetApp.getUi().alert(
      '⚠️ 設定が必要です\n\n' +
      'Config.gs の以下を設定してください:\n' +
      '  - MF_API.CLIENT_ID\n' +
      '  - MF_API.CLIENT_SECRET\n' +
      '  - SPREADSHEET_ID\n\n' +
      'MF会計の API連携設定ページでアプリを登録し、\n' +
      'クライアントID / シークレットを取得してください。'
    );
    return;
  }

  const state   = Utilities.getUuid();
  const authUrl = CONFIG.MF_API.AUTH_URL +
    `?client_id=${encodeURIComponent(CONFIG.MF_API.CLIENT_ID)}` +
    `&redirect_uri=${encodeURIComponent(CONFIG.MF_API.REDIRECT_URI)}` +
    `&response_type=code` +
    `&scope=${encodeURIComponent(CONFIG.MF_API.SCOPE)}` +
    `&state=${state}`;

  // stateをPropertiesServiceに一時保存
  PropertiesService.getScriptProperties().setProperty('MF_OAUTH_STATE', state);

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: 'Helvetica Neue', sans-serif; padding: 20px; font-size: 14px; color: #333; }
        h3   { color: #1a237e; margin-top: 0; }
        .url { background: #f5f5f5; padding: 12px; border-radius: 4px; word-break: break-all; font-size: 12px; }
        .btn { display: inline-block; margin-top: 12px; padding: 10px 20px; background: #1a237e; color: white; text-decoration: none; border-radius: 4px; font-weight: bold; }
        ol   { padding-left: 20px; line-height: 1.8; }
      </style>
    </head>
    <body>
      <h3>🔑 MF会計 認証手順</h3>
      <ol>
        <li>下の「MF会計で認証する」ボタンをクリック</li>
        <li>MF会計にログインして「許可する」をクリック</li>
        <li>表示された <strong>認証コード</strong> をコピー</li>
        <li>このダイアログを閉じて、メニューから<br>「🔑 認証コードを入力」を実行</li>
      </ol>
      <a class="btn" href="${authUrl}" target="_blank">MF会計で認証する</a>
      <p style="font-size:12px;color:#666;margin-top:16px;">
        URLが開かない場合は以下をコピーしてブラウザに貼り付けてください:<br>
        <span class="url">${authUrl}</span>
      </p>
    </body>
    </html>
  `).setWidth(500).setHeight(380);

  SpreadsheetApp.getUi().showModalDialog(html, '🔑 MF会計 認証');
}

/**
 * STEP 2: 認証コードをアクセストークンと交換する
 * authorize() 実行後にMF会計で許可すると表示されるコードを入力してください
 */
function exchangeAuthCode() {
  const ui     = SpreadsheetApp.getUi();
  const result = ui.prompt(
    '認証コード入力',
    'MF会計の認証画面に表示された認証コードを入力してください：',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) return;
  const code = result.getResponseText().trim();

  if (!code) {
    ui.alert('⚠️ 認証コードが入力されていません。');
    return;
  }

  try {
    MFApiClient.exchangeCodeForTokens(code);

    // 事業所IDも自動取得して保存
    const companies = MFApiClient.getCompanies();
    if (companies.length > 0) {
      PropertiesService.getScriptProperties().setProperty('MF_COMPANY_ID', companies[0].id);
      Logger.log(`事業所を設定: ${companies[0].name} (${companies[0].id})`);
    }

    ui.alert(
      '✅ 認証が完了しました\n\n' +
      '次のステップ:\n' +
      '1. 「勘定科目一覧を出力」で勘定科目名を確認\n' +
      '2. 「部門一覧を出力」でID確認 → Config.gs に設定\n' +
      '3. 「月次トリガーを設定」で自動実行を有効化\n' +
      '4. 「当期PLを更新」で動作確認'
    );
  } catch (e) {
    ui.alert(`❌ 認証失敗\n${e.message}\n\n認証コードが正しいか確認してください。`);
  }
}

/**
 * 認証情報をリセットする（再認証が必要な場合）
 */
function clearAuth() {
  const ui      = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    '確認',
    '認証情報（トークン）をすべてクリアします。\n再度 authorize() から認証が必要になります。\nよろしいですか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  MFApiClient.clearTokens();
  ui.alert('✅ 認証情報をクリアしました。\n「MF会計 認証を開始」から再認証してください。');
}

// ==============================
// トリガー管理
// ==============================

/**
 * 日次自動更新トリガーを設定する
 * 毎日 AM6:00 に dailyPLUpdate() を実行
 * → 当月を含む過去月のデータのみ更新（未来月の手入力予測値は保持）
 */
function setupTriggers() {
  // 既存の日次トリガーを削除
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'dailyPLUpdate' || t.getHandlerFunction() === 'monthlyPLUpdate')
    .forEach(t => ScriptApp.deleteTrigger(t));

  // 毎日 AM6:00 実行
  ScriptApp.newTrigger('dailyPLUpdate')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .inTimezone('Asia/Tokyo')
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ トリガー設定完了\n\n' +
    '毎日 AM6:00 に当期PLレポートを自動更新します。\n' +
    '（当月を含む過去月のみ更新。未来月の予測値は保持されます）\n\n' +
    '手動で今すぐ実行したい場合は\n「🔄 当期PLを更新（過去・当月のみ）」を選択してください。'
  );
  Logger.log('日次トリガー設定完了');
}

/**
 * すべてのトリガーを削除する
 */
function removeTriggers() {
  const ui      = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    '確認',
    '日次自動更新トリガーを削除します。\n自動更新が停止されますがよろしいですか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  let count = 0;
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'dailyPLUpdate' || t.getHandlerFunction() === 'monthlyPLUpdate')
    .forEach(t => { ScriptApp.deleteTrigger(t); count++; });

  ui.alert(`✅ ${count}件のトリガーを削除しました。\n再設定するには「月次トリガーを設定」を実行してください。`);
}

// ==============================
// 接続テスト
// ==============================

/**
 * API接続テスト（設定確認用）
 */
function testApiConnection() {
  const ui = SpreadsheetApp.getUi();
  try {
    const companies = MFApiClient.getCompanies();
    if (!companies.length) {
      ui.alert('⚠️ 事業所が見つかりません。');
      return;
    }
    const c = companies[0];
    ui.alert(
      `✅ API接続成功\n\n` +
      `事業所名: ${c.name}\n` +
      `事業所ID: ${c.id}\n\n` +
      `接続OK！次のステップに進んでください。`
    );
  } catch (e) {
    ui.alert(`❌ API接続失敗\n\n${e.message}\n\n認証が完了しているか確認してください。`);
  }
}
