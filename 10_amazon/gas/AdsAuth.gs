/**
 * Amazon Dashboard - Ads API 認証管理（Phase 3）
 *
 * ## 使い方（初回セットアップ）
 *
 * ### Step 1: Script Properties に認証情報を設定
 * GASエディタ → 歯車アイコン（プロジェクトの設定）→ スクリプトプロパティ
 * 以下を「プロパティを追加」：
 *   - ADS_CLIENT_ID       : amzn1.application-oa2-client.xxxxxxxx
 *   - ADS_CLIENT_SECRET   : （LWA Security Profile の「シークレットを表示」から）
 *
 * ### Step 2: 認可コードを取得
 * シークレットウィンドウで以下のOAuth URLを開き、
 * `kazuma.katoh@level1.biz` でログインして許可 → URL から code を取得:
 *
 *   https://apac.account.amazon.com/ap/oa?client_id=【CLIENT_ID】&scope=advertising%3A%3Acampaign_management&response_type=code&redirect_uri=https%3A%2F%2Fwww.example.com%2Fcallback
 *
 * ### Step 3: 下記 ADS_AUTH_CODE に貼り付けて exchangeAdsAuthCode() を実行
 * 認可コードは5分で失効するので素早く。
 *
 * ### Step 4: testAdsProfiles() を実行して /v2/profiles が動くか確認
 */

// ===== Ads API エンドポイント =====
const ADS_TOKEN_URL = 'https://api.amazon.com/auth/o2/token';
const ADS_REDIRECT_URI = 'https://www.example.com/callback';

/**
 * 認可コードを refresh_token に交換し Script Properties に保存
 *
 * 手順:
 *   1. 下記 ADS_AUTH_CODE に新しい認可コードを貼り付ける
 *   2. この関数を▶実行
 *   3. Logger に refresh_token 長さが表示されれば成功
 *   4. 実行後、ADS_AUTH_CODE を '' に戻してセーブ（セキュリティ）
 */
function exchangeAdsAuthCode() {
  // ═══════════════════════════════════════════════════════════════
  // ここに認可コードを貼り付けてください（５分で失効するので急ぐ）
  const ADS_AUTH_CODE = '';
  // ═══════════════════════════════════════════════════════════════

  if (!ADS_AUTH_CODE) {
    throw new Error('ADS_AUTH_CODE 変数に認可コードを貼り付けてから実行してください');
  }

  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('ADS_CLIENT_ID');
  const clientSecret = props.getProperty('ADS_CLIENT_SECRET');

  if (!clientId) throw new Error('ADS_CLIENT_ID が Script Properties に設定されていません');
  if (!clientSecret) throw new Error('ADS_CLIENT_SECRET が Script Properties に設定されていません');

  Logger.log('認可コード交換中...');

  const response = UrlFetchApp.fetch(ADS_TOKEN_URL, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      grant_type: 'authorization_code',
      code: ADS_AUTH_CODE,
      client_id: clientId,
      client_secret: clientSecret,
      redirect_uri: ADS_REDIRECT_URI,
    },
    muteHttpExceptions: true,
  });

  const status = response.getResponseCode();
  const body = response.getContentText();

  Logger.log('HTTP Status: ' + status);

  if (status !== 200) {
    Logger.log('❌ エラー レスポンス:');
    Logger.log(body);
    Logger.log('');
    Logger.log('よくある原因:');
    Logger.log('  - 認可コードが既に使用済み（1回しか使えない）');
    Logger.log('  - 認可コードが失効（取得から5分超）');
    Logger.log('  - ADS_CLIENT_SECRET が間違っている');
    Logger.log('  → 新しい認可コードを取得して再実行してください');
    return;
  }

  const data = JSON.parse(body);

  if (!data.refresh_token) {
    Logger.log('❌ refresh_token がレスポンスに含まれていません');
    Logger.log('Body: ' + body);
    return;
  }

  // Script Properties に保存
  props.setProperty('ADS_REFRESH_TOKEN', data.refresh_token);

  Logger.log('✅ ADS_REFRESH_TOKEN を保存しました');
  Logger.log('  refresh_token 長さ: ' + data.refresh_token.length + ' 文字');
  Logger.log('  access_token 長さ: ' + data.access_token.length + ' 文字');
  Logger.log('  expires_in: ' + data.expires_in + ' 秒');
  Logger.log('');
  Logger.log('次は testAdsProfiles() を実行して /v2/profiles テストしてください');
}

/**
 * access_token を取得（refresh_token から）
 * @returns {string} access_token
 */
function getAdsAccessToken() {
  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('ADS_CLIENT_ID');
  const clientSecret = props.getProperty('ADS_CLIENT_SECRET');
  const refreshToken = props.getProperty('ADS_REFRESH_TOKEN');

  if (!clientId || !clientSecret || !refreshToken) {
    throw new Error('ADS_CLIENT_ID / ADS_CLIENT_SECRET / ADS_REFRESH_TOKEN が未設定です');
  }

  const response = UrlFetchApp.fetch(ADS_TOKEN_URL, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      grant_type: 'refresh_token',
      refresh_token: refreshToken,
      client_id: clientId,
      client_secret: clientSecret,
    },
    muteHttpExceptions: true,
  });

  const status = response.getResponseCode();
  if (status !== 200) {
    throw new Error('access_token 取得失敗 (' + status + '): ' + response.getContentText());
  }

  return JSON.parse(response.getContentText()).access_token;
}

/**
 * /v2/profiles をテスト実行
 * プロファイルが返ってきたら自動で ADS_PROFILE_ID を保存
 */
function testAdsProfiles() {
  Logger.log('===== Ads API /v2/profiles テスト =====');

  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('ADS_CLIENT_ID');

  // access_token 取得
  let accessToken;
  try {
    accessToken = getAdsAccessToken();
    Logger.log('✅ access_token 取得成功');
  } catch (e) {
    Logger.log('❌ ' + e.message);
    return;
  }

  // /v2/profiles を叩く（極東エンドポイント）
  const response = UrlFetchApp.fetch(ADS_API_ENDPOINT + '/v2/profiles', {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + accessToken,
      'Amazon-Advertising-API-ClientId': clientId,
    },
    muteHttpExceptions: true,
  });

  const status = response.getResponseCode();
  const body = response.getContentText();

  Logger.log('--- レスポンス ---');
  Logger.log('Status: ' + status);
  Logger.log('Body: ' + body);

  if (status !== 200) {
    Logger.log('❌ API呼び出し失敗');
    return;
  }

  const profiles = JSON.parse(body);
  if (!Array.isArray(profiles) || profiles.length === 0) {
    Logger.log('⚠️ プロファイルは空配列です（顧客ID問題がまだ解決していない可能性）');
    return;
  }

  Logger.log('🎉 プロファイル取得成功！ ' + profiles.length + ' 件');
  profiles.forEach((p, i) => {
    Logger.log('');
    Logger.log('Profile ' + (i + 1) + ':');
    Logger.log('  profileId: ' + p.profileId);
    Logger.log('  countryCode: ' + p.countryCode);
    Logger.log('  currencyCode: ' + p.currencyCode);
    Logger.log('  timezone: ' + p.timezone);
    if (p.accountInfo) {
      Logger.log('  accountInfo.marketplaceStringId: ' + p.accountInfo.marketplaceStringId);
      Logger.log('  accountInfo.type: ' + p.accountInfo.type);
      Logger.log('  accountInfo.name: ' + p.accountInfo.name);
    }
  });

  // JP プロファイルがあれば自動保存
  const jpProfile = profiles.find(p => p.countryCode === 'JP');
  if (jpProfile) {
    props.setProperty('ADS_PROFILE_ID', String(jpProfile.profileId));
    Logger.log('');
    Logger.log('✅ JP Profile ID を Script Properties に保存: ' + jpProfile.profileId);
    Logger.log('これで Phase 3 の実装に進めます');
  } else {
    Logger.log('');
    Logger.log('⚠️ JP プロファイルが見つかりませんでした');
    Logger.log('必要なプロファイルの profileId を手動で ADS_PROFILE_ID に保存してください');
  }
}

/**
 * 現在の Ads API 認証状態を確認
 */
function checkAdsCredentials() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const keys = ['ADS_CLIENT_ID', 'ADS_CLIENT_SECRET', 'ADS_REFRESH_TOKEN', 'ADS_PROFILE_ID'];

  Logger.log('===== Ads API 認証情報 =====');
  keys.forEach(key => {
    const value = props[key];
    const status = value ? '✅ 設定済み (' + value.length + '文字)' : '❌ 未設定';
    Logger.log(key + ': ' + status);
  });
}
