/**
 * Amazon Dashboard - 認証モジュール
 *
 * LWA (Login with Amazon) を使った Access Token 取得
 * SP-API と Ads API の両方で使用
 */

/**
 * SP-API 用の Access Token を取得
 * @returns {string} Access Token
 */
function getSpApiAccessToken() {
  return getLwaAccessToken(
    getCredential('SP_CLIENT_ID'),
    getCredential('SP_CLIENT_SECRET'),
    getCredential('SP_REFRESH_TOKEN')
  );
}

/**
 * Ads API 用の Access Token を取得
 * @returns {string} Access Token
 */
function getAdsApiAccessToken() {
  return getLwaAccessToken(
    getCredential('ADS_CLIENT_ID'),
    getCredential('ADS_CLIENT_SECRET'),
    getCredential('ADS_REFRESH_TOKEN')
  );
}

/**
 * LWA (Login with Amazon) で Access Token を取得
 * Refresh Token を使って新しい Access Token を取得する
 *
 * @param {string} clientId - LWA Client ID
 * @param {string} clientSecret - LWA Client Secret
 * @param {string} refreshToken - Refresh Token
 * @returns {string} Access Token
 */
function getLwaAccessToken(clientId, clientSecret, refreshToken) {
  // キャッシュを確認（有効期限内ならキャッシュから返す）
  const cache = CacheService.getScriptCache();
  const cacheKey = 'access_token_' + clientId.slice(-8);
  const cached = cache.get(cacheKey);
  if (cached) {
    return cached;
  }

  const payload = {
    grant_type: 'refresh_token',
    refresh_token: refreshToken,
    client_id: clientId,
    client_secret: clientSecret,
  };

  const options = {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: payload,
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(LWA_TOKEN_URL, options);
  const statusCode = response.getResponseCode();
  const body = JSON.parse(response.getContentText());

  if (statusCode !== 200) {
    throw new Error('LWA認証エラー: ' + JSON.stringify(body));
  }

  // Access Token をキャッシュ（有効期限の5分前まで）
  const expiresIn = body.expires_in || 3600;
  cache.put(cacheKey, body.access_token, expiresIn - 300);

  return body.access_token;
}

/**
 * SP-API 認証テスト
 * GASエディタから手動実行して、認証が正しく動作するか確認する
 */
function testSpApiAuth() {
  try {
    const token = getSpApiAccessToken();
    Logger.log('✅ SP-API Access Token 取得成功（先頭20文字）: ' + token.substring(0, 20) + '...');

    // マーケットプレイス参加情報を取得してテスト
    const url = SP_API_ENDPOINT + '/sellers/v1/marketplaceParticipations';
    const options = {
      method: 'get',
      headers: {
        'x-amz-access-token': token,
        'Content-Type': 'application/json',
      },
      muteHttpExceptions: true,
    };
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();

    if (statusCode === 200) {
      const data = JSON.parse(response.getContentText());
      Logger.log('✅ SP-API 接続成功: ' + data.payload.length + ' 件のマーケットプレイス参加情報');
      data.payload.forEach(p => {
        Logger.log('  - ' + p.marketplace.name + ' (' + p.marketplace.id + ')');
      });
    } else {
      Logger.log('❌ SP-API 接続エラー: HTTP ' + statusCode);
      Logger.log(response.getContentText());
    }
  } catch (e) {
    Logger.log('❌ エラー: ' + e.message);
  }
}

/**
 * Ads API 認証テスト
 * Profiles取得を試みる（サポート回答後に実行）
 */
function testAdsApiAuth() {
  try {
    const token = getAdsApiAccessToken();
    Logger.log('✅ Ads API Access Token 取得成功（先頭20文字）: ' + token.substring(0, 20) + '...');

    const url = ADS_API_ENDPOINT + '/v2/profiles';
    const options = {
      method: 'get',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Amazon-Advertising-API-ClientId': getCredential('ADS_CLIENT_ID'),
        'Content-Type': 'application/json',
      },
      muteHttpExceptions: true,
    };
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    const data = JSON.parse(response.getContentText());

    if (statusCode === 200 && data.length > 0) {
      Logger.log('✅ Ads API Profiles 取得成功: ' + data.length + ' 件');
      data.forEach(p => {
        Logger.log('  - Profile ID: ' + p.profileId + ' (' + p.countryCode + ')');
      });
    } else if (statusCode === 200 && data.length === 0) {
      Logger.log('⚠️ Ads API Profiles: 0件（サポート回答待ち）');
    } else {
      Logger.log('❌ Ads API エラー: HTTP ' + statusCode);
      Logger.log(response.getContentText());
    }
  } catch (e) {
    Logger.log('❌ エラー: ' + e.message);
  }
}
