/**
 * 財務レポート自動化システム - マネーフォワード クラウド会計 APIクライアント
 *
 * ■ エンドポイント仕様（developers.api-accounting.moneyforward.com 準拠）
 *   BASE_URL: https://api-accounting.moneyforward.com
 *   PL試算表:  GET /api/v3/reports/trial_balance_pl
 *   BS試算表:  GET /api/v3/reports/trial_balance_bs
 *   部門一覧:  GET /api/v3/segments  （★ドキュメントで要確認）
 *   勘定科目:  GET /api/v3/account_items
 *
 * ■ OAuth2 仕様（developers.biz.moneyforward.com 準拠）
 *   Authorization Code Flow / Bearer Token
 *   認証基盤: MFビジネスプラットフォーム (accounts.biz.moneyforward.com)
 *
 * ★ 実際のエンドポイント名・パラメータ名はドキュメントで最終確認してください
 */

const MFApiClient = {

  // ==============================
  // 公開メソッド
  // ==============================

  /**
   * PL試算表を取得する
   *
   * @param {string} startDate    - 'YYYY-MM-DD'（期間開始日）
   * @param {string} endDate      - 'YYYY-MM-DD'（期間終了日）
   * @param {string} [segmentId]  - 部門/セグメントID（省略時は全社）
   * @return {Array} 試算表アイテム配列
   *
   * ★ パラメータ名はドキュメントで確認：
   *   start_date / end_date の名称、部門フィルタのキー名（segment_id? department_id?）
   */
  getTrialBalance(startDate, endDate, segmentId) {
    const params = {
      start_date: startDate,
      end_date:   endDate,
    };
    // 部門フィルタ: ★ドキュメントでパラメータ名を確認（segment_id / department_id など）
    if (segmentId) params[CONFIG.MF_API.SEGMENT_PARAM] = segmentId;

    const response = MFApiClient._request('GET', '/api/v3/reports/trial_balance_pl', params);

    // MF会計 API v3 試算表レスポンス構造:
    // { rows: [...], columns: [...] } または { items: [...] } など
    // 生データ確認は exportRawApiResponse() を参照
    const items = response.rows
      || response.items
      || response.account_items
      || (response.trial_balance_pl && response.trial_balance_pl.rows)
      || (response.trial_balance_pl && response.trial_balance_pl.items)
      || (response.trial_balance && response.trial_balance.items)
      || [];

    Logger.log(`試算表レスポンスキー: ${Object.keys(response).join(', ')} → ${items.length}件`);
    return items;
  },

  /**
   * BS試算表を取得する（将来のBS対応用に準備）
   */
  getBalanceSheet(startDate, endDate, segmentId) {
    const params = { start_date: startDate, end_date: endDate };
    if (segmentId) params[CONFIG.MF_API.SEGMENT_PARAM] = segmentId;
    const response = MFApiClient._request('GET', '/api/v3/reports/trial_balance_bs', params);
    return response.items || response.account_items || [];
  },

  /**
   * 部門一覧を取得する
   * スコープ: mfc/accounting/departments.read
   * エンドポイント候補: /api/v3/departments（ドキュメントで確認済みのスコープ名から推定）
   */
  getDepartments() {
    // 確認済みスコープ名 "departments" から推定する候補を順に試す
    const candidates = [
      '/api/v3/departments',
      '/api/v3/segments',
    ];

    for (const endpoint of candidates) {
      try {
        const response = MFApiClient._request('GET', endpoint, {});
        const items = response.departments || response.segments || response.items || [];
        if (items.length > 0) {
          Logger.log(`部門一覧取得成功: ${endpoint} (${items.length}件)`);
          return items;
        }
      } catch (e) {
        Logger.log(`部門エンドポイント ${endpoint} → ${e.message}`);
      }
    }
    return [];
  },

  /**
   * 勘定科目一覧を取得する（マッピング確認用）
   * ★ エンドポイント名をドキュメントで確認
   */
  getAccountItems() {
    const response = MFApiClient._request('GET', '/api/v3/account_items', {});
    return response.account_items || response.items || [];
  },

  /**
   * 接続テスト用：ユーザー・事業所情報を取得する
   * ★ エンドポイント名をドキュメントで確認
   */
  getMe() {
    // 事業所情報取得の候補エンドポイント
    const candidates = ['/api/v3/me', '/api/v3/user', '/api/v3/office'];
    for (const endpoint of candidates) {
      try {
        return MFApiClient._request('GET', endpoint, {});
      } catch (e) {
        Logger.log(`${endpoint} → ${e.message}`);
      }
    }
    throw new Error('ユーザー情報の取得に失敗しました。ドキュメントでエンドポイントを確認してください。');
  },

  // ==============================
  // 内部メソッド
  // ==============================

  /**
   * APIリクエストを実行する（トークンを自動管理）
   */
  _request(method, endpoint, params) {
    const token   = MFApiClient._getAccessToken();
    const baseUrl = CONFIG.MF_API.BASE_URL + endpoint;

    let url = baseUrl;
    const options = {
      method:  method.toLowerCase(),
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type':  'application/json',
        'Accept':        'application/json',
      },
      muteHttpExceptions: true,
    };

    if (method === 'GET' && Object.keys(params).length > 0) {
      const qs = Object.entries(params)
        .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
        .join('&');
      url = `${baseUrl}?${qs}`;
    } else if (method !== 'GET') {
      options.payload = JSON.stringify(params);
    }

    Logger.log(`API: ${method} ${url}`);
    const response   = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    const body       = response.getContentText();

    // 401: トークン期限切れ → リフレッシュして再試行（1回のみ）
    if (statusCode === 401) {
      Logger.log('Token expired → refreshing');
      MFApiClient._refreshAccessToken();
      const newToken = MFApiClient._getAccessToken();
      options.headers['Authorization'] = `Bearer ${newToken}`;
      const retry = UrlFetchApp.fetch(url, options);
      if (retry.getResponseCode() !== 200) {
        throw new Error(`再認証後もエラー [${retry.getResponseCode()}]: ${retry.getContentText()}`);
      }
      return JSON.parse(retry.getContentText());
    }

    if (statusCode < 200 || statusCode >= 300) {
      throw new Error(`MF API [${statusCode}] ${endpoint}: ${body}`);
    }

    return JSON.parse(body);
  },

  /**
   * 有効なアクセストークンを取得する
   */
  _getAccessToken() {
    const props = PropertiesService.getScriptProperties();
    const token = props.getProperty('MF_ACCESS_TOKEN');

    if (!token) {
      throw new Error(
        'MFアクセストークンが見つかりません。\n' +
        'メニュー「🔑 MF会計 認証を開始」から認証してください。'
      );
    }

    // 有効期限チェック（1分前にリフレッシュ）
    const expiresAt = parseInt(props.getProperty('MF_TOKEN_EXPIRES_AT') || '0');
    if (Date.now() > expiresAt - 60000) {
      MFApiClient._refreshAccessToken();
      return props.getProperty('MF_ACCESS_TOKEN');
    }

    return token;
  },

  /**
   * リフレッシュトークンを使ってアクセストークンを更新する
   */
  _refreshAccessToken() {
    const props        = PropertiesService.getScriptProperties();
    const refreshToken = props.getProperty('MF_REFRESH_TOKEN');

    if (!refreshToken) {
      throw new Error('リフレッシュトークンがありません。再認証が必要です。');
    }

    const response = UrlFetchApp.fetch(CONFIG.MF_API.TOKEN_URL, {
      method: 'post',
      payload: {
        grant_type:    'refresh_token',
        refresh_token: refreshToken,
        client_id:     CONFIG.MF_API.CLIENT_ID,
        client_secret: CONFIG.MF_API.CLIENT_SECRET,
      },
      muteHttpExceptions: true,
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(`トークンリフレッシュ失敗 [${response.getResponseCode()}]: ${response.getContentText()}`);
    }

    MFApiClient._storeTokens(JSON.parse(response.getContentText()));
    Logger.log('アクセストークン更新完了');
  },

  /**
   * 認証コードをトークンと交換する（Setup.gsから呼ばれる）
   */
  exchangeCodeForTokens(code) {
    const response = UrlFetchApp.fetch(CONFIG.MF_API.TOKEN_URL, {
      method: 'post',
      payload: {
        grant_type:    'authorization_code',
        code:          code,
        client_id:     CONFIG.MF_API.CLIENT_ID,
        client_secret: CONFIG.MF_API.CLIENT_SECRET,
        redirect_uri:  CONFIG.MF_API.REDIRECT_URI,
      },
      muteHttpExceptions: true,
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(`認証コード交換失敗 [${response.getResponseCode()}]: ${response.getContentText()}`);
    }

    const data = JSON.parse(response.getContentText());
    MFApiClient._storeTokens(data);
    Logger.log('認証完了 - トークン保存済み');
    return data;
  },

  /**
   * トークン情報をPropertiesServiceに保存する
   */
  _storeTokens(tokenData) {
    const props     = PropertiesService.getScriptProperties();
    const expiresAt = Date.now() + (tokenData.expires_in || 3600) * 1000;
    props.setProperties({
      'MF_ACCESS_TOKEN':     tokenData.access_token,
      'MF_REFRESH_TOKEN':    tokenData.refresh_token || props.getProperty('MF_REFRESH_TOKEN') || '',
      'MF_TOKEN_EXPIRES_AT': String(expiresAt),
    });
  },

  /**
   * 認証情報をすべてクリアする（再認証用）
   */
  clearTokens() {
    PropertiesService.getScriptProperties().deleteAllProperties();
    Logger.log('トークンをクリアしました');
  },
};
