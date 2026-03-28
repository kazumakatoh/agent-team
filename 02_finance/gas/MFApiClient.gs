/**
 * 財務レポート自動化システム - マネーフォワード クラウド会計 APIクライアント
 *
 * OAuth 2.0 Authorization Code Flow（OOBリダイレクト方式）
 * トークンは PropertiesService に保存・自動リフレッシュ
 */

const MFApiClient = {

  // ==============================
  // 公開メソッド
  // ==============================

  /**
   * 試算表（PL）を取得する
   * @param {string} startDate - 'YYYY-MM-DD'
   * @param {string} endDate   - 'YYYY-MM-DD'
   * @param {string} [departmentId] - 部門ID（省略時は全社）
   * @return {Array} 試算表アイテム配列
   */
  getTrialBalance(startDate, endDate, departmentId) {
    const companyId = MFApiClient._getCompanyId();
    const params = {
      start_date: startDate,
      end_date:   endDate,
    };
    if (departmentId) params['department_id'] = departmentId;

    const response = MFApiClient._request(
      'GET',
      `/api/v3/companies/${companyId}/trial_balance`,
      params
    );

    return (response.trial_balance && response.trial_balance.items) || [];
  },

  /**
   * 部門一覧を取得する
   * @return {Array} 部門配列
   */
  getDepartments() {
    const companyId = MFApiClient._getCompanyId();
    const response  = MFApiClient._request('GET', `/api/v3/companies/${companyId}/departments`, {});
    return response.departments || [];
  },

  /**
   * 事業所一覧を取得してIDを返す（初期設定用）
   * @return {Array} 事業所配列
   */
  getCompanies() {
    const response = MFApiClient._request('GET', '/api/v3/companies', {});
    return response.companies || [];
  },

  /**
   * 勘定科目一覧を取得する（マッピング確認用）
   * @return {Array} 勘定科目配列
   */
  getAccountItems() {
    const companyId = MFApiClient._getCompanyId();
    const response  = MFApiClient._request('GET', `/api/v3/companies/${companyId}/account_items`, {});
    return response.account_items || [];
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

    let url     = baseUrl;
    let options = {
      method:  method.toLowerCase(),
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type':  'application/json',
        'Accept':        'application/json',
      },
      muteHttpExceptions: true,
    };

    if (method === 'GET' && Object.keys(params).length > 0) {
      const qs = Object.entries(params).map(([k, v]) => `${k}=${encodeURIComponent(v)}`).join('&');
      url = `${baseUrl}?${qs}`;
    } else if (method !== 'GET') {
      options.payload = JSON.stringify(params);
    }

    Logger.log(`API Request: ${method} ${url}`);
    const response   = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    const body       = response.getContentText();

    // 401: トークン期限切れ → リフレッシュして再試行
    if (statusCode === 401) {
      Logger.log('Token expired, refreshing...');
      MFApiClient._refreshAccessToken();
      const newToken = MFApiClient._getAccessToken();
      options.headers['Authorization'] = `Bearer ${newToken}`;
      const retryResp = UrlFetchApp.fetch(url, options);
      return JSON.parse(retryResp.getContentText());
    }

    if (statusCode < 200 || statusCode >= 300) {
      throw new Error(`MF API Error [${statusCode}]: ${body}`);
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
        'Setup.gs の authorize() を実行して認証してください。'
      );
    }

    // 有効期限チェック
    const expiresAt = parseInt(props.getProperty('MF_TOKEN_EXPIRES_AT') || '0');
    if (Date.now() > expiresAt - 60000) { // 1分前にリフレッシュ
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
      throw new Error(`トークンリフレッシュ失敗: ${response.getContentText()}`);
    }

    const data = JSON.parse(response.getContentText());
    MFApiClient._storeTokens(data);
    Logger.log('アクセストークンを更新しました');
  },

  /**
   * 認証コードをトークンと交換する（Setup.gsから呼ばれる）
   * @param {string} code - 認証コード
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
      throw new Error(`認証コード交換失敗: ${response.getContentText()}`);
    }

    const data = JSON.parse(response.getContentText());
    MFApiClient._storeTokens(data);
    Logger.log('認証完了 - トークンを保存しました');
    return data;
  },

  /**
   * トークン情報をPropertiesServiceに保存する
   */
  _storeTokens(tokenData) {
    const props     = PropertiesService.getScriptProperties();
    const expiresAt = Date.now() + (tokenData.expires_in || 3600) * 1000;
    props.setProperties({
      'MF_ACCESS_TOKEN':   tokenData.access_token,
      'MF_REFRESH_TOKEN':  tokenData.refresh_token || props.getProperty('MF_REFRESH_TOKEN') || '',
      'MF_TOKEN_EXPIRES_AT': String(expiresAt),
    });
  },

  /**
   * 事業所IDを取得する（設定値 or 自動取得）
   */
  _getCompanyId() {
    if (CONFIG.MF_API.COMPANY_ID) return CONFIG.MF_API.COMPANY_ID;

    // 設定がない場合は自動取得して最初の事業所IDを使用
    const props = PropertiesService.getScriptProperties();
    const cached = props.getProperty('MF_COMPANY_ID');
    if (cached) return cached;

    const companies = MFApiClient.getCompanies();
    if (!companies.length) throw new Error('事業所が見つかりません');
    const id = companies[0].id;
    props.setProperty('MF_COMPANY_ID', id);
    Logger.log(`事業所ID自動取得: ${id} (${companies[0].name})`);
    return id;
  },

  /**
   * 認証情報をすべてクリアする（再認証用）
   */
  clearTokens() {
    PropertiesService.getScriptProperties().deleteAllProperties();
    Logger.log('トークンをクリアしました');
  },
};
