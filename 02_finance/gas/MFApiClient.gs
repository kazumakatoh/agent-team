/**
 * 財務レポート自動化システム - マネーフォワード クラウド会計 APIクライアント
 *
 * ■ 確認済みエンドポイント
 *   PL試算表（全社）: GET /api/v3/reports/trial_balance_pl?start_date=&end_date=
 *     → 部門フィルタ非対応（department_id は unsupported_query_parameter）
 *   仕訳一覧:         GET /api/v3/journals?start_date=&end_date=&page=&per_page=
 *     → 各branch の creditor/debtor に department_id(URL-encoded) が付与される
 *   部門一覧:         GET /api/v3/departments
 *
 * ■ 部門別PL取得方式
 *   trial_balance_pl は部門フィルタ不可のため、仕訳APIで全仕訳を取得し
 *   department_id でアプリ側フィルタ → 勘定科目ごとに credit/debit を集計する
 *
 * ■ 仕訳レスポンス構造
 *   journals[].branches[].creditor/debtor = {
 *     account_id, account_name, department_id(URL-encoded or null), value, ...
 *   }
 *   department_id は URL-encoded base64 (例: "s4afNgEETQf7VyrB0IqZ1Q%3D%3D")
 */

const MFApiClient = {

  // 実行中の仕訳キャッシュ（同一月を複数部門で処理する際の重複API呼び出しを防ぐ）
  _journalCache: {},

  // ==============================
  // 公開メソッド
  // ==============================

  /**
   * PL試算表を取得する（PLFormatter から呼ばれる）
   *
   * @param {string} startDate   - 'YYYY-MM-DD'
   * @param {string} endDate     - 'YYYY-MM-DD'
   * @param {string} [segmentId] - 部門ID（Config.DEPARTMENTS[].id の値）。空/nullなら全社
   * @return {Array} PLFormatter._buildAccountMap() に渡せる配列
   *   - segmentId なし: trial_balance_pl の rows 配列（ネスト構造）
   *   - segmentId あり: 仕訳から集計したフラット配列 [{name, type:'account', values:[...]}]
   */
  getTrialBalance(startDate, endDate, segmentId) {
    if (segmentId) {
      // 部門あり → 仕訳API経由で集計
      return MFApiClient._getTrialBalanceFromJournals(startDate, endDate, segmentId);
    }

    // 全社 → trial_balance_pl API（高速・ネスト構造で返る）
    const response = MFApiClient._request('GET', '/api/v3/reports/trial_balance_pl', {
      start_date: startDate,
      end_date:   endDate,
    });
    const rows = response.rows || [];
    Logger.log(`試算表（全社）: ${startDate}〜${endDate} → ${rows.length}カテゴリ`);
    return rows;
  },

  /**
   * 仕訳APIで全仕訳を取得し、部門でフィルタして勘定科目別残高を返す
   * @private
   */
  _getTrialBalanceFromJournals(startDate, endDate, segmentId) {
    // Config の segmentId は decoded base64 (例: "s4afNgEETQf7VyrB0IqZ1Q==")
    // API の department_id は URL-encoded (例: "s4afNgEETQf7VyrB0IqZ1Q%3D%3D")
    const encodedDeptId = encodeURIComponent(segmentId);

    const allJournals = MFApiClient._getAllJournals(startDate, endDate);
    Logger.log(`部門別試算表: ${segmentId} / ${startDate}〜${endDate} / 全仕訳${allJournals.length}件から集計`);

    // 勘定科目ごとに credit・debit を集計
    const accountData = {}; // { accountName: { credit: 0, debit: 0 } }

    allJournals.forEach(journal => {
      (journal.branches || []).forEach(branch => {
        const cr = branch.creditor;
        const dr = branch.debtor;

        // creditor（貸方）が指定部門に属する場合
        if (cr && cr.department_id === encodedDeptId && cr.account_name) {
          if (!accountData[cr.account_name]) accountData[cr.account_name] = { credit: 0, debit: 0 };
          accountData[cr.account_name].credit += (cr.value || 0);
        }

        // debtor（借方）が指定部門に属する場合
        if (dr && dr.department_id === encodedDeptId && dr.account_name) {
          if (!accountData[dr.account_name]) accountData[dr.account_name] = { credit: 0, debit: 0 };
          accountData[dr.account_name].debit += (dr.value || 0);
        }
      });
    });

    const accountCount = Object.keys(accountData).length;
    Logger.log(`  → 部門付き勘定科目: ${accountCount}件 (${Object.keys(accountData).slice(0, 5).join(', ')}...)`);

    // PLFormatter._buildAccountMap が期待する形式（type:'account', values[3]=closing_balance）に変換
    // closing_balance = credit - debit
    //   収益科目: credit > debit → 正値 (収益)
    //   費用科目: debit > credit → 負値 → Math.abs で費用額
    return Object.entries(accountData).map(([name, { credit, debit }]) => ({
      name,
      type:   'account',
      values: [0, debit, credit, credit - debit, 0], // [opening, debit, credit, closing, ratio]
      rows:   null,
    }));
  },

  /**
   * 指定期間の全仕訳を取得する（ページネーション対応・実行内キャッシュ）
   * per_page=10000 (APIの上限値) を使用して呼び出し回数を最小化する
   * @private
   */
  _getAllJournals(startDate, endDate) {
    const cacheKey = `${startDate}_${endDate}`;
    if (MFApiClient._journalCache[cacheKey]) {
      return MFApiClient._journalCache[cacheKey];
    }

    const allJournals = [];
    let page    = 1;
    const perPage = 10000; // OpenAPI spec の上限値（旧: 100）
    const MAX_PAGES = 10;  // 安全上限（月100,000件まで対応）

    while (page <= MAX_PAGES) {
      const response = MFApiClient._request('GET', '/api/v3/journals', {
        start_date: startDate,
        end_date:   endDate,
        page:       page,
        per_page:   perPage,
      });
      const journals = response.journals || [];
      allJournals.push(...journals);
      Logger.log(`  仕訳取得: page=${page}, ${journals.length}件 (累計 ${allJournals.length}件)`);

      if (journals.length < perPage) break; // 最終ページ
      page++;
    }

    MFApiClient._journalCache[cacheKey] = allJournals;
    return allJournals;
  },

  /**
   * 月次推移PLを一括取得する（全12ヶ月を1回のAPIコールで取得）
   * trial_balance_pl を12回呼ぶ代わりにこちらを使うと高速
   *
   * @param {number} fiscalYear  - 事業年度開始年（例: 2025）
   * @param {number} startMonth  - 取得開始月（デフォルト: 事業年度開始月）
   * @param {number} endMonth    - 取得終了月（デフォルト: 12）
   * @return {Object} APIレスポンス（rows構造はtrial_balance_plと同形だが
   *   values[]が月次配列になる）
   *
   * ■ レスポンス構造（予想・要実機確認）
   *   {
   *     columns: ["4月", "5月", ..., "3月"],
   *     rows: [
   *       { name:"売上高合計", type:"financial_statement_item",
   *         values:[月1合計, 月2合計, ...], rows:[
   *           { name:"売上（国内）", type:"account", values:[月1, 月2, ...] },
   *           ...
   *         ]}
   *     ]
   *   }
   */
  getTransitionPL(fiscalYear, startMonth, endMonth) {
    const params = {
      type:        'monthly',
      fiscal_year: fiscalYear,
    };
    if (startMonth) params.start_month = startMonth;
    if (endMonth)   params.end_month   = endMonth;

    const response = MFApiClient._request('GET', '/api/v3/reports/transition_pl', params);
    Logger.log(`推移PL取得: FY${fiscalYear} → カラム数: ${(response.columns || []).length}`);
    return response;
  },

  /**
   * BS試算表を取得する（将来のBS対応用に準備）
   */
  getBalanceSheet(startDate, endDate) {
    const response = MFApiClient._request('GET', '/api/v3/reports/trial_balance_bs', {
      start_date: startDate,
      end_date:   endDate,
    });
    return response.rows || response.items || response.account_items || [];
  },

  /**
   * 部門一覧を取得する
   */
  getDepartments() {
    const candidates = ['/api/v3/departments', '/api/v3/segments'];
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
   * 勘定科目一覧を取得する
   * 確認済み: /api/v3/accounts が正しいエンドポイント（account_items ではない）
   */
  getAccountItems() {
    const response = MFApiClient._request('GET', '/api/v3/accounts', {});
    return response.accounts || response.account_items || response.items || [];
  },

  /**
   * 接続テスト用
   */
  getMe() {
    const candidates = ['/api/v3/me', '/api/v3/user', '/api/v3/office'];
    for (const endpoint of candidates) {
      try {
        return MFApiClient._request('GET', endpoint, {});
      } catch (e) {
        Logger.log(`${endpoint} → ${e.message}`);
      }
    }
    throw new Error('ユーザー情報の取得に失敗しました。');
  },

  // ==============================
  // 内部メソッド
  // ==============================

  /**
   * APIリクエストを実行する（トークン自動管理）
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
   * 有効なアクセストークンを取得する（OAuth2のみ）
   * ※ app-portal.moneyforward.com の APIキーは api-accounting.moneyforward.com では
   *    401 を返すため使用しない（確認済み）
   */
  _getAccessToken() {
    const props = PropertiesService.getScriptProperties();

    // OAuth2トークン
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
   * MFビジネスプラットフォームは CLIENT_SECRET_BASIC 方式
   */
  _refreshAccessToken() {
    const props        = PropertiesService.getScriptProperties();
    const refreshToken = props.getProperty('MF_REFRESH_TOKEN');

    if (!refreshToken) {
      throw new Error('リフレッシュトークンがありません。再認証が必要です。');
    }

    const credentials = Utilities.base64Encode(
      CONFIG.MF_API.CLIENT_ID + ':' + CONFIG.MF_API.CLIENT_SECRET
    );
    const response = UrlFetchApp.fetch(CONFIG.MF_API.TOKEN_URL, {
      method: 'post',
      headers: { 'Authorization': 'Basic ' + credentials },
      payload: {
        grant_type:    'refresh_token',
        refresh_token: refreshToken,
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
   * CLIENT_SECRET_BASIC 方式（Authorizationヘッダに認証情報）
   */
  exchangeCodeForTokens(code) {
    const credentials = Utilities.base64Encode(
      CONFIG.MF_API.CLIENT_ID + ':' + CONFIG.MF_API.CLIENT_SECRET
    );
    const response = UrlFetchApp.fetch(CONFIG.MF_API.TOKEN_URL, {
      method: 'post',
      headers: { 'Authorization': 'Basic ' + credentials },
      payload: {
        grant_type:   'authorization_code',
        code:         code,
        redirect_uri: CONFIG.MF_API.REDIRECT_URI,
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
