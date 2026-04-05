/**
 * MoneyForward クラウド会計 → Google Drive 自動同期スクリプト
 *
 * 機能:
 * - MoneyForward API からレポートデータ（試算表・仕訳一覧等）を取得
 * - CSV/PDF形式で Google Drive の指定フォルダに自動保存
 * - タイマートリガーで定期実行（月次/週次）
 */

// ============================================================
// 設定（スクリプトプロパティから読み込み）
// ============================================================

/**
 * スクリプトプロパティのキー一覧:
 * - MF_CLIENT_ID       : MoneyForward APIのクライアントID
 * - MF_CLIENT_SECRET   : MoneyForward APIのクライアントシークレット
 * - MF_ACCESS_TOKEN    : アクセストークン（OAuth2認証後に自動設定）
 * - MF_REFRESH_TOKEN   : リフレッシュトークン（OAuth2認証後に自動設定）
 * - MF_OFFICE_ID       : 事業所ID
 * - GDRIVE_FOLDER_ID   : Google Drive保存先フォルダID
 */

var CONFIG = {
  MF_API_BASE: 'https://accounting.moneyforward.com/api/v3',
  MF_AUTH_URL: 'https://accounting.moneyforward.com/oauth/authorize',
  MF_TOKEN_URL: 'https://accounting.moneyforward.com/oauth/token',
  REDIRECT_URI: '' // initで設定
};

// ============================================================
// レポート名設定
// ============================================================

/**
 * 同期対象レポートの定義
 * ここでレポート名・取得対象・部門フィルタを自由にカスタマイズ可能
 *
 * 各エントリの設定:
 *   name     : ファイル名のプレフィックス（例: "物販_PL", "民泊_仕訳"）
 *   type     : レポート種別 "journals" | "trial_bs" | "trial_pl"
 *   department : 部門名（MF上の部門名と一致させる。nullなら全部門）
 *   enabled  : true/false で個別に有効・無効を切替
 */
var REPORT_DEFINITIONS = [
  // --- 物販部門 ---
  { name: '物販_PL',   type: 'trial_pl',  department: '物販部門', enabled: true },
  { name: '物販_BS',   type: 'trial_bs',  department: '物販部門', enabled: true },
  { name: '物販_仕訳', type: 'journals',   department: '物販部門', enabled: true },

  // --- 民泊部門 ---
  { name: '民泊_PL',   type: 'trial_pl',  department: '民泊部門', enabled: true },
  { name: '民泊_BS',   type: 'trial_bs',  department: '民泊部門', enabled: true },
  { name: '民泊_仕訳', type: 'journals',   department: '民泊部門', enabled: true },

  // --- 全社（部門指定なし） ---
  { name: '全社_PL',   type: 'trial_pl',  department: null, enabled: true },
  { name: '全社_BS',   type: 'trial_bs',  department: null, enabled: true },
  { name: '全社_仕訳', type: 'journals',   department: null, enabled: true }
];

/**
 * レポートのファイル名を生成
 * @param {string} reportName - レポート名プレフィックス（例: "物販_PL"）
 * @param {number} year - 年
 * @param {number} month - 月（nullなら年次）
 * @return {string} ファイル名（例: "物販_PL_2026_03.csv"）
 */
function buildFileName(reportName, year, month) {
  if (month) {
    return reportName + '_' + year + '_' + String(month).padStart(2, '0') + '.csv';
  }
  return reportName + '_' + year + '.csv';
}

// ============================================================
// OAuth2 認証
// ============================================================

/**
 * OAuth2サービスを取得
 */
function getMfOAuth2Service() {
  var props = PropertiesService.getScriptProperties();
  return OAuth2.createService('moneyforward')
    .setAuthorizationBaseUrl(CONFIG.MF_AUTH_URL)
    .setTokenUrl(CONFIG.MF_TOKEN_URL)
    .setClientId(props.getProperty('MF_CLIENT_ID'))
    .setClientSecret(props.getProperty('MF_CLIENT_SECRET'))
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('office_setting:read account_items:read journals:read reports:read')
    .setParam('access_type', 'offline');
}

/**
 * 認証URLを生成（初回セットアップ時にログで確認）
 */
function getAuthorizationUrl() {
  var service = getMfOAuth2Service();
  var authUrl = service.getAuthorizationUrl();
  Logger.log('以下のURLにアクセスして認証してください:');
  Logger.log(authUrl);
  return authUrl;
}

/**
 * OAuth2コールバック
 */
function authCallback(request) {
  var service = getMfOAuth2Service();
  var authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('認証成功！このタブを閉じてください。');
  } else {
    return HtmlService.createHtmlOutput('認証失敗。再度お試しください。');
  }
}

/**
 * 認証状態をリセット
 */
function resetAuth() {
  getMfOAuth2Service().reset();
  Logger.log('認証をリセットしました。');
}

// ============================================================
// MoneyForward API リクエスト
// ============================================================

/**
 * MoneyForward APIにリクエストを送信
 * @param {string} endpoint - APIエンドポイント（例: '/journals'）
 * @param {Object} params - クエリパラメータ
 * @return {Object} レスポンスデータ
 */
function mfApiRequest(endpoint, params) {
  var service = getMfOAuth2Service();
  if (!service.hasAccess()) {
    throw new Error('MoneyForward未認証です。getAuthorizationUrl()を実行して認証してください。');
  }

  var props = PropertiesService.getScriptProperties();
  var officeId = props.getProperty('MF_OFFICE_ID');

  var queryString = '';
  if (params) {
    var queryParts = [];
    for (var key in params) {
      queryParts.push(encodeURIComponent(key) + '=' + encodeURIComponent(params[key]));
    }
    queryString = '?' + queryParts.join('&');
  }

  var url = CONFIG.MF_API_BASE + '/offices/' + officeId + endpoint + queryString;
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + service.getAccessToken(),
      'Accept': 'application/json'
    },
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  if (code !== 200) {
    throw new Error('MF API エラー (' + code + '): ' + response.getContentText());
  }

  return JSON.parse(response.getContentText());
}

// ============================================================
// レポート取得関数
// ============================================================

/**
 * 仕訳一覧を取得
 * @param {number} year - 年
 * @param {number} month - 月
 * @param {string|null} department - 部門名（nullなら全部門）
 * @return {Array} 仕訳データ
 */
function getJournals(year, month, department) {
  var startDate = Utilities.formatDate(new Date(year, month - 1, 1), 'Asia/Tokyo', 'yyyy-MM-dd');
  var lastDay = new Date(year, month, 0).getDate();
  var endDate = Utilities.formatDate(new Date(year, month - 1, lastDay), 'Asia/Tokyo', 'yyyy-MM-dd');

  var params = {
    start_date: startDate,
    end_date: endDate
  };
  if (department) {
    params.department = department;
  }

  var data = mfApiRequest('/journals', params);
  return data.journals || [];
}

/**
 * 試算表を取得
 * @param {number} year - 年
 * @param {number} month - 月
 * @param {string|null} department - 部門名（nullなら全部門）
 * @return {Object} 試算表データ
 */
function getTrialBalance(year, month, department) {
  var params = {
    fiscal_year: year,
    month: month
  };
  if (department) {
    params.department = department;
  }
  var data = mfApiRequest('/reports/trial_bs', params);
  return data;
}

/**
 * 損益計算書データを取得
 * @param {number} year - 年
 * @param {number} month - 月
 * @param {string|null} department - 部門名（nullなら全部門）
 * @return {Object} PLデータ
 */
function getProfitAndLoss(year, month, department) {
  var params = {
    fiscal_year: year,
    month: month
  };
  if (department) {
    params.department = department;
  }
  var data = mfApiRequest('/reports/trial_pl', params);
  return data;
}

// ============================================================
// CSV変換
// ============================================================

/**
 * 仕訳データをCSV文字列に変換
 * @param {Array} journals - 仕訳データ配列
 * @return {string} CSV文字列
 */
function journalsToCsv(journals) {
  var headers = ['日付', '伝票番号', '借方勘定科目', '借方金額', '貸方勘定科目', '貸方金額', '摘要'];
  var rows = [headers.join(',')];

  journals.forEach(function(j) {
    var details = j.details || [];
    details.forEach(function(d) {
      var row = [
        j.date || '',
        j.slip_number || '',
        d.debit_account_item_name || '',
        d.debit_amount || 0,
        d.credit_account_item_name || '',
        d.credit_amount || 0,
        (j.description || '').replace(/,/g, '、').replace(/\n/g, ' ')
      ];
      rows.push(row.join(','));
    });
  });

  return rows.join('\n');
}

/**
 * 試算表データをCSV文字列に変換
 * @param {Object} trialData - 試算表データ
 * @param {string} type - 'bs' or 'pl'
 * @return {string} CSV文字列
 */
function trialBalanceToCsv(trialData, type) {
  var headers = ['勘定科目', '借方残高', '貸方残高', '借方発生', '貸方発生'];
  var rows = [headers.join(',')];

  var items = trialData.trial_balance_items || trialData.items || [];
  items.forEach(function(item) {
    var row = [
      (item.account_item_name || '').replace(/,/g, '、'),
      item.debit_closing_balance || 0,
      item.credit_closing_balance || 0,
      item.debit_amount || 0,
      item.credit_amount || 0
    ];
    rows.push(row.join(','));
  });

  return rows.join('\n');
}

// ============================================================
// Google Drive 保存
// ============================================================

/**
 * Google Driveの指定フォルダにファイルを保存
 * @param {string} fileName - ファイル名
 * @param {string} content - ファイル内容
 * @param {string} mimeType - MIMEタイプ
 * @return {File} 保存されたファイル
 */
function saveToDrive(fileName, content, mimeType) {
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty('GDRIVE_FOLDER_ID');

  var folder;
  if (folderId) {
    folder = DriveApp.getFolderById(folderId);
  } else {
    // フォルダ未設定の場合、ルートに「MoneyForward_Reports」フォルダを作成
    var folders = DriveApp.getFoldersByName('MoneyForward_Reports');
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder('MoneyForward_Reports');
      // フォルダIDを保存
      props.setProperty('GDRIVE_FOLDER_ID', folder.getId());
    }
    Logger.log('保存先フォルダ: ' + folder.getName() + ' (ID: ' + folder.getId() + ')');
  }

  // 年月サブフォルダを作成
  var now = new Date();
  var subFolderName = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM');
  var subFolders = folder.getFoldersByName(subFolderName);
  var subFolder;
  if (subFolders.hasNext()) {
    subFolder = subFolders.next();
  } else {
    subFolder = folder.createFolder(subFolderName);
  }

  // 同名ファイルがあれば上書き（削除→再作成）
  var existingFiles = subFolder.getFilesByName(fileName);
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }

  var blob = Utilities.newBlob(content, mimeType, fileName);
  var file = subFolder.createFile(blob);

  Logger.log('保存完了: ' + subFolder.getName() + '/' + fileName);
  return file;
}

// ============================================================
// メイン実行関数
// ============================================================

/**
 * REPORT_DEFINITIONS に基づいて1件のレポートを取得・保存
 * @param {Object} def - REPORT_DEFINITIONS の1エントリ
 * @param {number} year - 年
 * @param {number} month - 月
 * @return {string} 結果メッセージ
 */
function syncOneReport(def, year, month) {
  var fileName = buildFileName(def.name, year, month);
  var csv;

  switch (def.type) {
    case 'journals':
      var journals = getJournals(year, month, def.department);
      if (journals.length === 0) return def.name + ': データなし';
      csv = journalsToCsv(journals);
      saveToDrive(fileName, csv, 'text/csv');
      return def.name + ': ' + journals.length + '件';

    case 'trial_bs':
      var bsData = getTrialBalance(year, month, def.department);
      csv = trialBalanceToCsv(bsData, 'bs');
      saveToDrive(fileName, csv, 'text/csv');
      return def.name + ': 保存完了';

    case 'trial_pl':
      var plData = getProfitAndLoss(year, month, def.department);
      csv = trialBalanceToCsv(plData, 'pl');
      saveToDrive(fileName, csv, 'text/csv');
      return def.name + ': 保存完了';

    default:
      return def.name + ': 不明なtype "' + def.type + '"';
  }
}

/**
 * 月次レポートを取得してGoogle Driveに保存（メイン関数）
 * REPORT_DEFINITIONS で定義されたレポートを順に処理する
 * タイマートリガーからこの関数を呼び出す
 */
function syncMonthlyReports() {
  var now = new Date();
  // 前月のデータを取得（月初に前月分を同期する想定）
  var targetDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  var year = targetDate.getFullYear();
  var month = targetDate.getMonth() + 1;

  syncReportsByDefinitions(year, month);
}

/**
 * 指定年月のレポートを手動で取得
 * @param {number} year - 年（省略時: 今年）
 * @param {number} month - 月（省略時: 今月）
 */
function syncReportsManual(year, month) {
  if (!year || !month) {
    var now = new Date();
    year = now.getFullYear();
    month = now.getMonth() + 1;
  }
  syncReportsByDefinitions(year, month);
}

/**
 * 特定のレポートだけを手動で取得
 * 例: syncSingleReport('物販_PL', 2026, 3)
 * @param {string} reportName - REPORT_DEFINITIONS の name と一致させる
 * @param {number} year - 年
 * @param {number} month - 月
 */
function syncSingleReport(reportName, year, month) {
  var def = REPORT_DEFINITIONS.filter(function(d) { return d.name === reportName; })[0];
  if (!def) {
    Logger.log('エラー: "' + reportName + '" はREPORT_DEFINITIONSに定義されていません。');
    Logger.log('利用可能なレポート名:');
    REPORT_DEFINITIONS.forEach(function(d) { Logger.log('  - ' + d.name); });
    return;
  }

  Logger.log('=== 単体同期: ' + reportName + ' (' + year + '年' + month + '月) ===');
  try {
    var result = syncOneReport(def, year, month);
    Logger.log(result);
  } catch (e) {
    Logger.log(def.name + ': エラー - ' + e.message);
  }
}

/**
 * REPORT_DEFINITIONS の全有効レポートを同期する内部関数
 */
function syncReportsByDefinitions(year, month) {
  Logger.log('=== MoneyForward → Google Drive 同期開始 ===');
  Logger.log('対象期間: ' + year + '年' + month + '月');

  var enabledReports = REPORT_DEFINITIONS.filter(function(d) { return d.enabled; });
  Logger.log('対象レポート数: ' + enabledReports.length + '件');

  var results = [];
  enabledReports.forEach(function(def) {
    try {
      var result = syncOneReport(def, year, month);
      results.push(result);
    } catch (e) {
      results.push(def.name + ': エラー - ' + e.message);
    }
  });

  Logger.log('=== 実行結果 ===');
  results.forEach(function(r) { Logger.log(r); });
  Logger.log('=== 同期完了 ===');

  sendNotification(year, month, results);
}

// ============================================================
// 通知
// ============================================================

/**
 * 同期結果をメールで通知
 */
function sendNotification(year, month, results) {
  var props = PropertiesService.getScriptProperties();
  var email = props.getProperty('NOTIFICATION_EMAIL');
  if (!email) return;

  var subject = '【MF→Drive同期】' + year + '年' + month + '月 レポート保存完了';
  var body = 'MoneyForward → Google Drive 同期結果\n\n';
  body += '対象期間: ' + year + '年' + month + '月\n';
  body += '実行日時: ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss') + '\n\n';
  body += '--- 結果 ---\n';
  results.forEach(function(r) { body += '・' + r + '\n'; });

  GmailApp.sendEmail(email, subject, body);
}

// ============================================================
// トリガー管理
// ============================================================

/**
 * 月次自動実行トリガーを設定（毎月1日 AM9:00）
 */
function setupMonthlyTrigger() {
  // 既存トリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'syncMonthlyReports') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 新規トリガー作成：毎月1日 9:00-10:00
  ScriptApp.newTrigger('syncMonthlyReports')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();

  Logger.log('月次トリガーを設定しました（毎月1日 9:00-10:00）');
}

/**
 * 週次自動実行トリガーを設定（毎週月曜 AM9:00）
 */
function setupWeeklyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'syncMonthlyReports') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('syncMonthlyReports')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();

  Logger.log('週次トリガーを設定しました（毎週月曜 9:00-10:00）');
}

/**
 * 全トリガーを削除
 */
function removeAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    ScriptApp.deleteTrigger(trigger);
  });
  Logger.log('全トリガーを削除しました。');
}

// ============================================================
// 初期セットアップ確認
// ============================================================

/**
 * セットアップ状態を確認
 */
function checkSetup() {
  var props = PropertiesService.getScriptProperties();
  var checks = {
    'MF_CLIENT_ID': props.getProperty('MF_CLIENT_ID') ? 'OK' : '未設定',
    'MF_CLIENT_SECRET': props.getProperty('MF_CLIENT_SECRET') ? 'OK' : '未設定',
    'MF_OFFICE_ID': props.getProperty('MF_OFFICE_ID') ? 'OK' : '未設定',
    'GDRIVE_FOLDER_ID': props.getProperty('GDRIVE_FOLDER_ID') ? 'OK' : '未設定（自動作成されます）',
    'NOTIFICATION_EMAIL': props.getProperty('NOTIFICATION_EMAIL') ? 'OK' : '未設定（オプション）',
    'OAuth2認証': getMfOAuth2Service().hasAccess() ? '認証済み' : '未認証'
  };

  Logger.log('=== セットアップ状態 ===');
  for (var key in checks) {
    Logger.log(key + ': ' + checks[key]);
  }

  return checks;
}
