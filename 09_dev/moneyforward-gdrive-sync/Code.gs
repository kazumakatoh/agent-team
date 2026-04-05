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
 *   name       : ファイル名の部門プレフィックス（例: "物販", "民泊"）
 *   type       : レポート種別 "journals" | "trial_bs" | "trial_pl"
 *   department : 部門名（MF上の部門名と一致させる。nullなら全部門）
 *   enabled    : true/false で個別に有効・無効を切替
 *
 * ファイル名の形式: {部門名}_{レポート種別}_{決算年}.csv
 *   例: 物販_PL_2026.csv, 民泊_BS_2027.csv, 共通_仕訳_2026.csv
 *
 * 決算年について（LEVEL1: 決算月=2月）:
 *   第8期 (2025/3〜2026/2) → 決算年=2026
 *   第9期 (2026/3〜2027/2) → 決算年=2027
 */
var REPORT_DEFINITIONS = [
  // --- 共通（部門指定なし＝全社） ---
  { name: '共通', type: 'trial_pl',  department: null, enabled: true },
  { name: '共通', type: 'trial_bs',  department: null, enabled: true },
  { name: '共通', type: 'journals',  department: null, enabled: true },

  // --- 物販 ---
  { name: '物販', type: 'trial_pl',  department: '物販', enabled: true },
  { name: '物販', type: 'trial_bs',  department: '物販', enabled: true },
  { name: '物販', type: 'journals',  department: '物販', enabled: true },

  // --- ブランド ---
  { name: 'ブランド', type: 'trial_pl',  department: 'ブランド', enabled: true },
  { name: 'ブランド', type: 'trial_bs',  department: 'ブランド', enabled: true },
  { name: 'ブランド', type: 'journals',  department: 'ブランド', enabled: true },

  // --- 民泊 ---
  { name: '民泊', type: 'trial_pl',  department: '民泊', enabled: true },
  { name: '民泊', type: 'trial_bs',  department: '民泊', enabled: true },
  { name: '民泊', type: 'journals',  department: '民泊', enabled: true }
];

/**
 * レポート種別コードから日本語ラベルを返す
 */
var REPORT_TYPE_LABELS = {
  'trial_pl': 'PL',
  'trial_bs': 'BS',
  'journals': '仕訳'
};

/**
 * 現在の日付から決算年（期末年）を算出する
 * LEVEL1の決算月は2月なので、3月〜翌2月が1期
 *   例: 2026年3月〜2027年2月 → 決算年2027（第9期）
 *        2025年3月〜2026年2月 → 決算年2026（第8期）
 * @param {Date} date - 基準日（省略時は今日）
 * @return {number} 決算年
 */
function getCurrentFiscalYear(date) {
  date = date || new Date();
  var year = date.getFullYear();
  var month = date.getMonth() + 1;
  return month >= 3 ? year + 1 : year;
}

/**
 * 決算年から期の開始日・終了日を返す
 * @param {number} fiscalYear - 決算年（例: 2026）
 * @return {Object} { startDate: '2025-03-01', endDate: '2026-02-28' }
 */
function getFiscalYearRange(fiscalYear) {
  var startYear = fiscalYear - 1;
  var lastDay = new Date(fiscalYear, 2, 0).getDate(); // 2月の末日
  return {
    startDate: startYear + '-03-01',
    endDate: fiscalYear + '-02-' + String(lastDay).padStart(2, '0')
  };
}

/**
 * レポートのファイル名を生成
 * 形式: {部門名}_{レポート種別}_{決算年}.csv
 * 例: 物販_PL_2026.csv, 民泊_BS_2027.csv
 * @param {string} departmentName - 部門名（例: "物販"）
 * @param {string} reportType - レポート種別コード
 * @param {number} fiscalYear - 決算年
 * @return {string} ファイル名
 */
function buildFileName(departmentName, reportType, fiscalYear) {
  var typeLabel = REPORT_TYPE_LABELS[reportType] || reportType;
  return departmentName + '_' + typeLabel + '_' + fiscalYear + '.csv';
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
 * 年間仕訳一覧を取得（期全体: 3月〜翌2月）
 * @param {number} fiscalYear - 決算年（例: 2026）
 * @param {string|null} department - 部門名（nullなら全部門）
 * @return {Array} 仕訳データ
 */
function getJournals(fiscalYear, department) {
  var range = getFiscalYearRange(fiscalYear);
  var params = {
    start_date: range.startDate,
    end_date: range.endDate
  };
  if (department) {
    params.department = department;
  }

  var data = mfApiRequest('/journals', params);
  return data.journals || [];
}

/**
 * 試算表（BS）を取得（年間累計）
 * @param {number} fiscalYear - 決算年
 * @param {string|null} department - 部門名（nullなら全部門）
 * @return {Object} 試算表データ
 */
function getTrialBalance(fiscalYear, department) {
  var params = {
    fiscal_year: fiscalYear
  };
  if (department) {
    params.department = department;
  }
  var data = mfApiRequest('/reports/trial_bs', params);
  return data;
}

/**
 * 損益計算書（PL）を取得（年間累計）
 * @param {number} fiscalYear - 決算年
 * @param {string|null} department - 部門名（nullなら全部門）
 * @return {Object} PLデータ
 */
function getProfitAndLoss(fiscalYear, department) {
  var params = {
    fiscal_year: fiscalYear
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
      props.setProperty('GDRIVE_FOLDER_ID', folder.getId());
    }
    Logger.log('保存先フォルダ: ' + folder.getName() + ' (ID: ' + folder.getId() + ')');
  }

  // 同名ファイルがあれば上書き（削除→再作成）で常に最新版を維持
  var existingFiles = folder.getFilesByName(fileName);
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }

  var blob = Utilities.newBlob(content, mimeType, fileName);
  var file = folder.createFile(blob);

  Logger.log('保存完了: ' + fileName);
  return file;
}

// ============================================================
// メイン実行関数
// ============================================================

/**
 * REPORT_DEFINITIONS に基づいて1件のレポートを取得・保存
 * @param {Object} def - REPORT_DEFINITIONS の1エントリ
 * @param {number} fiscalYear - 決算年（例: 2026）
 * @return {string} 結果メッセージ
 */
function syncOneReport(def, fiscalYear) {
  var fileName = buildFileName(def.name, def.type, fiscalYear);
  var csv;

  switch (def.type) {
    case 'journals':
      var journals = getJournals(fiscalYear, def.department);
      if (journals.length === 0) return fileName + ': データなし';
      csv = journalsToCsv(journals);
      saveToDrive(fileName, csv, 'text/csv');
      return fileName + ' (' + journals.length + '件)';

    case 'trial_bs':
      var bsData = getTrialBalance(fiscalYear, def.department);
      csv = trialBalanceToCsv(bsData, 'bs');
      saveToDrive(fileName, csv, 'text/csv');
      return fileName + ' 保存完了';

    case 'trial_pl':
      var plData = getProfitAndLoss(fiscalYear, def.department);
      csv = trialBalanceToCsv(plData, 'pl');
      saveToDrive(fileName, csv, 'text/csv');
      return fileName + ' 保存完了';

    default:
      return fileName + ': 不明なtype "' + def.type + '"';
  }
}

/**
 * 年間レポート同期（メイン関数）
 * 現在の期の年間データをMFから取得し、Google Driveに上書き保存する
 * 毎月トリガーで実行 → 同じファイル名で最新データに更新される
 */
function syncReports() {
  var fiscalYear = getCurrentFiscalYear();
  syncReportsByFiscalYear(fiscalYear);
}

/**
 * 決算年を指定して手動で同期
 * 例: syncReportsForYear(2026)  → 第8期（2025/3〜2026/2）のデータ
 *     syncReportsForYear(2027)  → 第9期（2026/3〜2027/2）のデータ
 * @param {number} fiscalYear - 決算年
 */
function syncReportsForYear(fiscalYear) {
  if (!fiscalYear) {
    fiscalYear = getCurrentFiscalYear();
  }
  syncReportsByFiscalYear(fiscalYear);
}

/**
 * 特定のレポートだけを手動で取得
 * 例: syncSingleReport('物販', 'trial_pl', 2026)
 *     → 物販_PL_2026.csv
 * @param {string} departmentName - 部門名（"共通","物販","ブランド","民泊"）
 * @param {string} reportType - "trial_pl" | "trial_bs" | "journals"
 * @param {number} fiscalYear - 決算年
 */
function syncSingleReport(departmentName, reportType, fiscalYear) {
  var def = REPORT_DEFINITIONS.filter(function(d) {
    return d.name === departmentName && d.type === reportType;
  })[0];
  if (!def) {
    Logger.log('エラー: 部門="' + departmentName + '", type="' + reportType + '" の定義が見つかりません。');
    Logger.log('利用可能な組み合わせ:');
    REPORT_DEFINITIONS.forEach(function(d) {
      Logger.log('  - ' + d.name + ' / ' + d.type + ' → ' + buildFileName(d.name, d.type, fiscalYear));
    });
    return;
  }

  Logger.log('=== 単体同期: ' + buildFileName(def.name, def.type, fiscalYear) + ' ===');
  try {
    var result = syncOneReport(def, fiscalYear);
    Logger.log(result);
  } catch (e) {
    Logger.log(def.name + ': エラー - ' + e.message);
  }
}

/**
 * REPORT_DEFINITIONS の全有効レポートを同期する内部関数
 * @param {number} fiscalYear - 決算年
 */
function syncReportsByFiscalYear(fiscalYear) {
  var range = getFiscalYearRange(fiscalYear);
  Logger.log('=== MoneyForward → Google Drive 年間同期 ===');
  Logger.log('決算年: ' + fiscalYear + ' （対象期間: ' + range.startDate + ' 〜 ' + range.endDate + '）');

  var enabledReports = REPORT_DEFINITIONS.filter(function(d) { return d.enabled; });
  Logger.log('対象レポート数: ' + enabledReports.length + '件');

  var results = [];
  enabledReports.forEach(function(def) {
    try {
      var result = syncOneReport(def, fiscalYear);
      results.push(result);
    } catch (e) {
      results.push(def.name + ': エラー - ' + e.message);
    }
  });

  Logger.log('=== 実行結果 ===');
  results.forEach(function(r) { Logger.log(r); });
  Logger.log('=== 同期完了 ===');

  sendNotification(fiscalYear, results);
}

// ============================================================
// 通知
// ============================================================

/**
 * 同期結果をメールで通知
 */
function sendNotification(fiscalYear, results) {
  var props = PropertiesService.getScriptProperties();
  var email = props.getProperty('NOTIFICATION_EMAIL');
  if (!email) return;

  var range = getFiscalYearRange(fiscalYear);
  var subject = '【MF→Drive同期】決算年' + fiscalYear + ' レポート更新完了';
  var body = 'MoneyForward → Google Drive 同期結果\n\n';
  body += '決算年: ' + fiscalYear + '（' + range.startDate + ' 〜 ' + range.endDate + '）\n';
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
    if (trigger.getHandlerFunction() === 'syncReports') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 新規トリガー作成：毎月1日 9:00-10:00
  ScriptApp.newTrigger('syncReports')
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
    if (trigger.getHandlerFunction() === 'syncReports') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('syncReports')
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
