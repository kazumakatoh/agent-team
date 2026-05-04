/**
 * Amazon Dashboard - 設定・定数
 *
 * 認証情報は PropertiesService で管理（コードに直書きしない）
 * 初回セットアップ時に setupCredentials() を1回実行する
 */

// ===== マーケットプレイス設定 =====
const MARKETPLACE_ID_JP = 'A1VC38T7YXB528';
const SP_API_ENDPOINT = 'https://sellingpartnerapi-fe.amazon.com';
const ADS_API_ENDPOINT = 'https://advertising-api-fe.amazon.com';
const LWA_TOKEN_URL = 'https://api.amazon.com/auth/o2/token';

// ===== スプレッドシート設定 =====
// シートIDは PropertiesService から取得
function getMainSpreadsheetId() {
  return PropertiesService.getScriptProperties().getProperty('MAIN_SHEET_ID');
}

function getCfSpreadsheetId() {
  return PropertiesService.getScriptProperties().getProperty('CF_SHEET_ID');
}

function getIntermediateSheetId() {
  return PropertiesService.getScriptProperties().getProperty('INTERMEDIATE_SHEET_ID');
}

// ===== シート名定数 =====
const SHEET_NAMES = {
  // 見るシート（ダッシュボード）
  // 旧 事業ダッシュボード（L1）/ カテゴリ分析（L2）/ アカウント健全性（D5）は廃止。
  // カテゴリ別月次（CategoryMonthly.gs）に統合。
  L3_PRODUCT: '商品分析',
  // データシート
  D1_DAILY: '日次データ',
  D1S_DAILY_SALES: '日次販売実績',  // D1から日次集計（売上/CV/点数/経費比例配分/利益）
  D2_SETTLEMENT: '経費明細',
  D2S_SETTLEMENT_SUMMARY: '経費月次集計',  // ASIN×月の事前集計（高速化用）
  D2F_FINANCE_EVENTS: '日次フィー（Finance）',  // Finance APIから日次でフィー取得
  // 旧 広告詳細シート（D3_*）は廃止。Ads データは D1（日次データ）の広告列に集約され、
  // 集計はカテゴリ別月次シート（CategoryMonthly.gs）で行う。
  // マスターシート
  M1_PRODUCT_MASTER: '商品マスター',
  M2_PURCHASE_PRICE: '月次仕入単価',
  M3_PROMO_COST: '販促費マスター',
  // インポートシート
  HISTORICAL_IMPORT: 'インポート_履歴データ',  // 過去Excelデータの一括取り込み用
};

// ===== 中間スプシ（就業管理表）の構造 =====
// 「管理表」シートでの行マッピング（年×項目テーブル想定、列C=1月〜N=12月）
const INTERMEDIATE_MGMT_SHEET = '管理表';
const INTERMEDIATE_ROWS = {
  SALES: 2,             // 販売数
  PAY_REWARD: 5,        // 支払報酬（→ M3 納品人件費）
  FBA_SHIPPING: 6,      // FBA輸送手数料（参考表示・GAS自動更新）
  YAMATO: 7,            // ヤマト運輸（→ M3 荷造運賃に合算）
};
const INTERMEDIATE_MONTH_COL_START = 3;  // C列 = 1月

// ===== 対象商品の定義 =====
const ACTIVE_PRODUCT_DAYS = 60; // 直近60日以内に注文がある商品を対象

// ===== 認証情報の取得ヘルパー =====
function getCredential(key) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) {
    throw new Error('認証情報が未設定です: ' + key + '。setupCredentials() を実行してください。');
  }
  return value;
}

/**
 * 初回セットアップ: 認証情報を PropertiesService に保存
 * GASエディタから1回だけ手動実行する
 *
 * 実行前に以下の値を入力してください（実行後はこの関数からは値を削除してOK）
 */
function setupCredentials() {
  const props = PropertiesService.getScriptProperties();

  // ===== ここに値を入力して1回実行 → その後値を削除 =====
  props.setProperties({
    // SP-API
    'SP_SELLER_ID': '',          // 出品者ID
    'SP_CLIENT_ID': '',          // LWA Client ID
    'SP_CLIENT_SECRET': '',      // LWA Client Secret
    'SP_REFRESH_TOKEN': '',      // SP-API用 Refresh Token

    // Amazon Ads API
    'ADS_CLIENT_ID': '',         // Ads API Client ID
    'ADS_CLIENT_SECRET': '',     // Ads API Client Secret
    'ADS_REFRESH_TOKEN': '',     // Ads API用 Refresh Token（新規取得済み）
    'ADS_PROFILE_ID': '',        // 数値Profile ID（サポート回答後に設定）

    // スプレッドシート
    'MAIN_SHEET_ID': '',         // メインダッシュボードのスプレッドシートID
    'CF_SHEET_ID': '',           // キャッシュフロー管理シートのID
    'INTERMEDIATE_SHEET_ID': '', // 就業管理表のスプレッドシートID（中間スプシ）

    // Claude API
    'CLAUDE_API_KEY': '',        // Anthropic API Key

    // 通知
    'GMAIL_TO': '',              // 改善提案の送信先メールアドレス
    'LINE_CHANNEL_TOKEN': '',    // LINE Messaging API トークン
    'LINE_USER_ID': '',          // LINE ユーザーID
  });

  Logger.log('認証情報を保存しました。セキュリティのため、この関数内の値を空文字に戻してください。');
}

/**
 * 認証情報の設定状態を確認（値は表示しない）
 */
function checkCredentials() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const keys = [
    'SP_SELLER_ID', 'SP_CLIENT_ID', 'SP_CLIENT_SECRET', 'SP_REFRESH_TOKEN',
    'ADS_CLIENT_ID', 'ADS_CLIENT_SECRET', 'ADS_REFRESH_TOKEN', 'ADS_PROFILE_ID',
    'MAIN_SHEET_ID', 'CF_SHEET_ID', 'INTERMEDIATE_SHEET_ID',
    'CLAUDE_API_KEY', 'GMAIL_TO', 'LINE_CHANNEL_TOKEN', 'LINE_USER_ID'
  ];

  keys.forEach(key => {
    const status = props[key] ? '✅ 設定済み' : '❌ 未設定';
    Logger.log(key + ': ' + status);
  });
}
/**
 * 日次トリガーを設定
 * 1回だけ実行すればOK
 */
function setupDailyTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));

  // 毎日 6:00 - 注文データ取得
  ScriptApp.newTrigger('dailyFetchByReport')
    .timeBased().everyDays(1).atHour(6).create();

  // 毎日 6:15 - トラフィックデータ取得
  ScriptApp.newTrigger('dailyFetchTraffic')
    .timeBased().everyDays(1).atHour(6).nearMinute(15).create();

  // 毎日 6:30 - 商品マスター同期
  ScriptApp.newTrigger('syncMasterToDaily')
    .timeBased().everyDays(1).atHour(6).nearMinute(30).create();

  // 毎日 7:00 - Settlement Report
  ScriptApp.newTrigger('fetchSettlementReports')
    .timeBased().everyDays(1).atHour(7).create();

  // 毎日 7:30 - Finance API（日次フィー取得）
  ScriptApp.newTrigger('fetchDailyFinanceEvents')
    .timeBased().everyDays(1).atHour(7).nearMinute(30).create();

  // 毎日 8:00 - 中間スプシ ↔ M3 同期
  ScriptApp.newTrigger('syncIntermediateAndM3')
    .timeBased().everyDays(1).atHour(8).create();

  // 毎日 8:30 - 日次販売実績シート構築
  ScriptApp.newTrigger('buildDailySalesSheet')
    .timeBased().everyDays(1).atHour(8).nearMinute(30).create();

  // 毎日 9:00 - LINE 緊急アラート
  ScriptApp.newTrigger('runDailyAlerts')
    .timeBased().everyDays(1).atHour(9).create();

  // 毎日 9:30 - 競合価格取得（Phase 4b）
  ScriptApp.newTrigger('fetchCompetitorPricing')
    .timeBased().everyDays(1).atHour(9).nearMinute(30).create();

  // 毎日 11:00 - Amazon Ads レポート取得（前日分 × 3種）
  ScriptApp.newTrigger('dailyFetchAdsReports')
    .timeBased().everyDays(1).atHour(11).create();

  // 毎週月曜 7:30 - セールカレンダー点検
  ScriptApp.newTrigger('checkUpcomingSales')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).nearMinute(30).create();

  // 毎週月曜 8:00 - 週次AIレポート（Claude → Gmail）
  ScriptApp.newTrigger('sendWeeklyAiReport')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();

  // 毎日 10:00 - 在庫取得＋在庫切れアラート
  ScriptApp.newTrigger('fetchInventoryAndAlert')
    .timeBased().everyDays(1).atHour(10).create();

  // 毎日 10:30 - 発注管理表 / CF管理スプシへ在庫同期
  ScriptApp.newTrigger('syncInventoryToExternalSheets')
    .timeBased().everyDays(1).atHour(10).nearMinute(30).create();

  // 毎月1日 9:00 - 月次AI戦略レポート（Claude Opus → Gmail）
  ScriptApp.newTrigger('sendMonthlyAiReport')
    .timeBased().onMonthDay(1).atHour(9).create();

  // 毎月3日 6:00 - CFシート → M2 仕入単価同期
  ScriptApp.newTrigger('syncPurchasePriceFromCfSheet')
    .timeBased().onMonthDay(3).atHour(6).create();

  Logger.log('✅ トリガー設定完了');
  Logger.log('  毎日 6:00 - 注文データ');
  Logger.log('  毎日 6:15 - トラフィック');
  Logger.log('  毎日 6:30 - マスター同期');
  Logger.log('  毎日 7:00 - Settlement');
  Logger.log('  毎日 7:30 - Finance Events');
  Logger.log('  毎日 8:00 - 中間スプシ↔M3 同期');
  Logger.log('  毎日 8:30 - 日次販売実績シート');
  Logger.log('  毎日 9:00 - LINE 緊急アラート');
  Logger.log('  毎日 9:30 - 競合価格');
  Logger.log('  毎日 11:00 - Amazon Ads レポート');
  Logger.log('  毎日 10:00 - 在庫＋在庫切れアラート');
  Logger.log('  毎日 10:30 - 発注管理表/CF管理スプシへ在庫同期');
  Logger.log('  毎週月 7:30 - セール準備チェック');
  Logger.log('  毎週月 8:00 - 週次AIレポート');
  Logger.log('  毎月1日 9:00 - 月次AI戦略レポート');
  Logger.log('  毎月3日 6:00 - CF→M2 仕入単価同期');
}

/**
 * Phase 4 のトリガー5本だけを既存を保持したまま追加する
 * setupDailyTriggers() は全削除→再登録なので、既存トリガーの実行履歴を残したい場合はこちらを使う。
 *
 * 実行前に checkCredentials() で以下が設定済みか確認:
 *   - CLAUDE_API_KEY    (sendWeeklyAiReport, checkUpcomingSales)
 *   - GMAIL_TO          (sendWeeklyAiReport, checkUpcomingSales)
 *   - LINE_CHANNEL_TOKEN, LINE_USER_ID  (runDailyAlerts)
 */
function addPhase4Triggers() {
  const existing = ScriptApp.getProjectTriggers();
  const existingFns = new Set(existing.map(t => t.getHandlerFunction()));

  const specs = [
    { fn: 'runDailyAlerts',         desc: '毎日 9:00 - LINE緊急アラート',
      create: () => ScriptApp.newTrigger('runDailyAlerts').timeBased().everyDays(1).atHour(9).create() },
    { fn: 'fetchCompetitorPricing', desc: '毎日 9:30 - 競合価格',
      create: () => ScriptApp.newTrigger('fetchCompetitorPricing').timeBased().everyDays(1).atHour(9).nearMinute(30).create() },
    { fn: 'checkUpcomingSales',     desc: '毎週月 7:30 - セール準備チェック',
      create: () => ScriptApp.newTrigger('checkUpcomingSales').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).nearMinute(30).create() },
    { fn: 'sendWeeklyAiReport',     desc: '毎週月 8:00 - 週次AIレポート',
      create: () => ScriptApp.newTrigger('sendWeeklyAiReport').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create() },
  ];

  let added = 0;
  for (const s of specs) {
    if (existingFns.has(s.fn)) {
      Logger.log('⏭  既に登録済みスキップ: ' + s.fn);
      continue;
    }
    s.create();
    Logger.log('✅ 追加: ' + s.desc);
    added++;
  }
  Logger.log('完了: ' + added + '本追加（既存は保持）');
}

/**
 * Phase 5 のトリガーを既存を保持したまま追加する
 * - fetchInventoryAndAlert     毎日 10:00
 * - sendMonthlyAiReport        毎月1日 9:00
 */
function addPhase5Triggers() {
  const existing = ScriptApp.getProjectTriggers();
  const existingFns = new Set(existing.map(t => t.getHandlerFunction()));

  const specs = [
    { fn: 'fetchInventoryAndAlert',       desc: '毎日 10:00 - 在庫取得＋在庫切れアラート',
      create: () => ScriptApp.newTrigger('fetchInventoryAndAlert').timeBased().everyDays(1).atHour(10).create() },
    { fn: 'syncInventoryToExternalSheets', desc: '毎日 10:30 - 発注管理表/CF管理への在庫同期',
      create: () => ScriptApp.newTrigger('syncInventoryToExternalSheets').timeBased().everyDays(1).atHour(10).nearMinute(30).create() },
    { fn: 'sendMonthlyAiReport',           desc: '毎月1日 9:00 - 月次AI戦略レポート',
      create: () => ScriptApp.newTrigger('sendMonthlyAiReport').timeBased().onMonthDay(1).atHour(9).create() },
  ];

  let added = 0;
  for (const s of specs) {
    if (existingFns.has(s.fn)) {
      Logger.log('⏭  既に登録済みスキップ: ' + s.fn);
      continue;
    }
    s.create();
    Logger.log('✅ 追加: ' + s.desc);
    added++;
  }
  Logger.log('完了: ' + added + '本追加（既存は保持）');
}

/**
 * Phase 3 （Amazon Ads）のトリガーを既存を保持したまま追加する
 * - dailyFetchAdsReports   毎日 11:00
 *
 * 実行前に testAdsProfiles() で ADS_PROFILE_ID が保存済みか確認してください。
 */
function addPhase3Triggers() {
  const existing = ScriptApp.getProjectTriggers();
  const existingFns = new Set(existing.map(t => t.getHandlerFunction()));

  const specs = [
    { fn: 'dailyFetchAdsReports', desc: '毎日 11:00 - Amazon Adsレポート（前日分）',
      create: () => ScriptApp.newTrigger('dailyFetchAdsReports').timeBased().everyDays(1).atHour(11).create() },
  ];

  let added = 0;
  for (const s of specs) {
    if (existingFns.has(s.fn)) {
      Logger.log('⏭  既に登録済みスキップ: ' + s.fn);
      continue;
    }
    s.create();
    Logger.log('✅ 追加: ' + s.desc);
    added++;
  }
  Logger.log('完了: ' + added + '本追加（既存は保持）');
}

/**
 * Phase 4 で新規追加したシート（D4 競合価格 / M4 セールカレンダー）を一括初期化
 * 1回だけ実行すればOK
 */
function setupPhase4Sheets() {
  setupCompetitorSheet();        // D4 競合価格
  setupSaleCalendar();           // M4 セールカレンダー（プライムデー等を自動投入）
  Logger.log('✅ Phase 4 シート初期化完了');
}

