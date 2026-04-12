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

// ===== シート名定数 =====
const SHEET_NAMES = {
  // 見るシート（ダッシュボード）
  L1_DASHBOARD: '事業ダッシュボード',
  L2_CATEGORY: 'カテゴリ分析',
  L3_PRODUCT: '商品分析',
  // データシート
  D1_DAILY: '日次データ',
  D2_SETTLEMENT: '経費明細',
  D3_ADS_DETAIL: '広告詳細',
  // マスターシート
  M1_PRODUCT_MASTER: '商品マスター',
  M2_PURCHASE_PRICE: '月次仕入単価',
  M3_PROMO_COST: '販促費マスター',
};

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
    'MAIN_SHEET_ID', 'CF_SHEET_ID',
    'CLAUDE_API_KEY', 'GMAIL_TO', 'LINE_CHANNEL_TOKEN', 'LINE_USER_ID'
  ];

  keys.forEach(key => {
    const status = props[key] ? '✅ 設定済み' : '❌ 未設定';
    Logger.log(key + ': ' + status);
  });
}
