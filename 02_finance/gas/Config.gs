/**
 * キャッシュフロー管理システム - 設定ファイル
 * 株式会社LEVEL1 資金繰り管理
 *
 * ※ MF API の Client ID / Client Secret はコードに書かないこと！
 *   GASのスクリプトプロパティに設定してください。
 */

const CF_CONFIG = {

  // ==============================
  // スプレッドシート設定
  // ==============================
  SPREADSHEET_ID: '', // 空の場合はアクティブなスプレッドシートを使用

  SHEETS: {
    MONTHLY:     '月別',        // 月次集計（3口座集約）
    DAILY_SUMMARY: '日別サマリー', // 日別の3口座合計
    INVENTORY:   '在庫残高',     // 在庫残高シート
    SETTINGS:    '設定',        // カテゴリマスタ・設定
    CURRENT_BAL: '現残高'       // 各口座の現在残高
  },

  // ==============================
  // マネーフォワードクラウド会計 API設定
  // ==============================
  MF_API: {
    BASE_URL:      'https://api-accounting.moneyforward.com/api/v3',
    AUTH_URL:       'https://api.biz.moneyforward.com/authorize',
    TOKEN_URL:      'https://api.biz.moneyforward.com/token',
    SCOPE:          'mfc/accounting/offices.read mfc/accounting/journal.read mfc/accounting/accounts.read mfc/accounting/report.read mfc/accounting/connected_account.read'
  },

  // ==============================
  // 口座定義
  // 各口座に専用のDailyシートを持つ
  // 列構成: A:日付 / B:内容 / C:入金 / D:出金 / E:残高 / F:ソース
  // ==============================
  ACCOUNTS: {
    CF005: {
      name: 'PayPay銀行 ビジネス営業部',
      shortName: 'PayPay 005',
      dailySheet: 'Daily_005',   // 専用Dailyシート名
      walletId: ''
    },
    CF003: {
      name: 'PayPay銀行 はやぶさ支店',
      shortName: 'PayPay 003',
      dailySheet: 'Daily_003',
      walletId: ''
    },
    SEIBU: {
      name: '西武信用金庫 阿佐ヶ谷支店',
      shortName: '西武信金',
      dailySheet: 'Daily_西武',
      walletId: ''
    }
  },

  // 各Dailyシート共通の列定義（1始まり）
  DAILY_COLS: {
    DATE:       1,  // A: 日付
    CONTENT:    2,  // B: 内容
    DEPOSIT:    3,  // C: 入金
    WITHDRAWAL: 4,  // D: 出金
    BALANCE:    5,  // E: 残高
    SOURCE:     6   // F: ソース（MF/手入力/予定）
  },

  // Dailyシートのヘッダー行数
  DAILY_HEADER_ROWS: 1,

  // ==============================
  // 発注管理表（別スプレッドシート）
  // ==============================
  ORDER_MGMT: {
    SPREADSHEET_ID: '1S7LjgclM7teGzKay0usBDonH93NjszSsKtBONBBNzas',
    SHEET_NAME: '在庫一覧',
    COLS: {
      PRODUCT:   2,  // B: 商品名
      ASIN:      3,  // C: ASIN
      STOCK:     4,  // D: 在庫数（手入力）
      FBA_STOCK: 6   // F: FBA在庫（在庫残高に反映する列）
    }
  },

  // ==============================
  // データソースラベル
  // ==============================
  SOURCE: {
    MF:       'MF',       // マネーフォワードAPIから取込
    MANUAL:   '手入力',   // 実績の手入力
    PLANNED:  '予定'      // 未来の入出金予定
  },

  // ==============================
  // アラート基準（PayPay 005口座ベース）
  // ==============================
  ALERT: {
    DANGER_THRESHOLD:  5000000,   // 500万円以下 = 🔴 危険
    WARNING_THRESHOLD: 10000000,  // 1,000万円以下 = 🟡 注意
    ALERT_ACCOUNT: 'CF005'       // 監視対象口座
  },

  // ==============================
  // 表示設定
  // ==============================
  DISPLAY: {
    DATE_FORMAT: 'yyyy/MM/dd',
    CURRENCY_FORMAT: '#,##0',
    TIMEZONE: 'Asia/Tokyo'
  }
};
