/**
 * キャッシュフロー管理システム - 設定ファイル
 * 株式会社LEVEL1 資金繰り管理
 *
 * ※ SPREADSHEET_ID を実際のスプレッドシートIDに変更してください
 */

const CF_CONFIG = {

  // ==============================
  // スプレッドシート設定
  // ==============================
  SPREADSHEET_ID: '', // 空の場合はアクティブなスプレッドシートを使用

  SHEETS: {
    DAILY:       'Daily',       // 日次入出金記録
    CF005:       'CF005',       // PayPay 005 月別集計
    CF003:       'CF003',       // PayPay 003 月別集計
    SEIBU:       '西武信金',     // 西武信用金庫 月別集計
    MONTHLY:     '月別',        // 3口座合算サマリー
    SETTINGS:    '設定',        // カテゴリマスタ・設定
    CURRENT_BAL: '現残高'       // 各口座の現在残高
  },

  // ==============================
  // 口座定義
  // ==============================
  ACCOUNTS: {
    CF005: {
      name: 'PayPay銀行 ビジネス営業部',
      shortName: 'PayPay 005',
      sheetName: 'CF005',
      // Dailyシートでの列配置（1始まり）
      daily: {
        DATE:    2,  // B: 日付
        CONTENT: 4,  // D: 内容
        DEPOSIT: 5,  // E: 入金
        WITHDRAWAL: 6, // F: 出金
        BALANCE: 7,  // G: 残高
        SOURCE:  8   // H: データソース（CSV/手入力/予定）
      }
    },
    CF003: {
      name: 'PayPay銀行 はやぶさ支店',
      shortName: 'PayPay 003',
      sheetName: 'CF003',
      daily: {
        DATE:    10, // J: 日付
        CONTENT: 11, // K: 内容
        DEPOSIT: 12, // L: 入金
        WITHDRAWAL: 13, // M: 出金
        BALANCE: 14, // N: 残高
        SOURCE:  15  // O: データソース
      }
    },
    SEIBU: {
      name: '西武信用金庫 阿佐ヶ谷支店',
      shortName: '西武信金',
      sheetName: '西武信金',
      daily: {
        DATE:    17, // Q: 日付
        CONTENT: 18, // R: 内容
        DEPOSIT: 19, // S: 入金
        WITHDRAWAL: 20, // T: 出金
        BALANCE: 21, // U: 残高
        SOURCE:  22  // V: データソース
      }
    }
  },

  // Dailyシートの合計列（3口座合算残高）
  DAILY_TOTAL_COL: 1,  // A列: 合計残高
  // Dailyシートのヘッダー行数
  DAILY_HEADER_ROWS: 1,

  // ==============================
  // 銀行CSVフォーマット定義
  // ==============================
  CSV_FORMATS: {
    // PayPay銀行（ビジネス / はやぶさ共通）
    PAYPAY: {
      encoding: 'Shift_JIS',
      skipRows: 0,          // ヘッダー行のスキップ数（0=ヘッダーなし or 1行目がヘッダー）
      hasHeader: true,
      delimiter: ',',
      columns: {
        date: 0,            // 日付
        content: 1,         // 摘要
        deposit: 2,         // お預り金額（入金）
        withdrawal: 3,      // お支払金額（出金）
        balance: 4           // 残高
      },
      dateFormat: 'yyyy/MM/dd'
    },
    // 西武信用金庫
    SEIBU: {
      encoding: 'Shift_JIS',
      skipRows: 0,
      hasHeader: true,
      delimiter: ',',
      columns: {
        date: 0,
        content: 1,
        deposit: 2,
        withdrawal: 3,
        balance: 4
      },
      dateFormat: 'yyyy/MM/dd'
    }
  },

  // 口座とCSVフォーマットの対応
  ACCOUNT_CSV_MAP: {
    CF005: 'PAYPAY',
    CF003: 'PAYPAY',
    SEIBU: 'SEIBU'
  },

  // ==============================
  // 月次集計の入金カテゴリ
  // ==============================
  INCOME_CATEGORIES: [
    { key: 'amazon',     label: 'Amazon',     keywords: ['Amazon売上', 'AMAZON', 'アマゾン'] },
    { key: 'crowdfund',  label: 'クラファン',  keywords: ['Makuake', 'クラウドファンディング', 'クラファン'] },
    { key: 'ec',         label: 'ECサイト',    keywords: ['ECサイト', 'Shopify', 'BASE'] },
    { key: 'wholesale',  label: '卸',          keywords: ['卸', '卸売'] },
    { key: 'minpaku',    label: '民泊',        keywords: ['民泊', 'Airbnb', 'Booking', 'Marriott'] },
    { key: 'transfer',   label: '銀行移動',    keywords: ['振替', '資金移動', '口座間'] },
    { key: 'interest',   label: '受取利息',    keywords: ['受取利息', '利息'] },
    { key: 'other',      label: 'その他',      keywords: [] }
  ],

  // ==============================
  // 月次集計の出金カテゴリ
  // ==============================
  EXPENSE_CATEGORIES: [
    { key: 'saison',     label: 'セゾンプラチナ', keywords: ['セゾン', 'SAISON'] },
    { key: 'rakuten',    label: '楽天ビジネス',   keywords: ['楽天', 'RAKUTEN'] },
    { key: 'upsider',   label: 'UPSIDER',        keywords: ['UPSIDER'] },
    { key: 'lc',         label: 'LC',              keywords: ['LC'] },
    { key: 'jal',        label: 'JALカード',       keywords: ['JAL', 'ジャル'] },
    { key: 'purchase',   label: '仕入',            keywords: ['仕入', '仕入れ'] },
    { key: 'customs',    label: '関税・消費税',    keywords: ['関税', '消費税'] },
    { key: 'shipping',   label: '配送料',          keywords: ['配送', 'ヤマト', '佐川', '日本郵便'] },
    { key: 'expense',    label: '支払経費',        keywords: ['経費', '支払'] },
    { key: 'repayment',  label: '返済',            keywords: ['返済', '融資返済'] },
    { key: 'salary',     label: '役員報酬・手当',  keywords: ['役員報酬', '給与', '手当'] },
    { key: 'insurance',  label: '社会保険料',      keywords: ['社会保険', '厚生'] },
    { key: 'realestate', label: '不動産取引',      keywords: ['不動産', '家賃'] },
    { key: 'fee',        label: '振込手数料',      keywords: ['手数料', '振込手数料'] },
    { key: 'transfer',   label: '資金移動',        keywords: ['振替', '資金移動', '口座間'] },
    { key: 'tax',        label: '税理士',          keywords: ['税理士', '風間'] },
    { key: 'telecom',    label: '通信費',          keywords: ['ドコモ', 'docomo', '通信'] },
    { key: 'other',      label: 'その他',          keywords: [] }
  ],

  // ==============================
  // データソースラベル
  // ==============================
  SOURCE: {
    CSV:      'CSV',      // 銀行CSVから取込
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
