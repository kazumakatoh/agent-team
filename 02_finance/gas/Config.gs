/**
 * キャッシュフロー管理システム - 設定ファイル
 * 株式会社LEVEL1 資金繰り管理
 *
 * ※ MF API の Client ID / Client Secret はコードに書かないこと！
 *   スプレッドシートの「設定」シートに入力するか、
 *   GASのスクリプトプロパティに設定してください。
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
  // マネーフォワードクラウド会計 API設定
  // ==============================
  MF_API: {
    // Client ID / Secret はスクリプトプロパティから読み込む（getMfCredentials_()で取得）
    BASE_URL:      'https://api.biz.moneyforward.com/api/v3',
    AUTH_URL:       'https://api.biz.moneyforward.com/authorize',
    TOKEN_URL:      'https://api.biz.moneyforward.com/token',
    SCOPE:          'mfc/accounting/offices.read mfc/accounting/journal.read mfc/accounting/accounts.read mfc/accounting/report.read mfc/accounting/connected_account.read'
  },

  // ==============================
  // 口座定義
  // ==============================
  ACCOUNTS: {
    CF005: {
      name: 'PayPay銀行 ビジネス営業部',
      shortName: 'PayPay 005',
      sheetName: 'CF005',
      walletId: '',  // MFのwallet ID（初回連携時に設定シートから取得）
      // Dailyシートでの列配置（1始まり）
      daily: {
        DATE:       2,  // B: 日付
        CONTENT:    4,  // D: 内容
        DEPOSIT:    5,  // E: 入金
        WITHDRAWAL: 6,  // F: 出金
        BALANCE:    7,  // G: 残高
        SOURCE:     8   // H: データソース（MF/手入力/予定）
      }
    },
    CF003: {
      name: 'PayPay銀行 はやぶさ支店',
      shortName: 'PayPay 003',
      sheetName: 'CF003',
      walletId: '',
      daily: {
        DATE:       10, // J: 日付
        CONTENT:    11, // K: 内容
        DEPOSIT:    12, // L: 入金
        WITHDRAWAL: 13, // M: 出金
        BALANCE:    14, // N: 残高
        SOURCE:     15  // O: データソース
      }
    },
    SEIBU: {
      name: '西武信用金庫 阿佐ヶ谷支店',
      shortName: '西武信金',
      sheetName: '西武信金',
      walletId: '',
      daily: {
        DATE:       17, // Q: 日付
        CONTENT:    18, // R: 内容
        DEPOSIT:    19, // S: 入金
        WITHDRAWAL: 20, // T: 出金
        BALANCE:    21, // U: 残高
        SOURCE:     22  // V: データソース
      }
    }
  },

  // Dailyシートの合計列（3口座合算残高）
  DAILY_TOTAL_COL: 1,  // A列: 合計残高
  // Dailyシートのヘッダー行数
  DAILY_HEADER_ROWS: 1,

  // ==============================
  // 月次集計の入金カテゴリ
  // MFの勘定科目でも自動分類するが、摘要キーワードも併用
  // ==============================
  INCOME_CATEGORIES: [
    { key: 'amazon',     label: 'Amazon',     keywords: ['Amazon売上', 'AMAZON', 'アマゾン'],              accountItems: ['売上高'] },
    { key: 'crowdfund',  label: 'クラファン',  keywords: ['Makuake', 'クラウドファンディング', 'クラファン'], accountItems: [] },
    { key: 'ec',         label: 'ECサイト',    keywords: ['ECサイト', 'Shopify', 'BASE'],                   accountItems: [] },
    { key: 'wholesale',  label: '卸',          keywords: ['卸', '卸売'],                                    accountItems: [] },
    { key: 'minpaku',    label: '民泊',        keywords: ['民泊', 'Airbnb', 'Booking', 'Marriott'],         accountItems: [] },
    { key: 'transfer',   label: '銀行移動',    keywords: ['振替', '資金移動', '口座間'],                     accountItems: [] },
    { key: 'interest',   label: '受取利息',    keywords: ['受取利息', '利息'],                               accountItems: ['受取利息'] },
    { key: 'other',      label: 'その他',      keywords: [],                                                accountItems: [] }
  ],

  // ==============================
  // 月次集計の出金カテゴリ
  // ==============================
  EXPENSE_CATEGORIES: [
    { key: 'saison',     label: 'セゾンプラチナ', keywords: ['セゾン', 'SAISON'],                   accountItems: [] },
    { key: 'rakuten',    label: '楽天ビジネス',   keywords: ['楽天', 'RAKUTEN'],                    accountItems: [] },
    { key: 'upsider',    label: 'UPSIDER',        keywords: ['UPSIDER'],                           accountItems: [] },
    { key: 'lc',         label: 'LC',              keywords: ['LC'],                                accountItems: [] },
    { key: 'jal',        label: 'JALカード',       keywords: ['JAL', 'ジャル'],                     accountItems: [] },
    { key: 'purchase',   label: '仕入',            keywords: ['仕入', '仕入れ'],                    accountItems: ['仕入高'] },
    { key: 'customs',    label: '関税・消費税',    keywords: ['関税', '消費税'],                     accountItems: ['関税', '租税公課'] },
    { key: 'shipping',   label: '配送料',          keywords: ['配送', 'ヤマト', '佐川', '日本郵便'], accountItems: ['荷造運賃'] },
    { key: 'expense',    label: '支払経費',        keywords: ['経費', '支払'],                       accountItems: [] },
    { key: 'repayment',  label: '返済',            keywords: ['返済', '融資返済'],                   accountItems: ['長期借入金'] },
    { key: 'salary',     label: '役員報酬・手当',  keywords: ['役員報酬', '給与', '手当'],           accountItems: ['役員報酬'] },
    { key: 'insurance',  label: '社会保険料',      keywords: ['社会保険', '厚生'],                   accountItems: ['法定福利費'] },
    { key: 'realestate', label: '不動産取引',      keywords: ['不動産', '家賃'],                     accountItems: ['地代家賃'] },
    { key: 'fee',        label: '振込手数料',      keywords: ['手数料', '振込手数料'],               accountItems: ['支払手数料'] },
    { key: 'transfer',   label: '資金移動',        keywords: ['振替', '資金移動', '口座間'],         accountItems: [] },
    { key: 'tax',        label: '税理士',          keywords: ['税理士', '風間'],                     accountItems: ['支払報酬'] },
    { key: 'telecom',    label: '通信費',          keywords: ['ドコモ', 'docomo', '通信'],           accountItems: ['通信費'] },
    { key: 'other',      label: 'その他',          keywords: [],                                    accountItems: [] }
  ],

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
