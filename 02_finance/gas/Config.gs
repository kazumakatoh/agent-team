/**
 * 財務レポート自動化システム - 設定ファイル
 * マネーフォワード クラウド会計 API連携
 *
 * ⚠️ 初期設定が必要な箇所は「// ★要設定」コメントで明示しています
 */

const CONFIG = {

  // ==============================
  // MF会計 API設定
  // ※ OAuth認証は developers.biz.moneyforward.com で管理
  // ※ 会計API本体は api-accounting.moneyforward.com
  // ==============================
  MF_API: {
    // 会計API本体のベースURL（試算表・部門・勘定科目などのエンドポイント）
    BASE_URL:      'https://api-accounting.moneyforward.com',

    // OAuth2認証エンドポイント（MFビジネスプラットフォーム共通）
    // 出典: github.com/moneyforward/api-doc
    AUTH_URL:      'https://moneyforward.com/oauth/authorize',
    TOKEN_URL:     'https://moneyforward.com/oauth/v2/token',

    CLIENT_ID:     '', // ★要設定: MF会計アプリのクライアントID（STEP4で取得）
    CLIENT_SECRET: '', // ★要設定: MF会計アプリのクライアントシークレット（STEP4で取得）

    // ★要設定: GAS Web AppのURL（STEP2でデプロイ後に取得）
    // 例: 'https://script.google.com/macros/s/AKfycbxxxxxx/exec'
    REDIRECT_URI:  '',

    // 必要スコープ（スペース区切りで複数指定）
    // mfc/accounting/offices.read     : 事業所情報の読み取り
    // mfc/accounting/accounts.read    : 勘定科目・補助科目の読み取り
    // mfc/accounting/departments.read : 部門情報の読み取り
    // mfc/accounting/journal.read     : 試算表・仕訳の読み取り
    SCOPE: 'mfc/accounting/offices.read mfc/accounting/accounts.read mfc/accounting/departments.read mfc/accounting/journal.read',

    // 部門フィルタのパラメータ名（確認済み）
    SEGMENT_PARAM: 'department_id',
  },

  // ==============================
  // スプレッドシート設定
  // ==============================
  SPREADSHEET_ID: '', // ★要設定: 出力先スプレッドシートのID

  // ==============================
  // 事業年度設定
  // ★ 事業年度開始月（LEVEL1は3月始まり）
  // ==============================
  FISCAL: {
    START_MONTH: 3,  // 3月始まり（3〜翌2月）
  },

  // ==============================
  // 部門設定
  // ★ MF会計の部門IDと部門名の対応
  //   MF会計の「設定 > 部門」で確認してください
  // ==============================
  DEPARTMENTS: [
    { id: '',  name: '共通',   shortName: '共通'   }, // ★要設定: 部門ID
    { id: '',  name: '物販',   shortName: '物販'   }, // ★要設定: 部門ID
    { id: '',  name: 'ブランド', shortName: 'ブランド' }, // ★要設定: 部門ID
    { id: '',  name: '民泊',   shortName: '民泊'   }, // ★要設定: 部門ID
  ],

  // ==============================
  // PL構造定義
  // MF会計の勘定科目名 → PLカテゴリのマッピング
  // ★ 自社のMF会計の勘定科目名に合わせて調整してください
  // ==============================
  PL_STRUCTURE: [
    // ── 売上高 ──────────────────────────────
    {
      category:   'header',
      label:      '売上高',
      indent:     0,
      isTotal:    false,
      isBold:     true,
    },
    {
      category:   'revenue',
      label:      '売上（国内）',
      indent:     1,
      accountNames: ['売上高', '売上（国内）', '売上-国内'],
      sign:       1,   // 1=プラス, -1=マイナス
    },
    {
      category:   'revenue',
      label:      '売上（海外）',
      indent:     1,
      accountNames: ['売上（海外）', '売上-海外', '輸出売上'],
      sign:       1,
    },
    {
      category:   'revenue',
      label:      '売上（その他）',
      indent:     1,
      accountNames: ['売上（その他）', 'その他売上', '売上その他'],
      sign:       1,
    },
    {
      category:   'revenue',
      label:      '売上値引・返品',
      indent:     1,
      accountNames: ['売上値引', '売上返品', '売上値引・返品', '返品・値引き'],
      sign:       -1,  // マイナス項目
    },
    {
      category:   'subtotal',
      label:      '売上高合計',
      indent:     0,
      calcFrom:   'revenue',
      isBold:     true,
      isBorderTop: true,
    },

    // ── 売上原価 ──────────────────────────────
    {
      category:   'header',
      label:      '売上原価',
      indent:     0,
      isBold:     true,
    },
    {
      category:   'cogs',
      label:      '期首商品棚卸高',
      indent:     1,
      accountNames: ['期首商品棚卸高', '商品期首棚卸高'],
      sign:       1,
    },
    {
      category:   'cogs',
      label:      '仕入（国内）',
      indent:     1,
      accountNames: ['仕入高', '仕入（国内）', '仕入-国内', '商品仕入高'],
      sign:       1,
    },
    {
      category:   'cogs',
      label:      '仕入（輸入）',
      indent:     1,
      accountNames: ['仕入（輸入）', '輸入仕入高', '仕入-輸入'],
      sign:       1,
    },
    {
      category:   'cogs',
      label:      '仕入値引・返品',
      indent:     1,
      accountNames: ['仕入値引', '仕入返品', '仕入値引・返品'],
      sign:       -1,
    },
    {
      category:   'cogs',
      label:      '期末商品棚卸高',
      indent:     1,
      accountNames: ['期末商品棚卸高', '商品期末棚卸高'],
      sign:       -1,  // マイナス項目（期末棚卸は原価から控除）
    },
    {
      category:   'subtotal',
      label:      '売上原価合計',
      indent:     0,
      calcFrom:   'cogs',
      isBold:     true,
      isBorderTop: true,
    },

    // ── 売上総利益 ──────────────────────────────
    {
      category:   'grossProfit',
      label:      '売上総利益',
      indent:     0,
      calcType:   'grossProfit', // 売上高合計 - 売上原価合計
      isBold:     true,
      isBorderTop: true,
    },

    // ── 販売費及び一般管理費 ──────────────────────────────
    {
      category:   'header',
      label:      '販売費及び一般管理費',
      indent:     0,
      isBold:     true,
    },
    { category: 'sga', label: '役員賞与',      indent: 1, accountNames: ['役員賞与'],              sign: 1 },
    { category: 'sga', label: '役員報酬',      indent: 1, accountNames: ['役員報酬'],              sign: 1 },
    { category: 'sga', label: '法定福利費',    indent: 1, accountNames: ['法定福利費'],            sign: 1 },
    { category: 'sga', label: '研修採用費',    indent: 1, accountNames: ['研修採用費', '採用費', '研修費'], sign: 1 },
    { category: 'sga', label: '接待交際費',    indent: 1, accountNames: ['接待交際費', '交際費'],  sign: 1 },
    { category: 'sga', label: '旅費交通費',    indent: 1, accountNames: ['旅費交通費', '交通費'],  sign: 1 },
    { category: 'sga', label: '通信費',        indent: 1, accountNames: ['通信費'],                sign: 1 },
    { category: 'sga', label: '水道光熱費',    indent: 1, accountNames: ['水道光熱費', '電気代'],  sign: 1 },
    { category: 'sga', label: '保険料',        indent: 1, accountNames: ['保険料'],                sign: 1 },
    { category: 'sga', label: '租税公課',      indent: 1, accountNames: ['租税公課'],              sign: 1 },
    { category: 'sga', label: '支払手数料',    indent: 1, accountNames: ['支払手数料'],            sign: 1 },
    { category: 'sga', label: '支払報酬',      indent: 1, accountNames: ['支払報酬', '報酬費'],    sign: 1 },
    { category: 'sga', label: '会議費',        indent: 1, accountNames: ['会議費'],                sign: 1 },
    { category: 'sga', label: '新聞図書費',    indent: 1, accountNames: ['新聞図書費', '図書費'],  sign: 1 },
    { category: 'sga', label: '減価償却費',    indent: 1, accountNames: ['減価償却費'],            sign: 1 },
    { category: 'sga', label: '繰延資産償却',  indent: 1, accountNames: ['繰延資産償却', '開発費償却'], sign: 1 },
    { category: 'sga', label: '業務委託費',    indent: 1, accountNames: ['業務委託費', '外注費'],  sign: 1 },
    { category: 'sga', label: '荷造運賃',      indent: 1, accountNames: ['荷造運賃', '発送費'],    sign: 1 },
    { category: 'sga', label: '広告宣伝費',    indent: 1, accountNames: ['広告宣伝費', '広告費'],  sign: 1 },
    { category: 'sga', label: '備品・消耗品費', indent: 1, accountNames: ['備品費', '消耗品費', '備品・消耗品費'], sign: 1 },
    { category: 'sga', label: '地代家賃',      indent: 1, accountNames: ['地代家賃', '賃借料'],    sign: 1 },
    { category: 'sga', label: '修繕費',        indent: 1, accountNames: ['修繕費', '修理費'],      sign: 1 },
    { category: 'sga', label: '雑費',          indent: 1, accountNames: ['雑費', 'その他費用', 'その他経費'], sign: 1 },
    { category: 'sga', label: '福利厚生費',    indent: 1, accountNames: ['福利厚生費'],            sign: 1 },
    { category: 'sga', label: '給料手当',      indent: 1, accountNames: ['給料手当', '給与手当', '給料'],   sign: 1 },
    {
      category:   'subtotal',
      label:      '販売費及び一般管理費合計',
      indent:     0,
      calcFrom:   'sga',
      isBold:     true,
      isBorderTop: true,
    },

    // ── 営業利益 ──────────────────────────────
    {
      category:   'operatingProfit',
      label:      '営業利益',
      indent:     0,
      calcType:   'operatingProfit', // 売上総利益 - 販売費及び一般管理費合計
      isBold:     true,
      isBorderTop: true,
    },

    // ── 営業外収益 ──────────────────────────────
    {
      category:   'header',
      label:      '営業外収益',
      indent:     0,
      isBold:     true,
    },
    { category: 'nonOpIncome', label: '受取利息',   indent: 1, accountNames: ['受取利息'],               sign: 1 },
    { category: 'nonOpIncome', label: '受取配当金', indent: 1, accountNames: ['受取配当金', '受取配当'], sign: 1 },
    {
      category:   'subtotal',
      label:      '営業外収益合計',
      indent:     0,
      calcFrom:   'nonOpIncome',
      isBold:     true,
      isBorderTop: true,
    },

    // ── 営業外費用 ──────────────────────────────
    {
      category:   'header',
      label:      '営業外費用',
      indent:     0,
      isBold:     true,
    },
    { category: 'nonOpExpense', label: '支払利息', indent: 1, accountNames: ['支払利息'],          sign: 1 },
    { category: 'nonOpExpense', label: '為替差損', indent: 1, accountNames: ['為替差損', '為替損'], sign: 1 },
    {
      category:   'subtotal',
      label:      '営業外費用合計',
      indent:     0,
      calcFrom:   'nonOpExpense',
      isBold:     true,
      isBorderTop: true,
    },

    // ── 経常利益 ──────────────────────────────
    {
      category:   'ordinaryProfit',
      label:      '経常利益',
      indent:     0,
      calcType:   'ordinaryProfit', // 営業利益 + 営業外収益合計 - 営業外費用合計
      isBold:     true,
      isBorderTop: true,
    },

    // ── 特別利益 ──────────────────────────────
    {
      category:   'header',
      label:      '特別利益',
      indent:     0,
      isBold:     true,
    },
    {
      category:   'subtotal',
      label:      '特別利益合計',
      indent:     0,
      calcFrom:   'extraIncome',
      isBold:     true,
      isBorderTop: true,
    },

    // ── 特別損失 ──────────────────────────────
    {
      category:   'header',
      label:      '特別損失',
      indent:     0,
      isBold:     true,
    },
    {
      category:   'subtotal',
      label:      '特別損失合計',
      indent:     0,
      calcFrom:   'extraExpense',
      isBold:     true,
      isBorderTop: true,
    },

    // ── 税引前当期純利益 ──────────────────────────────
    {
      category:   'header',
      label:      '当期純損益',
      indent:     0,
      isBold:     true,
    },
    {
      category:   'pretaxProfit',
      label:      '税引前当期純利益',
      indent:     0,
      calcType:   'pretaxProfit', // 経常利益 + 特別利益合計 - 特別損失合計
      isBold:     true,
      isBorderTop: true,
    },
    { category: 'tax', label: '法人税、住民税及び事業税', indent: 1, accountNames: ['法人税', '住民税', '法人税等', '法人税、住民税及び事業税'], sign: 1 },
    {
      category:   'netProfit',
      label:      '当期純利益',
      indent:     0,
      calcType:   'netProfit', // 税引前当期純利益 - 税金
      isBold:     true,
      isBorderTop: true,
    },
  ],

  // ==============================
  // 部門別CSVインポート設定
  // MF会計「推移試算表」を部門別にCSVエクスポート → Driveにアップロード → インポート
  // ==============================
  CSV_IMPORT: {
    // ★要設定: インポート用CSVを置くGoogleドライブのフォルダID
    // フォルダURLの末尾: drive.google.com/drive/folders/[ここがID]
    FOLDER_ID: '',

    // ファイル命名規則: 部門名.csv（Config.gs の DEPARTMENTS[].name と一致させること）
    // 例: 民泊.csv / 物販.csv / ブランド.csv / 共通.csv
  },

  // ==============================
  // シート名プレフィックス（事業年度ごとに「第8期_」のようなプレフィックスがつく）
  // ==============================
  SHEET_NAMES: {
    PL_PREFIX:      'PL_',       // 部門別PL（例: PL_共通、PL_物販）
    PL_CONSOLIDATED: 'PL_全体',  // 統合PL
    TREND_DEPT:     '推移_部門別', // 部門別推移表
    TREND_TOTAL:    '推移_全体',   // 統合推移表
  },

  // ==============================
  // 月の表示名（事業年度 3月始まり）
  // ==============================
  FISCAL_MONTHS: ['3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月', '1月', '2月'],
};

/**
 * 現在の事業年度を返す（3月始まりの場合）
 * 例: 2025年3月〜2026年2月 → 8期目（開始年=2025）
 */
function getCurrentFiscalYear() {
  const today = new Date();
  const month = today.getMonth() + 1; // 1-12
  const year  = today.getFullYear();
  // 3月以上なら今年度、2月以下なら前年度
  return month >= CONFIG.FISCAL.START_MONTH ? year : year - 1;
}

/**
 * 事業年度の開始日・終了日を返す
 * @param {number} fiscalYear - 事業年度開始年（例: 2025）
 * @return {{ start: Date, end: Date }}
 */
function getFiscalYearRange(fiscalYear) {
  const start = new Date(fiscalYear, CONFIG.FISCAL.START_MONTH - 1, 1);    // 3/1
  const end   = new Date(fiscalYear + 1, CONFIG.FISCAL.START_MONTH - 2, 28); // 翌2月末（実際は月末を取得）
  // 月末を正確に取得
  const endActual = new Date(fiscalYear + 1, CONFIG.FISCAL.START_MONTH - 1, 0);
  return { start, end: endActual };
}

/**
 * 事業年度内の月リストを返す [{year, month, label}]
 * @param {number} fiscalYear
 */
function getFiscalMonths(fiscalYear) {
  const months = [];
  for (let i = 0; i < 12; i++) {
    const month = ((CONFIG.FISCAL.START_MONTH - 1 + i) % 12) + 1;
    const year  = month >= CONFIG.FISCAL.START_MONTH ? fiscalYear : fiscalYear + 1;
    months.push({
      year,
      month,
      label: month + '月',
      startDate: Utilities.formatDate(new Date(year, month - 1, 1), 'Asia/Tokyo', 'yyyy-MM-dd'),
      endDate:   Utilities.formatDate(new Date(year, month, 0),     'Asia/Tokyo', 'yyyy-MM-dd'),
    });
  }
  return months;
}

/**
 * 期の表示名を返す（例: 2025 → "第8期"）
 * 会社設立を第1期として計算
 * LEVEL1は2018年3月を第1期として仮定
 */
function getFiscalPeriodLabel(fiscalYear) {
  const BASE_YEAR   = 2018; // 第1期開始年
  const periodNumber = fiscalYear - BASE_YEAR + 1;
  return `第${periodNumber}期`;
}

/**
 * スプレッドシートのシート名（期別プレフィックス付き）を生成
 * 例: 第8期_PL_物販
 */
function buildSheetName(fiscalYear, sheetKey, deptName) {
  const period = getFiscalPeriodLabel(fiscalYear);
  const base   = CONFIG.SHEET_NAMES[sheetKey] || sheetKey;
  return deptName ? `${period}_${base}${deptName}` : `${period}_${base}`;
}
