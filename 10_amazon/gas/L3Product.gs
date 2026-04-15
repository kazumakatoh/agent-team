/**
 * Amazon Dashboard - L3 商品分析（ASIN別一覧）
 *
 * 各ASINを1行として、売上〜利益率までの主要指標を一覧表示。
 * カラースケール（条件付き書式のグラデーション）で商品間の相対順位を可視化。
 *
 * ## 広告詳細の取り扱い
 *
 * L3 には広告サマリ4指標（広告費 / 広告比率 / 利益率 / ROAS）のみ表示し、
 * キャンペーン/キーワード/検索用語といった広告詳細は D3 広告詳細シートへ委ねる。
 *
 * ## 列構造（16列）
 *
 * [固定] ASIN | 商品名 | カテゴリ
 * [指標] 売上 | CV | 点数 | 仕入 | 販売手数料 | 広告費 | その他経費 | 利益
 *        | 広告比率 | 粗利率 | 利益率 | ROAS | 注意
 *
 * ## カラースケール（緑 → 白 → 赤）
 *
 * **A類: 高いほど良い**（売上 / CV / 点数 / 利益 / 粗利率 / 利益率 / ROAS）
 * - 最小 → 🔴赤   中央 → ⚪白   最大 → 🟢緑
 *
 * **B類: 低いほど良い**（広告比率）
 * - 最小 → 🟢緑   中央 → ⚪白   最大 → 🔴赤
 *
 * **中立**（仕入 / 販売手数料 / 広告費 / その他経費）→ 色付けなし
 *
 * ## ソート
 *
 * 売上額の降順（最大値が先頭、売上ゼロは末尾）。updateDashboardL3 の実行時点で確定。
 *
 * ## 注意アイコン
 * - 利益マイナス: 🔴
 * - 広告費あり・売上ゼロ: ⚠️
 */

// ===== カラースケール色定義（ユーザー承認済み・Google Sheets標準パレット）=====
const L3_SCALE_GREEN = '#57bb8a';  // 最良（高いほど良いなら最大 / 低いほど良いなら最小）
const L3_SCALE_WHITE = '#ffffff';  // 中央（パーセンタイル50）
const L3_SCALE_RED   = '#e67c73';  // 最悪

// ===== 合計行・ヘッダー色 =====
const L3_COLOR_TOTAL_BG = '#d0e1f9';  // 全体合計行の薄青

// ===== ヘッダー定義 =====
const L3_HEADERS = [
  'ASIN', '商品名', 'カテゴリ',
  '売上', 'CV', '点数',
  '仕入', '販売手数料', '広告費', 'その他経費',
  '利益', '広告比率', '粗利率', '利益率', 'ROAS',
  '注意',
];

// カラースケール対象の列インデックス（1-indexed）と方向
// 'high-good' = 高い方が良い（min=赤, max=緑）
// 'low-good'  = 低い方が良い（min=緑, max=赤）
const L3_SCALE_COLS = [
  { col: 4,  type: 'high-good' },  // 売上
  { col: 5,  type: 'high-good' },  // CV
  { col: 6,  type: 'high-good' },  // 点数
  { col: 11, type: 'high-good' },  // 利益
  { col: 12, type: 'low-good'  },  // 広告比率
  { col: 13, type: 'high-good' },  // 粗利率
  { col: 14, type: 'high-good' },  // 利益率
  { col: 15, type: 'high-good' },  // ROAS
];

/**
 * メイン: L3 商品分析シートを更新
 */
function updateDashboardL3() {
  const t0 = Date.now();
  Logger.log('===== L3 商品分析 更新開始 =====');

  const sheet = getOrCreateSheet(SHEET_NAMES.L3_PRODUCT);

  // 既存のフィルタ・条件付き書式を削除
  const existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();
  sheet.clearConditionalFormatRules();
  sheet.clear();

  // データ準備
  const t1 = Date.now();
  const dailyData = getDailyDataAll();
  const allExpenses = readAllSettlement();
  const periods = getPeriods();
  Logger.log('データ読み込み: ' + (Date.now() - t1) + 'ms (' + dailyData.length + '行)');

  if (dailyData.length === 0) {
    sheet.getRange(1, 1).setValue('日次データがありません');
    return;
  }

  const thisMonthDaily = dailyData.filter(d => d.date >= periods.thisMonth.start && d.date <= periods.thisMonth.end);

  // 経費を年月別にインデックス化
  const expByMonthAsin = {};
  for (const e of allExpenses) {
    if (!expByMonthAsin[e.yearMonth]) expByMonthAsin[e.yearMonth] = {};
    expByMonthAsin[e.yearMonth][e.asin] = e;
  }

  // ASIN別集計
  const thisByAsin = aggregateByAsinWithExpense(thisMonthDaily, expByMonthAsin, periods.thisMonth.start.substring(0, 7));

  // ===== 書き込み =====

  // タイトル + ガイド
  sheet.getRange(1, 1).setValue('━━━ L3 商品分析 ━━━').setFontWeight('bold').setFontSize(14);
  sheet.getRange(2, 1)
    .setValue('※ カラースケールで商品間の相対順位を色分け（緑=良 / 赤=悪） ／ 売上降順でソート済み')
    .setFontStyle('italic').setFontColor('#666');

  // ヘッダー (行4)
  const headerRow = 4;
  sheet.getRange(headerRow, 1, 1, L3_HEADERS.length).setValues([L3_HEADERS])
    .setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center');

  // 全商品合計行 (行5)
  const totalRow = headerRow + 1;
  writeL3TotalRow(sheet, totalRow, thisByAsin);

  // 売上額の降順でソート（明示的・ユーザー要求 2026-04-15）
  const sortedAsins = Object.entries(thisByAsin)
    .sort((a, b) => b[1].sales - a[1].sales);

  // 商品行 (行6+)
  const t2 = Date.now();
  const dataStartRow = totalRow + 1;
  writeL3ProductRows(sheet, dataStartRow, sortedAsins);
  Logger.log('書き込み: ' + (Date.now() - t2) + 'ms (' + sortedAsins.length + '商品)');

  // カラースケール適用（商品行のみ、合計行は除外）
  const t3 = Date.now();
  if (sortedAsins.length > 0) {
    applyL3ColorScale(sheet, dataStartRow, sortedAsins.length);
  }
  Logger.log('カラースケール: ' + (Date.now() - t3) + 'ms');

  // フリーズ（列3まで・行4まで）
  sheet.setFrozenColumns(3);
  sheet.setFrozenRows(headerRow);

  // 列幅調整
  sheet.setColumnWidth(1, 110);    // ASIN
  sheet.setColumnWidth(2, 200);    // 商品名
  sheet.setColumnWidth(3, 120);    // カテゴリ

  Logger.log('===== L3 完了（合計 ' + (Date.now() - t0) + 'ms）=====');
}

/**
 * ASIN別に日次データ + 経費を集計
 */
function aggregateByAsinWithExpense(dailyData, expByMonthAsin, yearMonth) {
  const byAsin = {};
  for (const d of dailyData) {
    if (!d.asin) continue;
    if (!byAsin[d.asin]) {
      byAsin[d.asin] = {
        asin: d.asin,
        name: d.name || '',
        category: d.category || '(未分類)',
        sales: 0, cv: 0, units: 0,
        adCost: 0, adSales: 0,
      };
    }
    const a = byAsin[d.asin];
    a.sales += d.sales;
    a.cv += d.cv;
    a.units += d.units;
    a.adCost += d.adCost;
    a.adSales += d.adSales;
  }

  // 経費を付与
  const monthExp = expByMonthAsin[yearMonth] || {};
  for (const asin of Object.keys(byAsin)) {
    const exp = monthExp[asin] || { commission: 0, other: 0 };
    const a = byAsin[asin];
    a.commission = exp.commission;
    a.otherExpense = exp.other;
    a.cogs = 0; // TODO: CF連携
    a.profit = a.sales - a.cogs - a.commission - a.otherExpense - a.adCost;
    a.adRate = a.sales > 0 ? a.adCost / a.sales : 0;
    a.grossMargin = a.sales > 0 ? (a.sales - a.cogs) / a.sales : 0;
    a.profitMargin = a.sales > 0 ? a.profit / a.sales : 0;
    a.roas = a.adCost > 0 ? a.sales / a.adCost : 0;
  }
  return byAsin;
}

/**
 * 全商品合計行を書き込み
 */
function writeL3TotalRow(sheet, row, thisByAsin) {
  const t = sumAsinAggs(Object.values(thisByAsin));
  const values = [[
    '【全商品合計】', '-', '-',
    t.sales, t.cv, t.units,
    t.cogs, t.commission, t.adCost, t.otherExpense,
    t.profit, t.adRate, t.grossMargin, t.profitMargin, t.roas,
    '',
  ]];

  sheet.getRange(row, 1, 1, L3_HEADERS.length).setValues(values)
    .setFontWeight('bold').setBackground(L3_COLOR_TOTAL_BG);

  applyL3RowNumberFormats(sheet, row);
}

/**
 * 商品行を一括書き込み
 */
function writeL3ProductRows(sheet, startRow, sortedAsins) {
  const values = [];
  for (const [asin, a] of sortedAsins) {
    // 注意アイコン
    let alert = '';
    if (a.profit < 0) alert = '🔴';
    else if (a.sales === 0 && a.adCost > 0) alert = '⚠️';

    values.push([
      asin, a.name, a.category,
      a.sales, a.cv, a.units,
      a.cogs, a.commission, a.adCost, a.otherExpense,
      a.profit, a.adRate, a.grossMargin, a.profitMargin, a.roas,
      alert,
    ]);
  }

  if (values.length === 0) return;

  // 一括書き込み
  sheet.getRange(startRow, 1, values.length, L3_HEADERS.length).setValues(values);

  // 数値フォーマットを一括適用
  applyL3RowNumberFormats(sheet, startRow, values.length);
}

/**
 * L3 行の数値フォーマット設定
 */
function applyL3RowNumberFormats(sheet, startRow, numRows) {
  const n = numRows || 1;
  // 整数カラム: 売上(4), CV(5), 点数(6), 仕入(7), 販売手数料(8), 広告費(9), その他経費(10), 利益(11)
  [4, 5, 6, 7, 8, 9, 10, 11].forEach(col => {
    sheet.getRange(startRow, col, n, 1).setNumberFormat('#,##0');
  });
  // 率カラム: 広告比率(12), 粗利率(13), 利益率(14)
  [12, 13, 14].forEach(col => {
    sheet.getRange(startRow, col, n, 1).setNumberFormat('0.0%');
  });
  // ROAS(15)
  sheet.getRange(startRow, 15, n, 1).setNumberFormat('0.00');
  // 右寄せ
  sheet.getRange(startRow, 4, n, 12).setHorizontalAlignment('right');
  // 注意(16)は中央
  sheet.getRange(startRow, 16, n, 1).setHorizontalAlignment('center');
}

/**
 * カラースケール条件付き書式を適用
 *
 * 各指標列に 3色グラデーション（最小-中央パーセンタイル50-最大）を設定。
 * 高いほど良い指標は 赤→白→緑、低いほど良い指標は 緑→白→赤。
 *
 * @param {Sheet} sheet L3シート
 * @param {number} dataStartRow 商品行の開始行（合計行は除外）
 * @param {number} rowCount 商品行数
 */
function applyL3ColorScale(sheet, dataStartRow, rowCount) {
  const rules = [];

  for (const { col, type } of L3_SCALE_COLS) {
    const range = sheet.getRange(dataStartRow, col, rowCount, 1);

    if (type === 'high-good') {
      // 最小=赤 / 中央=白 / 最大=緑
      rules.push(buildGradientRule(range, L3_SCALE_RED, L3_SCALE_WHITE, L3_SCALE_GREEN));
    } else {
      // 低いほど良い: 最小=緑 / 中央=白 / 最大=赤
      rules.push(buildGradientRule(range, L3_SCALE_GREEN, L3_SCALE_WHITE, L3_SCALE_RED));
    }
  }

  sheet.setConditionalFormatRules(rules);
  Logger.log('✅ L3 カラースケール適用: ' + rules.length + ' 列');
}

/**
 * カラースケール（3色グラデーション）のルールを構築
 */
function buildGradientRule(range, minColor, midColor, maxColor) {
  return SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue(minColor, SpreadsheetApp.InterpolationType.MIN, '')
    .setGradientMidpointWithValue(midColor, SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpointWithValue(maxColor, SpreadsheetApp.InterpolationType.MAX, '')
    .setRanges([range])
    .build();
}

/**
 * ASIN集計の合計
 */
function sumAsinAggs(asinList) {
  const total = {
    sales: 0, cv: 0, units: 0,
    cogs: 0, commission: 0, adCost: 0, otherExpense: 0, adSales: 0,
    profit: 0,
  };
  for (const a of asinList) {
    total.sales += a.sales;
    total.cv += a.cv;
    total.units += a.units;
    total.cogs += a.cogs;
    total.commission += a.commission;
    total.adCost += a.adCost;
    total.otherExpense += a.otherExpense;
    total.adSales += a.adSales;
    total.profit += a.profit;
  }
  total.adRate = total.sales > 0 ? total.adCost / total.sales : 0;
  total.grossMargin = total.sales > 0 ? (total.sales - total.cogs) / total.sales : 0;
  total.profitMargin = total.sales > 0 ? total.profit / total.sales : 0;
  total.roas = total.adCost > 0 ? total.sales / total.adCost : 0;
  return total;
}
