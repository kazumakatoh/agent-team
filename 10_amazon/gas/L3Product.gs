/**
 * Amazon Dashboard - L3 商品分析（ASIN別一覧）
 *
 * 各ASINを1行として、売上〜利益率までの主要指標を一覧表示。
 * 前月同日数との比較で各セルに色付け（改善=緑 / 悪化=オレンジ）。
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
 * ## 色分けルール（前月同消化日数比較）
 *
 * **A類: 多いほど良い**（売上 / CV / 点数 / 利益 / 利益率 / ROAS）
 * - +10% 以上       : 🟢 緑
 * - -5%  〜 +10%    : ⚪ 色なし
 * - -15% 〜 -5%     : 🟠 薄オレンジ
 * - -15% 以下       : 🟠 濃オレンジ
 *
 * **B類: 少ないほど良い**（広告比率）
 * - -10% 以下       : 🟢 緑
 * - -5%  〜 +5%     : ⚪ 色なし
 * - +5%  〜 +15%    : 🟠 薄オレンジ
 * - +15% 以上       : 🟠 濃オレンジ
 *
 * **中立**（仕入 / 販売手数料 / 広告費 / その他経費 / 粗利率）→ 色付けなし
 *
 * ## 注意アイコン
 * - 利益マイナス: 🔴
 * - 広告費あり・売上ゼロ: ⚠️
 */

// ===== 色定義 =====
const L3_COLOR_POS_STRONG = '#93c47d';   // 緑（強）
const L3_COLOR_NEG_LIGHT  = '#fce5cd';   // オレンジ（薄）
const L3_COLOR_NEG_STRONG = '#f6b26b';   // オレンジ（濃）
const L3_COLOR_ROW_LOSS   = '#fbe9e7';   // 利益マイナス行の薄赤
const L3_COLOR_TOTAL_BG   = '#d0e1f9';   // 全体合計行の薄青

// ===== ヘッダー定義 =====
const L3_HEADERS = [
  'ASIN', '商品名', 'カテゴリ',
  '売上', 'CV', '点数',
  '仕入', '販売手数料', '広告費', 'その他経費',
  '利益', '広告比率', '粗利率', '利益率', 'ROAS',
  '注意',
];

// 色分け対象の列インデックス（1-indexed）と判定タイプ
// 'high-good' = A類（増で緑）, 'low-good' = B類（減で緑）
const L3_COLORED_COLS = [
  { col: 4,  metric: 'sales',        type: 'high-good' },  // 売上
  { col: 5,  metric: 'cv',           type: 'high-good' },  // CV
  { col: 6,  metric: 'units',        type: 'high-good' },  // 点数
  { col: 11, metric: 'profit',       type: 'high-good' },  // 利益
  { col: 12, metric: 'adRate',       type: 'low-good'  },  // 広告比率
  { col: 14, metric: 'profitMargin', type: 'high-good' },  // 利益率
  { col: 15, metric: 'roas',         type: 'high-good' },  // ROAS
];

/**
 * メイン: L3 商品分析シートを更新
 */
function updateDashboardL3() {
  const t0 = Date.now();
  Logger.log('===== L3 商品分析 更新開始 =====');

  const sheet = getOrCreateSheet(SHEET_NAMES.L3_PRODUCT);

  // 既存のフィルタを削除（clear前に実行する必要あり）
  const existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();

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

  // 期間別フィルタ
  const thisMonthDaily = dailyData.filter(d => d.date >= periods.thisMonth.start && d.date <= periods.thisMonth.end);
  const lastMonthDaily = dailyData.filter(d => d.date >= periods.lastMonthSameDay.start && d.date <= periods.lastMonthSameDay.end);

  // 経費を年月別にインデックス化
  const expByMonthAsin = {};
  for (const e of allExpenses) {
    if (!expByMonthAsin[e.yearMonth]) expByMonthAsin[e.yearMonth] = {};
    expByMonthAsin[e.yearMonth][e.asin] = e;
  }

  // ASIN別集計
  const thisByAsin = aggregateByAsinWithExpense(thisMonthDaily, expByMonthAsin, periods.thisMonth.start.substring(0, 7));
  const lastByAsin = aggregateByAsinWithExpense(lastMonthDaily, expByMonthAsin, periods.lastMonthSameDay.start.substring(0, 7));

  // ===== 書き込み =====

  // タイトル + ガイド
  sheet.getRange(1, 1).setValue('━━━ L3 商品分析 ━━━').setFontWeight('bold').setFontSize(14);
  sheet.getRange(2, 1)
    .setValue('※ 前月同消化日数比較で色分け（改善=緑 / 悪化=オレンジ） ／ フィルタ: メニュー → データ → フィルタを作成')
    .setFontStyle('italic').setFontColor('#666');

  // ヘッダー (行4)
  const headerRow = 4;
  sheet.getRange(headerRow, 1, 1, L3_HEADERS.length).setValues([L3_HEADERS])
    .setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center');

  // 全商品合計行 (行5)
  const totalRow = headerRow + 1;
  writeL3TotalRow(sheet, totalRow, thisByAsin, lastByAsin);

  // 商品行 (行6+)
  // 売上降順ソート
  const sortedAsins = Object.entries(thisByAsin)
    .sort((a, b) => b[1].sales - a[1].sales);

  const t2 = Date.now();
  const dataStartRow = totalRow + 1;
  writeL3ProductRows(sheet, dataStartRow, sortedAsins, lastByAsin);
  Logger.log('書き込み: ' + (Date.now() - t2) + 'ms (' + sortedAsins.length + '商品)');

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
function writeL3TotalRow(sheet, row, thisByAsin, lastByAsin) {
  const t = sumAsinAggs(Object.values(thisByAsin));
  const l = sumAsinAggs(Object.values(lastByAsin));

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
  applyL3RowColors(sheet, row, t, l);
}

/**
 * 商品行を一括書き込み
 */
function writeL3ProductRows(sheet, startRow, sortedAsins, lastByAsin) {
  const values = [];
  const rowLossFlags = []; // 利益マイナス行のフラグ

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
    rowLossFlags.push(a.profit < 0);
  }

  if (values.length === 0) return;

  // 一括書き込み
  sheet.getRange(startRow, 1, values.length, L3_HEADERS.length).setValues(values);

  // 数値フォーマットを一括適用
  applyL3RowNumberFormats(sheet, startRow, values.length);

  // 行ストライプ + 利益マイナス行の赤背景
  for (let i = 0; i < values.length; i++) {
    const row = startRow + i;
    if (rowLossFlags[i]) {
      sheet.getRange(row, 1, 1, L3_HEADERS.length).setBackground(L3_COLOR_ROW_LOSS);
    } else if (i % 2 === 1) {
      sheet.getRange(row, 1, 1, L3_HEADERS.length).setBackground('#f8f9fa'); // 薄グレーストライプ
    }
  }

  // 前月比較の色付け（各セル個別）
  for (let i = 0; i < sortedAsins.length; i++) {
    const [asin, a] = sortedAsins[i];
    const l = lastByAsin[asin];
    if (!l) continue; // 前月データがない新商品はスキップ
    applyL3RowColors(sheet, startRow + i, a, l);
  }
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
 * 前月比較で各セルに色を付ける
 */
function applyL3RowColors(sheet, row, current, previous) {
  for (const { col, metric, type } of L3_COLORED_COLS) {
    const cur = current[metric] || 0;
    const prev = previous[metric] || 0;
    const color = decideColor(cur, prev, type);
    if (color) {
      sheet.getRange(row, col).setBackground(color);
    }
  }
}

/**
 * 前月比の pct change から色を決定
 * @param {number} cur 当月値
 * @param {number} prev 前月値
 * @param {'high-good'|'low-good'} type
 * @returns {string|null} background color or null
 */
function decideColor(cur, prev, type) {
  if (!prev || prev === 0) return null;
  const pct = (cur - prev) / prev;

  if (type === 'high-good') {
    if (pct >= 0.10)  return L3_COLOR_POS_STRONG;
    if (pct >= -0.05) return null;
    if (pct >= -0.15) return L3_COLOR_NEG_LIGHT;
    return L3_COLOR_NEG_STRONG;
  }
  // low-good
  if (pct <= -0.10) return L3_COLOR_POS_STRONG;
  if (pct <= 0.05)  return null;
  if (pct <= 0.15)  return L3_COLOR_NEG_LIGHT;
  return L3_COLOR_NEG_STRONG;
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
