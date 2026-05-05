/**
 * Amazon Dashboard - L3 商品分析（ASIN別一覧）
 *
 * 各ASINを1行として、売上〜利益率までの主要指標を一覧表示。
 * **前月同消化日数との比較**で各セルをグラデーション色付け。
 *
 * 広告サマリ4指標（広告費 / 広告比率 / 利益率 / ROAS）まで表示。
 *
 * ## 列構造（16列）
 *
 * [固定] ASIN | 商品名 | カテゴリ
 * [指標] 売上 | CV | 点数 | 仕入 | 販売手数料 | 広告費 | その他経費 | 利益
 *        | 広告比率 | 粗利率 | 利益率 | ROAS | 注意
 *
 * ## カラーロジック（前月同消化日数比較・グラデーション）
 *
 * 前月比の変化率を ±20% の範囲にクランプし、白 → 緑 or 白 → 赤 に線形補間。
 * |前月比| < 5% は色なし（視覚ノイズ低減）。
 *
 * **A類: 高いほど良い**（売上 / CV / 点数 / 利益 / 粗利率 / 利益率 / ROAS）
 * - +20% 以上   : 🟢 最濃緑
 * - +10%        : 🟢 中緑
 * - ±5% 内      : ⚪ 色なし
 * - -10%        : 🔴 中赤
 * - -20% 以下   : 🔴 最濃赤
 *
 * **B類: 低いほど良い**（広告比率）
 * - -20% 以下   : 🟢 最濃緑（大幅改善）
 * - -10%        : 🟢 中緑
 * - ±5% 内      : ⚪ 色なし
 * - +10%        : 🔴 中赤
 * - +20% 以上   : 🔴 最濃赤（大幅悪化）
 *
 * **中立**（仕入 / 販売手数料 / 広告費 / その他経費）→ 色付けなし
 *
 * ## ソート
 *
 * 売上額の降順（updateDashboardL3 実行時点で確定）。
 *
 * ## 注意アイコン
 * - 利益マイナス: 🔴
 * - 広告費あり・売上ゼロ: ⚠️
 */

// ===== グラデーションのターゲット色（薄めで視認負荷を下げる）=====
const L3_RGB_GREEN = [183, 225, 205]; // #b7e1cd（前月比で伸びている）
const L3_RGB_RED   = [244, 199, 195]; // #f4c7c3（前月比で落ちている）
const L3_RGB_WHITE = [255, 255, 255]; // 白（±5% 以内）

// グラデーションの飽和ポイント: ±20% でフル彩度
const L3_GRADIENT_MAX_PCT = 0.20;
// ±5% 以内は色を付けない（視覚ノイズ低減）
const L3_GRADIENT_DEADBAND = 0.05;

// ===== 合計行の色 =====
const L3_COLOR_TOTAL_BG = '#d0e1f9';

// ===== ヘッダー定義 =====
const L3_HEADERS = [
  'ASIN', '商品名', 'カテゴリ',
  '売上', 'CV', '点数',
  '仕入', '販売手数料', '広告費', 'その他経費',
  '利益', '広告比率', '粗利率', '利益率', 'ROAS',
  '注意',
];

// 色付け対象の列（1-indexed）と判定タイプ
const L3_COLORED_COLS = [
  { col: 4,  metric: 'sales',        type: 'high-good' },
  { col: 5,  metric: 'cv',           type: 'high-good' },
  { col: 6,  metric: 'units',        type: 'high-good' },
  { col: 11, metric: 'profit',       type: 'high-good' },
  { col: 12, metric: 'adRate',       type: 'low-good'  },
  { col: 13, metric: 'grossMargin',  type: 'high-good' },
  { col: 14, metric: 'profitMargin', type: 'high-good' },
  { col: 15, metric: 'roas',         type: 'high-good' },
];

/**
 * メイン: L3 商品分析シートを更新
 */
function updateDashboardL3() {
  const t0 = Date.now();
  Logger.log('===== L3 商品分析 更新開始 =====');

  const sheet = getOrCreateSheet(SHEET_NAMES.L3_PRODUCT);

  // 既存のフィルタ・条件付き書式を削除（前回カラースケールの痕跡を除去）
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

  // 期間別フィルタ（当月 + 前月同日）
  const thisMonthDaily = dailyData.filter(d => d.date >= periods.thisMonth.start && d.date <= periods.thisMonth.end);
  const lastMonthDaily = dailyData.filter(d => d.date >= periods.lastMonthSameDay.start && d.date <= periods.lastMonthSameDay.end);

  // 経費を年月別にインデックス化
  const expByMonthAsin = {};
  for (const e of allExpenses) {
    if (!expByMonthAsin[e.yearMonth]) expByMonthAsin[e.yearMonth] = {};
    expByMonthAsin[e.yearMonth][e.asin] = e;
  }

  // ASIN別集計（当月 + 前月同日）
  const thisByAsin = aggregateByAsinWithExpense(thisMonthDaily, expByMonthAsin, periods.thisMonth.start.substring(0, 7));
  const lastByAsin = aggregateByAsinWithExpense(lastMonthDaily, expByMonthAsin, periods.lastMonthSameDay.start.substring(0, 7));

  // ===== 書き込み =====

  // タイトル + ガイド
  sheet.getRange(1, 1).setValue('━━━ L3 商品分析 ━━━').setFontWeight('bold').setFontSize(14);
  sheet.getRange(2, 1)
    .setValue('※ 前月同消化日数比較でグラデーション色分け（緑=伸びている / 赤=落ちている / 白=±5%以内） ／ 売上降順でソート済み')
    .setFontStyle('italic').setFontColor('#666');

  // ヘッダー (行4)
  const headerRow = 4;
  sheet.getRange(headerRow, 1, 1, L3_HEADERS.length).setValues([L3_HEADERS])
    .setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center');

  // 全商品合計行 (行5) - 合計ベースで前月比色も適用
  const totalRow = headerRow + 1;
  writeL3TotalRow(sheet, totalRow, thisByAsin, lastByAsin);

  // 売上額の降順でソート（明示・ユーザー要求 2026-04-15）
  const sortedAsins = Object.entries(thisByAsin)
    .sort((a, b) => b[1].sales - a[1].sales);

  // 商品行 (行6+) - 各商品ごとの前月比色
  const t2 = Date.now();
  const dataStartRow = totalRow + 1;
  writeL3ProductRows(sheet, dataStartRow, sortedAsins, lastByAsin);
  Logger.log('書き込み + 色付け: ' + (Date.now() - t2) + 'ms (' + sortedAsins.length + '商品)');

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
        adCost: 0, adSales: 0, cogs: 0,
      };
    }
    const a = byAsin[d.asin];
    a.sales += d.sales;
    a.cv += d.cv;
    a.units += d.units;
    a.adCost += d.adCost;
    a.adSales += d.adSales;
    a.cogs += d.cogs || 0; // CF連携: D1 仕入原価合計から集計
  }

  // その月の commission率 / other率 を算出（D2S の全ASIN合計から）
  const monthExp = expByMonthAsin[yearMonth] || {};
  let totalCommission = 0, totalOther = 0, totalPrincipal = 0;
  for (const exp of Object.values(monthExp)) {
    totalCommission += exp.commission || 0;
    totalOther += exp.other || 0;
    totalPrincipal += exp.principal || 0;
  }
  const commissionRate = totalPrincipal > 0 ? totalCommission / totalPrincipal : 0;
  const otherRate = totalPrincipal > 0 ? totalOther / totalPrincipal : 0;

  // 各ASINの経費を売上 × 率 で推定
  for (const asin of Object.keys(byAsin)) {
    const a = byAsin[asin];
    a.commission = a.sales * commissionRate;
    a.otherExpense = a.sales * otherRate;
    a.profit = a.sales - a.cogs - a.commission - a.otherExpense - a.adCost;
    a.adRate = a.sales > 0 ? a.adCost / a.sales : 0;
    a.grossMargin = a.sales > 0 ? (a.sales - a.cogs) / a.sales : 0;
    a.profitMargin = a.sales > 0 ? a.profit / a.sales : 0;
    a.roas = a.adCost > 0 ? a.sales / a.adCost : 0;
  }
  return byAsin;
}

/**
 * 全商品合計行を書き込み（前月比色付きで全体の伸び/縮みも可視化）
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

  // 合計行の前月比色（個別セル）
  for (const { col, metric, type } of L3_COLORED_COLS) {
    const color = pctToGradientColor(t[metric], l[metric], type);
    if (color) sheet.getRange(row, col).setBackground(color);
  }
}

/**
 * 商品行を一括書き込み + 前月比グラデーション色を setBackgrounds で適用
 */
function writeL3ProductRows(sheet, startRow, sortedAsins, lastByAsin) {
  const colCount = L3_HEADERS.length;
  const values = [];
  const bgColors = []; // 2次元配列: 色の文字列 or null

  for (const [asin, a] of sortedAsins) {
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

    // 行ごとの背景色配列を初期化（全 null = 色なし）
    const rowBg = new Array(colCount).fill(null);

    // 前月データがあれば色を計算
    const prev = lastByAsin[asin];
    if (prev) {
      for (const { col, metric, type } of L3_COLORED_COLS) {
        const color = pctToGradientColor(a[metric], prev[metric], type);
        if (color) rowBg[col - 1] = color;
      }
    }
    bgColors.push(rowBg);
  }

  if (values.length === 0) return;

  // 値を一括書き込み
  sheet.getRange(startRow, 1, values.length, colCount).setValues(values);

  // 数値フォーマット
  applyL3RowNumberFormats(sheet, startRow, values.length);

  // 背景色を一括適用（nullは色なし）
  sheet.getRange(startRow, 1, bgColors.length, colCount).setBackgrounds(bgColors);
}

/**
 * L3 行の数値フォーマット設定
 */
function applyL3RowNumberFormats(sheet, startRow, numRows) {
  const n = numRows || 1;
  [4, 5, 6, 7, 8, 9, 10, 11].forEach(col => {
    sheet.getRange(startRow, col, n, 1).setNumberFormat('#,##0');
  });
  [12, 13, 14].forEach(col => {
    sheet.getRange(startRow, col, n, 1).setNumberFormat('0.0%');
  });
  sheet.getRange(startRow, 15, n, 1).setNumberFormat('0.00');
  sheet.getRange(startRow, 4, n, 12).setHorizontalAlignment('right');
  sheet.getRange(startRow, 16, n, 1).setHorizontalAlignment('center');
}

/**
 * 前月比からグラデーション色を計算
 *
 * @param {number} curVal 当月値
 * @param {number} prevVal 前月値
 * @param {'high-good'|'low-good'} type 指標の性質
 * @returns {string|null} hex color（'#rrggbb'）or null（色なし）
 */
function pctToGradientColor(curVal, prevVal, type) {
  // 前月値がゼロまたは未定義 → 色なし（比較不能）
  if (!prevVal || prevVal === 0) return null;
  if (curVal === undefined || curVal === null) return null;

  const pct = (curVal - prevVal) / prevVal;
  if (!isFinite(pct)) return null;

  // low-good の場合は符号反転（減少 = 改善 = 緑）
  const signedPct = type === 'low-good' ? -pct : pct;

  // ±5% 以内は色なし
  if (Math.abs(signedPct) < L3_GRADIENT_DEADBAND) return null;

  // ±20% でフル彩度（クランプ）
  const clamped = Math.max(-L3_GRADIENT_MAX_PCT, Math.min(L3_GRADIENT_MAX_PCT, signedPct));
  const ratio = clamped / L3_GRADIENT_MAX_PCT; // -1 〜 +1
  const absRatio = Math.abs(ratio);

  // 白 → 緑（正）or 白 → 赤（負）に線形補間
  const targetRGB = ratio > 0 ? L3_RGB_GREEN : L3_RGB_RED;
  const r = Math.round(L3_RGB_WHITE[0] + (targetRGB[0] - L3_RGB_WHITE[0]) * absRatio);
  const g = Math.round(L3_RGB_WHITE[1] + (targetRGB[1] - L3_RGB_WHITE[1]) * absRatio);
  const b = Math.round(L3_RGB_WHITE[2] + (targetRGB[2] - L3_RGB_WHITE[2]) * absRatio);

  return '#' + [r, g, b].map(x => x.toString(16).padStart(2, '0')).join('');
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
    total.cogs += a.cogs || 0;
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
