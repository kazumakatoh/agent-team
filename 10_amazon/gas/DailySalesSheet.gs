/**
 * Amazon Dashboard - 日次販売実績シート
 *
 * D1 日次データ + 経費月次集計(D2S) + 販促費(M3) を組み合わせて、
 * 1日 = 1行 の販売実績を「日次販売実績」シートに集計表示。
 *
 * ## 列構造（13列）
 *
 * 日付 | 売上 | CV | 注文点数 | 仕入 | 販売手数料 | 広告費 | その他経費 | 利益
 *      | 広告比率 | 粗利率 | 利益率 | ROAS
 *
 * ## 経費の日割り按分
 *
 * 仕入・販売手数料・その他経費は月次集計しか持っていないため、
 * その月の日次売上比率で按分する。
 *
 *   日次経費 = 月次経費 × (日次売上 / 月次売上合計)
 *
 * 売上ゼロの日は経費もゼロ扱い。
 *
 * ## トリガー: 毎日 AM8:30 (buildDailySalesSheet)
 */

const D1S_HEADERS = [
  '日付', '売上', 'CV', '注文点数',
  '仕入', '販売手数料', '広告費', 'その他経費',
  '利益', '広告比率', '粗利率', '利益率', 'ROAS',
];
const D1S_COLS = D1S_HEADERS.length;

/**
 * メイン: 日次販売実績シートを構築
 */
function buildDailySalesSheet() {
  const t0 = Date.now();
  Logger.log('===== 日次販売実績シート構築 開始 =====');

  const sheet = getOrCreateSheet(SHEET_NAMES.D1S_DAILY_SALES);

  // 既存の条件付き書式・データを削除
  sheet.clearConditionalFormatRules();
  sheet.clear();

  // データソース読み込み
  const dailyData = getDailyDataAll();
  if (dailyData.length === 0) {
    sheet.getRange(1, 1).setValue('日次データがありません');
    return;
  }

  const allExpenses = readAllSettlement();

  // 経費を年月別にインデックス化（ASIN問わず月合計）
  const expByMonth = {};
  for (const e of allExpenses) {
    if (!expByMonth[e.yearMonth]) expByMonth[e.yearMonth] = { commission: 0, other: 0 };
    expByMonth[e.yearMonth].commission += e.commission;
    expByMonth[e.yearMonth].other += e.other;
  }

  // 日次データを「日付」で集計（全ASIN合計）
  const byDate = {};
  for (const d of dailyData) {
    if (!d.date) continue;
    if (!byDate[d.date]) {
      byDate[d.date] = { sales: 0, cv: 0, units: 0, adCost: 0, adSales: 0 };
    }
    const x = byDate[d.date];
    x.sales += d.sales;
    x.cv += d.cv;
    x.units += d.units;
    x.adCost += d.adCost;
    x.adSales += d.adSales;
  }

  // 月次売上合計（按分用の分母）
  const monthlySales = {};
  for (const [date, x] of Object.entries(byDate)) {
    const ym = date.substring(0, 7);
    monthlySales[ym] = (monthlySales[ym] || 0) + x.sales;
  }

  // 日付昇順でソートして行に変換
  const sortedDates = Object.keys(byDate).sort();
  const rows = [];
  for (const date of sortedDates) {
    const d = byDate[date];
    const ym = date.substring(0, 7);
    const monthSales = monthlySales[ym] || 0;
    const ratio = monthSales > 0 ? d.sales / monthSales : 0;

    // 月次経費を売上比率で按分
    const exp = expByMonth[ym] || { commission: 0, other: 0 };
    const commission = exp.commission * ratio;
    const otherExpense = exp.other * ratio;
    const cogs = 0; // TODO: CF連携後

    const profit = d.sales - cogs - commission - otherExpense - d.adCost;
    const adRate = d.sales > 0 ? d.adCost / d.sales : 0;
    const grossMargin = d.sales > 0 ? (d.sales - cogs) / d.sales : 0;
    const profitMargin = d.sales > 0 ? profit / d.sales : 0;
    const roas = d.adCost > 0 ? d.sales / d.adCost : 0;

    rows.push([
      date, d.sales, d.cv, d.units,
      cogs, commission, d.adCost, otherExpense,
      profit, adRate, grossMargin, profitMargin, roas,
    ]);
  }

  // タイトル + ガイド
  sheet.getRange(1, 1).setValue('━━━ 日次販売実績 ━━━').setFontWeight('bold').setFontSize(14);
  sheet.getRange(2, 1)
    .setValue('※ 経費は月次値を売上比率で按分 ／ 仕入は CF連携後に反映')
    .setFontStyle('italic').setFontColor('#666');

  // ヘッダー（行4）
  const headerRow = 4;
  sheet.getRange(headerRow, 1, 1, D1S_COLS).setValues([D1S_HEADERS])
    .setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center');

  // データ書き込み
  if (rows.length > 0) {
    const dataStartRow = headerRow + 1;
    sheet.getRange(dataStartRow, 1, rows.length, D1S_COLS).setValues(rows);

    // 数値フォーマット
    // 整数: 売上(2), CV(3), 点数(4), 仕入(5), 手数料(6), 広告費(7), 経費(8), 利益(9)
    [2, 3, 4, 5, 6, 7, 8, 9].forEach(col => {
      sheet.getRange(dataStartRow, col, rows.length, 1).setNumberFormat('#,##0');
    });
    // 比率: 広告比率(10), 粗利率(11), 利益率(12)
    [10, 11, 12].forEach(col => {
      sheet.getRange(dataStartRow, col, rows.length, 1).setNumberFormat('0.0%');
    });
    // ROAS(13)
    sheet.getRange(dataStartRow, 13, rows.length, 1).setNumberFormat('0.00');
    // 右寄せ
    sheet.getRange(dataStartRow, 2, rows.length, D1S_COLS - 1).setHorizontalAlignment('right');
    // 日付列センター揃え
    sheet.getRange(dataStartRow, 1, rows.length, 1).setHorizontalAlignment('center');
  }

  // フリーズ
  sheet.setFrozenColumns(1);
  sheet.setFrozenRows(headerRow);

  // 列幅
  sheet.setColumnWidth(1, 100);

  Logger.log('✅ 日次販売実績: ' + rows.length + ' 行 (' + (Date.now() - t0) + 'ms)');
}
