/**
 * Amazon Dashboard - 条件付き書式モジュール
 *
 * KPI のしきい値ハイライトを Range に適用するヘルパーを提供する。
 * Stage 2 で CategoryMonthly.gs から再利用予定。
 *
 * ## しきい値ルール
 *
 * | 指標 | 赤（注意） | 黄（警戒） | 緑（良好） |
 * |---|---|---|---|
 * | 粗利率 / 利益率 | < 0%   | < 10%  | ≥ 20% |
 * | TACOS          | > 30%  | ≥ 15%  | < 15% |
 * | ACOS           | > 40%  | ≥ 20%  | < 20% |
 * | ROAS           | < 3    | < 5    | ≥ 5   |
 *
 * 粗利率しきい値の根拠: 当社Amazon物販の目標粗利率 20% 以上（赤字=0% 未満）。
 * TACOSしきい値の根拠: kpi_targets.md「TACOS 30%超が3ヶ月連続で撤退検討」に準拠。
 */

// ===== 配色（薄め・目に優しい）=====
const FMT_COLOR_RED = '#f4c7c3';    // 注意（悪い）
const FMT_COLOR_YELLOW = '#fce8b2'; // 警戒（中間）
const FMT_COLOR_GREEN = '#b7e1cd';  // 良好（良い）

/**
 * しきい値ルールを1つのRangeに適用するヘルパー
 * @param {Array} rules 既存ルール配列（追加先）
 * @param {Range} range 対象レンジ
 * @param {string} metric 'profitMargin' | 'tacos' | 'acos' | 'roas'
 */
function addKpiRules(rules, range, metric) {
  if (!range) return;

  if (metric === 'profitMargin') {
    // 粗利率/利益率（比率 0.0〜1.0 形式）
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0)
        .setBackground(FMT_COLOR_RED)
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(0, 0.0999)
        .setBackground(FMT_COLOR_YELLOW)
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(0.20)
        .setBackground(FMT_COLOR_GREEN)
        .setRanges([range])
        .build()
    );
  } else if (metric === 'tacos') {
    // TACOS（比率 0.0〜1.0 形式）
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0.30)
        .setBackground(FMT_COLOR_RED)
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(0.15, 0.30)
        .setBackground(FMT_COLOR_YELLOW)
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(0.0001, 0.1499)
        .setBackground(FMT_COLOR_GREEN)
        .setRanges([range])
        .build()
    );
  } else if (metric === 'acos') {
    // ACOS（比率 0.0〜1.0 形式）
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0.40)
        .setBackground(FMT_COLOR_RED)
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(0.20, 0.40)
        .setBackground(FMT_COLOR_YELLOW)
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(0.0001, 0.1999)
        .setBackground(FMT_COLOR_GREEN)
        .setRanges([range])
        .build()
    );
  } else if (metric === 'roas') {
    // ROAS（絶対値・小数）
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(0.01, 2.9999)
        .setBackground(FMT_COLOR_RED)
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(3, 4.9999)
        .setBackground(FMT_COLOR_YELLOW)
        .setRanges([range])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(5)
        .setBackground(FMT_COLOR_GREEN)
        .setRanges([range])
        .build()
    );
  }
}
