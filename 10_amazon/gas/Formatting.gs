/**
 * Amazon Dashboard - 条件付き書式モジュール
 *
 * L1/L2 ダッシュボードの数値セルに KPI しきい値ハイライトを適用する。
 * Dashboard.gs が sheet.clear() を呼んだ後に再適用する必要あり。
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

/**
 * L1 事業ダッシュボードに条件付き書式を適用
 *
 * 対象範囲:
 *   全体サマリー（行3-7）: 粗利率(11)/広告比率=TACOS(12)/ROAS(13)/利益率(14)
 *   カテゴリ別サマリー（行 categoryStartRow〜categoryEndRow）:
 *     TACOS(8)/ACOS(10)/ROAS(12)/利益率(13)
 *
 * @param {Sheet} sheet L1シート
 * @param {number} categoryStartRow カテゴリデータ開始行
 * @param {number} categoryRowCount カテゴリ行数
 */
function applyL1ConditionalFormatting(sheet, categoryStartRow, categoryRowCount) {
  const rules = [];

  // 全体サマリー（行3-7 = 全体/広告/オーガニック/前月比/月末予測）
  // 前月比(行6)は pctChange なので除外、全体/月末予測を主対象に
  const overallRows = 5; // 行3〜7
  const overallStartRow = 3;

  // 粗利率（列11）
  addKpiRules(rules, sheet.getRange(overallStartRow, 11, overallRows, 1), 'profitMargin');
  // 広告比率=TACOS（列12）
  addKpiRules(rules, sheet.getRange(overallStartRow, 12, overallRows, 1), 'tacos');
  // ROAS（列13）
  addKpiRules(rules, sheet.getRange(overallStartRow, 13, overallRows, 1), 'roas');
  // 利益率（列14）
  addKpiRules(rules, sheet.getRange(overallStartRow, 14, overallRows, 1), 'profitMargin');

  // カテゴリ別サマリー
  if (categoryRowCount > 0) {
    // TACOS（列8）
    addKpiRules(rules, sheet.getRange(categoryStartRow, 8, categoryRowCount, 1), 'tacos');
    // ACOS（列10）
    addKpiRules(rules, sheet.getRange(categoryStartRow, 10, categoryRowCount, 1), 'acos');
    // ROAS（列12）
    addKpiRules(rules, sheet.getRange(categoryStartRow, 12, categoryRowCount, 1), 'roas');
    // 利益率（列13）
    addKpiRules(rules, sheet.getRange(categoryStartRow, 13, categoryRowCount, 1), 'profitMargin');
  }

  sheet.setConditionalFormatRules(rules);
  Logger.log('✅ L1 条件付き書式適用: ' + rules.length + ' ルール');
}

/**
 * L2 カテゴリ分析に条件付き書式を適用
 *
 * 対象範囲:
 *   左側 月次推移: TACOS(7)/ACOS(8)/ROAS(9)/利益率(10)
 *   右側 ASIN別:  TACOS(20)/ACOS(21)/利益率(22)
 *
 * カテゴリごとにブロックが縦に並ぶため、ブロック情報配列から全範囲を組み立てる。
 *
 * @param {Sheet} sheet L2シート
 * @param {Array<Object>} blocks [{ startRow, monthRowCount, asinRowCount }]
 */
function applyL2ConditionalFormatting(sheet, blocks) {
  const rules = [];

  for (const b of blocks) {
    const dataRow = b.startRow + 2; // ヘッダーの次の行

    // 左側 月次推移
    if (b.monthRowCount > 0) {
      addKpiRules(rules, sheet.getRange(dataRow, 7, b.monthRowCount, 1), 'tacos');
      addKpiRules(rules, sheet.getRange(dataRow, 8, b.monthRowCount, 1), 'acos');
      addKpiRules(rules, sheet.getRange(dataRow, 9, b.monthRowCount, 1), 'roas');
      addKpiRules(rules, sheet.getRange(dataRow, 10, b.monthRowCount, 1), 'profitMargin');
    }

    // 右側 ASIN別（右ブロックは列12開始、TACOS=20, ACOS=21, 利益率=22）
    if (b.asinRowCount > 0) {
      addKpiRules(rules, sheet.getRange(dataRow, 20, b.asinRowCount, 1), 'tacos');
      addKpiRules(rules, sheet.getRange(dataRow, 21, b.asinRowCount, 1), 'acos');
      addKpiRules(rules, sheet.getRange(dataRow, 22, b.asinRowCount, 1), 'profitMargin');
    }
  }

  sheet.setConditionalFormatRules(rules);
  Logger.log('✅ L2 条件付き書式適用: ' + rules.length + ' ルール / ' + blocks.length + ' ブロック');
}

/**
 * テスト: 現在のダッシュボードの条件付き書式を目視確認
 * 既にダッシュボードが書き出されている状態で実行する
 */
function testApplyFormatting() {
  const l1 = getOrCreateSheet(SHEET_NAMES.L1_DASHBOARD);
  // 暫定: カテゴリ行数を推定して適用
  Logger.log('L1 シートの最終行: ' + l1.getLastRow());
  // 実運用では Dashboard.gs から自動呼び出し
}
