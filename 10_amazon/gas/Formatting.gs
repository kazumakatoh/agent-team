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
 *   全体サマリー（行3-7）: 広告比率(8)/ROAS(9)/利益率(10)
 *     ※ 粗利率/原価率/販売手数料/経費等は P&L セクションへ集約したため削除
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
  const overallRows = 5;
  const overallStartRow = 3;

  // 広告比率=TACOS（列8）
  addKpiRules(rules, sheet.getRange(overallStartRow, 8, overallRows, 1), 'tacos');
  // ROAS（列9）
  addKpiRules(rules, sheet.getRange(overallStartRow, 9, overallRows, 1), 'roas');
  // 利益率（列10）
  addKpiRules(rules, sheet.getRange(overallStartRow, 10, overallRows, 1), 'profitMargin');

  // カテゴリ別サマリー（17列構成）
  // 列: カテゴリ(1), 売上(2), 売上比(3), CV(4), 点数(5), 仕入(6), 販売手数料(7),
  //     広告費(8), 経費等(9), 利益(10), 広告比率(11), 粗利率(12), 利益率(13), ROAS(14),
  //     TACOS(15), ACOS(16), 前月比(17)
  if (categoryRowCount > 0) {
    addKpiRules(rules, sheet.getRange(categoryStartRow, 11, categoryRowCount, 1), 'tacos');     // 広告比率
    addKpiRules(rules, sheet.getRange(categoryStartRow, 12, categoryRowCount, 1), 'profitMargin'); // 粗利率
    addKpiRules(rules, sheet.getRange(categoryStartRow, 13, categoryRowCount, 1), 'profitMargin'); // 利益率
    addKpiRules(rules, sheet.getRange(categoryStartRow, 14, categoryRowCount, 1), 'roas');      // ROAS
    addKpiRules(rules, sheet.getRange(categoryStartRow, 15, categoryRowCount, 1), 'tacos');     // TACOS
    addKpiRules(rules, sheet.getRange(categoryStartRow, 16, categoryRowCount, 1), 'acos');      // ACOS
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

    // 左側 月次推移（13列: 年月/売上/CV/点数/仕入/販売手数料/広告費/経費等/利益/TACOS/ACOS/ROAS/利益率）
    if (b.monthRowCount > 0) {
      addKpiRules(rules, sheet.getRange(dataRow, 10, b.monthRowCount, 1), 'tacos');      // TACOS
      addKpiRules(rules, sheet.getRange(dataRow, 11, b.monthRowCount, 1), 'acos');       // ACOS
      addKpiRules(rules, sheet.getRange(dataRow, 12, b.monthRowCount, 1), 'roas');       // ROAS
      addKpiRules(rules, sheet.getRange(dataRow, 13, b.monthRowCount, 1), 'profitMargin'); // 利益率
    }

    // 右側 ASIN別（列15開始、14列スパン: ASIN/商品名/売上/売上比/CV/点数/仕入/販売手数料/広告費/経費等/利益/TACOS/ACOS/利益率）
    // 絶対列番号: TACOS=26(15+11), ACOS=27(15+12), 利益率=28(15+13)
    if (b.asinRowCount > 0) {
      addKpiRules(rules, sheet.getRange(dataRow, 26, b.asinRowCount, 1), 'tacos');
      addKpiRules(rules, sheet.getRange(dataRow, 27, b.asinRowCount, 1), 'acos');
      addKpiRules(rules, sheet.getRange(dataRow, 28, b.asinRowCount, 1), 'profitMargin');
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
