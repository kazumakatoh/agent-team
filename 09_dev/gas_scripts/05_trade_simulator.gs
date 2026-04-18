/**
 * 取引シミュレーター
 * 銘柄：XAUUSD（ゴールド）/ BigBoss / レバレッジ2222倍
 *
 * 用途：
 *   - 取引値・SL・TP・ロット数を変更したリスクシミュレーション
 *   - 限界ロット数（最大リスク5%基準）の確認
 *   - 安全性判定（🟢🟡🔴）
 */

const SIM_CONFIG = {
  SHEET: '取引シミュレーター',
  GOLD_CONTRACT_SIZE: 100,    // 1ロット = 100oz
  PIP_VALUE_PER_LOT: 10,      // ゴールド: $10/pip/lot
  LEVERAGE: 2222,
  WIN_RATE: 0.60,             // 想定勝率
  MAX_RISK_PCT: 0.05,         // 限界ロット計算用最大リスク
  DEFAULT_RATE: 155
};

// ======================================================
// 構築
// ======================================================
function buildTradeSimulator() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(SIM_CONFIG.SHEET);
  if (!sheet) sheet = ss.insertSheet(SIM_CONFIG.SHEET);
  else sheet.clear();

  // Row 1: 通貨切替・現在情報
  sheet.getRange('A1').setValue('通貨:').setFontWeight('bold');
  sheet.getRange('B1').setValue('USD').setBackground('#fce5cd').setFontWeight('bold')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['USD','JPY'], true).build());
  sheet.getRange('C1').setValue('為替(¥/$):').setFontWeight('bold');
  sheet.getRange('D1').setValue(SIM_CONFIG.DEFAULT_RATE).setBackground('#fce5cd').setNumberFormat('0.00');
  sheet.getRange('E1').setValue('現在Equity($):').setFontWeight('bold');
  sheet.getRange('F1').setFormula("='FX_スナップショット'!B6").setBackground('#d9ead3');
  sheet.getRange('G1').setValue('現在Gold価格:').setFontWeight('bold');
  sheet.getRange('H1').setValue(4800).setBackground('#fce5cd');

  // 現状セクション
  setSection_(sheet, 3, '🔍 現状');
  const currentRows = [
    ['銘柄', 'XAUUSD（ゴールド/USD）'],
    ['取引所', 'BigBoss'],
    ['レバレッジ', SIM_CONFIG.LEVERAGE + '倍'],
    ['勝率想定', SIM_CONFIG.WIN_RATE],
    ['限界取引ロット数', '=ROUNDDOWN($F$1*' + SIM_CONFIG.MAX_RISK_PCT + '/(50*' + SIM_CONFIG.PIP_VALUE_PER_LOT + '),2)'],
    ['限界ロット説明', '=B8&"ロット：1ドル動くと約"&TEXT(B8*' + SIM_CONFIG.GOLD_CONTRACT_SIZE + '*$D$1,"#,##0")&"円動く"'],
    ['必要証拠金', "=IFERROR('FX_スナップショット'!B8, 0)"],
    ['余剰証拠金', "='FX_スナップショット'!B7"],
    ['証拠金維持率', '=IF(B10>0, $F$1/B10, 0)'],
    ['耐えられる値幅', '=IF(B8>0, $F$1/(B8*' + SIM_CONFIG.PIP_VALUE_PER_LOT + '), 0)']
  ];
  sheet.getRange(4, 1, currentRows.length, 2).setValues(currentRows);

  // 入力値セクション
  setSection_(sheet, 15, '✏️ 入力値（変更してシミュレーション）');
  const inputRows = [
    ['取引値', 4800],
    ['SL（ストップロス）', 4795],
    ['TP（利確）', 4810],
    ['ロット数', 0.30]
  ];
  sheet.getRange(16, 1, inputRows.length, 2).setValues(inputRows);
  sheet.getRange(16, 2, 4, 1).setBackground('#fff2cc');

  // リスク管理指標セクション
  setSection_(sheet, 21, '📊 リスク管理指標（自動計算）');
  const riskRows = [
    ['必要証拠金', '=B19*' + SIM_CONFIG.GOLD_CONTRACT_SIZE + '*B16/' + SIM_CONFIG.LEVERAGE],
    ['余剰証拠金', '=$F$1-B22'],
    ['証拠金維持率', '=IF(B22>0, $F$1/B22, 0)'],
    ['獲得値幅（$）', '=ABS(B18-B16)*B19*' + SIM_CONFIG.GOLD_CONTRACT_SIZE],
    ['獲得値幅（pips）', '=ABS(B18-B16)*10'],
    ['損失値幅（$）', '=ABS(B16-B17)*B19*' + SIM_CONFIG.GOLD_CONTRACT_SIZE],
    ['損失値幅（pips）', '=ABS(B16-B17)*10'],
    ['リスクリワード比', '=IFERROR(B25/B27, 0)']
  ];
  sheet.getRange(22, 1, riskRows.length, 2).setValues(riskRows);

  // 安全性判定
  setSection_(sheet, 31, '⚖️ 安全性判定');
  sheet.getRange(32, 1).setValue('現状との比較');
  sheet.getRange(32, 2).setFormula(
    '=IF(B19>B8, "🔴 限界ロット超過："&B19&" > "&B8, ' +
    'IF(B24<2, "🔴 維持率2倍未満は危険", ' +
    'IF(B29<0.7, "🟡 RR比0.7未満：薄利", ' +
    '"🟢 安全圏")))'
  );

  applySimulatorFormats_(sheet);
  Logger.log('✅ 取引シミュレーター構築完了');
}

function setSection_(sheet, row, label) {
  sheet.getRange(row, 1, 1, 8).setBackground('#37474f').setFontColor('white').setFontWeight('bold');
  sheet.getRange(row, 1).setValue(label);
}

function applySimulatorFormats_(sheet) {
  const lastRow = sheet.getLastRow();
  sheet.getRange(1, 1, lastRow, 8).setFontSize(10);
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidths(3, 6, 110);
  sheet.getRange('A:A').setHorizontalAlignment('center');
  sheet.getRange('B4:B6').setHorizontalAlignment('center');
  sheet.getRange('B7').setNumberFormat('0.0%');
  sheet.getRange('B8').setNumberFormat('0.00');
  sheet.getRange('B19').setNumberFormat('0.00');
  ['F1','B10','B11','B22','B23','B25','B27'].forEach(a => sheet.getRange(a).setNumberFormat('$#,##0'));
  ['H1','B16','B17','B18'].forEach(a => sheet.getRange(a).setNumberFormat('$#,##0'));
  ['B26','B28'].forEach(a => sheet.getRange(a).setNumberFormat('#,##0" pips"'));
  sheet.getRange('B12').setNumberFormat('0.0"倍"');
  sheet.getRange('B24').setNumberFormat('0.0"倍"');
  sheet.getRange('B13').setNumberFormat('#,##0" pips"');
  sheet.getRange('B29').setNumberFormat('0.00');
}

// ======================================================
// 修正適用（B9→C8移動・行9削除・短縮ラベル・B列幅統一）
// ======================================================
function applySimulatorFixes() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SIM_CONFIG.SHEET);

  sheet.getRange('B4').setValue('GOLD_USD');
  const b9Formula = sheet.getRange('B9').getFormula();
  if (b9Formula) sheet.getRange('C8').setFormula(b9Formula);
  sheet.deleteRow(9);

  sheet.getRange('A14').setValue('✏️ 入力値');
  sheet.getRange('A20').setValue('📊 リスク管理指標');

  sheet.setColumnWidth(2, 110);
  sheet.getRange('1:1').setHorizontalAlignment('right');
  sheet.getRange('B:B').setHorizontalAlignment('right');

  Logger.log('✅ シミュレーター修正完了');
}

// ======================================================
// 円/ドル切替（onEdit）
// ======================================================
function onEditSimulator(e) {
  if (!e || e.range.getA1Notation() !== 'B1') return;
  if (e.range.getSheet().getName() !== SIM_CONFIG.SHEET) return;

  const sheet = e.range.getSheet();
  const currency = e.range.getValue();
  // 行9削除後の参照
  const moneyCells = ['F1','H1','B9','B10','B15','B16','B17','B21','B22','B24','B26'];
  const fmt = currency === 'JPY' ? '¥#,##0' : '$#,##0';
  moneyCells.forEach(a => sheet.getRange(a).setNumberFormat(fmt));
}
