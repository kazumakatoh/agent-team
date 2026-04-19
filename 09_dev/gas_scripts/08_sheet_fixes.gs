/**
 * 各シートの追加修正
 * - 配信者スコアカード：レイアウト・RR比表記（1:X形式）
 * - 現物_スナップショット：USD/JPY切替、$表記削除
 * - FX_スナップショット：USD/JPY切替、E列に指標説明、$表記削除
 *
 * 重要：testWriteHoldingsToSheet / testWriteFXToSheet は sheet.clear() するため
 * これらの修正は月次更新 (runMonthlyUpdate) で毎回再適用される構造
 */

// ======================================================
// ① 配信者スコアカード 修正（V2）
// ======================================================
function applyScorecardFixesV2() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('配信者スコアカード');

  sheet.getRange('E1').setValue('GOLD_USD／TP Ladder／リスク5%');

  sheet.getRange('A1:A2').setHorizontalAlignment('right');
  sheet.getRange('D1').setHorizontalAlignment('right');
  sheet.getRange('B1').setHorizontalAlignment('left');
  sheet.getRange('E1').setHorizontalAlignment('left');
  sheet.getRange('A33').setHorizontalAlignment('left');
  sheet.getRange('B9:B30').setHorizontalAlignment('right');

  sheet.getRange('B18').setNumberFormat('"1:"0.00');
  sheet.setColumnWidths(2, 4, 134);

  Logger.log('✅ 配信者スコアカード修正完了');
}

// ======================================================
// ② 現物_スナップショット USD/JPY切替
// ======================================================
function addSpotSnapshotToggle() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('現物_スナップショット');

  sheet.getRange('D1').setValue('時価');
  sheet.getRange('E1').setValue('評価額');

  sheet.getRange('H1').setValue('通貨:').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange('I1').setValue('USD').setBackground('#fce5cd').setFontWeight('bold')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['USD','JPY'], true).build());
  sheet.getRange('J1').setValue('為替(¥/$):').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange('K1').setValue(155).setBackground('#fce5cd').setNumberFormat('0.00');

  Logger.log('✅ 現物スナップショット USD/JPY切替追加');
}

// ======================================================
// ③ FX_スナップショット USD/JPY切替 + E列説明
// ======================================================
function addFXSnapshotToggleAndExplanations() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('FX_スナップショット');

  sheet.getRange('B2').setValue('金額');

  sheet.getRange('H1').setValue('通貨:').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange('I1').setValue('USD').setBackground('#fce5cd').setFontWeight('bold')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['USD','JPY'], true).build());
  sheet.getRange('J1').setValue('為替(¥/$):').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange('K1').setValue(155).setBackground('#fce5cd').setNumberFormat('0.00');

  sheet.getRange('E9').setValue('💡 説明・判定').setFontWeight('bold')
    .setBackground('#cfe2f3').setHorizontalAlignment('center');
  sheet.getRange('E10').setFormula('=B10&"回の取引（勝ち"&ROUND(B10*B11,0)&"回、負け"&(B10-ROUND(B10*B11,0))&"回）"');
  sheet.getRange('E11').setFormula('=TEXT(B11,"0.0%")&"　損益分岐勝率"&TEXT(B20,"0.0%")&"と比較"');
  sheet.getRange('E12').setValue('PF = 総利益÷総損失　<1.0:赤字 / 1.0-1.5:黒字 / 1.5-2.0:優秀 / 2.0+:非常に優秀');
  sheet.getRange('E13').setFormula('="現在 1:"&TEXT(B13,"0.00")&"　目標 1:2（損益分岐勝率34%以上、3回に1回勝てば収支トントン）"');
  sheet.getRange('E14').setValue('1トレードあたり平均損益。プラスなら継続、マイナスなら戦略見直し');
  sheet.getRange('E18').setValue('高値からの最大下落額（$）');
  sheet.getRange('E19').setValue('資金に対する下落率。<10%安全 / 10-15%注意 / 15%+危険');
  sheet.getRange('E20').setValue('この勝率を超えないと長期的に赤字になる');

  sheet.getRange('B13').setNumberFormat('"1:"0.00');

  sheet.getRange('E10:E20').setFontSize(10).setWrap(false).setFontColor('#666666');
  sheet.setColumnWidth(5, 110);

  Logger.log('✅ FXスナップショット：切替＋説明追加完了');
}

// ======================================================
// ④ 配信者スコアカード RR比目標更新
// ======================================================
function updateScorecardRRTarget() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('配信者スコアカード');

  sheet.getRange('B18').setNumberFormat('"1:"0.00');
  sheet.getRange('C18').setValue('1:2');
  sheet.getRange('F18').setValue('目標 1:2（損益分岐勝率34%以上、3回に1回勝てば収支トントン）')
    .setFontSize(9).setFontColor('#666666');

  Logger.log('✅ スコアカード RR目標更新完了');
}

// ======================================================
// ⑤ FXスナップショット E列の追加調整
// ======================================================
function updateFXSnapshotExplanations() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('FX_スナップショット');

  sheet.getRange('E13').setFormula(
    '="現在 1:"&TEXT(B13,"0.00")&"　目標 1:2（損益分岐勝率34%以上、3回に1回勝てば収支トントン）"'
  );
  sheet.setColumnWidth(5, 110);
  sheet.getRange('E10:E20').setWrap(false);

  Logger.log('✅ FXスナップショット 説明文更新完了');
}

// ======================================================
// ⑥ onEdit ハンドラ（現物・FX）USD/JPY切替
// ======================================================
function onEditSpotSnapshot(e) {
  if (!e || !e.range) return;
  if (e.range.getA1Notation() !== 'I1') return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== '現物_スナップショット') return;

  const currency = e.range.getValue();
  const rate = Number(sheet.getRange('K1').getValue()) || 155;
  const lastRow = sheet.getLastRow();

  [4, 5].forEach(col => {
    for (let r = 2; r <= lastRow; r++) {
      const cell = sheet.getRange(r, col);
      const v = Number(cell.getValue());
      if (isNaN(v) || v === 0) continue;
      const fmt = cell.getNumberFormat();
      const isJPY = fmt.indexOf('¥') >= 0;

      if (currency === 'JPY' && !isJPY) {
        cell.setValue(v * rate).setNumberFormat('¥#,##0');
      } else if (currency === 'USD' && isJPY) {
        cell.setValue(v / rate).setNumberFormat(col === 4 ? '$#,##0.0000' : '$#,##0.00');
      }
    }
  });
}

function onEditFXSnapshot(e) {
  if (!e || !e.range) return;
  if (e.range.getA1Notation() !== 'I1') return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== 'FX_スナップショット') return;

  const currency = e.range.getValue();
  const rate = Number(sheet.getRange('K1').getValue()) || 155;
  const moneyCells = ['B3','B4','B5','B6','B7','B14','B15','B16','B17','B18'];

  moneyCells.forEach(addr => {
    const cell = sheet.getRange(addr);
    const v = Number(cell.getValue());
    if (isNaN(v) || v === 0) return;
    const fmt = cell.getNumberFormat();
    const isJPY = fmt.indexOf('¥') >= 0;

    if (currency === 'JPY' && !isJPY) {
      cell.setValue(v * rate).setNumberFormat('¥#,##0');
    } else if (currency === 'USD' && isJPY) {
      cell.setValue(v / rate).setNumberFormat('#,##0');
    }
  });
}

// ======================================================
// ⑦ トリガー一括セットアップ
// ======================================================
function setupSnapshotTriggers() {
  const ss = SpreadsheetApp.getActive();

  ScriptApp.getProjectTriggers()
    .filter(t => ['onEditSpotSnapshot', 'onEditFXSnapshot'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('onEditSpotSnapshot').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('onEditFXSnapshot').forSpreadsheet(ss).onEdit().create();

  Logger.log('✅ スナップショット編集トリガー再登録完了');
}
