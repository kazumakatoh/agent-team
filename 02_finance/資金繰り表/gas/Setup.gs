/**
 * 資金繰り表 セットアップスクリプト
 * 日本政策金融公庫フォーマット準拠（簡易レイアウト版）
 *
 * 使い方：
 * 1. 新規Googleスプレッドシートを作成
 * 2. 拡張機能 > Apps Script にこのコードを貼り付け
 * 3. setupAllSheets() を実行
 */

/** 作成するシート（年度 = 決算年） */
const SHEETS_TO_CREATE = [2025, 2026];

/** 月の並び（決算月=2月 → 3月始まり） */
const MONTHS = ['3月','4月','5月','6月','7月','8月','9月','10月','11月','12月','1月','2月'];

/**
 * 全シートをセットアップ
 */
function setupAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  SHEETS_TO_CREATE.forEach(function(year) {
    const sheetName = '資金繰り表_' + year;
    let sheet = ss.getSheetByName(sheetName);

    if (sheet) {
      sheet.clear();
      sheet.clearFormats();
    } else {
      sheet = ss.insertSheet(sheetName);
    }

    buildSheet_(sheet, year);
  });

  // デフォルトの「シート1」があれば削除
  const defaultSheet = ss.getSheetByName('シート1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  SpreadsheetApp.flush();
  Logger.log('セットアップ完了');
}

/**
 * 1つの年度シートを構築
 *
 * レイアウト:
 *   A列: カテゴリ  B列: 項目名  C〜N列: 3月〜2月  O列: 合計
 *
 *   行1:  (単位)
 *   行2:  ヘッダー
 *   行3:  売上（今期）
 *   行4:  売上（前期）
 *   行5:  前月繰越金
 *   行6:  ── 経常収支 ──（セクションヘッダー）
 *   行7:  収入（小計）
 *   行8:    現金売上
 *   行9:    売掛金回収
 *   行10: 支出（小計）
 *   行11:   現金仕入
 *   行12:   買掛金支払
 *   行13:   人件費
 *   行14:   商品棚卸高
 *   行15:   諸経費
 *   行16: 差引過不足
 *   行17: ── 経常外収支 ──（セクションヘッダー）
 *   行18: 経常外収入
 *   行19: 経常外支出
 *   行20: 合計
 *   行21: ── 財務収支 ──（セクションヘッダー）
 *   行22: 収入（小計）
 *   行23:   日本政策金融公庫
 *   行24:   西武信用金庫
 *   行25: 支出（小計）
 *   行26:   借入金返済（短期）
 *   行27:   借入金返済（長期）
 *   行28: 合計
 *   行29: 翌月繰越金
 */
function buildSheet_(sheet, year) {
  // --- 列幅設定 ---
  sheet.setColumnWidth(1, 40);   // A: カテゴリ
  sheet.setColumnWidth(2, 120);  // B: 項目名
  for (var c = 3; c <= 14; c++) {
    sheet.setColumnWidth(c, 75); // C〜N: 月データ
  }
  sheet.setColumnWidth(15, 85);  // O: 合計

  // 月データ列: C=3, D=4, ... N=14
  // 合計列: O=15
  var mStart = 3;  // 月データ開始列
  var mEnd = 14;   // 月データ終了列
  var colTotal = 15; // 合計列

  // --- 行1: 単位 ---
  sheet.getRange('O1').setValue('単位：千円').setHorizontalAlignment('right').setFontSize(9);

  // --- 行2: ヘッダー ---
  var headerValues = [['', '資金繰り表', ...MONTHS, '合計']];
  sheet.getRange('A2:O2').setValues(headerValues);
  sheet.getRange('A2:O2')
    .setBackground('#006060')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setFontSize(10);

  // --- 行3: 売上（今期） ---
  setRow_(sheet, 3, '', '売上（今期）', 'SUM');

  // --- 行4: 売上（前期） ---
  setRow_(sheet, 4, '', '売上（前期）', 'SUM');

  // --- 行5: 前月繰越金 ---
  setRow_(sheet, 5, '', '前月繰越金', 'NONE');
  sheet.getRange('O5').setValue('-');

  // === 行6: 経常収支（セクションヘッダー） ===
  setSectionHeader_(sheet, 6, '経常収支');

  // --- 行7: 収入（小計）= 行8 + 行9 ---
  setRow_(sheet, 7, '', '収入', 'SUM');
  setSubtotalFormulas_(sheet, 7, [8, 9], mStart, mEnd);
  sheet.getRange('B7').setFontWeight('bold');
  sheet.getRange('A7:O7').setBackground('#E0F0E0');

  // --- 行8: 現金売上 ---
  setRow_(sheet, 8, '', '現金売上', 'SUM');

  // --- 行9: 売掛金回収 ---
  setRow_(sheet, 9, '', '売掛金回収', 'SUM');

  // --- 行10: 支出（小計）= 行11〜15 ---
  setRow_(sheet, 10, '', '支出', 'SUM');
  setSubtotalFormulas_(sheet, 10, [11, 12, 13, 14, 15], mStart, mEnd);
  sheet.getRange('B10').setFontWeight('bold');
  sheet.getRange('A10:O10').setBackground('#E0F0E0');

  // --- 行11: 現金仕入 ---
  setRow_(sheet, 11, '', '現金仕入', 'SUM');

  // --- 行12: 買掛金支払 ---
  setRow_(sheet, 12, '', '買掛金支払', 'SUM');

  // --- 行13: 人件費 ---
  setRow_(sheet, 13, '', '人件費', 'SUM');

  // --- 行14: 商品棚卸高 ---
  setRow_(sheet, 14, '', '商品棚卸高', 'SUM');

  // --- 行15: 諸経費 ---
  setRow_(sheet, 15, '', '諸経費', 'SUM');

  // --- 行16: 差引過不足 = 収入(7) - 支出(10) ---
  setRow_(sheet, 16, '', '差引過不足', 'SUM');
  setDiffFormulas_(sheet, 16, 7, 10, mStart, mEnd);
  sheet.getRange('B16').setFontWeight('bold');
  sheet.getRange('A16:O16').setBackground('#E0F0E0');

  // === 行17: 経常外収支（セクションヘッダー） ===
  setSectionHeader_(sheet, 17, '経常外収支');

  // --- 行18: 経常外収入 ---
  setRow_(sheet, 18, '', '経常外収入', 'SUM');

  // --- 行19: 経常外支出 ---
  setRow_(sheet, 19, '', '経常外支出', 'SUM');

  // --- 行20: 合計 = 行18 - 行19 ---
  setRow_(sheet, 20, '', '合計', 'SUM');
  setDiffFormulas_(sheet, 20, 18, 19, mStart, mEnd);
  sheet.getRange('B20').setFontWeight('bold');
  sheet.getRange('A20:O20').setBackground('#E0F0E0');

  // === 行21: 財務収支（セクションヘッダー） ===
  setSectionHeader_(sheet, 21, '財務収支');

  // --- 行22: 収入（小計）= 行23 + 行24 ---
  setRow_(sheet, 22, '', '収入', 'SUM');
  setSubtotalFormulas_(sheet, 22, [23, 24], mStart, mEnd);
  sheet.getRange('B22').setFontWeight('bold');
  sheet.getRange('A22:O22').setBackground('#E0F0E0');

  // --- 行23: 日本政策金融公庫 ---
  setRow_(sheet, 23, '', '日本政策金融公庫', 'SUM');

  // --- 行24: 西武信用金庫 ---
  setRow_(sheet, 24, '', '西武信用金庫', 'SUM');

  // --- 行25: 支出（小計）= 行26 + 行27 ---
  setRow_(sheet, 25, '', '支出', 'SUM');
  setSubtotalFormulas_(sheet, 25, [26, 27], mStart, mEnd);
  sheet.getRange('B25').setFontWeight('bold');
  sheet.getRange('A25:O25').setBackground('#E0F0E0');

  // --- 行26: 借入金返済（短期） ---
  setRow_(sheet, 26, '', '借入金返済（短期）', 'SUM');

  // --- 行27: 借入金返済（長期） ---
  setRow_(sheet, 27, '', '借入金返済（長期）', 'SUM');

  // --- 行28: 合計 = 収入(22) - 支出(25) ---
  setRow_(sheet, 28, '', '合計', 'SUM');
  setDiffFormulas_(sheet, 28, 22, 25, mStart, mEnd);
  sheet.getRange('B28').setFontWeight('bold');
  sheet.getRange('A28:O28').setBackground('#E0F0E0');

  // --- 行29: 翌月繰越金 = 前月繰越金(5) + 差引過不足(16) + 経常外合計(20) + 財務合計(28) ---
  setRow_(sheet, 29, '', '翌月繰越金', 'NONE');
  for (var c = mStart; c <= mEnd; c++) {
    var col = colLetter_(c);
    sheet.getRange(col + '29').setFormula(
      '=' + col + '5+' + col + '16+' + col + '20+' + col + '28'
    );
  }
  sheet.getRange('B29').setFontWeight('bold');
  sheet.getRange('A29:O29').setBackground('#FFF8E1');

  // --- 前月繰越金の自動連携（4月以降 = 前月の翌月繰越金） ---
  for (var c = mStart + 1; c <= mEnd; c++) {
    var prevCol = colLetter_(c - 1);
    var col = colLetter_(c);
    sheet.getRange(col + '5').setFormula('=' + prevCol + '29');
  }

  // --- 前月繰越金の行を強調 ---
  sheet.getRange('A5:O5').setBackground('#FFF8E1');
  sheet.getRange('B5').setFontWeight('bold');

  // --- 書式設定 ---
  applyFormatting_(sheet, mStart, mEnd, colTotal);
}


// ==============================
// ヘルパー関数
// ==============================

/**
 * セクションヘッダー行（経常収支/経常外収支/財務収支）
 */
function setSectionHeader_(sheet, row, label) {
  sheet.getRange('A' + row + ':O' + row).merge();
  sheet.getRange('A' + row)
    .setValue('　　' + label)
    .setBackground('#555555')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(10);
  // merge後の背景色を全列に適用
  sheet.getRange('A' + row + ':O' + row).setBackground('#555555');
}

/**
 * 行にラベルと合計数式をセット
 */
function setRow_(sheet, row, category, label, type) {
  if (category) sheet.getRange('A' + row).setValue(category);
  sheet.getRange('B' + row).setValue(label);

  if (type === 'SUM') {
    sheet.getRange('O' + row).setFormula('=SUM(C' + row + ':N' + row + ')');
  }
}

/**
 * 小計行の数式（複数行の合計）
 */
function setSubtotalFormulas_(sheet, totalRow, sourceRows, mStart, mEnd) {
  for (var c = mStart; c <= mEnd; c++) {
    var col = colLetter_(c);
    var formula = '=' + sourceRows.map(function(r) { return col + r; }).join('+');
    sheet.getRange(col + totalRow).setFormula(formula);
  }
}

/**
 * 差引行の数式（row1 - row2）
 */
function setDiffFormulas_(sheet, resultRow, row1, row2, mStart, mEnd) {
  for (var c = mStart; c <= mEnd; c++) {
    var col = colLetter_(c);
    sheet.getRange(col + resultRow).setFormula('=' + col + row1 + '-' + col + row2);
  }
}

/**
 * 列番号 → 列文字
 */
function colLetter_(colNum) {
  return String.fromCharCode(64 + colNum);
}

/**
 * 書式設定
 */
function applyFormatting_(sheet, mStart, mEnd, colTotal) {
  var lastCol = colLetter_(colTotal); // O
  var dataRange = sheet.getRange('A1:' + lastCol + '29');

  // 数値フォーマット（カンマ区切り・マイナスは▲表示）
  sheet.getRange('C3:' + lastCol + '29').setNumberFormat('#,##0;▲#,##0');

  // 全体フォント
  dataRange.setFontFamily('游ゴシック');
  dataRange.setFontSize(10);

  // 数値セルを右寄せ
  sheet.getRange('C3:' + lastCol + '29').setHorizontalAlignment('right');

  // 項目名を左寄せ
  sheet.getRange('B3:B29').setHorizontalAlignment('left');

  // 罫線
  dataRange.setBorder(true, true, true, true, true, true,
    '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  // 行の高さ
  for (var r = 2; r <= 29; r++) {
    sheet.setRowHeight(r, 22);
  }
}
