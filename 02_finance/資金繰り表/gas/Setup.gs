/**
 * 資金繰り表 セットアップスクリプト
 * 日本政策金融公庫フォーマット準拠
 *
 * 使い方：
 * 1. 新規Googleスプレッドシートを作成
 * 2. 拡張機能 > Apps Script にこのコードを貼り付け
 * 3. setupAllSheets() を実行
 */

// ==============================
// 設定
// ==============================

/** 作成するシート（年度 = 決算年） */
const SHEETS_TO_CREATE = [2025, 2026];

/** 月の並び（決算月=2月 → 3月始まり） */
const MONTHS = ['3月','4月','5月','6月','7月','8月','9月','10月','11月','12月','1月','2月'];

/** 行ラベル定義 */
const ROW_LABELS = {
  HEADER:           '資金繰り表',
  SALES:            '今期売上',
  PREV_SALES:       '前期売上',
  CARRY_FORWARD:    '前月繰越金（A）',
  CASH_SALES:       '現金売上',
  AR_COLLECTION:    '売掛金回収',
  OTHER_INCOME:     'その他',
  INCOME_TOTAL:     '経常収入計（B）',
  CASH_PURCHASE:    '現金仕入',
  AP_PAYMENT:       '買掛金支払',
  PERSONNEL:        '人件費',
  OTHER_EXPENSE:    'その他',
  INVENTORY:        '商品棚卸高',
  MISC_EXPENSE:     '諸経費',
  EXPENSE_TOTAL:    '経常支出計（C）',
  OPERATING_DIFF:   '差引過不足（D）=（B）-（C）',
  NON_OP_INCOME:    '経常外収入',
  NON_OP_EXPENSE:   '経常外支出',
  NON_OP_TOTAL:     '経常外収支計（E）',
  LOAN_JFC:         '日本政策金融公庫',
  LOAN_SEIBU:       '西武信用金庫',
  LOAN_OTHER:       'その他',
  REPAY_SHORT:      '借入金返済（短期）',
  REPAY_LONG:       '借入金返済（長期）',
  FINANCE_TOTAL:    '財務収支計（F）',
  NEXT_CARRY:       '翌月繰越金（G）=（A）+（D）+（E）+（F）'
};

/** カテゴリラベル（A〜E列のセル結合用） */
const CATEGORIES = {
  INCOME:    { label: '経常収入', startRow: 6, endRow: 9, col: 'A' },
  EXPENSE:   { label: '経常支出', startRow: 10, endRow: 16, col: 'A' },
  NON_OP:    { label: '経常外収支', startRow: 18, endRow: 20, col: 'A' },
  FINANCE:   { label: '財務収支', startRow: 21, endRow: 26, col: 'A' }
};

const SUB_CATEGORIES = {
  IN:        { label: '収入', startRow: 6, endRow: 8 },
  OUT_LABEL: { label: '支出', startRow: 10, endRow: 15 }
};


// ==============================
// メイン
// ==============================

/**
 * 全シートをセットアップ
 */
function setupAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  SHEETS_TO_CREATE.forEach(function(year) {
    const sheetName = '資金繰り表_' + year;
    let sheet = ss.getSheetByName(sheetName);

    if (sheet) {
      // 既存シートをクリア
      sheet.clear();
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


// ==============================
// シート構築
// ==============================

/**
 * 1つの年度シートを構築
 */
function buildSheet_(sheet, year) {
  const period = year - 2018; // 第N期

  // --- 列幅設定 ---
  sheet.setColumnWidth(1, 30);   // A: カテゴリ
  sheet.setColumnWidth(2, 30);   // B: サブカテゴリ
  sheet.setColumnWidth(3, 30);   // C: サブ2
  sheet.setColumnWidth(4, 30);   // D: サブ3
  sheet.setColumnWidth(5, 140);  // E: 項目名
  for (var c = 6; c <= 17; c++) {
    sheet.setColumnWidth(c, 80); // F〜Q: 月データ
  }
  sheet.setColumnWidth(18, 90);  // R: 合計
  sheet.setColumnWidth(19, 90);  // S: 月平均

  // --- 行1: 単位 ---
  sheet.getRange('S1').setValue('単位：千円');
  sheet.getRange('S1').setHorizontalAlignment('right');

  // --- 行2: ヘッダー ---
  var headerRange = sheet.getRange('A2:S2');
  var headerValues = [['', '', '', '', '資金繰り表', ...MONTHS, '合計', '月平均']];
  headerRange.setValues(headerValues);
  sheet.getRange('A2:S2').merge(); // A2:D2は項目名と結合しない
  sheet.getRange('E2:S2').setBackground('#006060')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  // E2だけ左寄せ
  sheet.getRange('E2').setHorizontalAlignment('left');
  // A2:S2全体に背景色
  sheet.getRange('A2:D2').setBackground('#006060');

  // --- 行3: 今期売上 ---
  setRow_(sheet, 3, ROW_LABELS.SALES, 'SUM_AVG');

  // --- 行4: 前期売上 ---
  setRow_(sheet, 4, ROW_LABELS.PREV_SALES, 'SUM_AVG');

  // --- 行5: 前月繰越金(A) ---
  setRow_(sheet, 5, ROW_LABELS.CARRY_FORWARD, 'NONE');
  sheet.getRange('R5').setValue('-');
  sheet.getRange('S5').setValue('-');

  // === 経常収入 ===
  // 行6: 現金売上
  setRow_(sheet, 6, ROW_LABELS.CASH_SALES, 'SUM_AVG');
  setCategoryLabel_(sheet, 6, '収', 'B');

  // 行7: 売掛金回収
  setRow_(sheet, 7, ROW_LABELS.AR_COLLECTION, 'SUM_AVG');

  // 行8: その他
  setRow_(sheet, 8, ROW_LABELS.OTHER_INCOME, 'SUM_AVG');
  setCategoryLabel_(sheet, 8, '入', 'B');

  // 行9: 経常収入計(B) = 行6+7+8
  setRow_(sheet, 9, ROW_LABELS.INCOME_TOTAL, 'SUM_AVG');
  setCategoryLabel_(sheet, 9, '経', 'A');
  setSubtotalFormulas_(sheet, 9, [6, 7, 8]);
  sheet.getRange('E9').setFontWeight('bold');
  setCategoryLabel_(sheet, 9, '計', 'D');

  // === 経常支出 ===
  // 行10: 現金仕入
  setRow_(sheet, 10, ROW_LABELS.CASH_PURCHASE, 'SUM_AVG');
  setCategoryLabel_(sheet, 10, '常', 'A');

  // 行11: 買掛金支払
  setRow_(sheet, 11, ROW_LABELS.AP_PAYMENT, 'SUM_AVG');

  // 行12: 人件費
  setRow_(sheet, 12, ROW_LABELS.PERSONNEL, 'SUM_AVG');
  setCategoryLabel_(sheet, 12, '収', 'B');

  // 行13: その他
  setRow_(sheet, 13, ROW_LABELS.OTHER_EXPENSE, 'SUM_AVG');

  // 行14: 商品棚卸高
  setRow_(sheet, 14, ROW_LABELS.INVENTORY, 'SUM_AVG');
  setCategoryLabel_(sheet, 14, '出', 'C');

  // 行15: 諸経費
  setRow_(sheet, 15, ROW_LABELS.MISC_EXPENSE, 'SUM_AVG');

  // 行16: 経常支出計(C) = 行10〜15
  setRow_(sheet, 16, ROW_LABELS.EXPENSE_TOTAL, 'SUM_AVG');
  setSubtotalFormulas_(sheet, 16, [10, 11, 12, 13, 14, 15]);
  sheet.getRange('E16').setFontWeight('bold');
  setCategoryLabel_(sheet, 16, '支', 'B');
  setCategoryLabel_(sheet, 16, '計', 'D');

  // --- 行17: 差引過不足(D) = (B) - (C) ---
  setRow_(sheet, 17, ROW_LABELS.OPERATING_DIFF, 'SUM_AVG');
  setDiffFormulas_(sheet, 17, 9, 16);
  sheet.getRange('E17').setFontWeight('bold');

  // === 経常外収支 ===
  // 行18: 経常外収入
  setRow_(sheet, 18, ROW_LABELS.NON_OP_INCOME, 'SUM_AVG');
  setCategoryLabel_(sheet, 18, '経', 'A');
  setCategoryLabel_(sheet, 18, '常', 'B');

  // 行19: 経常外支出
  setRow_(sheet, 19, ROW_LABELS.NON_OP_EXPENSE, 'SUM_AVG');
  setCategoryLabel_(sheet, 19, '外', 'C');

  // 行20: 経常外収支計(E) = 行18 - 行19
  setRow_(sheet, 20, ROW_LABELS.NON_OP_TOTAL, 'SUM_AVG');
  setDiffFormulas_(sheet, 20, 18, 19);
  sheet.getRange('E20').setFontWeight('bold');
  setCategoryLabel_(sheet, 20, '収', 'A');
  setCategoryLabel_(sheet, 20, '支', 'B');
  setCategoryLabel_(sheet, 20, '計', 'D');

  // === 財務収支 ===
  // 行21: 日本政策金融公庫（手入力）
  setRow_(sheet, 21, ROW_LABELS.LOAN_JFC, 'SUM_ONLY');
  setCategoryLabel_(sheet, 21, '財', 'A');
  setCategoryLabel_(sheet, 21, '収', 'C');

  // 行22: 西武信用金庫（手入力）
  setRow_(sheet, 22, ROW_LABELS.LOAN_SEIBU, 'SUM_ONLY');
  setCategoryLabel_(sheet, 22, '入', 'C');

  // 行23: その他（手入力）
  setRow_(sheet, 23, ROW_LABELS.LOAN_OTHER, 'SUM_ONLY');
  setCategoryLabel_(sheet, 23, '務', 'A');

  // 行24: 借入金返済（短期）
  setRow_(sheet, 24, ROW_LABELS.REPAY_SHORT, 'SUM_ONLY');
  setCategoryLabel_(sheet, 24, '収', 'B');
  setCategoryLabel_(sheet, 24, '支', 'C');

  // 行25: 借入金返済（長期）
  setRow_(sheet, 25, ROW_LABELS.REPAY_LONG, 'SUM_ONLY');
  setCategoryLabel_(sheet, 25, '出', 'C');

  // 行26: 財務収支計(F) = (21+22+23) - (24+25)
  setRow_(sheet, 26, ROW_LABELS.FINANCE_TOTAL, 'SUM_AVG');
  sheet.getRange('E26').setFontWeight('bold');
  setCategoryLabel_(sheet, 26, '支', 'B');
  setCategoryLabel_(sheet, 26, '計', 'D');
  // 財務収支計の数式
  for (var c = 6; c <= 17; c++) {
    var col = columnLetter_(c);
    sheet.getRange(col + '26').setFormula(
      '=' + col + '21+' + col + '22+' + col + '23-' + col + '24-' + col + '25'
    );
  }
  // 合計
  sheet.getRange('R26').setFormula('=SUM(F26:Q26)');

  // --- 行27: 翌月繰越金(G) = (A)+(D)+(E)+(F) ---
  setRow_(sheet, 27, ROW_LABELS.NEXT_CARRY, 'NONE');
  sheet.getRange('E27').setFontWeight('bold');
  for (var c = 6; c <= 17; c++) {
    var col = columnLetter_(c);
    sheet.getRange(col + '27').setFormula(
      '=' + col + '5+' + col + '17+' + col + '20+' + col + '26'
    );
  }

  // --- 前月繰越金の自動連携（4月以降 = 前月の翌月繰越金） ---
  for (var c = 7; c <= 17; c++) {
    var prevCol = columnLetter_(c - 1);
    var col = columnLetter_(c);
    sheet.getRange(col + '5').setFormula('=' + prevCol + '27');
  }

  // --- 書式設定 ---
  applyFormatting_(sheet);
}


// ==============================
// ヘルパー関数
// ==============================

/**
 * 行にラベルと合計/平均の数式をセット
 * type: 'SUM_AVG' | 'SUM_ONLY' | 'NONE'
 */
function setRow_(sheet, row, label, type) {
  sheet.getRange('E' + row).setValue(label);

  if (type === 'SUM_AVG' || type === 'SUM_ONLY') {
    sheet.getRange('R' + row).setFormula('=SUM(F' + row + ':Q' + row + ')');
  }
  if (type === 'SUM_AVG') {
    sheet.getRange('S' + row).setFormula('=R' + row + '/12');
  }
}

/**
 * カテゴリラベルをセット（A〜D列）
 */
function setCategoryLabel_(sheet, row, label, col) {
  sheet.getRange(col + row).setValue(label)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
}

/**
 * 小計行の数式をセット（複数行の合計）
 */
function setSubtotalFormulas_(sheet, totalRow, sourceRows) {
  for (var c = 6; c <= 17; c++) {
    var col = columnLetter_(c);
    var formula = '=' + sourceRows.map(function(r) { return col + r; }).join('+');
    sheet.getRange(col + totalRow).setFormula(formula);
  }
}

/**
 * 差引行の数式をセット（row1 - row2）
 */
function setDiffFormulas_(sheet, resultRow, row1, row2) {
  for (var c = 6; c <= 17; c++) {
    var col = columnLetter_(c);
    sheet.getRange(col + resultRow).setFormula('=' + col + row1 + '-' + col + row2);
  }
}

/**
 * 列番号 → 列文字（1=A, 6=F, 18=R, 19=S）
 */
function columnLetter_(colNum) {
  return String.fromCharCode(64 + colNum);
}

/**
 * 書式設定
 */
function applyFormatting_(sheet) {
  var dataRange = sheet.getRange('A1:S27');

  // 数値フォーマット（千円・カンマ区切り・マイナスは▲表示）
  sheet.getRange('F3:S27').setNumberFormat('#,##0;▲#,##0');

  // 全体フォント
  dataRange.setFontFamily('游ゴシック');
  dataRange.setFontSize(10);

  // 小計行の背景色（薄いグレー）
  var subtotalRows = [9, 16, 17, 20, 26, 27];
  subtotalRows.forEach(function(row) {
    sheet.getRange('A' + row + ':S' + row).setBackground('#E8F0E8');
  });

  // ヘッダー行（行5: 前月繰越金）を強調
  sheet.getRange('A5:S5').setBackground('#FFF8E1');
  sheet.getRange('E5').setFontWeight('bold');

  // 翌月繰越金も強調
  sheet.getRange('A27:S27').setBackground('#FFF8E1');

  // 罫線
  dataRange.setBorder(true, true, true, true, true, true,
    '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  // ヘッダー行の下に太線
  sheet.getRange('A2:S2').setBorder(null, null, true, null, null, null,
    '#006060', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // 数値セルを右寄せ
  sheet.getRange('F3:S27').setHorizontalAlignment('right');

  // 項目名を左寄せ
  sheet.getRange('E3:E27').setHorizontalAlignment('left');

  // 行の高さ
  for (var r = 2; r <= 27; r++) {
    sheet.setRowHeight(r, 24);
  }
}
