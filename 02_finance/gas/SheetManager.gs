/**
 * 財務レポート自動化システム - スプレッドシート管理
 *
 * PLデータをスプレッドシートに書き込み・フォーマットする
 */

const SheetManager = {

  // ==============================
  // スプレッドシート取得
  // ==============================

  getSpreadsheet() {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  },

  /**
   * シートを取得または作成する
   * @param {string} sheetName
   * @return {Sheet}
   */
  getOrCreateSheet(sheetName) {
    const ss    = SheetManager.getSpreadsheet();
    let   sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      Logger.log(`シート作成: ${sheetName}`);
    }
    return sheet;
  },

  // ==============================
  // PL月別推移シート書き込み
  // ==============================

  /**
   * 部門別PLシートを書き込む
   * @param {number} fiscalYear
   * @param {string} deptName  - '共通' | '物販' | 'ブランド' | '民泊' | '全体'
   * @param {Object} monthlyRows - { monthLabel: [PLrows], ... }
   */
  writePLSheet(fiscalYear, deptName, monthlyRows) {
    const isConsolidated = (deptName === '全体');
    const sheetName = isConsolidated
      ? buildSheetName(fiscalYear, 'PL_CONSOLIDATED')
      : buildSheetName(fiscalYear, 'PL_PREFIX', deptName);

    const sheet  = SheetManager.getOrCreateSheet(sheetName);
    sheet.clearContents();
    sheet.clearFormats();

    const months = getFiscalMonths(fiscalYear);
    const { headers, rows } = PLFormatter.buildMonthlyMatrix(monthlyRows, months);
    const periodLabel = getFiscalPeriodLabel(fiscalYear);

    // ── ヘッダー行 ──────────────────────────────
    // 行1: タイトル
    sheet.getRange(1, 1).setValue(`${periodLabel} 損益計算書_月次推移_${deptName}`);
    sheet.getRange(1, 1).setFontSize(12).setFontWeight('bold');

    // 行2: 列ヘッダー（勘定科目 | 3月 | 4月 | ... | 2月 | 決算整理 | 合計）
    sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
    SheetManager._styleHeaderRow(sheet, 2, headers.length);

    // ── データ行 ──────────────────────────────
    const dataStartRow = 3;
    rows.forEach((row, i) => {
      const rowNum = dataStartRow + i;
      const numCols = headers.length;

      // 勘定科目名（インデント付き）
      sheet.getRange(rowNum, 1).setValue(row.labelCell);

      // 数値データ（ヘッダー行はスキップ）
      if (!row.isHeader) {
        const values = row.values.map(v => v || 0);
        sheet.getRange(rowNum, 2, 1, values.length).setValues([values]);
      }

      // スタイル適用
      SheetManager._styleDataRow(sheet, rowNum, numCols, row);
    });

    // ── 列幅・フォーマット ──────────────────────────────
    SheetManager._applyColumnFormats(sheet, headers.length, dataStartRow, dataStartRow + rows.length - 1);

    Logger.log(`PLシート書き込み完了: ${sheetName}`);
    return sheet;
  },

  // ==============================
  // 推移表シート書き込み
  // ==============================

  /**
   * 年度別×部門別推移表を書き込む
   * 構成: 縦=PL項目, 横=部門×月 または 部門×年度
   *
   * @param {number} fiscalYear
   * @param {Object} allDeptData - { deptName: { monthLabel: [PLrows] } }
   */
  writeTrendByDeptSheet(fiscalYear, allDeptData) {
    const sheetName = buildSheetName(fiscalYear, 'TREND_DEPT');
    const sheet     = SheetManager.getOrCreateSheet(sheetName);
    sheet.clearContents();
    sheet.clearFormats();

    const months      = getFiscalMonths(fiscalYear);
    const periodLabel = getFiscalPeriodLabel(fiscalYear);
    const depts       = CONFIG.DEPARTMENTS.map(d => d.name);

    // タイトル
    sheet.getRange(1, 1).setValue(`${periodLabel} 部門別推移表`);
    sheet.getRange(1, 1).setFontSize(12).setFontWeight('bold');

    // ヘッダー行2: 部門名（各部門が月数分スパン）
    // 行3: 月名
    let col = 2;
    const headerRow2 = ['勘定科目'];
    const headerRow3 = ['勘定科目'];

    depts.forEach(dept => {
      headerRow2.push(dept);
      months.forEach((m, mi) => {
        if (mi > 0) headerRow2.push(''); // 結合のため空白
        headerRow3.push(m.label);
      });
      headerRow2.push(''); headerRow3.push('合計'); // 合計列
      col += months.length + 1;
    });

    sheet.getRange(2, 1, 1, headerRow2.length).setValues([headerRow2]);
    sheet.getRange(3, 1, 1, headerRow3.length).setValues([headerRow3]);

    // 部門名の結合
    let mergeStartCol = 2;
    depts.forEach(dept => {
      if (months.length + 1 > 1) {
        sheet.getRange(2, mergeStartCol, 1, months.length + 1).merge();
      }
      sheet.getRange(2, mergeStartCol).setHorizontalAlignment('center').setFontWeight('bold');
      mergeStartCol += months.length + 1;
    });

    SheetManager._styleHeaderRow(sheet, 2, headerRow2.length);
    SheetManager._styleHeaderRow(sheet, 3, headerRow3.length);

    // データ行
    const dataStartRow = 4;
    CONFIG.PL_STRUCTURE.forEach((item, i) => {
      const rowNum    = dataStartRow + i;
      const labelCell = '　'.repeat(item.indent || 0) + item.label;
      sheet.getRange(rowNum, 1).setValue(labelCell);

      let dataCol = 2;
      depts.forEach(dept => {
        const monthlyRows = allDeptData[dept] || {};
        let deptTotal = 0;

        months.forEach(m => {
          const plRows = monthlyRows[m.label] || [];
          const plRow  = plRows.find(r => r.label === item.label);
          const val    = (plRow && plRow.amount !== null && !plRow.isHeader) ? (plRow.amount || 0) : '';

          if (typeof val === 'number') {
            sheet.getRange(rowNum, dataCol).setValue(val);
            deptTotal += val;
          }
          dataCol++;
        });

        // 部門合計
        if (item.category !== 'header') {
          sheet.getRange(rowNum, dataCol).setValue(deptTotal);
        }
        dataCol++;
      });

      // 行スタイル
      const rowObj = { isBold: item.isBold, isBorderTop: item.isBorderTop, isHeader: item.category === 'header', isSubtotal: item.category === 'subtotal' };
      SheetManager._styleDataRow(sheet, rowNum, headerRow2.length, rowObj);
    });

    SheetManager._applyColumnFormats(sheet, headerRow2.length, dataStartRow, dataStartRow + CONFIG.PL_STRUCTURE.length - 1);
    Logger.log(`部門別推移表書き込み完了: ${sheetName}`);
  },

  /**
   * 年度別×全体推移表を書き込む
   * 金額と前年比（%）を並べて表示
   *
   * @param {number} fiscalYear
   * @param {Object} allDeptData - { deptName: { monthLabel: [PLrows] } }
   */
  writeTrendConsolidatedSheet(fiscalYear, allDeptData) {
    const sheetName   = buildSheetName(fiscalYear, 'TREND_TOTAL');
    const sheet       = SheetManager.getOrCreateSheet(sheetName);
    sheet.clearContents();
    sheet.clearFormats();

    const months      = getFiscalMonths(fiscalYear);
    const periodLabel = getFiscalPeriodLabel(fiscalYear);

    // タイトル
    sheet.getRange(1, 1).setValue(`${periodLabel} 全体推移表（月別・金額・構成比）`);
    sheet.getRange(1, 1).setFontSize(12).setFontWeight('bold');

    // ヘッダー
    const headers = ['勘定科目', ...months.map(m => m.label), '決算整理', '合計', '売上比'];
    sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
    SheetManager._styleHeaderRow(sheet, 2, headers.length);

    const monthlyRows = allDeptData['全体'] || {};
    const { rows }    = PLFormatter.buildMonthlyMatrix(monthlyRows, months);

    // 売上高合計の値を取得（構成比計算用）
    const revenueRow  = rows.find(r => r.label === '売上高合計');
    const totalRevenue = revenueRow ? revenueRow.values[revenueRow.values.length - 1] : 0;

    const dataStartRow = 3;
    rows.forEach((row, i) => {
      const rowNum = dataStartRow + i;
      sheet.getRange(rowNum, 1).setValue(row.labelCell);

      if (!row.isHeader) {
        const values = row.values.map(v => v || 0);
        sheet.getRange(rowNum, 2, 1, values.length).setValues([values]);

        // 売上比（合計÷売上高合計）
        const total = values[values.length - 1];
        if (totalRevenue !== 0 && !row.isSubtotal) {
          const ratio = total / totalRevenue;
          sheet.getRange(rowNum, 2 + values.length).setValue(ratio).setNumberFormat('0.0%');
        }
      }

      SheetManager._styleDataRow(sheet, rowNum, headers.length, row);
    });

    SheetManager._applyColumnFormats(sheet, headers.length - 1, dataStartRow, dataStartRow + rows.length - 1);
    // 売上比列は%フォーマット済みなのでスキップ
    Logger.log(`全体推移表書き込み完了: ${sheetName}`);
  },

  // ==============================
  // スタイルユーティリティ
  // ==============================

  _styleHeaderRow(sheet, rowNum, numCols) {
    const range = sheet.getRange(rowNum, 1, 1, numCols);
    range.setBackground('#1a1a2e')
         .setFontColor('#ffffff')
         .setFontWeight('bold')
         .setHorizontalAlignment('center');
  },

  _styleDataRow(sheet, rowNum, numCols, row) {
    const range = sheet.getRange(rowNum, 1, 1, numCols);

    if (row.isHeader) {
      // カテゴリヘッダー（売上高、売上原価 など）
      range.setBackground('#e8eaf6').setFontWeight('bold');
      sheet.getRange(rowNum, 1).setFontWeight('bold');
    } else if (row.isBold && row.isSubtotal) {
      // 小計行（売上高合計、営業利益 など）
      range.setBackground('#c5cae9').setFontWeight('bold');
    } else if (row.isBold) {
      // 利益行（売上総利益、経常利益 など）
      range.setBackground('#bbdefb').setFontWeight('bold');
    }

    // 上線
    if (row.isBorderTop) {
      range.setBorder(true, null, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    }
  },

  _applyColumnFormats(sheet, numCols, dataStartRow, dataEndRow) {
    if (dataEndRow < dataStartRow) return;

    // A列: 勘定科目名の幅
    sheet.setColumnWidth(1, 200);

    // B列以降: 数値列の幅・フォーマット
    for (let c = 2; c <= numCols; c++) {
      sheet.setColumnWidth(c, 100);
    }

    // 数値フォーマット（カンマ区切り）
    if (numCols >= 2) {
      sheet.getRange(dataStartRow, 2, dataEndRow - dataStartRow + 1, numCols - 1)
           .setNumberFormat('#,##0;[RED]-#,##0;"-"');
    }

    // 行ヘッダー列を固定
    sheet.setFrozenRows(2);
    sheet.setFrozenColumns(1);
  },
};

/**
 * スプレッドシートを取得するグローバルヘルパー（他ファイルから呼ばれる）
 */
function getSpreadsheet() {
  return SheetManager.getSpreadsheet();
}
