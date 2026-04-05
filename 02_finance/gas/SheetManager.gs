/**
 * 財務レポート自動化システム - スプレッドシート管理 v1.2
 *
 * PLシート列構成（15列）:
 *   A     (col  1): 勘定科目
 *   B〜M  (col 2〜13): 各月 金額（3月〜2月）  monthIdx + 2
 *   N     (col 14): 決算整理 金額（手入力）
 *   O     (col 15): 合計 金額（数式）
 *
 * 更新ルール:
 *   - 初回: 全列にヘッダー・ラベル・数式を書き込む
 *   - 2回目以降: 過去・当月の「金額列」のみ上書き。数式列・未来月は一切触らない
 */

const SheetManager = {

  // ── 列インデックス定数 ──────────────────────────
  COL: {
    LABEL:    1,   // A: 勘定科目
    // 月金額: 2〜13  (monthIdx + 2)
    ADJ:     14,  // N: 決算整理
    TOTAL:   15,  // O: 合計
    NUM_COLS: 15, // 総列数
  },

  DATA_START_ROW: 3, // データ開始行（1=タイトル, 2=ヘッダー）

  // ==============================
  // スプレッドシート取得
  // ==============================

  getSpreadsheet() {
    if (CONFIG.SPREADSHEET_ID) {
      return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    }
    return SpreadsheetApp.getActiveSpreadsheet();
  },

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
  // PLシート書き込み（メイン）
  // ==============================

  /**
   * 部門別PLシートを書き込む
   *
   * @param {number} fiscalYear
   * @param {string} deptName       - '共通' | '物販' | 'ブランド' | '民泊' | '全体'
   * @param {Object} monthlyRows    - { monthLabel: [PLrows], ... }（更新対象月のみ）
   * @param {Array}  monthsToUpdate - 更新対象月リスト（getFiscalMonths()の部分集合）
   */
  writePLSheet(fiscalYear, deptName, monthlyRows, monthsToUpdate) {
    const isConsolidated = (deptName === '全体');
    const sheetName = isConsolidated
      ? buildSheetName(fiscalYear, 'PL_CONSOLIDATED')
      : buildSheetName(fiscalYear, 'PL_PREFIX', deptName);

    const sheet    = SheetManager.getOrCreateSheet(sheetName);
    const allMonths = getFiscalMonths(fiscalYear);
    const isNew    = (sheet.getLastRow() <= 1);

    if (isNew) {
      // ── 初回: 全体を初期化 ──
      Logger.log(`初期化: ${sheetName}`);
      SheetManager._initPLSheet(sheet, fiscalYear, deptName, allMonths);
    }

    // ── 金額列のみ更新（過去・当月分） ──
    monthsToUpdate.forEach(m => {
      const monthIdx = allMonths.findIndex(am => am.label === m.label);
      if (monthIdx < 0) return;

      const plRows = monthlyRows[m.label] || PLFormatter.buildPLRows([]);
      SheetManager._writeAmountColumn(sheet, monthIdx, plRows);
    });

    // ── 決算整理列の書き込み（CSVに決算整理データがある場合のみ） ──
    if (monthlyRows['決算整理']) {
      SheetManager._writeAdjColumn(sheet, monthlyRows['決算整理']);
    }

    Logger.log(`PLシート更新完了: ${sheetName}（${monthsToUpdate.length}ヶ月）`);
    return sheet;
  },

  /**
   * PLシートを初期化する（初回のみ）
   * ヘッダー・ラベル・合計数式を書き込む
   */
  _initPLSheet(sheet, fiscalYear, deptName, allMonths) {
    const C           = SheetManager.COL;
    const dataStart   = SheetManager.DATA_START_ROW;
    const periodLabel = getFiscalPeriodLabel(fiscalYear);
    const numItems    = CONFIG.PL_STRUCTURE.length;
    const revenueRow  = dataStart + CONFIG.PL_STRUCTURE.findIndex(i => i.label === '売上高合計');

    // タイトル行
    sheet.getRange(1, 1).setValue(`${periodLabel} 損益計算書_月次推移_${deptName}`)
         .setFontSize(12).setFontWeight('bold');

    // ヘッダー行（行2）
    const headers = SheetManager._buildHeaders(allMonths);
    sheet.getRange(2, 1, 1, C.NUM_COLS).setValues([headers]);
    SheetManager._styleHeaderRow(sheet, 2, C.NUM_COLS);

    // ラベル列（A列）
    const labels = CONFIG.PL_STRUCTURE.map(item =>
      '　'.repeat(item.indent || 0) + item.label
    );
    sheet.getRange(dataStart, 1, numItems, 1).setValues(labels.map(l => [l]));

    // 合計数式を書き込む（非ヘッダー行のみ）
    CONFIG.PL_STRUCTURE.forEach((item, i) => {
      const rowNum = dataStart + i;
      if (item.category === 'header') return;

      // 合計数式（O列）
      const amtCols  = Array.from({length: 12}, (_, mi) => SheetManager._colLetter(2 + mi));
      const adjLtr   = SheetManager._colLetter(C.ADJ);
      sheet.getRange(rowNum, C.TOTAL).setFormula(
        `=${amtCols.map(c => `${c}${rowNum}`).join('+')}+${adjLtr}${rowNum}`
      );
    });

    // スタイル適用
    CONFIG.PL_STRUCTURE.forEach((item, i) => {
      SheetManager._styleDataRow(sheet, dataStart + i, C.NUM_COLS, {
        isHeader:   item.category === 'header',
        isSubtotal: item.category === 'subtotal',
        isBold:     item.isBold || false,
        isBorderTop: item.isBorderTop || false,
      });
    });

    // 列幅・フォント・固定
    SheetManager._applyColumnFormats(sheet, numItems, dataStart, revenueRow);
  },

  /**
   * 決算整理列（N列）にデータを書き込む
   * @param {Sheet} sheet
   * @param {Array} plRows - PLFormatter.buildPLRows() の結果
   */
  _writeAdjColumn(sheet, plRows) {
    const dataStart = SheetManager.DATA_START_ROW;
    const adjCol    = SheetManager.COL.ADJ; // 14

    const values = CONFIG.PL_STRUCTURE.map((item, i) => {
      const plRow = plRows[i];
      if (!plRow || item.category === 'header') return [''];
      return [plRow.amount !== null ? (plRow.amount || 0) : 0];
    });

    sheet.getRange(dataStart, adjCol, values.length, 1).setValues(values);
  },

  /**
   * 特定月の金額列にデータを書き込む
   * @param {Sheet}  sheet
   * @param {number} monthIdx - 0=3月, 1=4月, ..., 11=2月
   * @param {Array}  plRows   - PLFormatter.buildPLRows() の結果
   */
  _writeAmountColumn(sheet, monthIdx, plRows) {
    const dataStart = SheetManager.DATA_START_ROW;
    const amtCol    = 2 + monthIdx;   // B=2(3月), C=3(4月), ... M=13(2月)

    const values = CONFIG.PL_STRUCTURE.map((item, i) => {
      const plRow = plRows[i];
      if (!plRow || item.category === 'header') return [''];
      return [plRow.amount !== null ? (plRow.amount || 0) : 0];
    });

    sheet.getRange(dataStart, amtCol, values.length, 1).setValues(values);
  },

  // ==============================
  // 推移表シート書き込み
  // ==============================

  /**
   * 部門別推移表を書き込む
   * 縦=PL項目, 横=部門×[月金額, 部門合計]
   *
   * @param {number} fiscalYear
   * @param {Object} allDeptData   - { deptName: { monthLabel: [PLrows] } }
   * @param {Array}  monthsToUpdate - 更新対象月リスト
   */
  writeTrendByDeptSheet(fiscalYear, allDeptData, monthsToUpdate) {
    const sheetName = buildSheetName(fiscalYear, 'TREND_DEPT');
    const sheet     = SheetManager.getOrCreateSheet(sheetName);
    sheet.clearContents();
    sheet.clearFormats();

    const allMonths   = getFiscalMonths(fiscalYear);
    const depts       = CONFIG.DEPARTMENTS.map(d => d.name);
    const periodLabel = getFiscalPeriodLabel(fiscalYear);
    const dataStart   = 4; // 行1=タイトル, 行2=部門名, 行3=列名

    sheet.getRange(1, 1).setValue(`${periodLabel} 部門別推移表（月別金額）`)
         .setFontSize(12).setFontWeight('bold');

    // ヘッダー行2: 部門名, 行3: 月ラベル
    const colsPerDept = allMonths.length + 1; // 各月金額 + 部門合計
    const headerRow2  = ['勘定科目'];
    const headerRow3  = ['勘定科目'];

    depts.forEach(dept => {
      headerRow2.push(dept);
      Array.from({length: colsPerDept - 1}, () => headerRow2.push(''));
      allMonths.forEach(m => headerRow3.push(m.label));
      headerRow3.push('合計');
    });

    const totalCols = 1 + depts.length * colsPerDept;
    sheet.getRange(2, 1, 1, totalCols).setValues([headerRow2]);
    sheet.getRange(3, 1, 1, totalCols).setValues([headerRow3]);

    // 部門名セルを結合・スタイル
    let mergeCol = 2;
    depts.forEach(() => {
      sheet.getRange(2, mergeCol, 1, colsPerDept).merge()
           .setHorizontalAlignment('center').setFontWeight('bold');
      mergeCol += colsPerDept;
    });
    SheetManager._styleHeaderRow(sheet, 2, totalCols);
    SheetManager._styleHeaderRow(sheet, 3, totalCols);

    // データ行
    const updatableLabels = new Set(monthsToUpdate.map(m => m.label));

    CONFIG.PL_STRUCTURE.forEach((item, i) => {
      const rowNum    = dataStart + i;
      const labelCell = '　'.repeat(item.indent || 0) + item.label;
      sheet.getRange(rowNum, 1).setValue(labelCell);

      let dataCol = 2;
      depts.forEach(dept => {
        const monthlyRows = allDeptData[dept] || {};
        let deptTotal = 0;

        allMonths.forEach(m => {
          const isUpdatable = updatableLabels.has(m.label);
          const plRows      = monthlyRows[m.label] || [];
          const plRow       = plRows[i];
          const val = (plRow && item.category !== 'header' && plRow.amount !== null)
            ? (plRow.amount || 0) : '';

          if (isUpdatable && typeof val === 'number') {
            sheet.getRange(rowNum, dataCol).setValue(val);
            deptTotal += val;
          }
          dataCol++;
        });

        // 部門合計
        if (item.category !== 'header') {
          sheet.getRange(rowNum, dataCol).setValue(deptTotal || 0);
        }
        dataCol++;
      });

      SheetManager._styleDataRow(sheet, rowNum, totalCols, {
        isHeader:   item.category === 'header',
        isSubtotal: item.category === 'subtotal',
        isBold:     item.isBold || false,
        isBorderTop: item.isBorderTop || false,
      });
    });

    sheet.setColumnWidth(1, 200);
    for (let c = 2; c <= totalCols; c++) {
      sheet.setColumnWidth(c, 90);
    }
    sheet.getRange(dataStart, 2, CONFIG.PL_STRUCTURE.length, totalCols - 1)
         .setNumberFormat('#,##0;[RED]-#,##0;"-"');
    sheet.setFrozenRows(3);
    sheet.setFrozenColumns(1);
    Logger.log(`部門別推移表書き込み完了: ${sheetName}`);
  },

  /**
   * 全体推移表を書き込む（月別金額）
   *
   * @param {number} fiscalYear
   * @param {Object} allDeptData   - { '全体': { monthLabel: [PLrows] } }
   * @param {Array}  monthsToUpdate - 更新対象月リスト
   */
  writeTrendConsolidatedSheet(fiscalYear, allDeptData, monthsToUpdate) {
    const sheetName = buildSheetName(fiscalYear, 'TREND_TOTAL');
    const sheet     = SheetManager.getOrCreateSheet(sheetName);
    const isNew     = (sheet.getLastRow() <= 1);
    const C         = SheetManager.COL;
    const dataStart = SheetManager.DATA_START_ROW;
    const allMonths = getFiscalMonths(fiscalYear);
    const periodLabel = getFiscalPeriodLabel(fiscalYear);

    if (isNew) {
      // 初回: ヘッダー・ラベル・数式を初期化
      const headers = SheetManager._buildHeaders(allMonths);
      sheet.getRange(1, 1).setValue(`${periodLabel} 全体推移表（月別金額）`)
           .setFontSize(12).setFontWeight('bold');
      sheet.getRange(2, 1, 1, C.NUM_COLS).setValues([headers]);
      SheetManager._styleHeaderRow(sheet, 2, C.NUM_COLS);

      const labels = CONFIG.PL_STRUCTURE.map(item =>
        '　'.repeat(item.indent || 0) + item.label
      );
      sheet.getRange(dataStart, 1, labels.length, 1).setValues(labels.map(l => [l]));

      CONFIG.PL_STRUCTURE.forEach((item, i) => {
        const rowNum = dataStart + i;
        if (item.category === 'header') return;

        const amtCols  = Array.from({length: 12}, (_, mi) => SheetManager._colLetter(2 + mi));
        const adjLtr   = SheetManager._colLetter(C.ADJ);
        sheet.getRange(rowNum, C.TOTAL).setFormula(
          `=${amtCols.map(c => `${c}${rowNum}`).join('+')}+${adjLtr}${rowNum}`
        );
      });

      CONFIG.PL_STRUCTURE.forEach((item, i) => {
        SheetManager._styleDataRow(sheet, dataStart + i, C.NUM_COLS, {
          isHeader:   item.category === 'header',
          isSubtotal: item.category === 'subtotal',
          isBold:     item.isBold || false,
          isBorderTop: item.isBorderTop || false,
        });
      });

      const revenueRowIdx = CONFIG.PL_STRUCTURE.findIndex(i => i.label === '売上高合計');
      SheetManager._applyColumnFormats(sheet, CONFIG.PL_STRUCTURE.length, dataStart, dataStart + revenueRowIdx);
    }

    // 金額列のみ更新（過去・当月分）
    const monthlyRows = allDeptData['全体'] || {};
    monthsToUpdate.forEach(m => {
      const monthIdx = allMonths.findIndex(am => am.label === m.label);
      if (monthIdx < 0) return;
      const plRows = monthlyRows[m.label] || PLFormatter.buildPLRows([]);
      SheetManager._writeAmountColumn(sheet, monthIdx, plRows);
    });

    Logger.log(`全体推移表更新完了: ${sheetName}`);
  },

  // ==============================
  // ユーティリティ
  // ==============================

  /**
   * 列ヘッダー配列を生成する
   * [勘定科目, 3月, 4月, ..., 2月, 決算整理, 合計]
   */
  _buildHeaders(allMonths) {
    const headers = ['勘定科目'];
    allMonths.forEach(m => headers.push(m.label));
    headers.push('決算整理', '合計');
    return headers;
  },

  /**
   * 列番号（1始まり）→列アルファベット（A, B, ..., Z, AA, ...）
   */
  _colLetter(n) {
    let result = '';
    while (n > 0) {
      const rem = (n - 1) % 26;
      result = String.fromCharCode(65 + rem) + result;
      n = Math.floor((n - 1) / 26);
    }
    return result;
  },

  _styleHeaderRow(sheet, rowNum, numCols) {
    sheet.getRange(rowNum, 1, 1, numCols)
         .setBackground('#1a1a2e')
         .setFontColor('#ffffff')
         .setFontWeight('bold')
         .setHorizontalAlignment('center');
  },

  _styleDataRow(sheet, rowNum, numCols, row) {
    const range = sheet.getRange(rowNum, 1, 1, numCols);
    if (row.isHeader) {
      range.setBackground('#e8eaf6').setFontWeight('bold');
    } else if (row.isBold && row.isSubtotal) {
      range.setBackground('#c5cae9').setFontWeight('bold');
    } else if (row.isBold) {
      range.setBackground('#bbdefb').setFontWeight('bold');
    }
    if (row.isBorderTop) {
      range.setBorder(true, null, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    }
  },

  _applyColumnFormats(sheet, numDataRows, dataStart, revenueRow) {
    const C = SheetManager.COL;
    sheet.setColumnWidth(1, 200);
    for (let mi = 0; mi < 12; mi++) {
      sheet.setColumnWidth(2 + mi, 90);
    }
    sheet.setColumnWidth(C.ADJ,   90);
    sheet.setColumnWidth(C.TOTAL, 100);

    // 数値フォーマット: 全金額列
    sheet.getRange(dataStart, 2, numDataRows, 14)
         .setNumberFormat('#,##0;[RED]-#,##0;"-"');

    sheet.setFrozenRows(2);
    sheet.setFrozenColumns(1);
  },

  // ==============================
  // 通期比較シート（全期一覧）
  // ==============================

  /**
   * 各期の 第X期_PL_全体 シートの合計列（O列）を横並びにした通期比較シートを作成する
   *
   * @param {Array} fiscalYears - [2018, 2019, ..., 2025] など
   */
  writePeriodComparisonSheet(fiscalYears) {
    const ss        = SheetManager.getSpreadsheet();
    const sheetName = '通期比較_全体';
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);
    sheet.clearContents();
    sheet.clearFormats();

    const dataStart     = 4; // 行1=タイトル, 行2=期ヘッダー, 行3=期間サブヘッダー
    const srcDataStart  = SheetManager.DATA_START_ROW; // ソースシートは常に行3
    const numItems      = CONFIG.PL_STRUCTURE.length;
    const numPeriods    = fiscalYears.length;
    const colsPerPeriod = 2; // 金額 + 売上比率
    const totalCols     = 1 + numPeriods * colsPerPeriod;

    // 売上高合計の行番号（売上比率の分母）
    const revenueRowIdx = CONFIG.PL_STRUCTURE.findIndex(item => item.label === '売上高合計');
    const revenueAbsRow = dataStart + revenueRowIdx;

    // タイトル
    sheet.getRange(1, 1).setValue('通期比較 損益計算書（年度別合計）')
         .setFontSize(12).setFontWeight('bold');

    // 行2: 期ヘッダー（第X期 | % | ...）
    const headers = ['勘定科目'];
    fiscalYears.forEach(y => {
      headers.push(getFiscalPeriodLabel(y));
      headers.push('%');
    });
    sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
    SheetManager._styleHeaderRow(sheet, 2, headers.length);

    // 行3: 事業期間サブヘッダー（2018年3月-2019年2月 | (空) | ...）
    const subHeaders = [''];
    fiscalYears.forEach(y => {
      subHeaders.push(`${y}年3月-${y+1}年2月`);
      subHeaders.push('');
    });
    sheet.getRange(3, 1, 1, subHeaders.length).setValues([subHeaders]);
    sheet.getRange(3, 1, 1, subHeaders.length)
         .setBackground('#3d3d5c')
         .setFontColor('#ccccdd')
         .setFontSize(9)
         .setHorizontalAlignment('center');

    // ラベル列（A列）
    const labels = CONFIG.PL_STRUCTURE.map(item => '　'.repeat(item.indent || 0) + item.label);
    sheet.getRange(dataStart, 1, numItems, 1).setValues(labels.map(l => [l]));

    // 各期：金額列 + 売上比率列
    fiscalYears.forEach((year, colIdx) => {
      const valueCol       = 2 + colIdx * colsPerPeriod;
      const ratioCol       = valueCol + 1;
      const valueColLetter = SheetManager._colLetter(valueCol);

      const srcName  = buildSheetName(year, 'PL_CONSOLIDATED');
      const srcSheet = ss.getSheetByName(srcName);
      if (!srcSheet) {
        Logger.log(`スキップ（シートなし）: ${srcName}`);
        return;
      }

      const lastRow = srcSheet.getLastRow();
      if (lastRow < srcDataStart) return;

      // ラベルマッチングで金額を取得
      const numSrcRows = lastRow - srcDataStart + 1;
      const srcLabels  = srcSheet.getRange(srcDataStart, 1, numSrcRows, 1).getValues()
                           .map(r => r[0].toString().trim());
      const srcTotals  = srcSheet.getRange(srcDataStart, SheetManager.COL.TOTAL, numSrcRows, 1).getValues()
                           .map(r => r[0]);
      const labelMap   = {};
      srcLabels.forEach((label, i) => { if (label) labelMap[label] = srcTotals[i]; });

      // 金額列
      const colValues = CONFIG.PL_STRUCTURE.map(item => {
        const val = labelMap[item.label];
        return [val !== undefined ? val : ''];
      });
      sheet.getRange(dataStart, valueCol, numItems, 1).setValues(colValues);

      // 売上比率列（指定項目のみ）
      const ratioLabels = new Set([
        '売上高合計', '売上原価合計', '売上総利益',
        '役員賞与', '役員報酬', '給料手当', '法定福利費', '福利厚生費',
        '研修採用費', '接待交際費', '旅費交通費', '通信費', '水道光熱費',
        '保険料', '租税公課', '支払手数料', '支払報酬', '業務委託費',
        '会議費', '新聞図書費', '減価償却費', '繰延資産償却', '長期前払費用償却',
        '荷造運賃', '広告宣伝費', '備品・消耗品費', '車両費', '地代家賃',
        '修繕費', '雑費',
        '販売費及び一般管理費合計', '営業利益', '経常利益', '税引前当期純利益', '当期純利益',
      ]);
      const ratioFormulas = CONFIG.PL_STRUCTURE.map((item, i) => {
        if (!ratioLabels.has(item.label)) return [''];
        const rowNum = dataStart + i;
        return [`=IF(${valueColLetter}${revenueAbsRow}=0,"",${valueColLetter}${rowNum}/${valueColLetter}${revenueAbsRow})`];
      });
      sheet.getRange(dataStart, ratioCol, numItems, 1).setFormulas(ratioFormulas);
    });

    // 行スタイル
    CONFIG.PL_STRUCTURE.forEach((item, i) => {
      SheetManager._styleDataRow(sheet, dataStart + i, totalCols, {
        isHeader:    item.category === 'header',
        isSubtotal:  item.category === 'subtotal',
        isBold:      item.isBold || false,
        isBorderTop: item.isBorderTop || false,
      });
    });

    // 列幅・数値フォーマット
    sheet.setColumnWidth(1, 200);
    for (let p = 0; p < numPeriods; p++) {
      const valueCol = 2 + p * colsPerPeriod;
      const ratioCol = valueCol + 1;
      sheet.setColumnWidth(valueCol, 110);
      sheet.setColumnWidth(ratioCol, 65);
      sheet.getRange(dataStart, valueCol, numItems, 1)
           .setNumberFormat('#,##0;[RED]-#,##0;"-"');
      sheet.getRange(dataStart, ratioCol, numItems, 1)
           .setNumberFormat('0.0%;[RED]-0.0%;"-"');
    }
    sheet.setFrozenRows(3);
    sheet.setFrozenColumns(1);

    Logger.log(`通期比較シート作成完了: ${sheetName}（${numPeriods}期分）`);
    return sheet;
  },
  // ==============================
  // 通期比較BSシート（全期一覧）
  // ==============================

  /**
   * 各期の 全体_BS_{年}.csv から読み込んだ期末残高を横並びにした通期比較BSシートを作成
   *
   * @param {Array}  fiscalYears  - [2018, 2019, ..., 2026]
   * @param {Object} yearDataMap  - { year: { level0: {}, level1: {} } }（BSImporter.importAllFromDrive()の戻り値）
   */
  writePeriodComparisonBSSheet(fiscalYears, yearDataMap) {
    const ss        = SheetManager.getSpreadsheet();
    const sheetName = '通期比較_BS';
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);
    sheet.clearContents();
    sheet.clearFormats();

    const dataStart  = 4;
    const numItems   = CONFIG.BS_STRUCTURE.length;
    const numPeriods = fiscalYears.length;
    const totalCols  = 1 + numPeriods;

    // タイトル
    sheet.getRange(1, 1).setValue('通期比較 貸借対照表（期末残高）')
         .setFontSize(12).setFontWeight('bold');

    // 行2: 期ヘッダー
    const headers = ['勘定科目'];
    fiscalYears.forEach(y => headers.push(getFiscalPeriodLabel(y)));
    sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
    SheetManager._styleHeaderRow(sheet, 2, headers.length);

    // 行3: 事業期間サブヘッダー
    const subHeaders = [''];
    fiscalYears.forEach(y => subHeaders.push(`${y}年3月-${y+1}年2月`));
    sheet.getRange(3, 1, 1, subHeaders.length).setValues([subHeaders]);
    sheet.getRange(3, 1, 1, subHeaders.length)
         .setBackground('#3d3d5c').setFontColor('#ccccdd')
         .setFontSize(9).setHorizontalAlignment('center');

    // ラベル列（A列）
    const labels = CONFIG.BS_STRUCTURE.map(item => '　'.repeat(item.indent || 0) + item.label);
    sheet.getRange(dataStart, 1, numItems, 1).setValues(labels.map(l => [l]));

    // 各期データ列
    fiscalYears.forEach((year, colIdx) => {
      const valueCol = 2 + colIdx;
      const data     = yearDataMap[year];
      if (!data) {
        Logger.log(`スキップ（データなし）: ${getFiscalPeriodLabel(year)}`);
        return;
      }

      const colValues = CONFIG.BS_STRUCTURE.map(item => {
        if (item.category === 'header') return [''];
        const map = item.srcLevel === 0 ? data.level0 : data.level1;
        const val = map[item.srcLabel];
        return [val !== undefined ? val : ''];
      });
      sheet.getRange(dataStart, valueCol, numItems, 1).setValues(colValues);
    });

    // 行スタイル
    CONFIG.BS_STRUCTURE.forEach((item, i) => {
      SheetManager._styleDataRow(sheet, dataStart + i, totalCols, {
        isHeader:    item.category === 'header',
        isSubtotal:  item.category === 'subtotal' || item.category === 'total',
        isBold:      item.isBold || false,
        isBorderTop: item.isBorderTop || false,
      });
    });

    // 列幅・数値フォーマット
    sheet.setColumnWidth(1, 200);
    for (let p = 0; p < numPeriods; p++) {
      const valueCol = 2 + p;
      sheet.setColumnWidth(valueCol, 120);
      sheet.getRange(dataStart, valueCol, numItems, 1)
           .setNumberFormat('#,##0;[RED]-#,##0;"-"');
    }
    sheet.setFrozenRows(3);
    sheet.setFrozenColumns(1);

    Logger.log(`通期比較BSシート作成完了: ${sheetName}（${numPeriods}期分）`);
    return sheet;
  },
};

function getSpreadsheet() {
  return SheetManager.getSpreadsheet();
}
