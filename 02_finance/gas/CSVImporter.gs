/**
 * CSVインポーター - MF会計 部門別推移試算表CSVを読み込んでPLシートに反映
 *
 * ■ 使い方
 *   1. MF会計 → レポート → 推移試算表 → 部門選択 → 全期間 → CSVダウンロード
 *   2. Googleドライブの指定フォルダ（Config.gs の CSV_IMPORT.FOLDER_ID）にアップロード
 *      ファイル名 = {部門名}_PL_{年}.csv（例: 物販_PL_2026.csv / 全体_PL_2026.csv）
 *   3. スプレッドシートのメニュー「📥 部門別CSVをインポート」を実行
 *
 * ■ MF会計 推移試算表のCSVフォーマット
 *   ヘッダー行: (空) | 勘定科目 | 補助科目 | 3月 | 4月 | ... | 2月 | 決算整理 | 合計
 *   Level 0（col 0が非空）: カテゴリ行・小計行 → スキップ
 *   Level 1（col 1が非空）: 勘定科目行 ← ここを使用
 *   Level 2（col 2が非空）: 補助科目行 → スキップ（Level 1が合計値）
 */

const CSVImporter = {

  /**
   * 現在の事業年度のCSVをインポートする（ダッシュボード用・年度入力なし）
   * 結果をオブジェクトで返す（UIアラートなし）
   *
   * @returns {{ ok: boolean, msg: string, imported: number, errors: string[] }}
   */
  importCurrentYear() {
    CSVImporter._lastUnmatched = [];
    const folderId = CONFIG.CSV_IMPORT && CONFIG.CSV_IMPORT.FOLDER_ID;
    if (!folderId) {
      return { ok: false, msg: '⚠️ CSV_IMPORT.FOLDER_ID が未設定', imported: 0, errors: [] };
    }

    const targetYear = getCurrentFiscalYear();
    const months     = getFiscalMonths(targetYear);
    const label      = getFiscalPeriodLabel(targetYear);
    let imported     = 0;
    const errors     = [];
    const allDeptData = {};
    let consolidatedFromCSV = false;

    try {
      const folder = DriveApp.getFolderById(folderId);
      const files  = folder.getFiles();

      while (files.hasNext()) {
        const file     = files.next();
        const fileName = file.getName();
        if (!/\.(csv|tsv|txt)$/i.test(fileName)) continue;

        const rawName = fileName.replace(/\.(csv|tsv|txt)$/i, '').trim();
        let fileYear, deptKey;
        const plMatch = rawName.match(/^(.+?)_PL_(\d{4})$/i);
        if (plMatch) {
          deptKey  = plMatch[1].trim();
          fileYear = parseInt(plMatch[2]);
        } else {
          if (/_BS_/i.test(rawName)) continue; // BSファイルスキップ
          const yearMatch = rawName.match(/^(\d{4})[-_](.+)$|^(.+?)[-_](\d{4})$/);
          fileYear = yearMatch ? parseInt(yearMatch[1] || yearMatch[4]) : null;
          deptKey  = yearMatch ? (yearMatch[2] || yearMatch[3]).trim() : rawName;
        }

        if (fileYear !== null && fileYear !== targetYear) continue;

        if (deptKey === '全体') {
          try {
            const csvText     = file.getBlob().getDataAsString('Shift_JIS');
            const monthlyRows = CSVImporter._parseToMonthlyRows(csvText, months);
            SheetManager.writePLSheet(targetYear, '全体', monthlyRows, months);
            consolidatedFromCSV = true;
            Logger.log('✅ 全体 (' + label + '): インポート完了');
            imported++;
          } catch (e) {
            errors.push('"' + fileName + '": ' + e.message);
          }
          continue;
        }

        const dept = CONFIG.DEPARTMENTS.find(function(d) { return d.name === deptKey || d.shortName === deptKey; });
        if (!dept) {
          errors.push('"' + fileName + '": 部門名が一致しません');
          continue;
        }

        try {
          const csvText     = file.getBlob().getDataAsString('Shift_JIS');
          const monthlyRows = CSVImporter._parseToMonthlyRows(csvText, months);
          SheetManager.writePLSheet(targetYear, dept.name, monthlyRows, months);
          allDeptData[dept.name] = monthlyRows;
          Logger.log('✅ ' + dept.name + ' (' + label + '): インポート完了');
          imported++;
        } catch (e) {
          errors.push('"' + fileName + '": ' + e.message);
        }
      }
    } catch (e) {
      return { ok: false, msg: '❌ Driveアクセスエラー: ' + e.message, imported: 0, errors: [] };
    }

    if (Object.keys(allDeptData).length > 0 && !consolidatedFromCSV) {
      try {
        const consolidatedRows = CSVImporter._mergeAllDepts(allDeptData, months);
        SheetManager.writePLSheet(targetYear, '全体', consolidatedRows, months);
      } catch (e) {
        errors.push('"全体合計": ' + e.message);
      }
    }

    const ok  = errors.length === 0;
    const msg = (ok ? '✅ ' : '⚠️ ') + label + ' ' + imported + '件インポート' + (errors.length ? ' / エラー' + errors.length + '件' : '完了');
    Logger.log(msg);
    return { ok, msg, imported, errors, label };
  },

  /**
   * Google DriveフォルダからすべてのCSVをインポートする（メニューから呼ぶ）
   */
  importAllFromDrive() {
    CSVImporter._lastUnmatched = []; // 未照合科目リセット
    const ui = SpreadsheetApp.getUi();
    const folderId = CONFIG.CSV_IMPORT && CONFIG.CSV_IMPORT.FOLDER_ID;

    if (!folderId) {
      ui.alert(
        '⚠️ Config.gs の CSV_IMPORT.FOLDER_ID が設定されていません。\n\n' +
        '1. Googleドライブでインポート用フォルダを作成\n' +
        '2. フォルダのURLからIDをコピー\n' +
        '   例: drive.google.com/drive/folders/[このID部分]\n' +
        '3. Config.gs の CSV_IMPORT.FOLDER_ID にセット'
      );
      return;
    }

    const currYear = getCurrentFiscalYear();
    const result = ui.prompt(
      '📥 部門別CSVインポート',
      `インポートする事業年度の開始年を入力してください。\n` +
      `例: ${currYear - 1} → ${getFiscalPeriodLabel(currYear - 1)}\n` +
      `（空白のまま OK を押すと ${currYear} 年度を使用）`,
      ui.ButtonSet.OK_CANCEL
    );
    if (result.getSelectedButton() !== ui.Button.OK) return;

    const targetYear = parseInt(result.getResponseText()) || currYear;
    const months     = getFiscalMonths(targetYear);
    const label      = getFiscalPeriodLabel(targetYear);

    let imported        = 0;
    const errors        = [];
    const allDeptData   = {};
    let consolidatedFromCSV = false;

    try {
      const folder = DriveApp.getFolderById(folderId);
      const files  = folder.getFiles();

      while (files.hasNext()) {
        const file     = files.next();
        const fileName = file.getName();

        if (!/\.(csv|tsv|txt)$/i.test(fileName)) continue;

        // ファイル名から年度と部門名を抽出
        // 主形式: 物販_PL_2026.csv → deptKey=物販, fileYear=2026
        // 旧形式: 民泊_2025.csv / 2024_全体.csv も引き続き対応
        const rawName  = fileName.replace(/\.(csv|tsv|txt)$/i, '').trim();
        let fileYear, deptKey;
        const plMatch  = rawName.match(/^(.+?)_PL_(\d{4})$/i);
        if (plMatch) {
          deptKey  = plMatch[1].trim();
          fileYear = parseInt(plMatch[2]);
        } else {
          // BSファイル（例: 全体_BS_2026.csv）はPLインポート対象外のためスキップ
          if (/_BS_/i.test(rawName)) {
            Logger.log(`BSファイルをスキップ (PLインポート対象外): ${fileName}`);
            continue;
          }
          const yearMatch = rawName.match(/^(\d{4})[-_](.+)$|^(.+?)[-_](\d{4})$/);
          fileYear = yearMatch ? parseInt(yearMatch[1] || yearMatch[4]) : null;
          deptKey  = yearMatch ? (yearMatch[2] || yearMatch[3]).trim() : rawName;
        }

        if (fileYear !== null && fileYear !== targetYear) continue;

        // 「全体」ファイルは直接 PL_全体 シートへ書き込む
        if (deptKey === '全体') {
          try {
            const csvText     = file.getBlob().getDataAsString('Shift_JIS');
            const monthlyRows = CSVImporter._parseToMonthlyRows(csvText, months);
            SheetManager.writePLSheet(targetYear, '全体', monthlyRows, months);
            consolidatedFromCSV = true;
            Logger.log(`✅ 全体 (${label}): CSVから直接インポート完了`);
            imported++;
          } catch (e) {
            errors.push(`"${fileName}": ${e.message}`);
            Logger.log(`❌ ${fileName}: ${e.message}`);
          }
          continue;
        }

        const dept = CONFIG.DEPARTMENTS.find(d => d.name === deptKey || d.shortName === deptKey);

        if (!dept) {
          errors.push(`"${fileName}": 部門名が一致しません（Config.gs の DEPARTMENTS を確認）`);
          continue;
        }

        try {
          const csvText     = file.getBlob().getDataAsString('Shift_JIS');
          const monthlyRows = CSVImporter._parseToMonthlyRows(csvText, months);
          SheetManager.writePLSheet(targetYear, dept.name, monthlyRows, months);
          allDeptData[dept.name] = monthlyRows;
          Logger.log(`✅ ${dept.name} (${label}): インポート完了`);
          imported++;
        } catch (e) {
          errors.push(`"${fileName}": ${e.message}`);
          Logger.log(`❌ ${fileName}: ${e.message}`);
        }
      }
    } catch (e) {
      ui.alert(`❌ Driveフォルダアクセスエラー: ${e.message}\n\nCSV_IMPORT.FOLDER_ID を確認してください。`);
      return;
    }

    // 部門CSVから全体を合算生成（全体CSVが明示的になかった場合のみ）
    if (Object.keys(allDeptData).length > 0 && !consolidatedFromCSV) {
      try {
        const consolidatedRows = CSVImporter._mergeAllDepts(allDeptData, months);
        SheetManager.writePLSheet(targetYear, '全体', consolidatedRows, months);
        Logger.log(`✅ 全体 (${label}): 部門合算シート生成完了`);
      } catch (e) {
        errors.push(`"全体合計": ${e.message}`);
        Logger.log(`❌ 全体合計: ${e.message}`);
      }
    }

    let msg = `✅ ${label} — ${imported}件をインポートしました。`;
    if (errors.length) msg += `\n\n⚠️ エラー（${errors.length}件）:\n` + errors.join('\n');
    if (CSVImporter._lastUnmatched && CSVImporter._lastUnmatched.length > 0) {
      msg += `\n\n📋 PL_STRUCTUREに未登録の勘定科目（Config.gsへの追加を検討）:\n` +
             CSVImporter._lastUnmatched.map(n => `  ・${n}`).join('\n');
    }
    ui.alert(msg);
  },

  /**
   * CSVのフォーマットを診断して内容を確認用シートに出力する（デバッグ用）
   */
  previewCSV() {
    const ui = SpreadsheetApp.getUi();
    const folderId = CONFIG.CSV_IMPORT && CONFIG.CSV_IMPORT.FOLDER_ID;
    if (!folderId) {
      ui.alert('⚠️ Config.gs の CSV_IMPORT.FOLDER_ID を設定してください。');
      return;
    }

    try {
      const folder = DriveApp.getFolderById(folderId);
      const files  = folder.getFiles();
      if (!files.hasNext()) {
        ui.alert('⚠️ フォルダにファイルがありません。');
        return;
      }

      const file    = files.next();
      const csvText = file.getBlob().getDataAsString('Shift_JIS');
      const lines   = csvText.split(/\r?\n/).filter(l => l.trim());

      const ss = SheetManager.getSpreadsheet();
      let sheet = ss.getSheetByName('_CSVプレビュー');
      if (!sheet) sheet = ss.insertSheet('_CSVプレビュー');
      sheet.clearContents();

      sheet.getRange(1, 1).setValue(`CSVプレビュー: ${file.getName()}\n解析行数: ${lines.length}`);
      sheet.getRange(1, 1).setWrap(true);

      const previewRows = lines.slice(0, 20).map(l => [l]);
      sheet.getRange(3, 1, previewRows.length, 1).setValues(previewRows);
      sheet.setColumnWidth(1, 800);

      const months = getFiscalMonths(getCurrentFiscalYear() - 1);
      const { numDataCols, monthOffsets } = CSVImporter._parseHeader(CSVImporter._splitLine(lines[0]), months);

      sheet.getRange(25, 1).setValue(`ヘッダー解析結果:\n  データ列数: ${numDataCols}\n  月ラベルオフセット: ${JSON.stringify(monthOffsets)}`);
      sheet.getRange(25, 1).setWrap(true);

      sheet.getRange(30, 1, 1, 3).setValues([['勘定科目名', '（1月）', '（2月）']]);
      const accountRows = [];
      for (let i = 1; i < lines.length; i++) {
        const cells = CSVImporter._splitLine(lines[i]);
        if (CSVImporter._getRowLevel(cells) !== 1) continue;
        const name = cells[1].trim();
        const data = cells.slice(cells.length - numDataCols);
        const jan  = data[monthOffsets['1月']] || 0;
        const feb  = data[monthOffsets['2月']] || 0;
        accountRows.push([name, jan, feb]);
        if (accountRows.length >= 30) break;
      }
      if (accountRows.length) {
        sheet.getRange(31, 1, accountRows.length, 3).setValues(accountRows);
      }

      ui.alert(`✅ "_CSVプレビュー" シートにプレビューを出力しました。\n\nファイル: ${file.getName()}\n検出勘定科目: ${accountRows.length}件`);
    } catch (e) {
      ui.alert(`❌ エラー: ${e.message}`);
    }
  },

  // ==============================
  // 内部処理
  // ==============================

  _parseToMonthlyRows(csvText, months) {
    const lines = csvText.split(/\r?\n/).filter(l => l.trim());
    if (lines.length < 2) throw new Error('CSVが空または不正です');

    const headerCells = CSVImporter._splitLine(lines[0]);
    const { numDataCols, monthOffsets } = CSVImporter._parseHeader(headerCells, months);

    if (Object.keys(monthOffsets).length === 0) {
      throw new Error(
        'ヘッダー行から月ラベルが検出できません。\n' +
        'ファイルが MF会計 推移試算表の形式か確認してください。\n' +
        `ヘッダー先頭: ${headerCells.slice(0, 5).join(' | ')}`
      );
    }

    const accountMonthly = {};

    for (let i = 1; i < lines.length; i++) {
      const cells = CSVImporter._splitLine(lines[i]);
      if (CSVImporter._getRowLevel(cells) !== 1) continue;

      const accountName = cells[1].trim();
      if (!accountName) continue;

      const dataValues = cells.slice(cells.length - numDataCols);
      accountMonthly[accountName] = {};

      // 月別金額を取得
      months.forEach(m => {
        const offset = monthOffsets[m.label];
        if (offset === undefined) return;
        const raw = (dataValues[offset] || '').replace(/,/g, '');
        accountMonthly[accountName][m.label] = parseInt(raw) || 0;
      });

      // 決算整理列も取得
      const adjOffset = monthOffsets['決算整理'];
      if (adjOffset !== undefined) {
        const raw = (dataValues[adjOffset] || '').replace(/,/g, '');
        accountMonthly[accountName]['決算整理'] = parseInt(raw) || 0;
      }
    }

    Logger.log(`CSV解析: ${Object.keys(accountMonthly).length}科目 → ${Object.keys(accountMonthly).slice(0, 5).join(', ')}...`);

    // PL_STRUCTURE に照合されない科目を警告（デバッグ用）
    const allMappedNames = new Set();
    CONFIG.PL_STRUCTURE.forEach(item => {
      if (item.accountNames) item.accountNames.forEach(n => allMappedNames.add(n));
    });
    const unmatchedAccounts = Object.keys(accountMonthly).filter(name => !allMappedNames.has(name));
    if (unmatchedAccounts.length > 0) {
      Logger.log(`⚠️ PL_STRUCTUREに未登録の勘定科目:\n  ${unmatchedAccounts.join('\n  ')}`);
    }
    CSVImporter._lastUnmatched = unmatchedAccounts;

    // 月別PLRowsを生成
    const monthlyRows = {};
    months.forEach(m => {
      const fakeItems = Object.entries(accountMonthly).map(([name, monthData]) => ({
        name,
        type:   'account',
        values: [0, 0, 0, monthData[m.label] || 0, 0],
        rows:   null,
      }));
      monthlyRows[m.label] = PLFormatter.buildPLRows(fakeItems);
    });

    // 決算整理列のPLRowsを生成
    if (monthOffsets['決算整理'] !== undefined) {
      const adjItems = Object.entries(accountMonthly).map(([name, monthData]) => ({
        name,
        type:   'account',
        values: [0, 0, 0, monthData['決算整理'] || 0, 0],
        rows:   null,
      }));
      monthlyRows['決算整理'] = PLFormatter.buildPLRows(adjItems);
    }

    return monthlyRows;
  },

  _parseHeader(headerCells, months) {
    const monthLabels = new Set(months.map(m => m.label));
    const firstMonthOffset = headerCells.findIndex(c => monthLabels.has(c.trim()));

    if (firstMonthOffset < 0) {
      return { numDataCols: 14, firstMonthOffset: 2, monthOffsets: {} };
    }

    const numDataCols = headerCells.length - firstMonthOffset;
    const monthOffsets = {};
    months.forEach((m, i) => {
      const headerIdx = firstMonthOffset + i;
      if (headerCells[headerIdx] && headerCells[headerIdx].trim() === m.label) {
        monthOffsets[m.label] = i;
      }
    });

    // 決算整理列を検出（12ヶ月の後ろにある）
    headerCells.forEach((cell, idx) => {
      if (idx >= firstMonthOffset && cell.trim() === '決算整理') {
        monthOffsets['決算整理'] = idx - firstMonthOffset;
      }
    });

    return { numDataCols, firstMonthOffset, monthOffsets };
  },

  _mergeAllDepts(allDeptData, months) {
    const numItems     = CONFIG.PL_STRUCTURE.length;
    const consolidated = {};

    // 月別データをマージ
    months.forEach(m => {
      const mergedRows = Array.from({length: numItems}, (_, i) => {
        const item = CONFIG.PL_STRUCTURE[i];
        if (item.category === 'header') return { amount: null };

        let total = 0;
        Object.values(allDeptData).forEach(monthlyRows => {
          const plRows = monthlyRows[m.label] || [];
          const plRow  = plRows[i];
          if (plRow && plRow.amount !== null && plRow.amount !== undefined) {
            total += plRow.amount || 0;
          }
        });
        return { amount: total };
      });
      consolidated[m.label] = mergedRows;
    });

    // 決算整理列もマージ
    const hasAdj = Object.values(allDeptData).some(d => d['決算整理']);
    if (hasAdj) {
      consolidated['決算整理'] = Array.from({length: numItems}, (_, i) => {
        const item = CONFIG.PL_STRUCTURE[i];
        if (item.category === 'header') return { amount: null };
        let total = 0;
        Object.values(allDeptData).forEach(monthlyRows => {
          const plRows = monthlyRows['決算整理'] || [];
          const plRow  = plRows[i];
          if (plRow && plRow.amount !== null && plRow.amount !== undefined) {
            total += plRow.amount || 0;
          }
        });
        return { amount: total };
      });
    }

    return consolidated;
  },

  _splitLine(line) {
    if (line.indexOf('\t') >= 0) return line.split('\t');

    const parts = [];
    let i = 0;
    while (i < line.length) {
      if (line[i] === '"') {
        const end = line.indexOf('"', i + 1);
        parts.push(end >= 0 ? line.slice(i + 1, end) : line.slice(i + 1));
        i = end >= 0 ? end + 2 : line.length;
      } else {
        const end = line.indexOf(',', i);
        parts.push(end >= 0 ? line.slice(i, end) : line.slice(i));
        i = end >= 0 ? end + 1 : line.length;
      }
    }
    return parts;
  },

  _getRowLevel(cells) {
    if (cells[0] && cells[0].trim()) return 0;
    if (cells.length > 1 && cells[1] && cells[1].trim()) return 1;
    if (cells.length > 2 && cells[2] && cells[2].trim()) return 2;
    return -1;
  },
};
