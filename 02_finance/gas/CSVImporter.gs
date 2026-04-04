/**
 * CSVインポーター - MF会計 部門別推移試算表TSVを読み込んでPLシートに反映
 *
 * ■ 使い方
 *   1. MF会計 → レポート → 推移試算表 → 部門選択 → 全期間 → CSVダウンロード
 *   2. Googleドライブの指定フォルダ（Config.gs の CSV_IMPORT.FOLDER_ID）にアップロード
 *      ファイル名 = 部門名.csv（例: 民泊.csv / 物販.csv / ブランド.csv / 共通.csv）
 *   3. スプレッドシートのメニュー「📥 部門別CSVをインポート」を実行
 *
 * ■ MF会計 推移試算表のCSVフォーマット
 *   ヘッダー行: 勘定科目[tab]補助科目[tab]3月[tab]4月...[tab]2月[tab]決算整理[tab]合計
 *   データ階層（タブでインデント）:
 *     Level 0（col 0が非空）: カテゴリ行・小計行 → スキップ
 *     Level 1（col 1が非空）: 勘定科目行 ← ここを使用
 *     Level 2（col 2が非空）: 補助科目行 → スキップ（Level 1が合計値）
 *   データ列は常に行末から N 列（ヘッダーから月ラベルが始まる列以降）
 */

const CSVImporter = {

  /**
   * Google DriveフォルダからすべてのCSVをインポートする（メニューから呼ぶ）
   */
  importAllFromDrive() {
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

    // 対象年度を選択
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

    let imported = 0;
    const errors  = [];

    try {
      const folder = DriveApp.getFolderById(folderId);
      const files  = folder.getFiles();

      while (files.hasNext()) {
        const file     = files.next();
        const fileName = file.getName();

        // .csv / .tsv / .txt のみ対象
        if (!/\.(csv|tsv|txt)$/i.test(fileName)) continue;

        const deptName = fileName.replace(/\.(csv|tsv|txt)$/i, '').trim();
        const dept     = CONFIG.DEPARTMENTS.find(d => d.name === deptName || d.shortName === deptName);

        if (!dept) {
          errors.push(`"${fileName}": 部門名が一致しません（Config.gs の DEPARTMENTS を確認）`);
          continue;
        }

        try {
          const csvText    = file.getBlob().getDataAsString('UTF-8');
          const monthlyRows = CSVImporter._parseToMonthlyRows(csvText, months);
          SheetManager.writePLSheet(targetYear, dept.name, monthlyRows, months);
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

    let msg = `✅ ${label} — ${imported}部門をインポートしました。`;
    if (errors.length) msg += `\n\n⚠️ エラー（${errors.length}件）:\n` + errors.join('\n');
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

      // 最初のファイルをプレビュー
      const file    = files.next();
      const csvText = file.getBlob().getDataAsString('UTF-8');
      const lines   = csvText.split(/\r?\n/).filter(l => l.trim());

      const ss = SheetManager.getSpreadsheet();
      let sheet = ss.getSheetByName('_CSVプレビュー');
      if (!sheet) sheet = ss.insertSheet('_CSVプレビュー');
      sheet.clearContents();

      sheet.getRange(1, 1).setValue(`CSVプレビュー: ${file.getName()}\n解析行数: ${lines.length}`);
      sheet.getRange(1, 1).setWrap(true);

      // 先頭20行を表示
      const previewRows = lines.slice(0, 20).map(l => [l]);
      sheet.getRange(3, 1, previewRows.length, 1).setValues(previewRows);
      sheet.setColumnWidth(1, 800);

      // 勘定科目として検出された行を表示
      const months = getFiscalMonths(getCurrentFiscalYear() - 1);
      const { numDataCols, firstMonthOffset, monthOffsets } = CSVImporter._parseHeader(lines[0].split('\t'), months);

      sheet.getRange(25, 1).setValue(`ヘッダー解析結果:\n  データ列数: ${numDataCols}\n  月ラベルオフセット: ${JSON.stringify(monthOffsets)}`);
      sheet.getRange(25, 1).setWrap(true);

      // Level 2行（勘定科目）を一覧表示
      sheet.getRange(30, 1, 1, 3).setValues([['勘定科目名', '（1月）', '（2月）']]);
      const accountRows = [];
      for (let i = 1; i < lines.length; i++) {
        const cells = lines[i].split('\t');
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

  /**
   * MF会計 推移試算表CSVを月別PLRowsに変換する
   *
   * @param {string} csvText - CSVテキスト
   * @param {Array}  months  - getFiscalMonths() の戻り値
   * @return {Object} { monthLabel: [PLrows], ... }
   */
  _parseToMonthlyRows(csvText, months) {
    const lines = csvText.split(/\r?\n/).filter(l => l.trim());
    if (lines.length < 2) throw new Error('CSVが空または不正です');

    const headerCells = lines[0].split('\t');
    const { numDataCols, monthOffsets } = CSVImporter._parseHeader(headerCells, months);

    if (Object.keys(monthOffsets).length === 0) {
      throw new Error(
        'ヘッダー行から月ラベルが検出できません。\n' +
        'ファイルが MF会計 推移試算表の形式か確認してください。\n' +
        `ヘッダー先頭: ${headerCells.slice(0, 5).join(' | ')}`
      );
    }

    // 勘定科目 → 月別金額 のマップを構築
    const accountMonthly = {}; // { accountName: { monthLabel: number } }

    for (let i = 1; i < lines.length; i++) {
      const cells = lines[i].split('\t');
      if (CSVImporter._getRowLevel(cells) !== 1) continue;

      const accountName = cells[1].trim();
      if (!accountName) continue;

      const dataValues = cells.slice(cells.length - numDataCols);
      accountMonthly[accountName] = {};

      months.forEach(m => {
        const offset = monthOffsets[m.label];
        if (offset === undefined) return;
        const raw = (dataValues[offset] || '').replace(/,/g, '');
        accountMonthly[accountName][m.label] = parseInt(raw) || 0;
      });
    }

    Logger.log(`CSV解析: ${Object.keys(accountMonthly).length}科目 → ${Object.keys(accountMonthly).slice(0, 5).join(', ')}...`);

    // 月別PLRowsを生成（PLFormatter.buildPLRows に渡せる形式に変換）
    const monthlyRows = {};
    months.forEach(m => {
      const fakeItems = Object.entries(accountMonthly).map(([name, monthData]) => ({
        name,
        type:   'account',
        values: [0, 0, 0, Math.abs(monthData[m.label] || 0), 0],
        rows:   null,
      }));
      monthlyRows[m.label] = PLFormatter.buildPLRows(fakeItems);
    });

    return monthlyRows;
  },

  /**
   * ヘッダー行を解析して月オフセットマップを返す
   *
   * @param {Array}  headerCells - ヘッダー行をタブ分割した配列
   * @param {Array}  months      - getFiscalMonths() の結果
   * @return {{ numDataCols, firstMonthOffset, monthOffsets }}
   *   numDataCols:     データ列の総数（月 + 決算整理 + 合計）
   *   firstMonthOffset: ヘッダー内で最初の月ラベルが出る列インデックス
   *   monthOffsets:   { monthLabel: データ配列内のオフセット }
   */
  _parseHeader(headerCells, months) {
    // ヘッダーの中から月ラベルが最初に現れる位置を探す
    const monthLabels = new Set(months.map(m => m.label));
    const firstMonthOffset = headerCells.findIndex(c => monthLabels.has(c.trim()));

    if (firstMonthOffset < 0) {
      return { numDataCols: 14, firstMonthOffset: 2, monthOffsets: {} };
    }

    // データ列数 = ヘッダーの残り列数（月ラベル以降すべて）
    const numDataCols = headerCells.length - firstMonthOffset;

    // 月ラベル → データ配列内のオフセット
    const monthOffsets = {};
    months.forEach((m, i) => {
      const headerIdx = firstMonthOffset + i;
      if (headerCells[headerIdx] && headerCells[headerIdx].trim() === m.label) {
        monthOffsets[m.label] = i;
      }
    });

    return { numDataCols, firstMonthOffset, monthOffsets };
  },

  /**
   * 行の階層レベルを判定する（先頭の空セル数で判定）
   * @return {number} 0=カテゴリ/小計, 1=勘定科目, 2=補助科目, -1=空行
   */
  _getRowLevel(cells) {
    if (cells[0] && cells[0].trim()) return 0;
    if (cells.length > 1 && cells[1] && cells[1].trim()) return 1;
    if (cells.length > 2 && cells[2] && cells[2].trim()) return 2;
    return -1;
  },
};
