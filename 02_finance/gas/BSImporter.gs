/**
 * BSインポーター - MF会計 貸借対照表CSVを読み込んで通期比較シートに反映
 *
 * ■ 使い方
 *   1. MF会計 → レポート → 貸借対照表 → 期間指定（事業年度全体）→ CSVダウンロード
 *   2. Googleドライブの指定フォルダ（Config.gs の CSV_IMPORT.FOLDER_ID）にアップロード
 *      ファイル名 = 全体_BS_{年}.csv（例: 全体_BS_2026.csv）
 *   3. スプレッドシートのメニュー「📊 BS通期比較シートを作成」を実行
 *
 * ■ MF会計 貸借対照表のCSV列構成
 *   A列(0): カテゴリ・合計行（Level 0）
 *   B列(1): 勘定科目行（Level 1）← 使用
 *   C列(2): 補助科目行（Level 2）← スキップ
 *   D列(3): 期首残高
 *   E列(4): 期間借方金額
 *   F列(5): 期間貸方金額
 *   G列(6): 期末残高（ヘッダー「終了月:YYYY-MM」）← 使用
 *   H列(7): 構成比
 */

const BSImporter = {

  /**
   * Google DriveフォルダからすべてのBS CSVを読み込んで返す
   * @returns {{ [fiscalYear]: { level0: {label:balance}, level1: {label:balance} } }}
   */
  importAllFromDrive() {
    const folderId = CONFIG.CSV_IMPORT && CONFIG.CSV_IMPORT.FOLDER_ID;
    if (!folderId) {
      Logger.log('⚠️ Config.gs の CSV_IMPORT.FOLDER_ID を設定してください。');
      return {};
    }

    const BASE_YEAR  = 2018;
    const currYear   = getCurrentFiscalYear();
    const yearDataMap = {};

    try {
      const folder = DriveApp.getFolderById(folderId);
      const files  = folder.getFiles();

      while (files.hasNext()) {
        const file     = files.next();
        const fileName = file.getName();
        if (!/\.(csv|tsv|txt)$/i.test(fileName)) continue;

        // ファイル名パース: 全体_BS_2026.csv
        const rawName = fileName.replace(/\.(csv|tsv|txt)$/i, '').trim();
        const bsMatch = rawName.match(/^(.+?)_BS_(\d{4})$/i);
        if (!bsMatch) continue; // PL等BSファイル以外はスキップ

        const fileYear = parseInt(bsMatch[2]);
        if (fileYear < BASE_YEAR || fileYear > currYear) continue;

        try {
          const csvText = file.getBlob().getDataAsString('Shift_JIS');
          yearDataMap[fileYear] = BSImporter._parseToBalances(csvText);
          Logger.log(`✅ BS ${getFiscalPeriodLabel(fileYear)}: level0=${Object.keys(yearDataMap[fileYear].level0).length}件, level1=${Object.keys(yearDataMap[fileYear].level1).length}件`);
        } catch (e) {
          Logger.log(`❌ ${fileName}: ${e.message}`);
        }
      }
    } catch (e) {
      Logger.log(`❌ Driveアクセスエラー: ${e.message}`);
      return {};
    }

    Logger.log(`BS CSVインポート完了: ${Object.keys(yearDataMap).length}期分`);
    return yearDataMap;
  },

  /**
   * BS CSVを解析して期末残高マップを返す
   * @param {string} csvText - Shift_JIS → string変換済みCSV
   * @returns {{ level0: {label:number}, level1: {label:number} }}
   */
  _parseToBalances(csvText) {
    const lines = csvText.split(/\r?\n/).filter(l => l.trim());
    if (lines.length < 2) throw new Error('CSVが空または不正です');

    // ヘッダー行から期末残高列を検出（"終了月:" で始まる列）
    const headerCells    = BSImporter._splitLine(lines[0]);
    const balanceColIdx  = headerCells.findIndex(c => c.trim().startsWith('終了月:'));
    if (balanceColIdx < 0) throw new Error(
      '期末残高列（終了月:）が見つかりません。\n' +
      `ヘッダー: ${headerCells.slice(0, 8).join(' | ')}`
    );

    const level0Map = {}; // A列（カテゴリ・合計行）
    const level1Map = {}; // B列（勘定科目行）

    for (let i = 1; i < lines.length; i++) {
      const cells = BSImporter._splitLine(lines[i]);
      const level = BSImporter._getRowLevel(cells);
      if (level < 0) continue;

      const raw = (cells[balanceColIdx] || '').replace(/,/g, '').trim();
      const val = parseInt(raw) || 0;

      if (level === 0) {
        const label = cells[0].trim();
        if (label) level0Map[label] = val;
      } else if (level === 1) {
        const label = cells[1].trim();
        if (label) level1Map[label] = val;
      }
      // level === 2 → 補助科目、スキップ
    }

    return { level0: level0Map, level1: level1Map };
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
