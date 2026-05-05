/**
 * Amazon Dashboard - スプレッドシート書き込みモジュール
 */

/**
 * メインスプレッドシートを取得
 */
function getMainSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  return SpreadsheetApp.openById(getMainSpreadsheetId());
}

/**
 * 指定シートを取得（なければ作成）
 */
function getOrCreateSheet(sheetName) {
  const ss = getMainSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

/**
 * D1 日次データにヘッダーを設定
 */
function setupDailyDataHeaders() {
  const sheet = getOrCreateSheet(SHEET_NAMES.D1_DAILY);
  const headers = [
    '日付', 'ASIN', '商品名', 'カテゴリ',
    '売上金額', 'CV(注文件数)', '注文点数',
    'セッション数', 'PV', 'CTR(%)', 'CVR(%)', 'BuyBox率(%)',
    'FBA手数料', '返品数', '返品額',
    '広告費', '広告売上', 'IMP', 'CT',
    '仕入単価', '仕入原価合計', 'ステータス'
  ];

  const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (existing[0] !== headers[0]) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    Logger.log('✅ D1 日次データ: ヘッダー設定完了');
  }
  return sheet;
}

/**
 * M1 商品マスターにヘッダーを設定（12列構成）
 *
 * 列構造:
 *   A: ASIN
 *   B: 商品名
 *   C: カテゴリ
 *   D: ステータス
 *   E: 仕入単価           （手入力 / D1 仕入原価バックフィルのソース）
 *   F: 仕入れ先
 *   G: 備考
 *   H: 販売単価           （手入力）
 *   I: 販売手数料率        （手入力 / カテゴリ別レート 例: 8%）
 *   J: 販売手数料          （数式 =H*I）
 *   K: FBA手数料           （手入力）
 *   L: 粗利単価           （数式 =H-E-J-K）
 */
const M1_HEADERS = [
  'ASIN', '商品名', 'カテゴリ', 'ステータス',
  '仕入単価', '仕入れ先', '備考',
  '販売単価', '販売手数料率', '販売手数料', 'FBA手数料', '粗利単価',
];

const M1_COLS = {
  ASIN: 1, NAME: 2, CATEGORY: 3, STATUS: 4,
  PURCHASE_PRICE: 5, SUPPLIER: 6, NOTE: 7,
  SELL_PRICE: 8, COMMISSION_RATE: 9, COMMISSION: 10, FBA_FEE: 11, GROSS_PROFIT: 12,
};

function setupProductMasterHeaders() {
  const sheet = getOrCreateSheet(SHEET_NAMES.M1_PRODUCT_MASTER);

  const existing = sheet.getRange(1, 1, 1, M1_HEADERS.length).getValues()[0];
  const needsUpdate = existing[0] !== M1_HEADERS[0] ||
                      existing[M1_COLS.SELL_PRICE - 1] !== '販売単価';

  if (needsUpdate) {
    sheet.getRange(1, 1, 1, M1_HEADERS.length).setValues([M1_HEADERS])
      .setFontWeight('bold').setBackground('#e8f0fe');
    sheet.setFrozenRows(1);
    Logger.log('✅ M1 商品マスター: ヘッダー設定完了（12列）');
  }

  // 既存データ行に販売手数料・粗利単価の数式を流し込む（空の行はスキップ）
  applyProductMasterFormulas(sheet);

  return sheet;
}

/**
 * M1 の販売手数料(J) / 粗利単価(L) に数式を流し込む。
 * ASIN が入っている全行に対して適用。手入力列が空でも数式自体は維持する。
 */
function applyProductMasterFormulas(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const asinCol = sheet.getRange(2, M1_COLS.ASIN, lastRow - 1, 1).getValues();
  const formulaJ = [];
  const formulaL = [];
  for (let i = 0; i < asinCol.length; i++) {
    const r = i + 2;
    const hasAsin = String(asinCol[i][0] || '').trim() !== '';
    formulaJ.push([hasAsin ? '=IFERROR(H' + r + '*I' + r + ', "")' : '']);
    formulaL.push([hasAsin ? '=IFERROR(H' + r + '-E' + r + '-J' + r + '-K' + r + ', "")' : '']);
  }
  sheet.getRange(2, M1_COLS.COMMISSION, formulaJ.length, 1).setFormulas(formulaJ);
  sheet.getRange(2, M1_COLS.GROSS_PROFIT, formulaL.length, 1).setFormulas(formulaL);

  // 表示フォーマット
  sheet.getRange(2, M1_COLS.PURCHASE_PRICE, lastRow - 1, 1).setNumberFormat('#,##0');
  sheet.getRange(2, M1_COLS.SELL_PRICE, lastRow - 1, 1).setNumberFormat('#,##0');
  sheet.getRange(2, M1_COLS.COMMISSION_RATE, lastRow - 1, 1).setNumberFormat('0.0%');
  sheet.getRange(2, M1_COLS.COMMISSION, lastRow - 1, 1).setNumberFormat('#,##0');
  sheet.getRange(2, M1_COLS.FBA_FEE, lastRow - 1, 1).setNumberFormat('#,##0');
  sheet.getRange(2, M1_COLS.GROSS_PROFIT, lastRow - 1, 1).setNumberFormat('#,##0');

  Logger.log('✅ M1 数式適用: ' + formulaJ.length + ' 行');
}

/**
 * D2 経費明細にヘッダーを設定
 */
function setupSettlementHeaders() {
  const sheet = getOrCreateSheet(SHEET_NAMES.D2_SETTLEMENT);
  const headers = [
    '決済期間開始', '決済期間終了', '日付', 'ASIN',
    'トランザクション種別', '明細種別', '金額', '数量'
  ];

  const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (existing[0] !== headers[0]) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    Logger.log('✅ D2 経費明細: ヘッダー設定完了');
  }
  return sheet;
}

/**
 * 全シートのヘッダーを初期設定
 */
function setupAllHeaders() {
  setupDailyDataHeaders();
  setupProductMasterHeaders();
  setupSettlementHeaders();
  Logger.log('✅ 全シートのヘッダー設定完了');
}

/**
 * シートの最終行にデータを追記
 */
function appendRows(sheetName, rows) {
  if (!rows || rows.length === 0) return;
  const sheet = getOrCreateSheet(sheetName);
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
  Logger.log(sheetName + ': ' + rows.length + ' 行追記');
}

/**
 * 指定シートを「最小サイズ」で新規作成 or 既存を返す
 * getOrCreateSheet は default 1000行×26列 のシートを作るが、
 * セル数上限（1000万）に近い場合は最小サイズで作る必要がある。
 */
function getOrCreateSheetCompact(sheetName, cols, rows) {
  cols = cols || 10;
  rows = rows || 100;
  const ss = getMainSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) return sheet;

  sheet = ss.insertSheet(sheetName);
  const maxCols = sheet.getMaxColumns();
  if (maxCols > cols) sheet.deleteColumns(cols + 1, maxCols - cols);
  const maxRows = sheet.getMaxRows();
  if (maxRows > rows) sheet.deleteRows(rows + 1, maxRows - rows);
  return sheet;
}

/**
 * 全シートのサイズ（行×列）を一覧表示してセル総計を算出
 * 容量警告が出たらこれで大きいシートを特定する
 */
function diagnoseSheetSizes() {
  const ss = getMainSpreadsheet();
  const sheets = ss.getSheets();
  let total = 0;
  const results = sheets.map(s => {
    const rows = s.getMaxRows();
    const cols = s.getMaxColumns();
    const cells = rows * cols;
    const lastRow = s.getLastRow();
    const lastCol = s.getLastColumn();
    total += cells;
    return { name: s.getName(), rows, cols, cells, lastRow, lastCol };
  }).sort((a, b) => b.cells - a.cells);

  Logger.log('シート名 | max行 | max列 | セル数 | データ最終行 | データ最終列');
  results.forEach(r => {
    Logger.log(r.name + ' | ' + r.rows + ' | ' + r.cols + ' | ' +
               r.cells.toLocaleString() + ' | ' + r.lastRow + ' | ' + r.lastCol);
  });
  Logger.log('合計: ' + total.toLocaleString() + ' / 10,000,000 (' +
             (total / 10000000 * 100).toFixed(1) + '%)');
  return { total, results };
}

/**
 * 指定シートの「データがある範囲の外」を削除してサイズを圧縮
 * 例: 1000行×26列のシートで実際データが 50行×10列なら、残りを削る
 */
function trimSheet(sheetName, keepExtraRows, keepExtraCols) {
  keepExtraRows = keepExtraRows == null ? 50 : keepExtraRows;
  keepExtraCols = keepExtraCols == null ? 5 : keepExtraCols;
  const ss = getMainSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) { Logger.log('シートなし: ' + sheetName); return; }

  const lastRow = Math.max(1, sheet.getLastRow());
  const lastCol = Math.max(1, sheet.getLastColumn());
  const targetRows = lastRow + keepExtraRows;
  const targetCols = lastCol + keepExtraCols;
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();

  let deletedRows = 0, deletedCols = 0;
  if (maxRows > targetRows) {
    sheet.deleteRows(targetRows + 1, maxRows - targetRows);
    deletedRows = maxRows - targetRows;
  }
  if (maxCols > targetCols) {
    sheet.deleteColumns(targetCols + 1, maxCols - targetCols);
    deletedCols = maxCols - targetCols;
  }
  Logger.log(sheetName + ': -' + deletedRows + ' 行 / -' + deletedCols + ' 列 削除');
}

/**
 * 全シートを一括 trim（データ末尾 + 少し余白、だけ残す）
 * 実行前に必ず diagnoseSheetSizes() で状況確認すること
 */
function trimAllSheets() {
  const ss = getMainSpreadsheet();
  ss.getSheets().forEach(s => {
    trimSheet(s.getName(), 100, 5);
  });
  Logger.log('✅ 全シート trim 完了');
  diagnoseSheetSizes();
}

/**
 * 商品マスターから ASIN → 商品情報のマッピングを取得
 *
 * 返却フィールド:
 *   name / category / status / purchasePrice / sellPrice /
 *   commissionRate / commission / fbaFee / grossProfit
 */
function getProductMasterMap() {
  const sheet = getOrCreateSheet(SHEET_NAMES.M1_PRODUCT_MASTER);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return {};

  const data = sheet.getRange(2, 1, lastRow - 1, M1_HEADERS.length).getValues();
  const map = {};
  data.forEach(row => {
    if (row[0]) {
      map[row[0]] = {
        name: row[M1_COLS.NAME - 1] || '',
        category: row[M1_COLS.CATEGORY - 1] || '',
        status: row[M1_COLS.STATUS - 1] || 'アクティブ',
        purchasePrice: parseFloat(row[M1_COLS.PURCHASE_PRICE - 1]) || 0,
        sellPrice: parseFloat(row[M1_COLS.SELL_PRICE - 1]) || 0,
        commissionRate: parseFloat(row[M1_COLS.COMMISSION_RATE - 1]) || 0,
        commission: parseFloat(row[M1_COLS.COMMISSION - 1]) || 0,
        fbaFee: parseFloat(row[M1_COLS.FBA_FEE - 1]) || 0,
        grossProfit: parseFloat(row[M1_COLS.GROSS_PROFIT - 1]) || 0,
      };
    }
  });
  return map;
}
function syncMasterToDaily() {
  Logger.log('===== 商品マスター → 日次データ 同期開始 =====');

  const masterMap = getProductMasterMap();
  const sheet = getOrCreateSheet(SHEET_NAMES.D1_DAILY);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    Logger.log('データなし');
    return;
  }

  const data = sheet.getRange(2, 2, lastRow - 1, 3).getValues();
  let updated = 0;

  for (let i = 0; i < data.length; i++) {
    const asin = data[i][0];
    if (!asin) continue;

    const master = masterMap[asin];
    if (!master) continue;

    let changed = false;
    if (!data[i][1] && master.name) {
      data[i][1] = master.name;
      changed = true;
    }
    if (!data[i][2] && master.category) {
      data[i][2] = master.category;
      changed = true;
    }
    if (changed) updated++;
  }

  if (updated > 0) {
    sheet.getRange(2, 3, lastRow - 1, 2).setValues(data.map(row => [row[1], row[2]]));
    Logger.log('✅ ' + updated + ' 行の商品名・カテゴリを更新');
  } else {
    Logger.log('更新対象なし');
  }
}
