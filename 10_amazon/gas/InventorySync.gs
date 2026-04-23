/**
 * Amazon Dashboard - 外部スプシ（発注管理表 / CF管理）への在庫同期
 *
 * ## 処理フロー（毎日 10:30）
 *
 *   1. Amazon Inventory API で ASIN別の FBA在庫 (fulfillableQuantity) を取得
 *   2. 発注管理表「発注タイミング」シートの F列「FBA在庫」に書き込み
 *      （ASIN は C列でマッチング）
 *   3. SpreadsheetApp.flush() で D列「在庫数」の数式を再計算
 *   4. 発注管理表の D列「在庫数」を読み取り
 *   5. CF管理「在庫残高」シートで当月の「在庫」行を特定
 *   6. 該当 ASIN 列（E列以降）に D列値を書き込む
 *
 *   → 毎日上書きされるが、月末時点の値が「そのまま次月以降に残り続ける」ため
 *     結果的に「月末最終在庫数」が保存される設計。
 *
 * ## 必要な ScriptProperties
 *   - ORDER_SHEET_ID: 発注管理表スプレッドシートID
 *   - CF_SHEET_ID:    CF管理スプレッドシートID（既存設定を再利用）
 *
 * ## トリガー: 毎日 AM10:30 (syncInventoryToExternalSheets)
 */

const ORDER_SHEET_NAME = '発注タイミング';
const ORDER_SHEET_ASIN_COL = 3;      // C列
const ORDER_SHEET_TOTAL_COL = 4;     // D列「在庫数」
const ORDER_SHEET_SALES_COL = 5;     // E列「販売set（1日）」= 過去7日平均販売数
const ORDER_SHEET_FBA_COL = 6;       // F列「FBA在庫」（在庫あり + FC転送）
const ORDER_SHEET_INBOUND_COL = 7;   // G列「FBA納品中」（inbound working+shipped+receiving）
const ORDER_SHEET_DATA_START_ROW = 2;
const ORDER_SALES_AVG_DAYS = 7;      // 平均販売数の算出期間

const CF_STOCK_SHEET_NAME = '在庫残高';
const CF_STOCK_ASIN_START_COL = 5;   // E列から ASIN別
const CF_STOCK_MONTH_COL = 2;        // B列: 年月（2026.04 等、number型のことが多い）
const CF_STOCK_KBN_COL = 3;          // C列: 区分（在庫/見込売上/...）
const CF_STOCK_KBN_VALUE = '在庫';

/**
 * メイン: 在庫データを発注管理表 + CF管理スプシに同期
 */
function syncInventoryToExternalSheets() {
  const t0 = Date.now();
  Logger.log('===== 外部スプシ 在庫同期 開始 =====');

  // 1. Amazon API から FBA 在庫取得（qty=在庫あり+FC転送 / inbound=納品中）
  const inventory = fetchInventoryData();
  const asinToFba = {};
  const asinToInbound = {};
  for (const inv of inventory) {
    if (!inv.asin) continue;
    // 同一ASINに複数SKUがある場合は合算
    asinToFba[inv.asin] = (asinToFba[inv.asin] || 0) + (inv.qty || 0);
    asinToInbound[inv.asin] = (asinToInbound[inv.asin] || 0) + (inv.inbound || 0);
  }
  Logger.log('Amazon 在庫取得: ' + Object.keys(asinToFba).length + ' ASIN');

  // 2. 発注管理表 F列「FBA在庫」+ G列「FBA納品中」+ E列「販売set(1日)」に書き込み
  const orderSheet = openOrderSheet();
  const updatedFba = writeFbaToOrderSheet(orderSheet, asinToFba);
  Logger.log('発注管理表 F列「FBA在庫」更新: ' + updatedFba + ' 行');
  const updatedInbound = writeInboundToOrderSheet(orderSheet, asinToInbound);
  Logger.log('発注管理表 G列「FBA納品中」更新: ' + updatedInbound + ' 行');

  // E列「販売set(1日)」= 過去7日平均販売数（売れなかった日も0として算入）
  // 発注管理表に記載のある全ASINに対して値を用意（売上なしは0）
  const salesAvgRaw = getSalesAvgByAsin(ORDER_SALES_AVG_DAYS);
  const asinMapForSales = getOrderSheetAsinMap(orderSheet);
  const asinToSalesAvg = {};
  for (const asin of Object.keys(asinMapForSales)) {
    asinToSalesAvg[asin] = salesAvgRaw[asin] || 0;
  }
  const updatedSales = writeSalesAvgToOrderSheet(orderSheet, asinToSalesAvg);
  Logger.log('発注管理表 E列「販売set(1日)」更新: ' + updatedSales + ' 行（' + ORDER_SALES_AVG_DAYS + '日平均）');

  // 3. 式再計算を待つ
  SpreadsheetApp.flush();
  Utilities.sleep(1000);

  // 4. D列「在庫数」を読み取り
  const asinToTotalStock = readOrderSheetTotalStock(orderSheet);
  Logger.log('発注管理表 D列「在庫数」読み取り: ' + Object.keys(asinToTotalStock).length + ' ASIN');

  // 5. CF管理スプシに書き込み
  const cfSheet = openCfStockSheet();
  const yearMonth = formatCurrentYearMonth();
  const written = writeCurrentMonthStockToCf(cfSheet, yearMonth, asinToTotalStock);
  Logger.log('CF管理「在庫残高」更新: ' + yearMonth + ' の在庫行 ' + written + ' 列');

  Logger.log('✅ 外部スプシ同期完了 (' + (Date.now() - t0) + 'ms)');
}

// ===== 発注管理表 =====

function openOrderSheet() {
  const id = PropertiesService.getScriptProperties().getProperty('ORDER_SHEET_ID');
  if (!id) throw new Error('ORDER_SHEET_ID が未設定');
  const ss = SpreadsheetApp.openById(id);
  const sheet = ss.getSheetByName(ORDER_SHEET_NAME);
  if (!sheet) throw new Error('発注管理表に「' + ORDER_SHEET_NAME + '」シートがない');
  return sheet;
}

/**
 * 発注管理表のC列からASIN → 行番号マップを作成
 */
function getOrderSheetAsinMap(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < ORDER_SHEET_DATA_START_ROW) return {};
  const asinCol = sheet.getRange(ORDER_SHEET_DATA_START_ROW, ORDER_SHEET_ASIN_COL,
                                  lastRow - ORDER_SHEET_DATA_START_ROW + 1, 1).getValues();
  const map = {};
  for (let i = 0; i < asinCol.length; i++) {
    const asin = String(asinCol[i][0] || '').trim();
    if (asin) map[asin] = ORDER_SHEET_DATA_START_ROW + i;
  }
  return map;
}

/**
 * F列「FBA在庫」に ASIN別の最新値を書き込み
 * @returns {number} 更新した行数
 */
function writeFbaToOrderSheet(sheet, asinToFba) {
  return writeColumnToOrderSheet(sheet, asinToFba, ORDER_SHEET_FBA_COL);
}

/**
 * G列「FBA納品中」に ASIN別の納品中数量を書き込み
 * @returns {number} 更新した行数
 */
function writeInboundToOrderSheet(sheet, asinToInbound) {
  return writeColumnToOrderSheet(sheet, asinToInbound, ORDER_SHEET_INBOUND_COL);
}

/**
 * E列「販売set（1日）」に 過去7日平均販売数を書き込み
 * @returns {number} 更新した行数
 */
function writeSalesAvgToOrderSheet(sheet, asinToAvg) {
  return writeColumnToOrderSheet(sheet, asinToAvg, ORDER_SHEET_SALES_COL);
}

/**
 * 指定列に ASIN→値マップを書き込む共通ロジック
 * 現在値との差分だけ更新することで再計算コストを抑える。
 */
function writeColumnToOrderSheet(sheet, asinToValue, colIndex) {
  const asinMap = getOrderSheetAsinMap(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < ORDER_SHEET_DATA_START_ROW) return 0;

  const range = sheet.getRange(ORDER_SHEET_DATA_START_ROW, colIndex,
                               lastRow - ORDER_SHEET_DATA_START_ROW + 1, 1);
  const current = range.getValues();
  let updated = 0;
  for (const [asin, qty] of Object.entries(asinToValue)) {
    const row = asinMap[asin];
    if (!row) continue;
    const idx = row - ORDER_SHEET_DATA_START_ROW;
    if (current[idx][0] !== qty) {
      current[idx][0] = qty;
      updated++;
    }
  }
  if (updated > 0) range.setValues(current);
  return updated;
}

/**
 * D列「在庫数」を読み取って ASIN → 在庫数 のマップを返す
 */
function readOrderSheetTotalStock(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < ORDER_SHEET_DATA_START_ROW) return {};

  // C列（ASIN）と D列（在庫数）を一括取得
  const data = sheet.getRange(ORDER_SHEET_DATA_START_ROW, ORDER_SHEET_ASIN_COL,
                              lastRow - ORDER_SHEET_DATA_START_ROW + 1, 2).getValues();
  const map = {};
  for (const row of data) {
    const asin = String(row[0] || '').trim();
    const qty = parseFloat(row[1]) || 0;
    if (asin) map[asin] = qty;
  }
  return map;
}

// ===== CF管理「在庫残高」 =====

function openCfStockSheet() {
  const id = PropertiesService.getScriptProperties().getProperty('CF_SHEET_ID');
  if (!id) throw new Error('CF_SHEET_ID が未設定');
  const ss = SpreadsheetApp.openById(id);
  const sheet = ss.getSheetByName(CF_STOCK_SHEET_NAME);
  if (!sheet) throw new Error('CF管理に「' + CF_STOCK_SHEET_NAME + '」シートがない');
  return sheet;
}

function formatCurrentYearMonth() {
  const d = new Date();
  return d.getFullYear() + '.' + String(d.getMonth() + 1).padStart(2, '0');
}

/**
 * セル値（文字列 / Date / 数値）を "YYYY.MM" 形式に正規化
 * 受け入れフォーマット:
 *   Date                    → "YYYY.MM"
 *   2026.04 (number)        → "2026.04"
 *   2026.1 (number)         → "2026.10" （10月は浮動小数点で .1 になる）
 *   "2026.10" (string)      → "2026.10"
 *   "2026.04" / "2026-04" / "2026/4" / "2026年4月" (string) → "YYYY.MM"
 */
function normalizeYearMonthCell(value) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date) {
    return value.getFullYear() + '.' + String(value.getMonth() + 1).padStart(2, '0');
  }
  if (typeof value === 'number') {
    const year = Math.floor(value);
    // 小数部分を 100 倍して月を復元。Math.round で浮動小数点誤差を吸収
    const monthFraction = Math.round((value - year) * 100);
    // 月が1桁の場合（例: 2026.1 → 月=10）は既に2桁の値として復元される
    const month = monthFraction;
    if (year >= 2000 && year <= 2100 && month >= 1 && month <= 12) {
      return year + '.' + String(month).padStart(2, '0');
    }
    return String(value);
  }
  const s = String(value).trim();
  // 4桁年 + 区切り + 1〜2桁月
  const m = s.match(/^(\d{4})[.\-\/年](\d{1,2})/);
  if (m) return m[1] + '.' + m[2].padStart(2, '0');
  return s;
}

/**
 * CF管理シートの 1行目 ASIN ヘッダーから ASIN → 列番号 マップを作成
 * ASIN は E列以降（CF_STOCK_ASIN_START_COL = 5）
 */
function getCfAsinColumnMap(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < CF_STOCK_ASIN_START_COL) return {};
  const headers = sheet.getRange(1, CF_STOCK_ASIN_START_COL, 1,
                                 lastCol - CF_STOCK_ASIN_START_COL + 1).getValues()[0];
  const map = {};
  for (let i = 0; i < headers.length; i++) {
    const asin = String(headers[i] || '').trim();
    if (asin) map[asin] = CF_STOCK_ASIN_START_COL + i;
  }
  return map;
}

/**
 * 指定年月の「在庫」行を特定（A列=年月 AND C列=在庫）
 * @returns {number} 行番号（見つからなければ -1）
 */
function findCfMonthlyStockRow(sheet, yearMonth) {
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) return -1;

  // B列（年月）と C列（区分）を一括取得
  // B列の月はブロック先頭行にのみあり、以降の行は空のため前方の値を保持して走査
  const range = sheet.getRange(1, 1, lastRow, 3).getValues();
  let currentYm = '';
  for (let i = 0; i < range.length; i++) {
    const b = normalizeYearMonthCell(range[i][CF_STOCK_MONTH_COL - 1]);
    const c = String(range[i][CF_STOCK_KBN_COL - 1] || '').trim();
    if (b && /^\d{4}\.\d{2}$/.test(b)) currentYm = b;
    if (currentYm === yearMonth && c === CF_STOCK_KBN_VALUE) {
      return i + 1;
    }
  }
  return -1;
}

/**
 * デバッグ: 「在庫」行のマッチング挙動を詳細に確認
 */
function debugStockRowLookup() {
  const sheet = openCfStockSheet();
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(1, 1, lastRow, 3).getValues();
  const target = formatCurrentYearMonth();
  Logger.log('Target yearMonth = ' + JSON.stringify(target));
  let currentYm = '';
  for (let i = 0; i < range.length; i++) {
    const rawB = range[i][1];
    const b = normalizeYearMonthCell(rawB);
    const c = String(range[i][2] || '').trim();
    const matchesRegex = b ? /^\d{4}\.\d{2}$/.test(b) : false;
    if (b && matchesRegex) currentYm = b;
    if (c === '在庫') {
      const match = (currentYm === target);
      Logger.log('row ' + (i+1) +
                 ' | rawB=' + JSON.stringify(rawB) + '(type=' + typeof rawB + ')' +
                 ' | normalized=' + JSON.stringify(b) +
                 ' | regexOk=' + matchesRegex +
                 ' | currentYm=' + JSON.stringify(currentYm) +
                 ' | 一致=' + match);
      if (match) {
        Logger.log('✅ 発見: row ' + (i+1));
        return;
      }
    }
  }
  Logger.log('❌ 一致なし');
}

/**
 * デバッグ: 年月ラベルが入っている列を特定する
 * A〜F列をスキャンして "2026" を含むセルを全部ログ出力
 */
function debugFindMonthColumn() {
  const sheet = openCfStockSheet();
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(1, 1, lastRow, 6).getValues();  // A-F列
  const colNames = ['A', 'B', 'C', 'D', 'E', 'F'];
  let shown = 0;
  for (let i = 0; i < range.length; i++) {
    for (let c = 0; c < 6; c++) {
      const v = range[i][c];
      if (v && String(v).match(/20\d{2}/)) {
        Logger.log('row ' + (i + 1) + ' / col ' + colNames[c] +
                   ' | value=' + JSON.stringify(v) +
                   ' (type=' + (v instanceof Date ? 'Date' : typeof v) + ')');
        shown++;
        if (shown >= 80) { Logger.log('... (表示上限)'); return; }
      }
    }
  }
  Logger.log('===== 終了（' + shown + '件発見） =====');
}

/**
 * デバッグ: CF管理シートのA列（年月）の実値を調べる
 * findCfMonthlyStockRow が「未発見」を返すときの原因調査用
 */
function debugCfMonthColumn() {
  const sheet = openCfStockSheet();
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(1, 1, lastRow, 3).getValues();
  let shown = 0;
  for (let i = 0; i < range.length; i++) {
    const rawA = range[i][0];
    const rawC = range[i][2];
    if (rawA !== '' || rawC === CF_STOCK_KBN_VALUE) {
      const norm = normalizeYearMonthCell(rawA);
      Logger.log('row ' + (i + 1) +
                 ' | A=' + JSON.stringify(rawA) +
                 ' (type=' + (rawA instanceof Date ? 'Date' : typeof rawA) + ')' +
                 ' | normalized=' + norm +
                 ' | C=' + JSON.stringify(rawC));
      shown++;
      if (shown >= 40) break;
    }
  }
  Logger.log('===== 終了 =====');
}

/**
 * 当月の「在庫」行に ASIN別の在庫数を書き込む
 * @returns {number} 書き込んだセル数
 */
function writeCurrentMonthStockToCf(sheet, yearMonth, asinToQty) {
  const row = findCfMonthlyStockRow(sheet, yearMonth);
  if (row < 0) {
    Logger.log('⚠️ CF管理に ' + yearMonth + ' の在庫行が見つかりません');
    return 0;
  }
  const asinCols = getCfAsinColumnMap(sheet);
  if (Object.keys(asinCols).length === 0) {
    Logger.log('⚠️ CF管理に ASIN ヘッダーが見つかりません');
    return 0;
  }

  // 現在値を取得してから差分更新（全体一括 setValues のため）
  const maxCol = Math.max.apply(null, Object.values(asinCols));
  const range = sheet.getRange(row, CF_STOCK_ASIN_START_COL, 1,
                               maxCol - CF_STOCK_ASIN_START_COL + 1);
  const current = range.getValues()[0];
  let updated = 0;
  for (const [asin, qty] of Object.entries(asinToQty)) {
    const col = asinCols[asin];
    if (!col) continue;
    const idx = col - CF_STOCK_ASIN_START_COL;
    if (current[idx] !== qty) {
      current[idx] = qty;
      updated++;
    }
  }
  if (updated > 0) range.setValues([current]);
  return updated;
}

// ===== テスト =====

/**
 * テスト: 実行前に各シートの構造を確認（書き込みなし）
 */
function testInventoryExternalSheets() {
  const orderSheet = openOrderSheet();
  const asinMap = getOrderSheetAsinMap(orderSheet);
  Logger.log('発注管理表 ASIN数: ' + Object.keys(asinMap).length);
  Logger.log('先頭5件:');
  Object.entries(asinMap).slice(0, 5).forEach(([a, r]) => Logger.log('  ' + a + ' → row ' + r));

  const cfSheet = openCfStockSheet();
  const cfAsinCols = getCfAsinColumnMap(cfSheet);
  Logger.log('CF管理 ASIN数: ' + Object.keys(cfAsinCols).length);
  Logger.log('先頭5件:');
  Object.entries(cfAsinCols).slice(0, 5).forEach(([a, c]) => Logger.log('  ' + a + ' → col ' + c));

  const ym = formatCurrentYearMonth();
  const row = findCfMonthlyStockRow(cfSheet, ym);
  Logger.log('CF管理 当月(' + ym + ')の「在庫」行: ' + (row > 0 ? 'row ' + row : '未発見'));
}
