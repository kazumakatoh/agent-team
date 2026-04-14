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
 * M1 商品マスターにヘッダーを設定
 */
function setupProductMasterHeaders() {
  const sheet = getOrCreateSheet(SHEET_NAMES.M1_PRODUCT_MASTER);
  const headers = [
    'ASIN', '商品名', 'カテゴリ', 'ステータス',
    '現在の仕入単価', '仕入れ先', '備考'
  ];

  const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (existing[0] !== headers[0]) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    Logger.log('✅ M1 商品マスター: ヘッダー設定完了');
  }
  return sheet;
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
 * 商品マスターからASIN→商品名のマッピングを取得
 */
function getProductMasterMap() {
  const sheet = getOrCreateSheet(SHEET_NAMES.M1_PRODUCT_MASTER);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return {};

  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const map = {};
  data.forEach(row => {
    if (row[0]) {
      map[row[0]] = {
        name: row[1] || '',
        category: row[2] || '',
        status: row[3] || 'アクティブ',
      };
    }
  });
  return map;
}

/**
 * 商品マスターの情報をD1日次データに反映
 * 商品名・カテゴリが空のセルを商品マスターから埋める
 */
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

  Logger.log('===== 同期完了 =====');
}
