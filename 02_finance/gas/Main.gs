/**
 * 財務レポート自動化システム - メインエントリーポイント v1.3
 * MF会計 推移試算表CSVインポート → PLレポート自動生成
 */

// ==============================
// メニュー設定
// ==============================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 財務レポート')
    .addItem('🔄 月次更新：PL（CSVインポート＋通期比較シート更新）', 'runCurrentYearUpdate')
    .addItem('🔄 月次更新：BS（貸借対照表 通期比較シート更新）', 'runBSPeriodComparison')
    .addSeparator()
    .addItem('📥 PL: CSVインポートのみ', 'runCSVImport')
    .addItem('📊 PL: 通期比較シートのみ作成（全期分）', 'runPeriodComparison')
    .addToUi();
}

// ==============================
// 実行関数
// ==============================

/**
 * 部門別CSVインポートを実行する
 * Driveフォルダ内の {部門名}_PL_{年}.csv を読み込んでPLシートに反映
 */
function runCSVImport() {
  CSVImporter.importAllFromDrive();
}

/**
 * 現在の事業年度をインポートして通期比較まで一括更新する
 * ダッシュボードのチェックボックスから呼ばれる（年度入力なし）
 */
function runCurrentYearUpdate() {
  Dashboard.updateStatus('⏳ インポート中...');

  // 1. PLインポート（現在の事業年度・UI入力なし）
  const result = CSVImporter.importCurrentYear();
  Logger.log(result.msg);

  if (!result.ok && result.imported === 0) {
    Dashboard.updateStatus('❌ インポート失敗: ' + result.msg);
    return;
  }

  // 2. 通期比較シート更新
  Dashboard.updateStatus('⏳ 通期比較シート更新中...');
  const BASE_YEAR   = 2018;
  const currYear    = getCurrentFiscalYear();
  const fiscalYears = [];
  for (let y = BASE_YEAR; y <= currYear; y++) fiscalYears.push(y);

  try {
    SheetManager.writePeriodComparisonSheet(fiscalYears);
    Dashboard.updateStatus(result.msg + ' ／ 通期比較更新完了 ✅');
    Logger.log('通期比較更新完了');
  } catch (e) {
    Dashboard.updateStatus(result.msg + ' ⚠️ 通期比較失敗: ' + e.message);
    Logger.log('通期比較更新失敗: ' + e.message);
  }
}

/**
 * 通期比較シートを作成する（第1期〜現在）
 */
function runPeriodComparison() {
  const BASE_YEAR = 2018;
  const currYear  = getCurrentFiscalYear();
  const fiscalYears = [];
  for (let y = BASE_YEAR; y <= currYear; y++) fiscalYears.push(y);

  try {
    SheetManager.writePeriodComparisonSheet(fiscalYears);
    const msg = `✅ 通期比較シートを作成しました。\n（${getFiscalPeriodLabel(BASE_YEAR)}〜${getFiscalPeriodLabel(currYear)}　${fiscalYears.length}期分）\nシート名: 通期比較_全体`;
    Logger.log(msg);
    try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  } catch (e) {
    Logger.log(`❌ エラー: ${e.message}`);
    try { SpreadsheetApp.getUi().alert(`❌ エラー: ${e.message}`); } catch (e2) {}
  }
}

// ==============================
// エラーログ
// ==============================

/**
 * BS通期比較シートを作成する（全体_BS_{年}.csv から）
 */
function runBSPeriodComparison() {
  try {
    Logger.log('BS通期比較シート作成開始...');
    const BASE_YEAR   = 2018;
    const currYear    = getCurrentFiscalYear();
    const fiscalYears = [];
    for (let y = BASE_YEAR; y <= currYear; y++) fiscalYears.push(y);

    const yearDataMap = BSImporter.importAllFromDrive();
    const foundYears  = Object.keys(yearDataMap).length;

    if (foundYears === 0) {
      const msg = '⚠️ BSデータが見つかりません。\nDriveフォルダに 全体_BS_{年}.csv を配置してください。\n例: 全体_BS_2026.csv';
      Logger.log(msg);
      try { SpreadsheetApp.getUi().alert(msg); } catch(e) {}
      return;
    }

    SheetManager.writePeriodComparisonBSSheet(fiscalYears, yearDataMap);
    const msg = `✅ 通期比較_BSシートを作成しました。\n（${foundYears}期分のデータを読み込み）\nシート名: 通期比較_BS`;
    Logger.log(msg);
    try { SpreadsheetApp.getUi().alert(msg); } catch(e) {}
  } catch (e) {
    Logger.log(`❌ エラー: ${e.message}\n${e.stack}`);
    try { SpreadsheetApp.getUi().alert(`❌ エラー: ${e.message}`); } catch(e2) {}
  }
}

function _notifyError(funcName, error) {
  try {
    const ss = SheetManager.getSpreadsheet();
    let logSheet = ss.getSheetByName('_エラーログ');
    if (!logSheet) {
      logSheet = ss.insertSheet('_エラーログ');
      logSheet.appendRow(['日時', '関数名', 'エラー', 'スタック']);
    }
    logSheet.appendRow([new Date(), funcName, error.message, error.stack || '']);
  } catch (e) {}
}
