/**
 * Amazon Dashboard - カスタムメニュー
 *
 * スプレッドシート上部に「🚀 Amazon」メニューを追加し、
 * ワンクリックでダッシュボード更新・レポート生成・診断を実行できるようにする。
 *
 * onOpen は Google Sheets の特殊トリガーで、ユーザーがスプシを開いた瞬間に自動実行される。
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Amazon')
    .addItem('📊 ダッシュボード更新（L1 + L2 + L3）', 'menuRefreshAllDashboards')
    .addItem('📅 日次販売実績シート更新', 'menuRebuildDailySales')
    .addSeparator()
    .addItem('📦 在庫取得 + アラート', 'menuFetchInventory')
    .addItem('🔄 外部スプシ 在庫同期', 'menuSyncExternalSheets')
    .addSeparator()
    .addItem('📈 週次AIレポートを今すぐ送信', 'menuSendWeeklyReport')
    .addItem('📑 月次AI戦略レポートを今すぐ送信', 'menuSendMonthlyReport')
    .addSeparator()
    .addSubMenu(ui.createMenu('🔍 診断・メンテナンス')
      .addItem('容量チェック', 'menuDiagnoseSheetSizes')
      .addItem('トラフィックカバレッジ（過去30日）', 'menuDiagnoseTraffic30')
      .addItem('認証情報チェック', 'menuCheckCredentials')
      .addItem('在庫アラート状態リセット', 'menuResetStockAlertState'))
    .addToUi();
}

// ===== ラッパー関数（進捗トーストを表示） =====

function menuRefreshAllDashboards() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('ダッシュボードを更新中...', '🚀 Amazon', 60);
  try {
    updateDashboardL1();
    updateDashboardL2();
    updateDashboardL3();
    ss.toast('✅ L1 + L2 + L3 更新完了', '🚀 Amazon', 5);
  } catch (e) {
    ui.alert('エラー', e.message, ui.ButtonSet.OK);
  }
}

function menuRebuildDailySales() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('日次販売実績を再構築中...', '🚀 Amazon', 60);
  try {
    buildDailySalesSheet();
    ss.toast('✅ 日次販売実績 更新完了', '🚀 Amazon', 5);
  } catch (e) {
    SpreadsheetApp.getUi().alert('エラー', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function menuFetchInventory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('在庫を取得してアラート判定中...', '🚀 Amazon', 60);
  try {
    fetchInventoryAndAlert();
    ss.toast('✅ 在庫取得 + アラート判定 完了', '🚀 Amazon', 5);
  } catch (e) {
    SpreadsheetApp.getUi().alert('エラー', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function menuSyncExternalSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('発注管理表 / CF管理に在庫同期中...', '🚀 Amazon', 120);
  try {
    syncInventoryToExternalSheets();
    ss.toast('✅ 外部スプシ同期完了', '🚀 Amazon', 5);
  } catch (e) {
    SpreadsheetApp.getUi().alert('エラー', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function menuSendWeeklyReport() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    '週次AIレポート',
    '今すぐ週次レポートを生成して Gmail に送信します。よろしいですか？',
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('週次AIレポート生成中（30秒〜1分）...', '🚀 Amazon', 90);
  try {
    sendWeeklyAiReport();
    ss.toast('✅ Gmail送信完了', '🚀 Amazon', 5);
  } catch (e) {
    ui.alert('エラー', e.message, ui.ButtonSet.OK);
  }
}

function menuSendMonthlyReport() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    '月次AI戦略レポート',
    '今すぐ月次戦略レポート（opus-4-6）を生成して Gmail に送信します。' +
    '\n生成に1〜2分かかります。よろしいですか？',
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('月次AI戦略レポート生成中（1〜2分）...', '🚀 Amazon', 180);
  try {
    sendMonthlyAiReport();
    ss.toast('✅ Gmail送信完了', '🚀 Amazon', 5);
  } catch (e) {
    ui.alert('エラー', e.message, ui.ButtonSet.OK);
  }
}

function menuDiagnoseSheetSizes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('容量を診断中...', '🚀 Amazon', 30);
  const result = diagnoseSheetSizes();
  const total = result.total;
  const pct = (total / 10000000 * 100).toFixed(1);
  const top3 = result.results.slice(0, 3).map(r =>
    '・' + r.name + ': ' + r.cells.toLocaleString() + ' セル').join('\n');
  SpreadsheetApp.getUi().alert('容量診断結果',
    '合計: ' + total.toLocaleString() + ' / 10,000,000 (' + pct + '%)\n\n上位3シート:\n' + top3 +
    '\n\n※詳細は実行ログを確認してください',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuDiagnoseTraffic30() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('過去30日のトラフィックカバレッジを確認中...', '🚀 Amazon', 30);
  diagnoseTrafficCoverage30();
  ss.toast('✅ 完了（実行ログ参照）', '🚀 Amazon', 5);
}

function menuCheckCredentials() {
  checkCredentials();
  SpreadsheetApp.getActiveSpreadsheet().toast('✅ 認証情報チェック完了（実行ログ参照）', '🚀 Amazon', 5);
}

function menuResetStockAlertState() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    '在庫アラート状態リセット',
    '在庫アラートの送信履歴をクリアします。\n' +
    '次回 fetchInventoryAndAlert 実行時に、該当商品すべてにアラートが再送されます。',
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;
  resetStockAlertState();
  ui.alert('✅ リセット完了', '在庫アラート状態をクリアしました。', ui.ButtonSet.OK);
}
