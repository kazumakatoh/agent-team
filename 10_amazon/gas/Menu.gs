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
    .addItem('⚡ 全部最新化（売上+トラフィック+広告+在庫+DB / 約5〜7分）', 'menuRefreshAll')
    .addSeparator()
    .addItem('📊 ダッシュボード更新（L1 + L2 + L3）', 'menuRefreshAllDashboards')
    .addItem('📅 日次販売実績シート更新', 'menuRebuildDailySales')
    .addSeparator()
    .addItem('📦 在庫取得 + アラート', 'menuFetchInventory')
    .addItem('🔄 外部スプシ 在庫同期', 'menuSyncExternalSheets')
    .addItem('🎯 Amazon Ads レポート取得（昨日）', 'menuFetchAdsReports')
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

/**
 * インストール型 onOpen トリガーを登録
 *
 * このGASはスタンドアロン（スプシにバインドされていない）ため、
 * 単純トリガーの onOpen は動作しない。インストール型トリガーを
 * 1回登録するとメニューが毎回表示される。
 *
 * GASエディタから1回だけ手動実行する。
 */
function setupOnOpenTrigger() {
  // 既存の onOpen トリガーを全削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'onOpen' &&
        t.getEventType() === ScriptApp.EventType.ON_OPEN) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // メインスプシ宛に新規作成
  const ss = SpreadsheetApp.openById(getMainSpreadsheetId());
  ScriptApp.newTrigger('onOpen')
    .forSpreadsheet(ss)
    .onOpen()
    .create();

  Logger.log('✅ onOpen トリガー登録完了');
  Logger.log('スプレッドシートを再読み込みすると「🚀 Amazon」メニューが表示されます');
}

// ===== ラッパー関数（進捗トーストを表示） =====

/**
 * 🔥 全部最新化
 *
 * 俯瞰確認用の統合ボタン。以下を順に実行する：
 *
 *   [1] 前日分の売上データ取得（Orders Report）← D1 新規行
 *   [2] 前日分のトラフィック取得（Sales & Traffic Report）← D1 セッション/PV/CVR等
 *   [3] 商品マスター → D1 へ商品名・カテゴリ同期
 *   [4] 前日分の広告レポート取得（spAdvertisedProduct/SearchTerm/Targeting）
 *       → D3 の3シート書き込み + D1 の広告4指標更新（最長 2〜3分）
 *   [5] FBA在庫取得 + 在庫シート更新 + LINE 在庫切れアラート判定
 *   [6] 発注管理表「発注タイミング」F列 + CF管理「在庫残高」へ在庫同期
 *   [7] 日次販売実績シート（D1S）再構築
 *   [8] L1 / L2 / L3 ダッシュボード更新
 *
 * 各ステップは try/catch で独立させ、1つ失敗しても残りは続行する。
 * 合計所要時間は概ね5〜7分（Ads API の混雑状況で変動）。
 */
function menuRefreshAll() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const t0 = Date.now();
  const errors = [];

  ss.toast('全部最新化を開始します（5〜7分）...', '🚀 Amazon', 480);

  // [1/8] 前日売上（Orders Report）
  try {
    ss.toast('[1/8] 売上データ取得中...', '🚀 Amazon', 120);
    dailyFetchByReport();
  } catch (e) {
    errors.push('売上: ' + e.message);
    Logger.log('❌ 売上取得失敗: ' + e.message);
  }

  // [2/8] 前日トラフィック（Sales & Traffic Report）
  try {
    ss.toast('[2/8] トラフィック取得中...', '🚀 Amazon', 120);
    dailyFetchTraffic();
  } catch (e) {
    errors.push('トラフィック: ' + e.message);
    Logger.log('❌ トラフィック取得失敗: ' + e.message);
  }

  // [3/8] 商品マスター → D1 同期
  try {
    ss.toast('[3/8] 商品マスター同期中...', '🚀 Amazon', 60);
    syncMasterToDaily();
  } catch (e) {
    errors.push('マスター同期: ' + e.message);
    Logger.log('❌ マスター同期失敗: ' + e.message);
  }

  // [4/8] 広告レポート（最長・昨日分）
  try {
    ss.toast('[4/8] 広告レポート取得中（2〜3分）...', '🚀 Amazon', 300);
    dailyFetchAdsReports();
  } catch (e) {
    errors.push('広告: ' + e.message);
    Logger.log('❌ 広告レポート失敗: ' + e.message);
  }

  // [5/8] 在庫取得 + アラート
  try {
    ss.toast('[5/8] 在庫取得中...', '🚀 Amazon', 120);
    fetchInventoryAndAlert();
  } catch (e) {
    errors.push('在庫: ' + e.message);
    Logger.log('❌ 在庫取得失敗: ' + e.message);
  }

  // [6/8] 外部スプシ同期（発注管理表 + CF管理）
  try {
    ss.toast('[6/8] 発注管理表/CF管理へ同期中...', '🚀 Amazon', 120);
    syncInventoryToExternalSheets();
  } catch (e) {
    errors.push('外部スプシ: ' + e.message);
    Logger.log('❌ 外部スプシ同期失敗: ' + e.message);
  }

  // [7/8] 日次販売実績 再構築
  try {
    ss.toast('[7/8] 日次販売実績 再構築中...', '🚀 Amazon', 120);
    buildDailySalesSheet();
  } catch (e) {
    errors.push('日次販売実績: ' + e.message);
    Logger.log('❌ 日次販売実績失敗: ' + e.message);
  }

  // [8/8] ダッシュボード更新（L1 + L2 + L3）
  try {
    ss.toast('[8/8] ダッシュボード更新中（L1 + L2 + L3）...', '🚀 Amazon', 120);
    updateDashboardL1();
    updateDashboardL2();
    updateDashboardL3();
  } catch (e) {
    errors.push('ダッシュボード: ' + e.message);
    Logger.log('❌ ダッシュボード失敗: ' + e.message);
  }

  const elapsed = Math.round((Date.now() - t0) / 1000);
  if (errors.length === 0) {
    ss.toast('✅ 全部最新化 完了（' + elapsed + '秒）', '🚀 Amazon', 10);
  } else {
    ss.toast('⚠️ ' + errors.length + '件失敗（' + elapsed + '秒） - 詳細は実行ログ', '🚀 Amazon', 15);
    ui.alert(
      '⚠️ 一部の更新が失敗しました',
      '失敗ステップ:\n' + errors.map(e => '・' + e).join('\n') +
      '\n\n成功したステップは反映済みです。失敗したものだけ個別メニューから再実行してください。',
      ui.ButtonSet.OK);
  }
}

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

function menuFetchAdsReports() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Amazon Ads レポート取得中（2〜3分）...', '🚀 Amazon', 240);
  try {
    dailyFetchAdsReports();
    ss.toast('✅ Ads レポート取得完了', '🚀 Amazon', 5);
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
