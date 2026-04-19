/**
 * スプシカスタムメニュー
 * スプシを開いた時に自動的にメニュー追加
 * 「🚀 SageMaster運用」として全主要関数へのショートカット提供
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 SageMaster運用')
    .addItem('📊 月次更新（ダッシュボード＋スコアカード）', 'runMonthlyUpdate')
    .addSeparator()
    .addItem('📈 現物データのみ更新', 'testWriteHoldingsToSheet')
    .addItem('💱 FXデータのみ更新', 'testWriteFXToSheet')
    .addItem('🎯 統合ダッシュボードのみ更新', 'populateIntegratedDashboard')
    .addItem('👤 スコアカードのみ更新', 'buildProviderScorecard')
    .addSeparator()
    .addItem('📝 週次レビュー即時生成（テスト）', 'testWeeklyReview')
    .addItem('🔍 MEXC接続テスト', 'testMEXCConnection')
    .addItem('🔍 MT4パーステスト', 'testMT4Parse')
    .addToUi();
}

/**
 * 月次更新：最新データ取得 → ダッシュボード＋スコアカード一括更新 → カスタマイズ再適用
 *
 * sheet.clear()でリセットされるトグル・説明文を必ず再適用する
 */
function runMonthlyUpdate() {
  const ui = SpreadsheetApp.getUi();
  try {
    // 1. データ取得（sheet.clear() でトグル等リセット）
    testWriteHoldingsToSheet();
    testWriteFXToSheet();

    // 2. ダッシュボード・スコアカード更新
    populateIntegratedDashboard();
    buildProviderScorecard();

    // 3. カスタマイズ再適用（リセット分の復元）
    addSpotSnapshotToggle();
    addFXSnapshotToggleAndExplanations();
    applyScorecardFixesV2();
    updateScorecardRRTarget();
    updateFXSnapshotExplanations();

    // 4. トリガー再確認
    setupSnapshotTriggers();

    ui.alert('✅ 月次更新完了',
      '以下を更新しました：\n\n' +
      '・現物スナップショット＋USD/JPY切替\n' +
      '・FXスナップショット＋説明文＋切替\n' +
      '・統合ダッシュボード\n' +
      '・配信者スコアカード＋RR目標\n' +
      '・編集トリガー再登録',
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('❌ エラー', `更新中にエラーが発生しました：\n${e.message}`, ui.ButtonSet.OK);
    Logger.log(e.stack);
  }
}
