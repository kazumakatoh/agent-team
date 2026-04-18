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
 * 月次更新：最新データ取得 → ダッシュボード＋スコアカード一括更新
 */
function runMonthlyUpdate() {
  const ui = SpreadsheetApp.getUi();
  try {
    testWriteHoldingsToSheet();
    testWriteFXToSheet();
    populateIntegratedDashboard();
    buildProviderScorecard();

    ui.alert('✅ 月次更新完了',
      '以下を更新しました：\n\n' +
      '・現物スナップショット（MEXC）\n' +
      '・FXスナップショット（MT4 HTML）\n' +
      '・統合ダッシュボード\n' +
      '・配信者スコアカード',
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('❌ エラー', `更新中にエラーが発生しました：\n${e.message}`, ui.ButtonSet.OK);
    Logger.log(e.stack);
  }
}
