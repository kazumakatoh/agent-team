/**
 * キャッシュフロー管理システム - メインエントリーポイント
 * 株式会社LEVEL1 資金繰り管理
 *
 * ■ 定期実行トリガー設定（setupTriggers() を1度だけ実行してください）
 *   - dailyCashFlowCheck() : 毎朝7時実行（MF同期＋アラートチェック）
 */

// ==============================
// スプレッドシートメニュー
// ==============================

/**
 * スプレッドシートを開いたときにカスタムメニューを追加する
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💰 CF管理')
    // MFデータ同期
    .addItem('🔄 MFデータ同期（当月）', 'syncCurrentMonth')
    .addItem('📅 MFデータ同期（期間指定）', 'syncWithDateRange')
    .addSeparator()
    // 入出金予定
    .addItem('📝 入出金予定を登録', 'addPlannedTransaction')
    .addSeparator()
    // 集計
    .addItem('📊 月次集計を更新', 'updateAllMonthlySheets')
    .addItem('💹 残高サマリー', 'showBalanceSummary')
    .addItem('⚠️ アラートチェック', 'checkCashOutRisk')
    .addSeparator()
    // MF連携
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🔗 MF連携')
      .addItem('MF連携開始', 'startMfAuth')
      .addItem('MF連携状態を確認', 'showMfStatus')
      .addItem('MF連携解除', 'disconnectMf')
      .addItem('リダイレクトURIを表示', 'showRedirectUri')
    )
    .addSeparator()
    // セットアップ
    .addItem('⚙️ 初期セットアップ', 'runSetup')
    .addItem('⏰ 定期実行トリガー設定', 'setupTriggers')
    .addToUi();
}

// ==============================
// トリガー設定
// ==============================

/**
 * 定期実行トリガーをセットアップする（初回に1度だけ実行）
 */
function setupTriggers() {
  // 既存の本システム関連トリガーを削除
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'dailyCashFlowCheck') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 毎朝7時: MFデータ同期＋アラートチェック
  ScriptApp.newTrigger('dailyCashFlowCheck')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .inTimezone('Asia/Tokyo')
    .create();

  Logger.log('トリガー設定完了');
  SpreadsheetApp.getUi().alert(
    '✅ トリガー設定完了\n\n' +
    '・毎朝7時: MFデータ同期＋アラートチェック\n\n' +
    'MFから入出金データを自動取得し、\nキャッシュアウトリスクを検知します。'
  );
}

// ==============================
// 手動実行ショートカット
// ==============================

/**
 * MF同期 → 月次集計 → アラートチェック を一括実行
 */
function runFullUpdate() {
  const ui = SpreadsheetApp.getUi();

  if (!isMfConnected()) {
    ui.alert('❌ MF未連携です。\n\nメニュー → MF連携 → MF連携開始 を実行してください。');
    return;
  }

  try {
    // 1. 当月のMFデータ同期
    const today = new Date();
    const year = today.getFullYear();
    const month = today.getMonth() + 1;
    const dateFrom = `${year}-${String(month).padStart(2, '0')}-01`;
    const dateTo = Utilities.formatDate(today, CF_CONFIG.DISPLAY.TIMEZONE, 'yyyy-MM-dd');

    const allTxns = fetchAllWalletTransactions(dateFrom, dateTo);
    const ss = getCfSpreadsheet();
    const sheet = ss.getSheetByName(CF_CONFIG.SHEETS.DAILY);

    if (sheet) {
      for (const [accountKey, txns] of Object.entries(allTxns)) {
        if (txns.length > 0) writeTransactionsToDaily_(sheet, accountKey, txns);
      }
      Object.keys(CF_CONFIG.ACCOUNTS).forEach(key => recalculateBalances(key));
      updateDailyTotals_();
    }

    // 2. 月次集計更新
    Object.keys(CF_CONFIG.ACCOUNTS).forEach(key => {
      updateAccountMonthlySheet(key, year);
    });
    updateConsolidatedMonthlySheet(year);

    // 3. 現残高シート更新
    updateCurrentBalanceSheet_();

    // 4. アラートチェック
    const alerts = checkCashOutRisk();

    const alertMsg = (alerts && alerts.length > 0)
      ? `\n\n⚠️ ${alerts.length}件のアラートがあります。`
      : '\n\n🟢 キャッシュアウトリスクなし。';

    ui.alert(
      `✅ 一括更新完了\n\n` +
      `・MFデータ同期: ${year}年${month}月\n` +
      `・月次集計: 更新済み\n` +
      `・現残高: 更新済み` +
      alertMsg
    );

  } catch (e) {
    ui.alert(`❌ エラーが発生しました\n\n${e.message}`);
    Logger.log(`runFullUpdate エラー: ${e.message}\n${e.stack}`);
  }
}

/**
 * 現残高シートを更新する
 */
function updateCurrentBalanceSheet_() {
  const ss = getCfSpreadsheet();
  const sheet = ss.getSheetByName(CF_CONFIG.SHEETS.CURRENT_BAL);
  if (!sheet) return;

  const now = Utilities.formatDate(new Date(), CF_CONFIG.DISPLAY.TIMEZONE, 'yyyy/MM/dd HH:mm');
  let row = 2;

  Object.entries(CF_CONFIG.ACCOUNTS).forEach(([key, account]) => {
    // 各口座のDailyシートから最新残高を取得
    const balance = getLatestBalance_(key);

    sheet.getRange(row, 2).setValue(balance).setNumberFormat('#,##0');
    sheet.getRange(row, 3).setValue(now);

    // ステータス（CF005のみアラート判定）
    if (key === CF_CONFIG.ALERT.ALERT_ACCOUNT) {
      if (balance <= CF_CONFIG.ALERT.DANGER_THRESHOLD) {
        sheet.getRange(row, 4).setValue('🔴 危険');
        sheet.getRange(row, 2).setFontColor('#b71c1c');
      } else if (balance <= CF_CONFIG.ALERT.WARNING_THRESHOLD) {
        sheet.getRange(row, 4).setValue('🟡 注意');
        sheet.getRange(row, 2).setFontColor('#f57f17');
      } else {
        sheet.getRange(row, 4).setValue('🟢 正常');
        sheet.getRange(row, 2).setFontColor('#2e7d32');
      }
    } else {
      sheet.getRange(row, 4).setValue('--');
    }
    row++;
  });
}

// ==============================
// ユーティリティ
// ==============================

/**
 * エラーをログシートに記録する
 */
function logError_(funcName, error) {
  try {
    const ss = getCfSpreadsheet();
    let logSheet = ss.getSheetByName('エラーログ');
    if (!logSheet) {
      logSheet = ss.insertSheet('エラーログ');
      logSheet.appendRow(['日時', '関数名', 'エラーメッセージ', 'スタックトレース']);
    }
    logSheet.appendRow([new Date(), funcName, error.message, error.stack || '']);
  } catch (e) {
    // 無視
  }
}
