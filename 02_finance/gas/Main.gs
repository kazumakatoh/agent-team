/**
 * キャッシュフロー管理システム - メインエントリーポイント
 * 株式会社LEVEL1 資金繰り管理
 *
 * ■ 定期実行トリガー設定（setupTriggers() を1度だけ実行してください）
 *   - dailyCashFlowCheck() : 毎朝7時実行（全シート一括更新）
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
    // メイン: 一括更新
    .addItem('🚀 最新情報に一括更新（当月＋前月）', 'runFullUpdate')
    .addSeparator()
    // 予定管理
    .addItem('📝 単発予定を登録', 'addPlannedTransaction')
    .addItem('🗓️ 予定マスタから一括展開', 'expandPlannedTransactions')
    .addSeparator()
    // データ更新
    .addItem('📦 在庫数を最新化', 'updateInventoryFromOrderMgmt')
    .addItem('📅 過去データを取込（期間指定）', 'syncWithDateRange')
    .addItem('📊 月別シートを再生成（期間/ALL）', 'updateAllMonthlySheets')
    .addItem('💼 実口座残高を個別更新', 'updateRealBalance')
    .addSeparator()
    // MF連携
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🔗 MF連携')
      .addItem('MF連携開始', 'startMfAuth')
      .addItem('MF連携状態を確認', 'showMfStatus')
      .addItem('🔍 口座マッピング診断', 'showWalletMapping')
      .addItem('MF連携解除', 'disconnectMf')
      .addItem('リダイレクトURIを表示', 'showRedirectUri')
    )
    .addSeparator()
    // 管理
    .addSubMenu(SpreadsheetApp.getUi().createMenu('⚙️ 管理')
      .addItem('初期セットアップ', 'runSetup')
      .addItem('Dailyシートをクリーンアップ', 'cleanupAllDailySheets')
      .addItem('定期実行トリガー設定', 'setupTriggers')
    )
    .addToUi();
}

// ==============================
// 一括更新（メインボタン）
// ==============================

/**
 * 当月＋前月の最新情報に全シートを一括更新する
 *
 * 処理内容（対象期間: 前月1日 〜 今日）:
 *  1. MFから入出金取得 → 各Daily_XXXに反映（会計年度をまたぐ場合は自動分割）
 *  2. 全Dailyシートの残高再計算
 *  3. 日別サマリー更新（全期間再描画）
 *  4. 月別シート更新（前月＋当月）
 *  5. 実口座残高更新（前月＋当月をMF試算表BSから取得）
 *  6. アラートチェック
 *
 * ※ 前月仕訳が当月に追加されるケースに対応するため、前月も再取込する
 */
function runFullUpdate() {
  const ui = SpreadsheetApp.getUi();

  if (!isMfConnected()) {
    ui.alert('❌ MF未連携です。\n\nメニュー → MF連携 → MF連携開始 を実行してください。');
    return;
  }

  const startTime = new Date();

  try {
    const today = new Date();
    const year = today.getFullYear();
    const month = today.getMonth() + 1;

    // 前月の年月を計算（1月実行時は前年12月）
    let prevYear = year;
    let prevMonth = month - 1;
    if (prevMonth === 0) { prevMonth = 12; prevYear--; }

    // 前月1日 〜 今日 の期間
    const dateFrom = `${prevYear}-${String(prevMonth).padStart(2, '0')}-01`;
    const dateTo = Utilities.formatDate(today, CF_CONFIG.DISPLAY.TIMEZONE, 'yyyy-MM-dd');

    // 1〜3. MFデータ同期 → 残高再計算 → 日別サマリー
    //       会計年度をまたぐ場合は分割して取得（splitByFiscalYear_）
    //       syncToDaily_が期間内の既存MF行を削除してから再挿入するためダブらない
    const periods = splitByFiscalYear_(dateFrom, dateTo);
    periods.forEach(p => {
      Logger.log(`MF同期: ${p.from} 〜 ${p.to}`);
      syncToDaily_(p.from, p.to);
    });

    // 4. 月別シート更新（前月〜当月）
    try {
      updateMonthlySheet(prevYear, prevMonth, year, month);
    } catch (e) {
      Logger.log(`月別シート更新エラー: ${e.message}`);
    }

    // 5. 実口座残高更新（前月＋当月）
    try {
      updateRealBalanceMonth_(prevYear, prevMonth);
    } catch (e) {
      Logger.log(`実口座残高更新エラー（前月）: ${e.message}`);
    }
    try {
      updateRealBalanceMonth_(year, month);
    } catch (e) {
      Logger.log(`実口座残高更新エラー（当月）: ${e.message}`);
    }

    // 6. アラートチェック
    const alerts = checkCashOutRisk();
    const alertMsg = (alerts && alerts.length > 0)
      ? `\n\n⚠️ ${alerts.length}件のアラートがあります`
      : '\n\n🟢 キャッシュアウトリスクなし';

    const elapsed = Math.round((new Date() - startTime) / 1000);

    const prevLabel = `${prevYear}.${String(prevMonth).padStart(2,'0')}`;
    const curLabel = `${year}.${String(month).padStart(2,'0')}`;

    ui.alert(
      `✅ 最新情報に更新しました（当月＋前月）\n\n` +
      `・MFデータ同期: ${prevLabel} 〜 ${curLabel}\n` +
      `・Daily各口座: 残高再計算\n` +
      `・日別サマリー: 更新\n` +
      `・月別: ${prevLabel} 〜 ${curLabel}\n` +
      `・実口座残高: ${prevLabel} / ${curLabel}` +
      alertMsg +
      `\n\n処理時間: ${elapsed}秒`
    );

  } catch (e) {
    ui.alert(`❌ エラーが発生しました\n\n${e.message}`);
    Logger.log(`runFullUpdate エラー: ${e.message}\n${e.stack}`);
  }
}

/**
 * 実口座残高の指定月をMFから更新する（内部関数）
 */
function updateRealBalanceMonth_(year, month) {
  const ss = getCfSpreadsheet();
  const sheet = ss.getSheetByName('実口座残高');
  if (!sheet) return;

  const targetRow = findRealBalanceRow_(sheet, year, month);
  if (targetRow === 0) return;

  const lastDay = new Date(year, month, 0).getDate();
  const bsData = mfApiRequest_('/reports/trial_balance_bs', {
    start_date: `${year}-${String(month).padStart(2, '0')}-01`,
    end_date: `${year}-${String(month).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`
  });

  const balanceMap = extractBalancesFromRows_(bsData.rows || []);

  sheet.getRange(targetRow, 3).setValue(balanceMap['普通預金'] || 0);
  sheet.getRange(targetRow, 4).setValue(balanceMap['売掛金'] || 0);
  sheet.getRange(targetRow, 5).setValue(balanceMap['未払金'] || 0);
  sheet.getRange(targetRow, 6).setValue(balanceMap['未払費用'] || 0);
  sheet.getRange(targetRow, 7).setValue(balanceMap['預り金'] || 0);
  sheet.getRange(targetRow, 9).setValue(balanceMap['商品'] || balanceMap['商品及び製品'] || 0);
  sheet.getRange(targetRow, 14).setValue(balanceMap['長期借入金'] || 0);
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
    const fn = t.getHandlerFunction();
    if (fn === 'dailyCashFlowCheck' || fn === 'runFullUpdate') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 毎朝7時: 全シート一括更新
  ScriptApp.newTrigger('runFullUpdate')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .inTimezone('Asia/Tokyo')
    .create();

  Logger.log('トリガー設定完了');
  SpreadsheetApp.getUi().alert(
    '✅ トリガー設定完了\n\n' +
    '・毎朝7時: 全シート一括更新（MF同期+集計+アラート）'
  );
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
