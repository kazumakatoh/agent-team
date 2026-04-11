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
    .addItem('🚀 最新情報に一括更新', 'runFullUpdate')
    .addSeparator()
    // 予定管理
    .addItem('📝 単発予定を登録', 'addPlannedTransaction')
    .addItem('🗓️ 予定マスタから一括展開', 'expandPlannedTransactions')
    .addSeparator()
    // 過去データ取込
    .addItem('📅 過去データを取込（期間指定）', 'syncWithDateRange')
    .addItem('💼 実口座残高を個別更新', 'updateRealBalance')
    .addSeparator()
    // MF連携
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🔗 MF連携')
      .addItem('MF連携開始', 'startMfAuth')
      .addItem('MF連携状態を確認', 'showMfStatus')
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
 * 今日時点の最新情報に全シートを一括更新する
 *
 * 処理内容:
 *  1. MFから当月データ取得 → 各Daily_XXXに反映
 *  2. 全Dailyシートの残高再計算
 *  3. 現残高シート更新
 *  4. 日別サマリー更新
 *  5. 月別シート更新（今年度）
 *  6. 実口座残高更新（当月のみMF API）
 *  7. アラートチェック
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
    const dateFrom = `${year}-${String(month).padStart(2, '0')}-01`;
    const dateTo = Utilities.formatDate(today, CF_CONFIG.DISPLAY.TIMEZONE, 'yyyy-MM-dd');

    // 1〜4. MFデータ同期（当月）→ 残高再計算 → 現残高 → 日別サマリー
    //       syncToDaily_が全てを実行（既存MF行を削除してから再挿入するためダブらない）
    syncToDaily_(dateFrom, dateTo);

    // 5. 月別シート更新（今年度の1月〜現在月）
    try {
      updateMonthlySheet(year, 1, year, month);
    } catch (e) {
      Logger.log(`月別シート更新エラー: ${e.message}`);
    }

    // 6. 実口座残高更新（当月）
    try {
      updateRealBalanceMonth_(year, month);
    } catch (e) {
      Logger.log(`実口座残高更新エラー: ${e.message}`);
    }

    // 7. アラートチェック
    const alerts = checkCashOutRisk();
    const alertMsg = (alerts && alerts.length > 0)
      ? `\n\n⚠️ ${alerts.length}件のアラートがあります`
      : '\n\n🟢 キャッシュアウトリスクなし';

    const elapsed = Math.round((new Date() - startTime) / 1000);

    ui.alert(
      `✅ 最新情報に更新しました\n\n` +
      `・MFデータ同期: ${year}年${month}月\n` +
      `・Daily各口座: 残高再計算\n` +
      `・日別サマリー: 更新\n` +
      `・月別: ${year}年1月〜${month}月\n` +
      `・実口座残高: ${year}.${String(month).padStart(2,'0')}\n` +
      `・現残高: 更新` +
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
// 現残高シート更新
// ==============================

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
