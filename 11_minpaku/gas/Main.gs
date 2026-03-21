/**
 * 民泊自動化システム - メインエントリーポイント
 * 和成也ハイム浅草 予約管理・KPI自動集計
 *
 * ■ 定期実行トリガー設定（setupTriggers() を1度だけ実行してください）
 *   - checkNewReservations()  : 毎時実行（Gmail監視）
 *   - dailyAggregation()      : 毎朝6時実行（集計更新）
 */

// ==============================
// 定期実行メイン処理
// ==============================

/**
 * 【毎時実行】Gmail を監視し、新しい予約メールを取り込む
 */
function checkNewReservations() {
  try {
    Logger.log('=== Gmail予約チェック開始 ===');

    const reservations = fetchNewReservationEmails();
    if (reservations.length === 0) {
      Logger.log('新規予約なし');
      return;
    }

    const added = writeReservations(reservations);
    Logger.log(`予約追加: ${added}件`);

    // 追加があれば集計も更新
    if (added > 0) {
      const fiscalYear = KPICalculator.getCurrentFiscalYear();
      updateMonthlySheet(fiscalYear);
      updateDashboard(fiscalYear);
      Logger.log('集計・ダッシュボード更新完了');
    }

  } catch (e) {
    Logger.log(`エラー: ${e.message}\n${e.stack}`);
    notifyError_('checkNewReservations', e);
  }
}

/**
 * 【毎朝6時実行】月別集計とダッシュボードを更新する
 */
function dailyAggregation() {
  try {
    Logger.log('=== 日次集計更新開始 ===');

    const fiscalYear = KPICalculator.getCurrentFiscalYear();
    updateMonthlySheet(fiscalYear);
    updateDashboard(fiscalYear);

    Logger.log(`日次集計完了 (${fiscalYear}年度)`);
  } catch (e) {
    Logger.log(`エラー: ${e.message}`);
    notifyError_('dailyAggregation', e);
  }
}

// ==============================
// 手動実行関数（メニューから実行）
// ==============================

/**
 * 今すぐGmailを確認して予約を取り込む（手動実行用）
 */
function runManualEmailCheck() {
  checkNewReservations();
  SpreadsheetApp.getUi().alert('✅ メールチェック完了\n予約リストとダッシュボードを更新しました。');
}

/**
 * 集計・ダッシュボードを手動更新する
 */
function runManualAggregation() {
  const ui = SpreadsheetApp.getUi();
  const year = KPICalculator.getCurrentFiscalYear();
  const result = ui.prompt(
    '年度指定',
    `集計する事業年度を入力してください（デフォルト: ${year}）\n例: 2025 → 2025年4月〜2026年3月`,
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) return;

  const inputYear = parseInt(result.getResponseText()) || year;
  updateMonthlySheet(inputYear);
  updateDashboard(inputYear);
  ui.alert(`✅ ${inputYear}年度の集計・ダッシュボードを更新しました。`);
}

/**
 * 過去のメールを遡って一括取込する
 */
function runBackfill() {
  const ui = SpreadsheetApp.getUi();

  const result = ui.prompt(
    '📧 過去メール一括取込',
    '取込開始日を入力してください（例: 2025/12/01）\n' +
    '空欄のままOKを押すと、Gmailの全期間を対象にします。',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) return;

  const input = result.getResponseText().trim();
  let sinceDate = null;

  if (input) {
    sinceDate = new Date(input.replace(/\//g, '-'));
    if (isNaN(sinceDate.getTime())) {
      ui.alert('⚠️ 日付の形式が正しくありません。\n例: 2025/12/01');
      return;
    }
  }

  const label = sinceDate
    ? `${input} 以降`
    : 'Gmail全期間';

  const confirm = ui.alert(
    '確認',
    `${label} の予約確定メールを取込みます。\n数分かかる場合があります。続行しますか？`,
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  try {
    const reservations = fetchReservationEmailsSince(sinceDate);

    if (reservations.length === 0) {
      ui.alert('ℹ️ 取込対象のメールが見つかりませんでした。\n\n送信元アドレスがConfig.gsの設定と一致しているか確認してください。');
      return;
    }

    const added = writeReservations(reservations);
    const skipped = reservations.length - added;

    const fiscalYear = KPICalculator.getCurrentFiscalYear();
    updateMonthlySheet(fiscalYear);
    updateDashboard(fiscalYear);

    ui.alert(
      `✅ 取込完了\n\n` +
      `・解析したメール: ${reservations.length}件\n` +
      `・新規追加: ${added}件\n` +
      `・重複スキップ: ${skipped}件\n\n` +
      `集計・ダッシュボードを更新しました。`
    );

  } catch (e) {
    ui.alert(`❌ エラーが発生しました\n${e.message}`);
    notifyError_('runBackfill', e);
  }
}

/**
 * 予約を手動で1件登録する（メールが来ない場合の手動入力）
 */
function addReservationManually() {
  const ui = SpreadsheetApp.getUi();

  const platform = ui.prompt('プラットフォーム', 'Airbnb / Booking.com / その他', ui.ButtonSet.OK_CANCEL);
  if (platform.getSelectedButton() !== ui.Button.OK) return;

  const checkin = ui.prompt('チェックイン日', '例: 2025/12/25', ui.ButtonSet.OK_CANCEL);
  if (checkin.getSelectedButton() !== ui.Button.OK) return;

  const checkout = ui.prompt('チェックアウト日', '例: 2025/12/27', ui.ButtonSet.OK_CANCEL);
  if (checkout.getSelectedButton() !== ui.Button.OK) return;

  const guests = ui.prompt('人数', '例: 2', ui.ButtonSet.OK_CANCEL);
  if (guests.getSelectedButton() !== ui.Button.OK) return;

  const revenue = ui.prompt('売上（円）', '例: 25000', ui.ButtonSet.OK_CANCEL);
  if (revenue.getSelectedButton() !== ui.Button.OK) return;

  const checkinDate  = new Date(checkin.getResponseText().replace(/\//g, '-'));
  const checkoutDate = new Date(checkout.getResponseText().replace(/\//g, '-'));
  const nights       = Math.round((checkoutDate - checkinDate) / (1000 * 60 * 60 * 24));
  const plt          = platform.getResponseText().trim();
  const rev          = parseInt(revenue.getResponseText()) || 0;
  const commRate     = CONFIG.COMMISSION_RATE[plt.toUpperCase()] || CONFIG.COMMISSION_RATE.OTHER;

  const reservation = {
    emailId:       `MANUAL_${Date.now()}`,
    reservationId: `MAN_${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss')}`,
    platform:      plt,
    bookedDate:    new Date(),
    checkin:       checkinDate,
    checkout:      checkoutDate,
    nights,
    guests:        parseInt(guests.getResponseText()) || 1,
    guestName:     '（手動入力）',
    revenue:       rev,
    commission:    Math.round(rev * commRate),
    cleaningFee:   0,
    status:        '確定',
    notes:         '手動入力'
  };

  writeReservations([reservation]);

  const fiscalYear = KPICalculator.getCurrentFiscalYear();
  updateMonthlySheet(fiscalYear);
  updateDashboard(fiscalYear);

  ui.alert(`✅ 予約を追加しました\n${plt} | ${checkin.getResponseText()} 〜 ${checkout.getResponseText()} (${nights}泊)`);
}

// ==============================
// トリガー設定
// ==============================

/**
 * 定期実行トリガーをセットアップする（初回に1度だけ実行）
 */
function setupTriggers() {
  // 既存トリガーを削除
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // 毎時: Gmail監視
  ScriptApp.newTrigger('checkNewReservations')
    .timeBased()
    .everyHours(1)
    .create();

  // 毎朝6時: 集計更新
  ScriptApp.newTrigger('dailyAggregation')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .inTimezone('Asia/Tokyo')
    .create();

  Logger.log('トリガー設定完了');
  SpreadsheetApp.getUi().alert(
    '✅ トリガー設定完了\n' +
    '・毎時: Gmail予約チェック\n' +
    '・毎朝6時: 集計・ダッシュボード更新'
  );
}

// ==============================
// スプレッドシートメニュー追加
// ==============================

/**
 * スプレッドシートを開いたときにカスタムメニューを追加する
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏠 民泊管理')
    .addItem('📧 今すぐメールをチェック', 'runManualEmailCheck')
    .addItem('📅 過去メールを一括取込', 'runBackfill')
    .addItem('📊 集計・ダッシュボード更新', 'runManualAggregation')
    .addSeparator()
    .addItem('✏️ 予約を手動入力', 'addReservationManually')
    .addSeparator()
    .addItem('⚙️ 初期セットアップ', 'runSetup')
    .addItem('⏰ 定期実行トリガー設定', 'setupTriggers')
    .addToUi();
}

// ==============================
// ユーティリティ
// ==============================

function notifyError_(funcName, error) {
  try {
    // スプレッドシートのログシートにエラーを記録
    const ss    = getSpreadsheet();
    let logSheet = ss.getSheetByName('エラーログ');
    if (!logSheet) {
      logSheet = ss.insertSheet('エラーログ');
      logSheet.appendRow(['日時', '関数名', 'エラーメッセージ', 'スタックトレース']);
    }
    logSheet.appendRow([
      new Date(),
      funcName,
      error.message,
      error.stack || ''
    ]);
  } catch (e) {
    // ログ書き込み自体が失敗した場合は無視
  }
}
