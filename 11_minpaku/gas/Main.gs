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
    const added = reservations.length > 0 ? writeReservations(reservations) : 0;
    if (added > 0) Logger.log(`予約追加: ${added}件`);

    // キャンセルメールも処理
    const cancelled = processCancellationEmails();
    if (cancelled > 0) Logger.log(`キャンセル更新: ${cancelled}件`);

    // 変更があれば集計も更新
    if (added > 0 || cancelled > 0) {
      const fiscalYear = KPICalculator.getCurrentFiscalYear();
      updateMonthlySheet(fiscalYear);
      updateDashboard(fiscalYear);
      Logger.log('集計・ダッシュボード更新完了');
    } else {
      Logger.log('新規予約・キャンセルなし');
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
 * 既存データのステータスを一括修正する（確定→予約、変更→キャンセル）
 */
function fixExistingStatuses() {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.RESERVATIONS);
  if (!sheet || sheet.getLastRow() <= 1) return;

  const C = CONFIG.RESERVATION_COLS;
  const lastRow = sheet.getLastRow();
  const statuses = sheet.getRange(2, C.STATUS, lastRow - 1, 1).getValues();
  let fixed = 0;

  statuses.forEach((row, i) => {
    const s = String(row[0]);
    if (s === '確定') {
      sheet.getRange(i + 2, C.STATUS).setValue('予約');
      fixed++;
    } else if (s === '変更') {
      sheet.getRange(i + 2, C.STATUS).setValue('キャンセル');
      fixed++;
    }
  });

  Logger.log(`ステータス一括修正: ${fixed}件`);
  SpreadsheetApp.getUi().alert(`✅ ステータス修正完了\n・${fixed}件を更新しました`);
}

/**
 * 今すぐGmailを確認して予約を取り込む（手動実行用）
 */
function runManualEmailCheck() {
  checkNewReservations();
  SpreadsheetApp.getUi().alert('✅ メールチェック完了\n予約リストとダッシュボードを更新しました。');
}

/**
 * キャンセルメールを手動処理する
 */
function runCancellationCheck() {
  try {
    const cancelled = processCancellationEmails();
    const fiscalYear = KPICalculator.getCurrentFiscalYear();
    if (cancelled > 0) {
      updateMonthlySheet(fiscalYear);
      updateDashboard(fiscalYear);
    }
    SpreadsheetApp.getUi().alert(`✅ キャンセル処理完了\n・ステータス更新: ${cancelled}件`);
  } catch (e) {
    SpreadsheetApp.getUi().alert(`❌ エラーが発生しました\n${e.message}`);
  }
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
 * 過去のメールを遡って一括取込する（HTMLダイアログでデフォルト日付を表示）
 */
function runBackfill() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; padding: 16px; font-size: 13px; }
        p { margin: 0 0 10px; color: #444; }
        input { width: 100%; padding: 8px; box-sizing: border-box; font-size: 14px; border: 1px solid #ccc; border-radius: 4px; }
        .buttons { display: flex; justify-content: flex-end; gap: 8px; margin-top: 16px; }
        button { padding: 8px 20px; cursor: pointer; border-radius: 4px; font-size: 13px; }
        .ok { background: #1a73e8; color: white; border: none; }
        .cancel { background: white; border: 1px solid #ccc; }
        button:disabled { opacity: 0.5; cursor: not-allowed; }
        #status { display: none; margin-top: 12px; color: #1a73e8; font-size: 12px; }
      </style>
    </head>
    <body>
      <p>取込開始日を入力してください（例: 2025/12/01）<br>空欄のままOKを押すと 2025/12/01 以降を対象にします。</p>
      <input type="text" id="d" value="2025/12/01">
      <div class="buttons">
        <button class="cancel" id="cancelBtn" onclick="google.script.host.close()">キャンセル</button>
        <button class="ok" id="okBtn" onclick="submit()">OK</button>
      </div>
      <div id="status">⏳ メールを取込中です。完了まで数分かかる場合があります…</div>
      <script>
        document.getElementById('d').select();
        function submit() {
          document.getElementById('okBtn').disabled = true;
          document.getElementById('cancelBtn').disabled = true;
          document.getElementById('status').style.display = 'block';
          google.script.run
            .withSuccessHandler(function() { google.script.host.close(); })
            .withFailureHandler(function(e) {
              document.getElementById('status').textContent = '❌ エラー: ' + e.message;
            })
            .continueBackfill(document.getElementById('d').value);
        }
      </script>
    </body>
    </html>
  `).setWidth(420).setHeight(220);

  SpreadsheetApp.getUi().showModalDialog(html, '📧 過去メール一括取込');
}

/**
 * runBackfill() のHTMLダイアログから呼ばれる実処理
 */
function continueBackfill(input) {
  const ui = SpreadsheetApp.getUi();
  const dateStr = (input || '').trim() || '2025/12/01';
  const sinceDate = new Date(dateStr.replace(/\//g, '-'));

  if (isNaN(sinceDate.getTime())) {
    ui.alert('⚠️ 日付の形式が正しくありません。\n例: 2025/12/01');
    return;
  }

  try {
    const reservations = fetchReservationEmailsSince(sinceDate);

    if (reservations.length === 0) {
      ui.alert('ℹ️ 取込対象のメールが見つかりませんでした。\n\n送信元アドレスがConfig.gsの設定と一致しているか確認してください。');
      return;
    }

    const added = writeReservations(reservations);
    const skipped = reservations.length - added;

    // キャンセルメールも処理
    const cancelled = processCancellationEmails(sinceDate);

    const fiscalYear = KPICalculator.getCurrentFiscalYear();
    updateMonthlySheet(fiscalYear);
    updateDashboard(fiscalYear);

    ui.alert(
      `✅ 取込完了\n\n` +
      `・解析したメール: ${reservations.length}件\n` +
      `・新規追加: ${added}件\n` +
      `・重複スキップ: ${skipped}件\n` +
      `・キャンセル反映: ${cancelled}件\n\n` +
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
    .addItem('❌ キャンセルメールを処理', 'runCancellationCheck')
    .addSeparator()
    .addItem('⚙️ 初期セットアップ', 'runSetup')
    .addItem('🔄 列構造マイグレーション（旧17列→新19列）', 'runReservationSheetMigration')
    .addItem('💴 経費入力シートを更新（列追加・60ヶ月入力）', 'migrateCostSheet')
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
