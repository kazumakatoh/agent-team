/**
 * 民泊自動化システム - 初期セットアップモジュール
 * 必要なシートをすべて作成し、書式・ヘッダーを設定する
 */

/**
 * 初期セットアップを実行する
 * スプレッドシートに必要なシートをすべて作成する
 */
function runSetup() {
  const ss = getSpreadsheet();

  Logger.log('=== 初期セットアップ開始 ===');

  setupReservationSheet_(ss);
  setupCostSheet_(ss);
  setupMonthlySheet_(ss);
  setupAnnualSheet_(ss);
  setupDashboardSheet_(ss);
  setupErrorLogSheet_(ss);

  // 不要なデフォルトシートを削除
  removeDefaultSheets_(ss);

  // ダッシュボードを最初のシートに移動
  const dashSheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (dashSheet) ss.setActiveSheet(dashSheet);

  Logger.log('=== 初期セットアップ完了 ===');

  try {
    SpreadsheetApp.getUi().alert(
      '✅ セットアップ完了！\n\n' +
      '作成されたシート:\n' +
      '  ・ダッシュボード\n' +
      '  ・予約リスト\n' +
      '  ・経費入力\n' +
      '  ・月別集計\n' +
      '  ・年間集計\n\n' +
      '次のステップ:\n' +
      '1. 「⚙️ 定期実行トリガー設定」を実行\n' +
      '2. 「経費入力」シートに毎月の固定費を入力'
    );
  } catch (e) {
    // UI なし環境（テスト実行等）では無視
  }
}

// ==============================
// 各シートのセットアップ
// ==============================

function setupReservationSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.RESERVATIONS);
  if (sheet) {
    Logger.log(`シート「${CONFIG.SHEETS.RESERVATIONS}」は既に存在します`);
    return;
  }

  sheet = ss.insertSheet(CONFIG.SHEETS.RESERVATIONS);
  applyReservationHeaders_(sheet);
  applyReservationColumnWidths_(sheet);

  Logger.log(`シート「${CONFIG.SHEETS.RESERVATIONS}」を作成しました`);
}

function setupCostSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.COSTS);
  if (sheet) {
    Logger.log(`シート「${CONFIG.SHEETS.COSTS}」は既に存在します`);
    return;
  }

  sheet = ss.insertSheet(CONFIG.SHEETS.COSTS);

  const headers = ['年月', '清掃費', '備品・消耗品費', '水光熱費', '家賃', 'その他経費', '備考'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleSheet_(sheet, headers.length, {
    headerBg:    '#ff6d00',
    headerColor: '#ffffff',
    freeze:       1
  });

  // サンプルデータ（説明用）
  const note = [
    ['例: 2025-12', 5000, 2000, 8000, 80000, 0, '清掃会社（ゲスト1人×2回）、水道・電気代']
  ];
  sheet.getRange(2, 1, note.length, headers.length).setValues(note);
  sheet.getRange(2, 1, note.length, headers.length).setFontColor('#9aa0a6'); // グレー（説明文）

  // 金額列のフォーマット
  sheet.getRange('B:F').setNumberFormat('¥#,##0');

  // 列幅
  sheet.setColumnWidth(1, 100);  // 年月
  sheet.setColumnWidth(7, 200);  // 備考

  // 入力ガイドのメモ
  sheet.getRange(1, 1).setNote(
    '年月はYYYY-MM形式で入力してください\n例: 2025-12'
  );
  sheet.getRange(1, 2).setNote(
    '自動集計される清掃費（予約リストより）に加算されます。\n清掃会社への支払いなど、予約外の清掃費を入力してください。'
  );

  Logger.log(`シート「${CONFIG.SHEETS.COSTS}」を作成しました`);
}

function setupMonthlySheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.MONTHLY);
  if (sheet) {
    Logger.log(`シート「${CONFIG.SHEETS.MONTHLY}」は既に存在します`);
    return;
  }

  sheet = ss.insertSheet(CONFIG.SHEETS.MONTHLY);

  const headers = [
    '年月', '問い合わせ数', '稼働日数', '利用可能日数',
    '利用件数', '利用人数', '売上', '手数料', '清掃費',
    '備品・消耗品費', '水光熱費', '家賃', 'その他経費', '総経費', '利益',
    'ROI(%)', 'ADR(円)', 'RevPAR(円)', '稼働率(%)'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleSheet_(sheet, headers.length, {
    headerBg:    '#1a73e8',
    headerColor: '#ffffff',
    freeze:       1
  });

  Logger.log(`シート「${CONFIG.SHEETS.MONTHLY}」を作成しました`);
}

function setupAnnualSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.ANNUAL);
  if (sheet) {
    Logger.log(`シート「${CONFIG.SHEETS.ANNUAL}」は既に存在します`);
    return;
  }

  sheet = ss.insertSheet(CONFIG.SHEETS.ANNUAL);

  const headers = [
    '事業年度', '稼働日数', '年間上限日数', '利用可能日数',
    '利用件数', '利用人数', '売上', '手数料', '清掃費',
    '備品・消耗品費', '水光熱費', '家賃', 'その他経費', '総経費', '利益',
    'ROI(%)', 'ADR(円)', 'RevPAR(円)', '稼働率(%)', '法定稼働率(%)'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleSheet_(sheet, headers.length, {
    headerBg:    '#7b1fa2',
    headerColor: '#ffffff',
    freeze:       1
  });

  Logger.log(`シート「${CONFIG.SHEETS.ANNUAL}」を作成しました`);
}

function setupDashboardSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (sheet) {
    Logger.log(`シート「${CONFIG.SHEETS.DASHBOARD}」は既に存在します`);
    return;
  }

  sheet = ss.insertSheet(CONFIG.SHEETS.DASHBOARD, 0); // 先頭に追加

  sheet.getRange('A1:L2').merge()
       .setValue('ダッシュボードは「集計・ダッシュボード更新」を実行すると表示されます。\nメニュー: 🏠 民泊管理 → 集計・ダッシュボード更新')
       .setFontSize(14)
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle')
       .setBackground('#e3f2fd')
       .setFontColor('#1565c0');
  sheet.setRowHeight(1, 80);

  Logger.log(`シート「${CONFIG.SHEETS.DASHBOARD}」を作成しました`);
}

function setupErrorLogSheet_(ss) {
  if (ss.getSheetByName('エラーログ')) return;

  const sheet = ss.insertSheet('エラーログ');
  sheet.appendRow(['日時', '関数名', 'エラーメッセージ', 'スタックトレース']);
  styleSheet_(sheet, 4, {
    headerBg:    '#e53935',
    headerColor: '#ffffff',
    freeze:       1
  });
}

// ==============================
// ユーティリティ
// ==============================

function styleSheet_(sheet, numCols, options) {
  const { headerBg, headerColor, freeze } = options;
  const headerRange = sheet.getRange(1, 1, 1, numCols);

  headerRange.setBackground(headerBg)
             .setFontColor(headerColor)
             .setFontWeight('bold')
             .setHorizontalAlignment('center')
             .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);
  if (freeze) sheet.setFrozenRows(freeze);
}

/**
 * 予約リストを旧15列から新17列構造にマイグレーションする
 * 既存データを保持しつつヘッダーと列構造を更新する
 * ※ 実行前に必ずスプレッドシートをバックアップしてください
 */
function runReservationSheetMigration() {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.RESERVATIONS);
  if (!sheet) {
    Logger.log('予約リストシートが存在しません。Setupを先に実行してください。');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log('データなし。ヘッダーのみ更新します。');
    applyReservationHeaders_(sheet);
    return;
  }

  // 既存データを全件読み込み（旧15列想定）
  const oldData = sheet.getRange(2, 1, lastRow - 1, 15).getValues();

  // 旧列インデックス（0始まり）
  // 旧: ID=0, Platform=1, BookedDate=2, Checkin=3, Checkout=4,
  //     Nights=5, Guests=6, GuestName=7, Revenue=8, Commission=9,
  //     CleaningFee=10, NetRevenue=11, Status=12, Notes=13, EmailId=14
  const newRows = oldData.map(row => {
    const newRow = new Array(17).fill('');
    newRow[0]  = row[0];  // 予約ID
    newRow[1]  = row[1];  // プラットフォーム
    newRow[2]  = row[2];  // 予約受付日
    newRow[3]  = row[3];  // チェックイン
    newRow[4]  = row[4];  // チェックアウト
    newRow[5]  = row[5];  // 宿泊数
    newRow[6]  = row[6];  // 人数
    newRow[7]  = row[7];  // ゲスト名
    newRow[8]  = row[8];  // 売上（旧: Revenue → そのまま）
    newRow[9]  = 0;       // 宿泊料（旧データなし → 0）
    newRow[10] = row[10]; // 清掃費（旧: CleaningFee → そのまま）
    newRow[11] = row[9];  // OTA手数料（旧: Commission → L列へ）
    newRow[12] = 0;       // 振込手数料（旧データなし → 0）
    newRow[13] = 0;       // 入金金額（旧データなし → 0）
    newRow[14] = row[12]; // ステータス（旧: Status → O列へ）
    newRow[15] = row[13]; // 備考（旧: Notes → P列へ）
    newRow[16] = row[14]; // メールID（旧: EmailId → Q列へ）
    return newRow;
  });

  // シートをクリアして新構造で書き直し
  sheet.clearContents();
  applyReservationHeaders_(sheet);
  if (newRows.length > 0) {
    sheet.getRange(2, 1, newRows.length, 17).setValues(newRows);
  }
  applyReservationColumnWidths_(sheet);

  Logger.log(`マイグレーション完了: ${newRows.length}件を17列構造に変換しました`);

  try {
    SpreadsheetApp.getUi().alert(
      `✅ マイグレーション完了\n\n${newRows.length}件のデータを新しい17列構造に変換しました。\n\n` +
      '※ 宿泊料・振込手数料・入金金額は旧データに存在しないため0になっています。\n' +
      '再バックフィルを実行すると正確な値が取り込まれます。'
    );
  } catch (e) {}
}

function applyReservationHeaders_(sheet) {
  const headers = [
    '予約ID', 'プラットフォーム', '予約受付日', 'チェックイン', 'チェックアウト',
    '宿泊数', '人数', 'ゲスト名',
    '売上', '宿泊料', '清掃費', 'OTA手数料', '振込手数料', '入金金額',
    'ステータス', '備考', 'メールID'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleSheet_(sheet, headers.length, { headerBg: '#34a853', headerColor: '#ffffff', freeze: 1 });
}

function applyReservationColumnWidths_(sheet) {
  const widths = [150, 120, 110, 110, 110, 70, 60, 150, 100, 100, 80, 100, 90, 110, 80, 200, 180];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
}

function removeDefaultSheets_(ss) {
  const defaultNames = ['Sheet1', 'シート1', 'sheet1'];
  defaultNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet && ss.getSheets().length > 1) {
      try {
        ss.deleteSheet(sheet);
      } catch (e) {
        // 削除できない場合は無視
      }
    }
  });
}
