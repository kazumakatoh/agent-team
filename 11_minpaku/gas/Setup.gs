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

  const headers = [
    '予約ID', 'プラットフォーム', '予約受付日', 'チェックイン', 'チェックアウト',
    '宿泊数', '人数', 'ゲスト名', '売上', '手数料', '清掃費', '純売上',
    'ステータス', '備考', 'メールID'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleSheet_(sheet, headers.length, {
    headerBg:     '#34a853',
    headerColor:  '#ffffff',
    freeze:       1
  });

  // 列幅設定
  sheet.setColumnWidth(1, 150);  // 予約ID
  sheet.setColumnWidth(2, 120);  // プラットフォーム
  sheet.setColumnWidth(3, 110);  // 予約受付日
  sheet.setColumnWidth(4, 110);  // チェックイン
  sheet.setColumnWidth(5, 110);  // チェックアウト
  sheet.setColumnWidth(6, 70);   // 宿泊数
  sheet.setColumnWidth(7, 60);   // 人数
  sheet.setColumnWidth(8, 150);  // ゲスト名
  sheet.setColumnWidth(9, 100);  // 売上
  sheet.setColumnWidth(10, 90);  // 手数料
  sheet.setColumnWidth(11, 80);  // 清掃費
  sheet.setColumnWidth(12, 100); // 純売上
  sheet.setColumnWidth(13, 80);  // ステータス
  sheet.setColumnWidth(14, 200); // 備考
  sheet.setColumnWidth(15, 180); // メールID

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
