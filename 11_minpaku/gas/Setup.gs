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

  // 列: A年月 B代行手数料 C清掃費 Dリネン費 E備品・消耗品費 F水光熱費 G家賃 Hその他経費 I備考
  const headers = ['年月', '代行手数料', '清掃費', 'リネン費', '備品・消耗品費', '水光熱費', '家賃', 'その他経費', '備考'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleSheet_(sheet, headers.length, {
    headerBg:    '#ff6d00',
    headerColor: '#ffffff',
    freeze:       1
  });

  // 2025/04〜2030/03 の60ヶ月分をデフォルト値で事前入力
  // デフォルト: 家賃115,500円、水光熱費10,000円、その他0円
  const rows = [];
  const startDate = new Date(2025, 3, 1); // 2025年4月
  const endDate   = new Date(2030, 2, 1); // 2030年3月
  const d = new Date(startDate);
  while (d <= endDate) {
    const ym = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
    // [年月, 代行手数料, 清掃費, リネン費, 備品・消耗品費, 水光熱費, 家賃, その他経費, 備考]
    rows.push([ym, 0, 0, 0, 0, 10000, 115500, 0, '']);
    d.setMonth(d.getMonth() + 1);
  }
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  // 金額列のフォーマット（B〜H列）
  sheet.getRange('B:H').setNumberFormat('¥#,##0');

  // 列幅
  sheet.setColumnWidth(1, 100);  // 年月
  sheet.setColumnWidth(2, 110);  // 代行手数料
  sheet.setColumnWidth(3, 90);   // 清掃費
  sheet.setColumnWidth(4, 90);   // リネン費
  sheet.setColumnWidth(5, 120);  // 備品・消耗品費
  sheet.setColumnWidth(6, 100);  // 水光熱費
  sheet.setColumnWidth(7, 100);  // 家賃
  sheet.setColumnWidth(8, 110);  // その他経費
  sheet.setColumnWidth(9, 200);  // 備考

  // 入力ガイドのメモ
  sheet.getRange(1, 1).setNote('年月はYYYY-MM形式で入力\n例: 2025-12');
  sheet.getRange(1, 2).setNote('運営代行会社への報酬（月額）');
  sheet.getRange(1, 3).setNote('清掃会社への支払いなど');
  sheet.getRange(1, 4).setNote('リネン・タオル等のレンタル費用');

  Logger.log(`シート「${CONFIG.SHEETS.COSTS}」を作成しました`);
}

/**
 * 経費入力シートを新列構造に移行する（既存データを保持しつつ列を追加）
 * メニュー「経費入力シートを更新」から実行
 */
function migrateCostSheet() {
  const ss    = getSpreadsheet();
  const ui    = SpreadsheetApp.getUi();
  let sheet   = ss.getSheetByName(CONFIG.SHEETS.COSTS);

  // 既存データ（年月・金額）を退避
  const existing = {};
  if (sheet && sheet.getLastRow() > 1) {
    const oldData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    oldData.forEach(row => {
      const ym = String(row[0]).trim();
      if (!ym || ym.startsWith('例:')) return;
      // 旧列順: 年月(0) 清掃費(1) 備品(2) 水光熱(3) 家賃(4) その他(5) 備考(6)
      existing[ym] = {
        cleaning:   Number(row[1]) || 0,
        supplies:   Number(row[2]) || 0,
        utilities:  Number(row[3]) || 0,
        rent:       Number(row[4]) || 0,
        other:      Number(row[5]) || 0,
        notes:      row[6] || ''
      };
    });
  }

  // シートを完全クリアまたは新規作成
  if (sheet) {
    sheet.clearContents();
    sheet.clearFormats();
    sheet.clearNotes();
  } else {
    sheet = ss.insertSheet(CONFIG.SHEETS.COSTS);
  }

  // 新ヘッダー
  const headers = ['年月', '代行手数料', '清掃費', 'リネン費', '備品・消耗品費', '水光熱費', '家賃', 'その他経費', '備考'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleSheet_(sheet, headers.length, { headerBg: '#ff6d00', headerColor: '#ffffff', freeze: 1 });

  // 2025/04〜2030/03 の60ヶ月分を生成（既存値を引き継ぎ、なければデフォルト値）
  const rows = [];
  const d = new Date(2025, 3, 1);
  const endDate = new Date(2030, 2, 1);
  while (d <= endDate) {
    const ym  = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
    const old = existing[ym] || {};
    rows.push([
      ym,
      0,                                           // 代行手数料（新項目）
      old.cleaning  != null ? old.cleaning  : 0,  // 清掃費
      0,                                           // リネン費（新項目）
      old.supplies  != null ? old.supplies  : 0,  // 備品・消耗品費
      old.utilities != null ? old.utilities : 10000,  // 水光熱費
      old.rent      != null ? old.rent      : 115500, // 家賃
      old.other     != null ? old.other     : 0,  // その他経費
      old.notes     || ''                          // 備考
    ]);
    d.setMonth(d.getMonth() + 1);
  }
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  // 書式
  sheet.getRange('B:H').setNumberFormat('¥#,##0');
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 90);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 110);
  sheet.setColumnWidth(9, 200);
  sheet.getRange(1, 1).setNote('年月はYYYY-MM形式で入力\n例: 2025-12');
  sheet.getRange(1, 2).setNote('運営代行会社への報酬（月額）');
  sheet.getRange(1, 3).setNote('清掃会社への支払いなど');
  sheet.getRange(1, 4).setNote('リネン・タオル等のレンタル費用');

  ui.alert('✅ 経費入力シートを更新しました\n\n' +
    '・列追加: 代行手数料（B列）、リネン費（D列）\n' +
    '・既存の清掃費・備品・水光熱費・家賃・その他は引き継ぎました\n' +
    '・2025/04〜2030/03 の60ヶ月分を入力済み');
}

function setupMonthlySheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.MONTHLY);
  if (sheet) {
    Logger.log(`シート「${CONFIG.SHEETS.MONTHLY}」は既に存在します`);
    return;
  }

  sheet = ss.insertSheet(CONFIG.SHEETS.MONTHLY);

  const headers = [
    '年月', '問い合わせ数', '稼働日数', '利用可能日数', '利用件数', '利用人数',
    'OTA売上', '宿泊料', '清掃料', 'OTA手数料', '振込手数料', '入金金額',
    '代行手数料', '清掃費', 'リネン費', '備品・消耗品費', '水光熱費', '家賃', 'その他経費',
    '売上', '流動費', '固定費', '利益', '利益率(%)',
    'ADR(円)', 'RevPAR(円)', '稼働率(%)'
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
    '事業年度', '稼働日数', '年間上限日数', '利用可能日数', '利用件数', '利用人数',
    'OTA売上', '宿泊料', '清掃料', 'OTA手数料', '振込手数料', '入金金額',
    '代行手数料', '清掃費', 'リネン費', '備品・消耗品費', '水光熱費', '家賃', 'その他経費',
    '売上', '流動費', '固定費', '利益', '利益率(%)',
    'ADR(円)', 'RevPAR(円)', '稼働率(%)', '法定稼働率(%)'
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

  // 既存データを全件読み込み（旧17列想定）
  const oldData = sheet.getRange(2, 1, lastRow - 1, 17).getValues();

  // 旧列インデックス（0始まり）
  // 旧17列: ID=0, Platform=1, BookedDate=2, Checkin=3, Checkout=4,
  //         Nights=5, Guests=6, GuestName=7, Revenue=8, Accommodation=9,
  //         CleaningFee=10, OtaFee=11, TransferFee=12, Payout=13,
  //         Status=14, Notes=15, EmailId=16
  const newRows = oldData.map(row => {
    const nights    = Number(row[5]) || 0;
    const guests    = Number(row[6]) || 1;
    const usageDays = nights + 1;
    const newRow = new Array(19).fill('');
    newRow[0]  = row[0];       // 予約ID
    newRow[1]  = row[1];       // プラットフォーム
    newRow[2]  = row[2];       // 予約受付日
    newRow[3]  = row[3];       // チェックイン
    newRow[4]  = row[4];       // チェックアウト
    newRow[5]  = nights;       // 宿泊数
    newRow[6]  = guests;       // 人数
    newRow[7]  = usageDays;    // 利用日数（新列）
    newRow[8]  = usageDays * guests; // 総利用人数（新列）
    newRow[9]  = row[7];       // ゲスト名
    newRow[10] = row[8];       // 売上
    newRow[11] = row[9];       // 宿泊料
    newRow[12] = row[10];      // 清掃費
    newRow[13] = row[11];      // OTA手数料
    newRow[14] = row[12];      // 振込手数料
    newRow[15] = row[13];      // 入金金額
    newRow[16] = row[14];      // ステータス
    newRow[17] = row[15];      // 備考
    newRow[18] = row[16];      // メールID
    return newRow;
  });

  // シートをクリアして新構造で書き直し
  sheet.clearContents();
  applyReservationHeaders_(sheet);
  if (newRows.length > 0) {
    sheet.getRange(2, 1, newRows.length, 19).setValues(newRows);
  }
  applyReservationColumnWidths_(sheet);

  Logger.log(`マイグレーション完了: ${newRows.length}件を19列構造に変換しました`);

  try {
    SpreadsheetApp.getUi().alert(
      `✅ マイグレーション完了\n\n${newRows.length}件のデータを新しい19列構造に変換しました。\n\n` +
      '追加列: 利用日数（宿泊数+1）・総利用人数（利用日数×人数）\n' +
      '※ 既存データは宿泊数・人数から自動計算しました。'
    );
  } catch (e) {}
}

function applyReservationHeaders_(sheet) {
  const headers = [
    '予約ID', 'プラットフォーム', '予約受付日', 'チェックイン', 'チェックアウト',
    '宿泊数', '人数', '利用日数', '総利用人数', 'ゲスト名',
    '売上', '宿泊料', '清掃料', 'OTA手数料', '振込手数料', '入金金額',
    'ステータス', '備考', 'メールID'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleSheet_(sheet, headers.length, { headerBg: '#34a853', headerColor: '#ffffff', freeze: 1 });
}

function applyReservationColumnWidths_(sheet) {
  const widths = [150, 120, 110, 110, 110, 70, 60, 80, 100, 150, 100, 100, 80, 100, 90, 110, 80, 200, 180];
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
