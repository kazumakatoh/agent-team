/**
 * キャッシュフロー管理システム - セットアップモジュール
 *
 * 初期シート作成・ヘッダー設定・トリガー設定
 */

/**
 * 初回セットアップを実行する
 */
function runSetup() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    '初期セットアップ',
    '以下のシートを作成します:\n\n' +
    '・Daily（日次入出金）\n' +
    '・CF005（PayPay 005 月別集計）\n' +
    '・CF003（PayPay 003 月別集計）\n' +
    '・西武信金（西武信用金庫 月別集計）\n' +
    '・月別（3口座合算サマリー）\n' +
    '・設定（マスタ設定）\n' +
    '・現残高（各口座の現在残高）\n\n' +
    '既にあるシートはスキップします。',
    ui.ButtonSet.OK_CANCEL
  );

  if (result !== ui.Button.OK) return;

  const ss = getCfSpreadsheet();

  createDailySheet_(ss);
  createAccountSheet_(ss, 'CF005');
  createAccountSheet_(ss, 'CF003');
  createAccountSheet_(ss, '西武信金');
  createMonthlySheet_(ss);
  createSettingsSheet_(ss);
  createCurrentBalanceSheet_(ss);

  ui.alert('✅ セットアップ完了！\n\n次にメニューから「MF連携開始」を実行してください。');
}

/**
 * Dailyシートを作成する
 */
function createDailySheet_(ss) {
  const name = CF_CONFIG.SHEETS.DAILY;
  let sheet = ss.getSheetByName(name);
  if (sheet) {
    Logger.log(`${name}シートは既に存在します。スキップ。`);
    return;
  }

  sheet = ss.insertSheet(name);

  // ヘッダー行
  const headers = [
    ['合計', '日付', '', '内容', '入金', '出金', '残高', 'ソース', '',
     '日付', '内容', '入金', '出金', '残高', 'ソース', '',
     '日付', '内容', '入金', '出金', '残高', 'ソース']
  ];

  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);

  // 口座名のマージヘッダー（行0に追加）
  sheet.insertRowBefore(1);
  sheet.getRange(1, 2).setValue(CF_CONFIG.ACCOUNTS.CF005.shortName);
  sheet.getRange(1, 2, 1, 7).merge()
    .setBackground('#1565c0').setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.getRange(1, 10).setValue(CF_CONFIG.ACCOUNTS.CF003.shortName);
  sheet.getRange(1, 10, 1, 6).merge()
    .setBackground('#2e7d32').setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.getRange(1, 17).setValue(CF_CONFIG.ACCOUNTS.SEIBU.shortName);
  sheet.getRange(1, 17, 1, 6).merge()
    .setBackground('#e65100').setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center');

  // 合計列のヘッダー
  sheet.getRange(1, 1, 1, 1)
    .setValue('3口座合計')
    .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center');

  // サブヘッダー行のスタイル
  sheet.getRange(2, 1, 1, 22)
    .setBackground('#e3f2fd').setFontWeight('bold')
    .setHorizontalAlignment('center');

  // 列幅設定
  sheet.setColumnWidth(1, 100);  // 合計

  // PayPay 005
  sheet.setColumnWidth(2, 85);   // 日付
  sheet.setColumnWidth(3, 10);   // 空白
  sheet.setColumnWidth(4, 150);  // 内容
  sheet.setColumnWidth(5, 100);  // 入金
  sheet.setColumnWidth(6, 100);  // 出金
  sheet.setColumnWidth(7, 100);  // 残高
  sheet.setColumnWidth(8, 50);   // ソース

  sheet.setColumnWidth(9, 10);   // 空白

  // PayPay 003
  sheet.setColumnWidth(10, 85);
  sheet.setColumnWidth(11, 150);
  sheet.setColumnWidth(12, 100);
  sheet.setColumnWidth(13, 100);
  sheet.setColumnWidth(14, 100);
  sheet.setColumnWidth(15, 50);

  sheet.setColumnWidth(16, 10);  // 空白

  // 西武信金
  sheet.setColumnWidth(17, 85);
  sheet.setColumnWidth(18, 150);
  sheet.setColumnWidth(19, 100);
  sheet.setColumnWidth(20, 100);
  sheet.setColumnWidth(21, 100);
  sheet.setColumnWidth(22, 50);

  // ヘッダー行を固定（マージヘッダー + サブヘッダー = 2行）
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(1);

  // DAILY_HEADER_ROWS を2に更新（実行時）
  Logger.log('Dailyシート作成完了（ヘッダー2行）');
}

/**
 * 口座別月次シートを作成する
 */
function createAccountSheet_(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    Logger.log(`${sheetName}シートは既に存在します。スキップ。`);
    return;
  }

  sheet = ss.insertSheet(sheetName);

  // 基本構造だけ作成（データは updateAccountMonthlySheet で書き込む）
  const today = new Date();
  sheet.getRange(1, 1).setValue(today.getFullYear() + '年');
  sheet.getRange(1, 1).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');

  for (let m = 1; m <= 12; m++) {
    sheet.getRange(1, m + 1).setValue(m + '月');
    sheet.getRange(1, m + 1)
      .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff')
      .setHorizontalAlignment('center');
  }

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  sheet.setColumnWidth(1, 140);

  Logger.log(`${sheetName}シート作成完了`);
}

/**
 * 月別シート（3口座合算）を作成する
 */
function createMonthlySheet_(ss) {
  const name = CF_CONFIG.SHEETS.MONTHLY;
  let sheet = ss.getSheetByName(name);
  if (sheet) {
    Logger.log(`${name}シートは既に存在します。スキップ。`);
    return;
  }

  sheet = ss.insertSheet(name);

  const headers = ['', '入金', '出金', '差',
    'PayPay銀行\n法人口座', 'PayPay銀行\n個人口座', '西武信用金庫', '残高'];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1a73e8').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center')
    .setWrap(true);

  for (let m = 1; m <= 12; m++) {
    sheet.getRange(m + 1, 1).setValue(m + '月').setFontWeight('bold');
  }

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 60);
  for (let i = 2; i <= headers.length; i++) sheet.setColumnWidth(i, 110);

  Logger.log(`${name}シート作成完了`);
}

/**
 * 設定シートを作成する
 */
function createSettingsSheet_(ss) {
  const name = CF_CONFIG.SHEETS.SETTINGS;
  let sheet = ss.getSheetByName(name);
  if (sheet) {
    Logger.log(`${name}シートは既に存在します。スキップ。`);
    return;
  }

  sheet = ss.insertSheet(name);

  // MF連携情報
  sheet.getRange(1, 1).setValue('マネーフォワード連携設定');
  sheet.getRange(1, 1, 1, 3).merge()
    .setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');

  const settings = [
    ['Client ID', '', '※MFアプリポータルで発行'],
    ['Client Secret', '', '※MFアプリポータルで発行'],
    ['事業所ID', '', '※MF連携後に自動設定'],
    ['', '', ''],
    ['口座マッピング', '', ''],
    ['CF005 (PayPay ビジネス)', '', 'MFの口座名（自動設定）'],
    ['CF003 (PayPay はやぶさ)', '', 'MFの口座名（自動設定）'],
    ['SEIBU (西武信金)', '', 'MFの口座名（自動設定）'],
    ['', '', ''],
    ['アラート設定', '', ''],
    ['危険水準', '5,000,000', 'PayPay 005残高がこの金額以下で🔴'],
    ['注意水準', '10,000,000', 'PayPay 005残高がこの金額以下で🟡']
  ];

  sheet.getRange(2, 1, settings.length, 3).setValues(settings);

  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 250);

  // ヘッダー行のスタイル
  [5, 10].forEach(row => {
    sheet.getRange(row + 1, 1, 1, 3).setBackground('#e8eaf6').setFontWeight('bold');
  });

  Logger.log(`${name}シート作成完了`);
}

/**
 * 現残高シートを作成する
 */
function createCurrentBalanceSheet_(ss) {
  const name = CF_CONFIG.SHEETS.CURRENT_BAL;
  let sheet = ss.getSheetByName(name);
  if (sheet) {
    Logger.log(`${name}シートは既に存在します。スキップ。`);
    return;
  }

  sheet = ss.insertSheet(name);

  const headers = ['口座', '残高', '最終同期', 'ステータス'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');

  // 口座行
  let row = 2;
  Object.entries(CF_CONFIG.ACCOUNTS).forEach(([key, account]) => {
    sheet.getRange(row, 1).setValue(account.name);
    sheet.getRange(row, 2).setValue(0).setNumberFormat('#,##0');
    sheet.getRange(row, 3).setValue('未同期');
    sheet.getRange(row, 4).setValue('--');
    row++;
  });

  // 合計行
  sheet.getRange(row, 1).setValue('合計').setFontWeight('bold');
  sheet.getRange(row, 2).setFormula('=SUM(B2:B4)').setNumberFormat('#,##0').setFontWeight('bold');
  sheet.getRange(row, 1, 1, 4).setBackground('#e8eaf6');

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 80);

  sheet.setFrozenRows(1);

  Logger.log(`${name}シート作成完了`);
}
