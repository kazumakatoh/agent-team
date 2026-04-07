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
    '・月別（月次集計）\n' +
    '・設定（マスタ設定）\n' +
    '・現残高（各口座の現在残高）\n\n' +
    '既にあるシートはスキップします。',
    ui.ButtonSet.OK_CANCEL
  );

  if (result !== ui.Button.OK) return;

  const ss = getCfSpreadsheet();

  // 口座別Dailyシートを作成
  Object.entries(CF_CONFIG.ACCOUNTS).forEach(([key, account]) => {
    createAccountDailySheet_(ss, account.dailySheet, account.shortName);
  });

  createMonthlySheet_(ss);
  createFixedCostSheet_(ss);
  createRealBalanceSheet_(ss);
  createSettingsSheet_(ss);
  createCurrentBalanceSheet_(ss);

  ui.alert('✅ セットアップ完了！\n\n次にメニューから「MF連携開始」を実行してください。');
}

/**
 * 口座別Dailyシートを作成する
 */
function createAccountDailySheet_(ss, sheetName, accountName) {
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    Logger.log(`${sheetName}は既に存在します。スキップ。`);
    return;
  }

  sheet = ss.insertSheet(sheetName);

  const headers = [['日付', '内容', '入金', '出金', '残高', 'ソース']];
  sheet.getRange(1, 1, 1, 6).setValues(headers);
  sheet.getRange(1, 1, 1, 6)
    .setBackground('#1a73e8').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');

  sheet.setColumnWidth(1, 90);   // 日付
  sheet.setColumnWidth(2, 200);  // 内容
  sheet.setColumnWidth(3, 100);  // 入金
  sheet.setColumnWidth(4, 100);  // 出金
  sheet.setColumnWidth(5, 110);  // 残高
  sheet.setColumnWidth(6, 55);   // ソース

  sheet.setFrozenRows(1);

  Logger.log(`${sheetName}シート作成完了`);
}

/**
 * 固定費・融資シートを作成する
 */
function createFixedCostSheet_(ss) {
  const name = '固定費・融資';
  let sheet = ss.getSheetByName(name);
  if (sheet) {
    Logger.log(`${name}は既に存在します。スキップ。`);
    return;
  }

  sheet = ss.insertSheet(name);

  // ===== 固定費セクション =====
  sheet.getRange(1, 1, 1, 3).setValues([['固定費', '月額（税込）', '備考']]);
  sheet.getRange(1, 1, 1, 3)
    .setBackground('#1a73e8').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');

  const fixedCosts = [
    ['融資返済', 757000, '下記融資一覧参照'],
    ['役員報酬①', 50000, ''],
    ['役員報酬②', 50000, ''],
    ['Amazonコンサル', 44000, ''],
    ['税理士報酬', 50000, '風間良光税理士事務所'],
    ['Amazon販売', 5500, ''],
    ['MF会計', 5654, ''],
    ['zoom', 2200, ''],
    ['サーバー', 1500, ''],
    ['固定電話', 1000, ''],
    ['山本さんコンサル', 33000, ''],
    ['Keepa', 2284, ''],
    ['Google One', 3033, ''],
    ['PLAUD', 1000, ''],
    ['Canva', 1500, ''],
    ['ChatGPT', 2000, ''],
    ['Adobe Pro', 1980, ''],
    ['MindMeister', 1340, ''],
    ['Genspark', 0, ''],
    ['Seller Sprite', 0, ''],
    ['CKB', 20860, ''],
    ['セゾンプラチナ年会費', 1833, '年額÷12'],
    ['楽天ビジネス年会費', 182, '年額÷12'],
    ['三井住友VISA', 0, ''],
    ['Marriott Bonvoy 年会費', 6875, '年額÷12'],
    ['JAL LC 年会費', 20167, '年額÷12'],
    ['不動産コミュニティ', 5500, ''],
    ['PL保険 日本', 2567, ''],
    ['交通費', 5000, ''],
    ['振込手数料', 1000, ''],
    ['支払利息', 80000, '']
  ];

  sheet.getRange(2, 1, fixedCosts.length, 3).setValues(fixedCosts);
  sheet.getRange(2, 2, fixedCosts.length, 1).setNumberFormat('#,##0');

  // 合計行
  const totalRow = fixedCosts.length + 2;
  sheet.getRange(totalRow, 1).setValue('合計').setFontWeight('bold');
  sheet.getRange(totalRow, 2)
    .setFormula(`=SUM(B2:B${totalRow - 1})`)
    .setNumberFormat('#,##0').setFontWeight('bold');
  sheet.getRange(totalRow, 1, 1, 3).setBackground('#e8eaf6');

  // 月額経費（融資返済除く）
  sheet.getRange(totalRow + 1, 1).setValue('月額経費（融資返済除く）').setFontWeight('bold');
  sheet.getRange(totalRow + 1, 2)
    .setFormula(`=B${totalRow}-B2`)
    .setNumberFormat('#,##0').setFontWeight('bold');

  // ===== 融資情報セクション =====
  const loanStartRow = totalRow + 4;

  sheet.getRange(loanStartRow, 1, 1, 10).setValues([[
    '金融機関', '内容', '融資額', '残高', '実質利率', '月額返済', '完済時期', '残年', '残月', '備考'
  ]]);
  sheet.getRange(loanStartRow, 1, 1, 10)
    .setBackground('#e65100').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');

  const loans = [
    ['東京東信用金庫', '創業融資', '200万円', 0, '', 0, '済', '', '', '完済'],
    ['日本政策金融公庫', '創業融資', '300万円', 0, '', 50000, '済', '', '', '完済'],
    ['東京東信用金庫', '創業融資', '1,000万円', 0, '', 170000, '済', '', '', '完済'],
    ['日本政策金融公庫', '資本性ローン', '700万円', 7000000, '', 0, '2031年8月', 5, 5, '一括返済'],
    ['日本政策金融公庫', '新企業育成', '2,000万円', 16560000, '', 345000, '2030年2月', 3, 11, ''],
    ['西武信用金庫', '制度融資', '300万円', 1632000, '', 38000, '2029年7月', 3, 5, ''],
    ['西武信用金庫', '制度融資', '1,000万円', 8581000, '', 129000, '2031年8月', 5, 6, ''],
    ['西武信用金庫', '制度融資', '900万円', 9000000, 0.60, 116000, '2032年12月', 6, 9, ''],
    ['西武信用金庫', '制度融資', '1,000万円', 10000000, 1.33, 129000, '2032年12月', 6, 9, '']
  ];

  sheet.getRange(loanStartRow + 1, 1, loans.length, 10).setValues(loans);

  // 金額フォーマット
  sheet.getRange(loanStartRow + 1, 4, loans.length, 1).setNumberFormat('#,##0');
  sheet.getRange(loanStartRow + 1, 6, loans.length, 1).setNumberFormat('#,##0');

  // 融資合計行
  const loanTotalRow = loanStartRow + loans.length + 1;
  sheet.getRange(loanTotalRow, 1).setValue('合計').setFontWeight('bold');
  sheet.getRange(loanTotalRow, 4)
    .setFormula(`=SUM(D${loanStartRow + 1}:D${loanTotalRow - 1})`)
    .setNumberFormat('#,##0').setFontWeight('bold');
  sheet.getRange(loanTotalRow, 6)
    .setFormula(`=SUM(F${loanStartRow + 1}:F${loanTotalRow - 1})`)
    .setNumberFormat('#,##0').setFontWeight('bold');
  sheet.getRange(loanTotalRow, 1, 1, 10).setBackground('#e8eaf6');

  // 列幅
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 70);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 90);
  sheet.setColumnWidth(8, 40);
  sheet.setColumnWidth(9, 40);
  sheet.setColumnWidth(10, 100);

  sheet.setFrozenRows(1);

  Logger.log('固定費・融資シート作成完了');
}

/**
 * 実口座残高管理シートを作成する
 *
 * ① 実口座残高（現実的）= 普通預金残高 − 未払金 − 未払費用 − 役員貸付金 + 売掛金 + Amazon未入金残高
 * ② 実口座残高（MAX）= ① + 在庫全売却想定入金額
 * ③ 実口座残高（無借金）= ② − 長期借入金
 */
function createRealBalanceSheet_(ss) {
  const name = '実口座残高';
  let sheet = ss.getSheetByName(name);
  if (sheet) {
    Logger.log(`${name}は既に存在します。スキップ。`);
    return;
  }

  sheet = ss.insertSheet(name);

  // タイトル
  sheet.getRange(1, 1, 1, 4).merge()
    .setValue('実口座残高管理')
    .setBackground('#1a73e8').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

  sheet.getRange(2, 1, 1, 4).merge()
    .setValue('作成日: ')
    .setHorizontalAlignment('right').setFontSize(10).setFontColor('#666666');

  // ===== 基礎データ セクション =====
  const baseRow = 4;
  sheet.getRange(baseRow, 1, 1, 4).setValues([['基礎データ', '', '金額', '備考']]);
  sheet.getRange(baseRow, 1, 1, 4)
    .setBackground('#37474f').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');

  const baseData = [
    ['普通預金残高（3口座合計）', '', 0, '現残高シートの合計値を転記'],
    ['', '', '', ''],
    ['  PayPay 005', '', 0, ''],
    ['  PayPay 003', '', 0, ''],
    ['  西武信用金庫', '', 0, ''],
  ];
  sheet.getRange(baseRow + 1, 1, baseData.length, 4).setValues(baseData);
  sheet.getRange(baseRow + 1, 3, baseData.length, 1).setNumberFormat('#,##0');
  sheet.getRange(baseRow + 1, 1, 1, 4).setBackground('#e3f2fd').setFontWeight('bold');

  // ===== 実口座残高① セクション =====
  const sec1Row = baseRow + baseData.length + 2;
  sheet.getRange(sec1Row, 1, 1, 4).setValues([['実口座残高①（現実的）', '', '金額', '備考']]);
  sheet.getRange(sec1Row, 1, 1, 4)
    .setBackground('#2e7d32').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');

  const sec1Data = [
    ['普通預金残高', '', 0, '上記の合計値'],
    ['', '', '', ''],
    ['【マイナス項目】', '', '', ''],
    ['(−) 未払金', '', 0, 'MF会計のBS → 未払金'],
    ['(−) 未払費用', '', 0, 'MF会計のBS → 未払費用'],
    ['(−) 役員貸付金', '', 0, 'MF会計のBS → 役員貸付金'],
    ['', '', '', ''],
    ['【プラス項目】', '', '', ''],
    ['(+) 売掛金', '', 0, 'MF会計のBS → 売掛金'],
    ['(+) Amazon未入金残高', '', 0, 'セラーセントラルの残高'],
    ['', '', '', ''],
  ];
  sheet.getRange(sec1Row + 1, 1, sec1Data.length, 4).setValues(sec1Data);
  sheet.getRange(sec1Row + 1, 3, sec1Data.length, 1).setNumberFormat('#,##0');

  // ①の計算結果行
  const result1Row = sec1Row + sec1Data.length + 1;
  sheet.getRange(result1Row, 1).setValue('★ 実口座残高①').setFontWeight('bold').setFontSize(12);
  sheet.getRange(result1Row, 3)
    .setFormula(`=C${sec1Row+1}-C${sec1Row+4}-C${sec1Row+5}-C${sec1Row+6}+C${sec1Row+9}+C${sec1Row+10}`)
    .setNumberFormat('#,##0').setFontWeight('bold').setFontSize(12);
  sheet.getRange(result1Row, 1, 1, 4).setBackground('#c8e6c9');
  sheet.getRange(result1Row, 4).setValue('現実的な手元資金');

  // ===== 実口座残高② セクション =====
  const sec2Row = result1Row + 3;
  sheet.getRange(sec2Row, 1, 1, 4).setValues([['実口座残高②（MAX）', '', '金額', '備考']]);
  sheet.getRange(sec2Row, 1, 1, 4)
    .setBackground('#1565c0').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');

  const sec2Data = [
    ['実口座残高①', '', 0, '上記①の値を転記'],
    ['', '', '', ''],
    ['【在庫全売却想定】', '', '', ''],
    ['(+) 在庫原価', '', 0, 'MF会計のBS → 商品'],
    ['(+) 想定売価（原価×掛け率）', '', 0, '掛け率は右に入力 →'],
    ['(−) Amazon手数料想定', '', 0, '手数料率は右に入力 →'],
    ['(=) 在庫売却後の想定入金額', '', 0, '売価 − 手数料'],
    ['', '', '', ''],
  ];
  sheet.getRange(sec2Row + 1, 1, sec2Data.length, 4).setValues(sec2Data);
  sheet.getRange(sec2Row + 1, 3, sec2Data.length, 1).setNumberFormat('#,##0');

  // 掛け率・手数料率入力セル
  sheet.getRange(sec2Row + 5, 4).setValue('掛け率: ');
  sheet.getRange(sec2Row + 5, 5).setValue(2.5).setNumberFormat('0.0').setBackground('#fff9c4');
  sheet.getRange(sec2Row + 6, 4).setValue('手数料率: ');
  sheet.getRange(sec2Row + 6, 5).setValue(0.15).setNumberFormat('0%').setBackground('#fff9c4');

  // ②の計算結果行
  const result2Row = sec2Row + sec2Data.length + 1;
  sheet.getRange(result2Row, 1).setValue('★ 実口座残高②').setFontWeight('bold').setFontSize(12);
  sheet.getRange(result2Row, 3).setValue(0)
    .setNumberFormat('#,##0').setFontWeight('bold').setFontSize(12);
  sheet.getRange(result2Row, 1, 1, 4).setBackground('#bbdefb');
  sheet.getRange(result2Row, 4).setValue('在庫含むMAX手元資金');

  // ===== 実口座残高③ セクション =====
  const sec3Row = result2Row + 3;
  sheet.getRange(sec3Row, 1, 1, 4).setValues([['実口座残高③（無借金指標）', '', '金額', '備考']]);
  sheet.getRange(sec3Row, 1, 1, 4)
    .setBackground('#b71c1c').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');

  const sec3Data = [
    ['実口座残高②', '', 0, '上記②の値を転記'],
    ['', '', '', ''],
    ['(−) 長期借入金', '', 0, '融資シートの残高合計を転記'],
    ['', '', '', ''],
  ];
  sheet.getRange(sec3Row + 1, 1, sec3Data.length, 4).setValues(sec3Data);
  sheet.getRange(sec3Row + 1, 3, sec3Data.length, 1).setNumberFormat('#,##0');

  // ③の計算結果行
  const result3Row = sec3Row + sec3Data.length + 1;
  sheet.getRange(result3Row, 1).setValue('★ 実口座残高③').setFontWeight('bold').setFontSize(12);
  sheet.getRange(result3Row, 3)
    .setFormula(`=C${sec3Row+1}-C${sec3Row+3}`)
    .setNumberFormat('#,##0').setFontWeight('bold').setFontSize(12);
  sheet.getRange(result3Row, 1, 1, 4).setBackground('#ffcdd2');
  sheet.getRange(result3Row, 4).setValue('全額返済後の残高（実質無借金指標）');

  // ===== 判定セクション =====
  const judgeRow = result3Row + 3;
  sheet.getRange(judgeRow, 1, 1, 4).setValues([['判定', '', '結果', '']]);
  sheet.getRange(judgeRow, 1, 1, 4)
    .setBackground('#37474f').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');

  sheet.getRange(judgeRow + 1, 1).setValue('③がプラス → 実質無借金経営');
  sheet.getRange(judgeRow + 2, 1).setValue('③がマイナス → 借入超過額');
  sheet.getRange(judgeRow + 2, 3)
    .setFormula(`=IF(C${result3Row}>=0,"✅ 実質無借金","❌ 借入超過 "&TEXT(ABS(C${result3Row}),"#,##0")&"円")`)
    .setFontWeight('bold').setFontSize(11);

  // 列幅
  sheet.setColumnWidth(1, 240);
  sheet.setColumnWidth(2, 30);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 200);
  sheet.setColumnWidth(5, 80);

  Logger.log('実口座残高シート作成完了');
}

/**
 * （廃止）口座別月次シートを作成する
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
