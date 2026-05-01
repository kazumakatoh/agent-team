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
    '・設定（マスタ設定）\n\n' +
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
  createPlanMasterSheet_(ss);
  createSettingsSheet_(ss);

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
 * 予定マスタシートを作成する
 *
 * 列構成:
 *  A: 口座 (CF005/CF003/SEIBU/RAKUTEN)
 *  B: 内容
 *  C: 金額
 *  D: 区分 (入/出)
 *  E: 頻度 (monthly/bimonthly/yearly)
 *  F: 発生日 (monthly/bimonthly: 1-31 or "last"(最終営業日) or "end"(月末日) / yearly: MM/DD形式)
 *  G: 開始年月 (2026.01)
 *  H: 終了年月 (2027.12)
 *  I: 備考
 */
function createPlanMasterSheet_(ss) {
  const name = '予定マスタ';
  let sheet = ss.getSheetByName(name);
  if (sheet) {
    Logger.log(`${name}は既に存在します。スキップ。`);
    return;
  }

  sheet = ss.insertSheet(name);

  const headers = ['口座', '内容', '金額', '区分', '頻度', '発生日', '開始年月', '終了年月', '備考'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1a73e8').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');

  // サンプル行
  const samples = [
    ['CF005', '家賃',           125830, '出', 'monthly', '26',    '2026.01', '2027.12', '事務所家賃'],
    ['CF005', '役員報酬①',        50000, '出', 'monthly', '22',    '2026.01', '2027.12', ''],
    ['CF005', '役員報酬②',        50000, '出', 'monthly', '22',    '2026.01', '2027.12', ''],
    ['CF005', '税理士',          49895, '出', 'monthly', 'last',  '2026.01', '2027.12', '最終営業日'],
    ['CF005', 'Amazon売上',     6000000, '入', 'monthly', '15',    '2026.01', '2027.12', 'Amazon入金1'],
    ['CF005', 'Amazon売上',     6000000, '入', 'monthly', 'end',   '2026.01', '2027.12', 'Amazon入金2（月末）'],
    ['SEIBU', '日本政策金融公庫', 345000, '出', 'monthly', '25',    '2026.01', '2030.02', '新企業育成'],
    ['SEIBU', '西武信用金庫返済', 129000, '出', 'monthly', '28',    '2026.01', '2031.08', '制度融資1000万'],
    ['CF005', 'JALカード年会費',  20167, '出', 'yearly',  '06/10', '2026.01', '2030.12', ''],
    ['CF005', 'セゾン年会費',     22000, '出', 'yearly',  '03/15', '2026.01', '2030.12', ''],
  ];

  sheet.getRange(2, 1, samples.length, headers.length).setValues(samples);
  sheet.getRange(2, 3, samples.length, 1).setNumberFormat('#,##0');

  // 説明行
  const noteRow = samples.length + 3;
  sheet.getRange(noteRow, 1).setValue('【入力ガイド】').setFontWeight('bold');
  const notes = [
    ['口座', 'CF005=PayPay005 / CF003=PayPay003 / SEIBU=西武信金 / RAKUTEN=楽天銀行'],
    ['金額', '0円でもOK（後から個別に変更可能）'],
    ['区分', '入 = 入金 / 出 = 出金'],
    ['頻度', 'monthly = 毎月 / bimonthly = 2ヶ月に1回 / yearly = 毎年'],
    ['発生日', 'monthly/bimonthly: 1〜31 の数字、または "last"(最終営業日) / "end"(月末日) / yearly: "MM/DD" 形式'],
    ['開始年月', '2026.01 形式'],
    ['終了年月', '2027.12 形式'],
  ];
  sheet.getRange(noteRow + 1, 1, notes.length, 2).setValues(notes);
  sheet.getRange(noteRow, 1, notes.length + 1, 2).setBackground('#f5f5f5');

  // 列幅
  sheet.setColumnWidth(1, 80);   // 口座
  sheet.setColumnWidth(2, 180);  // 内容
  sheet.setColumnWidth(3, 100);  // 金額
  sheet.setColumnWidth(4, 50);   // 区分
  sheet.setColumnWidth(5, 90);   // 頻度
  sheet.setColumnWidth(6, 70);   // 発生日
  sheet.setColumnWidth(7, 90);   // 開始年月
  sheet.setColumnWidth(8, 90);   // 終了年月
  sheet.setColumnWidth(9, 180);  // 備考

  // 年月列はテキスト形式
  sheet.getRange(2, 7, 100, 2).setNumberFormat('@');

  sheet.setFrozenRows(1);

  Logger.log('予定マスタシート作成完了');
}

/**
 * 実口座残高管理シートを作成する（月次推移形式）
 *
 * ① 実質残高 = 普通預金 − 未払金 − 未払費用 − 預り金 + 売掛金 + Amazon残高
 * ② 実質残高 = ① + 在庫想定入金（想定売上 × 入金割合）
 * ③ 実質残高 = ② − 長期借入金
 *
 * ■ MF API自動入力: 普通預金, 売掛金, 未払金, 未払費用, 預り金, 商品在庫, 融資残高
 * ■ 手入力: Amazon残高, 入金割合, 想定売上
 */
function createRealBalanceSheet_(ss) {
  const name = '実口座残高';
  let sheet = ss.getSheetByName(name);
  if (sheet) {
    Logger.log(`${name}は既に存在します。スキップ。`);
    return;
  }

  sheet = ss.insertSheet(name);

  // ===== ヘッダー行1: タイトル =====
  sheet.getRange(1, 1).setValue('■実口座残高算出表');
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(12);

  // ===== ヘッダー行2: 列ラベル =====
  const headers = [
    '年月',
    '', // B列予備
    '普通預金', '売掛金',                 // 資産 C,D
    '未払金', '未払費用', '預り金',       // 負債 E,F,G
    '実質残高①',                         // H = C-E-F-G+D
    '商品在庫', '想定売上', '入金割合', '想定入金',  // 在庫関連 I,J,K,L
    '実質残高②',                         // M = H+L
    '融資残高',                           // N
    '実質残高③',                         // O = M-N
    '判定①', '判定②', '判定③'           // P,Q,R
  ];

  sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, 1, headers.length)
    .setBackground('#1a73e8').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center')
    .setWrap(true);

  // 色分け
  // 自動入力（MF）= 濃い青
  [3,4,5,6,7,9,14].forEach(col => {
    sheet.getRange(2, col).setBackground('#1565c0');
  });
  // 手入力 = オレンジ（入金割合のみ）
  [11].forEach(col => {
    sheet.getRange(2, col).setBackground('#e65100');
  });
  // 計算結果 = 緑
  [8,12,13,15].forEach(col => {
    sheet.getRange(2, col).setBackground('#2e7d32');
  });
  // 判定 = 濃いグレー
  [16,17,18].forEach(col => {
    sheet.getRange(2, col).setBackground('#37474f');
  });

  // ===== 2025年度(4月〜2月) + 2026年度(3月〜) =====
  // 在庫残高シートの年月表示（2025.04形式）に統一
  // 在庫残高シートは5行/月構成: 在庫→見込売上→棚卸原価→仕入原価→販売単価
  // 見込売上の合計はD列

  // 2023.03 〜 2027.05 を自動生成
  const allMonths = [];
  const startYear = 2023, startMonth = 3;
  const endYear = 2027, endMonth = 5;
  let y = startYear, mo = startMonth;
  while (y < endYear || (y === endYear && mo <= endMonth)) {
    allMonths.push({
      label: `${y}.${String(mo).padStart(2, '0')}`,
      year: y,
      month: mo
    });
    mo++;
    if (mo > 12) { mo = 1; y++; }
  }

  // 在庫残高シートの見込売上行マップ（年月 → 行番号）
  // 在庫残高シートは5行/月: 在庫(+0), 見込売上(+1), 棚卸原価(+2), 仕入原価(+3), 販売単価(+4)
  // 2023.03が行3開始、各月5行 → 年月からの行番号計算
  // 基点: 2023.03 = 行3
  const inventoryBaseRow = 3;
  const inventoryBaseYear = 2023;
  const inventoryBaseMonth = 3;

  function getInventorySalesRow(year, month) {
    const monthsFromBase = (year - inventoryBaseYear) * 12 + (month - inventoryBaseMonth);
    return inventoryBaseRow + monthsFromBase * 5 + 1; // +1 で見込売上行
  }

  let currentRow = 3;

  allMonths.forEach(m => {
    const row = currentRow;

    // A列: 年月ラベル（2025.04形式）テキストとして設定
    sheet.getRange(row, 1).setNumberFormat('@').setValue(m.label);

    // B列は空（旧「月末」列を廃止、A列に統合）

    // 在庫残高シートの見込売上を自動参照（J列: 想定売上）
    const salesRow = getInventorySalesRow(m.year, m.month);
    sheet.getRange(row, 10).setFormula(`='在庫残高'!D${salesRow}`);

    // H: ① 実質残高 = C - E - F - G + D
    sheet.getRange(row, 8).setFormula(`=C${row}-E${row}-F${row}-G${row}+D${row}`)
      .setNumberFormat('#,##0').setBackground('#e8f5e9');

    // L: 想定入金 = J × K
    sheet.getRange(row, 12).setFormula(`=J${row}*K${row}`).setNumberFormat('#,##0');

    // M: ② 実質残高 = ① + 想定入金
    sheet.getRange(row, 13).setFormula(`=H${row}+L${row}`)
      .setNumberFormat('#,##0').setBackground('#e8f5e9');

    // O: ③ 実質残高 = ② - 融資残高
    sheet.getRange(row, 15).setFormula(`=M${row}-N${row}`)
      .setNumberFormat('#,##0').setBackground('#e8f5e9');

    // P: 判定① 実質残高① vs 融資残高
    sheet.getRange(row, 16).setFormula(
      `=IF(N${row}=0,"",IF(H${row}>=N${row},"✅","❌ "&TEXT(N${row}-H${row},"#,##0")))`
    );

    // Q: 判定② 実質残高② vs 融資残高
    sheet.getRange(row, 17).setFormula(
      `=IF(N${row}=0,"",IF(M${row}>=N${row},"✅","❌ "&TEXT(N${row}-M${row},"#,##0")))`
    );

    // R: 判定③ 実質残高③
    sheet.getRange(row, 18).setFormula(
      `=IF(N${row}=0,"",IF(O${row}>=0,"✅ 無借金","❌ "&TEXT(ABS(O${row}),"#,##0")))`
    );

    // 金額フォーマット
    [3,4,5,6,7,9,14].forEach(col => {
      sheet.getRange(row, col).setNumberFormat('#,##0');
    });
    sheet.getRange(row, 10).setNumberFormat('#,##0'); // 想定売上
    sheet.getRange(row, 11).setNumberFormat('0.0%');  // 入金割合

    // 手入力セルの背景色（入金割合のみ）
    [11].forEach(col => {
      sheet.getRange(row, col).setBackground('#fff9c4');
    });

    currentRow++;
  });

  // 列幅
  sheet.setColumnWidth(1, 60);  // 年月
  sheet.setColumnWidth(2, 30);  // 予備
  [3,4,5,6,7].forEach(col => sheet.setColumnWidth(col, 90));
  sheet.setColumnWidth(8, 100);  // ①
  sheet.setColumnWidth(9, 90);   // 商品在庫
  sheet.setColumnWidth(10, 90);  // 想定売上
  sheet.setColumnWidth(11, 60);  // 入金割合
  sheet.setColumnWidth(12, 90);  // 想定入金
  sheet.setColumnWidth(13, 100); // ②
  sheet.setColumnWidth(14, 100); // 融資残高
  sheet.setColumnWidth(15, 110); // ③
  sheet.setColumnWidth(16, 100); // 判定①
  sheet.setColumnWidth(17, 100); // 判定②
  sheet.setColumnWidth(18, 110); // 判定③

  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(1);

  Logger.log('実口座残高シート作成完了（月次推移形式）');
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
    ['RAKUTEN (楽天銀行)', '', 'MFの口座名（自動設定）'],
    ['', '', ''],
    ['アラート設定', '', ''],
    ['危険水準', '5,000,000', '監視対象口座（PayPay 005／楽天銀行）の残高がこの金額以下で🔴'],
    ['注意水準', '10,000,000', '監視対象口座（PayPay 005／楽天銀行）の残高がこの金額以下で🟡']
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
