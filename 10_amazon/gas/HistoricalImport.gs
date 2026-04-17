/**
 * Amazon Dashboard - 既存Excel過去3年データ一括インポート
 *
 * ## 使い方（社長向け手順）
 *
 * 1. GASエディタで `setupHistoricalImportSheet()` を1回実行
 *    → メインスプレッドシートに `インポート_履歴データ` シートが作成される
 *
 * 2. 既存Excelの月次/週次データを、そのシートの所定列にコピペする
 *    - A列: 年月（YYYY-MM）
 *    - B列: ASIN
 *    - C列: 商品名（任意・M1から自動補完）
 *    - D列: カテゴリ（任意・M1から自動補完）
 *    - E列: 売上金額
 *    - F列: CV（注文件数）
 *    - G列: 注文点数
 *    - H列: セッション数（任意）
 *    - I列: PV（任意）
 *    - J列: 広告費
 *    - K列: 広告売上
 *    - L列: 販売手数料合計（任意・Settlement代替）
 *    - M列: その他経費合計（任意・Settlement代替）
 *
 * 3. GASエディタで `importHistoricalData()` を実行
 *    → D1 日次データに YYYY-MM-01 として書き込み
 *    → D2S 経費月次集計に（L列/M列があれば）経費を書き込み
 *    → 同じ ASIN×年月 の既存行はスキップ（冪等）
 *
 * 4. `updateDashboardL1()` / `updateDashboardL2()` で反映確認
 *
 * ## 設計メモ
 *
 * - 既存ExcelはWeekly 10シート + Monthly 8シート構成（kpi_and_operations.md）
 * - 日次粒度での復元は不可能なため、月単位で YYYY-MM-01 として記録
 * - ステータスは「確定（履歴）」で通常の暫定/確定とは区別
 * - 経費列が空の場合、L1/L2 では経費ゼロとして計算されるので注意
 */

const HISTORICAL_IMPORT_COLS = 13; // A〜M列

const HISTORICAL_IMPORT_HEADERS = [
  '年月(YYYY-MM)', 'ASIN', '商品名', 'カテゴリ',
  '売上金額', 'CV(注文件数)', '注文点数',
  'セッション数', 'PV',
  '広告費', '広告売上',
  '販売手数料合計', 'その他経費合計',
];

const HISTORICAL_STATUS = '確定（履歴）';

/**
 * インポート用シートを作成・初期化
 * 既存のデータは保持し、ヘッダーと使い方ガイドのみ更新する
 */
function setupHistoricalImportSheet() {
  const ss = getMainSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.HISTORICAL_IMPORT);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.HISTORICAL_IMPORT);
    Logger.log('✅ インポート用シート新規作成: ' + SHEET_NAMES.HISTORICAL_IMPORT);
  } else {
    Logger.log('既存のインポート用シートを再利用: ' + SHEET_NAMES.HISTORICAL_IMPORT);
  }

  // 使い方ガイド（1〜3行目）
  const guide = [
    ['📥 既存Excel履歴データ インポート用', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['1. 4行目以降に月次データを貼り付け  2. importHistoricalData() を実行  3. ダッシュボードで確認',
     '', '', '', '', '', '', '', '', '', '', '', ''],
    HISTORICAL_IMPORT_HEADERS,
  ];

  sheet.getRange(1, 1, guide.length, HISTORICAL_IMPORT_COLS).setValues(guide);
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  sheet.getRange(2, 1, 1, HISTORICAL_IMPORT_COLS).setFontColor('#666').setFontStyle('italic');
  sheet.getRange(3, 1, 1, HISTORICAL_IMPORT_COLS)
    .setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center');
  sheet.setFrozenRows(3);

  // 列幅調整
  sheet.setColumnWidth(1, 110);   // 年月
  sheet.setColumnWidth(2, 110);   // ASIN
  sheet.setColumnWidth(3, 180);   // 商品名
  sheet.setColumnWidth(4, 120);   // カテゴリ

  Logger.log('✅ インポート用シート準備完了。4行目以降にデータを貼り付けてください。');
}

/**
 * メイン: インポート用シートから履歴データを取り込む
 */
function importHistoricalData() {
  const t0 = Date.now();
  Logger.log('===== 履歴データインポート開始 =====');

  const ss = getMainSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.HISTORICAL_IMPORT);
  if (!sheet) {
    Logger.log('❌ インポート用シートが未作成です。setupHistoricalImportSheet() を先に実行してください。');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 3) {
    Logger.log('❌ データが入力されていません（4行目以降にデータを貼り付けてください）');
    return;
  }

  const rows = sheet.getRange(4, 1, lastRow - 3, HISTORICAL_IMPORT_COLS).getValues();
  Logger.log('読み込み行数: ' + rows.length);

  // D1 既存データから (ASIN + 年月) の重複チェックキーを作る
  const existingKeys = getExistingD1MonthlyKeys();
  Logger.log('既存 D1 月次キー数: ' + existingKeys.size);

  // 商品マスター参照（商品名・カテゴリ補完用）
  const masterMap = getProductMasterMap();

  // 変換
  const d1Rows = [];
  const summaryRows = []; // D2S 経費月次集計用
  const skipped = [];
  const errors = [];

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const rawYm = String(r[0] || '').trim();
    const asin = String(r[1] || '').trim();

    if (!rawYm || !asin) continue; // 空行はスキップ

    // 年月パース（YYYY-MM または YYYY/MM または Date）
    const ym = normalizeYearMonth(rawYm);
    if (!ym) {
      errors.push(`行${i + 4}: 年月パース失敗 "${rawYm}"`);
      continue;
    }

    const key = asin + '_' + ym;
    if (existingKeys.has(key)) {
      skipped.push(key);
      continue;
    }

    const date = ym + '-01';
    const master = masterMap[asin] || {};
    const name = String(r[2] || '').trim() || master.name || '';
    const category = String(r[3] || '').trim() || master.category || '';

    const sales = parseFloat(r[4]) || 0;
    const cv = parseFloat(r[5]) || 0;
    const units = parseFloat(r[6]) || 0;
    const sessions = parseFloat(r[7]) || 0;
    const pv = parseFloat(r[8]) || 0;
    const adCost = parseFloat(r[9]) || 0;
    const adSales = parseFloat(r[10]) || 0;
    const commission = parseFloat(r[11]) || 0;
    const otherExpense = parseFloat(r[12]) || 0;

    // D1 日次データ形式（22列）
    d1Rows.push([
      date,              // 1 日付
      asin,              // 2 ASIN
      name,              // 3 商品名
      category,          // 4 カテゴリ
      sales,             // 5 売上金額
      cv,                // 6 CV(注文件数)
      units,             // 7 注文点数
      sessions,          // 8 セッション数
      pv,                // 9 PV
      '', '', '',        // 10 CTR, 11 CVR, 12 BuyBox率
      '', '', '',        // 13 FBA手数料, 14 返品数, 15 返品額
      adCost,            // 16 広告費
      adSales,           // 17 広告売上
      '', '',            // 18 IMP, 19 CT
      '', '',            // 20 仕入単価, 21 仕入原価合計
      HISTORICAL_STATUS, // 22 ステータス
    ]);

    // 経費が入力されていれば D2S にも追加（ダッシュボードの利益計算に反映される）
    if (commission > 0 || otherExpense > 0) {
      summaryRows.push([asin, ym, commission, otherExpense]);
    }
  }

  // 書き込み
  let d1Written = 0, d2sWritten = 0;
  if (d1Rows.length > 0) {
    appendRows(SHEET_NAMES.D1_DAILY, d1Rows);
    d1Written = d1Rows.length;
  }
  if (summaryRows.length > 0) {
    appendSettlementSummary(summaryRows);
    d2sWritten = summaryRows.length;
  }

  Logger.log('✅ D1 日次データ: ' + d1Written + ' 行追加');
  Logger.log('✅ D2S 経費月次集計: ' + d2sWritten + ' 行追加');
  Logger.log('⏭️ スキップ（既存）: ' + skipped.length + ' 件');
  if (errors.length > 0) {
    Logger.log('⚠️ エラー ' + errors.length + ' 件:');
    errors.forEach(e => Logger.log('  ' + e));
  }

  Logger.log('===== 履歴データインポート完了（' + (Date.now() - t0) + 'ms）=====');
}

/**
 * 年月文字列を YYYY-MM に正規化
 * 対応形式: "2025-03", "2025/03", "2025.3", Date object, "2025年3月"
 */
function normalizeYearMonth(raw) {
  if (raw instanceof Date) {
    const y = raw.getFullYear();
    const m = raw.getMonth() + 1;
    return y + '-' + (m < 10 ? '0' + m : String(m));
  }

  const s = String(raw).trim();

  // YYYY-MM / YYYY/MM / YYYY.MM
  let m = s.match(/^(\d{4})[-/.](\d{1,2})$/);
  if (m) {
    const mo = parseInt(m[2]);
    if (mo < 1 || mo > 12) return null;
    return m[1] + '-' + (mo < 10 ? '0' + mo : String(mo));
  }

  // YYYY年M月
  m = s.match(/^(\d{4})年(\d{1,2})月?$/);
  if (m) {
    const mo = parseInt(m[2]);
    if (mo < 1 || mo > 12) return null;
    return m[1] + '-' + (mo < 10 ? '0' + mo : String(mo));
  }

  // YYYYMM
  m = s.match(/^(\d{4})(\d{2})$/);
  if (m) {
    const mo = parseInt(m[2]);
    if (mo < 1 || mo > 12) return null;
    return m[1] + '-' + m[2];
  }

  return null;
}

/**
 * D1 日次データから (ASIN + 年月) の重複チェック用キーセットを取得
 */
function getExistingD1MonthlyKeys() {
  const sheet = getOrCreateSheet(SHEET_NAMES.D1_DAILY);
  const lastRow = sheet.getLastRow();
  const keys = new Set();
  if (lastRow <= 1) return keys;

  // 列1（日付）と列2（ASIN）だけ取得
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (const row of data) {
    const rawDate = row[0];
    const asin = String(row[1] || '').trim();
    if (!asin) continue;

    let ym = '';
    if (rawDate instanceof Date) {
      const y = rawDate.getFullYear();
      const m = rawDate.getMonth() + 1;
      ym = y + '-' + (m < 10 ? '0' + m : String(m));
    } else if (rawDate) {
      ym = String(rawDate).substring(0, 7);
    }
    if (ym) keys.add(asin + '_' + ym);
  }
  return keys;
}

/**
 * D2S 経費月次集計シートに履歴経費を追記
 * 既存の (ASIN + 年月) キーは上書きせず skip
 */
function appendSettlementSummary(rows) {
  const sheet = getOrCreateSheet(SHEET_NAMES.D2S_SETTLEMENT_SUMMARY);
  const lastRow = sheet.getLastRow();

  // 既存キーを収集
  const existing = new Set();
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    data.forEach(r => existing.add(String(r[0]).trim() + '_' + String(r[1]).trim()));
  }

  // ヘッダーがなければ追加（Principal売上列も含む5列）
  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, 5).setValues([['ASIN', '年月', '販売手数料', 'その他経費', 'Principal売上']])
      .setFontWeight('bold').setBackground('#e8f0fe');
    sheet.setFrozenRows(1);
  }

  const newRows = rows.filter(r => !existing.has(r[0] + '_' + r[1]));
  if (newRows.length === 0) return;

  const writeRow = Math.max(lastRow + 1, 2);
  sheet.getRange(writeRow, 1, newRows.length, 4).setValues(newRows);
  sheet.getRange(writeRow, 3, newRows.length, 2).setNumberFormat('#,##0');
}

/**
 * テスト: インポート用シートの状態を確認（書き込みなし）
 */
function testHistoricalImportDryRun() {
  const ss = getMainSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.HISTORICAL_IMPORT);
  if (!sheet) {
    Logger.log('❌ インポート用シート未作成');
    return;
  }

  const lastRow = sheet.getLastRow();
  Logger.log('インポート用シート 総行数: ' + lastRow);
  if (lastRow <= 3) {
    Logger.log('データ未入力');
    return;
  }

  const rows = sheet.getRange(4, 1, Math.min(5, lastRow - 3), HISTORICAL_IMPORT_COLS).getValues();
  Logger.log('先頭サンプル:');
  rows.forEach((r, i) => {
    const ym = normalizeYearMonth(r[0]);
    Logger.log('  行' + (i + 4) + ': 年月=' + ym + ' ASIN=' + r[1] + ' 売上=' + r[4]);
  });
}
