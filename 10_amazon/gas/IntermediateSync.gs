/**
 * Amazon Dashboard - 中間スプシ（就業管理表）⇔ M3 双方向同期
 *
 * ## 中間スプシ「管理表」シートの構造
 *
 * 1シート1年（例: 2026年）の年×項目テーブル。
 * 列: B=項目名 / C=1月 / D=2月 / ... / N=12月 / O=年間 / P=%
 *
 * | 行 | 項目         | データソース            | 同期方向          |
 * |----|--------------|------------------------|-------------------|
 * | 2  | 販売数        | Amazon D1 日次データ    | M3 → 中間（自動）  |
 * | 5  | 支払報酬      | 中間スプシ「入力」シート | 中間 → M3（納品人件費）|
 * | 6  | FBA輸送手数料 | Amazon Settlement+Finance| M3経由 → 中間    |
 * | 7  | ヤマト運輸    | MF会計から手入力         | 中間 → M3（荷造運賃の一部）|
 *
 * ## M3 への合算ロジック
 *
 *  M3.荷造運賃    = 中間スプシ.ヤマト運輸 + Settlement+Finance.FBA輸送手数料
 *  M3.納品人件費 = 中間スプシ.支払報酬
 *
 * ## トリガー: 毎日 AM8:00 (syncIntermediateAndM3)
 */

/**
 * メイン: 中間スプシと M3 の双方向同期
 */
function syncIntermediateAndM3() {
  const t0 = Date.now();
  Logger.log('===== 中間スプシ ↔ M3 同期 開始 =====');

  // 1. 中間スプシ → M3: ヤマト運輸 + 支払報酬
  syncFromIntermediateToM3();

  // 2. M3 / Amazon → 中間スプシ: 販売数 + FBA輸送手数料
  syncToIntermediateFromAmazon();

  Logger.log('===== 同期完了（' + (Date.now() - t0) + 'ms）=====');
}

/**
 * 中間スプシ → M3
 *
 * 中間スプシの「管理表」シート行5(支払報酬)・行7(ヤマト運輸)を読み取り、
 * M3「販促費マスター」の納品人件費・荷造運賃に反映。
 *
 * 荷造運賃は (中間スプシ.ヤマト + Finance/Settlement.FBA輸送手数料) の合算。
 */
function syncFromIntermediateToM3() {
  Logger.log('--- 中間スプシ → M3 ---');

  const intSheetId = getIntermediateSheetId();
  if (!intSheetId) {
    Logger.log('⚠️ INTERMEDIATE_SHEET_ID が未設定。スキップ');
    return;
  }

  const intSs = SpreadsheetApp.openById(intSheetId);
  const mgmtSheet = intSs.getSheetByName(INTERMEDIATE_MGMT_SHEET);
  if (!mgmtSheet) {
    Logger.log('⚠️ 「' + INTERMEDIATE_MGMT_SHEET + '」シートが見つかりません');
    return;
  }

  // 「管理表」の年を取得（B1セルが "2026年" 等の想定）
  const yearLabel = String(mgmtSheet.getRange(1, 2).getValue() || '').trim();
  const yearMatch = yearLabel.match(/^(\d{4})/);
  if (!yearMatch) {
    Logger.log('⚠️ B1セルから年を取得できません: "' + yearLabel + '"');
    return;
  }
  const year = parseInt(yearMatch[1]);
  Logger.log('対象年: ' + year);

  // 12ヶ月分の値を一括読み込み（行5: 支払報酬, 行7: ヤマト）
  const payRewards = mgmtSheet.getRange(INTERMEDIATE_ROWS.PAY_REWARD, INTERMEDIATE_MONTH_COL_START, 1, 12).getValues()[0];
  const yamatoCosts = mgmtSheet.getRange(INTERMEDIATE_ROWS.YAMATO, INTERMEDIATE_MONTH_COL_START, 1, 12).getValues()[0];

  // FBA輸送手数料を Finance/Settlement から取得（年単位で月別集計）
  const fbaShippingByMonth = aggregateFbaInboundShippingByMonth(year);

  // M3 の現在データを読み込み
  const m3Sheet = getOrCreateSheet(SHEET_NAMES.M3_PROMO_COST);
  const m3LastRow = m3Sheet.getLastRow();
  if (m3LastRow <= 1) {
    Logger.log('⚠️ M3 が空です。setupPromoCostSheet() を先に実行してください');
    return;
  }
  const m3Data = m3Sheet.getRange(2, 1, m3LastRow - 1, M3_COLS).getValues();

  // 年月→行番号のマップを作る
  const ymToRow = {};
  m3Data.forEach((row, i) => {
    const ym = formatYearMonth(row[0]);
    if (ym) ymToRow[ym] = i + 2; // シート上の行番号
  });

  let updated = 0;
  for (let m = 0; m < 12; m++) {
    const month = m + 1;
    const ym = year + '-' + String(month).padStart(2, '0');
    const m3Row = ymToRow[ym];
    if (!m3Row) {
      Logger.log('  ' + ym + ': M3 に該当行なし、スキップ');
      continue;
    }

    const payReward = parseFloat(payRewards[m]) || 0;
    const yamato = parseFloat(yamatoCosts[m]) || 0;
    const fba = fbaShippingByMonth[ym] || 0;
    const shippingTotal = yamato + fba;

    // M3 列順: 年月(1) | Amence(2) | その他ツール(3) | 荷造運賃(4) | 納品人件費(5) | 備考(6)
    m3Sheet.getRange(m3Row, 4).setValue(shippingTotal);  // 荷造運賃
    m3Sheet.getRange(m3Row, 5).setValue(payReward);      // 納品人件費

    Logger.log('  ' + ym + ': 荷造運賃=' + shippingTotal + ' (ヤマト' + yamato + '+FBA' + fba + ') / 納品人件費=' + payReward);
    updated++;
  }

  Logger.log('✅ M3 更新: ' + updated + ' ヶ月');
}

/**
 * Amazon → 中間スプシ
 *
 * D1 から月次販売数を集計して、中間スプシ「管理表」の行2(販売数)に書き込み。
 * Finance/Settlement の FBA輸送手数料を行6に書き込み。
 */
function syncToIntermediateFromAmazon() {
  Logger.log('--- Amazon → 中間スプシ ---');

  const intSheetId = getIntermediateSheetId();
  if (!intSheetId) {
    Logger.log('⚠️ INTERMEDIATE_SHEET_ID が未設定。スキップ');
    return;
  }

  const intSs = SpreadsheetApp.openById(intSheetId);
  const mgmtSheet = intSs.getSheetByName(INTERMEDIATE_MGMT_SHEET);
  if (!mgmtSheet) return;

  const yearLabel = String(mgmtSheet.getRange(1, 2).getValue() || '').trim();
  const yearMatch = yearLabel.match(/^(\d{4})/);
  if (!yearMatch) return;
  const year = parseInt(yearMatch[1]);

  // 販売数を D1 から月次集計
  const salesByMonth = aggregateMonthlyUnitsFromD1(year);

  // FBA輸送手数料を Finance/Settlement から月次集計
  const fbaByMonth = aggregateFbaInboundShippingByMonth(year);

  // 中間スプシに書き込み
  const salesRow = new Array(12).fill(0);
  const fbaRow = new Array(12).fill(0);
  for (let m = 0; m < 12; m++) {
    const ym = year + '-' + String(m + 1).padStart(2, '0');
    salesRow[m] = salesByMonth[ym] || 0;
    fbaRow[m] = fbaByMonth[ym] || 0;
  }

  mgmtSheet.getRange(INTERMEDIATE_ROWS.SALES, INTERMEDIATE_MONTH_COL_START, 1, 12).setValues([salesRow]);
  mgmtSheet.getRange(INTERMEDIATE_ROWS.FBA_SHIPPING, INTERMEDIATE_MONTH_COL_START, 1, 12).setValues([fbaRow]);

  Logger.log('✅ 中間スプシ更新: 販売数 + FBA輸送手数料（12ヶ月）');
}

/**
 * D1 日次データから月次販売数（注文点数合計）を集計
 * @param {number} year 対象年
 * @returns {Object} { 'YYYY-MM': units }
 */
function aggregateMonthlyUnitsFromD1(year) {
  const sheet = getOrCreateSheet(SHEET_NAMES.D1_DAILY);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return {};

  // 列1: 日付, 列7: 注文点数
  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const byMonth = {};

  for (const row of data) {
    const rawDate = row[0];
    let ym = '';
    if (rawDate instanceof Date) {
      if (rawDate.getFullYear() !== year) continue;
      ym = year + '-' + String(rawDate.getMonth() + 1).padStart(2, '0');
    } else if (rawDate) {
      const s = String(rawDate);
      if (s.substring(0, 4) !== String(year)) continue;
      ym = s.substring(0, 7);
    } else {
      continue;
    }

    byMonth[ym] = (byMonth[ym] || 0) + (parseFloat(row[6]) || 0);
  }
  return byMonth;
}

/**
 * Finance + Settlement から FBA Inbound Transportation Fee を月次集計
 * @param {number} year 対象年
 * @returns {Object} { 'YYYY-MM': totalFee（正の数）}
 */
function aggregateFbaInboundShippingByMonth(year) {
  const result = {};

  // Settlement (D2) から
  const d2Sheet = getOrCreateSheet(SHEET_NAMES.D2_SETTLEMENT);
  const d2LastRow = d2Sheet.getLastRow();
  if (d2LastRow > 1) {
    // 列3: 日付, 列6: 明細種別, 列7: 金額
    const data = d2Sheet.getRange(2, 3, d2LastRow - 1, 5).getValues();
    for (const row of data) {
      const rawDate = row[0];
      const itemType = String(row[3] || '');
      const amount = parseFloat(row[4]) || 0;

      // FBA Inbound Transportation Fee 系の明細を抽出
      // 候補: FBAInboundTransportationFee / FBA Inbound Transportation Fee /
      //       INBOUND / TransportationFee 等
      if (!isFbaInboundFeeType(itemType)) continue;

      let ym = '';
      if (rawDate instanceof Date) {
        if (rawDate.getFullYear() !== year) continue;
        ym = year + '-' + String(rawDate.getMonth() + 1).padStart(2, '0');
      } else {
        const s = String(rawDate);
        if (s.substring(0, 4) !== String(year)) continue;
        ym = s.substring(0, 7);
      }
      result[ym] = (result[ym] || 0) + Math.abs(amount);
    }
  }

  // Finance Events (D2F) から（重複しないものだけ）
  const d2fSheet = getOrCreateSheet(SHEET_NAMES.D2F_FINANCE_EVENTS);
  const d2fLastRow = d2fSheet.getLastRow();
  if (d2fLastRow > 1) {
    // 列1: 日付, 列2: イベント種別, 列3: FeeReason, 列4: 金額, 列5: ステータス(暫定/確定)
    const data = d2fSheet.getRange(2, 1, d2fLastRow - 1, 5).getValues();
    for (const row of data) {
      const rawDate = row[0];
      const feeReason = String(row[2] || '');
      const amount = parseFloat(row[3]) || 0;
      const status = String(row[4] || '');

      // 確定済み（Settlement で取り込んだ）はスキップ
      if (status === '確定') continue;
      if (!isFbaInboundFeeType(feeReason)) continue;

      let ym = '';
      if (rawDate instanceof Date) {
        if (rawDate.getFullYear() !== year) continue;
        ym = year + '-' + String(rawDate.getMonth() + 1).padStart(2, '0');
      } else {
        const s = String(rawDate);
        if (s.substring(0, 4) !== String(year)) continue;
        ym = s.substring(0, 7);
      }
      result[ym] = (result[ym] || 0) + Math.abs(amount);
    }
  }

  return result;
}

/**
 * 明細種別 / FeeReason 文字列が FBA Inbound Transportation Fee 系か判定
 */
function isFbaInboundFeeType(typeStr) {
  if (!typeStr) return false;
  const lower = String(typeStr).toLowerCase();
  return lower.includes('fbainboundtransportation') ||
         lower.includes('fba inbound transportation') ||
         lower.includes('inbound transportation fee') ||
         lower.includes('納品時の輸送手数料');
}

/**
 * 単体テスト: 中間スプシから読み取りだけ実行（書き込みなし）
 */
function testReadIntermediate() {
  const intSheetId = getIntermediateSheetId();
  if (!intSheetId) {
    Logger.log('❌ INTERMEDIATE_SHEET_ID 未設定');
    return;
  }
  const intSs = SpreadsheetApp.openById(intSheetId);
  const mgmtSheet = intSs.getSheetByName(INTERMEDIATE_MGMT_SHEET);
  if (!mgmtSheet) {
    Logger.log('❌ 「' + INTERMEDIATE_MGMT_SHEET + '」シートが見つかりません');
    Logger.log('シート一覧: ' + intSs.getSheets().map(s => s.getName()).join(', '));
    return;
  }
  const yearLabel = String(mgmtSheet.getRange(1, 2).getValue() || '').trim();
  Logger.log('B1 (年ラベル): "' + yearLabel + '"');

  const payRewards = mgmtSheet.getRange(INTERMEDIATE_ROWS.PAY_REWARD, INTERMEDIATE_MONTH_COL_START, 1, 12).getValues()[0];
  const yamatoCosts = mgmtSheet.getRange(INTERMEDIATE_ROWS.YAMATO, INTERMEDIATE_MONTH_COL_START, 1, 12).getValues()[0];
  Logger.log('支払報酬（行' + INTERMEDIATE_ROWS.PAY_REWARD + '）: ' + JSON.stringify(payRewards));
  Logger.log('ヤマト運輸（行' + INTERMEDIATE_ROWS.YAMATO + '）: ' + JSON.stringify(yamatoCosts));
}

/**
 * 診断: D2 経費明細の FBA/輸送関連 明細種別を集計
 *
 * 荷造運賃の数値検証用。isFbaInboundFeeType() が何を拾っているか確認する。
 */
function debugSettlementFbaItems() {
  const sheet = getOrCreateSheet(SHEET_NAMES.D2_SETTLEMENT);
  if (sheet.getLastRow() <= 1) {
    Logger.log('D2 経費明細にデータがありません');
    return;
  }

  // 列3: 日付, 列4: ASIN, 列5: トランザクション種別, 列6: 明細種別, 列7: 金額
  const data = sheet.getRange(2, 3, sheet.getLastRow() - 1, 5).getValues();

  const typesAll = {};       // FBA/輸送/納品 関連すべて
  const typesHit = {};       // isFbaInboundFeeType がヒットするもの
  const byMonthHit = {};     // ヒットするものの月別集計

  for (const row of data) {
    const rawDate = row[0];
    const itemType = String(row[3] || '');
    const amount = parseFloat(row[4]) || 0;
    const lower = itemType.toLowerCase();

    const isRelated = lower.includes('inbound') || lower.includes('fba') ||
                      lower.includes('transport') || itemType.includes('輸送') ||
                      itemType.includes('納品');
    if (!isRelated) continue;

    typesAll[itemType] = (typesAll[itemType] || 0) + Math.abs(amount);

    if (isFbaInboundFeeType(itemType)) {
      typesHit[itemType] = (typesHit[itemType] || 0) + Math.abs(amount);

      // 月別
      let ym = '';
      if (rawDate instanceof Date) {
        ym = rawDate.getFullYear() + '-' + String(rawDate.getMonth() + 1).padStart(2, '0');
      } else {
        ym = String(rawDate).substring(0, 7);
      }
      if (ym) byMonthHit[ym] = (byMonthHit[ym] || 0) + Math.abs(amount);
    }
  }

  Logger.log('═══════════════════════════════════════════');
  Logger.log('【A】 FBA/輸送/納品 関連すべての明細種別（金額降順）:');
  Logger.log('═══════════════════════════════════════════');
  Object.entries(typesAll)
    .sort((a, b) => b[1] - a[1])
    .forEach(([t, v]) => Logger.log('  "' + t + '": ¥' + v.toLocaleString()));

  Logger.log('');
  Logger.log('═══════════════════════════════════════════');
  Logger.log('【B】 isFbaInboundFeeType() で現在ヒットするもの:');
  Logger.log('═══════════════════════════════════════════');
  Object.entries(typesHit)
    .sort((a, b) => b[1] - a[1])
    .forEach(([t, v]) => Logger.log('  "' + t + '": ¥' + v.toLocaleString()));

  Logger.log('');
  Logger.log('═══════════════════════════════════════════');
  Logger.log('【C】 ヒットしたものの月別集計:');
  Logger.log('═══════════════════════════════════════════');
  Object.entries(byMonthHit)
    .sort()
    .forEach(([ym, v]) => Logger.log('  ' + ym + ': ¥' + v.toLocaleString()));
}
