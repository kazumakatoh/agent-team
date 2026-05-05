/**
 * Amazon Dashboard - D1 日次データ バックフィルモジュール（v2: M1ソース）
 *
 * 後追いで判明する単価・手数料・返品をD1各行へ反映する一連の関数:
 *
 *   backfillD1CogsFromM1()           M1 仕入単価 → D1 仕入単価/仕入原価合計
 *   backfillD1RefundsFromSettlement() D2 → D1 返品数/返品額（月別Settlement率推定）
 *   backfillD1FbaFeeFromSettlement()  D2 → D1 FBA手数料（月別Settlement率推定）
 *
 * 仕入単価は M1 商品マスター（手入力）が真のソース。
 * 過去は M2 月次仕入単価から取っていたが廃止し、最新値で全期間バックフィルする。
 */

// ===== 仕入原価 + 販売手数料（M1 → D1）=====

/**
 * D1 仕入原価 + 販売手数料 を M1 商品マスターから一括バックフィル
 *
 * - 仕入単価(列20) / 仕入原価合計(列21) ← M1.仕入単価 × 注文点数
 * - 販売手数料(列23)                    ← M1.販売手数料率 × 売上
 */
function backfillD1FromM1() {
  // 列23（販売手数料）ヘッダー未設定の場合は先に整える
  setupDailyDataHeaders();
  backfillD1CogsFromM1();
  backfillD1CommissionFromM1();
}

/**
 * D1 日次データの 仕入単価（列20）/ 仕入原価合計（列21）を M1 商品マスターから反映
 *
 * マッチロジック:
 *   1. ASIN で M1 商品マスターから 仕入単価 を取得
 *   2. 単価が無ければ 0
 *   3. 仕入原価合計 = 注文点数 × 仕入単価
 *
 * 全D1行を一括上書きする（M1 が真のソースのため、過去含め最新値で更新）。
 */
function backfillD1CogsFromM1() {
  Logger.log('===== D1 仕入原価バックフィル（M1ソース） 開始 =====');

  const masterMap = getProductMasterMap();
  const priceMap = {};
  for (const [asin, m] of Object.entries(masterMap)) {
    if (m.purchasePrice > 0) priceMap[asin] = m.purchasePrice;
  }
  Logger.log('  M1 仕入単価設定済みASIN数: ' + Object.keys(priceMap).length);

  const d1Sheet = getOrCreateSheet(SHEET_NAMES.D1_DAILY);
  const d1Last = d1Sheet.getLastRow();
  if (d1Last <= 1) { Logger.log('D1 空'); return 0; }

  // 列1:日付, 列2:ASIN, 列7:点数, 列20:仕入単価, 列21:仕入原価合計
  const d1Data = d1Sheet.getRange(2, 1, d1Last - 1, 21).getValues();
  const updates = [];
  let matched = 0, unmatched = 0;

  for (const row of d1Data) {
    const asin = String(row[1] || '').trim();
    const units = parseFloat(row[6]) || 0;
    const price = priceMap[asin] || 0;
    const cogs = price * units;
    updates.push([price || '', cogs || '']);
    if (price > 0) matched++;
    else if (asin) unmatched++;
  }

  Logger.log('  マッチ: ' + matched + ' / 単価未設定: ' + unmatched);

  if (updates.length > 0) {
    d1Sheet.getRange(2, 20, updates.length, 2).setValues(updates);
  }
  Logger.log('✅ D1 仕入原価バックフィル: ' + updates.length + ' 行更新');
  return updates.length;
}

/**
 * D1 日次データの 販売手数料（列23）を M1 商品マスターから反映
 *
 * 販売手数料 = D1.売上 × M1.販売手数料率（カテゴリ別レート）
 *
 * 全D1行を一括上書きする。
 */
function backfillD1CommissionFromM1() {
  Logger.log('===== D1 販売手数料バックフィル（M1ソース） 開始 =====');

  const masterMap = getProductMasterMap();
  const rateMap = {};
  for (const [asin, m] of Object.entries(masterMap)) {
    if (m.commissionRate > 0) rateMap[asin] = m.commissionRate;
  }
  Logger.log('  M1 販売手数料率設定済みASIN数: ' + Object.keys(rateMap).length);

  const d1Sheet = getOrCreateSheet(SHEET_NAMES.D1_DAILY);
  const d1Last = d1Sheet.getLastRow();
  if (d1Last <= 1) { Logger.log('D1 空'); return 0; }

  // 列1:日付, 列2:ASIN, 列5:売上, 列23:販売手数料
  const d1Data = d1Sheet.getRange(2, 1, d1Last - 1, 5).getValues();
  const updates = [];
  let matched = 0, unmatched = 0;

  for (const row of d1Data) {
    const asin = String(row[1] || '').trim();
    const sales = parseFloat(row[4]) || 0;
    const rate = rateMap[asin] || 0;
    const commission = Math.round(sales * rate);
    updates.push([commission || '']);
    if (rate > 0) matched++;
    else if (asin) unmatched++;
  }

  Logger.log('  マッチ: ' + matched + ' / 手数料率未設定: ' + unmatched);

  if (updates.length > 0) {
    d1Sheet.getRange(2, 23, updates.length, 1).setValues(updates);
  }
  Logger.log('✅ D1 販売手数料バックフィル: ' + updates.length + ' 行更新');
  return updates.length;
}

// ===== 返品（D2 → D1）=====

/**
 * D1 日次データの 返品数（列14）・返品額（列15）を Settlement rate で推定反映
 *
 * D2 の Refund トランザクションから月別返品率を算出し、日次売上に掛ける。
 */
function backfillD1RefundsFromSettlement() {
  Logger.log('===== D1 返品数・返品額バックフィル 開始 =====');

  const d2Sheet = getOrCreateSheet(SHEET_NAMES.D2_SETTLEMENT);
  const d2LastRow = d2Sheet.getLastRow();
  if (d2LastRow <= 1) { Logger.log('D2 空'); return; }

  const d2Data = d2Sheet.getRange(2, 3, d2LastRow - 1, 6).getValues();
  // 列3:日付, 列4:ASIN, 列5:トランザクション種別, 列6:明細種別, 列7:金額, 列8:数量
  const byMonth = {}; // ym → { refundAmount, refundQty, orderQty, principal }

  for (const row of d2Data) {
    const rawDate = row[0];
    const txType = String(row[2] || '').trim();
    const itemType = String(row[3] || '').trim();
    const amount = parseFloat(row[4]) || 0;
    const qty = parseInt(row[5]) || 0;

    let ym = '';
    if (rawDate instanceof Date) {
      ym = rawDate.getFullYear() + '-' + String(rawDate.getMonth() + 1).padStart(2, '0');
    } else if (rawDate) {
      ym = String(rawDate).substring(0, 7);
    }
    if (!ym) continue;

    if (!byMonth[ym]) byMonth[ym] = { refundAmount: 0, refundQty: 0, orderQty: 0, principal: 0 };

    if (txType === 'Refund' && itemType === 'Principal') {
      byMonth[ym].refundAmount += Math.abs(amount);
      byMonth[ym].refundQty += Math.abs(qty);
    } else if (txType === 'Order' && itemType === 'Principal') {
      byMonth[ym].orderQty += qty;
      byMonth[ym].principal += amount;
    }
  }

  // Settlement Report の Refund 行に qty が入っていない場合があるため、
  // qty率は「金額率」で代替する
  const rateByMonth = {};
  let totalRefundAmount = 0, totalRefundQty = 0, totalOrderQty = 0, totalPrincipal = 0;
  Object.keys(byMonth).forEach(ym => {
    const m = byMonth[ym];
    const amountRate = m.principal > 0 ? m.refundAmount / m.principal : 0;
    const qtyRate = m.orderQty > 0 && m.refundQty > 0 ? m.refundQty / m.orderQty : amountRate;
    rateByMonth[ym] = { refundAmountRate: amountRate, refundQtyRate: qtyRate };
    totalRefundAmount += m.refundAmount;
    totalRefundQty += m.refundQty;
    totalOrderQty += m.orderQty;
    totalPrincipal += m.principal;
  });
  const fallbackAmountRate = totalPrincipal > 0 ? totalRefundAmount / totalPrincipal : 0;
  const fallbackQtyRate = totalOrderQty > 0 && totalRefundQty > 0
    ? totalRefundQty / totalOrderQty
    : fallbackAmountRate;
  Logger.log('  月別rate: ' + Object.keys(rateByMonth).length + ' 月 / フォールバック率: qty=' +
    (fallbackQtyRate * 100).toFixed(2) + '% / amount=' + (fallbackAmountRate * 100).toFixed(2) + '%');

  // D1 を走査（列1:日付, 列5:売上, 列7:注文点数）
  const d1Sheet = getOrCreateSheet(SHEET_NAMES.D1_DAILY);
  const d1LastRow = d1Sheet.getLastRow();
  if (d1LastRow <= 1) return;

  const d1Data = d1Sheet.getRange(2, 1, d1LastRow - 1, 7).getValues();
  const updates = [];
  let applied = 0;

  for (const row of d1Data) {
    const ym = formatYearMonth(row[0]);
    const sales = parseFloat(row[4]) || 0;
    const units = parseFloat(row[6]) || 0;
    const rate = rateByMonth[ym];
    const qtyRate = rate ? rate.refundQtyRate : fallbackQtyRate;
    const amtRate = rate ? rate.refundAmountRate : fallbackAmountRate;

    const refundQty = Math.round(units * qtyRate);
    const refundAmount = Math.round(sales * amtRate);
    updates.push([refundQty || '', refundAmount || '']);
    if (refundQty > 0 || refundAmount > 0) applied++;
  }

  d1Sheet.getRange(2, 14, updates.length, 2).setValues(updates);
  Logger.log('✅ D1 返品バックフィル: ' + applied + ' / ' + updates.length + ' 行');
}

// ===== FBA手数料（D2 → D1）=====

/**
 * D1 日次データの FBA手数料列（列13）を Settlement のFBA率 × 日次売上 で推定反映
 *
 * D2 の明細種別 "FBAPerUnitFulfillmentFee" を月次集計して Principal に対する率を算出。
 */
function backfillD1FbaFeeFromSettlement() {
  Logger.log('===== D1 FBA手数料バックフィル 開始 =====');

  const d2Sheet = getOrCreateSheet(SHEET_NAMES.D2_SETTLEMENT);
  const d2LastRow = d2Sheet.getLastRow();
  if (d2LastRow <= 1) { Logger.log('D2 空'); return; }

  const d2Data = d2Sheet.getRange(2, 3, d2LastRow - 1, 5).getValues();
  const byMonth = {}; // ym → { fba, principal }
  for (const row of d2Data) {
    const rawDate = row[0];
    const itemType = String(row[3] || '').trim();
    const amount = parseFloat(row[4]) || 0;

    let ym = '';
    if (rawDate instanceof Date) {
      ym = rawDate.getFullYear() + '-' + String(rawDate.getMonth() + 1).padStart(2, '0');
    } else if (rawDate) {
      ym = String(rawDate).substring(0, 7);
    }
    if (!ym) continue;

    if (!byMonth[ym]) byMonth[ym] = { fba: 0, principal: 0 };
    if (itemType === 'FBAPerUnitFulfillmentFee') byMonth[ym].fba += Math.abs(amount);
    else if (itemType === 'Principal') byMonth[ym].principal += amount;
  }

  const rateByMonth = {};
  let totalFba = 0, totalPrincipalFba = 0;
  Object.keys(byMonth).forEach(ym => {
    const m = byMonth[ym];
    rateByMonth[ym] = m.principal > 0 ? m.fba / m.principal : 0;
    totalFba += m.fba;
    totalPrincipalFba += m.principal;
  });
  const fallbackFbaRate = totalPrincipalFba > 0 ? totalFba / totalPrincipalFba : 0;

  const d1Sheet = getOrCreateSheet(SHEET_NAMES.D1_DAILY);
  const d1LastRow = d1Sheet.getLastRow();
  if (d1LastRow <= 1) return;

  const d1Data = d1Sheet.getRange(2, 1, d1LastRow - 1, 5).getValues();
  const fbaFees = [];
  let applied = 0;
  for (const row of d1Data) {
    const ym = formatYearMonth(row[0]);
    const sales = parseFloat(row[4]) || 0;
    const rate = rateByMonth[ym] !== undefined ? rateByMonth[ym] : fallbackFbaRate;
    const fba = Math.round(sales * rate);
    fbaFees.push([fba || '']);
    if (fba > 0) applied++;
  }

  d1Sheet.getRange(2, 13, fbaFees.length, 1).setValues(fbaFees);
  Logger.log('✅ D1 FBA手数料バックフィル: ' + applied + ' / ' + fbaFees.length + ' 行');
}

// ===== 診断 =====

/**
 * 診断: D1 で売上ありながら cogs=0 の ASIN を上位表示
 *
 * 売上のうち原価未計上分の規模を把握し、M1 商品マスターへの追加入力が必要な
 * ASIN を特定する。
 */
function debugD1CogsGap() {
  const sheet = getOrCreateSheet(SHEET_NAMES.D1_DAILY);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) { Logger.log('D1 空'); return; }

  // 列1:日付, 列2:ASIN, 列3:商品名, 列5:売上, 列21:仕入原価合計
  const data = sheet.getRange(2, 1, lastRow - 1, 21).getValues();
  const byAsin = {};

  for (const row of data) {
    const asin = String(row[1] || '').trim();
    if (!asin) continue;
    const sales = parseFloat(row[4]) || 0;
    const cogs = parseFloat(row[20]) || 0;
    if (!byAsin[asin]) byAsin[asin] = { name: row[2] || '', sales: 0, cogs: 0, salesNoCogs: 0 };
    byAsin[asin].sales += sales;
    byAsin[asin].cogs += cogs;
    if (cogs === 0) byAsin[asin].salesNoCogs += sales;
  }

  const noCogsList = Object.entries(byAsin)
    .map(([asin, v]) => Object.assign({ asin }, v))
    .filter(v => v.salesNoCogs > 0)
    .sort((a, b) => b.salesNoCogs - a.salesNoCogs);

  Logger.log('===== D1 cogs ギャップ診断 =====');
  Logger.log('cogs=0 で売上計上ある ASIN: ' + noCogsList.length + ' 件 / 上位20');
  noCogsList.slice(0, 20).forEach(v => {
    Logger.log('  ' + v.asin + ' | 売上: ¥' + v.salesNoCogs.toLocaleString() + ' | ' + (v.name || '(名前なし)'));
  });
  if (noCogsList.length > 20) Logger.log('  ... 他 ' + (noCogsList.length - 20) + ' 件');
}

/**
 * 診断: D1 日次データの仕入原価合計を年月別に集計表示
 */
function debugD1CogsByMonth() {
  const sheet = getOrCreateSheet(SHEET_NAMES.D1_DAILY);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) { Logger.log('D1 空'); return; }

  // 列1: 日付, 列21: 仕入原価合計
  const data = sheet.getRange(2, 1, lastRow - 1, 21).getValues();
  const stats = {};

  for (const row of data) {
    const ym = formatYearMonth(row[0]);
    if (!ym) continue;
    if (!stats[ym]) stats[ym] = { rowCount: 0, cogsSum: 0, withCogsCount: 0 };

    stats[ym].rowCount++;
    const cogs = parseFloat(row[20]) || 0;
    stats[ym].cogsSum += cogs;
    if (cogs > 0) stats[ym].withCogsCount++;
  }

  Logger.log('===== D1 年月別 cogs 集計 =====');
  Logger.log('年月      | D1行数 | cogs有 | cogs合計');
  Object.keys(stats).sort().reverse().slice(0, 24).forEach(ym => {
    const s = stats[ym];
    Logger.log('  ' + ym + ' | ' + String(s.rowCount).padStart(5) + ' | ' +
               String(s.withCogsCount).padStart(5) + ' | ¥' + s.cogsSum.toLocaleString());
  });
}
