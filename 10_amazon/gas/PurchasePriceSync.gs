/**
 * Amazon Dashboard - 仕入単価 CF連携（M2 月次仕入単価）
 *
 * キャッシュフロー管理シートから月別仕入単価を自動読み取り、
 * M2 「月次仕入単価」テーブルに書き込む。
 * さらに D1 日次データの「仕入単価」「仕入原価合計」列をバックフィルする。
 *
 * ## CFシート想定構造
 *
 * 上部ヘッダー部（ProductMaster.gs が参照）:
 *   ASIN行 / 商品名行
 *   E列(index=4) 以降が各商品の列
 *
 * 月別ブロック（複数）:
 *   各月ブロックに「仕入単価」ラベルがある行。
 *   月はA〜C列のいずれかに「YYYY年M月」「YYYY.M」等の表記。
 *
 * ## 処理フロー
 *
 * 1. CFシート全行読み込み
 * 2. ASIN行・商品名行を特定（ProductMaster.gs と同ロジック）
 * 3. 「仕入単価」ラベル行を全て抽出
 * 4. 各行の直上/周辺を見て月を特定
 * 5. ASIN × 年月 × 仕入単価 のテーブルを M2 に書き込み
 * 6. D1 各行に仕入単価を紐付けて 仕入単価(列20) / 仕入原価合計(列21) を埋める
 *
 * ## トリガー: 毎月3日 AM6:00 (syncPurchasePriceFromCfSheet)
 */

const M2_HEADERS = ['ASIN', '年月', '仕入単価', '在庫数'];

/**
 * メイン: CFシート → M2 同期 + D1 バックフィル
 */
function syncPurchasePriceFromCfSheet() {
  const t0 = Date.now();
  Logger.log('===== CF → M2 仕入単価同期 開始 =====');

  const cfSheetId = getCredential('CF_SHEET_ID');
  const cfSs = SpreadsheetApp.openById(cfSheetId);

  // CFシート特定（ProductMaster.gs と同じ）
  const sheets = cfSs.getSheets();
  let cfSheet = null;
  for (const s of sheets) {
    if (s.getSheetId() === 347226196) {
      cfSheet = s;
      break;
    }
  }
  if (!cfSheet) cfSheet = sheets[0];
  Logger.log('CFシート: ' + cfSheet.getName());

  // 全データを一括読み込み
  const lastRow = cfSheet.getLastRow();
  const lastCol = cfSheet.getLastColumn();
  const allData = cfSheet.getRange(1, 1, lastRow, lastCol).getValues();

  // ASIN行・商品名行を特定（上部10行以内）
  const { asinRow, nameRow } = findHeaderRows(allData, Math.min(10, lastRow));

  if (asinRow < 0 && nameRow < 0) {
    Logger.log('❌ ASIN行も商品名行も見つかりません');
    return;
  }
  Logger.log('ASIN行: ' + (asinRow + 1) + ' / 商品名行: ' + (nameRow + 1));

  // 商品列リストを作成（E列以降）
  const products = extractProductColumns(allData, asinRow, nameRow, lastCol);
  Logger.log('商品列数: ' + products.length);

  // 「仕入単価」ラベル行を全て抽出
  const priceRows = findPriceRows(allData);
  Logger.log('仕入単価ラベル行: ' + priceRows.length + ' 件');

  // 各 仕入単価行に対して月を特定
  const priceEntries = [];
  for (const rowIdx of priceRows) {
    const ym = detectMonthForRow(allData, rowIdx);
    if (!ym) {
      Logger.log('  行' + (rowIdx + 1) + ': 月特定不可、スキップ');
      continue;
    }

    // その行の各商品列から単価を取得
    for (const p of products) {
      const rawVal = allData[rowIdx][p.col];
      const price = parseFloat(rawVal);
      if (!isFinite(price) || price <= 0) continue;

      priceEntries.push({
        asin: p.asin,
        name: p.name,
        yearMonth: ym,
        price: price,
      });
    }
  }

  Logger.log('有効な仕入単価エントリ: ' + priceEntries.length + ' 件');

  // M2 に書き込み（既存データはクリアして再生成）
  writeM2Sheet(priceEntries);

  // D1 をバックフィル
  const updated = backfillD1CogsFromM2();
  Logger.log('✅ D1 cogs バックフィル: ' + updated + ' 行更新');

  Logger.log('===== 同期完了（' + (Date.now() - t0) + 'ms）=====');
}

/**
 * ヘッダー領域から ASIN行 / 商品名行 を検出
 */
function findHeaderRows(data, maxRows) {
  let asinRow = -1;
  let nameRow = -1;

  for (let r = 0; r < maxRows; r++) {
    // ASIN ラベル or パターン検索
    for (let c = 0; c < Math.min(10, data[r].length); c++) {
      const v = String(data[r][c] || '').trim();
      if (v.toLowerCase() === 'asin' && asinRow === -1) asinRow = r;
      if (v.match(/^B0[A-Z0-9]{8,}$/) && asinRow === -1) asinRow = r;
    }

    // 商品名ラベル
    const firstThree = [0, 1, 2].map(c => String(data[r][c] || '').trim().toLowerCase());
    if (firstThree.some(v => v === '商品名') && nameRow === -1) nameRow = r;
  }

  if (asinRow >= 0 && nameRow === -1) nameRow = asinRow + 1;
  return { asinRow, nameRow };
}

/**
 * 商品列リストを抽出（E列以降で商品名がある列のみ）
 */
function extractProductColumns(data, asinRow, nameRow, lastCol) {
  const products = [];
  const startCol = 4; // E列

  for (let c = startCol; c < lastCol; c++) {
    const asin = asinRow >= 0 ? String(data[asinRow][c] || '').trim() : '';
    const name = nameRow >= 0 ? String(data[nameRow][c] || '').trim() : '';

    // 空・合計・残高系はスキップ
    if (!name || name === '合計' || name === '残高' || name.length < 2) continue;
    if (name.match(/^\d/)) continue; // 数字始まりはスキップ

    products.push({ col: c, asin, name });
  }
  return products;
}

/**
 * 「仕入単価」ラベルを含む行番号のリストを返す（0-indexed）
 */
function findPriceRows(data) {
  const rows = [];
  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < Math.min(4, data[r].length); c++) {
      const v = String(data[r][c] || '').trim();
      if (v.includes('仕入単価')) {
        rows.push(r);
        break;
      }
    }
  }
  return rows;
}

/**
 * 仕入単価行の周辺（上方向最大20行）から月を特定
 * 「YYYY年M月」「YYYY/M」「YYYY.M」等のパターンを探す
 */
function detectMonthForRow(data, rowIdx) {
  for (let r = rowIdx; r >= Math.max(0, rowIdx - 20); r--) {
    for (let c = 0; c < Math.min(5, data[r].length); c++) {
      const v = data[r][c];
      if (!v) continue;

      // Date オブジェクト
      if (v instanceof Date) {
        return v.getFullYear() + '-' + String(v.getMonth() + 1).padStart(2, '0');
      }

      const s = String(v).trim();

      // "YYYY年M月"
      let m = s.match(/(\d{4})\s*年\s*(\d{1,2})\s*月/);
      if (m) return m[1] + '-' + String(parseInt(m[2])).padStart(2, '0');

      // "YYYY/M" or "YYYY-M" or "YYYY.M"
      m = s.match(/^(\d{4})[\/\-.](\d{1,2})(?:\D|$)/);
      if (m) return m[1] + '-' + String(parseInt(m[2])).padStart(2, '0');

      // "YYYYMM"
      m = s.match(/^(\d{4})(\d{2})$/);
      if (m) {
        const mo = parseInt(m[2]);
        if (mo >= 1 && mo <= 12) return m[1] + '-' + m[2];
      }
    }
  }
  return null;
}

/**
 * M2 月次仕入単価テーブルを書き込み（完全再生成）
 */
function writeM2Sheet(entries) {
  const sheet = getOrCreateSheet(SHEET_NAMES.M2_PURCHASE_PRICE);
  sheet.clear();

  sheet.getRange(1, 1, 1, M2_HEADERS.length).setValues([M2_HEADERS])
    .setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  if (entries.length === 0) return;

  // ASIN × 年月 で一意化（複数ある場合は最後勝ち）
  const uniqueMap = {};
  for (const e of entries) {
    if (!e.asin) continue; // ASIN空はスキップ
    uniqueMap[e.asin + '_' + e.yearMonth] = e;
  }
  const rows = Object.values(uniqueMap).map(e => [e.asin, e.yearMonth, e.price, '']);

  // ASIN 昇順 → 年月昇順でソート
  rows.sort((a, b) => {
    if (a[0] !== b[0]) return a[0].localeCompare(b[0]);
    return a[1].localeCompare(b[1]);
  });

  // 年月列はテキスト形式に固定（自動Date変換を防ぐ）
  sheet.getRange(2, 2, rows.length, 1).setNumberFormat('@');

  sheet.getRange(2, 1, rows.length, M2_HEADERS.length).setValues(rows);
  sheet.getRange(2, 3, rows.length, 1).setNumberFormat('#,##0');
  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 110);
  Logger.log('✅ M2: ' + rows.length + ' 行書込み');
}

/**
 * D1 日次データに仕入単価・仕入原価合計を反映（バックフィル）
 *
 * マッチロジック:
 * 1. ASIN × 年月 で M2 から厳密一致する単価を探す
 * 2. 無ければ ASIN の**直近過去月**の単価を使う（価格はカテゴリ削除後も継続）
 * 3. それも無ければ 0
 *
 * @returns {number} 更新行数
 */
function backfillD1CogsFromM2() {
  // M2 を (ASIN_YM → price) マップと、ASIN別の月リストに
  const m2Sheet = getOrCreateSheet(SHEET_NAMES.M2_PURCHASE_PRICE);
  const m2Last = m2Sheet.getLastRow();
  if (m2Last <= 1) {
    Logger.log('M2 が空です、バックフィルスキップ');
    return 0;
  }
  const m2Data = m2Sheet.getRange(2, 1, m2Last - 1, 3).getValues();
  const priceMap = {};            // asin_ym → price（厳密一致）
  const monthsByAsin = {};         // asin → [ym昇順リスト]
  for (const row of m2Data) {
    const asin = String(row[0] || '').trim();
    const ym = formatYearMonth(row[1]);
    const price = parseFloat(row[2]) || 0;
    if (asin && ym && price > 0) {
      priceMap[asin + '_' + ym] = price;
      if (!monthsByAsin[asin]) monthsByAsin[asin] = [];
      monthsByAsin[asin].push(ym);
    }
  }
  Object.keys(monthsByAsin).forEach(asin => monthsByAsin[asin].sort());

  Logger.log('  priceMap 件数: ' + Object.keys(priceMap).length);
  Logger.log('  ASIN数: ' + Object.keys(monthsByAsin).length);

  // D1 を読み込み（列1: 日付, 列2: ASIN, 列7: 点数, 列20: 仕入単価, 列21: 仕入原価合計）
  const d1Sheet = getOrCreateSheet(SHEET_NAMES.D1_DAILY);
  const d1Last = d1Sheet.getLastRow();
  if (d1Last <= 1) return 0;

  const d1Data = d1Sheet.getRange(2, 1, d1Last - 1, 21).getValues();
  const updates = [];
  let exactMatch = 0, fallbackMatch = 0, noMatch = 0;

  for (const row of d1Data) {
    const rawDate = row[0];
    const asin = String(row[1] || '').trim();
    const units = parseFloat(row[6]) || 0;

    const ym = formatYearMonth(rawDate);
    let price = priceMap[asin + '_' + ym] || 0;

    if (price > 0) {
      exactMatch++;
    } else if (monthsByAsin[asin]) {
      // フォールバック: 該当月以前の最新単価（なければ最古単価）
      const months = monthsByAsin[asin];
      let bestYm = null;
      for (const m of months) {
        if (m <= ym) bestYm = m;  // 過去方向で最新
      }
      if (!bestYm) bestYm = months[0];  // 過去になければ最古
      price = priceMap[asin + '_' + bestYm] || 0;
      if (price > 0) fallbackMatch++;
    } else {
      noMatch++;
    }

    const cogs = price * units;
    updates.push([price, cogs]);
  }

  Logger.log('  マッチ詳細: 厳密=' + exactMatch + ' / フォールバック=' + fallbackMatch + ' / なし=' + noMatch);

  if (updates.length > 0) {
    d1Sheet.getRange(2, 20, updates.length, 2).setValues(updates);
  }
  return updates.length;
}

/**
 * D1 日次データの 返品数（列14）・返品額（列15）を Settlement rate で推定反映
 *
 * D2 の Refund トランザクションから月別返品率を算出し、日次売上に掛ける。
 */
function backfillD1RefundsFromSettlement() {
  Logger.log('===== D1 返品数・返品額バックフィル 開始 =====');

  // D2 から月別 返品率を算出
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

  // 率を算出
  // 注: Settlement Report の Refund 行に qty が入っていない場合があるため、
  //     qty率は「金額率」で代替する（refund_amount/principal を qty にも適用）
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
  // 全期間平均率
  const fallbackAmountRate = totalPrincipal > 0 ? totalRefundAmount / totalPrincipal : 0;
  const fallbackQtyRate = totalOrderQty > 0 && totalRefundQty > 0
    ? totalRefundQty / totalOrderQty
    : fallbackAmountRate;  // qtyが取れないなら金額率で代替
  Logger.log('  月別rate: ' + Object.keys(rateByMonth).length + ' 月 / フォールバック率: qty=' +
    (fallbackQtyRate * 100).toFixed(2) + '% / amount=' + (fallbackAmountRate * 100).toFixed(2) + '%');

  // D1 を走査（列1:日付, 列5:売上, 列7:注文点数）
  const d1Sheet = getOrCreateSheet(SHEET_NAMES.D1_DAILY);
  const d1LastRow = d1Sheet.getLastRow();
  if (d1LastRow <= 1) return;

  const d1Data = d1Sheet.getRange(2, 1, d1LastRow - 1, 7).getValues();
  const updates = [];  // [[返品数, 返品額], ...]
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

  // 列14（返品数）、列15（返品額）に書込
  d1Sheet.getRange(2, 14, updates.length, 2).setValues(updates);
  Logger.log('✅ D1 返品バックフィル: ' + applied + ' / ' + updates.length + ' 行');
}

/**
 * D1 日次データの FBA手数料列（列13）を Settlement のFBA率 × 日次売上 で推定反映
 *
 * D2 の明細種別 "FBAPerUnitFulfillmentFee" を月次集計して Principal に対する率を算出。
 * 各D1行の「売上 × 率」を FBA手数料列に書込む（推定値）。
 */
function backfillD1FbaFeeFromSettlement() {
  Logger.log('===== D1 FBA手数料バックフィル 開始 =====');

  // D2 から月別 FBA率を算出
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

  // D1 を走査し、FBA手数料列（列13）を書き換え
  const d1Sheet = getOrCreateSheet(SHEET_NAMES.D1_DAILY);
  const d1LastRow = d1Sheet.getLastRow();
  if (d1LastRow <= 1) return;

  // 列1: 日付, 列5: 売上
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

  // 列13 に一括書込
  d1Sheet.getRange(2, 13, fbaFees.length, 1).setValues(fbaFees);
  Logger.log('✅ D1 FBA手数料バックフィル: ' + applied + ' / ' + fbaFees.length + ' 行');
}

/**
 * 診断: D1 で売上ありながら cogs=0 の ASIN を上位表示
 *
 * 売上のうち原価未計上分の規模を把握し、CF シートへの追加入力が必要な
 * ASIN を特定する。
 */
function debugD1CogsGap() {
  const sheet = getOrCreateSheet(SHEET_NAMES.D1_DAILY);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  // 今月の範囲を想定（直近2ヶ月を対象）
  const now = new Date();
  const currentYm = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');
  const prevMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const prevYm = prevMonth.getFullYear() + '-' + String(prevMonth.getMonth() + 1).padStart(2, '0');

  const data = sheet.getRange(2, 1, lastRow - 1, 21).getValues();
  const byAsin = {};  // asin → { name, salesNoCogs, salesWithCogs, totalSales, ymList }

  for (const row of data) {
    const ym = formatYearMonth(row[0]);
    if (ym !== currentYm && ym !== prevYm) continue;

    const asin = String(row[1] || '').trim();
    const name = String(row[2] || '').trim();
    if (!asin) continue;

    const sales = parseFloat(row[4]) || 0;
    const cogs = parseFloat(row[20]) || 0;

    if (!byAsin[asin]) byAsin[asin] = { name, salesNoCogs: 0, salesWithCogs: 0 };
    if (cogs > 0) byAsin[asin].salesWithCogs += sales;
    else byAsin[asin].salesNoCogs += sales;
  }

  const noCogsList = Object.entries(byAsin)
    .filter(([, v]) => v.salesNoCogs > 0)
    .sort((a, b) => b[1].salesNoCogs - a[1].salesNoCogs);

  const totalSalesNoCogs = noCogsList.reduce((s, [, v]) => s + v.salesNoCogs, 0);
  const totalSalesAll = Object.values(byAsin).reduce((s, v) => s + v.salesNoCogs + v.salesWithCogs, 0);

  Logger.log('===== D1 cogs ギャップ診断（当月＋前月）=====');
  Logger.log('対象期間: ' + prevYm + ' / ' + currentYm);
  Logger.log('総売上: ¥' + totalSalesAll.toLocaleString());
  Logger.log('うち cogs=0 売上: ¥' + totalSalesNoCogs.toLocaleString() +
    ' (' + (totalSalesAll > 0 ? (totalSalesNoCogs / totalSalesAll * 100).toFixed(1) : '-') + '%)');
  Logger.log('');
  Logger.log('--- cogs=0 の売上上位20商品 ---');
  noCogsList.slice(0, 20).forEach(([asin, v]) => {
    Logger.log('  ' + asin + ' | 売上: ¥' + v.salesNoCogs.toLocaleString() + ' | ' + (v.name || '(名前なし)'));
  });
  if (noCogsList.length > 20) Logger.log('  ... 他 ' + (noCogsList.length - 20) + ' 件');
}

/**
 * 診断: D1 日次データの仕入原価合計を年月別に集計表示

/**
 * 指定年月の (ASIN → 仕入単価) マップを M2 から取得
 * @param {string} yearMonth 'YYYY-MM'
 * @returns {Object} { asin: price }
 */
function getPriceMapForMonth(yearMonth) {
  const sheet = getOrCreateSheet(SHEET_NAMES.M2_PURCHASE_PRICE);
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow <= 1) return map;

  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  for (const row of data) {
    const asin = String(row[0] || '').trim();
    const ym = formatYearMonth(row[1]);  // Date/文字列両対応
    const price = parseFloat(row[2]) || 0;
    if (asin && ym === yearMonth && price > 0) {
      map[asin] = price;
    }
  }
  return map;
}

/**
 * 診断: D1 日次データの仕入原価合計を年月別に集計表示
 *
 * 現月・前月に cogs が正しく入っているか確認する。
 */
function debugD1CogsByMonth() {
  const sheet = getOrCreateSheet(SHEET_NAMES.D1_DAILY);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) { Logger.log('D1 空'); return; }

  // 列1: 日付, 列20: 仕入単価, 列21: 仕入原価合計
  const data = sheet.getRange(2, 1, lastRow - 1, 21).getValues();

  const stats = {};  // ym => { rowCount, cogsSum, withCogsCount }
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

/**
 * 診断: CFシートの構造を表示（書き込みなし）
 *
 * 実行結果を見て CF シート構造を把握し、必要に応じて本体ロジックを調整する。
 */
function testReadCfSheet() {
  const cfSheetId = getCredential('CF_SHEET_ID');
  const cfSs = SpreadsheetApp.openById(cfSheetId);

  Logger.log('===== CFシート シート一覧 =====');
  cfSs.getSheets().forEach(s => {
    Logger.log('  ' + s.getName() + ' (gid=' + s.getSheetId() + ')');
  });

  // 対象シート特定
  let cfSheet = cfSs.getSheets().find(s => s.getSheetId() === 347226196) || cfSs.getSheets()[0];
  Logger.log('対象シート: ' + cfSheet.getName());

  const lastRow = cfSheet.getLastRow();
  const lastCol = cfSheet.getLastColumn();
  Logger.log('サイズ: ' + lastRow + ' 行 × ' + lastCol + ' 列');

  const data = cfSheet.getRange(1, 1, lastRow, lastCol).getValues();

  // ヘッダー行特定
  const { asinRow, nameRow } = findHeaderRows(data, Math.min(10, lastRow));
  Logger.log('ASIN行: ' + (asinRow + 1) + ' / 商品名行: ' + (nameRow + 1));

  // 商品列
  const products = extractProductColumns(data, asinRow, nameRow, lastCol);
  Logger.log('商品列数: ' + products.length);
  products.slice(0, 5).forEach(p => {
    Logger.log('  列' + (p.col + 1) + ': ASIN=' + p.asin + ' / 名=' + p.name);
  });
  if (products.length > 5) Logger.log('  ...他 ' + (products.length - 5) + ' 件');

  // 仕入単価ラベル行
  const priceRows = findPriceRows(data);
  Logger.log('仕入単価ラベル行: ' + priceRows.length + ' 件');
  priceRows.forEach(r => {
    const ym = detectMonthForRow(data, r);
    const labels = [0, 1, 2, 3].map(c => String(data[r][c] || '').trim()).filter(v => v).join(' / ');
    Logger.log('  行' + (r + 1) + ': 月=' + (ym || '未特定') + ' / ラベル="' + labels + '"');
  });
}
