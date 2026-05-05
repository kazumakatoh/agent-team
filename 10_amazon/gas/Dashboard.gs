/**
 * Amazon Dashboard - 共通データヘルパー
 *
 * D1（日次データ） / D2S（経費月次集計）からの読み込み + 期間集計の共通処理。
 * L3 商品分析 / 週次・月次 AIレポート / カテゴリ月次（CategoryMonthly.gs）から共有される。
 */

// ===== データ取得 =====

function getDailyDataAll() {
  const sheet = getOrCreateSheet(SHEET_NAMES.D1_DAILY);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  // 23列: 1-22 既存 + 23 販売手数料(M1ベース推定)
  const data = sheet.getRange(2, 1, lastRow - 1, 23).getValues();
  return data.map(row => ({
    date: row[0] instanceof Date ? Utilities.formatDate(row[0], 'Asia/Tokyo', 'yyyy-MM-dd') : String(row[0]).substring(0, 10),
    asin: row[1],
    name: row[2],
    category: row[3],
    sales: parseFloat(row[4]) || 0,
    cv: parseFloat(row[5]) || 0,
    units: parseFloat(row[6]) || 0,
    sessions: parseFloat(row[7]) || 0,
    pv: parseFloat(row[8]) || 0,
    fbaFee: parseFloat(row[12]) || 0,    // FBA手数料（列13）
    returnUnits: parseFloat(row[13]) || 0, // 返品数（列14）
    returnAmount: parseFloat(row[14]) || 0, // 返品額（列15）
    adCost: parseFloat(row[15]) || 0,
    adSales: parseFloat(row[16]) || 0,
    unitPrice: parseFloat(row[19]) || 0, // 仕入単価（列20）
    cogs: parseFloat(row[20]) || 0,      // 仕入原価合計（列21）
    commission: parseFloat(row[22]) || 0, // 販売手数料（列23）
  }));
}

/**
 * 経費月次集計シートから全データ読み込み（高速）
 * 事前にbuildSettlementSummary()で生成されたシートを使う
 */
function readAllSettlement() {
  const sheet = getOrCreateSheet(SHEET_NAMES.D2S_SETTLEMENT_SUMMARY);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log('⚠️ 経費月次集計シートが空です。buildSettlementSummary() を実行してください。');
    return [];
  }

  // 列1:ASIN / 2:年月 / 3:販売手数料 / 4:その他経費 / 5:Principal売上
  const lastCol = Math.max(5, sheet.getLastColumn());
  const data = sheet.getRange(2, 1, lastRow - 1, Math.min(5, lastCol)).getValues();
  return data.map(row => ({
    asin: String(row[0] || '').trim(),
    yearMonth: formatYearMonth(row[1]),  // Date/文字列両対応
    commission: parseFloat(row[2]) || 0,
    other: parseFloat(row[3]) || 0,
    principal: parseFloat(row[4]) || 0,  // Settlement確定売上（rate算出用）
  }));
}

/**
 * 月次集計から期間でフィルタして集計（月単位の精度）
 */
function aggregateExpenses(allExpenses, startDate, endDate) {
  const startMonth = startDate.substring(0, 7);
  const endMonth = endDate.substring(0, 7);

  let commission = 0, other = 0, principal = 0;
  const byAsin = {};

  // 同月内の部分期間の場合、日数按分
  const startDay = parseInt(startDate.substring(8, 10));
  const endDay = parseInt(endDate.substring(8, 10));
  const sameMonth = startMonth === endMonth;

  for (const row of allExpenses) {
    if (row.yearMonth < startMonth || row.yearMonth > endMonth) continue;

    // 部分月の按分（同月内で日数指定がある場合）
    let ratio = 1.0;
    if (sameMonth) {
      const y = parseInt(startMonth.substring(0, 4));
      const m = parseInt(startMonth.substring(5, 7));
      const daysInMonth = new Date(y, m, 0).getDate();
      const daysInRange = endDay - startDay + 1;
      ratio = daysInRange / daysInMonth;
    }

    const c = row.commission * ratio;
    const o = row.other * ratio;
    const p = (row.principal || 0) * ratio;

    commission += c;
    other += o;
    principal += p;

    if (row.asin) {
      if (!byAsin[row.asin]) byAsin[row.asin] = { commission: 0, other: 0 };
      byAsin[row.asin].commission += c;
      byAsin[row.asin].other += o;
    }
  }

  // Settlement確定分から率を算出（D1売上 vs D2 Principal のタイミング不一致を補正）
  const commissionRate = principal > 0 ? commission / principal : 0;
  const otherRate = principal > 0 ? other / principal : 0;

  return { commission, other, total: commission + other, byAsin, principal, commissionRate, otherRate };
}

/**
 * Settlement確定比率を使って D1売上相当の経費を推定
 *
 * @param {number} d1Sales D1日次データから集計した売上（確定+未確定）
 * @param {Object} expenses aggregateExpensesの戻り値
 * @returns {Object} { commission, other, isEstimated }
 */
function estimateExpensesFromRate(d1Sales, expenses) {
  // Settlement から算出した比率が有効 かつ
  // D1売上が D2 Principal より多い（未確定分が残っている）なら推定
  if (expenses.principal > 0 && d1Sales > expenses.principal && expenses.commissionRate > 0) {
    return {
      commission: d1Sales * expenses.commissionRate,
      other: d1Sales * expenses.otherRate,
      isEstimated: true,
    };
  }
  // Settlement 完全確定済み or D2 データなし → そのまま返す
  return {
    commission: expenses.commission,
    other: expenses.other,
    isEstimated: false,
  };
}

function getPeriods() {
  const today = new Date();
  const y = today.getFullYear();
  const m = today.getMonth();
  const d = today.getDate();

  const daysInMonth = new Date(y, m + 1, 0).getDate();
  const elapsedDays = d;

  const fmt = d => Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');

  return {
    thisMonth: { start: fmt(new Date(y, m, 1)), end: fmt(today) },
    lastMonth: { start: fmt(new Date(y, m - 1, 1)), end: fmt(new Date(y, m, 0)) },
    lastMonthSameDay: {
      start: fmt(new Date(y, m - 1, 1)),
      end: fmt(new Date(y, m - 1, Math.min(d, new Date(y, m, 0).getDate()))),
    },
    prevYear: { start: fmt(new Date(y - 1, m, 1)), end: fmt(new Date(y - 1, m + 1, 0)) },
    ytd: { start: fmt(new Date(y, 0, 1)), end: fmt(today) },
    daysInMonth,
    elapsedDays,
  };
}

// ===== カテゴリ別集計（1パスで全カテゴリ集計） =====

/**
 * 事前フィルタ済みの日次データから、カテゴリ別に1パスで集計
 * @returns {Object} { category: aggregatedData }
 */
function aggregateByCategory(filteredDaily, expenses) {
  const byCategory = {};

  for (const d of filteredDaily) {
    const cat = d.category || '(未分類)';
    if (!byCategory[cat]) {
      byCategory[cat] = {
        category: cat,
        sales: 0, cv: 0, units: 0, sessions: 0, pv: 0,
        adCost: 0, adSales: 0, cogs: 0,
        asins: new Set(),
      };
    }
    const c = byCategory[cat];
    c.sales += d.sales;
    c.cv += d.cv;
    c.units += d.units;
    c.sessions += d.sessions;
    c.pv += d.pv;
    c.adCost += d.adCost;
    c.adSales += d.adSales;
    c.cogs += d.cogs;
    if (d.asin) c.asins.add(d.asin);
  }

  // カテゴリごとに経費を集計
  for (const cat of Object.values(byCategory)) {
    let commission = 0, otherExpense = 0;
    for (const asin of cat.asins) {
      const exp = expenses.byAsin[asin];
      if (exp) {
        commission += exp.commission || 0;
        otherExpense += exp.other || 0;
      }
    }
    Object.assign(cat, computeDerivedMetrics(cat, commission, otherExpense));
  }

  return byCategory;
}

function computeDerivedMetrics(base, commission, otherExpense) {
  const sales = base.sales;
  const adCost = base.adCost;
  const adSales = base.adSales;
  const cogs = base.cogs || 0; // CF連携: D1.仕入原価合計 から集計
  const grossProfit = sales - cogs - commission - otherExpense;
  const profit = grossProfit - adCost;

  return {
    commission, otherExpense, expense: commission + otherExpense,
    cogs, grossProfit, profit,
    costRate: sales > 0 ? cogs / sales : 0,
    grossMargin: sales > 0 ? grossProfit / sales : 0,
    adRate: sales > 0 ? adCost / sales : 0,
    profitMargin: sales > 0 ? profit / sales : 0,
    tacos: sales > 0 ? (adCost / sales * 100) : 0,
    acos: adSales > 0 ? (adCost / adSales * 100) : 0,
    roas: adCost > 0 ? (sales / adCost) : 0,
    organicShare: sales > 0 ? ((sales - adSales) / sales * 100) : 0,
  };
}
