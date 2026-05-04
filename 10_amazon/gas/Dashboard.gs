/**
 * Amazon Dashboard - 共通ヘルパー
 *
 * D1（日次データ）/ D2S（経費月次集計）の読み込み + 期間ユーティリティ。
 * 旧 L1 事業ダッシュボード / L2 カテゴリ分析の構築コードはこのファイルから削除済み。
 * （新カテゴリ別月次シートは CategoryMonthly.gs で実装）
 */

/**
 * D1 日次データを全行読み込んでオブジェクト配列に変換。
 * 列マッピング:
 *   A(0)=日付 / B(1)=ASIN / C(2)=商品名 / D(3)=カテゴリ
 *   E(4)=売上 / F(5)=CV / G(6)=点数
 *   H(7)=セッション / I(8)=PV / J(9)=CTR / K(10)=CVR / L(11)=BuyBox率
 *   M(12)=FBA手数料 / N(13)=返品数 / O(14)=返品額
 *   P(15)=広告費 / Q(16)=広告売上 / R(17)=IMP / S(18)=CT
 *   T(19)=仕入単価 / U(20)=仕入原価合計 / V(21)=ステータス
 */
function getDailyDataAll() {
  const sheet = getOrCreateSheet(SHEET_NAMES.D1_DAILY);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 22).getValues();
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
    fbaFee: parseFloat(row[12]) || 0,
    returnUnits: parseFloat(row[13]) || 0,
    returnAmount: parseFloat(row[14]) || 0,
    adCost: parseFloat(row[15]) || 0,
    adSales: parseFloat(row[16]) || 0,
    adImp: parseFloat(row[17]) || 0,
    adCt: parseFloat(row[18]) || 0,
    unitPrice: parseFloat(row[19]) || 0,
    cogs: parseFloat(row[20]) || 0,
  }));
}

/**
 * 経費月次集計シートから全データ読み込み（高速）
 * 事前に buildSettlementSummary() で生成されたシートを使う
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
    yearMonth: formatYearMonth(row[1]),
    commission: parseFloat(row[2]) || 0,
    other: parseFloat(row[3]) || 0,
    principal: parseFloat(row[4]) || 0,
  }));
}

/**
 * 月次集計から期間でフィルタして集計（月単位の精度）
 *
 * @param {Array} allExpenses readAllSettlement() の戻り値
 * @param {string} startDate 'YYYY-MM-DD'
 * @param {string} endDate   'YYYY-MM-DD'
 * @returns {Object} { commission, other, total, byAsin, principal, commissionRate, otherRate }
 */
function aggregateExpenses(allExpenses, startDate, endDate) {
  const startMonth = startDate.substring(0, 7);
  const endMonth = endDate.substring(0, 7);

  let commission = 0, other = 0, principal = 0;
  const byAsin = {};

  const startDay = parseInt(startDate.substring(8, 10));
  const endDay = parseInt(endDate.substring(8, 10));
  const sameMonth = startMonth === endMonth;

  for (const row of allExpenses) {
    if (row.yearMonth < startMonth || row.yearMonth > endMonth) continue;

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

  const commissionRate = principal > 0 ? commission / principal : 0;
  const otherRate = principal > 0 ? other / principal : 0;

  return { commission, other, total: commission + other, byAsin, principal, commissionRate, otherRate };
}

/**
 * Settlement 確定比率を使って D1 売上相当の経費を推定
 */
function estimateExpensesFromRate(d1Sales, expenses) {
  if (expenses.principal > 0 && d1Sales > expenses.principal && expenses.commissionRate > 0) {
    return {
      commission: d1Sales * expenses.commissionRate,
      other: d1Sales * expenses.otherRate,
      isEstimated: true,
    };
  }
  return {
    commission: expenses.commission,
    other: expenses.other,
    isEstimated: false,
  };
}

/**
 * 期間定義（当月 / 前月 / 前月同消化日 / 前年同月 / YTD）
 */
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
