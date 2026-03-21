/**
 * 民泊自動化システム - KPI計算モジュール
 *
 * KPI定義：
 *  ADR     = 売上 / 稼働日数（1泊あたりの平均客室収益）
 *  RevPAR  = 売上 / 利用可能日数（1日あたりの潜在収益）
 *  稼働率  = 稼働日数 / 利用可能日数 × 100
 *  ROI     = 利益 / 総経費 × 100
 *
 * 民泊法上の年間集計期間：毎年4月1日〜翌年3月31日（上限180日）
 */

const KPICalculator = {

  /**
   * 月次KPIを計算する
   * @param {Object} params
   * @param {number} params.revenue      - 売上（手数料込み）
   * @param {number} params.commission   - 手数料
   * @param {number} params.usageDays    - 稼働日数（当月の宿泊日数合計）
   * @param {number} params.daysInMonth  - 当月の暦日数
   * @param {number} params.totalCosts   - 総経費（清掃費+備品+水光熱+家賃+その他）
   * @return {Object} { roi, adr, revpar, occupancyRate }
   */
  calcMonthlyKPIs(params) {
    const { revenue, commission, usageDays, daysInMonth, totalCosts } = params;
    const netRevenue = revenue - commission;
    const profit     = netRevenue - totalCosts;

    return {
      roi:           this.calcROI(profit, totalCosts),
      adr:           this.calcADR(revenue, usageDays),
      revpar:        this.calcRevPAR(revenue, daysInMonth),
      occupancyRate: this.calcOccupancyRate(usageDays, daysInMonth)
    };
  },

  /**
   * 年次KPIを計算する
   * @param {Array<Object>} monthlyDataArray - 12ヶ月分の集計データ
   * @return {Object}
   */
  calcAnnualKPIs(monthlyDataArray) {
    const totals = monthlyDataArray.reduce((acc, m) => {
      acc.revenue      += m.revenue      || 0;
      acc.commission   += m.commission   || 0;
      acc.usageDays    += m.usageDays    || 0;
      acc.totalCosts   += m.totalCosts   || 0;
      acc.daysInYear   += m.daysInMonth  || 0;
      acc.bookingCount += m.bookingCount || 0;
      return acc;
    }, { revenue: 0, commission: 0, usageDays: 0, totalCosts: 0, daysInYear: 0, bookingCount: 0 });

    const netRevenue = totals.revenue - totals.commission;
    const profit     = netRevenue - totals.totalCosts;

    return {
      ...totals,
      netRevenue,
      profit,
      roi:                this.calcROI(profit, totals.totalCosts),
      adr:                this.calcADR(totals.revenue, totals.usageDays),
      revpar:             this.calcRevPAR(totals.revenue, totals.daysInYear),
      occupancyRate:      this.calcOccupancyRate(totals.usageDays, totals.daysInYear),
      legalOccupancyRate: this.calcOccupancyRate(totals.usageDays, CONFIG.PROPERTY.MAX_ANNUAL_DAYS)
    };
  },

  /**
   * ROI（投資収益率）= 利益 / 総経費 × 100
   */
  calcROI(profit, totalCosts) {
    if (!totalCosts || totalCosts === 0) return 0;
    return Math.round((profit / totalCosts) * 1000) / 10; // 小数1位
  },

  /**
   * ADR（平均日次収益）= 売上 / 稼働日数
   * 稼働がない月は0を返す
   */
  calcADR(revenue, usageDays) {
    if (!usageDays || usageDays === 0) return 0;
    return Math.round(revenue / usageDays);
  },

  /**
   * RevPAR（利用可能室あたり収益）= 売上 / 利用可能日数
   */
  calcRevPAR(revenue, availableDays) {
    if (!availableDays || availableDays === 0) return 0;
    return Math.round(revenue / availableDays);
  },

  /**
   * 稼働率 = 稼働日数 / 利用可能日数 × 100
   */
  calcOccupancyRate(usageDays, availableDays) {
    if (!availableDays || availableDays === 0) return 0;
    return Math.round((usageDays / availableDays) * 1000) / 10; // 小数1位
  },

  /**
   * 残り利用可能日数（年間180日上限に対して）
   * @param {number} usedDays - 今年度の稼働日数合計
   * @return {number}
   */
  calcRemainingLegalDays(usedDays) {
    return Math.max(0, CONFIG.PROPERTY.MAX_ANNUAL_DAYS - usedDays);
  },

  /**
   * 現在の年度を取得する（4月始まり）
   * @return {number} 年度の開始年（例: 2025年4月〜2026年3月 → 2025）
   */
  getCurrentFiscalYear() {
    const now = new Date();
    const month = now.getMonth() + 1; // 1-12
    const year  = now.getFullYear();
    return month >= 4 ? year : year - 1;
  },

  /**
   * 指定日が何年度に属するかを返す
   * @param {Date} date
   * @return {number}
   */
  getFiscalYearOf(date) {
    const month = date.getMonth() + 1;
    const year  = date.getFullYear();
    return month >= 4 ? year : year - 1;
  }
};
