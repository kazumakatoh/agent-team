/**
 * Amazon Dashboard - L1 事業ダッシュボード構築
 */

/**
 * L1 事業ダッシュボードを更新
 */
function updateDashboardL1() {
  Logger.log('===== L1 事業ダッシュボード更新開始 =====');

  const sheet = getOrCreateSheet(SHEET_NAMES.L1_DASHBOARD);
  sheet.clear();

  const dailyData = getDailyDataAll();
  if (dailyData.length === 0) return;

  const periods = getPeriods();

  // 経費データを取得
  const thisMonthExp = getSettlementExpenses(periods.thisMonth.start, periods.thisMonth.end);
  const lastMonthSameDayExp = getSettlementExpenses(periods.lastMonthSameDay.start, periods.lastMonthSameDay.end);

  const totals = {
    thisMonth: aggregateData(dailyData, periods.thisMonth.start, periods.thisMonth.end, null, thisMonthExp),
    lastMonthSameDay: aggregateData(dailyData, periods.lastMonthSameDay.start, periods.lastMonthSameDay.end, null, lastMonthSameDayExp),
  };

  writeOverallSummary(sheet, totals, periods);
  const alertStartRow = writeCategorySummary(sheet, dailyData, periods, thisMonthExp, lastMonthSameDayExp);
  writeAlertProducts(sheet, dailyData, periods.thisMonth, alertStartRow);

  Logger.log('===== L1 完了 =====');
}



/**
 * 日次データを全行取得
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
    adCost: parseFloat(row[15]) || 0,
    adSales: parseFloat(row[16]) || 0,
  }));
}

/**
 * 期間を計算
 */
function getPeriods() {
  const today = new Date();
  const y = today.getFullYear();
  const m = today.getMonth();
  const d = today.getDate();

  // 当月の日数と現在の経過日数
  const daysInMonth = new Date(y, m + 1, 0).getDate();
  const elapsedDays = d;

  const thisMonthStart = new Date(y, m, 1);
  const thisMonthEnd = today;

  // 先月
  const lastMonthStart = new Date(y, m - 1, 1);
  const lastMonthEnd = new Date(y, m, 0);

  // 先月の同日数（比較用）
  const lastMonthSameDayStart = new Date(y, m - 1, 1);
  const lastMonthSameDayEnd = new Date(y, m - 1, Math.min(d, new Date(y, m, 0).getDate()));

  const prevYearStart = new Date(y - 1, m, 1);
  const prevYearEnd = new Date(y - 1, m + 1, 0);
  const ytdStart = new Date(y, 0, 1);
  const ytdEnd = today;

  const fmt = d => Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');

  return {
    thisMonth: { start: fmt(thisMonthStart), end: fmt(thisMonthEnd) },
    lastMonth: { start: fmt(lastMonthStart), end: fmt(lastMonthEnd) },
    lastMonthSameDay: { start: fmt(lastMonthSameDayStart), end: fmt(lastMonthSameDayEnd) },
    prevYear: { start: fmt(prevYearStart), end: fmt(prevYearEnd) },
    ytd: { start: fmt(ytdStart), end: fmt(ytdEnd) },
    daysInMonth: daysInMonth,
    elapsedDays: elapsedDays,
  };
}


/**
 * 期間内のデータを集計
 */
function aggregateData(dailyData, startDate, endDate, filterFn, expenses) {
  const filtered = dailyData.filter(d => {
    if (!d.date || d.date < startDate || d.date > endDate) return false;
    if (filterFn && !filterFn(d)) return false;
    return true;
  });

  const sales = filtered.reduce((s, d) => s + d.sales, 0);
  const cv = filtered.reduce((s, d) => s + d.cv, 0);
  const units = filtered.reduce((s, d) => s + d.units, 0);
  const sessions = filtered.reduce((s, d) => s + d.sessions, 0);
  const adCost = filtered.reduce((s, d) => s + d.adCost, 0);
  const adSales = filtered.reduce((s, d) => s + d.adSales, 0);

  let commission = 0, otherExpense = 0;
  if (expenses) {
    if (filterFn) {
      const asinsInFilter = new Set(filtered.map(d => d.asin));
      for (const [asin, exp] of Object.entries(expenses.byAsin)) {
        if (asinsInFilter.has(asin)) {
          commission += exp.commission || 0;
          otherExpense += exp.other || 0;
        }
      }
    } else {
      commission = expenses.commission;
      otherExpense = expenses.other;
    }
  }

  const cogs = 0;  // TODO: CFシート連携
  const grossProfit = sales - cogs - commission - otherExpense;
  const profit = grossProfit - adCost;

  return {
    sales, cv, units, sessions, adCost, adSales,
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



/**
 * 全体サマリーを書き込み
 */
function writeOverallSummary(sheet, totals, periods) {
  const t = totals.thisMonth;
  const lm = totals.lastMonthSameDay;
  const organicSales = t.sales - t.adSales;
  const fc = periods.daysInMonth / periods.elapsedDays;

  sheet.getRange(1, 1).setValue('━━━ Amazon事業 全体サマリー（当月）━━━').setFontWeight('bold').setFontSize(14);

  const headers = ['軸', '売上', '売上比', 'CV', '点数', '広告費', '販売手数料', '経費等', '利益', '原価率', '粗利率', '広告比率', 'ROAS', '利益率'];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center');

  const totalRow = ['全体', t.sales, 1, t.cv, t.units, t.adCost, t.commission, t.otherExpense, t.profit, t.costRate, t.grossMargin, t.adRate, t.roas, t.profitMargin];
  const adRow = ['広告', t.adSales, t.sales > 0 ? t.adSales/t.sales : 0, 0, 0, t.adCost, '-', '-', '-', '-', '-', '-', t.roas, '-'];
  const organicRow = ['オーガニック', organicSales, t.sales > 0 ? organicSales/t.sales : 0, t.cv, t.units, '-', '-', '-', '-', '-', '-', '-', '-', '-'];
  const pctRow = ['前月比',
    pctChangeNum(t.sales, lm.sales), '-',
    pctChangeNum(t.cv, lm.cv),
    pctChangeNum(t.units, lm.units),
    pctChangeNum(t.adCost, lm.adCost),
    pctChangeNum(t.commission, lm.commission),
    pctChangeNum(t.otherExpense, lm.otherExpense),
    pctChangeNum(t.profit, lm.profit),
    pctChangeNum(t.costRate, lm.costRate),
    pctChangeNum(t.grossMargin, lm.grossMargin),
    pctChangeNum(t.adRate, lm.adRate),
    pctChangeNum(t.roas, lm.roas),
    pctChangeNum(t.profitMargin, lm.profitMargin),
  ];
  const fcRow = ['月末予測', t.sales*fc, '-', t.cv*fc, t.units*fc, t.adCost*fc, t.commission*fc, t.otherExpense*fc, t.profit*fc, '-', '-', '-', '-', '-'];

  const rows = [totalRow, adRow, organicRow, pctRow, fcRow];
  sheet.getRange(3, 1, rows.length, headers.length).setValues(rows);

  sheet.getRange(3, 1, rows.length, 1).setHorizontalAlignment('center').setFontWeight('bold');
  sheet.getRange(3, 2, rows.length, headers.length - 1).setHorizontalAlignment('right');

  // 数値列: 売上(2), CV(4), 点数(5), 広告費(6), 販売手数料(7), 経費等(8), 利益(9)
  [2, 4, 5, 6, 7, 8, 9].forEach(col => {
    sheet.getRange(3, col, 3, 1).setNumberFormat('#,##0');
    sheet.getRange(7, col, 1, 1).setNumberFormat('#,##0');
  });
  // 率列: 売上比(3), 原価率(10), 粗利率(11), 広告比率(12), 利益率(14)
  [3, 10, 11, 12, 14].forEach(col => {
    sheet.getRange(3, col, 3, 1).setNumberFormat('0.0%');
  });
  // ROAS(13)
  sheet.getRange(3, 13, 3, 1).setNumberFormat('0.00');
  // 前月比行（6行目）
  sheet.getRange(6, 2, 1, headers.length - 1).setNumberFormat('+0.0%;-0.0%;-');
}





function writeOverallSummary(sheet, totals, periods) {
  const t = totals.thisMonth;
  const lm = totals.lastMonthSameDay;
  const organicSales = t.sales - t.adSales;
  const fc = periods.daysInMonth / periods.elapsedDays;

  sheet.getRange(1, 1).setValue('━━━ Amazon事業 全体サマリー（当月）━━━').setFontWeight('bold').setFontSize(14);

  const headers = ['軸', '売上', '売上比', 'CV', '点数', '広告費', '販売手数料', '経費等', '利益', '原価率', '粗利率', '広告比率', 'ROAS', '利益率'];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center');

  const totalRow = ['全体', t.sales, 1, t.cv, t.units, t.adCost, t.commission, t.otherExpense, t.profit, t.costRate, t.grossMargin, t.adRate, t.roas, t.profitMargin];
  const adRow = ['広告', t.adSales, t.sales > 0 ? t.adSales/t.sales : 0, 0, 0, t.adCost, '-', '-', '-', '-', '-', '-', t.roas, '-'];
  const organicRow = ['オーガニック', organicSales, t.sales > 0 ? organicSales/t.sales : 0, t.cv, t.units, '-', '-', '-', '-', '-', '-', '-', '-', '-'];
  const pctRow = ['前月比',
    pctChangeNum(t.sales, lm.sales), '-',
    pctChangeNum(t.cv, lm.cv),
    pctChangeNum(t.units, lm.units),
    pctChangeNum(t.adCost, lm.adCost),
    pctChangeNum(t.commission, lm.commission),
    pctChangeNum(t.otherExpense, lm.otherExpense),
    pctChangeNum(t.profit, lm.profit),
    pctChangeNum(t.costRate, lm.costRate),
    pctChangeNum(t.grossMargin, lm.grossMargin),
    pctChangeNum(t.adRate, lm.adRate),
    pctChangeNum(t.roas, lm.roas),
    pctChangeNum(t.profitMargin, lm.profitMargin),
  ];
  const fcRow = ['月末予測', t.sales*fc, '-', t.cv*fc, t.units*fc, t.adCost*fc, t.commission*fc, t.otherExpense*fc, t.profit*fc, '-', '-', '-', '-', '-'];

  const rows = [totalRow, adRow, organicRow, pctRow, fcRow];
  sheet.getRange(3, 1, rows.length, headers.length).setValues(rows);

  sheet.getRange(3, 1, rows.length, 1).setHorizontalAlignment('center').setFontWeight('bold');
  sheet.getRange(3, 2, rows.length, headers.length - 1).setHorizontalAlignment('right');

  // 数値列: 売上(2), CV(4), 点数(5), 広告費(6), 販売手数料(7), 経費等(8), 利益(9)
  [2, 4, 5, 6, 7, 8, 9].forEach(col => {
    sheet.getRange(3, col, 3, 1).setNumberFormat('#,##0');
    sheet.getRange(7, col, 1, 1).setNumberFormat('#,##0');
  });
  // 率列: 売上比(3), 原価率(10), 粗利率(11), 広告比率(12), 利益率(14)
  [3, 10, 11, 12, 14].forEach(col => {
    sheet.getRange(3, col, 3, 1).setNumberFormat('0.0%');
  });
  // ROAS(13)
  sheet.getRange(3, 13, 3, 1).setNumberFormat('0.00');
  // 前月比行（6行目）
  sheet.getRange(6, 2, 1, headers.length - 1).setNumberFormat('+0.0%;-0.0%;-');
}





function writeCategorySummary(sheet, dailyData, periods, thisMonthExp, lastMonthSameDayExp) {
  const startRow = 10;
  sheet.getRange(startRow, 1).setValue('━━━ カテゴリ別サマリー（当月）━━━').setFontWeight('bold').setFontSize(14);

  const headers = ['カテゴリ', '売上', '売上比', 'CV', '点数', '広告費', '利益', 'TACOS', 'TACOS前月比', 'ACOS', 'ACOS前月比', 'ROAS', '利益率'];
  sheet.getRange(startRow + 1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center');

  const categories = [...new Set(dailyData.filter(d => d.category).map(d => d.category))].sort();
  const totalSales = dailyData
    .filter(d => d.date >= periods.thisMonth.start && d.date <= periods.thisMonth.end)
    .reduce((s, d) => s + d.sales, 0);

  const catData = categories.map(cat => {
    const agg = aggregateData(dailyData, periods.thisMonth.start, periods.thisMonth.end, d => d.category === cat, thisMonthExp);
    const lmAgg = aggregateData(dailyData, periods.lastMonthSameDay.start, periods.lastMonthSameDay.end, d => d.category === cat, lastMonthSameDayExp);
    return { cat, agg, lmAgg };
  });

  const filtered = catData.filter(c => c.agg.sales > 0 || c.agg.cv > 0)
    .sort((a, b) => b.agg.sales - a.agg.sales);

  const rows = filtered.map(c => [
    c.cat,
    c.agg.sales,
    totalSales > 0 ? c.agg.sales / totalSales : 0,
    c.agg.cv,
    c.agg.units,
    c.agg.adCost,
    c.agg.profit,
    c.agg.tacos / 100,
    pctChangeNum(c.agg.tacos, c.lmAgg.tacos),
    c.agg.acos / 100,
    pctChangeNum(c.agg.acos, c.lmAgg.acos),
    c.agg.roas,
    c.agg.profitMargin,
  ]);

  if (rows.length > 0) {
    const dataRow = startRow + 2;
    sheet.getRange(dataRow, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(dataRow, 1, rows.length, 1).setHorizontalAlignment('center');
    sheet.getRange(dataRow, 2, rows.length, headers.length - 1).setHorizontalAlignment('right');
    sheet.getRange(dataRow, 2, rows.length, 1).setNumberFormat('#,##0');  // 売上
    sheet.getRange(dataRow, 3, rows.length, 1).setNumberFormat('0.0%');   // 売上比
    sheet.getRange(dataRow, 4, rows.length, 2).setNumberFormat('#,##0');  // CV, 点数
    sheet.getRange(dataRow, 6, rows.length, 2).setNumberFormat('#,##0');  // 広告費, 利益
    sheet.getRange(dataRow, 8, rows.length, 1).setNumberFormat('0.0%');   // TACOS
    sheet.getRange(dataRow, 9, rows.length, 1).setNumberFormat('+0.0%;-0.0%;-');
    sheet.getRange(dataRow, 10, rows.length, 1).setNumberFormat('0.0%');  // ACOS
    sheet.getRange(dataRow, 11, rows.length, 1).setNumberFormat('+0.0%;-0.0%;-');
    sheet.getRange(dataRow, 12, rows.length, 1).setNumberFormat('0.00');  // ROAS
    sheet.getRange(dataRow, 13, rows.length, 1).setNumberFormat('0.0%');  // 利益率
  }

  return startRow + 2 + rows.length + 2;
}







/**
 * カテゴリ別サマリー
 */
function writeOverallSummary(sheet, totals, periods) {
  const t = totals.thisMonth;
  const lm = totals.lastMonthSameDay;
  const organicSales = t.sales - t.adSales;
  const fc = periods.daysInMonth / periods.elapsedDays;

  sheet.getRange(1, 1).setValue('━━━ Amazon事業 全体サマリー（当月）━━━').setFontWeight('bold').setFontSize(14);

  const headers = ['軸', '売上', '売上比', 'CV', '点数', '広告費', '販売手数料', '経費等', '利益', '原価率', '粗利率', '広告比率', 'ROAS', '利益率'];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center');

  const totalRow = ['全体', t.sales, 1, t.cv, t.units, t.adCost, t.commission, t.otherExpense, t.profit, t.costRate, t.grossMargin, t.adRate, t.roas, t.profitMargin];
  const adRow = ['広告', t.adSales, t.sales > 0 ? t.adSales/t.sales : 0, 0, 0, t.adCost, '-', '-', '-', '-', '-', '-', t.roas, '-'];
  const organicRow = ['オーガニック', organicSales, t.sales > 0 ? organicSales/t.sales : 0, t.cv, t.units, '-', '-', '-', '-', '-', '-', '-', '-', '-'];
  const pctRow = ['前月比',
    pctChangeNum(t.sales, lm.sales), '-',
    pctChangeNum(t.cv, lm.cv),
    pctChangeNum(t.units, lm.units),
    pctChangeNum(t.adCost, lm.adCost),
    pctChangeNum(t.commission, lm.commission),
    pctChangeNum(t.otherExpense, lm.otherExpense),
    pctChangeNum(t.profit, lm.profit),
    pctChangeNum(t.costRate, lm.costRate),
    pctChangeNum(t.grossMargin, lm.grossMargin),
    pctChangeNum(t.adRate, lm.adRate),
    pctChangeNum(t.roas, lm.roas),
    pctChangeNum(t.profitMargin, lm.profitMargin),
  ];
  const fcRow = ['月末予測', t.sales*fc, '-', t.cv*fc, t.units*fc, t.adCost*fc, t.commission*fc, t.otherExpense*fc, t.profit*fc, '-', '-', '-', '-', '-'];

  const rows = [totalRow, adRow, organicRow, pctRow, fcRow];
  sheet.getRange(3, 1, rows.length, headers.length).setValues(rows);

  sheet.getRange(3, 1, rows.length, 1).setHorizontalAlignment('center').setFontWeight('bold');
  sheet.getRange(3, 2, rows.length, headers.length - 1).setHorizontalAlignment('right');

  // 数値列: 売上(2), CV(4), 点数(5), 広告費(6), 販売手数料(7), 経費等(8), 利益(9)
  [2, 4, 5, 6, 7, 8, 9].forEach(col => {
    sheet.getRange(3, col, 3, 1).setNumberFormat('#,##0');
    sheet.getRange(7, col, 1, 1).setNumberFormat('#,##0');
  });
  // 率列: 売上比(3), 原価率(10), 粗利率(11), 広告比率(12), 利益率(14)
  [3, 10, 11, 12, 14].forEach(col => {
    sheet.getRange(3, col, 3, 1).setNumberFormat('0.0%');
  });
  // ROAS(13)
  sheet.getRange(3, 13, 3, 1).setNumberFormat('0.00');
  // 前月比行（6行目）
  sheet.getRange(6, 2, 1, headers.length - 1).setNumberFormat('+0.0%;-0.0%;-');
}





function writeCategorySummary(sheet, dailyData, periods, thisMonthExp, lastMonthSameDayExp) {
  const startRow = 10;
  sheet.getRange(startRow, 1).setValue('━━━ カテゴリ別サマリー（当月）━━━').setFontWeight('bold').setFontSize(14);

  const headers = ['カテゴリ', '売上', '売上比', 'CV', '点数', '広告費', '利益', 'TACOS', 'TACOS前月比', 'ACOS', 'ACOS前月比', 'ROAS', '利益率'];
  sheet.getRange(startRow + 1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center');

  const categories = [...new Set(dailyData.filter(d => d.category).map(d => d.category))].sort();
  const totalSales = dailyData
    .filter(d => d.date >= periods.thisMonth.start && d.date <= periods.thisMonth.end)
    .reduce((s, d) => s + d.sales, 0);

  const catData = categories.map(cat => {
    const agg = aggregateData(dailyData, periods.thisMonth.start, periods.thisMonth.end, d => d.category === cat, thisMonthExp);
    const lmAgg = aggregateData(dailyData, periods.lastMonthSameDay.start, periods.lastMonthSameDay.end, d => d.category === cat, lastMonthSameDayExp);
    return { cat, agg, lmAgg };
  });

  const filtered = catData.filter(c => c.agg.sales > 0 || c.agg.cv > 0)
    .sort((a, b) => b.agg.sales - a.agg.sales);

  const rows = filtered.map(c => [
    c.cat,
    c.agg.sales,
    totalSales > 0 ? c.agg.sales / totalSales : 0,
    c.agg.cv,
    c.agg.units,
    c.agg.adCost,
    c.agg.profit,
    c.agg.tacos / 100,
    pctChangeNum(c.agg.tacos, c.lmAgg.tacos),
    c.agg.acos / 100,
    pctChangeNum(c.agg.acos, c.lmAgg.acos),
    c.agg.roas,
    c.agg.profitMargin,
  ]);

  if (rows.length > 0) {
    const dataRow = startRow + 2;
    sheet.getRange(dataRow, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(dataRow, 1, rows.length, 1).setHorizontalAlignment('center');
    sheet.getRange(dataRow, 2, rows.length, headers.length - 1).setHorizontalAlignment('right');
    sheet.getRange(dataRow, 2, rows.length, 1).setNumberFormat('#,##0');  // 売上
    sheet.getRange(dataRow, 3, rows.length, 1).setNumberFormat('0.0%');   // 売上比
    sheet.getRange(dataRow, 4, rows.length, 2).setNumberFormat('#,##0');  // CV, 点数
    sheet.getRange(dataRow, 6, rows.length, 2).setNumberFormat('#,##0');  // 広告費, 利益
    sheet.getRange(dataRow, 8, rows.length, 1).setNumberFormat('0.0%');   // TACOS
    sheet.getRange(dataRow, 9, rows.length, 1).setNumberFormat('+0.0%;-0.0%;-');
    sheet.getRange(dataRow, 10, rows.length, 1).setNumberFormat('0.0%');  // ACOS
    sheet.getRange(dataRow, 11, rows.length, 1).setNumberFormat('+0.0%;-0.0%;-');
    sheet.getRange(dataRow, 12, rows.length, 1).setNumberFormat('0.00');  // ROAS
    sheet.getRange(dataRow, 13, rows.length, 1).setNumberFormat('0.0%');  // 利益率
  }

  return startRow + 2 + rows.length + 2;
}






function writeAlertProducts(sheet, dailyData, period, startRow) {
  sheet.getRange(startRow, 1).setValue('━━━ 注意が必要な商品 ━━━').setFontWeight('bold').setFontSize(14);
  sheet.getRange(startRow + 1, 1, 1, 4).setValues([['ASIN', '商品名', 'カテゴリ', '理由']])
    .setFontWeight('bold').setBackground('#fce8e6').setHorizontalAlignment('center');

  const asinMap = {};
  dailyData.filter(d => d.date >= period.start && d.date <= period.end).forEach(d => {
    if (!asinMap[d.asin]) {
      asinMap[d.asin] = { name: d.name, category: d.category, sales: 0, adCost: 0, adSales: 0 };
    }
    asinMap[d.asin].sales += d.sales;
    asinMap[d.asin].adCost += d.adCost;
    asinMap[d.asin].adSales += d.adSales;
  });

  const alerts = [];
  for (const [asin, d] of Object.entries(asinMap)) {
    const tacos = d.sales > 0 ? (d.adCost / d.sales * 100) : 0;
    if (tacos > 30) {
      alerts.push([asin, d.name, d.category, 'TACOS高: ' + tacos.toFixed(1) + '%']);
    }
    if (d.sales === 0 && d.adCost > 0) {
      alerts.push([asin, d.name, d.category, '広告費あり・売上ゼロ']);
    }
  }

  if (alerts.length > 0) {
    const dataRow = startRow + 2;
    sheet.getRange(dataRow, 1, alerts.length, 4).setValues(alerts);
    // ASIN・カテゴリ列を中央寄せ
    sheet.getRange(dataRow, 1, alerts.length, 1).setHorizontalAlignment('center');  // ASIN
    sheet.getRange(dataRow, 3, alerts.length, 1).setHorizontalAlignment('center');  // カテゴリ
  } else {
    sheet.getRange(startRow + 2, 1).setValue('✅ アラート対象なし');
  }
}



/**
 * 注意商品リスト
 */
function writeOverallSummary(sheet, totals, periods) {
  const t = totals.thisMonth;
  const lm = totals.lastMonthSameDay;
  const organicSales = t.sales - t.adSales;
  const fc = periods.daysInMonth / periods.elapsedDays;

  sheet.getRange(1, 1).setValue('━━━ Amazon事業 全体サマリー（当月）━━━').setFontWeight('bold').setFontSize(14);

  const headers = ['軸', '売上', '売上比', 'CV', '点数', '広告費', '販売手数料', '経費等', '利益', '原価率', '粗利率', '広告比率', 'ROAS', '利益率'];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center');

  const totalRow = ['全体', t.sales, 1, t.cv, t.units, t.adCost, t.commission, t.otherExpense, t.profit, t.costRate, t.grossMargin, t.adRate, t.roas, t.profitMargin];
  const adRow = ['広告', t.adSales, t.sales > 0 ? t.adSales/t.sales : 0, 0, 0, t.adCost, '-', '-', '-', '-', '-', '-', t.roas, '-'];
  const organicRow = ['オーガニック', organicSales, t.sales > 0 ? organicSales/t.sales : 0, t.cv, t.units, '-', '-', '-', '-', '-', '-', '-', '-', '-'];
  const pctRow = ['前月比',
    pctChangeNum(t.sales, lm.sales), '-',
    pctChangeNum(t.cv, lm.cv),
    pctChangeNum(t.units, lm.units),
    pctChangeNum(t.adCost, lm.adCost),
    pctChangeNum(t.commission, lm.commission),
    pctChangeNum(t.otherExpense, lm.otherExpense),
    pctChangeNum(t.profit, lm.profit),
    pctChangeNum(t.costRate, lm.costRate),
    pctChangeNum(t.grossMargin, lm.grossMargin),
    pctChangeNum(t.adRate, lm.adRate),
    pctChangeNum(t.roas, lm.roas),
    pctChangeNum(t.profitMargin, lm.profitMargin),
  ];
  const fcRow = ['月末予測', t.sales*fc, '-', t.cv*fc, t.units*fc, t.adCost*fc, t.commission*fc, t.otherExpense*fc, t.profit*fc, '-', '-', '-', '-', '-'];

  const rows = [totalRow, adRow, organicRow, pctRow, fcRow];
  sheet.getRange(3, 1, rows.length, headers.length).setValues(rows);

  sheet.getRange(3, 1, rows.length, 1).setHorizontalAlignment('center').setFontWeight('bold');
  sheet.getRange(3, 2, rows.length, headers.length - 1).setHorizontalAlignment('right');

  // 数値列: 売上(2), CV(4), 点数(5), 広告費(6), 販売手数料(7), 経費等(8), 利益(9)
  [2, 4, 5, 6, 7, 8, 9].forEach(col => {
    sheet.getRange(3, col, 3, 1).setNumberFormat('#,##0');
    sheet.getRange(7, col, 1, 1).setNumberFormat('#,##0');
  });
  // 率列: 売上比(3), 原価率(10), 粗利率(11), 広告比率(12), 利益率(14)
  [3, 10, 11, 12, 14].forEach(col => {
    sheet.getRange(3, col, 3, 1).setNumberFormat('0.0%');
  });
  // ROAS(13)
  sheet.getRange(3, 13, 3, 1).setNumberFormat('0.00');
  // 前月比行（6行目）
  sheet.getRange(6, 2, 1, headers.length - 1).setNumberFormat('+0.0%;-0.0%;-');
}





function writeCategorySummary(sheet, dailyData, periods, thisMonthExp, lastMonthSameDayExp) {
  const startRow = 10;
  sheet.getRange(startRow, 1).setValue('━━━ カテゴリ別サマリー（当月）━━━').setFontWeight('bold').setFontSize(14);

  const headers = ['カテゴリ', '売上', '売上比', 'CV', '点数', '広告費', '利益', 'TACOS', 'TACOS前月比', 'ACOS', 'ACOS前月比', 'ROAS', '利益率'];
  sheet.getRange(startRow + 1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center');

  const categories = [...new Set(dailyData.filter(d => d.category).map(d => d.category))].sort();
  const totalSales = dailyData
    .filter(d => d.date >= periods.thisMonth.start && d.date <= periods.thisMonth.end)
    .reduce((s, d) => s + d.sales, 0);

  const catData = categories.map(cat => {
    const agg = aggregateData(dailyData, periods.thisMonth.start, periods.thisMonth.end, d => d.category === cat, thisMonthExp);
    const lmAgg = aggregateData(dailyData, periods.lastMonthSameDay.start, periods.lastMonthSameDay.end, d => d.category === cat, lastMonthSameDayExp);
    return { cat, agg, lmAgg };
  });

  const filtered = catData.filter(c => c.agg.sales > 0 || c.agg.cv > 0)
    .sort((a, b) => b.agg.sales - a.agg.sales);

  const rows = filtered.map(c => [
    c.cat,
    c.agg.sales,
    totalSales > 0 ? c.agg.sales / totalSales : 0,
    c.agg.cv,
    c.agg.units,
    c.agg.adCost,
    c.agg.profit,
    c.agg.tacos / 100,
    pctChangeNum(c.agg.tacos, c.lmAgg.tacos),
    c.agg.acos / 100,
    pctChangeNum(c.agg.acos, c.lmAgg.acos),
    c.agg.roas,
    c.agg.profitMargin,
  ]);

  if (rows.length > 0) {
    const dataRow = startRow + 2;
    sheet.getRange(dataRow, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(dataRow, 1, rows.length, 1).setHorizontalAlignment('center');
    sheet.getRange(dataRow, 2, rows.length, headers.length - 1).setHorizontalAlignment('right');
    sheet.getRange(dataRow, 2, rows.length, 1).setNumberFormat('#,##0');  // 売上
    sheet.getRange(dataRow, 3, rows.length, 1).setNumberFormat('0.0%');   // 売上比
    sheet.getRange(dataRow, 4, rows.length, 2).setNumberFormat('#,##0');  // CV, 点数
    sheet.getRange(dataRow, 6, rows.length, 2).setNumberFormat('#,##0');  // 広告費, 利益
    sheet.getRange(dataRow, 8, rows.length, 1).setNumberFormat('0.0%');   // TACOS
    sheet.getRange(dataRow, 9, rows.length, 1).setNumberFormat('+0.0%;-0.0%;-');
    sheet.getRange(dataRow, 10, rows.length, 1).setNumberFormat('0.0%');  // ACOS
    sheet.getRange(dataRow, 11, rows.length, 1).setNumberFormat('+0.0%;-0.0%;-');
    sheet.getRange(dataRow, 12, rows.length, 1).setNumberFormat('0.00');  // ROAS
    sheet.getRange(dataRow, 13, rows.length, 1).setNumberFormat('0.0%');  // 利益率
  }

  return startRow + 2 + rows.length + 2;
}








function writeAlertProducts(sheet, dailyData, period, startRow) {
  sheet.getRange(startRow, 1).setValue('━━━ 注意が必要な商品 ━━━').setFontWeight('bold').setFontSize(14);
  sheet.getRange(startRow + 1, 1, 1, 4).setValues([['ASIN', '商品名', 'カテゴリ', '理由']])
    .setFontWeight('bold').setBackground('#fce8e6').setHorizontalAlignment('center');

  const asinMap = {};
  dailyData.filter(d => d.date >= period.start && d.date <= period.end).forEach(d => {
    if (!asinMap[d.asin]) {
      asinMap[d.asin] = { name: d.name, category: d.category, sales: 0, adCost: 0, adSales: 0 };
    }
    asinMap[d.asin].sales += d.sales;
    asinMap[d.asin].adCost += d.adCost;
    asinMap[d.asin].adSales += d.adSales;
  });

  const alerts = [];
  for (const [asin, d] of Object.entries(asinMap)) {
    const tacos = d.sales > 0 ? (d.adCost / d.sales * 100) : 0;
    if (tacos > 30) {
      alerts.push([asin, d.name, d.category, 'TACOS高: ' + tacos.toFixed(1) + '%']);
    }
    if (d.sales === 0 && d.adCost > 0) {
      alerts.push([asin, d.name, d.category, '広告費あり・売上ゼロ']);
    }
  }

  if (alerts.length > 0) {
    const dataRow = startRow + 2;
    sheet.getRange(dataRow, 1, alerts.length, 4).setValues(alerts);
    // ASIN・カテゴリ列を中央寄せ
    sheet.getRange(dataRow, 1, alerts.length, 1).setHorizontalAlignment('center');  // ASIN
    sheet.getRange(dataRow, 3, alerts.length, 1).setHorizontalAlignment('center');  // カテゴリ
  } else {
    sheet.getRange(startRow + 2, 1).setValue('✅ アラート対象なし');
  }
}




function pctChangeNum(current, prev) {
  if (!prev || prev === 0) return '';
  if (!current && current !== 0) return '';
  return (current - prev) / prev;
}
/**
 * L2 カテゴリ分析を更新
 * カテゴリドロップダウンで選択されたカテゴリの月次推移とASIN別を表示
 */
function updateDashboardL2() {
  const sheet = getOrCreateSheet(SHEET_NAMES.L2_CATEGORY);
  const dailyData = getDailyDataAll();
  const periods = getPeriods();

  sheet.clear();

  // タイトル
  sheet.getRange(1, 1).setValue('━━━ L2 カテゴリ分析 ━━━').setFontWeight('bold').setFontSize(14);
  sheet.getRange(1, 9).setValue('※ 当月売上上位順に表示').setFontStyle('italic').setFontColor('#888');

  // カテゴリを当月売上順でソート
  const categories = [...new Set(dailyData.filter(d => d.category).map(d => d.category))];
  const catWithSales = categories.map(cat => {
    const thisMonthSales = dailyData
      .filter(d => d.category === cat && d.date >= periods.thisMonth.start && d.date <= periods.thisMonth.end)
      .reduce((s, d) => s + d.sales, 0);
    return { cat, sales: thisMonthSales };
  }).filter(c => c.sales > 0)
    .sort((a, b) => b.sales - a.sales);

  let currentRow = 3;

  catWithSales.forEach(({ cat }) => {
    const catData = dailyData.filter(d => d.category === cat);
    const blockHeight = writeCategoryBlock(sheet, catData, cat, currentRow, periods);
    currentRow += blockHeight + 2; // スペース行
  });

  if (catWithSales.length === 0) {
    sheet.getRange(3, 1).setValue('当月売上のあるカテゴリがありません');
  }
}

/**
 * カテゴリブロックを書き込み（月次推移 + ASIN別を横並び）
 * @returns {number} 使用した行数
 */
function writeCategoryBlock(sheet, catData, category, startRow, periods) {
  // 左側: 月次推移
  sheet.getRange(startRow, 1).setValue('━━━ ' + category + ' 月次推移（直近12ヶ月）━━━')
    .setFontWeight('bold').setFontSize(12);
  sheet.getRange(startRow + 1, 1, 1, 7).setValues([['年月', '売上', 'CV', '点数', 'TACOS', 'ACOS', 'ROAS']])
    .setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center');

  const monthMap = {};
  catData.forEach(d => {
    const month = d.date.substring(0, 7);
    if (!monthMap[month]) {
      monthMap[month] = { sales: 0, cv: 0, units: 0, adCost: 0, adSales: 0 };
    }
    monthMap[month].sales += d.sales;
    monthMap[month].cv += d.cv;
    monthMap[month].units += d.units;
    monthMap[month].adCost += d.adCost;
    monthMap[month].adSales += d.adSales;
  });

  const months = Object.keys(monthMap).sort().slice(-12);
  const monthRows = months.map(m => {
    const d = monthMap[m];
    return [
      m, d.sales, d.cv, d.units,
      d.sales > 0 ? d.adCost / d.sales : 0,
      d.adSales > 0 ? d.adCost / d.adSales : 0,
      d.adCost > 0 ? d.sales / d.adCost : 0,
    ];
  });

  if (monthRows.length > 0) {
    const r = startRow + 2;
    sheet.getRange(r, 1, monthRows.length, 7).setValues(monthRows);
    sheet.getRange(r, 1, monthRows.length, 1).setHorizontalAlignment('center');
    sheet.getRange(r, 2, monthRows.length, 6).setHorizontalAlignment('right');
    sheet.getRange(r, 2, monthRows.length, 3).setNumberFormat('#,##0');
    sheet.getRange(r, 5, monthRows.length, 2).setNumberFormat('0.0%');
    sheet.getRange(r, 7, monthRows.length, 1).setNumberFormat('0.00');
  }

  // 右側: ASIN別（当月）I列から開始
  sheet.getRange(startRow, 9).setValue('━━━ ' + category + ' 内ASIN別（当月）━━━')
    .setFontWeight('bold').setFontSize(12);
  sheet.getRange(startRow + 1, 9, 1, 8).setValues([['ASIN', '商品名', '売上', '売上比', 'CV', '点数', 'TACOS', 'ACOS']])
    .setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center');

  const thisMonth = catData.filter(d => d.date >= periods.thisMonth.start && d.date <= periods.thisMonth.end);
  const asinMap = {};
  thisMonth.forEach(d => {
    if (!asinMap[d.asin]) {
      asinMap[d.asin] = { name: d.name, sales: 0, cv: 0, units: 0, adCost: 0, adSales: 0 };
    }
    asinMap[d.asin].sales += d.sales;
    asinMap[d.asin].cv += d.cv;
    asinMap[d.asin].units += d.units;
    asinMap[d.asin].adCost += d.adCost;
    asinMap[d.asin].adSales += d.adSales;
  });

  const catTotal = Object.values(asinMap).reduce((s, d) => s + d.sales, 0);

  const asinRows = Object.entries(asinMap)
    .sort((a, b) => b[1].sales - a[1].sales)
    .map(([asin, d]) => [
      asin, d.name, d.sales,
      catTotal > 0 ? d.sales / catTotal : 0,
      d.cv, d.units,
      d.sales > 0 ? d.adCost / d.sales : 0,
      d.adSales > 0 ? d.adCost / d.adSales : 0,
    ]);

  if (asinRows.length > 0) {
    const r = startRow + 2;
    sheet.getRange(r, 9, asinRows.length, 8).setValues(asinRows);
    sheet.getRange(r, 9, asinRows.length, 1).setHorizontalAlignment('center');
    sheet.getRange(r, 11, asinRows.length, 6).setHorizontalAlignment('right');
    sheet.getRange(r, 11, asinRows.length, 1).setNumberFormat('#,##0');
    sheet.getRange(r, 12, asinRows.length, 1).setNumberFormat('0.0%');
    sheet.getRange(r, 13, asinRows.length, 2).setNumberFormat('#,##0');
    sheet.getRange(r, 15, asinRows.length, 2).setNumberFormat('0.0%');
  }

  // 使用した行数 = max(月次行数, ASIN行数) + ヘッダー2行
  return 2 + Math.max(monthRows.length, asinRows.length);
}


/**
 * 月次推移を書き込み
 */
function writeMonthlyTrend(sheet, catData, startRow, category) {
  sheet.getRange(startRow, 1).setValue('━━━ ' + category + ' 月次推移（直近12ヶ月）━━━').setFontWeight('bold').setFontSize(12);
  sheet.getRange(startRow + 1, 1, 1, 7).setValues([['年月', '売上', 'CV', '点数', 'TACOS', 'ACOS', 'ROAS']])
    .setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center');

  // 月別に集計
  const monthMap = {};
  catData.forEach(d => {
    const month = d.date.substring(0, 7); // YYYY-MM
    if (!monthMap[month]) {
      monthMap[month] = { sales: 0, cv: 0, units: 0, adCost: 0, adSales: 0 };
    }
    monthMap[month].sales += d.sales;
    monthMap[month].cv += d.cv;
    monthMap[month].units += d.units;
    monthMap[month].adCost += d.adCost;
    monthMap[month].adSales += d.adSales;
  });

  // 直近12ヶ月を昇順でソート
  const months = Object.keys(monthMap).sort().slice(-12);

  const rows = months.map(m => {
    const d = monthMap[m];
    return [
      m,
      d.sales,
      d.cv,
      d.units,
      d.sales > 0 ? d.adCost / d.sales : 0,
      d.adSales > 0 ? d.adCost / d.adSales : 0,
      d.adCost > 0 ? d.sales / d.adCost : 0,
    ];
  });

  if (rows.length > 0) {
    const dataRow = startRow + 2;
    sheet.getRange(dataRow, 1, rows.length, 7).setValues(rows);
    sheet.getRange(dataRow, 1, rows.length, 1).setHorizontalAlignment('center');
    sheet.getRange(dataRow, 2, rows.length, 6).setHorizontalAlignment('right');
    sheet.getRange(dataRow, 2, rows.length, 3).setNumberFormat('#,##0');
    sheet.getRange(dataRow, 5, rows.length, 2).setNumberFormat('0.0%');
    sheet.getRange(dataRow, 7, rows.length, 1).setNumberFormat('0.00');
  }
}

/**
 * カテゴリ内ASIN別（当月）
 */
function writeCategoryAsins(sheet, catData, startRow, category) {
  sheet.getRange(startRow, 1).setValue('━━━ ' + category + ' 内ASIN別（当月）━━━').setFontWeight('bold').setFontSize(12);
  sheet.getRange(startRow + 1, 1, 1, 8).setValues([['ASIN', '商品名', '売上', '売上比', 'CV', '点数', 'TACOS', 'ACOS']])
    .setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center');

  const periods = getPeriods();
  const thisMonth = catData.filter(d => d.date >= periods.thisMonth.start && d.date <= periods.thisMonth.end);

  const asinMap = {};
  thisMonth.forEach(d => {
    if (!asinMap[d.asin]) {
      asinMap[d.asin] = { name: d.name, sales: 0, cv: 0, units: 0, adCost: 0, adSales: 0 };
    }
    asinMap[d.asin].sales += d.sales;
    asinMap[d.asin].cv += d.cv;
    asinMap[d.asin].units += d.units;
    asinMap[d.asin].adCost += d.adCost;
    asinMap[d.asin].adSales += d.adSales;
  });

  const catTotalSales = Object.values(asinMap).reduce((s, d) => s + d.sales, 0);

  const rows = Object.entries(asinMap)
    .sort((a, b) => b[1].sales - a[1].sales)
    .map(([asin, d]) => [
      asin,
      d.name,
      d.sales,
      catTotalSales > 0 ? d.sales / catTotalSales : 0,
      d.cv,
      d.units,
      d.sales > 0 ? d.adCost / d.sales : 0,
      d.adSales > 0 ? d.adCost / d.adSales : 0,
    ]);

  if (rows.length > 0) {
    const dataRow = startRow + 2;
    sheet.getRange(dataRow, 1, rows.length, 8).setValues(rows);
    sheet.getRange(dataRow, 1, rows.length, 1).setHorizontalAlignment('center');
    sheet.getRange(dataRow, 3, rows.length, 6).setHorizontalAlignment('right');
    sheet.getRange(dataRow, 3, rows.length, 1).setNumberFormat('#,##0');
    sheet.getRange(dataRow, 4, rows.length, 1).setNumberFormat('0.0%');
    sheet.getRange(dataRow, 5, rows.length, 2).setNumberFormat('#,##0');
    sheet.getRange(dataRow, 7, rows.length, 2).setNumberFormat('0.0%');
  } else {
    sheet.getRange(startRow + 2, 1).setValue('データなし');
  }
}

/**
 * Settlement Report から期間内の経費を集計
 * @returns {Object} { total, byAsin }
 */
function getSettlementExpenses(startDate, endDate) {
  const sheet = getOrCreateSheet(SHEET_NAMES.D2_SETTLEMENT);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { commission: 0, other: 0, total: 0, byAsin: {} };

  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  let commission = 0;
  let other = 0;
  const byAsin = {};

  data.forEach(row => {
    const postedDate = row[2] instanceof Date
      ? Utilities.formatDate(row[2], 'Asia/Tokyo', 'yyyy-MM-dd')
      : String(row[2]).substring(0, 10);
    if (!postedDate || postedDate < startDate || postedDate > endDate) return;

    const asin = String(row[3] || '').trim();
    const itemType = String(row[5]).trim();
    const amount = parseFloat(row[6]) || 0;

    if (itemType === 'Principal' || itemType === 'Tax') return;

    const expense = -amount;

    if (itemType === 'Commission') {
      commission += expense;
      if (asin) {
        if (!byAsin[asin]) byAsin[asin] = { commission: 0, other: 0 };
        byAsin[asin].commission += expense;
      }
    } else {
      other += expense;
      if (asin) {
        if (!byAsin[asin]) byAsin[asin] = { commission: 0, other: 0 };
        byAsin[asin].other += expense;
      }
    }
  });

  return { commission, other, total: commission + other, byAsin };
}




