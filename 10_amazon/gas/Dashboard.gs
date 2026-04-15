/**
 * Amazon Dashboard - L1/L2 ダッシュボード構築（最適化版）
 */

// ===== L1 事業ダッシュボード =====

function updateDashboardL1() {
  const t0 = Date.now();
  Logger.log('===== L1 事業ダッシュボード更新開始 =====');

  const sheet = getOrCreateSheet(SHEET_NAMES.L1_DASHBOARD);
  sheet.clear();

  // データ読み込み（シートアクセスは最小回数）
  const t1 = Date.now();
  const dailyData = getDailyDataAll();
  Logger.log('日次データ読み込み: ' + (Date.now()-t1) + 'ms (' + dailyData.length + '行)');

  if (dailyData.length === 0) { Logger.log('データなし'); return; }

  const periods = getPeriods();

  // 経費データは2期間まとめて1回だけ読む
  const t2 = Date.now();
  const allExpenses = readAllSettlement();
  Logger.log('経費データ読み込み: ' + (Date.now()-t2) + 'ms (' + allExpenses.length + '行)');

  // 経費を期間別に集計（1パス）
  const thisMonthExp = aggregateExpenses(allExpenses, periods.thisMonth.start, periods.thisMonth.end);
  const lastMonthSameDayExp = aggregateExpenses(allExpenses, periods.lastMonthSameDay.start, periods.lastMonthSameDay.end);

  // 日次データを期間別に事前フィルタ
  const thisMonthDaily = dailyData.filter(d => d.date >= periods.thisMonth.start && d.date <= periods.thisMonth.end);
  const lastMonthDaily = dailyData.filter(d => d.date >= periods.lastMonthSameDay.start && d.date <= periods.lastMonthSameDay.end);

  // カテゴリ別の集計を1パスで作る
  const t3 = Date.now();
  const thisMonthByCategory = aggregateByCategory(thisMonthDaily, thisMonthExp);
  const lastMonthByCategory = aggregateByCategory(lastMonthDaily, lastMonthSameDayExp);
  Logger.log('カテゴリ集計: ' + (Date.now()-t3) + 'ms');

  // 全体合計
  const totals = {
    thisMonth: sumCategoryAggs(thisMonthByCategory),
    lastMonthSameDay: sumCategoryAggs(lastMonthByCategory),
  };

  // シート書き込み
  const t4 = Date.now();
  const summaryEndRow = writeOverallSummary(sheet, totals, periods);
  const finalProfitEndRow = writeFinalProfitSection(sheet, totals, periods, summaryEndRow);
  const categoryResult = writeCategorySummary(sheet, thisMonthByCategory, lastMonthByCategory, periods, finalProfitEndRow);
  writeAlertProducts(sheet, thisMonthDaily, categoryResult.nextRow);
  Logger.log('シート書き込み: ' + (Date.now()-t4) + 'ms');

  // 条件付き書式（粗利率/TACOS/ACOS/ROAS のハイライト）
  const t5 = Date.now();
  applyL1ConditionalFormatting(sheet, categoryResult.dataStartRow, categoryResult.rowCount);
  Logger.log('条件付き書式: ' + (Date.now()-t5) + 'ms');

  Logger.log('===== L1 完了（合計 ' + (Date.now()-t0) + 'ms）=====');
}

// ===== データ取得 =====

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

  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  return data.map(row => ({
    asin: String(row[0] || '').trim(),
    yearMonth: String(row[1] || '').trim(),
    commission: parseFloat(row[2]) || 0,
    other: parseFloat(row[3]) || 0,
  }));
}

/**
 * 月次集計から期間でフィルタして集計（月単位の精度）
 */
function aggregateExpenses(allExpenses, startDate, endDate) {
  const startMonth = startDate.substring(0, 7);
  const endMonth = endDate.substring(0, 7);

  let commission = 0, other = 0;
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

    commission += c;
    other += o;

    if (row.asin) {
      if (!byAsin[row.asin]) byAsin[row.asin] = { commission: 0, other: 0 };
      byAsin[row.asin].commission += c;
      byAsin[row.asin].other += o;
    }
  }

  return { commission, other, total: commission + other, byAsin };
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
        adCost: 0, adSales: 0,
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
  const cogs = 0; // TODO: CFシート連携
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

function sumCategoryAggs(byCategory) {
  const total = {
    sales: 0, cv: 0, units: 0, sessions: 0, pv: 0,
    adCost: 0, adSales: 0, commission: 0, otherExpense: 0,
  };
  for (const cat of Object.values(byCategory)) {
    total.sales += cat.sales;
    total.cv += cat.cv;
    total.units += cat.units;
    total.sessions += cat.sessions;
    total.pv += cat.pv;
    total.adCost += cat.adCost;
    total.adSales += cat.adSales;
    total.commission += cat.commission;
    total.otherExpense += cat.otherExpense;
  }
  Object.assign(total, computeDerivedMetrics(total, total.commission, total.otherExpense));
  return total;
}

// ===== 全体サマリー =====

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

  [2, 4, 5, 6, 7, 8, 9].forEach(col => {
    sheet.getRange(3, col, 3, 1).setNumberFormat('#,##0');
    sheet.getRange(7, col, 1, 1).setNumberFormat('#,##0');
  });
  [3, 10, 11, 12, 14].forEach(col => {
    sheet.getRange(3, col, 3, 1).setNumberFormat('0.0%');
  });
  sheet.getRange(3, 13, 3, 1).setNumberFormat('0.00');
  sheet.getRange(6, 2, 1, headers.length - 1).setNumberFormat('+0.0%;-0.0%;-');

  // 全体サマリー末尾の行番号を返す（3 + 5行データ = 7行目まで使用）
  return 8;
}

// ===== 最終利益セクション（レイヤー2） =====

function writeFinalProfitSection(sheet, totals, periods, startRow) {
  const t = totals.thisMonth;
  const lm = totals.lastMonthSameDay;

  // 当月と前月同日の販促費を計算
  const thisPromo = calcPromoCostForPeriod(periods.thisMonth.start, periods.thisMonth.end, t.adCost);
  const lastPromo = calcPromoCostForPeriod(periods.lastMonthSameDay.start, periods.lastMonthSameDay.end, lm.adCost);

  // 最終利益 = Amazon内粗利 − 販促費合計
  // 注意: t.profit は既に広告費も引かれている = Amazon内粗利
  const thisFinalProfit = t.profit - thisPromo.total;
  const lastFinalProfit = lm.profit - lastPromo.total;
  const thisFinalMargin = t.sales > 0 ? thisFinalProfit / t.sales : 0;
  const lastFinalMargin = lm.sales > 0 ? lastFinalProfit / lm.sales : 0;

  sheet.getRange(startRow, 1).setValue('━━━ 最終利益（レイヤー2: 販促費控除後）━━━')
    .setFontWeight('bold').setFontSize(14);

  const headers = ['項目', '当月', '前月同日', '前月比'];
  sheet.getRange(startRow + 1, 1, 1, 4).setValues([headers])
    .setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center');

  const rows = [
    ['Amazon内粗利（①）',  t.profit,              lm.profit,              pctChangeNum(t.profit, lm.profit)],
    ['  Amence',           -thisPromo.amence,     -lastPromo.amence,      pctChangeNum(thisPromo.amence, lastPromo.amence)],
    ['  その他ツール',      -thisPromo.otherTools, -lastPromo.otherTools,  pctChangeNum(thisPromo.otherTools, lastPromo.otherTools)],
    ['  荷造運賃',         -thisPromo.shipping,   -lastPromo.shipping,    pctChangeNum(thisPromo.shipping, lastPromo.shipping)],
    ['  納品人件費',        -thisPromo.labor,      -lastPromo.labor,       pctChangeNum(thisPromo.labor, lastPromo.labor)],
    ['  M19成果報酬(6.0%)', -thisPromo.m19Performance, -lastPromo.m19Performance, pctChangeNum(thisPromo.m19Performance, lastPromo.m19Performance)],
    ['販促費合計（②）',    -thisPromo.total,      -lastPromo.total,       pctChangeNum(thisPromo.total, lastPromo.total)],
    ['最終利益（①−②）',    thisFinalProfit,       lastFinalProfit,        pctChangeNum(thisFinalProfit, lastFinalProfit)],
    ['最終利益率',          thisFinalMargin,       lastFinalMargin,        pctChangeNum(thisFinalMargin, lastFinalMargin)],
  ];

  const dataStartRow = startRow + 2;
  sheet.getRange(dataStartRow, 1, rows.length, 4).setValues(rows);

  // フォーマット
  sheet.getRange(dataStartRow, 2, rows.length, 2).setHorizontalAlignment('right');
  // 金額行（1〜8行目、最終利益率除く）
  sheet.getRange(dataStartRow, 2, rows.length - 1, 2).setNumberFormat('#,##0');
  // 最終利益率行
  sheet.getRange(dataStartRow + rows.length - 1, 2, 1, 2).setNumberFormat('0.0%');
  // 前月比（D列）
  sheet.getRange(dataStartRow, 4, rows.length, 1).setNumberFormat('+0.0%;-0.0%;-');

  // 太字にする行: Amazon内粗利（①）、販促費合計（②）、最終利益、最終利益率
  sheet.getRange(dataStartRow, 1, 1, 4).setFontWeight('bold');           // ①
  sheet.getRange(dataStartRow + 6, 1, 1, 4).setFontWeight('bold');        // ②
  sheet.getRange(dataStartRow + 7, 1, 2, 4).setFontWeight('bold');        // 最終利益 + 最終利益率
  sheet.getRange(dataStartRow + 7, 1, 2, 4).setBackground('#fff2cc');     // 最終利益行を薄黄色で強調

  return dataStartRow + rows.length + 2; // セクション末尾（次のセクション開始行）
}

// ===== カテゴリ別サマリー =====

function writeCategorySummary(sheet, thisMonthByCategory, lastMonthByCategory, periods, startRow) {
  if (!startRow) startRow = 10; // デフォルト（後方互換）
  sheet.getRange(startRow, 1).setValue('━━━ カテゴリ別サマリー（当月）━━━').setFontWeight('bold').setFontSize(14);

  const headers = ['カテゴリ', '売上', '売上比', 'CV', '点数', '広告費', '利益', 'TACOS', 'TACOS前月比', 'ACOS', 'ACOS前月比', 'ROAS', '利益率'];
  sheet.getRange(startRow + 1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center');

  const totalSales = Object.values(thisMonthByCategory).reduce((s, c) => s + c.sales, 0);

  const filtered = Object.values(thisMonthByCategory)
    .filter(c => c.sales > 0 || c.cv > 0)
    .sort((a, b) => b.sales - a.sales);

  const rows = filtered.map(c => {
    const lm = lastMonthByCategory[c.category] || { tacos: 0, acos: 0 };
    return [
      c.category,
      c.sales,
      totalSales > 0 ? c.sales / totalSales : 0,
      c.cv,
      c.units,
      c.adCost,
      c.profit,
      c.tacos / 100,
      pctChangeNum(c.tacos, lm.tacos),
      c.acos / 100,
      pctChangeNum(c.acos, lm.acos),
      c.roas,
      c.profitMargin,
    ];
  });

  let dataStartRow = startRow + 2;
  if (rows.length > 0) {
    sheet.getRange(dataStartRow, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(dataStartRow, 1, rows.length, 1).setHorizontalAlignment('center');
    sheet.getRange(dataStartRow, 2, rows.length, headers.length - 1).setHorizontalAlignment('right');
    sheet.getRange(dataStartRow, 2, rows.length, 1).setNumberFormat('#,##0');
    sheet.getRange(dataStartRow, 3, rows.length, 1).setNumberFormat('0.0%');
    sheet.getRange(dataStartRow, 4, rows.length, 2).setNumberFormat('#,##0');
    sheet.getRange(dataStartRow, 6, rows.length, 2).setNumberFormat('#,##0');
    sheet.getRange(dataStartRow, 8, rows.length, 1).setNumberFormat('0.0%');
    sheet.getRange(dataStartRow, 9, rows.length, 1).setNumberFormat('+0.0%;-0.0%;-');
    sheet.getRange(dataStartRow, 10, rows.length, 1).setNumberFormat('0.0%');
    sheet.getRange(dataStartRow, 11, rows.length, 1).setNumberFormat('+0.0%;-0.0%;-');
    sheet.getRange(dataStartRow, 12, rows.length, 1).setNumberFormat('0.00');
    sheet.getRange(dataStartRow, 13, rows.length, 1).setNumberFormat('0.0%');
  }

  return {
    nextRow: startRow + 2 + rows.length + 2,
    dataStartRow: dataStartRow,
    rowCount: rows.length,
  };
}

// ===== 注意商品 =====

function writeAlertProducts(sheet, thisMonthDaily, startRow) {
  sheet.getRange(startRow, 1).setValue('━━━ 注意が必要な商品 ━━━').setFontWeight('bold').setFontSize(14);
  sheet.getRange(startRow + 1, 1, 1, 4).setValues([['ASIN', '商品名', 'カテゴリ', '理由']])
    .setFontWeight('bold').setBackground('#fce8e6').setHorizontalAlignment('center');

  const asinMap = {};
  for (const d of thisMonthDaily) {
    if (!asinMap[d.asin]) {
      asinMap[d.asin] = { name: d.name, category: d.category, sales: 0, adCost: 0, adSales: 0 };
    }
    asinMap[d.asin].sales += d.sales;
    asinMap[d.asin].adCost += d.adCost;
    asinMap[d.asin].adSales += d.adSales;
  }

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
    sheet.getRange(dataRow, 1, alerts.length, 1).setHorizontalAlignment('center');
    sheet.getRange(dataRow, 3, alerts.length, 1).setHorizontalAlignment('center');
  } else {
    sheet.getRange(startRow + 2, 1).setValue('✅ アラート対象なし');
  }
}

// ===== L2 カテゴリ分析 =====

function updateDashboardL2() {
  const t0 = Date.now();
  const sheet = getOrCreateSheet(SHEET_NAMES.L2_CATEGORY);
  const dailyData = getDailyDataAll();
  const periods = getPeriods();

  // 経費データを読み込み（高速化用）
  const allExpenses = readAllSettlement();

  sheet.clear();

  sheet.getRange(1, 1).setValue('━━━ L2 カテゴリ分析 ━━━').setFontWeight('bold').setFontSize(14);
  sheet.getRange(1, 11).setValue('※ 当月売上上位順に表示').setFontStyle('italic').setFontColor('#888');

  // 当月売上順でカテゴリをソート
  const thisMonthDaily = dailyData.filter(d => d.date >= periods.thisMonth.start && d.date <= periods.thisMonth.end);
  const catSalesMap = {};
  for (const d of thisMonthDaily) {
    if (!d.category) continue;
    catSalesMap[d.category] = (catSalesMap[d.category] || 0) + d.sales;
  }

  const sortedCategories = Object.entries(catSalesMap)
    .filter(([, sales]) => sales > 0)
    .sort((a, b) => b[1] - a[1])
    .map(([cat]) => cat);

  // カテゴリごとにデータを事前分類
  const dataByCategory = {};
  for (const d of dailyData) {
    if (!d.category) continue;
    if (!dataByCategory[d.category]) dataByCategory[d.category] = [];
    dataByCategory[d.category].push(d);
  }

  const blocks = [];
  let currentRow = 3;
  sortedCategories.forEach(cat => {
    const catData = dataByCategory[cat] || [];
    const result = writeCategoryBlock(sheet, catData, cat, currentRow, periods, allExpenses);
    blocks.push({
      startRow: currentRow,
      monthRowCount: result.monthRowCount,
      asinRowCount: result.asinRowCount,
    });
    currentRow += result.blockHeight + 2;
  });

  if (sortedCategories.length === 0) {
    sheet.getRange(3, 1).setValue('当月売上のあるカテゴリがありません');
  }

  // 条件付き書式（粗利率/TACOS/ACOS/ROAS のハイライト）
  if (blocks.length > 0) {
    const t5 = Date.now();
    applyL2ConditionalFormatting(sheet, blocks);
    Logger.log('条件付き書式: ' + (Date.now()-t5) + 'ms');
  }

  Logger.log('L2 完了（' + (Date.now()-t0) + 'ms）');
}

function writeCategoryBlock(sheet, catData, category, startRow, periods, allExpenses) {
  // 左側: 月次推移（10列: 年月, 売上, CV, 点数, 広告費, 利益, TACOS, ACOS, ROAS, 利益率）
  sheet.getRange(startRow, 1).setValue('━━━ ' + category + ' 月次推移（直近12ヶ月）━━━')
    .setFontWeight('bold').setFontSize(12);
  sheet.getRange(startRow + 1, 1, 1, 10).setValues([['年月', '売上', 'CV', '点数', '広告費', '利益', 'TACOS', 'ACOS', 'ROAS', '利益率']])
    .setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center');

  // 月別集計（月ごとに ASIN セットも記録）
  const monthMap = {};
  for (const d of catData) {
    const month = d.date.substring(0, 7);
    if (!monthMap[month]) {
      monthMap[month] = { sales: 0, cv: 0, units: 0, adCost: 0, adSales: 0, asins: new Set() };
    }
    monthMap[month].sales += d.sales;
    monthMap[month].cv += d.cv;
    monthMap[month].units += d.units;
    monthMap[month].adCost += d.adCost;
    monthMap[month].adSales += d.adSales;
    if (d.asin) monthMap[month].asins.add(d.asin);
  }

  // 月ごとに経費を集計
  const expenseByMonthAsin = {};
  for (const exp of allExpenses) {
    if (!expenseByMonthAsin[exp.yearMonth]) expenseByMonthAsin[exp.yearMonth] = {};
    expenseByMonthAsin[exp.yearMonth][exp.asin] = exp;
  }

  const months = Object.keys(monthMap).sort().slice(-12);
  const monthRows = months.map(m => {
    const d = monthMap[m];
    let commission = 0, otherExpense = 0;
    const monthExp = expenseByMonthAsin[m] || {};
    for (const asin of d.asins) {
      if (monthExp[asin]) {
        commission += monthExp[asin].commission || 0;
        otherExpense += monthExp[asin].other || 0;
      }
    }
    const profit = d.sales - commission - otherExpense - d.adCost;
    return [
      m, d.sales, d.cv, d.units, d.adCost, profit,
      d.sales > 0 ? d.adCost / d.sales : 0,
      d.adSales > 0 ? d.adCost / d.adSales : 0,
      d.adCost > 0 ? d.sales / d.adCost : 0,
      d.sales > 0 ? profit / d.sales : 0,
    ];
  });

  if (monthRows.length > 0) {
    const r = startRow + 2;
    sheet.getRange(r, 1, monthRows.length, 10).setValues(monthRows);
    sheet.getRange(r, 1, monthRows.length, 1).setHorizontalAlignment('center');
    sheet.getRange(r, 2, monthRows.length, 9).setHorizontalAlignment('right');
    sheet.getRange(r, 2, monthRows.length, 5).setNumberFormat('#,##0');  // 売上,CV,点数,広告費,利益
    sheet.getRange(r, 7, monthRows.length, 2).setNumberFormat('0.0%');   // TACOS,ACOS
    sheet.getRange(r, 9, monthRows.length, 1).setNumberFormat('0.00');   // ROAS
    sheet.getRange(r, 10, monthRows.length, 1).setNumberFormat('0.0%');  // 利益率
  }

  // 右側: ASIN別（11列目以降）
  const rightCol = 12;
  sheet.getRange(startRow, rightCol).setValue('━━━ ' + category + ' 内ASIN別（当月）━━━')
    .setFontWeight('bold').setFontSize(12);
  sheet.getRange(startRow + 1, rightCol, 1, 11).setValues([['ASIN', '商品名', '売上', '売上比', 'CV', '点数', '広告費', '利益', 'TACOS', 'ACOS', '利益率']])
    .setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center');

  const asinMap = {};
  for (const d of catData) {
    if (d.date < periods.thisMonth.start || d.date > periods.thisMonth.end) continue;
    if (!asinMap[d.asin]) {
      asinMap[d.asin] = { name: d.name, sales: 0, cv: 0, units: 0, adCost: 0, adSales: 0 };
    }
    asinMap[d.asin].sales += d.sales;
    asinMap[d.asin].cv += d.cv;
    asinMap[d.asin].units += d.units;
    asinMap[d.asin].adCost += d.adCost;
    asinMap[d.asin].adSales += d.adSales;
  }

  // ASIN別の経費を当月分から取得
  const thisMonthYM = periods.thisMonth.start.substring(0, 7);
  const thisMonthExpByAsin = expenseByMonthAsin[thisMonthYM] || {};

  const catTotal = Object.values(asinMap).reduce((s, d) => s + d.sales, 0);

  const asinRows = Object.entries(asinMap)
    .sort((a, b) => b[1].sales - a[1].sales)
    .map(([asin, d]) => {
      const exp = thisMonthExpByAsin[asin] || { commission: 0, other: 0 };
      const profit = d.sales - exp.commission - exp.other - d.adCost;
      return [
        asin, d.name, d.sales,
        catTotal > 0 ? d.sales / catTotal : 0,
        d.cv, d.units, d.adCost, profit,
        d.sales > 0 ? d.adCost / d.sales : 0,
        d.adSales > 0 ? d.adCost / d.adSales : 0,
        d.sales > 0 ? profit / d.sales : 0,
      ];
    });

  if (asinRows.length > 0) {
    const r = startRow + 2;
    sheet.getRange(r, rightCol, asinRows.length, 11).setValues(asinRows);
    sheet.getRange(r, rightCol, asinRows.length, 1).setHorizontalAlignment('center');
    sheet.getRange(r, rightCol + 2, asinRows.length, 9).setHorizontalAlignment('right');
    sheet.getRange(r, rightCol + 2, asinRows.length, 1).setNumberFormat('#,##0');     // 売上
    sheet.getRange(r, rightCol + 3, asinRows.length, 1).setNumberFormat('0.0%');      // 売上比
    sheet.getRange(r, rightCol + 4, asinRows.length, 4).setNumberFormat('#,##0');     // CV,点数,広告費,利益
    sheet.getRange(r, rightCol + 8, asinRows.length, 2).setNumberFormat('0.0%');      // TACOS,ACOS
    sheet.getRange(r, rightCol + 10, asinRows.length, 1).setNumberFormat('0.0%');     // 利益率
  }

  return {
    blockHeight: 2 + Math.max(monthRows.length, asinRows.length),
    monthRowCount: monthRows.length,
    asinRowCount: asinRows.length,
  };
}

// ===== ヘルパー =====

function pctChangeNum(current, prev) {
  if (!prev || prev === 0) return '';
  if (!current && current !== 0) return '';
  return (current - prev) / prev;
}
