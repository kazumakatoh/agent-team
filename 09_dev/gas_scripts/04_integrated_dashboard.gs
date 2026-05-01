/**
 * 統合ダッシュボード構築
 * 2026/4 〜 2030/12（57ヶ月）
 * - 元本：取引所への入出金で変動（手入力 or 入金履歴）
 * - 評価額：現物評価額＋FX Balance（毎月キャリー）
 * - 月次利益＝当月評価額−前月評価額
 * - 月次利回り＝月次利益÷元本
 * - 累計利回り＝月次利回りの累計（gross 純利益÷元本）
 */

const DASH_CONFIG = {
  SHEET_MAIN: '統合ダッシュボード',
  SHEET_RAW: 'raw_USD',
  START_YEAR: 2026,
  START_MONTH: 4,
  END_YEAR: 2030,
  END_MONTH: 12,
  DEFAULT_RATE: 155,
  FIRST_DATA_COL: 2
};

const KPI_ROWS = [
  { type: 'header', label: '💰 統合（合計）' },
  { type: 'data', label: '初期費用', fmt: 'money' },
  { type: 'data', label: '元本', fmt: 'money' },
  { type: 'data', label: '評価額', fmt: 'money' },
  { type: 'data', label: '純利益', fmt: 'money' },
  { type: 'data', label: '月次利回り(%)', fmt: 'percent' },
  { type: 'data', label: '経費', fmt: 'money' },
  { type: 'data', label: '累計利益', fmt: 'money' },
  { type: 'data', label: '累計利回り(%)', fmt: 'percent' },
  { type: 'blank' },
  { type: 'header', label: '📈 現物（AIグリッド + AIバスケット）' },
  { type: 'data', label: '元本', fmt: 'money' },
  { type: 'data', label: '評価額', fmt: 'money' },
  { type: 'data', label: '月次利益', fmt: 'money' },
  { type: 'data', label: '月次利回り(%)', fmt: 'percent' },
  { type: 'data', label: '累計利益', fmt: 'money' },
  { type: 'data', label: '累計利回り(%)', fmt: 'percent' },
  { type: 'data', label: '手数料', fmt: 'money' },
  { type: 'blank' },
  { type: 'header', label: '💱 FX(BigBoss)' },
  { type: 'data', label: '元本', fmt: 'money' },
  { type: 'data', label: 'Balance', fmt: 'money' },
  { type: 'data', label: '月次利益', fmt: 'money' },
  { type: 'data', label: '月次利回り(%)', fmt: 'percent' },
  { type: 'data', label: '累計利益', fmt: 'money' },
  { type: 'data', label: '累計利回り(%)', fmt: 'percent' },
  { type: 'data', label: '取引回数', fmt: 'int' },
  { type: 'data', label: '勝率(%)', fmt: 'percent' },
  { type: 'data', label: '損益分岐勝率(%)', fmt: 'percent' },
  { type: 'data', label: 'PF', fmt: 'decimal' },
  { type: 'data', label: 'RR比', fmt: 'decimal' },
  { type: 'data', label: '損益平均/回', fmt: 'money' },
  { type: 'data', label: '最大DD(%)', fmt: 'percent' },
  { type: 'data', label: '取引コスト', fmt: 'money' },
  { type: 'data', label: '平均ロット数', fmt: 'decimal' },
  { type: 'data', label: '必要証拠金', fmt: 'money' }
];

function initIntegratedDashboard() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(DASH_CONFIG.SHEET_MAIN);
  if (!sheet) sheet = ss.insertSheet(DASH_CONFIG.SHEET_MAIN);
  else sheet.clear();

  const months = [];
  for (let y = DASH_CONFIG.START_YEAR; y <= DASH_CONFIG.END_YEAR; y++) {
    const s = (y === DASH_CONFIG.START_YEAR) ? DASH_CONFIG.START_MONTH : 1;
    const e = (y === DASH_CONFIG.END_YEAR) ? DASH_CONFIG.END_MONTH : 12;
    for (let m = s; m <= e; m++) months.push({ year: y, month: m });
  }
  const lastCol = DASH_CONFIG.FIRST_DATA_COL + months.length - 1;

  sheet.getRange('A1').setValue('通貨:').setFontWeight('bold');
  sheet.getRange('B1').setValue('USD');
  sheet.getRange('B1').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['USD', 'JPY'], true).build()
  );
  sheet.getRange('B1').setBackground('#fce5cd').setFontWeight('bold');
  sheet.getRange('C1').setValue('為替(¥/$):').setFontWeight('bold');
  sheet.getRange('D1').setValue(DASH_CONFIG.DEFAULT_RATE);
  sheet.getRange('D1').setBackground('#fce5cd').setNumberFormat('0.00');
  sheet.getRange('E1').setValue('最終更新:').setFontWeight('bold');
  sheet.getRange('F1').setValue(new Date()).setNumberFormat('yyyy/MM/dd HH:mm:ss');

  let col = DASH_CONFIG.FIRST_DATA_COL;
  for (let y = DASH_CONFIG.START_YEAR; y <= DASH_CONFIG.END_YEAR; y++) {
    const s = (y === DASH_CONFIG.START_YEAR) ? DASH_CONFIG.START_MONTH : 1;
    const e = (y === DASH_CONFIG.END_YEAR) ? DASH_CONFIG.END_MONTH : 12;
    const count = e - s + 1;
    const cell = sheet.getRange(3, col);
    cell.setValue(y + '年').setHorizontalAlignment('center')
      .setFontWeight('bold').setBackground('#1c4587').setFontColor('white');
    if (count > 1) sheet.getRange(3, col, 1, count).merge()
      .setHorizontalAlignment('center').setBackground('#1c4587').setFontColor('white').setFontWeight('bold');
    col += count;
  }

  const monthLabels = months.map(m => m.month + '月');
  sheet.getRange(4, DASH_CONFIG.FIRST_DATA_COL, 1, months.length)
    .setValues([monthLabels])
    .setHorizontalAlignment('center').setFontWeight('bold').setBackground('#cfe2f3');

  const DATA_START = 5;
  KPI_ROWS.forEach((kpi, i) => {
    const row = DATA_START + i;
    if (kpi.type === 'header') {
      sheet.getRange(row, 1, 1, lastCol)
        .setBackground('#37474f').setFontColor('white').setFontWeight('bold').setFontSize(11);
      sheet.getRange(row, 1).setValue(kpi.label);
    } else if (kpi.type === 'data') {
      sheet.getRange(row, 1).setValue(kpi.label).setFontWeight('bold').setBackground('#f3f3f3');
    }
  });

  const totalRows = DATA_START + KPI_ROWS.length - 1;
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidths(DASH_CONFIG.FIRST_DATA_COL, months.length, 90);
  sheet.getRange(1, 1, totalRows, lastCol).setFontSize(10);
  sheet.setFrozenRows(4);
  sheet.setFrozenColumns(1);

  Logger.log(`✅ 統合ダッシュボード初期化完了（${months.length}ヶ月分・${lastCol}列）`);
}

const DASH_ASSUMPTIONS = {
  INITIAL_SPOT_USD: 2972,
  INITIAL_FX_USD: 6253.51,
  SPOT_MONTHLY_YIELD: 0.10,
  FX_MONTHLY_YIELD: 0.05,
  FX_MONTHLY_TRADES: 60,
  FX_WIN_RATE: 0.70,
  SAGEMASTER_MONTHLY: 149,
  INITIAL_SETUP_USD: 2366,
  INITIAL_SETUP_YEAR: 2026,
  INITIAL_SETUP_MONTH: 4
};

function populateIntegratedDashboard() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(DASH_CONFIG.SHEET_MAIN);
  if (!sheet) throw new Error('統合ダッシュボードが未作成です');

  const spot = readSpotForDashboard_(ss);
  const fx = readFXForDashboard_(ss);
  const months = generateMonthList_();
  const currentMonthIdx = findCurrentMonthIndex_(months);

  let cumulativeDeposits = {};
  try { cumulativeDeposits = getCumulativePrincipalByMonth_(); } catch (e) {}

  const ROW = {
    TOTAL_SETUP: 6, TOTAL_PRINCIPAL: 7, TOTAL_VALUE: 8, TOTAL_NET: 9, TOTAL_YIELD: 10,
    TOTAL_EXPENSE: 11, TOTAL_CUMPROFIT: 12, TOTAL_CUMYIELD: 13,
    SPOT_PRINCIPAL: 16, SPOT_VALUE: 17, SPOT_PROFIT: 18, SPOT_YIELD: 19,
    SPOT_CUMPROFIT: 20, SPOT_CUMYIELD: 21, SPOT_FEE: 22,
    FX_PRINCIPAL: 25,  // 手入力管理（書き込まない）
    FX_BALANCE: 26, FX_PROFIT: 27, FX_YIELD: 28,
    FX_CUMPROFIT: 29, FX_CUMYIELD: 30,
    FX_TRADES: 31, FX_WINRATE: 32, FX_BREAKEVEN: 33,
    FX_PF: 34, FX_RR: 35, FX_AVG_PROFIT: 36,
    FX_DD: 37, FX_COST: 38,
    FX_AVG_LOT: 39, FX_REQ_MARGIN: 40
  };

  let cumulativeNet = 0;
  let cumulativeGrossProfit = 0;
  let cumulativeSpotProfit = 0;
  let cumulativeFxProfit = 0;
  let spotPrincipal = DASH_ASSUMPTIONS.INITIAL_SPOT_USD;
  let fxPrincipal = DASH_ASSUMPTIONS.INITIAL_FX_USD;
  let prevSpotValue = DASH_ASSUMPTIONS.INITIAL_SPOT_USD;
  let prevFxBalance = DASH_ASSUMPTIONS.INITIAL_FX_USD;

  months.forEach((m, idx) => {
    const col = DASH_CONFIG.FIRST_DATA_COL + idx;
    const isCurrent = (idx === currentMonthIdx);
    const isPast = (idx < currentMonthIdx);

    const monthKey = `${m.year}-${String(m.month).padStart(2, '0')}`;
    if (cumulativeDeposits[monthKey] !== undefined) {
      spotPrincipal = cumulativeDeposits[monthKey];
    } else {
      const sorted = Object.keys(cumulativeDeposits).sort();
      for (const k of sorted) {
        if (k <= monthKey) spotPrincipal = cumulativeDeposits[k];
      }
    }

    if (isPast) {
      const sv = sheet.getRange(ROW.SPOT_VALUE, col).getValue();
      if (typeof sv === 'number' && sv > 0) prevSpotValue = sv;
      const fb = sheet.getRange(ROW.FX_BALANCE, col).getValue();
      if (typeof fb === 'number' && fb > 0) prevFxBalance = fb;
      const cp = sheet.getRange(ROW.SPOT_CUMPROFIT, col).getValue();
      if (typeof cp === 'number') cumulativeSpotProfit = cp;
      const fcp = sheet.getRange(ROW.FX_CUMPROFIT, col).getValue();
      if (typeof fcp === 'number') cumulativeFxProfit = fcp;
      const tcp = sheet.getRange(ROW.TOTAL_CUMPROFIT, col).getValue();
      if (typeof tcp === 'number') cumulativeNet = tcp;
      cumulativeGrossProfit = cumulativeSpotProfit + cumulativeFxProfit;
      return;
    }

    let spotValue, spotProfit, spotYield, spotFee = 0;
    let fxBalance, fxProfit, fxYield, fxTrades, fxWins, fxLosses, fxWinRate;
    let fxPF, fxRR, fxDD, fxCost;

    if (isCurrent) {
      spotValue = spot.totalUSD;
      spotProfit = spotValue - prevSpotValue;
      spotYield = spotPrincipal > 0 ? spotProfit / spotPrincipal : 0;

      fxBalance = fx ? fx.balance : prevFxBalance;
      fxProfit = fxBalance - prevFxBalance;
      fxYield = fxPrincipal > 0 ? fxProfit / fxPrincipal : 0;
      fxTrades = fx ? fx.totalTrades : 0;
      fxWins = fx ? fx.wins : 0;
      fxLosses = fx ? fx.losses : 0;
      fxWinRate = fx ? fx.winRate / 100 : 0;
      fxPF = fx ? fx.profitFactor : null;
      fxRR = fx ? fx.rrRatio : null;
      fxDD = fx ? fx.maxDDPct / 100 : null;
      fxCost = fx ? Math.abs(fx.commission) : null;
    } else {
      spotProfit = prevSpotValue * DASH_ASSUMPTIONS.SPOT_MONTHLY_YIELD;
      spotValue = prevSpotValue + spotProfit;
      spotYield = spotPrincipal > 0 ? spotProfit / spotPrincipal : 0;

      fxProfit = prevFxBalance * DASH_ASSUMPTIONS.FX_MONTHLY_YIELD;
      fxBalance = prevFxBalance + fxProfit;
      fxYield = fxPrincipal > 0 ? fxProfit / fxPrincipal : 0;
      fxTrades = DASH_ASSUMPTIONS.FX_MONTHLY_TRADES;
      fxWins = Math.round(fxTrades * DASH_ASSUMPTIONS.FX_WIN_RATE);
      fxLosses = fxTrades - fxWins;
      fxWinRate = DASH_ASSUMPTIONS.FX_WIN_RATE;
      fxPF = null; fxRR = null; fxDD = null; fxCost = null;
    }

    cumulativeSpotProfit += spotProfit;
    cumulativeFxProfit += fxProfit;
    const spotCumYield = spotPrincipal > 0 ? cumulativeSpotProfit / spotPrincipal : 0;
    const fxCumYield = fxPrincipal > 0 ? cumulativeFxProfit / fxPrincipal : 0;

    const isInitMonth = (m.year === DASH_ASSUMPTIONS.INITIAL_SETUP_YEAR && m.month === DASH_ASSUMPTIONS.INITIAL_SETUP_MONTH);
    const initialSetup = isInitMonth ? DASH_ASSUMPTIONS.INITIAL_SETUP_USD : 0;
    const totalPrincipal = spotPrincipal + fxPrincipal;
    const totalValue = spotValue + fxBalance;
    const totalGrossProfit = spotProfit + fxProfit;
    const totalExpense = DASH_ASSUMPTIONS.SAGEMASTER_MONTHLY + spotFee;
    const totalYield = totalPrincipal > 0 ? totalGrossProfit / totalPrincipal : 0;

    cumulativeGrossProfit += totalGrossProfit;
    cumulativeNet += totalGrossProfit - totalExpense - initialSetup;
    const totalCumYield = totalPrincipal > 0 ? cumulativeGrossProfit / totalPrincipal : 0;

    if (initialSetup > 0) sheet.getRange(ROW.TOTAL_SETUP, col).setValue(initialSetup);
    sheet.getRange(ROW.TOTAL_PRINCIPAL, col).setValue(totalPrincipal);
    sheet.getRange(ROW.TOTAL_VALUE, col).setValue(totalValue);
    sheet.getRange(ROW.TOTAL_NET, col).setValue(totalGrossProfit);
    sheet.getRange(ROW.TOTAL_YIELD, col).setValue(totalYield);
    sheet.getRange(ROW.TOTAL_EXPENSE, col).setValue(totalExpense);
    sheet.getRange(ROW.TOTAL_CUMPROFIT, col).setValue(cumulativeNet);
    sheet.getRange(ROW.TOTAL_CUMYIELD, col).setValue(totalCumYield);

    sheet.getRange(ROW.SPOT_PRINCIPAL, col).setValue(spotPrincipal);
    sheet.getRange(ROW.SPOT_VALUE, col).setValue(spotValue);
    sheet.getRange(ROW.SPOT_PROFIT, col).setValue(spotProfit);
    sheet.getRange(ROW.SPOT_YIELD, col).setValue(spotYield);
    sheet.getRange(ROW.SPOT_CUMPROFIT, col).setValue(cumulativeSpotProfit);
    sheet.getRange(ROW.SPOT_CUMYIELD, col).setValue(spotCumYield);
    sheet.getRange(ROW.SPOT_FEE, col).setValue(spotFee);

    sheet.getRange(ROW.FX_BALANCE, col).setValue(fxBalance);
    sheet.getRange(ROW.FX_PROFIT, col).setValue(fxProfit);
    sheet.getRange(ROW.FX_YIELD, col).setValue(fxYield);
    sheet.getRange(ROW.FX_CUMPROFIT, col).setValue(cumulativeFxProfit);
    sheet.getRange(ROW.FX_CUMYIELD, col).setValue(fxCumYield);
    sheet.getRange(ROW.FX_TRADES, col).setValue(fxTrades);
    sheet.getRange(ROW.FX_WINRATE, col).setValue(fxWinRate);

    if (fxRR !== null && fxRR > 0) {
      sheet.getRange(ROW.FX_BREAKEVEN, col).setValue(1 / (1 + fxRR));
    }
    if (fxPF !== null) sheet.getRange(ROW.FX_PF, col).setValue(fxPF);
    if (fxRR !== null) sheet.getRange(ROW.FX_RR, col).setValue(fxRR);

    if (fxTrades > 0) {
      sheet.getRange(ROW.FX_AVG_PROFIT, col).setValue(fxProfit / fxTrades);
    }
    if (fxDD !== null) sheet.getRange(ROW.FX_DD, col).setValue(fxDD);
    if (fxCost !== null) sheet.getRange(ROW.FX_COST, col).setValue(fxCost);

    const fxAvgLot = fx ? fx.avgLot : 0.24;
    const fxReqMargin = fx ? fx.reqMargin : 50;
    sheet.getRange(ROW.FX_AVG_LOT, col).setValue(fxAvgLot);
    sheet.getRange(ROW.FX_REQ_MARGIN, col).setValue(fxReqMargin);

    prevSpotValue = spotValue;
    prevFxBalance = fxBalance;
  });

  applyDashboardFormats_(sheet, months.length);
  sheet.getRange('F1').setValue(new Date());
  Logger.log(`✅ データ投入完了（${months.length}ヶ月）`);
}

function readSpotForDashboard_(ss) {
  const sheet = ss.getSheetByName('現物_スナップショット');
  if (!sheet) return { totalUSD: DASH_ASSUMPTIONS.INITIAL_SPOT_USD };
  const data = sheet.getDataRange().getValues();
  const lastRow = data[data.length - 1];
  return { totalUSD: Number(lastRow[4]) || DASH_ASSUMPTIONS.INITIAL_SPOT_USD };
}

function readFXForDashboard_(ss) {
  const sheet = ss.getSheetByName('FX_スナップショット');
  if (!sheet) return null;
  const getNum = (range) => Number(sheet.getRange(range).getValue()) || 0;

  let html = null;
  try { html = getLatestReportFromDrive_(); } catch (e) { Logger.log(`Drive取得スキップ: ${e.message}`); }
  let avgLot = 0, totalPips = 0, totalLots = 0, tradeCount = 0, commission = 0;
  if (html) {
    const parsed = parseMT4HTML_(html);
    const trading = parsed.trades.filter(t => ['buy', 'sell'].includes(t.type));
    tradeCount = trading.length;
    trading.forEach(t => {
      totalLots += t.size;
      const diff = t.closePrice - t.openPrice;
      const sign = (t.type === 'sell') ? -1 : 1;
      totalPips += diff * sign * 10;
      commission += (t.commission || 0);
    });
    avgLot = tradeCount > 0 ? totalLots / tradeCount : 0;
  }

  const goldPrice = 4800;
  const reqMargin = avgLot * 100 * goldPrice / 2222;
  const balance = getNum('B5');
  const equity = getNum('B6');
  const freeMargin = getNum('B7');
  const marginLevel = reqMargin > 0 ? (equity / reqMargin) : null;

  return {
    balance: balance, equity: equity, freeMargin: freeMargin,
    totalTrades: getNum('B10'),
    winRate: getNum('B11') * 100,
    profitFactor: getNum('B12'),
    rrRatio: getNum('B13'),
    maxDDPct: getNum('B19') * 100,
    commission: commission,
    wins: Math.round(getNum('B10') * getNum('B11')),
    losses: getNum('B10') - Math.round(getNum('B10') * getNum('B11')),
    avgLot: avgLot, reqMargin: reqMargin, marginLevel: marginLevel, totalPips: totalPips
  };
}

function generateMonthList_() {
  const months = [];
  for (let y = DASH_CONFIG.START_YEAR; y <= DASH_CONFIG.END_YEAR; y++) {
    const s = (y === DASH_CONFIG.START_YEAR) ? DASH_CONFIG.START_MONTH : 1;
    const e = (y === DASH_CONFIG.END_YEAR) ? DASH_CONFIG.END_MONTH : 12;
    for (let m = s; m <= e; m++) months.push({ year: y, month: m });
  }
  return months;
}

function findCurrentMonthIndex_(months) {
  const now = new Date();
  return months.findIndex(mm => mm.year === now.getFullYear() && mm.month === now.getMonth() + 1);
}

function applyDashboardFormats_(sheet, monthCount) {
  const col1 = DASH_CONFIG.FIRST_DATA_COL;
  const moneyRows = [6, 7, 8, 9, 11, 12, 16, 17, 18, 20, 22, 25, 26, 27, 29, 36, 38, 40];
  const pctRows = [10, 13, 19, 21, 28, 30, 32, 33, 37];
  const intRows = [31];
  const decRows = [34, 35, 39];

  moneyRows.forEach(r => sheet.getRange(r, col1, 1, monthCount).setNumberFormat('$#,##0'));
  pctRows.forEach(r => sheet.getRange(r, col1, 1, monthCount).setNumberFormat('0.0%'));
  intRows.forEach(r => sheet.getRange(r, col1, 1, monthCount).setNumberFormat('#,##0'));
  decRows.forEach(r => sheet.getRange(r, col1, 1, monthCount).setNumberFormat('0.00'));
}

function onEditDashboard(e) {
  if (!e) return;
  const range = e.range;
  if (range.getA1Notation() !== 'B1') return;
  if (range.getSheet().getName() !== DASH_CONFIG.SHEET_MAIN) return;

  const currency = range.getValue();
  const sheet = range.getSheet();
  const rate = Number(sheet.getRange('D1').getValue()) || 155;

  const moneyRows = [6, 7, 8, 9, 11, 12, 16, 17, 18, 20, 22, 25, 26, 27, 29, 36, 38, 40];
  const lastCol = sheet.getLastColumn();
  const firstCol = DASH_CONFIG.FIRST_DATA_COL;

  moneyRows.forEach(r => {
    const rng = sheet.getRange(r, firstCol, 1, lastCol - firstCol + 1);
    const vals = rng.getValues()[0];
    const currentFmt = rng.getNumberFormats()[0][0];
    const isCurrentlyJPY = currentFmt.includes('¥');
    const newVals = vals.map(v => {
      if (typeof v !== 'number') return v;
      if (currency === 'JPY' && !isCurrentlyJPY) return v * rate;
      if (currency === 'USD' && isCurrentlyJPY) return v / rate;
      return v;
    });
    rng.setValues([newVals]);
    rng.setNumberFormat(currency === 'JPY' ? '¥#,##0' : '$#,##0');
  });
}

function addFXAdditionalMetrics() {
  Logger.log('FX追加指標は KPI_ROWS に統合済み（不要）');
}

function applyDashboardFixes() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(DASH_CONFIG.SHEET_MAIN);
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  sheet.getRange(1, 1, lastRow, lastCol).setFontSize(10);
  Logger.log('✅ ダッシュボード修正完了');
}
