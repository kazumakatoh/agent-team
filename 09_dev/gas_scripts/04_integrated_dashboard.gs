/**
 * 統合ダッシュボード構築
 * 2026/4 〜 2030/12（57ヶ月分）
 * 円/ドル切替対応
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
  { type: 'data', label: '元本（$）', fmt: 'money' },
  { type: 'data', label: '純利益（$）', fmt: 'money' },
  { type: 'data', label: '利回り（%）', fmt: 'percent' },
  { type: 'data', label: '経費（$）', fmt: 'money' },
  { type: 'blank' },
  { type: 'header', label: '📈 現物（AIグリッド）' },
  { type: 'data', label: '元本（$）', fmt: 'money' },
  { type: 'data', label: '月次利益（$）', fmt: 'money' },
  { type: 'data', label: '月次利回り（%）', fmt: 'percent' },
  { type: 'data', label: '手数料（$）', fmt: 'money' },
  { type: 'blank' },
  { type: 'header', label: '💱 FX（BigBoss）' },
  { type: 'data', label: 'Balance（$）', fmt: 'money' },
  { type: 'data', label: '月次利益（$）', fmt: 'money' },
  { type: 'data', label: '月次利回り（%）', fmt: 'percent' },
  { type: 'data', label: '取引回数', fmt: 'int' },
  { type: 'data', label: '勝ち', fmt: 'int' },
  { type: 'data', label: '負け', fmt: 'int' },
  { type: 'data', label: '勝率（%）', fmt: 'percent' },
  { type: 'data', label: 'PF', fmt: 'decimal' },
  { type: 'data', label: 'RR比', fmt: 'decimal' },
  { type: 'data', label: '最大DD（%）', fmt: 'percent' },
  { type: 'data', label: '取引コスト（$）', fmt: 'money' }
];

const DASH_ASSUMPTIONS = {
  INITIAL_SPOT_USD: 2972,
  INITIAL_FX_USD: 6253.51,
  SPOT_MONTHLY_YIELD: 0.10,
  FX_MONTHLY_YIELD: 0.05,
  FX_MONTHLY_TRADES: 60,
  FX_WIN_RATE: 0.70,
  SAGEMASTER_MONTHLY: 149
};

// ======================================================
// 初期構築
// ======================================================
function initIntegratedDashboard() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(DASH_CONFIG.SHEET_MAIN);
  if (!sheet) sheet = ss.insertSheet(DASH_CONFIG.SHEET_MAIN);
  else sheet.clear();

  const months = generateMonthList_();
  const lastCol = DASH_CONFIG.FIRST_DATA_COL + months.length - 1;

  // Row 1: 通貨切替・為替・更新日時
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

  // Row 3: 年ヘッダー（マージ）
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

  // Row 4: 月ヘッダー
  const monthLabels = months.map(m => m.month + '月');
  sheet.getRange(4, DASH_CONFIG.FIRST_DATA_COL, 1, months.length)
    .setValues([monthLabels])
    .setHorizontalAlignment('center').setFontWeight('bold').setBackground('#cfe2f3');

  // データ行
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

  // 当月列を黄色ハイライト
  const totalRows = DATA_START + KPI_ROWS.length - 1;
  sheet.getRange(5, DASH_CONFIG.FIRST_DATA_COL, KPI_ROWS.length, 1).setBackground('#fff2cc');

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidths(DASH_CONFIG.FIRST_DATA_COL, months.length, 90);
  sheet.getRange(1, 1, totalRows, lastCol).setFontSize(10);
  sheet.setFrozenRows(4);
  sheet.setFrozenColumns(1);

  Logger.log(`✅ 統合ダッシュボード初期化完了（${months.length}ヶ月分）`);
}

// ======================================================
// FX追加指標（5項目）追加
// ======================================================
function addFXAdditionalMetrics() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(DASH_CONFIG.SHEET_MAIN);
  const labels = [
    ['平均取引ロット数'],
    ['必要証拠金（$）'],
    ['証拠金維持率（%）'],
    ['余剰証拠金（$）'],
    ['獲得値幅（pips）']
  ];
  sheet.getRange(29, 1, 5, 1).setValues(labels);
  Logger.log('✅ 5項目追加完了');
}

// ======================================================
// データ投入
// ======================================================
function populateIntegratedDashboard() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(DASH_CONFIG.SHEET_MAIN);
  if (!sheet) throw new Error('統合ダッシュボードが未作成です');

  const spot = readSpotForDashboard_(ss);
  const fx = readFXForDashboard_(ss);
  const months = generateMonthList_();
  const currentMonthIdx = findCurrentMonthIndex_(months);

  const ROW = {
    TOTAL_PRINCIPAL: 6, TOTAL_NET: 7, TOTAL_YIELD: 8, TOTAL_EXPENSE: 9,
    SPOT_PRINCIPAL: 12, SPOT_PROFIT: 13, SPOT_YIELD: 14, SPOT_FEE: 15,
    FX_BALANCE: 18, FX_PROFIT: 19, FX_YIELD: 20,
    FX_TRADES: 21, FX_WINS: 22, FX_LOSSES: 23,
    FX_WINRATE: 24, FX_PF: 25, FX_RR: 26, FX_DD: 27, FX_COST: 28,
    FX_AVG_LOT: 29, FX_REQ_MARGIN: 30, FX_MARGIN_LEVEL: 31,
    FX_FREE_MARGIN: 32, FX_PIPS: 33
  };

  let spotPrincipal = DASH_ASSUMPTIONS.INITIAL_SPOT_USD;
  let fxPrincipal = DASH_ASSUMPTIONS.INITIAL_FX_USD;

  months.forEach((m, idx) => {
    const col = DASH_CONFIG.FIRST_DATA_COL + idx;
    const isCurrent = (idx === currentMonthIdx);
    const isPast = (idx < currentMonthIdx);
    if (isPast) return;

    let spotProfit, spotYield, spotFee;
    let fxBalance, fxProfit, fxYield, fxTrades, fxWins, fxLosses, fxWinRate;
    let fxPF, fxRR, fxDD, fxCost;
    let fxAvgLot, fxReqMargin, fxMarginLevel, fxFreeMargin, fxPips;

    if (isCurrent) {
      spotProfit = spot.totalUSD - spotPrincipal;
      spotYield = spotProfit / spotPrincipal;
      spotFee = 0;

      fxBalance = fx.balance;
      fxProfit = fx.balance - fxPrincipal;
      fxYield = fxProfit / fxPrincipal;
      fxTrades = fx.totalTrades;
      fxWins = fx.wins;
      fxLosses = fx.losses;
      fxWinRate = fx.winRate / 100;
      fxPF = fx.profitFactor;
      fxRR = fx.rrRatio;
      fxDD = fx.maxDDPct / 100;
      fxCost = Math.abs(fx.commission);

      fxAvgLot = fx.avgLot;
      fxReqMargin = fx.reqMargin;
      fxMarginLevel = fx.marginLevel;
      fxFreeMargin = fx.freeMargin;
      fxPips = fx.totalPips;
    } else {
      spotProfit = spotPrincipal * DASH_ASSUMPTIONS.SPOT_MONTHLY_YIELD;
      spotYield = DASH_ASSUMPTIONS.SPOT_MONTHLY_YIELD;
      spotFee = 0;

      fxProfit = fxPrincipal * DASH_ASSUMPTIONS.FX_MONTHLY_YIELD;
      fxBalance = fxPrincipal + fxProfit;
      fxYield = DASH_ASSUMPTIONS.FX_MONTHLY_YIELD;
      fxTrades = DASH_ASSUMPTIONS.FX_MONTHLY_TRADES;
      fxWins = Math.round(fxTrades * DASH_ASSUMPTIONS.FX_WIN_RATE);
      fxLosses = fxTrades - fxWins;
      fxWinRate = DASH_ASSUMPTIONS.FX_WIN_RATE;
      fxPF = null; fxRR = null; fxDD = null; fxCost = null;

      fxAvgLot = fx ? fx.avgLot : 0.24;
      fxReqMargin = fx ? fx.reqMargin : 50;
      fxMarginLevel = fxPrincipal / fxReqMargin;
      fxFreeMargin = fxPrincipal - fxReqMargin;
      fxPips = DASH_ASSUMPTIONS.FX_MONTHLY_TRADES * 10;
    }

    // 現物
    sheet.getRange(ROW.SPOT_PRINCIPAL, col).setValue(spotPrincipal);
    sheet.getRange(ROW.SPOT_PROFIT, col).setValue(spotProfit);
    sheet.getRange(ROW.SPOT_YIELD, col).setValue(spotYield);
    sheet.getRange(ROW.SPOT_FEE, col).setValue(spotFee);

    // FX
    sheet.getRange(ROW.FX_BALANCE, col).setValue(fxPrincipal);
    sheet.getRange(ROW.FX_PROFIT, col).setValue(fxProfit);
    sheet.getRange(ROW.FX_YIELD, col).setValue(fxYield);
    sheet.getRange(ROW.FX_TRADES, col).setValue(fxTrades);
    sheet.getRange(ROW.FX_WINS, col).setValue(fxWins);
    sheet.getRange(ROW.FX_LOSSES, col).setValue(fxLosses);
    sheet.getRange(ROW.FX_WINRATE, col).setValue(fxWinRate);
    if (fxPF !== null) sheet.getRange(ROW.FX_PF, col).setValue(fxPF);
    if (fxRR !== null) sheet.getRange(ROW.FX_RR, col).setValue(fxRR);
    if (fxDD !== null) sheet.getRange(ROW.FX_DD, col).setValue(fxDD);
    if (fxCost !== null) sheet.getRange(ROW.FX_COST, col).setValue(fxCost);

    sheet.getRange(ROW.FX_AVG_LOT, col).setValue(fxAvgLot);
    sheet.getRange(ROW.FX_REQ_MARGIN, col).setValue(fxReqMargin);
    if (fxMarginLevel !== null) sheet.getRange(ROW.FX_MARGIN_LEVEL, col).setValue(fxMarginLevel);
    sheet.getRange(ROW.FX_FREE_MARGIN, col).setValue(fxFreeMargin);
    sheet.getRange(ROW.FX_PIPS, col).setValue(fxPips);

    // 統合
    const totalPrincipal = spotPrincipal + fxPrincipal;
    const totalNet = spotProfit + fxProfit - DASH_ASSUMPTIONS.SAGEMASTER_MONTHLY - spotFee - (fxCost || 0);
    const totalYield = totalNet / totalPrincipal;
    const totalExpense = DASH_ASSUMPTIONS.SAGEMASTER_MONTHLY + spotFee + (fxCost || 0);

    sheet.getRange(ROW.TOTAL_PRINCIPAL, col).setValue(totalPrincipal);
    sheet.getRange(ROW.TOTAL_NET, col).setValue(totalNet);
    sheet.getRange(ROW.TOTAL_YIELD, col).setValue(totalYield);
    sheet.getRange(ROW.TOTAL_EXPENSE, col).setValue(totalExpense);

    spotPrincipal = spotPrincipal + spotProfit;
    fxPrincipal = fxBalance;
  });

  applyDashboardFormats_(sheet, months.length);
  sheet.getRange('F1').setValue(new Date());

  Logger.log(`✅ データ投入完了（${months.length}ヶ月）`);
}

// ======================================================
// スナップショット読み込み
// ======================================================
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

  const html = getLatestReportFromDrive_();
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
    avgLot: avgLot,
    reqMargin: reqMargin,
    marginLevel: marginLevel,
    totalPips: totalPips
  };
}

// ======================================================
// ユーティリティ
// ======================================================
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
  const moneyRows = [6, 7, 9, 12, 13, 15, 18, 19, 28, 30, 32];
  const pctRows = [8, 14, 20, 24, 27];
  const intRows = [21, 22, 23, 33];
  const decRows = [25, 26, 29];

  moneyRows.forEach(r => sheet.getRange(r, col1, 1, monthCount).setNumberFormat('$#,##0'));
  pctRows.forEach(r => sheet.getRange(r, col1, 1, monthCount).setNumberFormat('0.0%'));
  intRows.forEach(r => sheet.getRange(r, col1, 1, monthCount).setNumberFormat('#,##0'));
  decRows.forEach(r => sheet.getRange(r, col1, 1, monthCount).setNumberFormat('0.00'));
  // 証拠金維持率は倍率表記
  sheet.getRange(31, col1, 1, monthCount).setNumberFormat('0.0"倍"');
}

// ======================================================
// 修正適用（A列ラベル整理・倍率表記・フォントサイズ統一）
// ======================================================
function applyDashboardFixes() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(DASH_CONFIG.SHEET_MAIN);

  const newLabels = {
    6: '元本', 7: '純利益', 9: '経費',
    12: '元本', 13: '月次利益', 15: '手数料',
    18: 'Balance', 19: '月次利益', 28: '取引コスト',
    30: '必要証拠金', 32: '余剰証拠金'
  };
  Object.keys(newLabels).forEach(r => sheet.getRange(Number(r), 1).setValue(newLabels[r]));

  sheet.getRange(31, 1).setValue('証拠金維持率（倍率）');
  const lastCol = sheet.getLastColumn();
  sheet.getRange(31, 2, 1, lastCol - 1).setNumberFormat('0.0"倍"');

  const lastRow = sheet.getLastRow();
  sheet.getRange(1, 1, lastRow, lastCol).setFontSize(10);

  Logger.log('✅ ダッシュボード修正完了');
}

// ======================================================
// 円/ドル切替（onEdit トリガー）
// ======================================================
function onEditDashboard(e) {
  if (!e || e.range.getA1Notation() !== 'B1') return;
  if (e.range.getSheet().getName() !== DASH_CONFIG.SHEET_MAIN) return;

  const sheet = e.range.getSheet();
  const currency = e.range.getValue();
  const rate = Number(sheet.getRange('D1').getValue()) || 155;

  const moneyRows = [6, 7, 9, 12, 13, 15, 18, 19, 28, 30, 32];
  const lastCol = sheet.getLastColumn();
  const firstCol = DASH_CONFIG.FIRST_DATA_COL;

  moneyRows.forEach(r => {
    const range = sheet.getRange(r, firstCol, 1, lastCol - firstCol + 1);
    const vals = range.getValues()[0];
    const currentFmt = range.getNumberFormats()[0][0];
    const isCurrentlyJPY = currentFmt.includes('¥');
    const newVals = vals.map(v => {
      if (typeof v !== 'number') return v;
      if (currency === 'JPY' && !isCurrentlyJPY) return v * rate;
      if (currency === 'USD' && isCurrentlyJPY) return v / rate;
      return v;
    });
    range.setValues([newVals]);
    range.setNumberFormat(currency === 'JPY' ? '¥#,##0' : '$#,##0');
  });
}
