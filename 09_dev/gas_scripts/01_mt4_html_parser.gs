/**
 * MT4 HTML詳細レポート自動パーサー
 *
 * 運用フロー：
 *   1. 社長がMT4で「口座履歴→詳細レポートの保存」実行
 *   2. DetailedStatement.htm を Google Drive「SageMaster/FX/MT4_Reports/」に保存
 *   3. 本スクリプトがトリガーで自動実行（月曜7:00 JST）
 *   4. パース結果をスプシ「FX_月次」「raw_FX_trades」に反映
 *
 * 入力ファイル例：BigBoss MT4 Detailed Statement HTML
 * アカウント：1139773 - KAZUMA KATO
 */

// ======================================================
// 設定値
// ======================================================
const MT4_CONFIG = {
  FOLDER_ID: '1JzrH902S5Vy1ZYFgQy3FzxD6exK7Q0js',  // 後方互換用（BigBoss）
  SHEET_RAW: 'raw_FX_trades',
  SHEET_MONTHLY: 'FX_月次',
  SHEET_SUMMARY: 'FX_サマリー',
  ACCOUNT_ID: '1139773',
  CREDIT_BONUS_USD: 3128.76
};

// ブローカー定義（active=true のみが合算対象）
const BROKERS = [
  {
    name: 'BigBoss',
    folderId: '1JzrH902S5Vy1ZYFgQy3FzxD6exK7Q0js',
    accountId: '1139773',
    leverage: 2222,
    snapshotSheet: 'FX_スナップショット_BigBoss',
    active: false  // クローズ済（過去月の参照用にデータは残す）
  },
  {
    name: 'FXTRADING',
    folderId: '1FkUTsBqwTPg-jIj570nYVBwjfBu_hKMH',
    accountId: '888077276',
    leverage: 500,
    snapshotSheet: 'FX_スナップショット_FXTRADING',
    active: true
  }
];

// ======================================================
// メイン関数（トリガー実行対象）
// ======================================================
function runMT4Aggregation() {
  const latestHTML = getLatestReportFromDrive_();
  if (!latestHTML) {
    Logger.log('新しいレポートなし');
    return;
  }

  const parsed = parseMT4HTML_(latestHTML);
  upsertRawTrades_(parsed.trades);
  updateMonthlyStats_(parsed.stats, parsed.summary);
  updateSummarySheet_(parsed.summary);

  // 週次レビューの自動トリガー（別スクリプト）
  if (isMonday_()) {
    generateWeeklyReview();
  }
}

// ======================================================
// Drive から最新HTML取得
// ======================================================
function getLatestReportFromDrive_() {
  // active ブローカーで最新の HTML を返す（後方互換）
  const perBroker = getLatestReportPerBroker_();
  const activeItems = Object.values(perBroker).filter(r => r.broker.active);
  if (activeItems.length === 0) return null;
  const latest = activeItems.reduce((best, cur) => cur.date > best.date ? cur : best);
  Logger.log(`📄 最新（${latest.broker.name}）：${latest.fileName}（更新：${latest.date}）`);
  return latest.html;
}

function getLatestReportPerBroker_() {
  const result = {};
  BROKERS.forEach(broker => {
    try {
      const folder = DriveApp.getFolderById(broker.folderId);
      const files = folder.getFiles();
      let latest = null;
      let latestDate = new Date(0);
      while (files.hasNext()) {
        const f = files.next();
        const name = f.getName();
        if (!name.toLowerCase().endsWith('.htm') && !name.toLowerCase().endsWith('.html')) continue;
        if (f.getLastUpdated() > latestDate) {
          latest = f;
          latestDate = f.getLastUpdated();
        }
      }
      if (latest) {
        result[broker.name] = {
          broker: broker,
          html: latest.getBlob().getDataAsString('UTF-8'),
          date: latestDate,
          fileName: latest.getName()
        };
        Logger.log(`📄 ${broker.name}: ${latest.getName()} (${latestDate})`);
      }
    } catch (e) {
      Logger.log(`${broker.name} 取得失敗: ${e.message}`);
    }
  });
  return result;
}

// ======================================================
// HTMLパース（XMLServiceでは不安定なので正規表現ベース）
// ======================================================
function parseMT4HTML_(html) {
  return {
    trades: extractClosedTrades_(html),
    stats: extractDetailedStats_(html),
    summary: extractSummary_(html)
  };
}

/**
 * Closed Transactionsテーブルから取引明細を抽出
 */
function extractClosedTrades_(html) {
  const trades = [];
  // <tr>...<td title="...">TICKET</td>...</tr> パターン
  const trRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/g;
  const trades_section = extractSection_(html, 'Closed Transactions', 'Open Trades');

  let match;
  while ((match = trRegex.exec(trades_section)) !== null) {
    const cells = extractCells_(match[1]);
    if (cells.length < 14) continue;

    const type = cells[2];
    if (!['buy', 'sell', 'balance', 'credit'].includes(type)) continue;

    trades.push({
      ticket: cells[0],
      openTime: parseMT4Date_(cells[1]),
      type: type,
      size: parseFloat(cells[3]) || 0,
      item: cells[4],
      openPrice: parseFloat(cells[5]) || 0,
      sl: parseFloat(cells[6]) || 0,
      tp: parseFloat(cells[7]) || 0,
      closeTime: parseMT4Date_(cells[8]),
      closePrice: parseFloat(cells[9]) || 0,
      commission: parseFloat(cells[10].replace(/\s/g, '')) || 0,
      taxes: parseFloat(cells[11].replace(/\s/g, '')) || 0,
      swap: parseFloat(cells[12].replace(/\s/g, '')) || 0,
      profit: parseFloat(cells[13].replace(/\s/g, '')) || 0
    });
  }
  return trades;
}

/**
 * Details セクションから詳細統計を抽出（複数行マッチ対応）
 */
function extractDetailedStats_(html) {
  return {
    grossProfit: extractNumber_(html, /Gross Profit:[\s\S]*?<b>([\d\s.,]+)<\/b>/),
    grossLoss: extractNumber_(html, /Gross Loss:[\s\S]*?<b>([\d\s.,]+)<\/b>/),
    totalNetProfit: extractNumber_(html, /Total Net Profit:[\s\S]*?<b>([-\d\s.,]+)<\/b>/),
    profitFactor: extractNumber_(html, /Profit Factor:[\s\S]*?<b>([\d.]+)<\/b>/),
    expectedPayoff: extractNumber_(html, /Expected Payoff:[\s\S]*?<b>([-\d.,]+)<\/b>/),
    absoluteDrawdown: extractNumber_(html, /Absolute Drawdown:[\s\S]*?<b>([\d\s.,]+)<\/b>/),
    maximalDrawdown: extractNumber_(html, /Maximal Drawdown:[\s\S]*?<b>([\d\s.,]+)/),
    maximalDrawdownPct: extractNumber_(html, /Maximal Drawdown:[\s\S]*?\(([\d.]+)%\)/),
    totalTrades: extractNumber_(html, /Total Trades:[\s\S]*?<b>(\d+)<\/b>/),
    winRate: extractNumber_(html, /Profit Trades[\s\S]*?\(([\d.]+)%\)/),
    largestProfit: extractNumber_(html, /Largest[\s\S]*?profit trade:[\s\S]*?<b>([\d.,]+)<\/b>/),
    largestLoss: extractNumber_(html, /Largest[\s\S]*?loss trade:[\s\S]*?<b>([-\d.,]+)<\/b>/),
    avgProfit: extractNumber_(html, /Average[\s\S]*?profit trade:[\s\S]*?<b>([\d.,]+)<\/b>/),
    avgLoss: extractNumber_(html, /Average[\s\S]*?loss trade:[\s\S]*?<b>([-\d.,]+)<\/b>/),
    maxConsecWins: extractNumber_(html, /consecutive wins[\s\S]*?<b>(\d+)/),
    maxConsecLosses: extractNumber_(html, /consecutive losses[\s\S]*?<b>(\d+)/)
  };
}

/**
 * Summary セクションから残高情報を抽出
 */
function extractSummary_(html) {
  return {
    deposit: extractNumber_(html, /Deposit\/Withdrawal:[\s\S]*?<b>([\d\s.,]+)<\/b>/),
    creditFacility: extractNumber_(html, /Credit Facility:[\s\S]*?<b>([\d\s.,]+)<\/b>/),
    closedTradePL: extractNumber_(html, /Closed Trade P\/L:[\s\S]*?<b>([-\d\s.,]+)<\/b>/),
    floatingPL: extractNumber_(html, /Floating P\/L:[\s\S]*?<b>([-\d\s.,]+)<\/b>/),
    balance: extractNumber_(html, /Balance:[\s\S]*?<b>([\d\s.,]+)<\/b>/),
    equity: extractNumber_(html, /Equity:[\s\S]*?<b>([\d\s.,]+)<\/b>/),
    freeMargin: extractNumber_(html, /Free Margin:[\s\S]*?<b>([\d\s.,]+)<\/b>/),
    margin: extractNumber_(html, /Margin:[\s\S]*?<b>([\d\s.,]+)<\/b>/),
    realBalance: 0
  };
}

// ======================================================
// KPI計算（収益性 / 安定性 / 持続性の3軸）
// ======================================================

function calculateKPIs_(trades, stats) {
  const tradingOnly = trades.filter(t => ['buy', 'sell'].includes(t.type));

  const rrRatio = stats.avgLoss !== 0
    ? Math.abs(stats.avgProfit / stats.avgLoss)
    : 0;

  // 損益分岐勝率 = 1 / (1 + RR比)
  const breakEvenWinRate = 1 / (1 + rrRatio) * 100;

  return {
    // 収益性
    profitFactor: stats.profitFactor,
    netProfit: stats.totalNetProfit,
    expectedPayoff: stats.expectedPayoff,

    // 安定性
    winRate: stats.winRate,
    rrRatio: rrRatio,
    breakEvenWinRate: breakEvenWinRate,
    winRateMargin: stats.winRate - breakEvenWinRate,  // 損益分岐勝率との差
    maxDD: stats.maximalDrawdown,
    maxDDPct: stats.maximalDrawdownPct,

    // 持続性
    totalTrades: stats.totalTrades,
    maxConsecWins: stats.maxConsecWins,
    maxConsecLosses: stats.maxConsecLosses,

    // 配信者スコア（3軸加重）
    profitabilityScore: scoreProfitability_(stats),
    stabilityScore: scoreStability_(stats, rrRatio),
    sustainabilityScore: scoreSustainability_(stats, tradingOnly.length),
    totalScore: 0  // 以下で計算
  };
}

function scoreProfitability_(stats) {
  if (stats.profitFactor >= 2.0) return 10;
  if (stats.profitFactor >= 1.5) return 8;
  if (stats.profitFactor >= 1.3) return 6;
  if (stats.profitFactor >= 1.0) return 4;
  return 2;
}

function scoreStability_(stats, rrRatio) {
  let score = 0;
  if (stats.winRate >= 70) score += 3;
  else if (stats.winRate >= 60) score += 2;
  if (rrRatio >= 1.0) score += 3;
  else if (rrRatio >= 0.7) score += 2;
  else if (rrRatio >= 0.5) score += 1;
  if (stats.maximalDrawdownPct < 10) score += 3;
  else if (stats.maximalDrawdownPct < 15) score += 2;
  return Math.min(score, 10);
}

function scoreSustainability_(stats, tradeCount) {
  let score = 0;
  if (tradeCount >= 100) score += 4;
  else if (tradeCount >= 50) score += 2;
  if (stats.maxConsecLosses <= 3) score += 3;
  if (stats.maxConsecWins >= 5) score += 3;
  return Math.min(score, 10);
}

// ======================================================
// ロット最適化計算
// ======================================================

/**
 * リスク%に基づく推奨ロット算出
 * XAUUSD（ゴールド）：1ロット = $10/pip
 */
function calculateRecommendedLot_(balance, riskPct, slPips) {
  const pipValue = 10;  // XAUUSD 1ロット = $10/pip
  return Math.floor(
    (balance * riskPct / 100) / (slPips * pipValue) * 100
  ) / 100;
}

function generateLotAnalysisTable_(balance, avgSLPips) {
  const scenarios = [1, 2, 3, 5, 7, 10];
  return scenarios.map(riskPct => ({
    riskPct: riskPct,
    maxLot: calculateRecommendedLot_(balance, riskPct, avgSLPips),
    maxLoss: balance * riskPct / 100,
    label: getRiskLabel_(riskPct)
  }));
}

function getRiskLabel_(riskPct) {
  if (riskPct <= 1) return '超安全';
  if (riskPct <= 2) return '安全（推奨）';
  if (riskPct <= 3) return '標準';
  if (riskPct <= 5) return '現状';
  if (riskPct <= 7) return '攻め（PF>1.3必須）';
  return '危険';
}

// ======================================================
// ユーティリティ
// ======================================================

function extractCells_(trHtml) {
  const cells = [];
  const tdRegex = /<td[^>]*>([\s\S]*?)<\/td>/g;
  let match;
  while ((match = tdRegex.exec(trHtml)) !== null) {
    cells.push(match[1].replace(/<[^>]+>/g, '').trim());
  }
  return cells;
}

function extractSection_(html, startMarker, endMarker) {
  const start = html.indexOf(startMarker);
  const end = html.indexOf(endMarker, start);
  if (start < 0) return '';
  return end > 0 ? html.substring(start, end) : html.substring(start);
}

function extractNumber_(html, regex) {
  const match = html.match(regex);
  if (!match) return 0;
  return parseFloat(match[1].replace(/[\s,]/g, '')) || 0;
}

function parseMT4Date_(str) {
  const m = str.match(/(\d{4})\.(\d{2})\.(\d{2})\s+(\d{2}):(\d{2}):(\d{2})/);
  if (!m) return null;
  return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]);
}

function getOrCreateFolder_(path) {
  const parts = path.split('/');
  let folder = DriveApp.getRootFolder();
  for (const name of parts) {
    const sub = folder.getFoldersByName(name);
    folder = sub.hasNext() ? sub.next() : folder.createFolder(name);
  }
  return folder;
}

function isMonday_() {
  return new Date().getDay() === 1;
}

function upsertRawTrades_(trades) {
  // raw_FX_trades シートにチケット番号重複排除で追記
  // ...実装詳細はDay2で
}

function updateMonthlyStats_(stats, summary) {
  // FX_月次シートを更新
  // ...実装詳細はDay2で
}

function updateSummarySheet_(summary) {
  // 真正Balance = Balance - Credit Facility
  summary.realBalance = summary.balance;  // Credit除外済み
  // ...実装詳細はDay2で
}

// ======================================================
// 接続テスト＆データ確認用
// ======================================================
function testMT4Parse() {
  const html = getLatestReportFromDrive_();
  if (!html) {
    Logger.log('❌ Driveに HTMLファイルが見つかりません');
    Logger.log(`   フォルダID：${MT4_CONFIG.FOLDER_ID}`);
    return;
  }
  Logger.log(`✅ HTMLファイル読み込みOK（${html.length.toLocaleString()}文字）`);

  const parsed = parseMT4HTML_(html);

  Logger.log('━━━━━━━━━━━━━━━━');
  Logger.log('📊 サマリー');
  Logger.log(`  Deposit: $${parsed.summary.deposit}`);
  Logger.log(`  Credit: $${parsed.summary.creditFacility}`);
  Logger.log(`  Balance（真正資産）: $${parsed.summary.balance}`);
  Logger.log(`  Equity: $${parsed.summary.equity}`);
  Logger.log(`  Free Margin: $${parsed.summary.freeMargin}`);

  Logger.log('━━━━━━━━━━━━━━━━');
  Logger.log('📈 パフォーマンス');
  Logger.log(`  Total Trades: ${parsed.stats.totalTrades}`);
  Logger.log(`  勝率: ${parsed.stats.winRate}%`);
  Logger.log(`  PF: ${parsed.stats.profitFactor}`);
  Logger.log(`  Net Profit: $${parsed.stats.totalNetProfit}`);
  Logger.log(`  最大DD: $${parsed.stats.maximalDrawdown} (${parsed.stats.maximalDrawdownPct}%)`);

  const rrRatio = parsed.stats.avgLoss !== 0
    ? Math.abs(parsed.stats.avgProfit / parsed.stats.avgLoss) : 0;
  const breakEven = 1 / (1 + rrRatio) * 100;
  Logger.log(`  RR比: ${rrRatio.toFixed(2)} / 損益分岐勝率: ${breakEven.toFixed(2)}%`);

  Logger.log('━━━━━━━━━━━━━━━━');
  Logger.log(`🔍 取引明細：${parsed.trades.length}件抽出`);
  Logger.log('🎉 MT4パース完了！');
}

// ======================================================
// FX実績をスナップショットシートへ書き込み（per-broker + 合算）
// ======================================================
function testWriteFXToSheet() {
  const ss = SpreadsheetApp.getActive();
  const perBroker = getLatestReportPerBroker_();
  if (Object.keys(perBroker).length === 0) {
    Logger.log('❌ どのブローカーフォルダにもHTMLなし');
    return;
  }

  // ブローカー別シート書き込み
  Object.values(perBroker).forEach(item => {
    const parsed = parseMT4HTML_(item.html);
    writeFXSnapshotSheet_(ss, item.broker.snapshotSheet, parsed, item.broker.name);
  });

  // 合算シート書き込み（active のみ）
  const activeItems = Object.values(perBroker).filter(r => r.broker.active);
  if (activeItems.length > 0) {
    const aggregated = aggregateBrokers_(activeItems);
    writeFXSnapshotSheet_(ss, 'FX_スナップショット', aggregated.parsed,
      activeItems.map(i => i.broker.name).join('+'));
  }

  Logger.log(`✅ FX_スナップショット書き込み完了（${Object.keys(perBroker).length}ブローカー + 合算）`);
}

function aggregateBrokers_(items) {
  const allTrades = [];
  let balance = 0, equity = 0, freeMargin = 0, deposit = 0, credit = 0;
  let maxDD = 0, maxDDPct = 0, largestProfit = 0, largestLoss = 0;

  items.forEach(item => {
    const parsed = parseMT4HTML_(item.html);
    allTrades.push(...(parsed.trades || []));
    balance += parsed.summary.balance || 0;
    equity += parsed.summary.equity || 0;
    freeMargin += parsed.summary.freeMargin || 0;
    deposit += parsed.summary.deposit || 0;
    credit += parsed.summary.creditFacility || 0;
    if ((parsed.stats.maximalDrawdownPct || 0) > maxDDPct) maxDDPct = parsed.stats.maximalDrawdownPct;
    if ((parsed.stats.maximalDrawdown || 0) > maxDD) maxDD = parsed.stats.maximalDrawdown;
    if ((parsed.stats.largestProfit || 0) > largestProfit) largestProfit = parsed.stats.largestProfit;
    if ((parsed.stats.largestLoss || 0) < largestLoss) largestLoss = parsed.stats.largestLoss;
  });

  const tradingOnly = allTrades.filter(t => ['buy', 'sell'].includes(t.type));
  const wins = tradingOnly.filter(t => t.profit > 0);
  const losses = tradingOnly.filter(t => t.profit < 0);
  const grossProfit = wins.reduce((s, t) => s + t.profit, 0);
  const grossLoss = Math.abs(losses.reduce((s, t) => s + t.profit, 0));
  const avgProfit = wins.length > 0 ? grossProfit / wins.length : 0;
  const avgLoss = losses.length > 0 ? -grossLoss / losses.length : 0;
  const totalNet = grossProfit - grossLoss;

  return {
    parsed: {
      summary: {
        balance, equity, freeMargin, deposit, creditFacility: credit, realBalance: balance,
        floatingPL: 0, closedTradePL: totalNet, margin: 0
      },
      stats: {
        totalTrades: tradingOnly.length,
        winRate: tradingOnly.length > 0 ? (wins.length / tradingOnly.length) * 100 : 0,
        profitFactor: grossLoss > 0 ? grossProfit / grossLoss : 0,
        grossProfit, grossLoss,
        totalNetProfit: totalNet,
        avgProfit, avgLoss,
        expectedPayoff: tradingOnly.length > 0 ? totalNet / tradingOnly.length : 0,
        maximalDrawdown: maxDD,
        maximalDrawdownPct: maxDDPct,
        largestProfit, largestLoss,
        absoluteDrawdown: 0,
        maxConsecWins: 0, maxConsecLosses: 0
      },
      trades: allTrades
    }
  };
}

function writeFXSnapshotSheet_(ss, sheetName, parsed, brokerLabel) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  else sheet.clear();

  const rrRatio = parsed.stats.avgLoss !== 0
    ? Math.abs(parsed.stats.avgProfit / parsed.stats.avgLoss) : 0;
  const breakEven = 1 / (1 + rrRatio) * 100;
  const now = new Date();

  // 残高サマリー
  const summary = [
    ['📊 残高サマリー', '', '取得日時', now],
    ['項目', '金額($)', '備考', ''],
    ['Deposit（入金額）', parsed.summary.deposit, '初期入金', ''],
    ['Credit（ボーナス）', parsed.summary.creditFacility, '⚠️ 出金不可・参考値', ''],
    ['Balance（真正資産）', parsed.summary.balance, '✅ 実資産', ''],
    ['Equity', parsed.summary.equity, 'Balance＋Credit＋含み損益', ''],
    ['Free Margin', parsed.summary.freeMargin, '取引可能証拠金', '']
  ];

  // パフォーマンス
  const perfHeader = [['📈 パフォーマンス指標', '実績', '目標', '判定']];
  const perf = [
    ['総トレード数', parsed.stats.totalTrades, '—', ''],
    ['勝率', parsed.stats.winRate / 100, breakEven / 100, judgeFX_(parsed.stats.winRate, breakEven)],
    ['プロフィットファクター（PF）', parsed.stats.profitFactor, 1.3, judgeFX_(parsed.stats.profitFactor, 1.3)],
    ['RR比', rrRatio, 0.7, judgeFX_(rrRatio, 0.7)],
    ['期待値（$/トレード）', parsed.stats.expectedPayoff, 0, judgeFX_(parsed.stats.expectedPayoff, 0)],
    ['Gross Profit', parsed.stats.grossProfit, '—', ''],
    ['Gross Loss', parsed.stats.grossLoss, '—', ''],
    ['Net Profit', parsed.stats.totalNetProfit, 149, judgeFX_(parsed.stats.totalNetProfit, 149)],
    ['最大DD（$）', parsed.stats.maximalDrawdown, '—', ''],
    ['最大DD（%）', parsed.stats.maximalDrawdownPct / 100, 0.10, judgeFX_(0.10, parsed.stats.maximalDrawdownPct / 100)],
    ['損益分岐勝率', breakEven / 100, '—', '']
  ];

  // 取引明細
  const tradeHeader = [[`📋 取引明細（${brokerLabel || '直近'}）`, '', '', '', '', '', '']];
  const tradeCols = [['チケット', '建玉時刻', 'タイプ', 'ロット', '建値', '決済', 'P/L($)']];
  const tradingOnly = (parsed.trades || []).filter(t => ['buy', 'sell'].includes(t.type));
  const tradeRows = tradingOnly.map(t => [
    t.ticket, t.openTime, t.type, t.size, t.openPrice, t.closePrice, t.profit
  ]);

  // 書き込み
  let row = 1;
  sheet.getRange(row, 1, summary.length, 4).setValues(summary);
  sheet.getRange(row, 1, 1, 4).setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');
  sheet.getRange(row + 1, 1, 1, 4).setFontWeight('bold').setBackground('#cfe2f3');
  row += summary.length + 1;

  sheet.getRange(row, 1, 1, 4).setValues(perfHeader)
    .setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');
  row++;
  sheet.getRange(row, 1, perf.length, 4).setValues(perf);
  row += perf.length + 1;

  sheet.getRange(row, 1, 1, 7).setValues(tradeHeader)
    .setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');
  row++;
  sheet.getRange(row, 1, 1, 7).setValues(tradeCols)
    .setFontWeight('bold').setBackground('#cfe2f3');
  row++;
  if (tradeRows.length > 0) {
    sheet.getRange(row, 1, tradeRows.length, 7).setValues(tradeRows);
  }

  // 書式
  const lastRow = sheet.getLastRow();
  sheet.getRange(1, 1, lastRow, 7).setFontSize(10);
  sheet.getRange('A1:D1').setHorizontalAlignment('center');
  sheet.getRange('A2:C2').setHorizontalAlignment('center');
  sheet.getRange('A9:D9').setHorizontalAlignment('center');
  sheet.getRange('A22:G22').setHorizontalAlignment('center');
  sheet.getRange('A23:G23').setHorizontalAlignment('center');
  sheet.getRange(3, 1, 5, 1).setHorizontalAlignment('center');
  sheet.getRange(10, 1, 11, 1).setHorizontalAlignment('center');

  const tradeDataRows = lastRow - 23;
  if (tradeDataRows > 0) {
    sheet.getRange(24, 1, tradeDataRows, 3).setHorizontalAlignment('center');
    sheet.getRange(24, 2, tradeDataRows, 1).setNumberFormat('yyyy/MM/dd HH:mm');
    sheet.getRange(24, 5, tradeDataRows, 3).setNumberFormat('#,##0');
  }

  sheet.getRange('B11:C11').setNumberFormat('0.0%');
  sheet.getRange('B13:C13').setNumberFormat('0.0%');
  sheet.getRange('B19:C19').setNumberFormat('0.0%');
  sheet.getRange('B20:C20').setNumberFormat('0.0%');
  sheet.getRange('B3:B7').setNumberFormat('#,##0');
  sheet.getRange('D1').setNumberFormat('yyyy/MM/dd HH:mm:ss');

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidths(2, 6, 170);

  Logger.log(`✅ ${sheetName} 書き込み完了`);
}

function judgeFX_(actual, target) {
  if (typeof actual !== 'number' || typeof target !== 'number') return '';
  if (actual >= target) return '🟢';
  if (actual >= target * 0.85) return '🟡';
  return '🔴';
}
