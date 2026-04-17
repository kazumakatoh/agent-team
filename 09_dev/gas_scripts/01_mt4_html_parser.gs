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
  FOLDER_NAME: 'SageMaster/FX/MT4_Reports',
  SHEET_RAW: 'raw_FX_trades',
  SHEET_MONTHLY: 'FX_月次',
  SHEET_SUMMARY: 'FX_サマリー',
  ACCOUNT_ID: '1139773',
  CREDIT_BONUS_USD: 3128.76  // 出金不可のクレジットボーナス（参考値扱い）
};

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
  const folder = getOrCreateFolder_(MT4_CONFIG.FOLDER_NAME);
  const files = folder.getFilesByType(MimeType.HTML);
  let latest = null;
  let latestDate = new Date(0);

  while (files.hasNext()) {
    const f = files.next();
    if (f.getLastUpdated() > latestDate) {
      latest = f;
      latestDate = f.getLastUpdated();
    }
  }
  return latest ? latest.getBlob().getDataAsString('UTF-8') : null;
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
 * Details セクションから詳細統計を抽出
 */
function extractDetailedStats_(html) {
  return {
    grossProfit: extractNumber_(html, /Gross Profit:.*?<b>([\d\s.,]+)<\/b>/),
    grossLoss: extractNumber_(html, /Gross Loss:.*?<b>([\d\s.,]+)<\/b>/),
    totalNetProfit: extractNumber_(html, /Total Net Profit:.*?<b>([-\d\s.,]+)<\/b>/),
    profitFactor: extractNumber_(html, /Profit Factor:.*?<b>([\d.]+)<\/b>/),
    expectedPayoff: extractNumber_(html, /Expected Payoff:.*?<b>([-\d.,]+)<\/b>/),
    absoluteDrawdown: extractNumber_(html, /Absolute Drawdown:.*?<b>([\d\s.,]+)<\/b>/),
    maximalDrawdown: extractNumber_(html, /Maximal Drawdown:.*?<b>([\d\s.,]+)/),
    maximalDrawdownPct: extractNumber_(html, /Maximal Drawdown:.*?\(([\d.]+)%\)/),
    totalTrades: extractNumber_(html, /Total Trades:.*?<b>(\d+)<\/b>/),
    winRate: extractNumber_(html, /Profit Trades.*?\(([\d.]+)%\)/),
    largestProfit: extractNumber_(html, /Largest[\s\S]*?profit trade:.*?<b>([\d.,]+)<\/b>/),
    largestLoss: extractNumber_(html, /Largest[\s\S]*?loss trade:.*?<b>([-\d.,]+)<\/b>/),
    avgProfit: extractNumber_(html, /Average[\s\S]*?profit trade:.*?<b>([\d.,]+)<\/b>/),
    avgLoss: extractNumber_(html, /Average[\s\S]*?loss trade:.*?<b>([-\d.,]+)<\/b>/),
    maxConsecWins: extractNumber_(html, /consecutive wins.*?<b>(\d+)/),
    maxConsecLosses: extractNumber_(html, /consecutive losses.*?<b>(\d+)/)
  };
}

/**
 * Summary セクションから残高情報を抽出
 */
function extractSummary_(html) {
  return {
    deposit: extractNumber_(html, /Deposit\/Withdrawal:.*?<b>([\d\s.,]+)<\/b>/),
    creditFacility: extractNumber_(html, /Credit Facility:.*?<b>([\d\s.,]+)<\/b>/),
    closedTradePL: extractNumber_(html, /Closed Trade P\/L:.*?<b>([-\d\s.,]+)<\/b>/),
    floatingPL: extractNumber_(html, /Floating P\/L:.*?<b>([-\d\s.,]+)<\/b>/),
    balance: extractNumber_(html, /Balance:.*?<b>([\d\s.,]+)<\/b>/),
    equity: extractNumber_(html, /Equity:.*?<b>([\d\s.,]+)<\/b>/),
    freeMargin: extractNumber_(html, /Free Margin:.*?<b>([\d\s.,]+)<\/b>/),
    margin: extractNumber_(html, /Margin:.*?<b>([\d\s.,]+)<\/b>/),
    // 真正資産（クレジット除外）
    realBalance: 0  // calculateRealBalance_() で上書き
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
