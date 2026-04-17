/**
 * 週次レビュー自動生成
 *
 * 実行タイミング：毎週月曜 7:00 JST（トリガー）
 * 出力先：
 *   1. スプシ「週次レビュー」シートに追記
 *   2. 社長のGmailに通知
 *   3. Looker Studioダッシュボードは自動更新
 *
 * 設計原則：
 *   - 一喜一憂しない（短期ノイズは強調しない）
 *   - 目標vs実績を明示
 *   - アクションは閾値判定で自動提案
 *   - FX用語には★マーク（用語集へのリンク）
 */

// ======================================================
// 目標値（スプシ ④目標vs実績対比 シートと同期）
// ======================================================
const TARGETS = {
  fxMonthlyYield: 5.0,      // FX 月次利回り %
  fxWinRate: 71,            // FX 勝率 %（損益分岐）
  fxProfitFactor: 1.3,      // FX PF
  fxRRRatio: 0.7,           // FX RR比
  fxMaxDD: 10,              // FX 最大DD %
  spotMonthlyYield: 8.0,    // 現物 月次利回り %
  spotFeeRate: 0.5,         // 現物 手数料率 %
  totalMoMGrowth: 6.5,      // 総資産 前月比 %
  sageMasterFee: 149        // サブスク費用 $
};

// ======================================================
// メイン関数
// ======================================================
function generateWeeklyReview() {
  const data = collectCurrentWeekData_();
  const lastWeek = getLastWeekSnapshot_();
  const review = buildReviewText_(data, lastWeek);

  writeReviewToSheet_(review, data);
  sendEmailNotification_(review);
  takeSnapshot_(data);  // 来週の前週比計算用
}

// ======================================================
// データ収集
// ======================================================
function collectCurrentWeekData_() {
  const ss = SpreadsheetApp.getActive();
  return {
    date: new Date(),
    spot: readSpotData_(ss),
    fx: readFXData_(ss),
    mtd: readMonthToDate_(ss),
    scenarios: runFutureSimulations_(ss)
  };
}

function readSpotData_(ss) {
  // 現物_月次 シートから今週の残高・通貨別内訳
  // return { totalUSD, holdings: [{asset, principal, current, profit, yieldPct}] }
  return { totalUSD: 0, holdings: [] };  // プレースホルダ
}

function readFXData_(ss) {
  // FX_月次 シートから今週の残高・統計
  return {
    balance: 0,          // 真正資産
    equity: 0,           // クレジット込み
    credit: 0,           // 参考値
    weeklyTrades: 0,
    winRate: 0,
    profitFactor: 0,
    rrRatio: 0,
    maxDDPct: 0,
    expectedPayoff: 0,
    avgLotSize: 0,
    maxLotSize: 0
  };
}

function readMonthToDate_(ss) {
  return {
    spotYieldMTD: 0,
    fxYieldMTD: 0,
    totalNetProfit: 0,
    totalExpenses: 0
  };
}

function runFutureSimulations_(ss) {
  // 3シナリオ（A/B/C）× 12ヶ月後の予測総資産
  return {
    scenarioA: 0,  // ハイ
    scenarioB: 0,  // ミドル
    scenarioC: 0   // ロー
  };
}

// ======================================================
// レビュー文面生成
// ======================================================
function buildReviewText_(data, lastWeek) {
  const weekNum = getWeekNumber_(data.date);
  const lines = [];

  lines.push(`📊 ${data.date.getFullYear()}年${data.date.getMonth() + 1}月第${weekNum}週 運用レビュー`);
  lines.push('━━━━━━━━━━━━━━━━━━━━━━');
  lines.push('');

  // ----- 総資産 -----
  const totalCurr = data.spot.totalUSD + data.fx.balance;
  const totalPrev = lastWeek ? (lastWeek.spot.totalUSD + lastWeek.fx.balance) : totalCurr;
  const totalDiff = totalCurr - totalPrev;
  const totalPct = totalPrev > 0 ? (totalDiff / totalPrev * 100) : 0;

  lines.push(`💰 総資産：$${fmt_(totalPrev)} → $${fmt_(totalCurr)}（${signed_(totalDiff)} / ${signed_(totalPct, 2)}%）`);
  lines.push(`  現物（MEXC）：$${fmt_(data.spot.totalUSD)}`);
  lines.push(`  FX（BigBoss Balance）：$${fmt_(data.fx.balance)}`);
  lines.push(`  ※ FX Equity $${fmt_(data.fx.equity)} にはクレジット$${fmt_(data.fx.credit)}含む（出金不可）`);
  lines.push('');

  // ----- 現物 -----
  lines.push('📈 現物（MEXC／AIグリッド）通貨別内訳');
  lines.push('  ┌──────────┬────────┬────────┬────────┬────────┐');
  lines.push('  │ 通貨     │ 元本($) │ 現在($) │ 利益($) │ 利回り │');
  lines.push('  ├──────────┼────────┼────────┼────────┼────────┤');
  data.spot.holdings.forEach(h => {
    lines.push(`  │ ${pad_(h.asset, 8)} │ ${fmt_(h.principal, 8)} │ ${fmt_(h.current, 8)} │ ${signed_(h.profit, 2, 8)} │ ${signed_(h.yieldPct, 2, 6)}% │`);
  });
  lines.push('  └──────────┴────────┴────────┴────────┴────────┘');
  lines.push('');

  // ----- FX -----
  lines.push('💱 FX（BigBoss／MelodyTrade）');
  lines.push('  📊 週次パフォーマンス');
  lines.push(`    新規トレード数：${data.fx.weeklyTrades}`);
  lines.push(`    勝率：${fmt_(data.fx.winRate, 2)}%（目標${TARGETS.fxWinRate}%★）${judge_(data.fx.winRate, TARGETS.fxWinRate)}`);
  lines.push(`    PF★：${fmt_(data.fx.profitFactor, 2)}（目標${TARGETS.fxProfitFactor}）${judge_(data.fx.profitFactor, TARGETS.fxProfitFactor)}`);
  lines.push(`    RR比★：${fmt_(data.fx.rrRatio, 2)}（目標${TARGETS.fxRRRatio}）${judge_(data.fx.rrRatio, TARGETS.fxRRRatio)}`);
  lines.push(`    最大DD：${fmt_(data.fx.maxDDPct, 2)}%（目標${TARGETS.fxMaxDD}%以下）${judge_(TARGETS.fxMaxDD, data.fx.maxDDPct)}`);
  lines.push(`    期待値：$${fmt_(data.fx.expectedPayoff, 2)}`);
  lines.push('');

  // ----- ロット分析 -----
  lines.push('  💡 ロット分析');
  const lotTable = buildLotAnalysis_(data.fx.balance, 6);  // 平均SL6pips
  lines.push(`    今週の平均ロット：${fmt_(data.fx.avgLotSize, 2)}`);
  lines.push(`    今週の最大ロット：${fmt_(data.fx.maxLotSize, 2)}`);
  lines.push('    元本ベース推奨ロット：');
  lotTable.forEach(s => {
    lines.push(`      リスク${s.riskPct}%：${fmt_(s.maxLot, 2)}ロット（最大損失$${fmt_(s.maxLoss, 0)}）[${s.label}]`);
  });
  lines.push('');

  // ----- 月次目標達成 -----
  lines.push('🎯 月次目標達成状況');
  lines.push(`  FX 月次利回り：実績${fmt_(data.mtd.fxYieldMTD, 2)}% / 目標${TARGETS.fxMonthlyYield}%（達成率${progress_(data.mtd.fxYieldMTD, TARGETS.fxMonthlyYield)}）`);
  lines.push(`  現物 月次利回り：実績${fmt_(data.mtd.spotYieldMTD, 2)}% / 目標${TARGETS.spotMonthlyYield}%（達成率${progress_(data.mtd.spotYieldMTD, TARGETS.spotMonthlyYield)}）`);
  lines.push(`  純利益 vs SageMaster費：$${fmt_(data.mtd.totalNetProfit, 0)} vs $${TARGETS.sageMasterFee}`);
  lines.push('');

  // ----- 配信者スコア -----
  lines.push('👤 配信者スコア（MelodyTrade - TP Ladder）');
  const scores = calculateProviderScores_(data.fx);
  lines.push(`  収益性：${scores.profitability}/10`);
  lines.push(`  安定性：${scores.stability}/10`);
  lines.push(`  持続性：${scores.sustainability}/10`);
  lines.push(`  総合：${scores.total}/30（判定：${scoreLabel_(scores.total)}）`);
  lines.push('');

  // ----- 将来予測 -----
  lines.push('🎲 将来予測（12ヶ月後）');
  lines.push(`  Aシナリオ（ハイ）：$${fmt_(data.scenarios.scenarioA, 0)}`);
  lines.push(`  Bシナリオ（ミドル）：$${fmt_(data.scenarios.scenarioB, 0)}`);
  lines.push(`  Cシナリオ（ロー）：$${fmt_(data.scenarios.scenarioC, 0)}`);
  lines.push('');

  // ----- アラート -----
  const alerts = detectAlerts_(data);
  if (alerts.length > 0) {
    lines.push('⚠️ アラート');
    alerts.forEach(a => lines.push(`  ${a}`));
    lines.push('');
  }

  // ----- アクション -----
  lines.push('💡 来週のアクション');
  suggestActions_(data).forEach(a => lines.push(`  ${a}`));
  lines.push('');

  // ----- 用語解説 -----
  lines.push('📖 今週の用語解説');
  lines.push('  ★ PF（プロフィットファクター）：総利益÷総損失。1超えで黒字戦略');
  lines.push('  ★ RR比：平均利益÷平均損失。勝率と並ぶ重要指標');
  lines.push('  → 全用語集は スプシ「📖 FX用語集」シート参照');

  return lines.join('\n');
}

// ======================================================
// 判定・アラート・アクション提案
// ======================================================
function detectAlerts_(data) {
  const alerts = [];
  if (data.fx.profitFactor < 1.0 && data.fx.weeklyTrades >= 20) {
    alerts.push('🔴 FX：PFが1.0を下回る。戦略見直し検討');
  }
  const breakEven = 1 / (1 + data.fx.rrRatio) * 100;
  if (data.fx.winRate < breakEven) {
    alerts.push(`🔴 FX：損益分岐勝率${fmt_(breakEven, 1)}%未達（実際${fmt_(data.fx.winRate, 1)}%）`);
  }
  if (data.fx.maxDDPct > TARGETS.fxMaxDD) {
    alerts.push(`🟡 FX：最大DD${fmt_(data.fx.maxDDPct, 2)}%が目標${TARGETS.fxMaxDD}%超過`);
  }
  return alerts;
}

function suggestActions_(data) {
  const actions = [];
  if (data.fx.profitFactor < 1.0 && data.fx.weeklyTrades >= 30) {
    actions.push('🔸 FX：SageMasterでリスク5%→3%に引き下げ検討');
  } else if (data.fx.profitFactor >= 1.3 && data.fx.profitFactor < 1.5) {
    actions.push('🟢 FX：PF良好。リスク5%継続、1.5達成でリスク7%検討');
  } else if (data.fx.profitFactor >= 1.5) {
    actions.push('🟢 FX：PF優秀。リスク5%→7%引き上げ検討可能');
  } else {
    actions.push('🔸 FX：データ蓄積中。判断保留、様子見継続');
  }

  if (data.spot.totalUSD > 0) {
    actions.push('🔸 現物：このまま継続、月末時点で月利8%達成か確認');
  }

  actions.push(`🔸 経費計上：SageMaster $${TARGETS.sageMasterFee}を当月コストに反映`);
  return actions;
}

function calculateProviderScores_(fx) {
  // 簡易版（詳細ロジックは 01_mt4_html_parser.gs の scoreXXX_ を利用）
  let profitability = 0, stability = 0, sustainability = 0;

  if (fx.profitFactor >= 2.0) profitability = 10;
  else if (fx.profitFactor >= 1.5) profitability = 8;
  else if (fx.profitFactor >= 1.3) profitability = 6;
  else if (fx.profitFactor >= 1.0) profitability = 4;
  else profitability = 2;

  if (fx.winRate >= 70) stability += 3;
  if (fx.rrRatio >= 0.7) stability += 3;
  if (fx.maxDDPct < 10) stability += 3;

  // 持続性は取引数で判定（簡易）
  if (fx.weeklyTrades >= 30) sustainability = 8;
  else if (fx.weeklyTrades >= 10) sustainability = 5;
  else sustainability = 3;

  return {
    profitability: profitability,
    stability: Math.min(stability, 10),
    sustainability: sustainability,
    total: profitability + Math.min(stability, 10) + sustainability
  };
}

function scoreLabel_(total) {
  if (total >= 24) return 'A（優秀・継続）';
  if (total >= 18) return 'B（良好・継続）';
  if (total >= 12) return 'C（要観察）';
  if (total >= 6) return 'D（データ不足 or 改善要）';
  return 'E（解約検討）';
}

// ======================================================
// ユーティリティ
// ======================================================
function fmt_(n, digits, width) {
  if (n === undefined || n === null || isNaN(n)) return '-';
  const d = digits === undefined ? 2 : digits;
  let s = n.toLocaleString('en-US', {
    minimumFractionDigits: d,
    maximumFractionDigits: d
  });
  if (width) s = s.padStart(width);
  return s;
}

function signed_(n, digits, width) {
  const s = fmt_(n, digits, width);
  if (n > 0) return '+' + s.trim();
  return s;
}

function pad_(s, width) {
  return String(s || '').padEnd(width);
}

function judge_(actual, target) {
  if (actual >= target) return '🟢';
  if (actual >= target * 0.85) return '🟡';
  return '🔴';
}

function progress_(actual, target) {
  if (!target) return '-';
  return `${Math.round(actual / target * 100)}%`;
}

function getWeekNumber_(date) {
  const first = new Date(date.getFullYear(), date.getMonth(), 1);
  return Math.ceil((date.getDate() + first.getDay()) / 7);
}

function buildLotAnalysis_(balance, avgSLPips) {
  const scenarios = [2, 3, 5, 7];
  const pipValue = 10;  // XAUUSD
  return scenarios.map(riskPct => {
    const maxLot = Math.floor(balance * riskPct / 100 / (avgSLPips * pipValue) * 100) / 100;
    return {
      riskPct: riskPct,
      maxLot: maxLot,
      maxLoss: balance * riskPct / 100,
      label: riskPct === 2 ? '安全' : riskPct === 3 ? '標準' : riskPct === 5 ? '現状' : '攻め'
    };
  });
}

function getLastWeekSnapshot_() {
  // Script Propertiesから前週スナップショット取得
  const raw = PropertiesService.getScriptProperties().getProperty('LAST_WEEK_SNAPSHOT');
  return raw ? JSON.parse(raw) : null;
}

function takeSnapshot_(data) {
  PropertiesService.getScriptProperties()
    .setProperty('LAST_WEEK_SNAPSHOT', JSON.stringify({
      date: data.date.toISOString(),
      spot: { totalUSD: data.spot.totalUSD },
      fx: { balance: data.fx.balance, equity: data.fx.equity }
    }));
}

function writeReviewToSheet_(review, data) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('週次レビュー') || ss.insertSheet('週次レビュー');
  sheet.insertRowBefore(1);
  sheet.getRange(1, 1).setValue(data.date);
  sheet.getRange(1, 2).setValue(review);
}

function sendEmailNotification_(review) {
  const email = Session.getActiveUser().getEmail();
  MailApp.sendEmail({
    to: email,
    subject: `📊 SageMaster週次レビュー ${new Date().toLocaleDateString('ja-JP')}`,
    body: review
  });
}

// ======================================================
// トリガー設定（初回セットアップ時に手動実行）
// ======================================================
function setupTriggers() {
  // 既存のトリガーを削除
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // 月曜7:00 JSTで週次レビュー
  ScriptApp.newTrigger('generateWeeklyReview')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .inTimezone('Asia/Tokyo')
    .create();

  // 毎日6:00 JSTでMT4/MEXC集計
  ScriptApp.newTrigger('runMT4Aggregation')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .inTimezone('Asia/Tokyo')
    .create();

  ScriptApp.newTrigger('runMEXCAggregation')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .inTimezone('Asia/Tokyo')
    .create();
}
