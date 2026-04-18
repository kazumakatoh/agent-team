/**
 * 配信者スコアカード（MelodyTrade - TP Ladder）
 * 3軸評価：収益性 / 安定性 / 持続性 → 30点満点 → A〜E判定
 *
 * 評価ロジック：
 *   - データから自動算出
 *   - 月末判定で継続/解約意思決定
 */

const SCORECARD_CONFIG = {
  SHEET: '配信者スコアカード',
  PROVIDER_NAME: 'MelodyTrade - TP Ladder',
  STRATEGY: 'XAUUSD（ゴールド）／TP Ladder／リスク5%'
};

// ======================================================
// 構築
// ======================================================
function buildProviderScorecard() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(SCORECARD_CONFIG.SHEET);
  if (!sheet) sheet = ss.insertSheet(SCORECARD_CONFIG.SHEET);
  else sheet.clear();

  const fxSheet = ss.getSheetByName('FX_スナップショット');
  if (!fxSheet) {
    Logger.log('❌ FX_スナップショットがありません');
    return;
  }
  const g = (a) => Number(fxSheet.getRange(a).getValue()) || 0;

  const stats = {
    trades: g('B10'),
    winRate: g('B11') * 100,
    pf: g('B12'),
    rr: g('B13'),
    expectedPayoff: g('B14'),
    grossProfit: g('B15'),
    grossLoss: g('B16'),
    netProfit: g('B17'),
    maxDD: g('B18'),
    maxDDPct: g('B19') * 100,
    breakEven: g('B20') * 100
  };

  const scores = calculateScores_(stats);

  // ヘッダー
  sheet.getRange('A1').setValue('配信者:').setFontWeight('bold');
  sheet.getRange('B1').setValue(SCORECARD_CONFIG.PROVIDER_NAME).setFontWeight('bold');
  sheet.getRange('D1').setValue('戦略:').setFontWeight('bold');
  sheet.getRange('E1').setValue(SCORECARD_CONFIG.STRATEGY);
  sheet.getRange('A2').setValue('集計日時:').setFontWeight('bold');
  sheet.getRange('B2').setValue(new Date()).setNumberFormat('yyyy/MM/dd HH:mm:ss');

  // 総合判定
  setScoreSection_(sheet, 4, '🏆 総合判定');
  sheet.getRange('A5').setValue('総合スコア').setFontWeight('bold').setBackground('#f3f3f3');
  sheet.getRange('B5').setValue(scores.total + ' / 30').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('A6').setValue('ランク').setFontWeight('bold').setBackground('#f3f3f3');
  sheet.getRange('B6').setValue(scores.rank).setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground(scores.rankColor);
  sheet.getRange('A7').setValue('判定').setFontWeight('bold').setBackground('#f3f3f3');
  sheet.getRange('B7').setValue(scores.judgement);

  // 軸1：収益性
  setScoreSection_(sheet, 9, '💰 収益性（10点満点）');
  const profRows = [
    ['プロフィットファクター（PF）', stats.pf.toFixed(2), '≥1.0', scores.pf, judgeS_(stats.pf, 1.0)],
    ['月次純利益', '$' + stats.netProfit.toFixed(2), '> $0', scores.netProfit, judgeS_(stats.netProfit, 0)],
    ['期待値（$/トレード）', '$' + stats.expectedPayoff.toFixed(2), '> $0', scores.expectedPayoff, judgeS_(stats.expectedPayoff, 0)],
    ['', '', '', '', ''],
    ['収益性スコア', '', '', scores.profitability + ' / 10', '']
  ];
  sheet.getRange(10, 1, profRows.length, 5).setValues(profRows);
  sheet.getRange(14, 1, 1, 5).setFontWeight('bold').setBackground('#fff2cc');

  // 軸2：安定性
  setScoreSection_(sheet, 16, '🛡️ 安定性（10点満点）');
  const stabRows = [
    ['勝率', stats.winRate.toFixed(2) + '%', '≥70%', scores.winRate, judgeS_(stats.winRate, 70)],
    ['RR比', stats.rr.toFixed(2), '≥0.7', scores.rr, judgeS_(stats.rr, 0.7)],
    ['最大DD', stats.maxDDPct.toFixed(2) + '%', '<10%', scores.dd, judgeMaxS_(stats.maxDDPct, 10)],
    ['損益分岐勝率', stats.breakEven.toFixed(2) + '%', '実勝率と比較', '', judgeS_(stats.winRate, stats.breakEven)],
    ['', '', '', '', ''],
    ['安定性スコア', '', '', scores.stability + ' / 10', '']
  ];
  sheet.getRange(17, 1, stabRows.length, 5).setValues(stabRows);
  sheet.getRange(22, 1, 1, 5).setFontWeight('bold').setBackground('#fff2cc');

  // 軸3：持続性
  setScoreSection_(sheet, 24, '⏳ 持続性（10点満点）');
  const sustRows = [
    ['総トレード数', stats.trades, '≥100', scores.tradeCount, judgeS_(stats.trades, 100)],
    ['Gross Profit', '$' + stats.grossProfit.toFixed(2), '—', '', ''],
    ['Gross Loss', '$' + stats.grossLoss.toFixed(2), '—', '', ''],
    ['データ蓄積期間', '4日（2026/4/13〜）', '≥30日', scores.dataPeriod, judgeS_(4, 30)],
    ['', '', '', '', ''],
    ['持続性スコア', '', '', scores.sustainability + ' / 10', '']
  ];
  sheet.getRange(25, 1, sustRows.length, 5).setValues(sustRows);
  sheet.getRange(30, 1, 1, 5).setFontWeight('bold').setBackground('#fff2cc');

  // 推奨アクション
  setScoreSection_(sheet, 32, '💡 推奨アクション');
  sheet.getRange(33, 1).setValue(scores.action);

  // 継続判定ライン
  setScoreSection_(sheet, 35, '⚖️ 継続判定ライン');
  const judgeRows = [
    ['初回判定', '2026年4月末（運用開始3週間時点）'],
    ['継続ライン', 'PF > 1.0 かつ 月次純利益 > 0'],
    ['要観察ライン', 'PF 0.85〜1.0'],
    ['解約検討ライン', '3ヶ月連続でPF < 1.0']
  ];
  sheet.getRange(36, 1, judgeRows.length, 2).setValues(judgeRows);

  applyScorecardFormats_(sheet);
  Logger.log(`✅ スコアカード構築完了：${scores.total}/30（${scores.rank}）`);
}

// ======================================================
// スコア計算
// ======================================================
function calculateScores_(s) {
  let profPF = 0;
  if (s.pf >= 2.0) profPF = 5;
  else if (s.pf >= 1.5) profPF = 4;
  else if (s.pf >= 1.3) profPF = 3;
  else if (s.pf >= 1.0) profPF = 2;

  const profNet = s.netProfit > 0 ? 3 : 0;
  const profExp = s.expectedPayoff > 0 ? 2 : 0;
  const profitability = Math.min(profPF + profNet + profExp, 10);

  const winRateScore = s.winRate >= 70 ? 3 : (s.winRate >= 60 ? 2 : 1);
  const rrScore = s.rr >= 1.0 ? 4 : (s.rr >= 0.7 ? 3 : (s.rr >= 0.5 ? 2 : 1));
  const ddScore = s.maxDDPct < 10 ? 3 : (s.maxDDPct < 15 ? 2 : 1);
  const stability = Math.min(winRateScore + rrScore + ddScore, 10);

  const tradeScore = s.trades >= 100 ? 5 : (s.trades >= 50 ? 3 : (s.trades >= 20 ? 2 : 1));
  const dataPeriodScore = s.trades >= 100 ? 5 : (s.trades >= 30 ? 3 : 1);
  const sustainability = Math.min(tradeScore + dataPeriodScore, 10);

  const total = profitability + stability + sustainability;

  let rank, rankColor, judgement, action;
  if (total >= 24) {
    rank = 'A'; rankColor = '#b6d7a8';
    judgement = '優秀。継続＋ロット増検討可能';
    action = '🟢 PF>1.5維持なら、リスク%を5%→7%へ引き上げ検討';
  } else if (total >= 18) {
    rank = 'B'; rankColor = '#c9daf8';
    judgement = '良好。現状維持';
    action = '🔸 現状リスク5%継続。月末まで様子観察';
  } else if (total >= 12) {
    rank = 'C'; rankColor = '#fff2cc';
    judgement = '要観察。改善傾向の確認が必要';
    action = '🔸 SageMasterリスク5%→3%に引き下げ検討。データ蓄積継続';
  } else if (total >= 6) {
    rank = 'D'; rankColor = '#f4cccc';
    judgement = 'データ不足 or 改善要';
    action = '⚠️ 30トレード達成まで観察。改善なければ戦略見直し';
  } else {
    rank = 'E'; rankColor = '#ea9999';
    judgement = '解約検討';
    action = '🔴 SageMaster解約 or 別配信者への切替検討';
  }

  return {
    pf: profPF, netProfit: profNet, expectedPayoff: profExp,
    winRate: winRateScore, rr: rrScore, dd: ddScore,
    tradeCount: tradeScore, dataPeriod: dataPeriodScore,
    profitability, stability, sustainability, total,
    rank, rankColor, judgement, action
  };
}

// ======================================================
// 補助
// ======================================================
function setScoreSection_(sheet, row, label) {
  sheet.getRange(row, 1, 1, 5).setBackground('#37474f').setFontColor('white').setFontWeight('bold');
  sheet.getRange(row, 1).setValue(label);
}

function applyScorecardFormats_(sheet) {
  const lastRow = sheet.getLastRow();
  sheet.getRange(1, 1, lastRow, 5).setFontSize(10);
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 80);
  sheet.getRange('A:A').setHorizontalAlignment('center');
  sheet.getRange('C:E').setHorizontalAlignment('center');
}

function judgeS_(actual, target) {
  if (actual >= target) return '🟢';
  if (actual >= target * 0.85) return '🟡';
  return '🔴';
}
function judgeMaxS_(actual, target) {
  if (actual <= target) return '🟢';
  if (actual <= target * 1.15) return '🟡';
  return '🔴';
}
