/**
 * 週次レビュー自動生成
 *
 * 実行：毎週月曜 7:00 JST（トリガー）
 * 出力：
 *   1. スプシ「週次レビュー」シートに追記（timeline式・最新が一番上）
 *   2. 社長のGmailに通知
 *   3. MEXC・MT4データを自動更新後にレビュー作成
 */

const REVIEW_CONFIG = {
  SHEET: '週次レビュー',
  DASH_SHEET: '統合ダッシュボード',
  SPOT_SNAP: '現物_スナップショット',
  FX_SNAP: 'FX_スナップショット'
};

// 目標値（統合ダッシュボードと同期）
const TARGETS = {
  fxMonthlyYield: 0.05,
  fxWinRate: 0.71,
  fxPF: 1.3,
  fxRR: 0.7,
  fxMaxDD: 0.10,
  spotMonthlyYield: 0.10,
  totalMoMGrowth: 0.065,
  sageMasterFee: 149
};

// FX用語集（週次ローテーション）
const GLOSSARY = [
  ['PF（プロフィットファクター）', '総利益÷総損失。1超えで黒字戦略。1.3以上が安定ライン'],
  ['RR比（リスクリワード比）', '平均利益÷平均損失。0.7以上推奨。低いと勝率でカバー必要'],
  ['損益分岐勝率', '1÷(1+RR比)×100%。この勝率を超えないと赤字'],
  ['最大DD（ドローダウン）', '高値からの最大下落率。10%以下が安全圏'],
  ['期待値', '1トレードあたり平均損益。プラスなら続ける意味あり'],
  ['pips', '価格変動の単位。ゴールドは1pip=$0.1'],
  ['レバレッジ', '証拠金の何倍取引可能か。BigBossは最大2222倍'],
  ['証拠金維持率', '有効証拠金÷必要証拠金。100%割れで強制決済'],
  ['スプレッド', '買値と売値の差。実質的な取引手数料'],
  ['スワップ', '通貨間金利差の日割り。長期保有時のコスト/収益'],
  ['TP Ladder', '利確を段階的に設定する手法。MelodyTradeの戦略名'],
  ['SL/TP', 'ストップロス（損切）/テイクプロフィット（利確）']
];

// ======================================================
// メイン関数
// ======================================================
function generateWeeklyReview() {
  const ss = SpreadsheetApp.getActive();

  // 最新データ取得（MEXC・MT4パース）
  try { testWriteHoldingsToSheet(); } catch (e) { Logger.log(`MEXC更新失敗: ${e.message}`); }
  try { testWriteFXToSheet(); } catch (e) { Logger.log(`FX更新失敗: ${e.message}`); }

  const data = collectReviewData_(ss);
  const lastWeek = getLastSnapshot_();
  const review = buildReviewText_(data, lastWeek);

  writeReviewToSheet_(ss, review, data);
  sendEmailNotification_(review);
  saveSnapshot_(data);

  Logger.log('✅ 週次レビュー生成完了');
}

// ======================================================
// データ収集
// ======================================================
function collectReviewData_(ss) {
  const now = new Date();
  const data = {
    date: now,
    weekNum: getWeekNumber_(now),
    spot: { totalUSD: 0, holdings: [] },
    fx: {},
    mtd: { spotYield: 0, fxYield: 0, netProfit: 0 }
  };

  const spotSheet = ss.getSheetByName(REVIEW_CONFIG.SPOT_SNAP);
  if (spotSheet) {
    const rows = spotSheet.getDataRange().getValues();
    for (let i = 1; i < rows.length - 1; i++) {
      const r = rows[i];
      data.spot.holdings.push({
        asset: r[1], amount: r[2], price: r[3], value: r[4], ratio: r[5]
      });
    }
    data.spot.totalUSD = rows[rows.length - 1][4];
  }

  const fxSheet = ss.getSheetByName(REVIEW_CONFIG.FX_SNAP);
  if (fxSheet) {
    const g = (a) => Number(fxSheet.getRange(a).getValue()) || 0;
    data.fx = {
      balance: g('B5'), equity: g('B6'), credit: g('B4'),
      trades: g('B10'), winRate: g('B11'), pf: g('B12'),
      rr: g('B13'), expectedPayoff: g('B14'),
      netProfit: g('B17'), maxDDPct: g('B19'), breakEven: g('B20')
    };
  }

  const dashSheet = ss.getSheetByName(REVIEW_CONFIG.DASH_SHEET);
  if (dashSheet) {
    const months = generateMonthList_();
    const idx = months.findIndex(m => m.year === now.getFullYear() && m.month === now.getMonth() + 1);
    if (idx >= 0) {
      const col = DASH_CONFIG.FIRST_DATA_COL + idx;
      data.mtd.netProfit = Number(dashSheet.getRange(7, col).getValue()) || 0;
      data.mtd.spotYield = Number(dashSheet.getRange(14, col).getValue()) || 0;
      data.mtd.fxYield = Number(dashSheet.getRange(20, col).getValue()) || 0;
    }
  }

  return data;
}

// ======================================================
// レビュー文面生成
// ======================================================
function buildReviewText_(data, lastWeek) {
  const lines = [];
  const total = data.spot.totalUSD + data.fx.balance;
  const totalPrev = lastWeek ? (lastWeek.spotUSD + lastWeek.fxBalance) : total;
  const totalDiff = total - totalPrev;
  const totalPct = totalPrev > 0 ? totalDiff / totalPrev : 0;

  lines.push(`📊 ${data.date.getFullYear()}年${data.date.getMonth() + 1}月第${data.weekNum}週 運用レビュー`);
  lines.push('━━━━━━━━━━━━━━━━━━━━━━');
  lines.push('');

  lines.push(`💰 総資産：$${fmtN_(totalPrev)} → $${fmtN_(total)}（${sgn_(totalDiff)} / ${sgnPct_(totalPct)}）`);
  lines.push(`  現物（MEXC）：$${fmtN_(data.spot.totalUSD)}`);
  lines.push(`  FX Balance：$${fmtN_(data.fx.balance)}（真正資産）`);
  lines.push(`  FX Equity：$${fmtN_(data.fx.equity)}（クレジット$${fmtN_(data.fx.credit)}含む・出金不可）`);
  lines.push('');

  lines.push('📈 現物（MEXC／AIグリッド）');
  if (data.spot.holdings.length > 0) {
    lines.push('  通貨別内訳：');
    data.spot.holdings.slice(0, 8).forEach(h => {
      lines.push(`  ・${h.asset.padEnd(8)}：$${fmtN_(h.value).padStart(8)}（${fmtPct_(h.ratio)}）`);
    });
  }
  lines.push('');

  lines.push('💱 FX（BigBoss／MelodyTrade - TP Ladder）');
  lines.push(`  📊 累計パフォーマンス（${data.fx.trades}トレード）`);
  lines.push(`    勝率：${fmtPct_(data.fx.winRate)}（目標${fmtPct_(TARGETS.fxWinRate)}★）${judge_(data.fx.winRate, TARGETS.fxWinRate)}`);
  lines.push(`    PF★：${data.fx.pf.toFixed(2)}（目標${TARGETS.fxPF}）${judge_(data.fx.pf, TARGETS.fxPF)}`);
  lines.push(`    RR比★：${data.fx.rr.toFixed(2)}（目標${TARGETS.fxRR}）${judge_(data.fx.rr, TARGETS.fxRR)}`);
  lines.push(`    最大DD：${fmtPct_(data.fx.maxDDPct)}（目標${fmtPct_(TARGETS.fxMaxDD)}以下）${judgeMax_(data.fx.maxDDPct, TARGETS.fxMaxDD)}`);
  lines.push(`    期待値：$${data.fx.expectedPayoff.toFixed(2)}/トレード`);
  lines.push(`    損益分岐勝率：${fmtPct_(data.fx.breakEven)}`);
  lines.push('');

  lines.push('🎯 月次目標達成状況');
  lines.push(`  FX 月次利回り：${fmtPct_(data.mtd.fxYield)} / 目標${fmtPct_(TARGETS.fxMonthlyYield)}（達成率${progress_(data.mtd.fxYield, TARGETS.fxMonthlyYield)}）`);
  lines.push(`  現物 月次利回り：${fmtPct_(data.mtd.spotYield)} / 目標${fmtPct_(TARGETS.spotMonthlyYield)}（達成率${progress_(data.mtd.spotYield, TARGETS.spotMonthlyYield)}）`);
  lines.push(`  純利益：$${fmtN_(data.mtd.netProfit)}（SageMaster費$${TARGETS.sageMasterFee}）`);
  lines.push('');

  const alerts = detectAlerts_(data);
  if (alerts.length > 0) {
    lines.push('⚠️ アラート');
    alerts.forEach(a => lines.push(`  ${a}`));
    lines.push('');
  }

  lines.push('💡 来週のアクション');
  suggestActions_(data).forEach(a => lines.push(`  ${a}`));
  lines.push('');

  const wkIdx = Math.floor(data.date.getTime() / (7 * 24 * 60 * 60 * 1000));
  lines.push('📖 今週の用語解説');
  lines.push(`  ★ ${GLOSSARY[wkIdx % GLOSSARY.length][0]}：${GLOSSARY[wkIdx % GLOSSARY.length][1]}`);
  lines.push(`  ★ ${GLOSSARY[(wkIdx + 1) % GLOSSARY.length][0]}：${GLOSSARY[(wkIdx + 1) % GLOSSARY.length][1]}`);

  return lines.join('\n');
}

// ======================================================
// 判定ロジック
// ======================================================
function detectAlerts_(data) {
  const alerts = [];
  if (data.fx.pf < 1.0 && data.fx.trades >= 20) {
    alerts.push(`🔴 FX：PFが1.0を下回る（${data.fx.pf.toFixed(2)}）。赤字戦略の可能性`);
  }
  if (data.fx.winRate < data.fx.breakEven && data.fx.trades >= 10) {
    alerts.push(`🔴 FX：損益分岐勝率${fmtPct_(data.fx.breakEven)}未達（実際${fmtPct_(data.fx.winRate)}）`);
  }
  if (data.fx.maxDDPct > TARGETS.fxMaxDD) {
    alerts.push(`🟡 FX：最大DD${fmtPct_(data.fx.maxDDPct)}が目標${fmtPct_(TARGETS.fxMaxDD)}超過`);
  }
  return alerts;
}

function suggestActions_(data) {
  const actions = [];
  if (data.fx.trades < 30) {
    actions.push('🔸 FX：データ蓄積中（30トレード未満）、様子見継続');
  } else if (data.fx.pf < 1.0) {
    actions.push('🔸 FX：SageMasterでリスク5%→3%に引き下げ検討');
  } else if (data.fx.pf >= 1.5) {
    actions.push('🟢 FX：PF優秀。リスク5%→7%引き上げ検討可能');
  } else if (data.fx.pf >= 1.3) {
    actions.push('🟢 FX：PF良好。現状維持');
  }

  if (data.spot.totalUSD > 0) {
    actions.push('🔸 現物：継続、月末時点で月利10%達成か確認');
  }
  actions.push(`🔸 経費計上：SageMaster $${TARGETS.sageMasterFee}を当月コストに反映`);
  return actions;
}

// ======================================================
// 書き込み・通知
// ======================================================
function writeReviewToSheet_(ss, review, data) {
  let sheet = ss.getSheetByName(REVIEW_CONFIG.SHEET);
  if (!sheet) sheet = ss.insertSheet(REVIEW_CONFIG.SHEET);

  sheet.insertRowsBefore(1, 2);
  sheet.getRange(1, 1).setValue(data.date).setNumberFormat('yyyy/MM/dd HH:mm');
  sheet.getRange(1, 2).setValue(review);
  sheet.getRange(1, 1).setVerticalAlignment('top').setFontWeight('bold').setBackground('#cfe2f3');
  sheet.getRange(1, 2).setVerticalAlignment('top').setWrap(true);
  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 800);
  sheet.setRowHeight(1, 600);
  sheet.getRange(1, 1, 1, 2).setFontSize(10);
}

function sendEmailNotification_(review) {
  try {
    const email = Session.getActiveUser().getEmail();
    MailApp.sendEmail({
      to: email,
      subject: `📊 SageMaster週次レビュー ${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd')}`,
      body: review
    });
  } catch (e) {
    Logger.log(`メール送信失敗: ${e.message}`);
  }
}

// ======================================================
// スナップショット
// ======================================================
function getLastSnapshot_() {
  const raw = PropertiesService.getScriptProperties().getProperty('LAST_REVIEW_SNAPSHOT');
  return raw ? JSON.parse(raw) : null;
}

function saveSnapshot_(data) {
  PropertiesService.getScriptProperties()
    .setProperty('LAST_REVIEW_SNAPSHOT', JSON.stringify({
      date: data.date.toISOString(),
      spotUSD: data.spot.totalUSD,
      fxBalance: data.fx.balance
    }));
}

// ======================================================
// フォーマット
// ======================================================
function fmtN_(n) {
  if (typeof n !== 'number' || isNaN(n)) return '0';
  return n.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
}
function fmtPct_(n) { return (n * 100).toFixed(2) + '%'; }
function sgn_(n) { return (n > 0 ? '+' : '') + fmtN_(n); }
function sgnPct_(n) { return (n > 0 ? '+' : '') + fmtPct_(n); }
function judge_(actual, target) {
  if (actual >= target) return '🟢';
  if (actual >= target * 0.85) return '🟡';
  return '🔴';
}
function judgeMax_(actual, target) {
  if (actual <= target) return '🟢';
  if (actual <= target * 1.15) return '🟡';
  return '🔴';
}
function progress_(actual, target) {
  if (!target) return '-';
  return Math.round(actual / target * 100) + '%';
}
function getWeekNumber_(date) {
  const first = new Date(date.getFullYear(), date.getMonth(), 1);
  return Math.ceil((date.getDate() + first.getDay()) / 7);
}

// ======================================================
// テスト＆トリガー設定
// ======================================================
function testWeeklyReview() {
  generateWeeklyReview();
  Logger.log('テスト完了。スプシ「週次レビュー」シート＋Gmailを確認してください');
}

function setupWeeklyReviewTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'generateWeeklyReview')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('generateWeeklyReview')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .inTimezone('Asia/Tokyo')
    .create();

  Logger.log('✅ 月曜7:00 JSTトリガー設定完了');
}
