/**
 * Amazon Dashboard - 週次AI改善提案（Claude API → Gmail）
 *
 * ## 流れ
 *
 *   collectWeeklyMetrics()   D1/D2S/M3 から直近1週間 vs 前週・前月比を集計
 *     ↓
 *   buildWeeklyPrompt()      集計結果をプロンプト化（KPI比較表 + 商品TOP/BOTTOM）
 *     ↓
 *   callClaude(sonnet-4-6)   改善提案を生成
 *     ↓
 *   sendWeeklyMail()         Gmail（プレーン+HTML）で社長宛送信
 *
 * ## トリガー: 毎週月曜 8:00 (sendWeeklyAiReport)
 *
 * ## 出力フォーマット
 * Claude が以下5セクションのMarkdownを返すよう指示する:
 *   1. 全体サマリー（数値ハイライト3つ）
 *   2. 良かった点 / 改善ポイント
 *   3. 商品別の注力アクション（TOP3）
 *   4. 広告配分の見直し提案
 *   5. 来週の優先タスク（チェックリスト形式）
 */

const WEEKLY_REPORT_DAYS = 7;
const WEEKLY_TOP_N = 5;     // プロンプトに含める上位N商品

/**
 * メイン: 週次AIレポートを生成して Gmail 送信
 */
function sendWeeklyAiReport() {
  const t0 = Date.now();
  Logger.log('===== 週次AIレポート生成 開始 =====');

  const metrics = collectWeeklyMetrics();
  if (metrics.thisWeek.sales === 0) {
    Logger.log('⚠️ 今週の売上が0のためレポートをスキップ');
    return;
  }

  const prompt = buildWeeklyPrompt(metrics);
  Logger.log('プロンプト長: ' + prompt.length + ' 文字');

  const aiBody = callClaude({
    model: CLAUDE_MODELS.WEEKLY,
    system: WEEKLY_SYSTEM_PROMPT,
    prompt: prompt,
    maxTokens: 4096,
    temperature: 0.4,
  });

  sendWeeklyMail(metrics, aiBody);
  Logger.log('✅ 週次AIレポート送信完了 (' + (Date.now() - t0) + 'ms)');
}

const WEEKLY_SYSTEM_PROMPT =
  '株式会社LEVEL1 Amazon物販事業の週次レビューを担当する分析AIです。' +
  '社長への報告として、忖度せず数値根拠に基づき、結論ファースト・行動につながる提案を行ってください。' +
  '出力は日本語Markdown。冗長な前置きは不要です。';

/**
 * D1 / D2S / M3 から直近1週間の指標を集計
 * 比較対象: 前週同曜日（7日前〜13日前）/ 前月同期間
 */
function collectWeeklyMetrics() {
  const today = new Date();
  const fmt = d => Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');

  const thisEnd = new Date(today); thisEnd.setDate(today.getDate() - 1);  // 昨日まで
  const thisStart = new Date(thisEnd); thisStart.setDate(thisEnd.getDate() - (WEEKLY_REPORT_DAYS - 1));
  const lastEnd = new Date(thisStart); lastEnd.setDate(thisStart.getDate() - 1);
  const lastStart = new Date(lastEnd); lastStart.setDate(lastEnd.getDate() - (WEEKLY_REPORT_DAYS - 1));

  const periods = {
    thisWeek: { start: fmt(thisStart), end: fmt(thisEnd) },
    lastWeek: { start: fmt(lastStart), end: fmt(lastEnd) },
  };

  const dailyData = getDailyDataAll();
  const allExpenses = readAllSettlement();

  const thisFiltered = dailyData.filter(d => d.date >= periods.thisWeek.start && d.date <= periods.thisWeek.end);
  const lastFiltered = dailyData.filter(d => d.date >= periods.lastWeek.start && d.date <= periods.lastWeek.end);

  const thisExp = aggregateExpenses(allExpenses, periods.thisWeek.start, periods.thisWeek.end);
  const lastExp = aggregateExpenses(allExpenses, periods.lastWeek.start, periods.lastWeek.end);

  // 期間サマリー
  const thisSummary = summarizePeriod(thisFiltered, thisExp);
  const lastSummary = summarizePeriod(lastFiltered, lastExp);

  // ASIN別集計（今週）→ TOP / BOTTOM 抽出
  const byAsin = {};
  for (const d of thisFiltered) {
    if (!d.asin) continue;
    if (!byAsin[d.asin]) {
      byAsin[d.asin] = {
        asin: d.asin, name: d.name, category: d.category,
        sales: 0, units: 0, cv: 0, adCost: 0, adSales: 0, cogs: 0,
      };
    }
    const x = byAsin[d.asin];
    x.sales += d.sales; x.units += d.units; x.cv += d.cv;
    x.adCost += d.adCost; x.adSales += d.adSales; x.cogs += d.cogs || 0;
  }
  const asinList = Object.values(byAsin).map(a => {
    a.profit = a.sales - a.cogs - a.adCost;
    a.profitRate = a.sales > 0 ? a.profit / a.sales : 0;
    a.acos = a.adSales > 0 ? a.adCost / a.adSales : 0;
    a.tacos = a.sales > 0 ? a.adCost / a.sales : 0;
    return a;
  });
  asinList.sort((a, b) => b.sales - a.sales);
  const topAsins = asinList.slice(0, WEEKLY_TOP_N);
  const bottomAsins = asinList.filter(a => a.profit < 0).slice(0, WEEKLY_TOP_N);

  // 当月累計（MTD）= 今月1日 〜 昨日
  const mtdStart = fmt(new Date(today.getFullYear(), today.getMonth(), 1));
  const mtdEnd = fmt(thisEnd);
  const mtdFiltered = dailyData.filter(d => d.date >= mtdStart && d.date <= mtdEnd);
  const mtdExp = aggregateExpenses(allExpenses, mtdStart, mtdEnd);
  const mtdSummary = summarizePeriod(mtdFiltered, mtdExp);
  const mtdPeriod = { start: mtdStart, end: mtdEnd };

  // アカウント健全性ログ（D5）を直近7日分だけ読み取り
  let health = { latestScore: null, latestRow: null, issues: [], avgReturn7: null, avgReturn30: null };
  try {
    health = readRecentAccountHealth(periods.thisWeek.start, periods.thisWeek.end);
  } catch (e) {
    Logger.log('⚠️ アカウント健全性ログ読み取り失敗: ' + e.message);
  }

  return {
    periods: Object.assign({}, periods, { mtd: mtdPeriod }),
    thisWeek: thisSummary,
    lastWeek: lastSummary,
    mtd: mtdSummary,
    topAsins,
    bottomAsins,
    activeAsinCount: asinList.length,
    health: health,
  };
}

/**
 * D5 アカウント健全性シートから直近期間のログを集計
 * ヘッダー: 日付 | 総合 | 返品率(直近7日) | 返品率(直近30日) | 注意点
 */
function readRecentAccountHealth(startDate, endDate) {
  const ss = getMainSpreadsheet();
  const sheet = ss.getSheetByName(D5_HEALTH);
  if (!sheet || sheet.getLastRow() <= 1) {
    return { latestScore: null, latestRow: null, issues: [], avgReturn7: null, avgReturn30: null };
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  const inRange = data.filter(row => {
    const d = row[0] instanceof Date
      ? Utilities.formatDate(row[0], 'Asia/Tokyo', 'yyyy-MM-dd')
      : String(row[0]).substring(0, 10);
    return d >= startDate && d <= endDate;
  });

  if (inRange.length === 0) {
    // 期間内のログがない場合は最新行を返す
    const last = data[data.length - 1];
    return {
      latestScore: last[1] != null ? parseFloat(last[1]) : null,
      latestRow: {
        date: last[0] instanceof Date ? Utilities.formatDate(last[0], 'Asia/Tokyo', 'yyyy-MM-dd') : String(last[0]).substring(0, 10),
        score: last[1],
        ret7: last[2],
        ret30: last[3],
        notes: last[4],
      },
      issues: [],
      avgReturn7: null,
      avgReturn30: null,
    };
  }

  // 一番新しい行を基準値として取得
  const latest = inRange[inRange.length - 1];
  const parsePct = v => {
    const s = String(v || '').replace('%', '').trim();
    const n = parseFloat(s);
    return isNaN(n) ? null : n / 100;
  };

  // 期間中に登場した「注意点」のユニーク一覧（異常なし行は除外）
  const issueSet = new Set();
  inRange.forEach(row => {
    const notes = String(row[4] || '');
    if (!notes || notes.includes('異常なし')) return;
    // "/" 区切り要素を個別 issue として採用（アカウント状態プレフィックスはスキップ）
    notes.split('/').map(s => s.trim()).forEach(part => {
      if (!part) return;
      // 数値付きの明確な問題だけ拾う
      if (/件|日|率|急増|非参加|切れ|取得失敗/.test(part)) issueSet.add(part);
    });
  });

  // 期間中の返品率平均
  const ret7Vals = inRange.map(r => parsePct(r[2])).filter(v => v != null);
  const ret30Vals = inRange.map(r => parsePct(r[3])).filter(v => v != null);
  const avg = arr => arr.length > 0 ? arr.reduce((s, x) => s + x, 0) / arr.length : null;

  return {
    latestScore: latest[1] != null ? parseFloat(latest[1]) : null,
    latestRow: {
      date: latest[0] instanceof Date ? Utilities.formatDate(latest[0], 'Asia/Tokyo', 'yyyy-MM-dd') : String(latest[0]).substring(0, 10),
      score: latest[1],
      ret7: latest[2],
      ret30: latest[3],
      notes: latest[4],
    },
    issues: Array.from(issueSet),
    avgReturn7: avg(ret7Vals),
    avgReturn30: avg(ret30Vals),
  };
}

function summarizePeriod(daily, expenses) {
  let sales = 0, cv = 0, units = 0, sessions = 0, pv = 0, adCost = 0, adSales = 0, cogs = 0;
  for (const d of daily) {
    sales += d.sales; cv += d.cv; units += d.units;
    sessions += d.sessions; pv += d.pv;
    adCost += d.adCost; adSales += d.adSales; cogs += d.cogs || 0;
  }
  // Settlement 比率推定で経費を反映（aggregateExpenses 由来の既存ロジックを再利用）
  const estimated = estimateExpensesFromRate(sales, expenses);
  const commission = estimated.commission;
  const otherExpense = estimated.other;
  const grossProfit = sales - cogs - commission - otherExpense;
  const profit = grossProfit - adCost;
  return {
    sales, cv, units, sessions, pv,
    adCost, adSales, cogs, commission, otherExpense,
    grossProfit, profit,
    grossMargin: sales > 0 ? grossProfit / sales : 0,
    profitMargin: sales > 0 ? profit / sales : 0,
    tacos: sales > 0 ? adCost / sales : 0,
    acos: adSales > 0 ? adCost / adSales : 0,
    roas: adCost > 0 ? sales / adCost : 0,
    cvr: sessions > 0 ? cv / sessions : 0,
  };
}

/**
 * Claude に渡すプロンプトを組み立て（人間も読める指標表）
 */
function buildWeeklyPrompt(m) {
  const fmtN = n => (n == null) ? '-' : Math.round(n).toLocaleString();
  const fmtP = (n, digits) => (n == null) ? '-' : (n * 100).toFixed(digits == null ? 1 : digits) + '%';
  const fmtR = n => (n == null) ? '-' : n.toFixed(2);
  const dPct = (a, b) => b > 0 ? (((a - b) / b) * 100).toFixed(1) + '%' : '-';

  const t = m.thisWeek, l = m.lastWeek;

  let s = '';
  s += '## 期間\n';
  s += '- 今週: ' + m.periods.thisWeek.start + ' 〜 ' + m.periods.thisWeek.end + '\n';
  s += '- 前週: ' + m.periods.lastWeek.start + ' 〜 ' + m.periods.lastWeek.end + '\n';
  if (m.periods.mtd) {
    s += '- 当月累計: ' + m.periods.mtd.start + ' 〜 ' + m.periods.mtd.end + '\n';
  }
  s += '- アクティブASIN: ' + m.activeAsinCount + '\n\n';

  // 売上 / 利益 / 広告費 — 今週・先週・先週比・当月累計
  if (m.mtd) {
    s += '## 売上・利益・広告費 — 主要指標サマリー\n\n';
    s += '| 指標 | 今週 | 先週 | 先週比 | 当月累計 |\n|---|---:|---:|---:|---:|\n';
    const focus = [
      ['売上',   t.sales,  l.sales,  m.mtd.sales],
      ['利益（広告後）', t.profit, l.profit, m.mtd.profit],
      ['広告費', t.adCost, l.adCost, m.mtd.adCost],
    ];
    focus.forEach(r => {
      s += '| ' + [
        r[0], fmtN(r[1]), fmtN(r[2]), dPct(r[1], r[2]), fmtN(r[3]),
      ].join(' | ') + ' |\n';
    });
    s += '\n';
  }

  s += '## 全体KPI（今週 vs 前週）\n\n';
  s += '| 指標 | 今週 | 前週 | 前週比 |\n|---|---:|---:|---:|\n';
  const rows = [
    ['売上', fmtN(t.sales), fmtN(l.sales), dPct(t.sales, l.sales)],
    ['CV(注文件数)', fmtN(t.cv), fmtN(l.cv), dPct(t.cv, l.cv)],
    ['注文点数', fmtN(t.units), fmtN(l.units), dPct(t.units, l.units)],
    ['セッション数', fmtN(t.sessions), fmtN(l.sessions), dPct(t.sessions, l.sessions)],
    ['CVR', fmtP(t.cvr, 2), fmtP(l.cvr, 2), '-'],
    ['広告費', fmtN(t.adCost), fmtN(l.adCost), dPct(t.adCost, l.adCost)],
    ['広告売上', fmtN(t.adSales), fmtN(l.adSales), dPct(t.adSales, l.adSales)],
    ['ACOS', fmtP(t.acos), fmtP(l.acos), '-'],
    ['TACOS', fmtP(t.tacos), fmtP(l.tacos), '-'],
    ['ROAS', fmtR(t.roas), fmtR(l.roas), '-'],
    ['仕入原価', fmtN(t.cogs), fmtN(l.cogs), dPct(t.cogs, l.cogs)],
    ['販売手数料(推定)', fmtN(t.commission), fmtN(l.commission), dPct(t.commission, l.commission)],
    ['粗利率', fmtP(t.grossMargin), fmtP(l.grossMargin), '-'],
    ['利益（広告後）', fmtN(t.profit), fmtN(l.profit), dPct(t.profit, l.profit)],
    ['利益率', fmtP(t.profitMargin), fmtP(l.profitMargin), '-'],
  ];
  rows.forEach(r => { s += '| ' + r.join(' | ') + ' |\n'; });

  s += '\n## 売上TOP' + m.topAsins.length + '商品（今週）\n\n';
  s += '| ASIN | 商品名 | カテゴリ | 売上 | 点数 | 広告費 | 利益 | 利益率 | ACOS |\n|---|---|---|---:|---:|---:|---:|---:|---:|\n';
  m.topAsins.forEach(a => {
    s += '| ' + [
      a.asin, (a.name || '').substring(0, 20), a.category || '-',
      fmtN(a.sales), fmtN(a.units), fmtN(a.adCost),
      fmtN(a.profit), fmtP(a.profitRate), fmtP(a.acos),
    ].join(' | ') + ' |\n';
  });

  if (m.bottomAsins.length > 0) {
    s += '\n## 利益マイナス商品（今週・上位' + m.bottomAsins.length + '）\n\n';
    s += '| ASIN | 商品名 | 売上 | 広告費 | 利益 | ACOS |\n|---|---|---:|---:|---:|---:|\n';
    m.bottomAsins.forEach(a => {
      s += '| ' + [
        a.asin, (a.name || '').substring(0, 20),
        fmtN(a.sales), fmtN(a.adCost), fmtN(a.profit), fmtP(a.acos),
      ].join(' | ') + ' |\n';
    });
  }

  // アカウント健全性（D5 より）
  const h = m.health || { latestScore: null, issues: [], avgReturn7: null, avgReturn30: null, latestRow: null };
  s += '\n## 🏥 アカウント健全性（今週の状態）\n\n';
  if (h.latestScore == null && !h.latestRow) {
    s += '- ログ未取得\n';
  } else {
    const score = h.latestScore != null ? h.latestScore : '-';
    const ret7 = h.avgReturn7 != null ? fmtP(h.avgReturn7, 2) : (h.latestRow ? h.latestRow.ret7 : '-');
    const ret30 = h.avgReturn30 != null ? fmtP(h.avgReturn30, 2) : (h.latestRow ? h.latestRow.ret30 : '-');
    s += '- 最新スコア: ' + score + ' / 100\n';
    s += '- 直近7日返品率: ' + ret7 + ' / 直近30日返品率: ' + ret30 + '\n';
    if (h.issues && h.issues.length > 0) {
      s += '- 期間中に記録された注意点（重複排除）:\n';
      h.issues.forEach(i => { s += '  - ' + i + '\n'; });
    } else {
      s += '- 期間中の注意事項: 特になし\n';
    }
  }

  s += '\n---\n\n';
  s += '上記データを踏まえ、以下の構成でMarkdownレポートを出力してください。\n\n';
  s += '## 1. 今週のハイライト\n';
  s += '（売上・利益・広告費の3点について、**「今週 / 先週比 / 当月累計」を必ず1行で併記** してください。例: ' +
       '「売上: 今週¥X (先週比+Y%) / 当月累計¥Z」）\n\n';
  s += '## 2. 良かった点 / 課題\n（箇条書きそれぞれ2〜4点）\n\n';
  s += '## 3. 商品別アクション（TOP3）\n（ASIN/商品名・状況・推奨アクション・優先度A/B/C）\n\n';
  s += '## 4. 広告配分の提案\n（ACOS/ROAS/オーガニック比率の観点で、増減すべきASIN・キャンペーン方針）\n\n';
  s += '## 5. アカウント健全性の評価\n（スコア・返品率・注意点を踏まえた状態診断と対応要否。即対応/様子見/問題なし のいずれか + 具体策）\n\n';
  s += '## 6. 来週の優先タスク\n（チェックリスト形式・5件以内）\n';
  return s;
}

/**
 * Gmail送信（プレーン本文 + Markdown→簡易HTML 化）
 */
function sendWeeklyMail(metrics, aiBody) {
  const to = getCredential('GMAIL_TO');
  const period = metrics.periods.thisWeek;
  const mtdPeriod = metrics.periods.mtd;
  const mtd = metrics.mtd;
  const subject = '【Amazon週次レポート】' + period.start + '〜' + period.end +
    '（売上 ' + Math.round(metrics.thisWeek.sales).toLocaleString() + '円）';

  const fmtY = n => Math.round(n).toLocaleString() + ' 円';
  const dPct = (a, b) => b > 0 ? (((a - b) / b) * 100).toFixed(1) + '%' : '-';
  const t = metrics.thisWeek, l = metrics.lastWeek;

  let summaryText =
    '■ 集計期間（今週）: ' + period.start + ' 〜 ' + period.end + '\n' +
    '■ 売上: ' + fmtY(t.sales) + '\n' +
    '■ 利益: ' + fmtY(t.profit) + ' (利益率 ' + (t.profitMargin * 100).toFixed(1) + '%)\n' +
    '■ 広告費: ' + fmtY(t.adCost) + ' (TACOS ' + (t.tacos * 100).toFixed(1) + '%)\n';

  if (mtd && mtdPeriod) {
    summaryText +=
      '\n● 主要3指標 (今週 / 先週 / 先週比 / 当月累計)\n' +
      '  集計期間（当月累計）: ' + mtdPeriod.start + ' 〜 ' + mtdPeriod.end + '\n' +
      '  売上 : ' + fmtY(t.sales) + ' / ' + fmtY(l.sales) + ' / ' + dPct(t.sales, l.sales) + ' / ' + fmtY(mtd.sales) + '\n' +
      '  利益 : ' + fmtY(t.profit) + ' / ' + fmtY(l.profit) + ' / ' + dPct(t.profit, l.profit) + ' / ' + fmtY(mtd.profit) + '\n' +
      '  広告費: ' + fmtY(t.adCost) + ' / ' + fmtY(l.adCost) + ' / ' + dPct(t.adCost, l.adCost) + ' / ' + fmtY(mtd.adCost) + '\n';
  }

  summaryText +=
    '\n------ Claude AI 改善提案 ------\n\n' + aiBody +
    '\n\n--\nAmazon Dashboard 自動配信（' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') + '）';

  const htmlBody = '<pre style="font-family: -apple-system, BlinkMacSystemFont, sans-serif; ' +
    'white-space: pre-wrap; line-height: 1.6;">' + escapeHtml(summaryText) + '</pre>';

  GmailApp.sendEmail(to, subject, summaryText, { htmlBody: htmlBody, name: 'Amazon Dashboard', charset: 'UTF-8' });
  Logger.log('📧 Gmail送信完了 → ' + to);
}

function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/**
 * テスト: プロンプトのみ生成してログ出力（API呼び出しなし）
 */
function testWeeklyPromptOnly() {
  const m = collectWeeklyMetrics();
  Logger.log(buildWeeklyPrompt(m));
}

/**
 * テスト: 実際にAPI叩いてレポート本文をログ出力（メール送信なし）
 */
function testWeeklyAiReportNoMail() {
  const m = collectWeeklyMetrics();
  if (m.thisWeek.sales === 0) { Logger.log('売上0のため中止'); return; }
  const prompt = buildWeeklyPrompt(m);
  const body = callClaude({
    model: CLAUDE_MODELS.WEEKLY,
    system: WEEKLY_SYSTEM_PROMPT,
    prompt: prompt,
    maxTokens: 4096,
    temperature: 0.4,
  });
  Logger.log('===== Claude応答 =====\n' + body);
}
