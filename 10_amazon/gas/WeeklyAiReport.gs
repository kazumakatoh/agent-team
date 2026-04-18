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

  return {
    periods,
    thisWeek: thisSummary,
    lastWeek: lastSummary,
    topAsins,
    bottomAsins,
    activeAsinCount: asinList.length,
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
  s += '- アクティブASIN: ' + m.activeAsinCount + '\n\n';

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

  s += '\n---\n\n';
  s += '上記データを踏まえ、以下の構成でMarkdownレポートを出力してください。\n\n';
  s += '## 1. 今週のハイライト\n（売上・利益・広告効率の主要3つを1〜2行で）\n\n';
  s += '## 2. 良かった点 / 課題\n（箇条書きそれぞれ2〜4点）\n\n';
  s += '## 3. 商品別アクション（TOP3）\n（ASIN/商品名・状況・推奨アクション・優先度A/B/C）\n\n';
  s += '## 4. 広告配分の提案\n（ACOS/ROAS/オーガニック比率の観点で、増減すべきASIN・キャンペーン方針）\n\n';
  s += '## 5. 来週の優先タスク\n（チェックリスト形式・5件以内）\n';
  return s;
}

/**
 * Gmail送信（プレーン本文 + Markdown→簡易HTML 化）
 */
function sendWeeklyMail(metrics, aiBody) {
  const to = getCredential('GMAIL_TO');
  const period = metrics.periods.thisWeek;
  const subject = '【Amazon週次レポート】' + period.start + '〜' + period.end +
    '（売上 ' + Math.round(metrics.thisWeek.sales).toLocaleString() + '円）';

  const summaryText =
    '■ 集計期間: ' + period.start + ' 〜 ' + period.end + '\n' +
    '■ 売上: ' + Math.round(metrics.thisWeek.sales).toLocaleString() + ' 円\n' +
    '■ 利益: ' + Math.round(metrics.thisWeek.profit).toLocaleString() + ' 円' +
    ' (利益率 ' + (metrics.thisWeek.profitMargin * 100).toFixed(1) + '%)\n' +
    '■ 広告費: ' + Math.round(metrics.thisWeek.adCost).toLocaleString() + ' 円' +
    ' (TACOS ' + (metrics.thisWeek.tacos * 100).toFixed(1) + '%)\n\n' +
    '------ Claude AI 改善提案 ------\n\n' + aiBody +
    '\n\n--\nAmazon Dashboard 自動配信（' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') + '）';

  const htmlBody = '<pre style="font-family: -apple-system, BlinkMacSystemFont, sans-serif; ' +
    'white-space: pre-wrap; line-height: 1.6;">' + escapeHtml(summaryText) + '</pre>';

  GmailApp.sendEmail(to, subject, summaryText, { htmlBody: htmlBody, name: 'Amazon Dashboard' });
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
