/**
 * Amazon Dashboard - 月次AI戦略レポート（Claude Opus → Gmail）
 *
 * 週次の sonnet 分析より踏み込んだ「戦略立案」を opus に依頼する。
 * 前月の全データを使い、逆算分析（目標利益→必要売上→必要セッション数）、
 * 広告予算の最適配分、撤退・注力の判断材料を作らせる。
 *
 * ## 流れ
 *   collectMonthlyMetrics()    D1/D2S/M3 から前月 vs 前々月 vs 前年同月
 *     ↓
 *   buildMonthlyPrompt()       カテゴリ別 / 商品TOP10 / BOTTOM5 を含む包括データ
 *     ↓
 *   callClaude(opus-4-6)       戦略レポート生成（max_tokens 大きめ）
 *     ↓
 *   sendMonthlyMail()          Gmail 送信
 *
 * ## トリガー: 毎月1日 9:00 (sendMonthlyAiReport)
 */

const MONTHLY_TOP_N = 10;
const MONTHLY_BOTTOM_N = 5;

/**
 * メイン: 月次AI戦略レポートを生成して Gmail 送信
 */
function sendMonthlyAiReport() {
  const t0 = Date.now();
  Logger.log('===== 月次AI戦略レポート生成 開始 =====');

  const metrics = collectMonthlyMetrics();
  if (metrics.lastMonth.sales === 0) {
    Logger.log('⚠️ 前月の売上が0のためレポートをスキップ');
    return;
  }

  const prompt = buildMonthlyPrompt(metrics);
  Logger.log('プロンプト長: ' + prompt.length + ' 文字');

  const aiBody = callClaude({
    model: CLAUDE_MODELS.MONTHLY,
    system: MONTHLY_SYSTEM_PROMPT,
    prompt: prompt,
    maxTokens: 8192,      // 戦略レポートは長めに
    temperature: 0.5,
  });

  sendMonthlyMail(metrics, aiBody);
  Logger.log('✅ 月次AIレポート送信完了 (' + (Date.now() - t0) + 'ms)');
}

const MONTHLY_SYSTEM_PROMPT =
  '株式会社LEVEL1 Amazon物販事業の月次戦略レビューを担当する経営参謀AIです。' +
  '数値根拠に基づき、忖度せず、意思決定できる具体策を提案してください。' +
  '社長は多忙なため、結論ファースト・箇条書き・アクション単位で構造化した Markdown を出力してください。' +
  '「もし◯◯なら△△すべき」という条件付き提案ではなく、「何をいつまでにやる」を明示すること。';

/**
 * 月次指標を集計（前月 / 前々月 / 前年同月 の3期間比較）
 */
function collectMonthlyMetrics() {
  const today = new Date();
  const fmt = d => Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');

  // 前月
  const lastMonth = {
    start: fmt(new Date(today.getFullYear(), today.getMonth() - 1, 1)),
    end:   fmt(new Date(today.getFullYear(), today.getMonth(), 0)),
  };
  // 前々月
  const prevMonth = {
    start: fmt(new Date(today.getFullYear(), today.getMonth() - 2, 1)),
    end:   fmt(new Date(today.getFullYear(), today.getMonth() - 1, 0)),
  };
  // 前年同月
  const yoyMonth = {
    start: fmt(new Date(today.getFullYear() - 1, today.getMonth() - 1, 1)),
    end:   fmt(new Date(today.getFullYear() - 1, today.getMonth(), 0)),
  };

  const periods = { lastMonth, prevMonth, yoyMonth };

  const dailyData = getDailyDataAll();
  const allExpenses = readAllSettlement();

  const last = filterAndSummarize(dailyData, allExpenses, lastMonth.start, lastMonth.end);
  const prev = filterAndSummarize(dailyData, allExpenses, prevMonth.start, prevMonth.end);
  const yoy  = filterAndSummarize(dailyData, allExpenses, yoyMonth.start, yoyMonth.end);

  // ASIN別（前月） → TOP10 / BOTTOM5
  const byAsin = {};
  for (const d of dailyData) {
    if (d.date < lastMonth.start || d.date > lastMonth.end) continue;
    if (!d.asin) continue;
    if (!byAsin[d.asin]) {
      byAsin[d.asin] = {
        asin: d.asin, name: d.name, category: d.category,
        sales: 0, units: 0, cv: 0, sessions: 0,
        adCost: 0, adSales: 0, cogs: 0,
      };
    }
    const x = byAsin[d.asin];
    x.sales += d.sales; x.units += d.units; x.cv += d.cv;
    x.sessions += d.sessions;
    x.adCost += d.adCost; x.adSales += d.adSales; x.cogs += d.cogs || 0;
  }
  const asinList = Object.values(byAsin).map(a => {
    a.profit = a.sales - a.cogs - a.adCost;
    a.profitRate = a.sales > 0 ? a.profit / a.sales : 0;
    a.acos = a.adSales > 0 ? a.adCost / a.adSales : 0;
    a.tacos = a.sales > 0 ? a.adCost / a.sales : 0;
    a.roas = a.adCost > 0 ? a.adSales / a.adCost : 0;
    a.cvr = a.sessions > 0 ? a.cv / a.sessions : 0;
    return a;
  });
  asinList.sort((a, b) => b.sales - a.sales);
  const topAsins = asinList.slice(0, MONTHLY_TOP_N);
  const bottomAsins = asinList.filter(a => a.profit < 0).sort((a, b) => a.profit - b.profit).slice(0, MONTHLY_BOTTOM_N);

  // カテゴリ別（前月）
  const byCategory = {};
  for (const d of dailyData) {
    if (d.date < lastMonth.start || d.date > lastMonth.end) continue;
    const cat = d.category || '(未分類)';
    if (!byCategory[cat]) byCategory[cat] = { category: cat, sales: 0, units: 0, adCost: 0, cogs: 0 };
    byCategory[cat].sales += d.sales;
    byCategory[cat].units += d.units;
    byCategory[cat].adCost += d.adCost;
    byCategory[cat].cogs += d.cogs || 0;
  }
  const categoryList = Object.values(byCategory).map(c => {
    c.profit = c.sales - c.cogs - c.adCost;
    c.profitRate = c.sales > 0 ? c.profit / c.sales : 0;
    c.tacos = c.sales > 0 ? c.adCost / c.sales : 0;
    return c;
  }).sort((a, b) => b.sales - a.sales);

  return {
    periods,
    lastMonth: last,
    prevMonth: prev,
    yoyMonth: yoy,
    topAsins,
    bottomAsins,
    categories: categoryList,
    activeAsinCount: asinList.length,
  };
}

function filterAndSummarize(dailyData, allExpenses, start, end) {
  const filtered = dailyData.filter(d => d.date >= start && d.date <= end);
  const exp = aggregateExpenses(allExpenses, start, end);
  return summarizePeriod(filtered, exp);
}

/**
 * Claude Opus 向けプロンプト。週次より広範 + 戦略立案指示を含む。
 */
function buildMonthlyPrompt(m) {
  const fmtN = n => (n == null) ? '-' : Math.round(n).toLocaleString();
  const fmtP = (n, digits) => (n == null) ? '-' : (n * 100).toFixed(digits == null ? 1 : digits) + '%';
  const fmtR = n => (n == null) ? '-' : n.toFixed(2);
  const dPct = (a, b) => b > 0 ? (((a - b) / b) * 100).toFixed(1) + '%' : '-';

  const L = m.lastMonth, P = m.prevMonth, Y = m.yoyMonth;

  let s = '';
  s += '## 期間\n';
  s += '- 前月: ' + m.periods.lastMonth.start + ' 〜 ' + m.periods.lastMonth.end + '\n';
  s += '- 前々月: ' + m.periods.prevMonth.start + ' 〜 ' + m.periods.prevMonth.end + '\n';
  s += '- 前年同月: ' + m.periods.yoyMonth.start + ' 〜 ' + m.periods.yoyMonth.end + '\n';
  s += '- アクティブASIN: ' + m.activeAsinCount + '\n\n';

  s += '## 月次KPI（前月 vs 前々月 vs 前年同月）\n\n';
  s += '| 指標 | 前月 | 前々月 | 前年同月 | 前月比 | 前年比 |\n|---|---:|---:|---:|---:|---:|\n';
  const rows = [
    ['売上', L.sales, P.sales, Y.sales, fmtN, null],
    ['CV(注文件数)', L.cv, P.cv, Y.cv, fmtN, null],
    ['注文点数', L.units, P.units, Y.units, fmtN, null],
    ['セッション数', L.sessions, P.sessions, Y.sessions, fmtN, null],
    ['広告費', L.adCost, P.adCost, Y.adCost, fmtN, null],
    ['広告売上', L.adSales, P.adSales, Y.adSales, fmtN, null],
    ['仕入原価', L.cogs, P.cogs, Y.cogs, fmtN, null],
    ['販売手数料(推定)', L.commission, P.commission, Y.commission, fmtN, null],
    ['利益', L.profit, P.profit, Y.profit, fmtN, null],
  ];
  rows.forEach(r => {
    s += '| ' + r[0] + ' | ' + r[5](r[1]) + ' | ' + r[5](r[2]) + ' | ' + r[5](r[3]) +
         ' | ' + dPct(r[1], r[2]) + ' | ' + dPct(r[1], r[3]) + ' |\n';
  });
  // 比率系は別枠
  s += '| CVR | ' + fmtP(L.cvr, 2) + ' | ' + fmtP(P.cvr, 2) + ' | ' + fmtP(Y.cvr, 2) + ' | - | - |\n';
  s += '| TACOS | ' + fmtP(L.tacos) + ' | ' + fmtP(P.tacos) + ' | ' + fmtP(Y.tacos) + ' | - | - |\n';
  s += '| ROAS | ' + fmtR(L.roas) + ' | ' + fmtR(P.roas) + ' | ' + fmtR(Y.roas) + ' | - | - |\n';
  s += '| 粗利率 | ' + fmtP(L.grossMargin) + ' | ' + fmtP(P.grossMargin) + ' | ' + fmtP(Y.grossMargin) + ' | - | - |\n';
  s += '| 利益率 | ' + fmtP(L.profitMargin) + ' | ' + fmtP(P.profitMargin) + ' | ' + fmtP(Y.profitMargin) + ' | - | - |\n';

  s += '\n## カテゴリ別（前月）\n\n';
  s += '| カテゴリ | 売上 | 点数 | 広告費 | 利益 | 利益率 | TACOS |\n|---|---:|---:|---:|---:|---:|---:|\n';
  m.categories.forEach(c => {
    s += '| ' + c.category + ' | ' + fmtN(c.sales) + ' | ' + fmtN(c.units) + ' | ' +
         fmtN(c.adCost) + ' | ' + fmtN(c.profit) + ' | ' + fmtP(c.profitRate) + ' | ' + fmtP(c.tacos) + ' |\n';
  });

  s += '\n## 売上TOP' + m.topAsins.length + '商品（前月）\n\n';
  s += '| ASIN | 商品名 | カテゴリ | 売上 | 点数 | 広告費 | 利益 | 利益率 | ACOS | CVR |\n|---|---|---|---:|---:|---:|---:|---:|---:|---:|\n';
  m.topAsins.forEach(a => {
    s += '| ' + [
      a.asin, (a.name || '').substring(0, 20), a.category || '-',
      fmtN(a.sales), fmtN(a.units), fmtN(a.adCost),
      fmtN(a.profit), fmtP(a.profitRate), fmtP(a.acos), fmtP(a.cvr, 2),
    ].join(' | ') + ' |\n';
  });

  if (m.bottomAsins.length > 0) {
    s += '\n## 赤字商品（前月・損失の大きい順）\n\n';
    s += '| ASIN | 商品名 | 売上 | 広告費 | 利益 | ACOS |\n|---|---|---:|---:|---:|---:|\n';
    m.bottomAsins.forEach(a => {
      s += '| ' + [
        a.asin, (a.name || '').substring(0, 20),
        fmtN(a.sales), fmtN(a.adCost), fmtN(a.profit), fmtP(a.acos),
      ].join(' | ') + ' |\n';
    });
  }

  s += '\n---\n\n';
  s += '上記の数値を踏まえ、以下の構成で戦略レポートをMarkdownで作成してください。\n\n';
  s += '## 1. エグゼクティブサマリー\n';
  s += '（前月の業績要点を3行以内。利益・成長・リスクの観点で）\n\n';
  s += '## 2. 伸びている / 萎んでいるセグメント\n';
  s += '（カテゴリ別・ASIN別の傾向。原因推定含む）\n\n';
  s += '## 3. 逆算分析（来月の目標設定）\n';
  s += '（例: 利益 XX万円達成するには売上◯◯必要→CVR △% で セッション □□ 必要）\n\n';
  s += '## 4. 広告予算の最適配分提案\n';
  s += '（ROAS/ACOS の観点でASIN別に増額・減額・停止の判断。具体的なシフト額を提示）\n\n';
  s += '## 5. 注力 / 撤退の判断\n';
  s += '（赤字商品の扱い、カテゴリ別の投資判断。「いつまでに・どうする」を明示）\n\n';
  s += '## 6. 来月の重点タスク（優先度付き）\n';
  s += '（A: 今週中 / B: 今月中 / C: 四半期内 の3ランクで5〜8件）\n\n';
  s += '## 7. 経営上のリスク / 要注意点\n';
  s += '（在庫・競合・規約・アカウント健全性の観点で懸念事項）\n';
  return s;
}

function sendMonthlyMail(metrics, aiBody) {
  const to = getCredential('GMAIL_TO');
  const period = metrics.periods.lastMonth;
  const subject = '【Amazon月次戦略レポート】' + period.start.substring(0, 7) +
    '（売上 ' + Math.round(metrics.lastMonth.sales).toLocaleString() + '円 / 利益率 ' +
    (metrics.lastMonth.profitMargin * 100).toFixed(1) + '%）';

  const body =
    '■ 対象月: ' + period.start + ' 〜 ' + period.end + '\n' +
    '■ 売上: ' + Math.round(metrics.lastMonth.sales).toLocaleString() + ' 円\n' +
    '■ 利益: ' + Math.round(metrics.lastMonth.profit).toLocaleString() + ' 円' +
    ' (利益率 ' + (metrics.lastMonth.profitMargin * 100).toFixed(1) + '%)\n' +
    '■ 広告費: ' + Math.round(metrics.lastMonth.adCost).toLocaleString() + ' 円\n' +
    '■ 前月比売上: ' + (metrics.prevMonth.sales > 0
      ? (((metrics.lastMonth.sales - metrics.prevMonth.sales) / metrics.prevMonth.sales) * 100).toFixed(1) + '%'
      : '-') + '\n' +
    '■ 前年比売上: ' + (metrics.yoyMonth.sales > 0
      ? (((metrics.lastMonth.sales - metrics.yoyMonth.sales) / metrics.yoyMonth.sales) * 100).toFixed(1) + '%'
      : '-') + '\n\n' +
    '------ Claude Opus 戦略レポート ------\n\n' + aiBody +
    '\n\n--\nAmazon Dashboard 自動配信（' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') + '）';

  GmailApp.sendEmail(to, subject, body, { name: 'Amazon Dashboard' });
  Logger.log('📧 月次Gmail送信完了 → ' + to);
}

/**
 * テスト: プロンプトのみログ出力
 */
function testMonthlyPromptOnly() {
  const m = collectMonthlyMetrics();
  Logger.log(buildMonthlyPrompt(m));
}

/**
 * テスト: Claude を呼ぶが Gmail 送信はしない
 */
function testMonthlyAiReportNoMail() {
  const m = collectMonthlyMetrics();
  if (m.lastMonth.sales === 0) { Logger.log('前月売上0のため中止'); return; }
  const prompt = buildMonthlyPrompt(m);
  const body = callClaude({
    model: CLAUDE_MODELS.MONTHLY,
    system: MONTHLY_SYSTEM_PROMPT,
    prompt: prompt,
    maxTokens: 8192,
    temperature: 0.5,
  });
  Logger.log('===== Claude Opus 応答 =====\n' + body);
}
