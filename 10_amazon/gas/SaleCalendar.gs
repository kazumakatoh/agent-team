/**
 * Amazon Dashboard - セール対策（Phase 4b ④）
 *
 * 年間セールカレンダー（プライムデー / ブラックフライデー / 年末等）を
 * シート M4 に保持し、セール6週前 / 4週前 / 2週前 のタイミングで
 * Claude API に「準備状況の点検＋施策提案」をしてもらう。
 *
 * ## M4 シート構造
 *   イベント名 | 開始日 | 終了日 | 準備開始日 | ステータス | メモ
 *
 * 初期データはセットアップ時に自動投入。社長が必要に応じて手で更新。
 *
 * ## トリガー: 毎週月曜 7:30（checkUpcomingSales）
 *   → 該当する準備期間に入ったら Claude が施策レポート → Gmail
 */

const M4_SALE_CALENDAR = 'セールカレンダー';
const M4_SALE_HEADERS = ['イベント名', '開始日', '終了日', '準備開始日', 'ステータス', 'メモ'];

const SALE_LEAD_WEEKS = [6, 4, 2];   // この週数以内に入ったら通知

/**
 * M4 セットアップ: 当年と来年の主要セールを自動投入
 */
function setupSaleCalendar() {
  const sheet = getOrCreateSheet(M4_SALE_CALENDAR);
  const existing = sheet.getRange(1, 1, 1, M4_SALE_HEADERS.length).getValues()[0];
  if (existing[0] !== M4_SALE_HEADERS[0]) {
    sheet.getRange(1, 1, 1, M4_SALE_HEADERS.length).setValues([M4_SALE_HEADERS])
      .setFontWeight('bold').setBackground('#e8f0fe');
    sheet.setFrozenRows(1);
  }

  const lastRow = sheet.getLastRow();
  const existingNames = new Set();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(r => existingNames.add(String(r[0])));
  }

  const year = new Date().getFullYear();
  // 概算日付（実日付は毎年公式発表で更新する）
  const defaults = [];
  [year, year + 1].forEach(y => {
    defaults.push(['プライムデー ' + y, y + '-07-15', y + '-07-16', y + '-06-01', '未開始', '6週前から準備開始']);
    defaults.push(['ブラックフライデー ' + y, y + '-11-22', y + '-11-29', y + '-10-11', '未開始', '6週前から準備開始']);
    defaults.push(['年末セール ' + y, y + '-12-20', y + '-12-31', y + '-11-20', '未開始', '1ヶ月前から準備']);
    defaults.push(['新生活セール ' + y, y + '-03-25', y + '-04-04', y + '-02-15', '未開始', '6週前から準備開始']);
  });

  const newRows = defaults.filter(r => !existingNames.has(r[0]));
  if (newRows.length > 0) {
    sheet.getRange(lastRow + 1, 1, newRows.length, M4_SALE_HEADERS.length).setValues(newRows);
    Logger.log('✅ セールカレンダー: ' + newRows.length + ' 件追加');
  }
}

/**
 * 直近のセールに対して「準備期間の点検」レポートをAIで生成
 * SALE_LEAD_WEEKS のいずれかの境界週に該当したら送信。
 */
function checkUpcomingSales() {
  const sheet = getOrCreateSheet(M4_SALE_CALENDAR);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) { Logger.log('セールカレンダー未設定'); return; }

  const today = new Date();
  const data = sheet.getRange(2, 1, lastRow - 1, M4_SALE_HEADERS.length).getValues();

  for (const row of data) {
    const name = String(row[0]);
    const start = parseDateLoose(row[1]);
    if (!start) continue;
    const daysUntil = Math.floor((start - today) / (24 * 3600 * 1000));
    if (daysUntil <= 0 || daysUntil > 50) continue; // セール中・終了 or 50日以上先 はスキップ

    const weeksUntil = Math.ceil(daysUntil / 7);
    if (!SALE_LEAD_WEEKS.includes(weeksUntil)) continue;

    Logger.log('📅 ' + name + ' まで ' + daysUntil + '日 → AIレポート生成');
    sendSalePrepReport(name, daysUntil, start);
  }
}

function sendSalePrepReport(name, daysUntil, startDate) {
  const m = collectWeeklyMetrics();
  const fmtN = n => Math.round(n).toLocaleString();

  const dataSection =
    '## 直近1週間の業績\n' +
    '- 売上: ' + fmtN(m.thisWeek.sales) + '円\n' +
    '- 利益: ' + fmtN(m.thisWeek.profit) + '円（利益率 ' + (m.thisWeek.profitMargin * 100).toFixed(1) + '%）\n' +
    '- 広告費: ' + fmtN(m.thisWeek.adCost) + '円（TACOS ' + (m.thisWeek.tacos * 100).toFixed(1) + '%）\n\n' +
    '## 売上TOP商品\n' +
    m.topAsins.slice(0, 5).map(a => '- ' + a.asin + ' ' + (a.name || '') +
      '：売上 ' + fmtN(a.sales) + '円 / 利益 ' + fmtN(a.profit) + '円').join('\n');

  const prompt =
    '【セール準備レポート依頼】\n' +
    'イベント: ' + name + '\n' +
    '開始まで: ' + daysUntil + '日（' + Utilities.formatDate(startDate, 'Asia/Tokyo', 'M月d日') + '）\n\n' +
    dataSection + '\n\n' +
    '上記を踏まえ、以下の観点でMarkdownレポートを作成してください。\n\n' +
    '## 1. このセールの位置づけ\n（年間で何位の重要度か / 過去実績の傾向）\n\n' +
    '## 2. 在庫準備の必要量\n（TOP商品ごとに推奨発注数の目安と理由）\n\n' +
    '## 3. 価格戦略\n（クーポン/値引き率の推奨。利益率を維持できる範囲で）\n\n' +
    '## 4. 広告戦略\n（セール期間に集中投下すべきASIN・キャンペーンタイプ）\n\n' +
    '## 5. やるべきタスク（優先度付き）\n（チェックリスト形式・期日入り）';

  const aiBody = callClaude({
    model: CLAUDE_MODELS.WEEKLY,
    system: 'あなたは Amazon 物販のセール対策専門コンサルです。実行可能な具体策を提示してください。',
    prompt: prompt,
    maxTokens: 4096,
    temperature: 0.5,
  });

  const to = getCredential('GMAIL_TO');
  const subject = '【セール準備】' + name + '（' + daysUntil + '日前・準備チェック）';
  const body =
    '■ イベント: ' + name + '\n' +
    '■ 開始日: ' + Utilities.formatDate(startDate, 'Asia/Tokyo', 'yyyy-MM-dd') + '\n' +
    '■ 残り日数: ' + daysUntil + '日\n\n' +
    '------ Claude AI セール対策レポート ------\n\n' + aiBody +
    '\n\n--\nAmazon Dashboard 自動配信';
  GmailApp.sendEmail(to, subject, body, { name: 'Amazon Dashboard' });
  Logger.log('📧 セール準備メール送信: ' + name);
}

function parseDateLoose(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  const s = String(v).trim();
  const m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (!m) return null;
  return new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
}
