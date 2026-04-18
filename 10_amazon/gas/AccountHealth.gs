/**
 * Amazon Dashboard - アカウント健全性チェック（Phase 4b ⑤）
 *
 * SP-API には完全な Account Health 取得 API は限定的なため、以下を組み合わせて
 * 健全性スコアを推定する:
 *
 *   1. /sellers/v1/account              … アカウントステータス
 *   2. Performance Notifications（SQS） … 公式違反通知（要設定）
 *   3. 返品率の急増                       … D1 / D2 から計算
 *   4. 在庫切れ商品比率                   … fetchInventory（実装済）から計算
 *
 * このモジュールでは "(3) 返品率の急増" のみ自動チェックを実装し、
 * 残りはログ + 手動Seller Central確認の運用とする（権限取得後に拡張）。
 *
 * ## D5 シート: 健全性ログ
 *   日付 | スコア | 主な問題 | アクション
 *
 * ## トリガー: 毎日 AM9:15（runAccountHealthCheck）
 */

const D5_HEALTH = 'アカウント健全性';
const D5_HEALTH_HEADERS = ['日付', '総合', '返品率(直近7日)', '返品率(直近30日)', '注意点'];

const RETURN_RATE_WARNING = 0.05;   // 直近7日返品率が 5% 超で警告
const RETURN_RATE_SPIKE = 1.5;      // 7日 / 30日 > 1.5 で「急増」判定

/**
 * メイン: 健全性ログを更新
 *
 * 以下を組み合わせた総合健全性チェック:
 *   1. SP-API /sellers/v1/account でアカウント状態
 *   2. 返品率（7日 / 30日、急増判定）
 *   3. 在庫切れ商品の比率
 */
function runAccountHealthCheck() {
  Logger.log('===== アカウント健全性チェック =====');
  setupHealthSheet();

  const today = new Date();
  const fmt = d => Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  const dailyData = getDailyDataAll();

  const issues = [];
  let score = 100;

  // --- 1. SP-API アカウントステータス ---
  let accountStatus = '-';
  try {
    const acc = callSpApi('GET', '/sellers/v1/account');
    const participations = acc.payload || [];
    const jp = participations.find(p => (p.marketplace && p.marketplace.id) === MARKETPLACE_ID_JP);
    if (jp) {
      const status = jp.participation && jp.participation.isParticipating;
      accountStatus = status ? '✅ 参加中' : '⚠️ 非参加';
      if (!status) { issues.push('JP マーケットプレイスが非参加状態'); score -= 30; }
    } else {
      accountStatus = '⚠️ JP参加情報なし';
      issues.push('JP 参加情報が取得できない');
      score -= 20;
    }
  } catch (e) {
    accountStatus = '⚠️ 取得失敗';
    Logger.log('sellers/v1/account 取得失敗: ' + e.message);
  }

  // --- 2. 返品率 ---
  const recentReturn7 = aggregateReturnRate(dailyData, fmt(addDays(today, -7)), fmt(addDays(today, -1)));
  const recentReturn30 = aggregateReturnRate(dailyData, fmt(addDays(today, -30)), fmt(addDays(today, -1)));

  if (recentReturn7.rate >= RETURN_RATE_WARNING) {
    issues.push('返品率が ' + (recentReturn7.rate * 100).toFixed(2) + '% / 警戒水準');
  }
  if (recentReturn30.rate > 0 && recentReturn7.rate / recentReturn30.rate >= RETURN_RATE_SPIKE) {
    issues.push('返品率が30日比 ' + (recentReturn7.rate / recentReturn30.rate).toFixed(2) + '倍に急増');
  }
  if (recentReturn7.rate >= 0.03) score -= 10;
  if (recentReturn7.rate >= 0.05) score -= 20;
  if (recentReturn7.rate >= 0.08) score -= 30;

  // --- 3. 在庫切れ比率（D6 在庫シートが存在する場合のみ） ---
  const stockStat = summarizeStockStatus();
  if (stockStat.total > 0) {
    const outRatio = stockStat.outOfStock / stockStat.total;
    const criticalRatio = stockStat.critical / stockStat.total;
    if (outRatio >= 0.2) { issues.push('在庫切れ商品が ' + stockStat.outOfStock + ' 件（' + (outRatio * 100).toFixed(0) + '%）'); score -= 15; }
    else if (outRatio >= 0.1) { issues.push('在庫切れ ' + stockStat.outOfStock + ' 件（' + (outRatio * 100).toFixed(0) + '%）'); score -= 10; }
    if (criticalRatio >= 0.1) { issues.push('在庫3日分未満の緊急商品 ' + stockStat.critical + ' 件'); score -= 10; }
  }

  score = Math.max(0, score - issues.length * 2);

  appendRows(D5_HEALTH, [[
    fmt(today),
    score,
    (recentReturn7.rate * 100).toFixed(2) + '%',
    (recentReturn30.rate * 100).toFixed(2) + '%',
    accountStatus +
      (stockStat.total > 0 ? ' / 在庫切れ ' + stockStat.outOfStock + '/' + stockStat.total : '') +
      (issues.length === 0 ? ' / 異常なし' : ' / ' + issues.join(' / ')),
  ]]);

  Logger.log('スコア: ' + score + ' / 7日返品率: ' + (recentReturn7.rate * 100).toFixed(2) +
             '% / アカウント: ' + accountStatus);

  if (issues.length > 0) {
    notifyHealthIssues(issues, score, recentReturn7);
  }
}

/**
 * D6 在庫シートから在庫ステータスの内訳を集計
 * シート未作成の場合は全ゼロを返す
 */
function summarizeStockStatus() {
  const ss = getMainSpreadsheet();
  const sheet = ss.getSheetByName(D6_INVENTORY);
  if (!sheet || sheet.getLastRow() <= 1) return { total: 0, outOfStock: 0, critical: 0, warning: 0 };

  const data = sheet.getRange(2, 8, sheet.getLastRow() - 1, 1).getValues(); // 8列目: ステータス
  let total = 0, outOfStock = 0, critical = 0, warning = 0;
  for (const row of data) {
    const status = row[0];
    if (!status) continue;
    total++;
    if (status === '在庫切れ') outOfStock++;
    else if (status === '緊急') critical++;
    else if (status === '警告') warning++;
  }
  return { total, outOfStock, critical, warning };
}

function aggregateReturnRate(dailyData, start, end) {
  let units = 0, returnUnits = 0;
  for (const d of dailyData) {
    if (d.date < start || d.date > end) continue;
    units += d.units;
    // D1 列14: 返品数（getDailyDataAll の列マッピングが拡張されたら直接拾う）
    // 現状ロジックは units に閉じる。getDailyDataAll の改修で returnUnits 拡張可能。
    returnUnits += d.returnUnits || 0;
  }
  return {
    units, returnUnits,
    rate: units > 0 ? returnUnits / units : 0,
  };
}

function notifyHealthIssues(issues, score, returnInfo) {
  const alerts = issues.map((line, i) => ({
    key: 'HEALTH_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd') + '_' + i,
    type: '🏥アカウント健全性',
    line: line,
  }));
  // LineAlert.gs の重複抑止＋送信
  const sent = getSentAlertMap();
  const now = Date.now();
  const fresh = alerts.filter(a => {
    const last = sent[a.key];
    return !last || (now - last) > ALERT_DEDUP_HOURS * 3600 * 1000;
  });
  if (fresh.length === 0) return;
  const head = '【アカウント健全性】スコア ' + score + '/100\n';
  pushLineAlert(head + formatAlertMessage(fresh));
  fresh.forEach(a => { sent[a.key] = now; });
  saveSentAlertMap(sent);
}

function setupHealthSheet() {
  const sheet = getOrCreateSheetCompact(D5_HEALTH, D5_HEALTH_HEADERS.length, 400);
  const existing = sheet.getRange(1, 1, 1, D5_HEALTH_HEADERS.length).getValues()[0];
  if (existing[0] !== D5_HEALTH_HEADERS[0]) {
    sheet.getRange(1, 1, 1, D5_HEALTH_HEADERS.length).setValues([D5_HEALTH_HEADERS])
      .setFontWeight('bold').setBackground('#e8f0fe');
    sheet.setFrozenRows(1);
  }
}
