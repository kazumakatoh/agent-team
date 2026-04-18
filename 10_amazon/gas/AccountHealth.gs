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
 */
function runAccountHealthCheck() {
  Logger.log('===== アカウント健全性チェック =====');
  setupHealthSheet();

  const today = new Date();
  const fmt = d => Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  const dailyData = getDailyDataAll();

  const recentReturn7 = aggregateReturnRate(dailyData, fmt(addDays(today, -7)), fmt(addDays(today, -1)));
  const recentReturn30 = aggregateReturnRate(dailyData, fmt(addDays(today, -30)), fmt(addDays(today, -1)));

  const issues = [];
  if (recentReturn7.rate >= RETURN_RATE_WARNING) {
    issues.push('返品率が ' + (recentReturn7.rate * 100).toFixed(2) + '% / 警戒水準');
  }
  if (recentReturn30.rate > 0 && recentReturn7.rate / recentReturn30.rate >= RETURN_RATE_SPIKE) {
    issues.push('返品率が30日比 ' + (recentReturn7.rate / recentReturn30.rate).toFixed(2) + '倍に急増');
  }

  // スコア（簡易）: 返品率と問題件数からざっくり算出
  let score = 100;
  if (recentReturn7.rate >= 0.03) score -= 10;
  if (recentReturn7.rate >= 0.05) score -= 20;
  if (recentReturn7.rate >= 0.08) score -= 30;
  score = Math.max(0, score - issues.length * 5);

  appendRows(D5_HEALTH, [[
    fmt(today),
    score,
    (recentReturn7.rate * 100).toFixed(2) + '%',
    (recentReturn30.rate * 100).toFixed(2) + '%',
    issues.length === 0 ? '異常なし' : issues.join(' / '),
  ]]);

  Logger.log('スコア: ' + score + ' / 7日返品率: ' + (recentReturn7.rate * 100).toFixed(2) + '%');

  // LINE 通知（重要な問題のみ）
  if (issues.length > 0) {
    notifyHealthIssues(issues, score, recentReturn7);
  }
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
