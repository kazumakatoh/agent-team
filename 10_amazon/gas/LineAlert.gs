/**
 * Amazon Dashboard - LINE 緊急アラート
 *
 * 在庫切れ間近・広告費異常・大型返金・アカウント警告など、即時対応が必要な事象を
 * LINE Messaging API（Push Message）で社長個人宛に通知する。
 *
 * ## トリガー条件（kpi_and_operations.md 準拠）
 *   1. 在庫切れ間近: FBA在庫 < 直近7日平均日販 × 3日分
 *   2. 広告費異常: 直近1日の広告費が 過去7日平均 × 2 を超過
 *   3. 大型返金: Settlement で1日あたり Refund/Adjustment > 10,000円
 *   4. アカウント健全性警告: Account Health API（Phase 4b 実装後）
 *
 * ## 重複抑止
 *   同一アラートを毎日送らないよう、ScriptProperties に「日付＋アラートキー」を保存し
 *   24時間以内の再送をスキップする。
 *
 * ## トリガー: 毎日 AM9:00（runDailyAlerts）
 */

const LINE_PUSH_URL = 'https://api.line.me/v2/bot/message/push';
const ALERT_DEDUP_PROP = 'LINE_ALERT_SENT';
const ALERT_DEDUP_HOURS = 22;       // 22時間 → 翌朝の同条件は再送

// 閾値
const STOCK_DAYS_THRESHOLD = 3;     // 残り日数（日販×N日分を下回ったら警告）
const STOCK_LOOKBACK_DAYS = 7;      // 日販平均の参照期間
const AD_SPIKE_RATIO = 2.0;         // 広告費スパイク倍率
const REFUND_THRESHOLD = 10000;     // 大型返金の閾値（円/日）

/**
 * メイン: 全アラートをチェック → 該当があれば LINE 送信
 */
function runDailyAlerts() {
  Logger.log('===== 緊急アラートチェック 開始 =====');
  const alerts = [];

  try { alerts.push(...checkAdSpikes()); } catch (e) { Logger.log('広告費チェックエラー: ' + e.message); }
  try { alerts.push(...checkLargeRefunds()); } catch (e) { Logger.log('返金チェックエラー: ' + e.message); }
  // 在庫チェックは Inventory API 取得後に有効化（fetchInventory が D1 在庫列を埋める想定）
  // try { alerts.push(...checkLowStock()); } catch (e) { Logger.log('在庫チェックエラー: ' + e.message); }

  if (alerts.length === 0) {
    Logger.log('✅ アラート該当なし');
    return;
  }

  // 重複抑止: 過去22時間以内に同キーで送信済みなら除外
  const sent = getSentAlertMap();
  const now = Date.now();
  const fresh = alerts.filter(a => {
    const last = sent[a.key];
    return !last || (now - last) > ALERT_DEDUP_HOURS * 3600 * 1000;
  });

  if (fresh.length === 0) {
    Logger.log('全アラート重複抑止により送信スキップ');
    return;
  }

  pushLineAlert(formatAlertMessage(fresh));
  fresh.forEach(a => { sent[a.key] = now; });
  saveSentAlertMap(sent);
  Logger.log('🔔 LINE送信: ' + fresh.length + ' 件');
}

// ===== 検知ロジック =====

/**
 * 広告費スパイク: 昨日の広告費が直近7日平均×2を超えたASIN一覧
 */
function checkAdSpikes() {
  const data = getDailyDataAll();
  if (data.length === 0) return [];

  const today = new Date();
  const fmt = d => Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  const yesterday = fmt(addDays(today, -1));
  const window7Start = fmt(addDays(today, -8));
  const window7End = fmt(addDays(today, -2));

  // ASINごとに昨日と過去7日平均を比較
  const yMap = {}, baseMap = {}, baseDays = {};
  for (const d of data) {
    if (!d.asin) continue;
    if (d.date === yesterday) {
      yMap[d.asin] = (yMap[d.asin] || 0) + d.adCost;
    }
    if (d.date >= window7Start && d.date <= window7End) {
      baseMap[d.asin] = (baseMap[d.asin] || 0) + d.adCost;
      if (!baseDays[d.asin]) baseDays[d.asin] = new Set();
      baseDays[d.asin].add(d.date);
    }
  }

  const alerts = [];
  for (const [asin, ySpend] of Object.entries(yMap)) {
    if (ySpend < 1000) continue; // 100円未満はノイズ
    const days = baseDays[asin] ? baseDays[asin].size : 0;
    if (days < 3) continue;       // 過去データが少なすぎる
    const avg = (baseMap[asin] || 0) / days;
    if (avg <= 0) continue;
    if (ySpend / avg >= AD_SPIKE_RATIO) {
      alerts.push({
        key: 'AD_SPIKE_' + yesterday + '_' + asin,
        type: '🔺広告費急増',
        line: asin + ': 昨日 ' + Math.round(ySpend).toLocaleString() + '円 ' +
              '（過去7日平均 ' + Math.round(avg).toLocaleString() + '円 × ' + (ySpend / avg).toFixed(1) + ')',
      });
    }
  }
  return alerts;
}

/**
 * 大型返金: 直近の経費明細から1日あたり返金額 > 閾値 の日を検出
 * D2 経費明細（明細種別が "Refund" を含むもの）を対象に集計。
 */
function checkLargeRefunds() {
  const sheet = getOrCreateSheet(SHEET_NAMES.D2_SETTLEMENT);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  // D2 列: 0:決済期間開始 1:決済期間終了 2:日付 3:ASIN 4:トランザクション種別 5:明細種別 6:金額
  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const today = new Date();
  const since = Utilities.formatDate(addDays(today, -2), 'Asia/Tokyo', 'yyyy-MM-dd');

  const dailyRefund = {};
  for (const row of data) {
    const date = row[2] instanceof Date
      ? Utilities.formatDate(row[2], 'Asia/Tokyo', 'yyyy-MM-dd')
      : String(row[2] || '').substring(0, 10);
    if (!date || date < since) continue;
    const itemType = String(row[5] || '');
    if (!/Refund|Adjustment|return/i.test(itemType)) continue;
    const amount = parseFloat(row[6]) || 0;
    if (amount >= 0) continue;       // 返金はマイナス値で記録される
    dailyRefund[date] = (dailyRefund[date] || 0) + Math.abs(amount);
  }

  const alerts = [];
  for (const [date, amount] of Object.entries(dailyRefund)) {
    if (amount >= REFUND_THRESHOLD) {
      alerts.push({
        key: 'REFUND_' + date,
        type: '💸大型返金/調整',
        line: date + ': ' + Math.round(amount).toLocaleString() + '円（閾値 ' + REFUND_THRESHOLD.toLocaleString() + '円）',
      });
    }
  }
  return alerts;
}

/**
 * （在庫チェック・将来）FBA在庫 < 直近7日日販 × 3 を検出
 * fetchInventory() が D1 等に在庫を書き込むようになったら呼び出す。
 */
function checkLowStock() {
  Logger.log('checkLowStock: 在庫データソース未実装のためスキップ');
  return [];
}

// ===== LINE 送信 =====

function pushLineAlert(text) {
  const token = getCredential('LINE_CHANNEL_TOKEN');
  const userId = getCredential('LINE_USER_ID');

  const payload = {
    to: userId,
    messages: [{ type: 'text', text: text.substring(0, 4900) }],   // LINE上限5000
  };
  const res = UrlFetchApp.fetch(LINE_PUSH_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code !== 200) {
    throw new Error('LINE送信エラー HTTP ' + code + ': ' + res.getContentText());
  }
}

function formatAlertMessage(alerts) {
  const head = '【Amazon緊急アラート】' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M/d HH:mm') + '\n';
  const grouped = {};
  for (const a of alerts) {
    if (!grouped[a.type]) grouped[a.type] = [];
    grouped[a.type].push(a.line);
  }
  let body = '';
  for (const [type, lines] of Object.entries(grouped)) {
    body += '\n■ ' + type + '\n' + lines.map(l => '  ・' + l).join('\n') + '\n';
  }
  return head + body;
}

// ===== ユーティリティ =====

function addDays(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; }

function getSentAlertMap() {
  const raw = PropertiesService.getScriptProperties().getProperty(ALERT_DEDUP_PROP);
  if (!raw) return {};
  try { return JSON.parse(raw); } catch (e) { return {}; }
}

function saveSentAlertMap(map) {
  // 古いキー（48h以上前）は掃除
  const cutoff = Date.now() - 48 * 3600 * 1000;
  const cleaned = {};
  for (const [k, v] of Object.entries(map)) {
    if (v >= cutoff) cleaned[k] = v;
  }
  PropertiesService.getScriptProperties().setProperty(ALERT_DEDUP_PROP, JSON.stringify(cleaned));
}

// ===== テスト =====

function testLineAlert() {
  pushLineAlert('【テスト】Amazon Dashboard からの通知テストです。\n' +
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'));
  Logger.log('✅ テスト送信完了');
}

function testCheckAdSpikes() {
  const a = checkAdSpikes();
  Logger.log('検出 ' + a.length + ' 件');
  a.forEach(x => Logger.log('  ' + x.type + ' / ' + x.line));
}

function testCheckRefunds() {
  const a = checkLargeRefunds();
  Logger.log('検出 ' + a.length + ' 件');
  a.forEach(x => Logger.log('  ' + x.type + ' / ' + x.line));
}
