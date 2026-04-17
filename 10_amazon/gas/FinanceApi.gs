/**
 * Amazon Dashboard - Finance API（日次フィー取得）
 *
 * Settlement Report は14日サイクルで確定するため、リアルタイムに近いフィー情報は
 * Finance API から取得して D2F「日次フィー（Finance）」シートに保存する。
 *
 * ## エンドポイント
 *
 * GET /finances/v0/financialEvents
 *   ?PostedAfter=YYYY-MM-DDTHH:mm:ssZ
 *   &PostedBefore=YYYY-MM-DDTHH:mm:ssZ
 *   &MaxResultsPerPage=100
 *   &NextToken=...（ページング）
 *
 * ## 取得対象イベント
 *
 * ServiceFeeEventList（FBA Inbound Transportation Fee 等の手数料）を中心に取得。
 * 他にも ShipmentEventList / RefundEventList 等あるが、本実装では絞る。
 *
 * ## D2F シート構造（5列）
 *
 * 日付 | イベント種別 | FeeReason | 金額 | ステータス
 *
 * ステータス: '暫定' (Finance API取得時) / '確定' (Settlement Report で確認後)
 *
 * ## トリガー: 毎日 AM7:30 (fetchDailyFinanceEvents)
 */

const D2F_HEADERS = ['日付', 'イベント種別', 'FeeReason', '金額', 'ステータス'];

/**
 * メイン: 前日分の Finance Events を取得して D2F に保存
 */
function fetchDailyFinanceEvents() {
  const t0 = Date.now();
  Logger.log('===== Finance API 日次フィー取得 開始 =====');

  // 日付範囲: 過去2日 〜 5分前（PostedBefore は「現在から2分以上過去」が必須）
  const now = new Date();
  const postedBefore = new Date(now.getTime() - 5 * 60 * 1000).toISOString();
  const postedAfter = new Date(now.getTime() - 2 * 24 * 60 * 60 * 1000).toISOString();

  Logger.log('期間: ' + postedAfter + ' 〜 ' + postedBefore);

  setupD2fHeaders();

  // 既存データのキー（日付+種別+金額）を集めて重複防止
  const existingKeys = getExistingD2fKeys();
  Logger.log('既存D2F件数: ' + existingKeys.size);

  // ページング込みで全件取得
  const events = fetchAllFinancialEvents(postedAfter, postedBefore);
  Logger.log('取得イベント総数: ' + events.length);

  // 重複除外しつつ書き込み
  const newRows = [];
  for (const ev of events) {
    const key = ev.date + '|' + ev.eventType + '|' + ev.feeReason + '|' + ev.amount;
    if (existingKeys.has(key)) continue;
    newRows.push([ev.date, ev.eventType, ev.feeReason, ev.amount, '暫定']);
    existingKeys.add(key);
  }

  if (newRows.length > 0) {
    appendRows(SHEET_NAMES.D2F_FINANCE_EVENTS, newRows);
    Logger.log('✅ D2F: ' + newRows.length + ' 件 追記');
  } else {
    Logger.log('新規イベントなし');
  }

  Logger.log('===== Finance API 完了（' + (Date.now() - t0) + 'ms）=====');
}

/**
 * Finance API から指定期間のイベントをページング込みで全件取得
 *
 * @param {string} postedAfter ISO 8601
 * @param {string} postedBefore ISO 8601
 * @returns {Array<Object>} { date, eventType, feeReason, amount }[]
 */
function fetchAllFinancialEvents(postedAfter, postedBefore) {
  const events = [];
  let nextToken = null;
  let page = 0;
  const maxPages = 50; // 安全弁

  while (page < maxPages) {
    page++;
    const params = {
      PostedAfter: postedAfter,
      PostedBefore: postedBefore,
      MaxResultsPerPage: 100,
    };
    if (nextToken) params.NextToken = nextToken;

    const result = callSpApi('GET', '/finances/v0/financialEvents', params);
    const payload = result.payload || {};
    const fe = payload.FinancialEvents || {};

    // ServiceFeeEventList: FBA手数料・倉庫手数料等
    const serviceFees = fe.ServiceFeeEventList || [];
    for (const sf of serviceFees) {
      const feeList = sf.FeeList || [];
      for (const fee of feeList) {
        events.push({
          date: extractDate(sf.PostedDate || ''),
          eventType: 'ServiceFee',
          feeReason: sf.FeeReason || sf.FeeDescription || '',
          amount: fee.FeeAmount ? parseFloat(fee.FeeAmount.CurrencyAmount) : 0,
        });
      }
    }

    nextToken = payload.NextToken || null;
    if (!nextToken) break;

    Utilities.sleep(500); // レート制限対策
  }

  return events;
}

/**
 * D2F シートのヘッダーを設定
 */
function setupD2fHeaders() {
  const sheet = getOrCreateSheet(SHEET_NAMES.D2F_FINANCE_EVENTS);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, D2F_HEADERS.length).setValues([D2F_HEADERS])
      .setFontWeight('bold').setBackground('#e8f0fe');
    sheet.setFrozenRows(1);
  }
}

/**
 * 既存 D2F のキー集合を取得（重複防止用）
 */
function getExistingD2fKeys() {
  const sheet = getOrCreateSheet(SHEET_NAMES.D2F_FINANCE_EVENTS);
  const lastRow = sheet.getLastRow();
  const keys = new Set();
  if (lastRow <= 1) return keys;

  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  for (const row of data) {
    const date = row[0] instanceof Date
      ? Utilities.formatDate(row[0], 'Asia/Tokyo', 'yyyy-MM-dd')
      : String(row[0] || '').substring(0, 10);
    const key = date + '|' + row[1] + '|' + row[2] + '|' + row[3];
    keys.add(key);
  }
  return keys;
}

/**
 * ISO 8601 日時から YYYY-MM-DD 部分を抽出
 */
function extractDate(isoStr) {
  if (!isoStr) return '';
  return String(isoStr).substring(0, 10);
}

/**
 * Settlement Report 取込時に呼び出す: D2F の暫定データを「確定」にマーク
 *
 * Settlement Report で同じ取引が確認された日付範囲について、
 * D2F の status を「暫定」→「確定」に変更（重複集計防止）。
 *
 * @param {string} startDate 'YYYY-MM-DD'
 * @param {string} endDate   'YYYY-MM-DD'
 */
function markFinanceEventsAsConfirmed(startDate, endDate) {
  const sheet = getOrCreateSheet(SHEET_NAMES.D2F_FINANCE_EVENTS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  let updated = 0;
  for (let i = 0; i < data.length; i++) {
    const rawDate = data[i][0];
    const date = rawDate instanceof Date
      ? Utilities.formatDate(rawDate, 'Asia/Tokyo', 'yyyy-MM-dd')
      : String(rawDate || '').substring(0, 10);
    if (date < startDate || date > endDate) continue;

    if (data[i][4] !== '確定') {
      sheet.getRange(i + 2, 5).setValue('確定');
      updated++;
    }
  }

  if (updated > 0) {
    Logger.log('✅ D2F 確定マーク: ' + updated + ' 件 (' + startDate + ' 〜 ' + endDate + ')');
  }
}

/**
 * テスト: 直近1日分の Finance Events を取得して内訳を表示
 */
function testFinanceApi() {
  const now = new Date();
  // PostedBefore は「現在から2分以上過去」が必須。5分前にしてバッファ確保
  const postedBefore = new Date(now.getTime() - 5 * 60 * 1000).toISOString();
  const postedAfter = new Date(now.getTime() - 24 * 60 * 60 * 1000).toISOString();
  const events = fetchAllFinancialEvents(postedAfter, postedBefore);

  Logger.log('取得件数: ' + events.length);
  const byReason = {};
  events.forEach(e => {
    const key = e.feeReason || '(empty)';
    byReason[key] = (byReason[key] || 0) + 1;
  });
  Logger.log('FeeReason 別件数:');
  Object.entries(byReason).forEach(([k, v]) => Logger.log('  ' + k + ': ' + v));
}
