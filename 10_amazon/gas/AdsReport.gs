/**
 * Amazon Dashboard - Amazon Ads API v3 レポート取得（Phase 3）
 *
 * Sponsored Products を対象に以下3種のレポートを取得して D3 へ蓄積し、
 * 併せて D1 日次データの広告4指標（広告費/広告売上/IMP/CT）を更新する：
 *
 *   - spAdvertisedProduct  : ASIN×日次の広告実績   → D3_CAMPAIGN + D1 更新
 *   - spSearchTerm         : 検索用語別            → D3_SEARCHTERM
 *   - spTargeting          : キーワード/ターゲット  → D3_TARGET
 *
 * 帰属期間（attribution window）は 7 日で統一（purchases7d / sales7d）。
 *
 * Amazon Ads v3 の非同期レポートフロー：
 *   1) POST /reporting/reports  → reportId 取得
 *   2) GET  /reporting/reports/{reportId}  → status が COMPLETED になるまでポーリング
 *   3) レスポンスの url（S3 pre-signed）を UrlFetchApp.fetch → gzip ungzip → JSON.parse
 */

// ===== v3 Content-Type（vnd 形式が必須） =====
const ADS_V3_CT_CREATE_REPORT = 'application/vnd.createasyncreportrequest.v3+json';

// ===== ポーリング設定 =====
const ADS_REPORT_POLL_INTERVAL_MS = 15000;  // 15秒間隔
const ADS_REPORT_POLL_MAX_ATTEMPTS = 20;    // 最大 5 分（15s × 20）

// ===== レポート別カラム定義 =====
// 各 reportTypeId に対して groupBy と columns を定義
const ADS_REPORT_CONFIGS = {
  spAdvertisedProduct: {
    adProduct: 'SPONSORED_PRODUCTS',
    reportTypeId: 'spAdvertisedProduct',
    groupBy: ['advertiser'],
    timeUnit: 'DAILY',
    columns: [
      'date', 'campaignId', 'campaignName', 'adGroupId', 'adGroupName',
      'advertisedAsin', 'advertisedSku',
      'impressions', 'clicks', 'cost', 'clickThroughRate', 'costPerClick',
      'purchases7d', 'sales7d', 'unitsSoldClicks7d',
      'acosClicks7d', 'roasClicks7d',
    ],
  },
  spSearchTerm: {
    adProduct: 'SPONSORED_PRODUCTS',
    reportTypeId: 'spSearchTerm',
    groupBy: ['searchTerm'],
    timeUnit: 'DAILY',
    columns: [
      'date', 'campaignId', 'campaignName', 'adGroupId', 'adGroupName',
      'keywordId', 'keyword', 'matchType', 'searchTerm',
      'impressions', 'clicks', 'cost', 'clickThroughRate', 'costPerClick',
      'purchases7d', 'sales7d',
      'acosClicks7d', 'roasClicks7d',
    ],
  },
  spTargeting: {
    adProduct: 'SPONSORED_PRODUCTS',
    reportTypeId: 'spTargeting',
    groupBy: ['targeting'],
    timeUnit: 'DAILY',
    columns: [
      'date', 'campaignId', 'campaignName', 'adGroupId', 'adGroupName',
      'keywordId', 'keyword', 'matchType', 'targeting',
      'impressions', 'clicks', 'cost', 'clickThroughRate', 'costPerClick',
      'purchases7d', 'sales7d',
      'acosClicks7d', 'roasClicks7d',
    ],
  },
};

// ===== 公開関数：1日分を3種まとめて取得し D3 / D1 に反映 =====

/**
 * 昨日分の広告レポートを取得 → D3 / D1 反映
 * 毎日トリガーから呼ばれる想定
 */
function dailyFetchAdsReports() {
  const yesterday = getYesterday();
  Logger.log('===== Ads Reports 日次取得 (' + yesterday + ') =====');
  fetchAdsReportsForRange(yesterday, yesterday);
}

/**
 * 指定レンジで3種取得 → D3 へ追記 + D1 の広告列を更新
 *
 * @param {string} startDate YYYY-MM-DD
 * @param {string} endDate   YYYY-MM-DD
 */
function fetchAdsReportsForRange(startDate, endDate) {
  setupAdsDetailSheets();

  // 1) spAdvertisedProduct → D3_ADS_ASIN + D1 更新
  try {
    const rows = fetchAdsReport('spAdvertisedProduct', startDate, endDate);
    writeAdvertisedProductRows(rows);
    updateDailyAdsFromAdvertisedProduct(rows);
    Logger.log('✅ spAdvertisedProduct (広告_商品別): ' + rows.length + ' 行');
  } catch (e) {
    Logger.log('❌ spAdvertisedProduct 失敗: ' + e.message);
  }

  // 2) spSearchTerm → D3_ADS_SEARCHTERM
  try {
    const rows = fetchAdsReport('spSearchTerm', startDate, endDate);
    writeSearchTermRows(rows);
    Logger.log('✅ spSearchTerm: ' + rows.length + ' 行');
  } catch (e) {
    Logger.log('❌ spSearchTerm 失敗: ' + e.message);
  }

  // 3) spTargeting → D3_ADS_TARGET
  try {
    const rows = fetchAdsReport('spTargeting', startDate, endDate);
    writeTargetingRows(rows);
    Logger.log('✅ spTargeting: ' + rows.length + ' 行');
  } catch (e) {
    Logger.log('❌ spTargeting 失敗: ' + e.message);
  }

  Logger.log('===== Ads Reports 完了 =====');
}

/**
 * 過去 N 日分を1日ずつ取得（手動バックフィル用）
 * GASの6分制限を考慮して N は 3〜5 程度が目安。
 */
function backfillAdsReportsRange(fromDaysAgo, toDaysAgo) {
  const today = new Date();
  const fmt = d => Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  for (let d = fromDaysAgo; d <= toDaysAgo; d++) {
    const target = new Date(today);
    target.setDate(target.getDate() - d);
    const dateStr = fmt(target);
    Logger.log('--- Ads ' + dateStr + ' ---');
    try {
      fetchAdsReportsForRange(dateStr, dateStr);
    } catch (e) {
      Logger.log('エラー: ' + e.message);
    }
  }
}

// ==========================================================
//  低レベル: レポート作成→ポーリング→DL→parse
// ==========================================================

/**
 * v3 非同期レポートを1本取得
 *
 * @param {string} reportKey  ADS_REPORT_CONFIGS のキー
 * @param {string} startDate  YYYY-MM-DD
 * @param {string} endDate    YYYY-MM-DD
 * @returns {Array<Object>}   レポート行（JSON 配列）
 */
function fetchAdsReport(reportKey, startDate, endDate) {
  const cfg = ADS_REPORT_CONFIGS[reportKey];
  if (!cfg) throw new Error('未知の reportKey: ' + reportKey);

  const payload = {
    name: reportKey + '_' + startDate + '_' + endDate + '_' + Date.now(),
    startDate: startDate,
    endDate: endDate,
    configuration: {
      adProduct: cfg.adProduct,
      groupBy: cfg.groupBy,
      columns: cfg.columns,
      reportTypeId: cfg.reportTypeId,
      timeUnit: cfg.timeUnit,
      format: 'GZIP_JSON',
    },
  };

  // 1) 作成
  const createRes = callAdsApi('POST', '/reporting/reports', payload, {
    contentType: ADS_V3_CT_CREATE_REPORT,
  });
  const reportId = createRes.reportId;
  if (!reportId) throw new Error('reportId が返りませんでした: ' + JSON.stringify(createRes));
  Logger.log('  [' + reportKey + '] reportId=' + reportId);

  // 2) ポーリング
  let downloadUrl = null;
  for (let i = 0; i < ADS_REPORT_POLL_MAX_ATTEMPTS; i++) {
    Utilities.sleep(ADS_REPORT_POLL_INTERVAL_MS);
    const st = callAdsApi('GET', '/reporting/reports/' + reportId);
    const s = st.status;
    if (s === 'COMPLETED') {
      downloadUrl = st.url || (st.location);
      break;
    }
    if (s === 'FAILED' || s === 'CANCELLED') {
      throw new Error('レポート生成失敗: ' + s + ' / ' + (st.failureReason || ''));
    }
    if ((i + 1) % 4 === 0) Logger.log('  [' + reportKey + '] status=' + s + ' (' + (i + 1) + '/' + ADS_REPORT_POLL_MAX_ATTEMPTS + ')');
  }
  if (!downloadUrl) throw new Error('レポート生成タイムアウト: reportId=' + reportId);

  // 3) ダウンロード + gzip 解凍 + JSON パース
  return downloadAndParseAdsReport(downloadUrl);
}

/**
 * pre-signed URL から gzip JSON を取得して配列化
 */
function downloadAndParseAdsReport(url) {
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) throw new Error('レポートDL失敗 ' + code);

  // Amazon Ads v3 は GZIP_JSON 指定なら gzip 済み
  let jsonText;
  try {
    const blob = res.getBlob().setContentType('application/x-gzip');
    jsonText = Utilities.ungzip(blob).getDataAsString('UTF-8');
  } catch (e) {
    // 素のJSONの可能性もあるのでフォールバック
    jsonText = res.getContentText();
  }

  const data = JSON.parse(jsonText);
  return Array.isArray(data) ? data : (data.data || []);
}

// ==========================================================
//  D3 シート作成 + 書き込み
// ==========================================================

function setupAdsDetailSheets() {
  setupAdsAsinSheet();
  setupAdsSearchTermSheet();
  setupAdsTargetSheet();
}

function setupAdsAsinSheet() {
  const sheet = getOrCreateSheet(SHEET_NAMES.D3_ADS_ASIN);
  const headers = [
    '日付', 'キャンペーンID', 'キャンペーン名', '広告グループID', '広告グループ名',
    'ASIN', 'SKU',
    'IMP', 'CT', '広告費', 'CTR', 'CPC',
    '注文数(7d)', '広告売上(7d)', '点数(7d)',
    'ACOS(7d)', 'ROAS(7d)',
  ];
  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() !== headers[0]) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#fff2cc');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function setupAdsSearchTermSheet() {
  const sheet = getOrCreateSheet(SHEET_NAMES.D3_ADS_SEARCHTERM);
  const headers = [
    '日付', 'キャンペーンID', 'キャンペーン名', '広告グループID', '広告グループ名',
    'キーワードID', 'キーワード', 'マッチタイプ', '検索用語',
    'IMP', 'CT', '広告費', 'CTR', 'CPC',
    '注文数(7d)', '広告売上(7d)',
    'ACOS(7d)', 'ROAS(7d)',
  ];
  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() !== headers[0]) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#fff2cc');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function setupAdsTargetSheet() {
  const sheet = getOrCreateSheet(SHEET_NAMES.D3_ADS_TARGET);
  const headers = [
    '日付', 'キャンペーンID', 'キャンペーン名', '広告グループID', '広告グループ名',
    'キーワードID', 'キーワード', 'マッチタイプ', 'ターゲット',
    'IMP', 'CT', '広告費', 'CTR', 'CPC',
    '注文数(7d)', '広告売上(7d)',
    'ACOS(7d)', 'ROAS(7d)',
  ];
  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() !== headers[0]) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#fff2cc');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 同じ（日付×キャンペーン×広告グループ×ASIN）の既存行を削除してから新規分を追記。
 * 二重取得時の重複を防ぐ。
 */
function writeAdvertisedProductRows(rows) {
  if (!rows || rows.length === 0) return;
  const sheet = setupAdsAsinSheet();
  const dates = new Set(rows.map(r => r.date));
  deleteAdsRowsForDates(sheet, dates, 1);

  const out = rows.map(r => [
    r.date || '',
    r.campaignId || '',
    r.campaignName || '',
    r.adGroupId || '',
    r.adGroupName || '',
    r.advertisedAsin || '',
    r.advertisedSku || '',
    r.impressions || 0,
    r.clicks || 0,
    r.cost || 0,
    r.clickThroughRate || 0,
    r.costPerClick || 0,
    r.purchases7d || 0,
    r.sales7d || 0,
    r.unitsSoldClicks7d || 0,
    r.acosClicks7d || 0,
    r.roasClicks7d || 0,
  ]);
  appendAdsRows(sheet, out);
}

function writeSearchTermRows(rows) {
  if (!rows || rows.length === 0) return;
  const sheet = setupAdsSearchTermSheet();
  const dates = new Set(rows.map(r => r.date));
  deleteAdsRowsForDates(sheet, dates, 1);

  const out = rows.map(r => [
    r.date || '',
    r.campaignId || '',
    r.campaignName || '',
    r.adGroupId || '',
    r.adGroupName || '',
    r.keywordId || '',
    r.keyword || '',
    r.matchType || '',
    r.searchTerm || '',
    r.impressions || 0,
    r.clicks || 0,
    r.cost || 0,
    r.clickThroughRate || 0,
    r.costPerClick || 0,
    r.purchases7d || 0,
    r.sales7d || 0,
    r.acosClicks7d || 0,
    r.roasClicks7d || 0,
  ]);
  appendAdsRows(sheet, out);
}

function writeTargetingRows(rows) {
  if (!rows || rows.length === 0) return;
  const sheet = setupAdsTargetSheet();
  const dates = new Set(rows.map(r => r.date));
  deleteAdsRowsForDates(sheet, dates, 1);

  const out = rows.map(r => [
    r.date || '',
    r.campaignId || '',
    r.campaignName || '',
    r.adGroupId || '',
    r.adGroupName || '',
    r.keywordId || '',
    r.keyword || '',
    r.matchType || '',
    r.targeting || '',
    r.impressions || 0,
    r.clicks || 0,
    r.cost || 0,
    r.clickThroughRate || 0,
    r.costPerClick || 0,
    r.purchases7d || 0,
    r.sales7d || 0,
    r.acosClicks7d || 0,
    r.roasClicks7d || 0,
  ]);
  appendAdsRows(sheet, out);
}

/**
 * 指定日集合に一致する行を削除（冪等性確保）
 */
function deleteAdsRowsForDates(sheet, datesSet, dateCol) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  const values = sheet.getRange(2, dateCol, lastRow - 1, 1).getValues();
  // 下から削除（インデックスずれ回避）
  for (let i = values.length - 1; i >= 0; i--) {
    const v = values[i][0];
    const dateStr = (v instanceof Date)
      ? Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd')
      : String(v).substring(0, 10);
    if (datesSet.has(dateStr)) {
      sheet.deleteRow(i + 2);
    }
  }
}

function appendAdsRows(sheet, rows) {
  if (!rows.length) return;
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

// ==========================================================
//  D1 日次データへの広告指標統合
// ==========================================================

/**
 * spAdvertisedProduct レポート行を ASIN×日付で集計し、
 * D1 日次データの広告4指標（広告費 / 広告売上 / IMP / CT）を更新する。
 *
 * D1 列（1-indexed）:
 *   P(16)=広告費 / Q(17)=広告売上 / R(18)=IMP / S(19)=CT
 */
function updateDailyAdsFromAdvertisedProduct(rows) {
  if (!rows || rows.length === 0) return;

  // ASIN × 日付で合算（同一ASINが複数キャンペーンに跨る場合がある）
  const agg = {};
  for (const r of rows) {
    const asin = r.advertisedAsin;
    const date = r.date;
    if (!asin || !date) continue;
    const key = date + '_' + asin;
    if (!agg[key]) agg[key] = { cost: 0, sales: 0, imp: 0, clicks: 0 };
    agg[key].cost += r.cost || 0;
    agg[key].sales += r.sales7d || 0;
    agg[key].imp += r.impressions || 0;
    agg[key].clicks += r.clicks || 0;
  }

  const sheet = getOrCreateSheet(SHEET_NAMES.D1_DAILY);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log('⚠️ D1 が空。先に売上取得が必要。');
    return;
  }

  // D1 の 日付/ASIN を読み込み → 一致行の広告列を更新
  const range = sheet.getRange(2, 1, lastRow - 1, 19);  // A〜S列（広告列まで）
  const data = range.getValues();

  let updated = 0;
  for (let i = 0; i < data.length; i++) {
    const d = data[i][0];
    const asin = data[i][1];
    if (!d || !asin) continue;
    const dateStr = (d instanceof Date)
      ? Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd')
      : String(d).substring(0, 10);
    const key = dateStr + '_' + asin;
    const a = agg[key];
    if (!a) continue;

    data[i][15] = Math.round(a.cost);   // P: 広告費
    data[i][16] = Math.round(a.sales);  // Q: 広告売上
    data[i][17] = a.imp;                // R: IMP
    data[i][18] = a.clicks;             // S: CT
    updated++;
  }

  // 広告列（P〜S）だけ書き戻し
  if (updated > 0) {
    const adCols = data.map(row => [row[15], row[16], row[17], row[18]]);
    sheet.getRange(2, 16, adCols.length, 4).setValues(adCols);
    Logger.log('  D1 更新: ' + updated + ' 行の広告4指標');
  } else {
    Logger.log('  ⚠️ D1 に一致する行なし（売上取得が先行しているか確認）');
  }
}

// ==========================================================
//  テスト / 診断用
// ==========================================================

/**
 * 昨日の spAdvertisedProduct レポートだけを試しに取得してログ出力
 */
function testFetchAdsReportYesterday() {
  const yesterday = getYesterday();
  Logger.log('===== test: spAdvertisedProduct ' + yesterday + ' =====');
  const rows = fetchAdsReport('spAdvertisedProduct', yesterday, yesterday);
  Logger.log('行数: ' + rows.length);
  if (rows.length > 0) {
    Logger.log('先頭行: ' + JSON.stringify(rows[0]));
    // ASIN別の広告費 TOP5
    const byAsin = {};
    for (const r of rows) {
      const a = r.advertisedAsin || '(none)';
      if (!byAsin[a]) byAsin[a] = { cost: 0, sales: 0 };
      byAsin[a].cost += r.cost || 0;
      byAsin[a].sales += r.sales7d || 0;
    }
    const top = Object.entries(byAsin).sort((a, b) => b[1].cost - a[1].cost).slice(0, 5);
    top.forEach(([asin, v]) => Logger.log('  ' + asin + ': cost=' + v.cost + ' sales=' + v.sales));
  }
}

/**
 * Ads レポート取得全体の疎通確認
 *   - 認証、プロファイル、レポート作成→DL、D3 書き込み、D1 更新まで通す
 */
function testAdsEndToEnd() {
  const yesterday = getYesterday();
  Logger.log('===== testAdsEndToEnd ' + yesterday + ' =====');
  fetchAdsReportsForRange(yesterday, yesterday);
}
