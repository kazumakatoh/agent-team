/**
 * MEXC取引所 API連携（現物／AIグリッド運用）
 *
 * 取得データ：
 *   - 保有通貨別残高（USDT, TIA, RENDER, LINK, ALGO, CAKE, ENA, PENDLE等）
 *   - 各通貨の時価（USD建て）
 *   - 取引手数料履歴（Maker 0.02% / Taker 0.06%）
 *   - SageMaster AIグリッドの実績損益
 *
 * 権限：読み取り専用（取引権限なし）
 * APIキー管理：Google Script Properties
 *
 * 参考：https://mexcdevelop.github.io/apidocs/spot_v3_en/
 */

// ======================================================
// 設定
// ======================================================
const MEXC_CONFIG = {
  BASE_URL: 'https://api.mexc.com',
  SHEET_HOLDINGS: '現物_月次',
  SHEET_RAW: 'raw_MEXC_trades',
  TARGET_COINS: ['USDT', 'TIA', 'RENDER', 'LINK', 'ALGO', 'CAKE', 'ENA', 'PENDLE'],
  MAKER_FEE: 0.0002,  // 0.02%
  TAKER_FEE: 0.0006   // 0.06%
};

// ======================================================
// メイン関数
// ======================================================
function runMEXCAggregation() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('MEXC_API_KEY');
  const secretKey = props.getProperty('MEXC_SECRET_KEY');

  if (!apiKey || !secretKey) {
    throw new Error('MEXC APIキーが未設定。Script Properties で設定してください。');
  }

  const balances = fetchAccountBalances_(apiKey, secretKey);
  const prices = fetchCurrentPrices_();
  const holdings = calculateHoldings_(balances, prices);

  updateHoldingsSheet_(holdings);

  // 取引履歴（手数料計算のため）
  const trades = fetchRecentTrades_(apiKey, secretKey, 30);  // 過去30日
  appendRawTrades_(trades);
}

// ======================================================
// 署名付きAPIリクエスト
// ======================================================
function mexcSignedRequest_(endpoint, params, apiKey, secretKey) {
  const timestamp = Date.now();
  const query = Object.keys(params)
    .map(k => `${k}=${encodeURIComponent(params[k])}`)
    .join('&');
  const queryWithTs = `${query}${query ? '&' : ''}timestamp=${timestamp}`;

  const signature = Utilities.computeHmacSha256Signature(queryWithTs, secretKey)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2))
    .join('');

  const url = `${MEXC_CONFIG.BASE_URL}${endpoint}?${queryWithTs}&signature=${signature}`;

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'X-MEXC-APIKEY': apiKey },
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  if (code !== 200) {
    throw new Error(`MEXC API error ${code}: ${response.getContentText()}`);
  }
  return JSON.parse(response.getContentText());
}

// ======================================================
// 口座残高取得
// ======================================================
function fetchAccountBalances_(apiKey, secretKey) {
  const account = mexcSignedRequest_('/api/v3/account', {}, apiKey, secretKey);
  return account.balances
    .filter(b => parseFloat(b.free) + parseFloat(b.locked) > 0)
    .map(b => ({
      asset: b.asset,
      free: parseFloat(b.free),
      locked: parseFloat(b.locked),
      total: parseFloat(b.free) + parseFloat(b.locked)
    }));
}

// ======================================================
// 現在価格取得（USDT建て）
// ======================================================
function fetchCurrentPrices_() {
  const url = `${MEXC_CONFIG.BASE_URL}/api/v3/ticker/price`;
  const response = UrlFetchApp.fetch(url);
  const data = JSON.parse(response.getContentText());

  const prices = { 'USDT': 1.0 };
  data.forEach(p => {
    if (p.symbol.endsWith('USDT')) {
      const coin = p.symbol.replace('USDT', '');
      prices[coin] = parseFloat(p.price);
    }
  });
  return prices;
}

// ======================================================
// 保有資産評価
// ======================================================
function calculateHoldings_(balances, prices) {
  const holdings = [];
  let totalUSD = 0;

  balances.forEach(b => {
    const price = prices[b.asset] || 0;
    const valueUSD = b.total * price;
    totalUSD += valueUSD;
    holdings.push({
      asset: b.asset,
      amount: b.total,
      priceUSD: price,
      valueUSD: valueUSD,
      ratio: 0  // 後で計算
    });
  });

  // 構成比率
  holdings.forEach(h => {
    h.ratio = totalUSD > 0 ? h.valueUSD / totalUSD : 0;
  });

  return {
    items: holdings.sort((a, b) => b.valueUSD - a.valueUSD),
    totalUSD: totalUSD
  };
}

// ======================================================
// 取引履歴取得
// ======================================================
function fetchRecentTrades_(apiKey, secretKey, days) {
  const startTime = Date.now() - days * 24 * 60 * 60 * 1000;
  const allTrades = [];

  MEXC_CONFIG.TARGET_COINS.forEach(coin => {
    if (coin === 'USDT') return;  // USDTペアは対象外
    try {
      const trades = mexcSignedRequest_(
        '/api/v3/myTrades',
        { symbol: `${coin}USDT`, startTime: startTime, limit: 1000 },
        apiKey, secretKey
      );
      trades.forEach(t => {
        allTrades.push({
          symbol: t.symbol,
          id: t.id,
          orderId: t.orderId,
          time: new Date(t.time),
          side: t.isBuyer ? 'BUY' : 'SELL',
          qty: parseFloat(t.qty),
          price: parseFloat(t.price),
          quoteQty: parseFloat(t.quoteQty),
          commission: parseFloat(t.commission),
          commissionAsset: t.commissionAsset,
          isMaker: t.isMaker
        });
      });
    } catch (e) {
      Logger.log(`取引履歴取得失敗（${coin}）: ${e.message}`);
    }
  });

  return allTrades;
}

// ======================================================
// 通貨別元本・利回り算出
// ======================================================

/**
 * 通貨別の初期元本を追跡（手動入力シート or 初回取得時刻）
 * 利回り = (現在価値 - 元本) / 元本
 */
function calculateYieldPerCoin_(holdings, initialPrincipals) {
  return holdings.items.map(h => {
    const principal = initialPrincipals[h.asset] || h.valueUSD;
    const profit = h.valueUSD - principal;
    const yieldPct = principal > 0 ? (profit / principal) * 100 : 0;
    return {
      ...h,
      principalUSD: principal,
      profitUSD: profit,
      yieldPct: yieldPct
    };
  });
}

// ======================================================
// 手数料集計
// ======================================================
function calculateTotalFees_(trades) {
  let totalFeeUSD = 0;
  trades.forEach(t => {
    // commission は commissionAsset 建て
    // USD換算は別途時価から
    const rate = t.commissionAsset === 'USDT' ? 1 : 0;  // 簡易版
    totalFeeUSD += t.commission * rate;
  });
  return totalFeeUSD;
}

// ======================================================
// シート書き込み
// ======================================================
function updateHoldingsSheet_(holdings) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(MEXC_CONFIG.SHEET_HOLDINGS);
  if (!sheet) {
    throw new Error(`シート「${MEXC_CONFIG.SHEET_HOLDINGS}」が見つかりません`);
  }

  // 今月列を特定して書き込み
  // 実装詳細はDay2で
}

function appendRawTrades_(trades) {
  // raw_MEXC_trades シートに重複排除で追記
  // 実装詳細はDay2で
}

// ======================================================
// 接続テスト
// ======================================================
function testMEXCConnection() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('MEXC_API_KEY');
  const secretKey = props.getProperty('MEXC_SECRET_KEY');

  if (!apiKey || !secretKey) {
    Logger.log('❌ APIキー未設定');
    return;
  }
  Logger.log('✅ APIキー読み込みOK');

  try {
    const balances = fetchAccountBalances_(apiKey, secretKey);
    Logger.log(`✅ 口座残高取得成功：${balances.length}通貨`);
    balances.forEach(b => Logger.log(`  ${b.asset}: ${b.total}`));
  } catch (e) {
    Logger.log(`❌ 残高取得失敗：${e.message}`);
    return;
  }

  try {
    const prices = fetchCurrentPrices_();
    Logger.log(`✅ 価格取得成功：${Object.keys(prices).length}銘柄`);
    MEXC_CONFIG.TARGET_COINS.forEach(c => {
      Logger.log(`  ${c}USDT: $${prices[c] || 'N/A'}`);
    });
  } catch (e) {
    Logger.log(`❌ 価格取得失敗：${e.message}`);
    return;
  }

  Logger.log('🎉 MEXC接続テスト完了！');
}

// ======================================================
// スプシへ書き込みテスト（現物スナップショット）
// ======================================================
function testWriteHoldingsToSheet() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('MEXC_API_KEY');
  const secretKey = props.getProperty('MEXC_SECRET_KEY');

  const balances = fetchAccountBalances_(apiKey, secretKey);
  const prices = fetchCurrentPrices_();
  const holdings = calculateHoldings_(balances, prices);

  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName('現物_スナップショット');
  if (!sheet) sheet = ss.insertSheet('現物_スナップショット');
  else sheet.clear();

  // ヘッダー
  sheet.getRange('A1:F1')
    .setValues([['取得日時', '通貨', '保有量', '時価($)', '評価額($)', '構成比']])
    .setFontWeight('bold')
    .setBackground('#4a86e8')
    .setFontColor('white')
    .setHorizontalAlignment('center');

  const now = new Date();
  const rows = holdings.items.map(h => [
    now, h.asset, h.amount, h.priceUSD, h.valueUSD, h.ratio
  ]);
  rows.push(['', '合計', '', '', holdings.totalUSD, 1]);

  sheet.getRange(2, 1, rows.length, 6).setValues(rows);

  const lastRow = rows.length + 1;
  sheet.getRange(2, 1, rows.length, 1).setNumberFormat('yyyy/MM/dd HH:mm:ss').setHorizontalAlignment('center');
  sheet.getRange(2, 2, rows.length, 1).setHorizontalAlignment('center');
  sheet.getRange(2, 3, rows.length, 1).setNumberFormat('#,##0');
  sheet.getRange(2, 4, rows.length, 1).setNumberFormat('$#,##0.0000');
  sheet.getRange(2, 5, rows.length, 1).setNumberFormat('$#,##0.00');
  sheet.getRange(2, 6, rows.length, 1).setNumberFormat('0.00%');
  sheet.getRange(lastRow, 1, 1, 6).setFontWeight('bold').setBackground('#fff2cc');

  sheet.getRange(1, 1, lastRow, 6).setFontSize(10);
  sheet.autoResizeColumn(1);
  sheet.setColumnWidths(2, 5, 110);

  Logger.log(`✅ 書き込み完了：${holdings.items.length}通貨、総資産$${holdings.totalUSD.toFixed(2)}`);
}
