"""
トレードシグナル ダッシュボード (Flask版)
==========================================
MEXC取引高上位銘柄を移動平均線ロジックでスクリーニングし、
トレードシグナルを一覧表示する。
4時間足・日足のシグナルを並列表示。

起動方法:
  cd 13_trade_signals
  pip install -r requirements.txt
  python app.py

ブラウザで http://localhost:5000 を開く
"""

import sys
import os
import time
import pandas as pd
from datetime import datetime
from threading import Thread, Lock

sys.path.insert(0, os.path.dirname(__file__))

from flask import Flask, render_template_string, jsonify, request

from config import (
    SIGNAL_LABELS,
    TIMEFRAMES,
    MA_PERIODS,
    TOP_N_SYMBOLS,
    REFRESH_INTERVAL,
    USD_JPY_RATE,
)
from mexc_api import get_top_symbols, get_klines, get_klines_batch, get_usdjpy_rate, get_market_caps
from analyzer import add_moving_averages, analyze_symbol

app = Flask(__name__)

# === データキャッシュ ===
cache = {
    "data": None,
    "klines_1d": None,
    "klines_4h": None,
    "chart_data_1d": {},  # 事前計算済みチャートデータ（銘柄 -> dict）
    "chart_data_4h": {},
    "last_update": None,
    "loading": False,
    "error": None,
    "jpy_rate": None,
}


def build_chart_data(df: pd.DataFrame) -> dict:
    """ローソク足DataFrameからフロント用のチャートデータを構築"""
    if df is None or df.empty:
        return None
    df = add_moving_averages(df)
    data = {
        "timestamp": df["timestamp"].astype(str).tolist(),
        "open": df["open"].tolist(),
        "high": df["high"].tolist(),
        "low": df["low"].tolist(),
        "close": df["close"].tolist(),
    }
    for p in MA_PERIODS:
        col = f"MA{p}"
        if col in df.columns:
            data[col] = [None if pd.isna(v) else v for v in df[col].tolist()]
    return data
cache_lock = Lock()


def load_data(top_n=TOP_N_SYMBOLS):
    """4h・1d両方のデータを取得→分析する。"""
    # USD/JPYレートをリアルタイム取得
    jpy_rate = get_usdjpy_rate()

    top_symbols = get_top_symbols(n=top_n, jpy_rate=jpy_rate)
    symbol_list = [s["symbol"] for s in top_symbols]
    symbol_info = {s["symbol"]: s for s in top_symbols}

    # 時価総額を取得
    market_caps = get_market_caps(symbol_list)

    print(f"=== {len(symbol_list)}銘柄の日足データを取得中... ===")

    # 両方の時間足を取得
    klines_1d = get_klines_batch(symbol_list, TIMEFRAMES["1d"])
    print(f"=== {len(symbol_list)}銘柄の4時間足データを取得中... ===")
    klines_4h = get_klines_batch(symbol_list, TIMEFRAMES["4h"])
    print(f"=== 分析中... ===")

    signal_order = ["strong_bull", "bull", "bull_hint", "bear_hint", "bear", "strong_bear", "sideways"]

    results = []
    for symbol in symbol_list:
        info = symbol_info.get(symbol, {})

        # 日足分析
        sig_1d = {"signal": "sideways", "signal_label": SIGNAL_LABELS["sideways"],
                  "detail": "データなし", "close": 0, "ma_values": {}, "trend_strength": 0}
        if symbol in klines_1d:
            sig_1d_raw = analyze_symbol(klines_1d[symbol])
            sig_1d = {**sig_1d_raw, "signal_label": SIGNAL_LABELS.get(sig_1d_raw["signal"], "不明")}

        # 4時間足分析
        sig_4h = {"signal": "sideways", "signal_label": SIGNAL_LABELS["sideways"],
                  "detail": "データなし", "close": 0, "ma_values": {}, "trend_strength": 0}
        if symbol in klines_4h:
            sig_4h_raw = analyze_symbol(klines_4h[symbol])
            sig_4h = {**sig_4h_raw, "signal_label": SIGNAL_LABELS.get(sig_4h_raw["signal"], "不明")}

        price = sig_1d["close"] if sig_1d["close"] else sig_4h["close"]
        volume_usdt = info.get("volume", 0)
        volume_jpy = volume_usdt * jpy_rate

        avg_volume_30d_jpy = 0
        if symbol in klines_1d and len(klines_1d[symbol]) >= 30:
            df_1d = klines_1d[symbol]
            avg_vol = df_1d.tail(30)["volume"].mean()
            avg_volume_30d_jpy = avg_vol * price * jpy_rate

        market_cap_usd = market_caps.get(symbol, 0)
        market_cap_jpy = market_cap_usd * jpy_rate
        # 出来高/時価総額 回転率(%) = 24h出来高 / 時価総額 × 100
        volume_ratio = (volume_usdt / market_cap_usd * 100) if market_cap_usd > 0 else 0

        results.append({
            "symbol": symbol,
            "symbol_display": symbol.replace("USDT", "/USDT"),
            "price": price,
            "signal_1d": sig_1d["signal"],
            "signal_label_1d": sig_1d["signal_label"],
            "detail_1d": sig_1d["detail"],
            "flags_1d": sig_1d.get("flags", {}),
            "signal_4h": sig_4h["signal"],
            "signal_label_4h": sig_4h["signal_label"],
            "detail_4h": sig_4h["detail"],
            "flags_4h": sig_4h.get("flags", {}),
            "signal_1d_order": signal_order.index(sig_1d["signal"]) if sig_1d["signal"] in signal_order else 99,
            "signal_4h_order": signal_order.index(sig_4h["signal"]) if sig_4h["signal"] in signal_order else 99,
            "market_cap_jpy": market_cap_jpy,
            "volume_ratio": volume_ratio,
            "volume_jpy": volume_jpy,
            "avg_volume_30d_jpy": avg_volume_30d_jpy,
            "change_pct": info.get("priceChangePercent", 0),
            "ma5": sig_1d["ma_values"].get(5, 0),
            "ma10": sig_1d["ma_values"].get(10, 0),
            "ma20": sig_1d["ma_values"].get(20, 0),
            "ma50": sig_1d["ma_values"].get(50, 0),
            "ma100": sig_1d["ma_values"].get(100, 0),
        })

    # 日足シグナル優先度順にソート
    results.sort(key=lambda r: r["signal_1d_order"])

    return results, klines_1d, klines_4h, jpy_rate


def refresh_cache(top_n=TOP_N_SYMBOLS):
    """バックグラウンドでデータを更新する。"""
    with cache_lock:
        if cache["loading"]:
            return
        cache["loading"] = True
        cache["error"] = None

    try:
        results, klines_1d, klines_4h, jpy_rate = load_data(top_n)
        # チャートデータを事前計算
        print(f"=== チャートデータを事前構築中... ===")
        chart_data_1d = {s: build_chart_data(df) for s, df in klines_1d.items()}
        chart_data_4h = {s: build_chart_data(df) for s, df in klines_4h.items()}
        print(f"=== 事前構築完了: {len(chart_data_1d)}銘柄 ===")
        with cache_lock:
            cache["data"] = results
            cache["klines_1d"] = klines_1d
            cache["klines_4h"] = klines_4h
            cache["chart_data_1d"] = chart_data_1d
            cache["chart_data_4h"] = chart_data_4h
            cache["jpy_rate"] = jpy_rate
            cache["last_update"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    except Exception as e:
        with cache_lock:
            cache["error"] = str(e)
    finally:
        with cache_lock:
            cache["loading"] = False


# === HTMLテンプレート ===
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>トレードシグナル</title>
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI','Meiryo',sans-serif; background:#0e1117; color:#fafafa; }
        .header { background:#1a1d23; padding:20px 30px; border-bottom:1px solid #333; }
        .header h1 { font-size:24px; }
        .header .sub { color:#888; font-size:13px; margin-top:4px; }
        .controls { display:flex; gap:15px; align-items:center; padding:15px 30px; background:#1a1d23; border-bottom:1px solid #333; flex-wrap:wrap; }
        .controls select, .controls button, .controls input {
            padding:8px 16px; background:#262730; color:#fafafa; border:1px solid #444; border-radius:6px; font-size:14px;
        }
        .controls input { width:200px; }
        .controls button { background:#ff4b4b; border-color:#ff4b4b; font-weight:bold; cursor:pointer; }
        .controls button:hover { background:#ff6b6b; }
        .controls button:disabled { background:#555; border-color:#555; cursor:wait; }
        .summary { display:flex; gap:15px; padding:20px 30px; flex-wrap:wrap; }
        .summary-card { background:#1a1d23; border-radius:8px; padding:15px 20px; flex:1; min-width:140px; text-align:center; }
        .summary-card .num { font-size:28px; font-weight:bold; }
        .summary-card .label { font-size:12px; color:#888; margin-top:4px; }
        .content { padding:0 30px 30px; }
        .section-title { font-size:18px; margin:25px 0 12px; }
        table { width:100%; border-collapse:collapse; background:#1a1d23; border-radius:8px; overflow:hidden; }
        th { background:#262730; padding:10px 12px; text-align:left; font-size:12px; color:#888; white-space:nowrap; cursor:pointer; user-select:none; }
        th:hover { color:#fafafa; }
        th .sort-icon { margin-left:4px; font-size:10px; }
        td { padding:8px 12px; border-top:1px solid #262730; font-size:13px; }
        tr:hover td { background:#1e2028; }
        .signal-badge { display:inline-block; padding:2px 8px; border-radius:10px; font-size:11px; font-weight:bold; white-space:nowrap; }
        .signal-strong_bull { background:#00c853; color:#000; }
        .signal-bull { background:#2979ff; color:#fff; }
        .signal-bull_hint { background:#ffd600; color:#000; }
        .signal-bear_hint { background:#ff9100; color:#000; }
        .signal-bear { background:#ff1744; color:#fff; }
        .signal-strong_bear { background:#424242; color:#fff; }
        .signal-sideways { background:#616161; color:#fff; }
        .sig-badge { display:inline-block; padding:2px 6px; border-radius:6px; font-size:10px; font-weight:bold; margin-right:3px; min-width:18px; text-align:center; }
        .sb-half { background:#ff1744; color:#fff; }              /* 下半身=赤 */
        .sb-half-down { background:#7b1fa2; color:#fff; }         /* 逆下半身=紫 */
        .sb-kuchi { background:#ffd600; color:#000; }             /* クチバシ=黄 */
        .sb-kuchi-strong { background:#00c853; color:#000; }      /* 強いクチバシ=緑 */
        .sb-kuchi-down { background:#ff9100; color:#000; }        /* 逆クチバシ=橙 */
        .sb-kuchi-strong-down { background:#424242; color:#fff; } /* 強い逆クチバシ=黒 */
        .sb-ppp { background:#2979ff; color:#fff; }               /* PPP=青 */
        .sb-ppp-down { background:#5d4037; color:#fff; }          /* 逆PPP=茶 */
        .sb-mono { background:#9c27b0; color:#fff; }              /* ものわかれ=紫 */
        .sb-mono-down { background:#6a1b9a; color:#fff; }         /* 逆ものわかれ=濃紫 */
        .sb-100 { background:#e91e63; color:#fff; }               /* 100日線突破=ピンク */
        .sb-100-down { background:#37474f; color:#fff; }          /* 100日線割れ=暗灰 */
        .sb-warn { background:#ff5722; color:#fff; }              /* 15本警戒=橙赤 */
        .positive { color:#00c853; }
        .negative { color:#ff1744; }
        .charts-row { display:flex; gap:15px; margin-top:10px; flex-wrap:wrap; }
        .chart-box { flex:1; min-width:400px; background:#1a1d23; border-radius:8px; padding:10px; }
        .alert-box { background:#1a1d23; border-left:4px solid #ffd600; padding:12px 20px; margin:5px 0; border-radius:0 8px 8px 0; }
        .alert-box.bear-alert { border-left-color:#ff9100; }
        .loading { text-align:center; padding:60px; color:#888; font-size:18px; }
        .footer { text-align:center; padding:20px; color:#555; font-size:12px; }
        .ma-info { display:flex; gap:30px; padding:15px; flex-wrap:wrap; }
        .ma-info div { flex:1; min-width:200px; }
        .ma-info ul { list-style:none; margin-top:8px; }
        .ma-info li { padding:2px 0; font-size:13px; }
    </style>
</head>
<body>
    <div class="header">
        <h1>📊 トレードシグナル ダッシュボード</h1>
        <div class="sub">MEXC | 24h出来高5000万円以上 | 日足・4時間足 並列分析 | <span id="last-update">--</span> | USD/JPY: <span id="jpy-rate">--</span></div>
    </div>
    <div class="controls">
        <label>最低銘柄数: <select id="topn"><option value="10">10</option><option value="20" selected>20</option><option value="30">30</option><option value="50">50</option><option value="100">100</option></select></label>
        <label>フィルタ: <select id="signal-filter">
            <option value="">全て表示</option>
            <option value="strong_bull">🟢 強い上昇</option><option value="bull">🔵 上昇</option>
            <option value="bull_hint">🟡 上昇の兆し</option><option value="bear_hint">🟠 下落の兆し</option>
            <option value="bear">🔴 下落</option><option value="strong_bear">⚫ 強い下落</option>
            <option value="sideways">⚪ 横ばい</option>
        </select></label>
        <input type="text" id="search-box" placeholder="🔍 銘柄検索（例: BTC）" oninput="renderDashboard()">
        <button id="refresh-btn" onclick="refreshData()">🔄 データ更新</button>
    </div>
    <div id="main-content"><div class="loading">📡 データ取得中...（初回は2〜3分かかります）</div></div>
    <div class="footer">
        ⚠️ このツールは投資助言ではありません。投資判断は自己責任でお願いします。<br>
        📌 ロジック: 相場流5つの道具（下半身/本数/しこり/PPP/ものわかれ）+ 3段ロジック（風向き→合図→加速）
    </div>
<script>
let allData = [];
let symbolList = [];
let sortCol = '';
let sortAsc = true;
let chartCache = {};  // 'SYMBOL_tf' -> 描画データ
let dataVersion = 0;  // データ更新時にキャッシュを無効化するためのバージョン

async function refreshData() {
    const btn = document.getElementById('refresh-btn');
    btn.disabled = true; btn.textContent = '⏳ 取得中...';
    document.getElementById('main-content').innerHTML = '<div class="loading">📡 データ取得中...<br><small>50銘柄×2時間足を取得中（2〜3分）</small></div>';
    const n = document.getElementById('topn').value;
    try {
        await fetch('/api/refresh?topn=' + n);
        // ポーリングで完了を待つ
        pollForData();
    } catch(e) {
        document.getElementById('main-content').innerHTML = '<div class="loading">❌ ' + e.message + '</div>';
        btn.disabled = false; btn.textContent = '🔄 データ更新';
    }
}

function pollForData() {
    const btn = document.getElementById('refresh-btn');
    const poll = setInterval(async () => {
        try {
            const resp = await fetch('/api/data');
            const result = await resp.json();
            if (result.error) {
                clearInterval(poll);
                document.getElementById('main-content').innerHTML = '<div class="loading">❌ ' + result.error + '</div>';
                btn.disabled = false; btn.textContent = '🔄 データ更新';
            } else if (result.data && result.data.length > 0 && !result.loading) {
                clearInterval(poll);
                allData = result.data;
                symbolList = allData.map(d => d.symbol);
                document.getElementById('last-update').textContent = result.last_update || '--';
                renderDashboard();
                btn.disabled = false; btn.textContent = '🔄 データ更新';
            }
            // まだloading中なら待ち続ける
        } catch(e) { /* ネットワークエラーは無視して再試行 */ }
    }, 3000); // 3秒ごとにチェック
}

async function loadDashboard() {
    try {
        const resp = await fetch('/api/data');
        const result = await resp.json();
        if (!result.data || result.data.length === 0) {
            if (result.loading) {
                document.getElementById('main-content').innerHTML = '<div class="loading">📡 データ取得中...</div>';
                setTimeout(loadDashboard, 3000); return;
            }
            refreshData(); return;
        }
        // データ更新を検知したらキャッシュを無効化
        const newLastUpdate = result.last_update;
        const prevLastUpdate = document.getElementById('last-update').textContent;
        if (newLastUpdate && newLastUpdate !== prevLastUpdate && prevLastUpdate !== '--') {
            chartCache = {};
            dataVersion++;
        }
        allData = result.data;
        symbolList = allData.map(d => d.symbol);
        document.getElementById('last-update').textContent = newLastUpdate || '--';
        if (result.jpy_rate) document.getElementById('jpy-rate').textContent = result.jpy_rate.toFixed(2) + '円';
        renderDashboard();
        // バックグラウンドで全銘柄のチャートを先読み
        preloadCharts();
    } catch(e) { document.getElementById('main-content').innerHTML = '<div class="loading">❌ ' + e.message + '</div>'; }
}

function getFilteredData() {
    let data = allData;
    const filter = document.getElementById('signal-filter').value;
    if (filter) data = data.filter(d => d.signal_1d === filter);
    const search = (document.getElementById('search-box').value || '').toUpperCase().trim();
    if (search) data = data.filter(d => d.symbol.includes(search));
    return data;
}

function sortData(data, col) {
    const cmp = sortAsc ? 1 : -1;
    return [...data].sort((a, b) => {
        let va = a[col], vb = b[col];
        if (typeof va === 'string') return va.localeCompare(vb) * cmp;
        return ((va||0) - (vb||0)) * cmp;
    });
}

function onSort(col) {
    if (sortCol === col) { sortAsc = !sortAsc; } else { sortCol = col; sortAsc = true; }
    renderDashboard();
}

function sortIcon(col) {
    if (sortCol !== col) return '<span class="sort-icon">⇅</span>';
    return sortAsc ? '<span class="sort-icon">▲</span>' : '<span class="sort-icon">▼</span>';
}

function renderDashboard() {
    let data = getFilteredData();
    if (sortCol) data = sortData(data, sortCol);

    // サマリー（日足ベース）
    const counts = {};
    allData.forEach(d => { counts[d.signal_1d] = (counts[d.signal_1d]||0) + 1; });
    const bullCount = (counts.strong_bull||0) + (counts.bull||0);
    const bearCount = (counts.strong_bear||0) + (counts.bear||0);

    let html = '<div class="summary">' +
        '<div class="summary-card"><div class="num positive">' + bullCount + '</div><div class="label">🟢🔵 上昇（日足）</div></div>' +
        '<div class="summary-card"><div class="num" style="color:#ffd600">' + (counts.bull_hint||0) + '</div><div class="label">🟡 上昇の兆し</div></div>' +
        '<div class="summary-card"><div class="num">' + (counts.sideways||0) + '</div><div class="label">⚪ 横ばい</div></div>' +
        '<div class="summary-card"><div class="num" style="color:#ff9100">' + (counts.bear_hint||0) + '</div><div class="label">🟠 下落の兆し</div></div>' +
        '<div class="summary-card"><div class="num negative">' + bearCount + '</div><div class="label">🔴⚫ 下落（日足）</div></div>' +
    '</div><div class="content">';

    // === 個別チャート（一覧の上に表示）===
    html += '<h2 class="section-title">📈 個別チャート（日足・4時間足 並列表示）</h2>' +
        '<select id="chart-select" onchange="showChart(this.value)" style="padding:8px 16px;background:#262730;color:#fafafa;border:1px solid #444;border-radius:6px;font-size:14px;margin-bottom:10px">' +
        symbolList.map(s => '<option value="' + s + '">' + s.replace('USDT','/USDT') + '</option>').join('') + '</select>' +
        '<div class="charts-row"><div class="chart-box" id="chart-1d"></div><div class="chart-box" id="chart-4h"></div></div>' +
        '<div id="ma-details" class="ma-info"></div>';

    // === シグナル一覧 ===
    html += '<h2 class="section-title">📋 シグナル一覧（' + data.length + '銘柄）</h2>' +
        '<div style="font-size:11px;color:#888;margin-bottom:8px">凡例: ' +
        '<span class="sig-badge sb-half" title="下半身/逆下半身">半</span>' +
        '<span class="sig-badge sb-kuchi" title="基本クチバシ">ク</span>' +
        '<span class="sig-badge sb-kuchi-strong" title="強いクチバシ">強</span>' +
        '<span class="sig-badge sb-ppp" title="PPP/逆PPP">P</span>' +
        '<span class="sig-badge sb-mono" title="ものわかれ継続">も</span>' +
        '<span class="sig-badge sb-100" title="100日線突破">100</span>' +
        '<span class="sig-badge sb-warn" title="15本超え 利確警戒">15</span>' +
        '</div>' +
        '<table><thead><tr>' +
        '<th onclick="onSort(&quot;symbol_display&quot;)">銘柄' + sortIcon('symbol_display') + '</th>' +
        '<th onclick="onSort(&quot;signal_1d_order&quot;)">日足シグナル' + sortIcon('signal_1d_order') + '</th>' +
        '<th onclick="onSort(&quot;signal_4h_order&quot;)">4hシグナル' + sortIcon('signal_4h_order') + '</th>' +
        '<th>出現サイン (日足)</th>' +
        '<th onclick="onSort(&quot;price&quot;)">現在値' + sortIcon('price') + '</th>' +
        '<th onclick="onSort(&quot;market_cap_jpy&quot;)">時価総額(円)' + sortIcon('market_cap_jpy') + '</th>' +
        '<th onclick="onSort(&quot;volume_jpy&quot;)">24h出来高(円)' + sortIcon('volume_jpy') + '</th>' +
        '<th onclick="onSort(&quot;volume_ratio&quot;)" title="24h出来高÷時価総額。10%以上=非常に好条件、5%前後=良好、1%以下=効率低">回転率(%)' + sortIcon('volume_ratio') + '</th>' +
        '<th onclick="onSort(&quot;avg_volume_30d_jpy&quot;)">30日平均/日(円)' + sortIcon('avg_volume_30d_jpy') + '</th>' +
        '<th onclick="onSort(&quot;change_pct&quot;)">24h変動(%)' + sortIcon('change_pct') + '</th>' +
    '</tr></thead><tbody>';

    data.forEach(d => {
        const cc = d.change_pct >= 0 ? 'positive' : 'negative';
        html += '<tr onclick="showChart(&quot;' + d.symbol + '&quot;)" style="cursor:pointer">' +
            '<td><a href="https://www.tradingview.com/chart/?symbol=MEXC%3A' + d.symbol + '" target="_blank" style="color:#4fc3f7;text-decoration:none;font-weight:bold" onclick="event.stopPropagation()">' + d.symbol_display + '</a></td>' +
            '<td><span class="signal-badge signal-' + d.signal_1d + '">' + d.signal_label_1d + '</span></td>' +
            '<td><span class="signal-badge signal-' + d.signal_4h + '">' + d.signal_label_4h + '</span></td>' +
            '<td>' + renderFlagBadges(d.flags_1d || {}) + '</td>' +
            '<td>' + formatPrice(d.price) + '</td>' +
            '<td>' + (d.market_cap_jpy ? formatJPY(d.market_cap_jpy) : '-') + '</td>' +
            '<td>' + formatJPY(d.volume_jpy) + '</td>' +
            '<td>' + formatRatio(d.volume_ratio) + '</td>' +
            '<td>' + formatJPY(d.avg_volume_30d_jpy) + '</td>' +
            '<td class="' + cc + '">' + (d.change_pct >= 0 ? '+' : '') + d.change_pct.toFixed(2) + '%</td></tr>';
    });

    html += '</tbody></table>';

    html += '<h2 class="section-title">🔔 注目銘柄（シグナル変化あり）</h2>';
    const alerts = allData.filter(d => ['bull_hint','bear_hint'].includes(d.signal_1d) || ['bull_hint','bear_hint'].includes(d.signal_4h));
    if (!alerts.length) { html += '<p style="color:#888;padding:10px 0">現在、シグナル変化のある銘柄はありません。</p>'; }
    else {
        alerts.forEach(d => {
            let parts = [];
            if (['bull_hint','bear_hint'].includes(d.signal_1d)) parts.push('日足: ' + d.detail_1d);
            if (['bull_hint','bear_hint'].includes(d.signal_4h)) parts.push('4h: ' + d.detail_4h);
            const icon = (d.signal_1d === 'bull_hint' || d.signal_4h === 'bull_hint') ? '🟡' : '🟠';
            const cls = icon === '🟠' ? 'bear-alert' : '';
            html += '<div class="alert-box ' + cls + '">' + icon + ' <strong>' + d.symbol_display + '</strong> | ' + parts.join(' / ') + '</div>';
        });
    }
    html += '</div>';
    document.getElementById('main-content').innerHTML = html;

    // 初回レンダリング後、デフォルトで先頭銘柄のチャートを表示（スクロールなし）
    if (symbolList.length > 0) {
        showChart(symbolList[0], false);
    }
}

function formatPrice(p) {
    if (p >= 1000) return p.toFixed(2);
    if (p >= 1) return p.toFixed(4);
    return p.toFixed(6);
}
function formatJPY(v) {
    if (v >= 1e12) return (v/1e12).toFixed(1) + '兆円';
    if (v >= 1e8) return (v/1e8).toFixed(1) + '億円';
    if (v >= 1e4) return (v/1e4).toFixed(0) + '万円';
    return Math.round(v).toLocaleString() + '円';
}

function formatRatio(r) {
    if (!r || r === 0) return '-';
    let color = '#888';
    let icon = '';
    if (r >= 10) { color = '#00c853'; icon = '🟢 '; }       // 非常に好条件
    else if (r >= 5) { color = '#4fc3f7'; icon = '🔵 '; }   // 通常OK
    else if (r >= 1) { color = '#ffd600'; icon = '🟡 '; }   // やや低い
    else { color = '#ff6b6b'; icon = '🔴 '; }               // 低い
    return '<span style="color:' + color + ';font-weight:bold">' + icon + r.toFixed(2) + '%</span>';
}

function renderFlagBadges(f) {
    if (!f || Object.keys(f).length === 0) return '';
    let html = '';
    // 下半身/逆下半身
    if (f.lower_half_body) html += '<span class="sig-badge sb-half" title="下半身（陽線が5日線を上抜け）">半</span>';
    if (f.reverse_lower_half_body) html += '<span class="sig-badge sb-half-down" title="逆下半身（陰線が5日線を下抜け）">半</span>';
    // クチバシ（強いものを優先表示）
    if (f.kuchibashi_strong) html += '<span class="sig-badge sb-kuchi-strong" title="強いクチバシ（5/10/20日すべて上向き）">強</span>';
    else if (f.kuchibashi_basic) html += '<span class="sig-badge sb-kuchi" title="基本クチバシ（5日が10日を上抜け、同方向）">ク</span>';
    if (f.rev_kuchibashi_strong) html += '<span class="sig-badge sb-kuchi-strong-down" title="強い逆クチバシ（5/10/20日すべて下向き）">強</span>';
    else if (f.rev_kuchibashi_basic) html += '<span class="sig-badge sb-kuchi-down" title="基本逆クチバシ">ク</span>';
    // PPP/逆PPP
    if (f.ppp) html += '<span class="sig-badge sb-ppp" title="PPP成立（5>10>20>50>100）">P</span>';
    if (f.reverse_ppp) html += '<span class="sig-badge sb-ppp-down" title="逆PPP成立（5<10<20<50<100）">P</span>';
    // ものわかれ
    if (f.monowakare_up) html += '<span class="sig-badge sb-mono" title="ものわかれ継続（上昇トレンド継続）">も</span>';
    if (f.monowakare_down) html += '<span class="sig-badge sb-mono-down" title="ものわかれ継続（下降トレンド継続）">も</span>';
    // 100日線
    if (f.above_ma100) html += '<span class="sig-badge sb-100" title="100日線突破（大局上昇）">100</span>';
    if (f.below_ma100) html += '<span class="sig-badge sb-100-down" title="100日線割れ（大局下降）">100</span>';
    // 15本ルール
    if (f.bars_warn) html += '<span class="sig-badge sb-warn" title="' + f.bars_count + '本継続（15本超え 利確警戒）">15</span>';
    return html || '<span style="color:#666">-</span>';
}

async function fetchChartData(symbol, tf) {
    const key = symbol + '_' + tf + '_v' + dataVersion;
    if (chartCache[key]) return chartCache[key];
    try {
        const resp = await fetch('/api/chart?symbol=' + symbol + '&timeframe=' + tf);
        const result = await resp.json();
        if (result.error) return null;
        chartCache[key] = result.data;
        return result.data;
    } catch(e) { return null; }
}

async function renderChart(container, symbol, tf, tfLabel) {
    try {
        const d = await fetchChartData(symbol, tf);
        if (!d) { document.getElementById(container).innerHTML = '<p style="color:red">データ取得失敗</p>'; return; }
        const traces = [{x:d.timestamp,open:d.open,high:d.high,low:d.low,close:d.close,type:'candlestick',name:'ローソク足'}];
        const mc = {5:'#FF6B6B',10:'#4ECDC4',20:'#45B7D1',50:'#FFA07A',100:'#9B59B6'};
        [5,10,20,50,100].forEach(p => {
            if (d['MA'+p]) traces.push({x:d.timestamp,y:d['MA'+p],type:'scatter',mode:'lines',name:'MA'+p,line:{color:mc[p],width:1.5}});
        });
        Plotly.newPlot(container, traces, {
            title: symbol.replace('USDT','/USDT') + ' - ' + tfLabel,
            template:'plotly_dark', paper_bgcolor:'#1a1d23', plot_bgcolor:'#1a1d23',
            xaxis:{rangeslider:{visible:false}}, yaxis:{title:'USDT'},
            height:450, margin:{l:50,r:10,t:35,b:30}, legend:{orientation:'h',y:-0.15}
        }, {responsive:true});
    } catch(e) { document.getElementById(container).innerHTML = '<p style="color:red">' + e.message + '</p>'; }
}

// バックグラウンドで全銘柄のチャートデータを一括先読み
async function preloadCharts() {
    const startVersion = dataVersion;
    try {
        const resp = await fetch('/api/charts_bulk');
        if (startVersion !== dataVersion) return;  // 取得中に更新があれば破棄
        const result = await resp.json();
        const c1d = result.chart_1d || {};
        const c4h = result.chart_4h || {};
        for (const sym in c1d) chartCache[sym + '_1d_v' + dataVersion] = c1d[sym];
        for (const sym in c4h) chartCache[sym + '_4h_v' + dataVersion] = c4h[sym];
        console.log('チャート先読み完了:', Object.keys(c1d).length + '銘柄');
    } catch(e) { console.warn('先読み失敗:', e); }
}

async function showChart(symbol, scroll) {
    const sel = document.getElementById('chart-select');
    if (sel) sel.value = symbol;
    await Promise.all([renderChart('chart-1d',symbol,'1d','日足'), renderChart('chart-4h',symbol,'4h','4時間足')]);
    const info = allData.find(x => x.symbol === symbol);
    if (info) {
        document.getElementById('ma-details').innerHTML =
            '<div><strong>移動平均線（日足）:</strong><ul>' +
            '<li style="color:#FF6B6B">MA5: ' + formatPrice(info.ma5) + '</li>' +
            '<li style="color:#4ECDC4">MA10: ' + formatPrice(info.ma10) + '</li>' +
            '<li style="color:#45B7D1">MA20（風向き）: ' + formatPrice(info.ma20) + '</li>' +
            '<li style="color:#FFA07A">MA50: ' + formatPrice(info.ma50) + '</li>' +
            '<li style="color:#9B59B6">MA100: ' + formatPrice(info.ma100) + '</li></ul></div>' +
            '<div><strong>日足:</strong> <span class="signal-badge signal-' + info.signal_1d + '">' + info.signal_label_1d + '</span><br>' +
            '<span style="font-size:12px;color:#aaa">' + info.detail_1d + '</span><br><br>' +
            '<strong>4時間足:</strong> <span class="signal-badge signal-' + info.signal_4h + '">' + info.signal_label_4h + '</span><br>' +
            '<span style="font-size:12px;color:#aaa">' + info.detail_4h + '</span></div>';
    }
    // 行クリック時はチャート位置までスクロール
    if (scroll !== false) {
        const chartEl = document.getElementById('chart-1d');
        if (chartEl) chartEl.scrollIntoView({behavior:'smooth', block:'start'});
    }
}

document.getElementById('signal-filter').addEventListener('change', renderDashboard);
loadDashboard();
setInterval(loadDashboard, {{ refresh_interval }} * 1000);
</script>
</body>
</html>
"""


@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE, refresh_interval=REFRESH_INTERVAL)


@app.route("/api/data")
def api_data():
    with cache_lock:
        return jsonify({
            "data": cache["data"],
            "last_update": cache["last_update"],
            "loading": cache["loading"],
            "error": cache["error"],
            "jpy_rate": cache["jpy_rate"],
        })


@app.route("/api/refresh")
def api_refresh():
    n = int(request.args.get("topn", TOP_N_SYMBOLS))
    # 非同期で開始し、即座にレスポンスを返す
    thread = Thread(target=refresh_cache, args=(n,), daemon=True)
    thread.start()
    return jsonify({"status": "started"})


@app.route("/api/charts_bulk")
def api_charts_bulk():
    """
    全銘柄のチャートデータを一括取得（先読み用）。
    ネットワーク往復を減らし、一度の呼び出しで全て取得できる。
    """
    with cache_lock:
        chart_1d = cache.get("chart_data_1d") or {}
        chart_4h = cache.get("chart_data_4h") or {}
    return jsonify({
        "chart_1d": chart_1d,
        "chart_4h": chart_4h,
    })


@app.route("/api/chart")
def api_chart():
    symbol = request.args.get("symbol")
    tf = request.args.get("timeframe", "1d")

    # 事前計算済みデータから即座に返す（高速）
    with cache_lock:
        chart_cache = cache.get("chart_data_1d") if tf == "1d" else cache.get("chart_data_4h")
        if chart_cache and symbol in chart_cache:
            return jsonify({"data": chart_cache[symbol]})

    # キャッシュにない場合は動的に取得（フォールバック）
    try:
        df = get_klines(symbol, TIMEFRAMES[tf])
    except Exception as e:
        return jsonify({"error": str(e)})

    if df.empty:
        return jsonify({"error": "データなし"})

    data = build_chart_data(df)
    return jsonify({"data": data})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print("=" * 50)
    print("📊 トレードシグナル ダッシュボード")
    print("=" * 50)
    print(f"ブラウザで http://localhost:{port} を開いてください")
    print("終了するには Ctrl+C を押してください")
    print("=" * 50)
    app.run(host="0.0.0.0", port=port, debug=False)
