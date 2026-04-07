"""
トレードシグナル ダッシュボード (Flask版)
==========================================
MEXC取引高上位銘柄を移動平均線ロジックでスクリーニングし、
トレードシグナルを一覧表示する。

起動方法:
  cd 13_trade_signals
  pip install -r requirements.txt
  python app.py

ブラウザで http://localhost:5000 を開く
"""

import sys
import os
import json
import time
from datetime import datetime
from threading import Thread, Lock

sys.path.insert(0, os.path.dirname(__file__))

from flask import Flask, render_template_string, jsonify, request
import plotly
import plotly.graph_objects as go

from config import (
    SIGNAL_LABELS,
    TIMEFRAMES,
    MA_PERIODS,
    TOP_N_SYMBOLS,
    REFRESH_INTERVAL,
)
from mexc_api import get_top_symbols, get_klines, get_klines_batch
from analyzer import add_moving_averages, analyze_symbol

app = Flask(__name__)

# === データキャッシュ ===
cache = {
    "data": None,
    "klines": None,
    "last_update": None,
    "loading": False,
    "error": None,
}
cache_lock = Lock()


def load_data(timeframe="1d", top_n=TOP_N_SYMBOLS):
    """データ取得→分析を実行する。"""
    top_symbols = get_top_symbols(top_n)
    symbol_list = [s["symbol"] for s in top_symbols]
    symbol_info = {s["symbol"]: s for s in top_symbols}

    klines_data = get_klines_batch(symbol_list, TIMEFRAMES[timeframe])

    results = []
    for symbol, df in klines_data.items():
        signal = analyze_symbol(df)
        info = symbol_info.get(symbol, {})
        results.append({
            "symbol": symbol,
            "symbol_display": symbol.replace("USDT", "/USDT"),
            "price": signal["close"],
            "signal": signal["signal"],
            "signal_label": SIGNAL_LABELS.get(signal["signal"], "不明"),
            "trend_strength": signal["trend_strength"],
            "detail": signal["detail"],
            "volume": info.get("volume", 0),
            "change_pct": info.get("priceChangePercent", 0),
            "ma5": signal["ma_values"].get(5, 0),
            "ma10": signal["ma_values"].get(10, 0),
            "ma30": signal["ma_values"].get(30, 0),
            "ma50": signal["ma_values"].get(50, 0),
            "ma100": signal["ma_values"].get(100, 0),
        })

    # シグナル優先度順にソート
    signal_order = ["strong_bull", "bull", "bull_hint", "bear_hint", "bear", "strong_bear", "sideways"]
    results.sort(key=lambda r: signal_order.index(r["signal"]) if r["signal"] in signal_order else 99)

    return results, klines_data


def refresh_cache(timeframe="1d", top_n=TOP_N_SYMBOLS):
    """バックグラウンドでデータを更新する。"""
    with cache_lock:
        if cache["loading"]:
            return
        cache["loading"] = True
        cache["error"] = None

    try:
        results, klines_data = load_data(timeframe, top_n)
        with cache_lock:
            cache["data"] = results
            cache["klines"] = klines_data
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
        body {
            font-family: 'Segoe UI', 'Meiryo', sans-serif;
            background: #0e1117;
            color: #fafafa;
        }
        .header {
            background: #1a1d23;
            padding: 20px 30px;
            border-bottom: 1px solid #333;
        }
        .header h1 { font-size: 24px; }
        .header .sub { color: #888; font-size: 13px; margin-top: 4px; }
        .controls {
            display: flex;
            gap: 15px;
            align-items: center;
            padding: 15px 30px;
            background: #1a1d23;
            border-bottom: 1px solid #333;
            flex-wrap: wrap;
        }
        .controls select, .controls button {
            padding: 8px 16px;
            background: #262730;
            color: #fafafa;
            border: 1px solid #444;
            border-radius: 6px;
            font-size: 14px;
            cursor: pointer;
        }
        .controls button { background: #ff4b4b; border-color: #ff4b4b; font-weight: bold; }
        .controls button:hover { background: #ff6b6b; }
        .controls button:disabled { background: #555; border-color: #555; cursor: wait; }
        .summary {
            display: flex;
            gap: 15px;
            padding: 20px 30px;
            flex-wrap: wrap;
        }
        .summary-card {
            background: #1a1d23;
            border-radius: 8px;
            padding: 15px 20px;
            flex: 1;
            min-width: 140px;
            text-align: center;
        }
        .summary-card .num { font-size: 28px; font-weight: bold; }
        .summary-card .label { font-size: 12px; color: #888; margin-top: 4px; }
        .content { padding: 0 30px 30px; }
        .section-title { font-size: 18px; margin: 25px 0 12px; }
        table {
            width: 100%;
            border-collapse: collapse;
            background: #1a1d23;
            border-radius: 8px;
            overflow: hidden;
        }
        th {
            background: #262730;
            padding: 12px 15px;
            text-align: left;
            font-size: 13px;
            color: #888;
            white-space: nowrap;
        }
        td {
            padding: 10px 15px;
            border-top: 1px solid #262730;
            font-size: 14px;
        }
        tr:hover td { background: #1e2028; }
        .signal-badge {
            display: inline-block;
            padding: 3px 10px;
            border-radius: 12px;
            font-size: 12px;
            font-weight: bold;
            white-space: nowrap;
        }
        .signal-strong_bull { background: #00c853; color: #000; }
        .signal-bull { background: #2979ff; color: #fff; }
        .signal-bull_hint { background: #ffd600; color: #000; }
        .signal-bear_hint { background: #ff9100; color: #000; }
        .signal-bear { background: #ff1744; color: #fff; }
        .signal-strong_bear { background: #424242; color: #fff; }
        .signal-sideways { background: #616161; color: #fff; }
        .positive { color: #00c853; }
        .negative { color: #ff1744; }
        #chart-container {
            background: #1a1d23;
            border-radius: 8px;
            padding: 15px;
            margin-top: 10px;
        }
        .alert-box {
            background: #1a1d23;
            border-left: 4px solid #ffd600;
            padding: 12px 20px;
            margin: 5px 0;
            border-radius: 0 8px 8px 0;
        }
        .alert-box.bear-alert { border-left-color: #ff9100; }
        .loading {
            text-align: center;
            padding: 60px;
            color: #888;
            font-size: 18px;
        }
        .footer {
            text-align: center;
            padding: 20px;
            color: #555;
            font-size: 12px;
        }
        .ma-info { display: flex; gap: 30px; padding: 15px; flex-wrap: wrap; }
        .ma-info div { flex: 1; min-width: 200px; }
        .ma-info ul { list-style: none; margin-top: 8px; }
        .ma-info li { padding: 2px 0; font-size: 13px; }
    </style>
</head>
<body>
    <div class="header">
        <h1>📊 トレードシグナル ダッシュボード</h1>
        <div class="sub">MEXC 取引高上位銘柄 | 移動平均線分析 | <span id="last-update">--</span></div>
    </div>

    <div class="controls">
        <label>時間足:
            <select id="timeframe">
                <option value="1d" selected>日足</option>
                <option value="4h">4時間足</option>
            </select>
        </label>
        <label>銘柄数:
            <select id="topn">
                <option value="20">20</option>
                <option value="30">30</option>
                <option value="50" selected>50</option>
                <option value="100">100</option>
            </select>
        </label>
        <label>フィルタ:
            <select id="signal-filter">
                <option value="">全て表示</option>
                <option value="strong_bull">🟢 強い上昇</option>
                <option value="bull">🔵 上昇</option>
                <option value="bull_hint">🟡 上昇の兆し</option>
                <option value="bear_hint">🟠 下落の兆し</option>
                <option value="bear">🔴 下落</option>
                <option value="strong_bear">⚫ 強い下落</option>
                <option value="sideways">⚪ 横ばい</option>
            </select>
        </label>
        <button id="refresh-btn" onclick="refreshData()">🔄 データ更新</button>
    </div>

    <div id="main-content">
        <div class="loading" id="loading-msg">📡 データ取得中...（初回は1〜2分かかります）</div>
    </div>

    <div class="footer">
        ⚠️ このツールは投資助言ではありません。投資判断は自己責任でお願いします。<br>
        📌 ロジックは暫定版です。相場師朗ロジックのPDF受領後に差し替えます。
    </div>

    <script>
        let allData = [];
        let symbolList = [];

        async function refreshData() {
            const btn = document.getElementById('refresh-btn');
            btn.disabled = true;
            btn.textContent = '⏳ 取得中...';
            document.getElementById('main-content').innerHTML =
                '<div class="loading">📡 データ取得中...（1〜2分かかります）</div>';

            const tf = document.getElementById('timeframe').value;
            const n = document.getElementById('topn').value;

            try {
                const resp = await fetch(`/api/refresh?timeframe=${tf}&topn=${n}`);
                const result = await resp.json();
                if (result.error) {
                    document.getElementById('main-content').innerHTML =
                        `<div class="loading">❌ エラー: ${result.error}</div>`;
                } else {
                    loadDashboard();
                }
            } catch(e) {
                document.getElementById('main-content').innerHTML =
                    `<div class="loading">❌ 通信エラー: ${e.message}</div>`;
            } finally {
                btn.disabled = false;
                btn.textContent = '🔄 データ更新';
            }
        }

        async function loadDashboard() {
            try {
                const resp = await fetch('/api/data');
                const result = await resp.json();

                if (!result.data || result.data.length === 0) {
                    if (result.loading) {
                        document.getElementById('main-content').innerHTML =
                            '<div class="loading">📡 データ取得中...</div>';
                        setTimeout(loadDashboard, 3000);
                        return;
                    }
                    refreshData();
                    return;
                }

                allData = result.data;
                symbolList = allData.map(d => d.symbol);
                document.getElementById('last-update').textContent = result.last_update || '--';
                renderDashboard();
            } catch(e) {
                document.getElementById('main-content').innerHTML =
                    `<div class="loading">❌ ${e.message}</div>`;
            }
        }

        function renderDashboard() {
            const filter = document.getElementById('signal-filter').value;
            const data = filter ? allData.filter(d => d.signal === filter) : allData;

            // サマリー集計
            const counts = {};
            allData.forEach(d => { counts[d.signal] = (counts[d.signal] || 0) + 1; });
            const bullCount = (counts.strong_bull || 0) + (counts.bull || 0);
            const bearCount = (counts.strong_bear || 0) + (counts.bear || 0);

            let html = `
            <div class="summary">
                <div class="summary-card"><div class="num positive">${bullCount}</div><div class="label">🟢🔵 上昇相場</div></div>
                <div class="summary-card"><div class="num" style="color:#ffd600">${counts.bull_hint || 0}</div><div class="label">🟡 上昇の兆し</div></div>
                <div class="summary-card"><div class="num">${counts.sideways || 0}</div><div class="label">⚪ 横ばい</div></div>
                <div class="summary-card"><div class="num" style="color:#ff9100">${counts.bear_hint || 0}</div><div class="label">🟠 下落の兆し</div></div>
                <div class="summary-card"><div class="num negative">${bearCount}</div><div class="label">🔴⚫ 下落相場</div></div>
            </div>
            <div class="content">
                <h2 class="section-title">📋 シグナル一覧（${data.length}銘柄）</h2>
                <table>
                    <thead><tr>
                        <th>銘柄</th><th>シグナル</th><th>現在値</th>
                        <th>強度(%)</th><th>24h変動(%)</th><th>判定理由</th>
                    </tr></thead>
                    <tbody>`;

            data.forEach(d => {
                const changeClass = d.change_pct >= 0 ? 'positive' : 'negative';
                html += `<tr onclick="showChart('${d.symbol}')" style="cursor:pointer">
                    <td><strong>${d.symbol_display}</strong></td>
                    <td><span class="signal-badge signal-${d.signal}">${d.signal_label}</span></td>
                    <td>${formatPrice(d.price)}</td>
                    <td>${d.trend_strength.toFixed(2)}</td>
                    <td class="${changeClass}">${d.change_pct >= 0 ? '+' : ''}${d.change_pct.toFixed(2)}%</td>
                    <td style="font-size:12px;color:#aaa">${d.detail}</td>
                </tr>`;
            });

            html += `</tbody></table>

                <h2 class="section-title">📈 個別チャート</h2>
                <select id="chart-select" onchange="showChart(this.value)" style="padding:8px 16px;background:#262730;color:#fafafa;border:1px solid #444;border-radius:6px;font-size:14px;margin-bottom:10px">
                    ${symbolList.map(s => `<option value="${s}">${s.replace('USDT','/USDT')}</option>`).join('')}
                </select>
                <div id="chart-container"><div style="text-align:center;padding:40px;color:#888">銘柄を選択またはテーブルの行をクリック</div></div>
                <div id="ma-details" class="ma-info"></div>

                <h2 class="section-title">🔔 注目銘柄（シグナル変化あり）</h2>`;

            const alerts = allData.filter(d => d.signal === 'bull_hint' || d.signal === 'bear_hint');
            if (alerts.length === 0) {
                html += '<p style="color:#888;padding:10px 0">現在、シグナル変化のある銘柄はありません。</p>';
            } else {
                alerts.forEach(d => {
                    const cls = d.signal === 'bull_hint' ? '' : 'bear-alert';
                    const icon = d.signal === 'bull_hint' ? '🟡' : '🟠';
                    html += `<div class="alert-box ${cls}">${icon} <strong>${d.symbol_display}</strong> | 現在値: ${formatPrice(d.price)} | ${d.detail}</div>`;
                });
            }

            html += '</div>';
            document.getElementById('main-content').innerHTML = html;
        }

        function formatPrice(p) {
            if (p >= 1000) return p.toFixed(2);
            if (p >= 1) return p.toFixed(4);
            return p.toFixed(6);
        }

        async function showChart(symbol) {
            const sel = document.getElementById('chart-select');
            if (sel) sel.value = symbol;

            const tf = document.getElementById('timeframe').value;
            const tfLabel = tf === '1d' ? '日足' : '4時間足';

            try {
                const resp = await fetch(`/api/chart?symbol=${symbol}&timeframe=${tf}`);
                const result = await resp.json();

                if (result.error) {
                    document.getElementById('chart-container').innerHTML = `<p style="color:red">${result.error}</p>`;
                    return;
                }

                const d = result.data;
                const traces = [{
                    x: d.timestamp, open: d.open, high: d.high, low: d.low, close: d.close,
                    type: 'candlestick', name: 'ローソク足'
                }];

                const maColors = {5:'#FF6B6B',10:'#4ECDC4',30:'#45B7D1',50:'#FFA07A',100:'#9B59B6'};
                [5,10,30,50,100].forEach(p => {
                    if (d['MA'+p]) {
                        traces.push({x:d.timestamp, y:d['MA'+p], type:'scatter', mode:'lines',
                            name:'MA'+p, line:{color:maColors[p],width:1.5}});
                    }
                });

                const layout = {
                    title: `${symbol.replace('USDT','/USDT')} - ${tfLabel}`,
                    template: 'plotly_dark',
                    paper_bgcolor: '#1a1d23', plot_bgcolor: '#1a1d23',
                    xaxis: {title:'日時', rangeslider:{visible:false}},
                    yaxis: {title:'価格 (USDT)'},
                    height: 500,
                    margin: {l:60,r:20,t:40,b:40},
                };

                Plotly.newPlot('chart-container', traces, layout, {responsive:true});

                // MA詳細
                const info = allData.find(x => x.symbol === symbol);
                if (info) {
                    document.getElementById('ma-details').innerHTML = `
                        <div>
                            <strong>移動平均線の値:</strong>
                            <ul>
                                <li style="color:#FF6B6B">MA5: ${formatPrice(info.ma5)}</li>
                                <li style="color:#4ECDC4">MA10: ${formatPrice(info.ma10)}</li>
                                <li style="color:#45B7D1">MA30: ${formatPrice(info.ma30)}</li>
                                <li style="color:#FFA07A">MA50: ${formatPrice(info.ma50)}</li>
                                <li style="color:#9B59B6">MA100: ${formatPrice(info.ma100)}</li>
                            </ul>
                        </div>
                        <div>
                            <strong>シグナル:</strong> <span class="signal-badge signal-${info.signal}">${info.signal_label}</span><br><br>
                            <strong>判定理由:</strong> ${info.detail}<br>
                            <strong>トレンド強度:</strong> ${info.trend_strength.toFixed(2)}%
                        </div>`;
                }
            } catch(e) {
                document.getElementById('chart-container').innerHTML = `<p style="color:red">チャート取得エラー: ${e.message}</p>`;
            }
        }

        // フィルタ変更時に再描画
        document.getElementById('signal-filter').addEventListener('change', renderDashboard);

        // 初回読み込み
        loadDashboard();

        // 自動更新
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
        })


@app.route("/api/refresh")
def api_refresh():
    tf = request.args.get("timeframe", "1d")
    n = int(request.args.get("topn", TOP_N_SYMBOLS))

    # バックグラウンドで更新
    thread = Thread(target=refresh_cache, args=(tf, n))
    thread.start()
    thread.join(timeout=180)  # 最大3分待つ

    with cache_lock:
        if cache["error"]:
            return jsonify({"error": cache["error"]})
        return jsonify({"status": "ok"})


@app.route("/api/chart")
def api_chart():
    symbol = request.args.get("symbol")
    tf = request.args.get("timeframe", "1d")

    with cache_lock:
        klines = cache.get("klines")

    if not klines or symbol not in klines:
        # キャッシュにない場合は直接取得
        try:
            df = get_klines(symbol, TIMEFRAMES[tf])
        except Exception as e:
            return jsonify({"error": str(e)})
    else:
        df = klines[symbol]

    if df.empty:
        return jsonify({"error": "データなし"})

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
            data[col] = df[col].tolist()

    return jsonify({"data": data})


if __name__ == "__main__":
    print("=" * 50)
    print("📊 トレードシグナル ダッシュボード")
    print("=" * 50)
    print(f"ブラウザで http://localhost:5000 を開いてください")
    print(f"終了するには Ctrl+C を押してください")
    print("=" * 50)
    app.run(host="0.0.0.0", port=5000, debug=False)
