"""
トレードシグナル ダッシュボード
=================================
MEXC取引高上位銘柄を移動平均線ロジックでスクリーニングし、
トレードシグナルを一覧表示する。

起動方法:
  cd 13_trade_signals
  pip install -r requirements.txt
  streamlit run app.py
"""

import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime

from config import (
    SIGNAL_LABELS,
    TIMEFRAMES,
    MA_PERIODS,
    TOP_N_SYMBOLS,
    REFRESH_INTERVAL,
)
from mexc_api import get_top_symbols, get_klines, get_klines_batch
from analyzer import add_moving_averages, analyze_symbol

# === ページ設定 ===
st.set_page_config(
    page_title="トレードシグナル",
    page_icon="📊",
    layout="wide",
)

st.title("📊 トレードシグナル ダッシュボード")
st.caption(f"MEXC 取引高上位{TOP_N_SYMBOLS}銘柄 | 自動更新: {REFRESH_INTERVAL // 60}分ごと")


# === サイドバー ===
with st.sidebar:
    st.header("⚙️ 設定")

    timeframe = st.selectbox(
        "時間足",
        options=list(TIMEFRAMES.keys()),
        format_func=lambda x: {"4h": "4時間足", "1d": "日足"}[x],
        index=1,  # デフォルト: 日足
    )

    top_n = st.slider("対象銘柄数", min_value=10, max_value=100, value=TOP_N_SYMBOLS, step=10)

    signal_filter = st.multiselect(
        "シグナルフィルタ",
        options=list(SIGNAL_LABELS.keys()),
        format_func=lambda x: SIGNAL_LABELS[x],
        default=None,
        help="選択したシグナルのみ表示。空欄なら全表示。",
    )

    if st.button("🔄 データ更新", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.divider()
    st.markdown("### 📖 シグナル凡例")
    for key, label in SIGNAL_LABELS.items():
        st.markdown(f"- {label}")

    st.divider()
    st.markdown(
        "**⚠️ 暫定ロジック**\n\n"
        "パーフェクトオーダー＋\n"
        "ゴールデンクロス/デッドクロス\n\n"
        "相場師朗ロジックは\nPDF受領後に差し替え予定"
    )


# === データ取得・分析 ===
@st.cache_data(ttl=REFRESH_INTERVAL)
def load_data(tf: str, n: int):
    """データ取得→分析を実行してキャッシュする。"""

    # 1. 取引高上位銘柄を取得
    top_symbols = get_top_symbols(n)
    symbol_list = [s["symbol"] for s in top_symbols]
    symbol_info = {s["symbol"]: s for s in top_symbols}

    # 2. ローソク足を一括取得
    klines_data = get_klines_batch(symbol_list, TIMEFRAMES[tf])

    # 3. 各銘柄を分析
    results = []
    for symbol, df in klines_data.items():
        signal = analyze_symbol(df)
        info = symbol_info.get(symbol, {})
        results.append({
            "銘柄": symbol.replace("USDT", "/USDT"),
            "symbol_raw": symbol,
            "現在値": signal["close"],
            "シグナル": signal["signal"],
            "シグナル表示": SIGNAL_LABELS.get(signal["signal"], "不明"),
            "トレンド強度(%)": signal["trend_strength"],
            "判定理由": signal["detail"],
            "24h出来高(USDT)": info.get("volume", 0),
            "24h変動率(%)": info.get("priceChangePercent", 0),
            "MA5": signal["ma_values"].get(5, 0),
            "MA10": signal["ma_values"].get(10, 0),
            "MA30": signal["ma_values"].get(30, 0),
            "MA50": signal["ma_values"].get(50, 0),
            "MA100": signal["ma_values"].get(100, 0),
        })

    return pd.DataFrame(results), klines_data


# データ読み込み
with st.spinner("📡 MEXC APIからデータ取得中...（初回は1〜2分かかります）"):
    try:
        df_results, klines_data = load_data(timeframe, top_n)
    except Exception as e:
        st.error(f"データ取得に失敗しました: {e}")
        st.stop()

# === シグナルフィルタ適用 ===
if signal_filter:
    df_display = df_results[df_results["シグナル"].isin(signal_filter)]
else:
    df_display = df_results

# === サマリー表示 ===
st.markdown(f"**最終更新:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | **時間足:** {'日足' if timeframe == '1d' else '4時間足'}")

col1, col2, col3, col4, col5 = st.columns(5)

signal_counts = df_results["シグナル"].value_counts()

with col1:
    bull_count = signal_counts.get("strong_bull", 0) + signal_counts.get("bull", 0)
    st.metric("🟢🔵 上昇相場", f"{bull_count} 銘柄")

with col2:
    st.metric("🟡 上昇の兆し", f"{signal_counts.get('bull_hint', 0)} 銘柄")

with col3:
    st.metric("⚪ 横ばい", f"{signal_counts.get('sideways', 0)} 銘柄")

with col4:
    st.metric("🟠 下落の兆し", f"{signal_counts.get('bear_hint', 0)} 銘柄")

with col5:
    bear_count = signal_counts.get("strong_bear", 0) + signal_counts.get("bear", 0)
    st.metric("🔴⚫ 下落相場", f"{bear_count} 銘柄")

st.divider()

# === シグナル一覧テーブル ===
st.subheader("📋 シグナル一覧")

# シグナル優先度順にソート
signal_order = ["strong_bull", "bull", "bull_hint", "bear_hint", "bear", "strong_bear", "sideways"]
df_display = df_display.copy()
df_display["_sort"] = df_display["シグナル"].map({s: i for i, s in enumerate(signal_order)})
df_display = df_display.sort_values("_sort").drop(columns=["_sort"])

# 表示用カラム
display_cols = ["銘柄", "シグナル表示", "現在値", "トレンド強度(%)", "24h変動率(%)", "判定理由"]
st.dataframe(
    df_display[display_cols],
    use_container_width=True,
    hide_index=True,
    height=600,
    column_config={
        "銘柄": st.column_config.TextColumn("銘柄", width="small"),
        "シグナル表示": st.column_config.TextColumn("シグナル", width="medium"),
        "現在値": st.column_config.NumberColumn("現在値", format="%.6f"),
        "トレンド強度(%)": st.column_config.NumberColumn("強度(%)", format="%.2f"),
        "24h変動率(%)": st.column_config.NumberColumn("24h変動(%)", format="%.2f"),
        "判定理由": st.column_config.TextColumn("判定理由", width="large"),
    },
)

# === 個別チャート ===
st.divider()
st.subheader("📈 個別チャート")

chart_symbol = st.selectbox(
    "銘柄を選択",
    options=df_results["symbol_raw"].tolist(),
    format_func=lambda x: x.replace("USDT", "/USDT"),
)

if chart_symbol and chart_symbol in klines_data:
    df_chart = klines_data[chart_symbol].copy()
    df_chart = add_moving_averages(df_chart)

    fig = go.Figure()

    # ローソク足
    fig.add_trace(go.Candlestick(
        x=df_chart["timestamp"],
        open=df_chart["open"],
        high=df_chart["high"],
        low=df_chart["low"],
        close=df_chart["close"],
        name="ローソク足",
    ))

    # 移動平均線
    ma_colors = {5: "#FF6B6B", 10: "#4ECDC4", 30: "#45B7D1", 50: "#FFA07A", 100: "#9B59B6"}
    for period in MA_PERIODS:
        if f"MA{period}" in df_chart.columns:
            fig.add_trace(go.Scatter(
                x=df_chart["timestamp"],
                y=df_chart[f"MA{period}"],
                mode="lines",
                name=f"MA{period}",
                line=dict(color=ma_colors.get(period, "#888"), width=1.5),
            ))

    fig.update_layout(
        title=f"{chart_symbol.replace('USDT', '/USDT')} - {'日足' if timeframe == '1d' else '4時間足'}",
        xaxis_title="日時",
        yaxis_title="価格 (USDT)",
        template="plotly_dark",
        height=600,
        xaxis_rangeslider_visible=False,
    )

    st.plotly_chart(fig, use_container_width=True)

    # MA値の詳細表示
    signal_info = df_results[df_results["symbol_raw"] == chart_symbol].iloc[0]
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**移動平均線の値:**")
        for period in MA_PERIODS:
            st.markdown(f"- MA{period}: `{signal_info[f'MA{period}']:.6f}`")
    with col_b:
        st.markdown(f"**シグナル:** {signal_info['シグナル表示']}")
        st.markdown(f"**判定理由:** {signal_info['判定理由']}")
        st.markdown(f"**トレンド強度:** {signal_info['トレンド強度(%)']:.2f}%")

# === アラートログ ===
st.divider()
st.subheader("🔔 注目銘柄（シグナル変化あり）")
st.info(
    "ゴールデンクロス・デッドクロスが直近で発生した銘柄を表示します。\n"
    "ロング準備・ショート準備のシグナルが出ている銘柄に注目してください。"
)

alert_signals = ["bull_hint", "bear_hint"]
df_alerts = df_results[df_results["シグナル"].isin(alert_signals)]

if df_alerts.empty:
    st.write("現在、シグナル変化のある銘柄はありません。")
else:
    for _, row in df_alerts.iterrows():
        icon = "🟡" if row["シグナル"] == "bull_hint" else "🟠"
        st.markdown(
            f"{icon} **{row['銘柄']}** | "
            f"現在値: {row['現在値']:.6f} | "
            f"{row['判定理由']}"
        )

# === フッター ===
st.divider()
st.caption(
    "⚠️ このツールは投資助言ではありません。投資判断は自己責任でお願いします。\n"
    "📌 ロジックは暫定版です。相場師朗ロジックのPDF受領後に差し替えます。"
)
