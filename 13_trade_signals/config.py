"""
トレードシグナルシステム - 設定ファイル
"""

# === MEXC API ===
MEXC_BASE_URL = "https://api.mexc.com"
MEXC_KLINES_ENDPOINT = "/api/v3/klines"
MEXC_TICKER_ENDPOINT = "/api/v3/ticker/24hr"

# === 分析対象 ===
TOP_N_SYMBOLS = 50  # 取引高上位N銘柄
QUOTE_ASSET = "USDT"  # 対象ペア

# === 時間足 ===
TIMEFRAMES = {
    "4h": "4h",    # 4時間足
    "1d": "1d",    # 日足
}

# === 移動平均線 ===
MA_PERIODS = [5, 10, 30, 50, 100]

# === ローソク足取得本数 ===
KLINE_LIMIT = 150  # MA100を計算するために十分な本数

# === 自動更新間隔（秒）===
REFRESH_INTERVAL = 600  # 10分

# === シグナル定義 ===
SIGNAL_LABELS = {
    "strong_bull": "🟢 強い上昇相場（ロング推奨）",
    "bull": "🔵 上昇相場（ロング推奨）",
    "bull_hint": "🟡 上昇の兆し（ロング準備）",
    "bear_hint": "🟠 下落の兆し（ショート準備）",
    "bear": "🔴 下落相場（ショート推奨）",
    "strong_bear": "⚫ 強い下落相場（ショート推奨）",
    "sideways": "⚪ 横ばい（様子見）",
}

# === 為替レート ===
USD_JPY_RATE = 150  # 1ドル = 150円（概算）

# === LINE通知（Phase 4で実装） ===
LINE_NOTIFY_TOKEN = ""  # LINE Notify トークンをここに設定
