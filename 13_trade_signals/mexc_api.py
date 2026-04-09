"""
MEXC Public API データ取得モジュール
APIキー不要。取引高上位銘柄のローソク足データを取得する。
"""

import time
import requests
import pandas as pd
from config import (
    MEXC_BASE_URL,
    MEXC_KLINES_ENDPOINT,
    MEXC_TICKER_ENDPOINT,
    TOP_N_SYMBOLS,
    QUOTE_ASSET,
    KLINE_LIMIT,
    USD_JPY_RATE,
)


def get_usdjpy_rate() -> float:
    """
    USD/JPYの最新レートを取得する。
    取得失敗時はconfig.pyのデフォルト値を返す。
    """
    try:
        resp = requests.get(
            "https://open.er-api.com/v6/latest/USD",
            timeout=5,
        )
        resp.raise_for_status()
        data = resp.json()
        rate = data["rates"]["JPY"]
        print(f"[為替] USD/JPY = {rate:.2f}（リアルタイム）")
        return float(rate)
    except Exception as e:
        print(f"[為替] レート取得失敗、デフォルト値 {USD_JPY_RATE}円 を使用: {e}")
        return USD_JPY_RATE


def get_market_caps(symbols: list[str]) -> dict[str, float]:
    """
    CoinGecko APIで時価総額（USD）を取得する。
    MEXCのシンボル（BTCUSDT）→ CoinGeckoのID変換が必要。
    取得失敗時は空辞書を返す。
    """
    # MEXCシンボルからコインシンボルを抽出（BTCUSDT→btc）
    coin_symbols = {}
    for s in symbols:
        coin = s.replace("USDT", "").lower()
        coin_symbols[coin] = s

    try:
        # CoinGecko: vs_currency=usdで上位250コインの時価総額を取得
        resp = requests.get(
            "https://api.coingecko.com/api/v3/coins/markets",
            params={
                "vs_currency": "usd",
                "order": "market_cap_desc",
                "per_page": 250,
                "page": 1,
                "sparkline": "false",
            },
            timeout=10,
        )
        resp.raise_for_status()
        coins = resp.json()

        result = {}
        for coin in coins:
            symbol_lower = coin.get("symbol", "").lower()
            if symbol_lower in coin_symbols:
                mexc_symbol = coin_symbols[symbol_lower]
                mcap = coin.get("market_cap", 0)
                if mcap:
                    result[mexc_symbol] = float(mcap)

        print(f"[時価総額] {len(result)}/{len(symbols)} 銘柄の時価総額を取得")
        return result
    except Exception as e:
        print(f"[時価総額] 取得失敗（スキップ）: {e}")
        return {}


def get_top_symbols(n: int = TOP_N_SYMBOLS) -> list[dict]:
    """
    取引高上位N銘柄のUSDTペアを取得する。

    Returns:
        [{"symbol": "BTCUSDT", "volume": 123456.78, "lastPrice": 68000.0}, ...]
    """
    url = f"{MEXC_BASE_URL}{MEXC_TICKER_ENDPOINT}"
    resp = requests.get(url, timeout=10)
    resp.raise_for_status()
    tickers = resp.json()

    # USDTペアのみフィルタ
    usdt_tickers = [
        t for t in tickers
        if t["symbol"].endswith(QUOTE_ASSET)
        and float(t.get("quoteVolume", 0)) > 0
    ]

    # 取引高（USDT建て）で降順ソート
    usdt_tickers.sort(key=lambda t: float(t.get("quoteVolume", 0)), reverse=True)

    top = usdt_tickers[:n]
    return [
        {
            "symbol": t["symbol"],
            "volume": float(t.get("quoteVolume", 0)),
            "lastPrice": float(t.get("lastPrice", 0)),
            "priceChangePercent": float(t.get("priceChangePercent", 0)),
        }
        for t in top
    ]


def get_klines(symbol: str, interval: str, limit: int = KLINE_LIMIT) -> pd.DataFrame:
    """
    指定銘柄のローソク足データを取得する。

    Args:
        symbol: 銘柄（例: "BTCUSDT"）
        interval: 時間足（例: "4h", "1d"）
        limit: 取得本数

    Returns:
        DataFrame（columns: open, high, low, close, volume, timestamp）
    """
    url = f"{MEXC_BASE_URL}{MEXC_KLINES_ENDPOINT}"
    params = {
        "symbol": symbol,
        "interval": interval,
        "limit": limit,
    }
    resp = requests.get(url, params=params, timeout=10)
    resp.raise_for_status()
    data = resp.json()

    if not data:
        return pd.DataFrame()

    # MEXCのklineは8カラム: [timestamp, open, high, low, close, volume, close_time, quote_volume]
    col_names = [
        "timestamp", "open", "high", "low", "close",
        "volume", "close_time", "quote_volume",
    ]
    # カラム数が異なる場合に対応
    actual_cols = len(data[0]) if data else 0
    if actual_cols < len(col_names):
        col_names = col_names[:actual_cols]
    elif actual_cols > len(col_names):
        col_names += [f"extra_{i}" for i in range(actual_cols - len(col_names))]

    df = pd.DataFrame(data, columns=col_names)

    # 数値型に変換
    for col in ["open", "high", "low", "close", "volume"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # タイムスタンプを日時に変換
    df["timestamp"] = pd.to_datetime(df["timestamp"], unit="ms")
    df = df[["timestamp", "open", "high", "low", "close", "volume"]]
    df = df.sort_values("timestamp").reset_index(drop=True)

    return df


def get_klines_batch(symbols: list[str], interval: str, limit: int = KLINE_LIMIT) -> dict[str, pd.DataFrame]:
    """
    複数銘柄のローソク足データを一括取得する。
    API制限を考慮して少し間隔を空ける。

    Returns:
        {"BTCUSDT": DataFrame, "ETHUSDT": DataFrame, ...}
    """
    total = len(symbols)
    result = {}
    for i, symbol in enumerate(symbols):
        try:
            df = get_klines(symbol, interval, limit)
            if not df.empty:
                result[symbol] = df
            print(f"[{interval}] {i+1}/{total} {symbol} OK")
        except Exception as e:
            print(f"[{interval}] {i+1}/{total} {symbol} SKIP: {e}")

        # API制限対策: 10リクエストごとに0.2秒待機
        if (i + 1) % 10 == 0:
            time.sleep(0.2)

    print(f"[{interval}] 完了: {len(result)}/{total} 銘柄取得")
    return result
