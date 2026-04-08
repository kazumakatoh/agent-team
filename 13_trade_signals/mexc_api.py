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
)


def get_top_symbols(n: int = TOP_N_SYMBOLS) -> list[dict]:
    """
    取引高上位N銘柄のUSDTペアを取得する。

    Returns:
        [{"symbol": "BTCUSDT", "volume": 123456.78, "lastPrice": 68000.0}, ...]
    """
    url = f"{MEXC_BASE_URL}{MEXC_TICKER_ENDPOINT}"
    resp = requests.get(url, timeout=30)
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
    resp = requests.get(url, params=params, timeout=30)
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
    result = {}
    for i, symbol in enumerate(symbols):
        try:
            df = get_klines(symbol, interval, limit)
            if not df.empty:
                result[symbol] = df
        except Exception as e:
            print(f"[WARN] {symbol} のデータ取得に失敗: {e}")

        # API制限対策: 10リクエストごとに0.3秒待機
        if (i + 1) % 10 == 0:
            time.sleep(0.3)

    return result
