"""
移動平均線シグナル分析モジュール

【暫定ロジック】
正式版ロジックのPDF受領後に差し替え予定。
現在はパーフェクトオーダー＋ゴールデンクロス/デッドクロスで判定。
"""

import pandas as pd
from config import MA_PERIODS


def add_moving_averages(df: pd.DataFrame) -> pd.DataFrame:
    """ローソク足DataFrameに移動平均線を追加する。"""
    for period in MA_PERIODS:
        df[f"MA{period}"] = df["close"].rolling(window=period).mean()
    return df


def detect_signal(df: pd.DataFrame) -> dict:
    """
    最新のローソク足＋移動平均線からシグナルを判定する。

    Returns:
        {
            "signal": "strong_bull" | "bull" | "bull_hint" | "bear_hint" | "bear" | "strong_bear" | "sideways",
            "ma_values": {5: x, 10: x, ...},
            "close": 現在値,
            "detail": "判定理由の説明",
        }
    """
    if len(df) < max(MA_PERIODS):
        return {
            "signal": "sideways",
            "ma_values": {},
            "close": df["close"].iloc[-1] if len(df) > 0 else 0,
            "detail": "データ不足（様子見）",
        }

    latest = df.iloc[-1]
    prev = df.iloc[-2]

    close = latest["close"]
    ma_values = {p: latest[f"MA{p}"] for p in MA_PERIODS}

    ma5 = ma_values[5]
    ma10 = ma_values[10]
    ma30 = ma_values[30]
    ma50 = ma_values[50]
    ma100 = ma_values[100]

    prev_ma5 = prev["MA5"]
    prev_ma10 = prev["MA10"]

    # === シグナル判定ロジック（暫定版） ===

    # 1. 強い上昇相場: 完全パーフェクトオーダー（5>10>30>50>100）かつ価格がMA5より上
    if ma5 > ma10 > ma30 > ma50 > ma100 and close > ma5:
        return {
            "signal": "strong_bull",
            "ma_values": ma_values,
            "close": close,
            "detail": "パーフェクトオーダー成立（MA5>10>30>50>100）＋価格がMA5の上",
        }

    # 2. 上昇相場: 短期線が中長期線の上（5>10>30）
    if ma5 > ma10 > ma30:
        return {
            "signal": "bull",
            "ma_values": ma_values,
            "close": close,
            "detail": "上昇トレンド（MA5>10>30）",
        }

    # 3. 上昇の兆し: MA5がMA10をゴールデンクロス
    if prev_ma5 <= prev_ma10 and ma5 > ma10:
        return {
            "signal": "bull_hint",
            "ma_values": ma_values,
            "close": close,
            "detail": "MA5がMA10をゴールデンクロス（上昇転換の兆し）",
        }

    # 4. 強い下落相場: 逆パーフェクトオーダー（5<10<30<50<100）かつ価格がMA5より下
    if ma5 < ma10 < ma30 < ma50 < ma100 and close < ma5:
        return {
            "signal": "strong_bear",
            "ma_values": ma_values,
            "close": close,
            "detail": "逆パーフェクトオーダー成立（MA5<10<30<50<100）＋価格がMA5の下",
        }

    # 5. 下落相場: 短期線が中長期線の下（5<10<30）
    if ma5 < ma10 < ma30:
        return {
            "signal": "bear",
            "ma_values": ma_values,
            "close": close,
            "detail": "下落トレンド（MA5<10<30）",
        }

    # 6. 下落の兆し: MA5がMA10をデッドクロス
    if prev_ma5 >= prev_ma10 and ma5 < ma10:
        return {
            "signal": "bear_hint",
            "ma_values": ma_values,
            "close": close,
            "detail": "MA5がMA10をデッドクロス（下落転換の兆し）",
        }

    # 7. 横ばい: 上記いずれにも該当しない
    return {
        "signal": "sideways",
        "ma_values": ma_values,
        "close": close,
        "detail": "MA線が交錯中（方向感なし）",
    }


def analyze_symbol(df: pd.DataFrame) -> dict:
    """
    1銘柄分の分析を実行する。
    MA追加 → シグナル判定 → トレンド強度を返す。
    """
    df = add_moving_averages(df)
    signal = detect_signal(df)

    # トレンド強度: MA5とMA30の乖離率（%）
    if signal["ma_values"]:
        ma5 = signal["ma_values"].get(5, 0)
        ma30 = signal["ma_values"].get(30, 0)
        if ma30 and ma30 != 0:
            signal["trend_strength"] = round((ma5 - ma30) / ma30 * 100, 2)
        else:
            signal["trend_strength"] = 0.0
    else:
        signal["trend_strength"] = 0.0

    return signal


def check_exit_signal(df: pd.DataFrame, position: str) -> dict | None:
    """
    利確・撤退シグナルを判定する。（暫定版）

    Args:
        df: MA付きローソク足データ
        position: "long" or "short"

    Returns:
        利確シグナルがあれば辞書、なければNone
    """
    if len(df) < max(MA_PERIODS):
        return None

    df = add_moving_averages(df)
    latest = df.iloc[-1]
    prev = df.iloc[-2]

    ma5 = latest["MA5"]
    ma10 = latest["MA10"]
    prev_ma5 = prev["MA5"]
    prev_ma10 = prev["MA10"]

    if position == "long":
        # ロング利確: MA5がMA10をデッドクロス
        if prev_ma5 >= prev_ma10 and ma5 < ma10:
            return {
                "action": "利確推奨",
                "reason": "MA5がMA10をデッドクロス（上昇トレンド終了の兆し）",
            }

    elif position == "short":
        # ショート利確: MA5がMA10をゴールデンクロス
        if prev_ma5 <= prev_ma10 and ma5 > ma10:
            return {
                "action": "利確推奨",
                "reason": "MA5がMA10をゴールデンクロス（下落トレンド終了の兆し）",
            }

    return None
