"""
相場流 移動平均線シグナル分析モジュール
========================================

【5つの基本道具】
1. 下半身（逆下半身）: トレンド転換の初動シグナル
2. 本数: 上昇・下落は15本前後で転換（継続・終局の目安）
3. しこり: 前回高値・安値の抵抗/サポート
4. PPP（逆PPP）: 移動平均線の並びでトレンド判断
5. ものわかれ: 5日線が10日線に近づくが離れる→トレンド継続

【クチバシ（相場流のGC/DC）】
一般的なGC/DCより条件を厳しくしたもの。
- 基本クチバシ: 5日が10日を交差 + 5日と10日の向きが同じ
- 強いクチバシ: 上記 + 20日線も同じ向き（さらに信頼性が高い）

【3段ロジック】風向き → 合図 → 加速
- 風向き: 100日線との位置 + 20日線の傾き（大局）
- 合図:   下半身/逆下半身、クチバシ
- 加速:   100日線突破、PPP成立

移動平均線（相場流）:
  MA5  = 赤（直近の勢い）
  MA10 = 緑（少し落ち着いた勢い）
  MA20 = 青（風向き / 約1ヶ月）
  MA50 = 黄（中期）
  MA100= オレンジ（大局 / 20週・5ヶ月）
"""

import pandas as pd
from config import MA_PERIODS


def add_moving_averages(df: pd.DataFrame) -> pd.DataFrame:
    """ローソク足DataFrameに移動平均線を追加する。"""
    for period in MA_PERIODS:
        df[f"MA{period}"] = df["close"].rolling(window=period).mean()
    return df


# === 5つの基本道具の判定ヘルパー ===

def is_ppp(row) -> bool:
    """PPP（パーフェクトオーダー上昇）: 5 > 10 > 20 > 50 > 100"""
    return row["MA5"] > row["MA10"] > row["MA20"] > row["MA50"] > row["MA100"]


def is_reverse_ppp(row) -> bool:
    """逆PPP（下降）: 5 < 10 < 20 < 50 < 100"""
    return row["MA5"] < row["MA10"] < row["MA20"] < row["MA50"] < row["MA100"]


def detect_lower_half_body(df: pd.DataFrame) -> bool:
    """
    下半身シグナル（買い）
    条件: 5日線が横ばい or 上向き、かつ
          ローソク足が陽線で実体の半分以上が5日線の上
    """
    if len(df) < 3:
        return False
    latest = df.iloc[-1]
    prev = df.iloc[-2]

    ma5 = latest["MA5"]
    prev_ma5 = prev["MA5"]

    # 5日線が横ばい or 上向き
    if ma5 < prev_ma5:
        return False

    # 陽線（close > open）
    if latest["close"] <= latest["open"]:
        return False

    # 実体の半分以上が5日線の上
    body_mid = (latest["open"] + latest["close"]) / 2
    return body_mid > ma5 and latest["close"] > ma5


def detect_reverse_lower_half_body(df: pd.DataFrame) -> bool:
    """
    逆下半身シグナル（売り）
    条件: 5日線が横ばい or 下向き、かつ
          ローソク足が陰線で実体の半分以上が5日線の下
    """
    if len(df) < 3:
        return False
    latest = df.iloc[-1]
    prev = df.iloc[-2]

    ma5 = latest["MA5"]
    prev_ma5 = prev["MA5"]

    # 5日線が横ばい or 下向き
    if ma5 > prev_ma5:
        return False

    # 陰線（close < open）
    if latest["close"] >= latest["open"]:
        return False

    # 実体の半分以上が5日線の下
    body_mid = (latest["open"] + latest["close"]) / 2
    return body_mid < ma5 and latest["close"] < ma5


def detect_monowakare(df: pd.DataFrame, direction: str = "up") -> bool:
    """
    ものわかれシグナル（トレンド継続）
    上昇: 5日線が10日線に近づいた後、再び離れて上昇
    下降: 5日線が10日線に近づいた後、再び離れて下降
    直近5本で5日線と10日線の距離が一度縮まって広がっているかを確認
    """
    if len(df) < 6:
        return False

    recent = df.tail(6)
    distances = (recent["MA5"] - recent["MA10"]).abs().tolist()

    if direction == "up":
        # 上昇トレンド中に5日>10日が継続
        if not all(recent["MA5"].iloc[i] > recent["MA10"].iloc[i] for i in range(len(recent))):
            return False
    else:
        # 下降トレンド中に5日<10日が継続
        if not all(recent["MA5"].iloc[i] < recent["MA10"].iloc[i] for i in range(len(recent))):
            return False

    # 距離が一度縮まって（min）、再び広がっている
    min_idx = distances.index(min(distances))
    if min_idx == 0 or min_idx == len(distances) - 1:
        return False  # 最初 or 最後が最小なら「接近→離れる」ではない

    # 最後の距離 > 最小距離（離れている）
    return distances[-1] > distances[min_idx] * 1.1


def detect_golden_cross(df: pd.DataFrame, fast: int = 5, slow: int = 10) -> bool:
    """MA5がMA10をゴールデンクロス（単純な交差のみ）"""
    if len(df) < 2:
        return False
    latest = df.iloc[-1]
    prev = df.iloc[-2]
    return prev[f"MA{fast}"] <= prev[f"MA{slow}"] and latest[f"MA{fast}"] > latest[f"MA{slow}"]


def detect_dead_cross(df: pd.DataFrame, fast: int = 5, slow: int = 10) -> bool:
    """MA5がMA10をデッドクロス（単純な交差のみ）"""
    if len(df) < 2:
        return False
    latest = df.iloc[-1]
    prev = df.iloc[-2]
    return prev[f"MA{fast}"] >= prev[f"MA{slow}"] and latest[f"MA{fast}"] < latest[f"MA{slow}"]


def detect_kuchibashi(df: pd.DataFrame, direction: str = "up") -> tuple[bool, bool]:
    """
    クチバシシグナル（相場流のGC/DCより厳格な判定）

    基本クチバシ条件:
    1. 5日線が10日線を交差（上抜け=上、下抜け=下）
    2. 5日線と10日線の向きが同じ方向

    強いクチバシ（信頼度高）条件:
    - 基本クチバシ + 20日線の向きも同じ方向
      （20日線を抜くのではなく、向きが揃うこと）

    Args:
        df: ローソク足データ
        direction: "up" または "down"

    Returns:
        (基本クチバシ成立, 強いクチバシ成立)
    """
    if len(df) < 3:
        return (False, False)

    latest = df.iloc[-1]
    prev = df.iloc[-2]

    ma5_slope = latest["MA5"] - prev["MA5"]
    ma10_slope = latest["MA10"] - prev["MA10"]
    ma20_slope = latest["MA20"] - prev["MA20"]

    if direction == "up":
        # 5日が10日を上抜けているか（直近で交差 or 既に上）
        crossed = prev["MA5"] <= prev["MA10"] and latest["MA5"] > latest["MA10"]
        # 5日と10日の向きが同じ（上向き）
        same_direction = ma5_slope > 0 and ma10_slope > 0
        basic = crossed and same_direction
        strong = basic and ma20_slope > 0
    else:
        crossed = prev["MA5"] >= prev["MA10"] and latest["MA5"] < latest["MA10"]
        same_direction = ma5_slope < 0 and ma10_slope < 0
        basic = crossed and same_direction
        strong = basic and ma20_slope < 0

    return (basic, strong)


def check_wind_direction(latest) -> str:
    """
    風向き判定（大局）
    100日線との位置 + 20日線の傾き
    Returns: "up" | "down" | "neutral"
    """
    close = latest["close"]
    ma20 = latest["MA20"]
    ma100 = latest["MA100"]

    # 価格が100日線の上 → 上昇地合い
    if close > ma100 and ma20 > ma100:
        return "up"
    # 価格が100日線の下 → 下降地合い
    if close < ma100 and ma20 < ma100:
        return "down"
    return "neutral"


def count_trend_bars(df: pd.DataFrame, direction: str = "up") -> int:
    """
    トレンドの本数を数える（15本前後で転換の目安）
    MA5の傾きが同じ方向に連続している本数
    """
    if len(df) < 3:
        return 0

    count = 0
    ma5_series = df["MA5"].tolist()

    for i in range(len(ma5_series) - 1, 0, -1):
        diff = ma5_series[i] - ma5_series[i - 1]
        if direction == "up" and diff > 0:
            count += 1
        elif direction == "down" and diff < 0:
            count += 1
        else:
            break

    return count


# === メインのシグナル判定 ===

def detect_signal(df: pd.DataFrame) -> dict:
    """
    相場流 3段ロジックで最新のシグナルを判定する。
    風向き → 合図 → 加速 の順に評価。

    Returns:
        {
            "signal": "strong_bull" | "bull" | "bull_hint" | "bear_hint" | "bear" | "strong_bear" | "sideways",
            "ma_values": {5: x, 10: x, ...},
            "close": 現在値,
            "detail": "判定理由（相場流用語）",
            "flags": {
                "lower_half_body": bool,        # 下半身
                "reverse_lower_half_body": bool, # 逆下半身
                "kuchibashi_basic": bool,       # 基本クチバシ（上）
                "kuchibashi_strong": bool,      # 強いクチバシ（上）
                "rev_kuchibashi_basic": bool,   # 基本逆クチバシ
                "rev_kuchibashi_strong": bool,  # 強い逆クチバシ
                "ppp": bool,                    # PPP成立
                "reverse_ppp": bool,            # 逆PPP成立
                "monowakare_up": bool,          # ものわかれ（上）
                "monowakare_down": bool,        # ものわかれ（下）
                "above_ma100": bool,            # 100日線突破中
                "below_ma100": bool,            # 100日線割れ中
                "bars_warn": bool,              # 15本超え
                "bars_count": int,              # 連続本数
            }
        }
    """
    if len(df) < max(MA_PERIODS):
        return {
            "signal": "sideways",
            "ma_values": {},
            "close": df["close"].iloc[-1] if len(df) > 0 else 0,
            "detail": "データ不足（様子見）",
            "flags": {},
        }

    latest = df.iloc[-1]
    close = latest["close"]
    ma_values = {p: latest[f"MA{p}"] for p in MA_PERIODS}

    # 風向き判定
    wind = check_wind_direction(latest)

    # 基本道具の判定
    ppp = is_ppp(latest)
    rppp = is_reverse_ppp(latest)
    lhb = detect_lower_half_body(df)
    rlhb = detect_reverse_lower_half_body(df)
    kuchibashi_up_basic, kuchibashi_up_strong = detect_kuchibashi(df, "up")
    kuchibashi_dn_basic, kuchibashi_dn_strong = detect_kuchibashi(df, "down")
    mono_up = detect_monowakare(df, "up")
    mono_down = detect_monowakare(df, "down")

    bars_up = count_trend_bars(df, "up")
    bars_down = count_trend_bars(df, "down")
    bars_count = max(bars_up, bars_down)

    # フラグ辞書（バッジ表示用）。numpy.boolをPython boolに変換
    flags = {
        "lower_half_body": bool(lhb),
        "reverse_lower_half_body": bool(rlhb),
        "kuchibashi_basic": bool(kuchibashi_up_basic),
        "kuchibashi_strong": bool(kuchibashi_up_strong),
        "rev_kuchibashi_basic": bool(kuchibashi_dn_basic),
        "rev_kuchibashi_strong": bool(kuchibashi_dn_strong),
        "ppp": bool(ppp),
        "reverse_ppp": bool(rppp),
        "monowakare_up": bool(mono_up),
        "monowakare_down": bool(mono_down),
        "above_ma100": bool(close > latest["MA100"]),
        "below_ma100": bool(close < latest["MA100"]),
        "bars_warn": bool(bars_count >= 15),
        "bars_count": int(bars_count),
    }

    # === 強い上昇: PPP成立 + 風向き上 + 加速中（価格>MA5） ===
    if ppp and wind == "up" and close > latest["MA5"]:
        reason = "PPP成立（5>10>20>50>100）+ 100日線上で加速中"
        if mono_up:
            reason += " + ものわかれ継続"
        if bars_up >= 10:
            reason += f" + {bars_up}本継続"
            if bars_up >= 15:
                reason += "（15本超え、利確も意識）"
        return {
            "signal": "strong_bull",
            "ma_values": ma_values,
            "close": close,
            "detail": reason,
            "flags": flags,
        }

    # === 強い下落: 逆PPP + 風向き下 + 加速中 ===
    if rppp and wind == "down" and close < latest["MA5"]:
        reason = "逆PPP成立（5<10<20<50<100）+ 100日線下で加速中"
        if mono_down:
            reason += " + ものわかれ継続"
        if bars_down >= 10:
            reason += f" + {bars_down}本継続"
            if bars_down >= 15:
                reason += "（15本超え、戻し警戒）"
        return {
            "signal": "strong_bear",
            "ma_values": ma_values,
            "close": close,
            "detail": reason,
            "flags": flags,
        }

    # === 上昇相場: MA5>MA10>MA20 かつ 風向き上以外でない ===
    if latest["MA5"] > latest["MA10"] > latest["MA20"] and wind != "down":
        reason = "上昇トレンド（5>10>20）"
        if close > latest["MA100"]:
            reason += " / 100日線上"
        if mono_up:
            reason += " + ものわかれ継続"
        return {
            "signal": "bull",
            "ma_values": ma_values,
            "close": close,
            "detail": reason,
            "flags": flags,
        }

    # === 下落相場: MA5<MA10<MA20 かつ 風向き下以外でない ===
    if latest["MA5"] < latest["MA10"] < latest["MA20"] and wind != "up":
        reason = "下落トレンド（5<10<20）"
        if close < latest["MA100"]:
            reason += " / 100日線下"
        if mono_down:
            reason += " + ものわかれ継続"
        return {
            "signal": "bear",
            "ma_values": ma_values,
            "close": close,
            "detail": reason,
            "flags": flags,
        }

    # === 上昇の兆し: 下半身 or クチバシ ===
    if lhb:
        reason = "下半身出現（陽線が5日線を実体半分以上で上抜け）"
        if kuchibashi_up_strong:
            reason += " + 強いクチバシ（5/10/20日すべて上向き）"
        elif kuchibashi_up_basic:
            reason += " + クチバシ（5日が10日を上抜け、同方向）"
        if wind == "up":
            reason += " / 風向き上"
        return {
            "signal": "bull_hint",
            "ma_values": ma_values,
            "close": close,
            "detail": reason,
            "flags": flags,
        }

    if kuchibashi_up_strong:
        return {
            "signal": "bull_hint",
            "ma_values": ma_values,
            "close": close,
            "detail": "強いクチバシ出現（5日が10日を上抜け+5/10/20日すべて上向き）",
            "flags": flags,
        }

    if kuchibashi_up_basic:
        return {
            "signal": "bull_hint",
            "ma_values": ma_values,
            "close": close,
            "detail": "クチバシ出現（5日が10日を上抜け+5日と10日が同方向）",
            "flags": flags,
        }

    # === 下落の兆し: 逆下半身 or 逆クチバシ ===
    if rlhb:
        reason = "逆下半身出現（陰線が5日線を実体半分以上で下抜け）"
        if kuchibashi_dn_strong:
            reason += " + 強い逆クチバシ（5/10/20日すべて下向き）"
        elif kuchibashi_dn_basic:
            reason += " + 逆クチバシ（5日が10日を下抜け、同方向）"
        if wind == "down":
            reason += " / 風向き下"
        return {
            "signal": "bear_hint",
            "ma_values": ma_values,
            "close": close,
            "detail": reason,
            "flags": flags,
        }

    if kuchibashi_dn_strong:
        return {
            "signal": "bear_hint",
            "ma_values": ma_values,
            "close": close,
            "detail": "強い逆クチバシ出現（5日が10日を下抜け+5/10/20日すべて下向き）",
            "flags": flags,
        }

    if kuchibashi_dn_basic:
        return {
            "signal": "bear_hint",
            "ma_values": ma_values,
            "close": close,
            "detail": "逆クチバシ出現（5日が10日を下抜け+5日と10日が同方向）",
            "flags": flags,
        }

    # === 横ばい ===
    return {
        "signal": "sideways",
        "ma_values": ma_values,
        "close": close,
        "detail": "MA線が収束・交錯中（風向き待ち）",
        "flags": flags,
    }


def analyze_symbol(df: pd.DataFrame) -> dict:
    """
    1銘柄分の分析を実行する。
    MA追加 → シグナル判定 → トレンド強度を返す。
    """
    df = add_moving_averages(df)
    signal = detect_signal(df)

    # トレンド強度: MA5とMA20の乖離率（%）
    if signal["ma_values"]:
        ma5 = signal["ma_values"].get(5, 0)
        ma20 = signal["ma_values"].get(20, 0)
        if ma20 and ma20 != 0:
            signal["trend_strength"] = round((ma5 - ma20) / ma20 * 100, 2)
        else:
            signal["trend_strength"] = 0.0
    else:
        signal["trend_strength"] = 0.0

    return signal


def check_exit_signal(df: pd.DataFrame, position: str) -> dict | None:
    """
    利確・撤退シグナルを判定する。

    相場流の利確ロジック:
    - ロング: 逆下半身 or MA5がMA10をDC、または本数が15本超え
    - ショート: 下半身 or MA5がMA10をGC、または本数が15本超え
    """
    if len(df) < max(MA_PERIODS):
        return None

    df = add_moving_averages(df)

    if position == "long":
        if detect_reverse_lower_half_body(df):
            return {"action": "利確推奨", "reason": "逆下半身出現（上昇トレンド終了の初動）"}
        basic, strong = detect_kuchibashi(df, "down")
        if strong:
            return {"action": "利確推奨", "reason": "強い逆クチバシ出現"}
        if basic:
            return {"action": "利確推奨", "reason": "逆クチバシ出現"}
        bars = count_trend_bars(df, "up")
        if bars >= 15:
            return {"action": "利確警戒", "reason": f"上昇が{bars}本継続（15本ルール）"}

    elif position == "short":
        if detect_lower_half_body(df):
            return {"action": "利確推奨", "reason": "下半身出現（下落トレンド終了の初動）"}
        basic, strong = detect_kuchibashi(df, "up")
        if strong:
            return {"action": "利確推奨", "reason": "強いクチバシ出現"}
        if basic:
            return {"action": "利確推奨", "reason": "クチバシ出現"}
        bars = count_trend_bars(df, "down")
        if bars >= 15:
            return {"action": "利確警戒", "reason": f"下落が{bars}本継続（15本ルール）"}

    return None
