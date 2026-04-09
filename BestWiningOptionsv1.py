import io
import math
from datetime import datetime

import pandas as pd
import yfinance as yf


# =========================================================
# USER SETTINGS
# =========================================================
TICKERS = ["AAPL", "MSFT", "NVDA", "AMD", "TSLA", "META", "AMZN", "SPY", "QQQ"]

LOOKBACK_DAYS = 120
RESISTANCE_LOOKBACK = 20
SUPPORT_LOOKBACK = 20
ATR_PERIOD = 14
AFTER_HOURS_MOVE_THRESHOLD_PCT = 1.0
EARNINGS_SOON_DAYS = 7

TOP_STOCKS_TO_SCAN_OPTIONS = 5
MAX_OPTION_EXPIRIES = 6

MIN_VOLUME = 1000
MIN_OPEN_INTEREST = 3000
MAX_BID_ASK_SPREAD_PCT = 5.0
MIN_OI_VOLUME_RATIO = 0.5
MAX_OI_VOLUME_RATIO = 3.0

CALL_DELTA_MIN = 0.45
CALL_DELTA_MAX = 0.60
PUT_DELTA_MIN = -0.60
PUT_DELTA_MAX = -0.45

RISK_FREE_RATE = 0.05
OUTPUT_FILE = "next_day_options_upgraded.xlsx"

# Contract selection controls
PREFER_SAME_WEEK_EXPIRY = True
MAX_STRIKE_DISTANCE_FROM_EXPECTED_MOVE_MULT = 1.00
MAX_STRIKE_DISTANCE_FROM_ATR_MULT = 0.75

# Risk management defaults
CONSERVATIVE_STOP_LOSS_PCT = 25
CONSERVATIVE_TAKE_PROFIT_PCT = 40

AGGRESSIVE_STOP_LOSS_PCT = 35
AGGRESSIVE_TAKE_PROFIT_PCT = 70

WEIGHTS = {
    "close_strength": 20,
    "breakout_breakdown": 20,
    "after_hours": 15,
    "earnings_catalyst": 10,
    "sector_momentum": 20,
    "trend_confirmation": 15,
}


# =========================================================
# BASIC HELPERS
# =========================================================
def safe_float(value, default=0.0):
    try:
        if value is None or pd.isna(value):
            return default
        return float(value)
    except Exception:
        return default


def pct_change(new_value, old_value):
    new_value = safe_float(new_value, 0.0)
    old_value = safe_float(old_value, 0.0)
    if old_value == 0:
        return 0.0
    return ((new_value - old_value) / old_value) * 100.0


def clamp(value, low, high):
    return max(low, min(value, high))


def norm_cdf(x):
    return 0.5 * (1.0 + math.erf(x / math.sqrt(2.0)))


def get_history(symbol, period_days=LOOKBACK_DAYS):
    ticker = yf.Ticker(symbol)
    df = ticker.history(period=f"{period_days}d", interval="1d", auto_adjust=False)
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()
    try:
        df.index = pd.to_datetime(df.index).tz_localize(None)
    except Exception:
        df.index = pd.to_datetime(df.index)

    return df


def get_fast_info_field(ticker, field_name, default=None):
    try:
        fi = ticker.fast_info
        if fi is None:
            return default
        return fi.get(field_name, default)
    except Exception:
        return default


def trading_days_to_expiry(expiry_str):
    try:
        expiry_date = pd.to_datetime(expiry_str).date()
        today = datetime.now().date()
        days = (expiry_date - today).days
        return max(days, 0)
    except Exception:
        return None


def time_to_expiry_in_years(expiry_str):
    days = trading_days_to_expiry(expiry_str)
    if days is None:
        return None
    return max(days, 1) / 365.0


# =========================================================
# EARNINGS / AFTER HOURS
# =========================================================
def get_after_hours_snapshot(ticker):
    regular_close = get_fast_info_field(ticker, "regularMarketPreviousClose", None)
    last_price = get_fast_info_field(ticker, "lastPrice", None)

    if last_price is None:
        last_price = get_fast_info_field(ticker, "regularMarketPrice", None)

    return safe_float(last_price, 0.0), safe_float(regular_close, 0.0)


def get_earnings_date(ticker):
    try:
        cal = ticker.calendar
        if isinstance(cal, pd.DataFrame) and not cal.empty:
            for col in cal.columns:
                if "earn" in str(col).lower():
                    val = cal[col].iloc[0]
                    if pd.notna(val):
                        return pd.to_datetime(val).date()
        elif isinstance(cal, dict):
            for key, val in cal.items():
                if "earn" in str(key).lower() and val is not None:
                    if isinstance(val, (list, tuple)) and len(val) > 0:
                        return pd.to_datetime(val[0]).date()
                    return pd.to_datetime(val).date()
    except Exception:
        pass

    try:
        ed = ticker.earnings_dates
        if ed is not None and not ed.empty:
            idx = pd.to_datetime(ed.index)
            future_dates = [d.date() for d in idx if d.date() >= datetime.now().date()]
            if future_dates:
                return min(future_dates)
    except Exception:
        pass

    return None


def compute_after_hours_signal(ticker):
    last_price, regular_close = get_after_hours_snapshot(ticker)

    if regular_close <= 0 or last_price <= 0:
        return {
            "after_hours_price": None,
            "regular_close": None,
            "after_hours_change_pct": 0.0,
            "signal": "UNKNOWN",
            "score": 0.0,
        }

    ah_change = pct_change(last_price, regular_close)

    if ah_change >= AFTER_HOURS_MOVE_THRESHOLD_PCT:
        return {
            "after_hours_price": round(last_price, 2),
            "regular_close": round(regular_close, 2),
            "after_hours_change_pct": round(ah_change, 2),
            "signal": "BULLISH_AFTER_HOURS",
            "score": 1.0,
        }
    elif ah_change <= -AFTER_HOURS_MOVE_THRESHOLD_PCT:
        return {
            "after_hours_price": round(last_price, 2),
            "regular_close": round(regular_close, 2),
            "after_hours_change_pct": round(ah_change, 2),
            "signal": "BEARISH_AFTER_HOURS",
            "score": -1.0,
        }
    else:
        return {
            "after_hours_price": round(last_price, 2),
            "regular_close": round(regular_close, 2),
            "after_hours_change_pct": round(ah_change, 2),
            "signal": "NEUTRAL",
            "score": 0.0,
        }


def compute_earnings_signal(ticker):
    earnings_date = get_earnings_date(ticker)

    if earnings_date is None:
        return {
            "earnings_date": None,
            "days_to_earnings": None,
            "signal": "UNKNOWN",
            "score": 0.0,
        }

    today = datetime.now().date()
    dte = (earnings_date - today).days

    if 0 <= dte <= EARNINGS_SOON_DAYS:
        return {
            "earnings_date": earnings_date.isoformat(),
            "days_to_earnings": dte,
            "signal": "EARNINGS_SOON",
            "score": 1.0,
        }

    return {
        "earnings_date": earnings_date.isoformat(),
        "days_to_earnings": dte,
        "signal": "NO_NEAR_EARNINGS",
        "score": 0.0,
    }


# =========================================================
# MARKET CONTEXT
# =========================================================
def analyze_benchmark(symbol):
    df = get_history(symbol, 60)
    if df.empty or len(df) < 25:
        return {
            "symbol": symbol,
            "close": None,
            "daily_change_pct": 0.0,
            "above_20sma": False,
            "trend_score": 0.0,
            "direction": "NEUTRAL",
        }

    df = df.copy()
    df["SMA20"] = df["Close"].rolling(20).mean()

    last = df.iloc[-1]
    prev = df.iloc[-2]

    close_price = safe_float(last["Close"])
    prev_close = safe_float(prev["Close"])
    sma20 = safe_float(last["SMA20"])

    daily_change_pct = pct_change(close_price, prev_close)
    above_20sma = close_price > sma20 if sma20 else False

    trend_score = 0.0
    trend_score += 0.5 if daily_change_pct > 0 else -0.5 if daily_change_pct < 0 else 0.0
    trend_score += 0.5 if above_20sma else -0.5

    if trend_score > 0.25:
        direction = "BULLISH"
    elif trend_score < -0.25:
        direction = "BEARISH"
    else:
        direction = "NEUTRAL"

    return {
        "symbol": symbol,
        "close": close_price,
        "daily_change_pct": round(daily_change_pct, 2),
        "above_20sma": above_20sma,
        "trend_score": trend_score,
        "direction": direction,
    }


def get_sector_context():
    spy = analyze_benchmark("SPY")
    qqq = analyze_benchmark("QQQ")

    market_score = (spy["trend_score"] + qqq["trend_score"]) / 2.0

    if market_score > 0.25:
        market_direction = "BULLISH"
    elif market_score < -0.25:
        market_direction = "BEARISH"
    else:
        market_direction = "NEUTRAL"

    return {
        "SPY_direction": spy["direction"],
        "QQQ_direction": qqq["direction"],
        "SPY_daily_change_pct": spy["daily_change_pct"],
        "QQQ_daily_change_pct": qqq["daily_change_pct"],
        "market_direction": market_direction,
        "market_score": market_score,
    }


# =========================================================
# STOCK ANALYSIS
# =========================================================
def compute_atr(df, period=ATR_PERIOD):
    df = df.copy()
    df["prev_close"] = df["Close"].shift(1)
    df["tr1"] = df["High"] - df["Low"]
    df["tr2"] = (df["High"] - df["prev_close"]).abs()
    df["tr3"] = (df["Low"] - df["prev_close"]).abs()
    df["TR"] = df[["tr1", "tr2", "tr3"]].max(axis=1)
    df["ATR"] = df["TR"].rolling(period).mean()
    return df


def compute_close_strength(last_row):
    high_ = safe_float(last_row["High"])
    low_ = safe_float(last_row["Low"])
    close_ = safe_float(last_row["Close"])

    day_range = high_ - low_
    if day_range <= 0:
        return 0.5, "NEUTRAL", 0.0

    close_pos = (close_ - low_) / day_range

    if close_pos >= 0.8:
        return round(close_pos, 4), "BULLISH_CLOSE", 1.0
    elif close_pos <= 0.2:
        return round(close_pos, 4), "BEARISH_CLOSE", -1.0
    else:
        return round(close_pos, 4), "NEUTRAL", 0.0


def compute_breakout_breakdown(df):
    if len(df) < max(RESISTANCE_LOOKBACK, SUPPORT_LOOKBACK) + 2:
        return {
            "resistance": None,
            "support": None,
            "breakout": False,
            "breakdown": False,
            "signal": "NONE",
            "score": 0.0,
        }

    prev_window_high = df["High"].iloc[-(RESISTANCE_LOOKBACK + 1):-1].max()
    prev_window_low = df["Low"].iloc[-(SUPPORT_LOOKBACK + 1):-1].min()
    last_close = safe_float(df["Close"].iloc[-1])

    breakout = last_close > safe_float(prev_window_high)
    breakdown = last_close < safe_float(prev_window_low)

    if breakout and not breakdown:
        signal = "BREAKOUT"
        score = 1.0
    elif breakdown and not breakout:
        signal = "BREAKDOWN"
        score = -1.0
    else:
        signal = "NONE"
        score = 0.0

    return {
        "resistance": round(safe_float(prev_window_high), 2),
        "support": round(safe_float(prev_window_low), 2),
        "breakout": breakout,
        "breakdown": breakdown,
        "signal": signal,
        "score": score,
    }


def compute_trend_confirmation(df):
    if len(df) < 55:
        return {"SMA20": None, "SMA50": None, "signal": "NEUTRAL", "score": 0.0}

    df = df.copy()
    df["SMA20"] = df["Close"].rolling(20).mean()
    df["SMA50"] = df["Close"].rolling(50).mean()

    last = df.iloc[-1]
    prev = df.iloc[-2]

    close_ = safe_float(last["Close"])
    prev_close = safe_float(prev["Close"])
    sma20 = safe_float(last["SMA20"])
    sma50 = safe_float(last["SMA50"])
    daily_change = pct_change(close_, prev_close)

    score = 0.0
    score += 0.4 if close_ > sma20 else -0.4
    score += 0.4 if close_ > sma50 else -0.4
    score += 0.2 if daily_change > 0 else -0.2 if daily_change < 0 else 0.0

    if score > 0.3:
        signal = "BULLISH_TREND"
    elif score < -0.3:
        signal = "BEARISH_TREND"
    else:
        signal = "NEUTRAL"

    return {
        "SMA20": round(sma20, 2) if sma20 else None,
        "SMA50": round(sma50, 2) if sma50 else None,
        "signal": signal,
        "score": score,
    }


def compute_sector_alignment(stock_direction_score, market_score):
    if stock_direction_score > 0.25 and market_score > 0.25:
        return "BULLISH_ALIGNED", 1.0
    elif stock_direction_score < -0.25 and market_score < -0.25:
        return "BEARISH_ALIGNED", 1.0
    elif abs(stock_direction_score) < 0.25 or abs(market_score) < 0.25:
        return "NEUTRAL_ALIGNMENT", 0.0
    else:
        return "CONFLICT", -1.0


def analyze_ticker(symbol, market_context):
    ticker = yf.Ticker(symbol)
    df = get_history(symbol)

    if df.empty or len(df) < 55:
        return {"ticker": symbol, "error": "Not enough historical data"}

    df = compute_atr(df, ATR_PERIOD)

    last = df.iloc[-1]
    prev = df.iloc[-2]

    close_price = safe_float(last["Close"])
    prev_close = safe_float(prev["Close"])
    daily_change_pct = pct_change(close_price, prev_close)
    atr_value = safe_float(last.get("ATR"), 0.0)

    close_pos, close_signal, close_strength_score = compute_close_strength(last)
    breakout_info = compute_breakout_breakdown(df)
    trend_info = compute_trend_confirmation(df)
    ah_info = compute_after_hours_signal(ticker)
    earnings_info = compute_earnings_signal(ticker)

    raw_stock_direction_score = (
        0.35 * close_strength_score
        + 0.35 * breakout_info["score"]
        + 0.30 * trend_info["score"]
    )

    alignment_signal, alignment_score = compute_sector_alignment(
        raw_stock_direction_score,
        market_context["market_score"]
    )

    total_score = 0.0
    total_score += abs(close_strength_score) * WEIGHTS["close_strength"]
    total_score += abs(breakout_info["score"]) * WEIGHTS["breakout_breakdown"]
    total_score += abs(ah_info["score"]) * WEIGHTS["after_hours"]
    total_score += earnings_info["score"] * WEIGHTS["earnings_catalyst"]
    total_score += max(0.0, alignment_score) * WEIGHTS["sector_momentum"]
    total_score += abs(trend_info["score"]) * WEIGHTS["trend_confirmation"]
    total_score = round(clamp(total_score, 0, 100), 2)

    bullish_votes = 0
    bearish_votes = 0
    for score in [close_strength_score, breakout_info["score"], ah_info["score"], trend_info["score"]]:
        if score > 0:
            bullish_votes += 1
        elif score < 0:
            bearish_votes += 1

    if bullish_votes > bearish_votes:
        trade_bias = "CALL"
    elif bearish_votes > bullish_votes:
        trade_bias = "PUT"
    else:
        trade_bias = "NEUTRAL"

    setup_tags = []
    if close_signal != "NEUTRAL":
        setup_tags.append(close_signal)
    if breakout_info["signal"] != "NONE":
        setup_tags.append(breakout_info["signal"])
    if ah_info["signal"] not in ("NEUTRAL", "UNKNOWN"):
        setup_tags.append(ah_info["signal"])
    if earnings_info["signal"] == "EARNINGS_SOON":
        setup_tags.append("EARNINGS_SOON")
    if alignment_signal not in ("NEUTRAL_ALIGNMENT", "CONFLICT"):
        setup_tags.append(alignment_signal)
    if trend_info["signal"] != "NEUTRAL":
        setup_tags.append(trend_info["signal"])
    if not setup_tags:
        setup_tags.append("NO_STRONG_EDGE")

    return {
        "ticker": symbol,
        "close": round(close_price, 2),
        "atr": round(atr_value, 2) if atr_value else None,
        "daily_change_pct": round(daily_change_pct, 2),
        "close_position_in_day_range": close_pos,
        "close_signal": close_signal,
        "resistance": breakout_info["resistance"],
        "support": breakout_info["support"],
        "breakout_signal": breakout_info["signal"],
        "after_hours_change_pct": ah_info["after_hours_change_pct"],
        "after_hours_signal": ah_info["signal"],
        "earnings_date": earnings_info["earnings_date"],
        "days_to_earnings": earnings_info["days_to_earnings"],
        "earnings_signal": earnings_info["signal"],
        "trend_signal": trend_info["signal"],
        "market_direction": market_context["market_direction"],
        "alignment_signal": alignment_signal,
        "bullish_votes": bullish_votes,
        "bearish_votes": bearish_votes,
        "trade_bias": trade_bias,
        "setup_score": total_score,
        "setup_tags": ", ".join(setup_tags),
    }


# =========================================================
# OPTION PRICING HELPERS
# =========================================================
def black_scholes_delta(option_type, S, K, T, r, sigma):
    S = safe_float(S, 0.0)
    K = safe_float(K, 0.0)
    T = safe_float(T, 0.0)
    sigma = safe_float(sigma, 0.0)
    r = safe_float(r, 0.0)

    if S <= 0 or K <= 0 or T <= 0 or sigma <= 0:
        return None

    try:
        d1 = (math.log(S / K) + (r + 0.5 * sigma ** 2) * T) / (sigma * math.sqrt(T))
        if option_type == "CALL":
            return norm_cdf(d1)
        else:
            return norm_cdf(d1) - 1.0
    except Exception:
        return None


def calc_expected_move_price(underlying_price, iv_decimal, dte_days):
    underlying_price = safe_float(underlying_price, 0.0)
    iv_decimal = safe_float(iv_decimal, 0.0)
    dte_days = max(int(safe_float(dte_days, 1)), 1)

    if underlying_price <= 0 or iv_decimal <= 0:
        return None

    return underlying_price * iv_decimal * math.sqrt(dte_days / 252.0)


def label_contract_style(delta, spread_pct, oi_volume_ratio, dte_days, trade_bias):
    delta = safe_float(delta, 0.0)
    spread_pct = safe_float(spread_pct, 999.0)
    oi_volume_ratio = safe_float(oi_volume_ratio, 999.0)
    dte_days = int(safe_float(dte_days, 999))

    if trade_bias == "CALL":
        conservative_delta = 0.52 <= delta <= 0.60
        aggressive_delta = 0.45 <= delta < 0.52
    else:
        conservative_delta = -0.60 <= delta <= -0.52
        aggressive_delta = -0.52 < delta <= -0.45

    tight_liquidity = spread_pct <= 3.0 and 0.8 <= oi_volume_ratio <= 2.0
    same_week = dte_days <= 7

    if conservative_delta and tight_liquidity and same_week:
        return "Conservative"
    if aggressive_delta:
        return "Aggressive"
    return "Balanced"


def risk_levels(mid_price, label):
    mid_price = safe_float(mid_price, 0.0)
    if mid_price <= 0:
        return None, None

    if label == "Conservative":
        stop_loss = mid_price * (1 - CONSERVATIVE_STOP_LOSS_PCT / 100.0)
        take_profit = mid_price * (1 + CONSERVATIVE_TAKE_PROFIT_PCT / 100.0)
    elif label == "Aggressive":
        stop_loss = mid_price * (1 - AGGRESSIVE_STOP_LOSS_PCT / 100.0)
        take_profit = mid_price * (1 + AGGRESSIVE_TAKE_PROFIT_PCT / 100.0)
    else:
        stop_loss = mid_price * (1 - 0.30)
        take_profit = mid_price * (1 + 0.55)

    return round(stop_loss, 2), round(take_profit, 2)


# =========================================================
# OPTION CONTRACT SELECTION
# =========================================================
def get_best_contracts_for_ticker(ticker_symbol, trade_bias, stock_setup_score=None, atr_value=None):
    if trade_bias not in ("CALL", "PUT"):
        return pd.DataFrame()

    ticker = yf.Ticker(ticker_symbol)
    hist = get_history(ticker_symbol, 20)
    if hist.empty:
        return pd.DataFrame()

    underlying_price = safe_float(hist["Close"].iloc[-1], 0.0)
    if underlying_price <= 0:
        return pd.DataFrame()

    try:
        expiries = list(ticker.options)
    except Exception as e:
        print(f"[{ticker_symbol}] options fetch failed: {e}")
        return pd.DataFrame()

    if not expiries:
        return pd.DataFrame()

    expiries = expiries[:MAX_OPTION_EXPIRIES]
    all_contracts = []

    for expiry in expiries:
        try:
            chain = ticker.option_chain(expiry)
        except Exception:
            continue

        df = chain.calls.copy() if trade_bias == "CALL" else chain.puts.copy()
        if df.empty:
            continue

        numeric_cols = ["bid", "ask", "volume", "openInterest", "lastPrice", "strike", "impliedVolatility"]
        for col in numeric_cols:
            if col not in df.columns:
                df[col] = 0
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        df = df[(df["bid"] > 0) & (df["ask"] > 0) & (df["volume"] > 0)]
        if df.empty:
            continue

        dte_days = trading_days_to_expiry(expiry)
        T = time_to_expiry_in_years(expiry)
        if dte_days is None or T is None:
            continue

        df["spread_pct"] = ((df["ask"] - df["bid"]) / df["ask"]) * 100.0
        df["oi_volume_ratio"] = df["openInterest"] / df["volume"].replace(0, pd.NA)
        df["mid_price"] = (df["bid"] + df["ask"]) / 2.0
        df["expiry"] = expiry
        df["dte_days"] = dte_days
        df["ticker"] = ticker_symbol
        df["trade_bias"] = trade_bias
        df["underlying_price"] = underlying_price

        def calc_delta(row):
            iv = safe_float(row["impliedVolatility"], 0.0)
            strike = safe_float(row["strike"], 0.0)
            return black_scholes_delta(trade_bias, underlying_price, strike, T, RISK_FREE_RATE, iv)

        df["delta"] = df.apply(calc_delta, axis=1)
        df = df[df["delta"].notna()]
        if df.empty:
            continue

        # Main liquidity filters
        filtered = df[
            (df["volume"] > MIN_VOLUME) &
            (df["openInterest"] > MIN_OPEN_INTEREST) &
            (df["spread_pct"] < MAX_BID_ASK_SPREAD_PCT) &
            (df["oi_volume_ratio"] >= MIN_OI_VOLUME_RATIO) &
            (df["oi_volume_ratio"] <= MAX_OI_VOLUME_RATIO)
        ].copy()

        if filtered.empty:
            continue

        # Delta filters
        if trade_bias == "CALL":
            filtered = filtered[
                (filtered["delta"] >= CALL_DELTA_MIN) &
                (filtered["delta"] <= CALL_DELTA_MAX)
            ].copy()
        else:
            filtered = filtered[
                (filtered["delta"] >= PUT_DELTA_MIN) &
                (filtered["delta"] <= PUT_DELTA_MAX)
            ].copy()

        if filtered.empty:
            continue

        # Expected move filter + ATR strike distance filter
        filtered["expected_move"] = filtered.apply(
            lambda row: calc_expected_move_price(underlying_price, safe_float(row["impliedVolatility"], 0.0), dte_days),
            axis=1
        )

        if trade_bias == "CALL":
            filtered["strike_distance"] = filtered["strike"] - underlying_price
        else:
            filtered["strike_distance"] = underlying_price - filtered["strike"]

        filtered["strike_distance"] = filtered["strike_distance"].clip(lower=0)

        # Remove contracts too far away from underlying
        filtered = filtered[filtered["expected_move"].notna()].copy()
        if filtered.empty:
            continue

        filtered["max_allowed_by_em"] = filtered["expected_move"] * MAX_STRIKE_DISTANCE_FROM_EXPECTED_MOVE_MULT
        filtered = filtered[filtered["strike_distance"] <= filtered["max_allowed_by_em"]].copy()
        if filtered.empty:
            continue

        if atr_value and atr_value > 0:
            filtered["max_allowed_by_atr"] = atr_value * MAX_STRIKE_DISTANCE_FROM_ATR_MULT
            filtered = filtered[filtered["strike_distance"] <= filtered["max_allowed_by_atr"]].copy()
            if filtered.empty:
                continue
        else:
            filtered["max_allowed_by_atr"] = None

        # Same-week expiry preference
        filtered["same_week_expiry"] = filtered["dte_days"] <= 7
        if PREFER_SAME_WEEK_EXPIRY and filtered["same_week_expiry"].any():
            filtered = filtered[filtered["same_week_expiry"]].copy()
            if filtered.empty:
                continue

        # Delta center preference
        ideal_delta = 0.55 if trade_bias == "CALL" else -0.55
        filtered["delta_distance"] = (filtered["delta"] - ideal_delta).abs()

        # Contract style label
        filtered["label"] = filtered.apply(
            lambda row: label_contract_style(
                row["delta"], row["spread_pct"], row["oi_volume_ratio"], row["dte_days"], trade_bias
            ),
            axis=1
        )

        # Stop / target
        risk_df = filtered["mid_price"].apply(lambda x: risk_levels(x, "Balanced"))
        filtered["stop_loss"] = [x[0] if x else None for x in risk_df]
        filtered["take_profit"] = [x[1] if x else None for x in risk_df]

        # Recompute with actual labels
        labels_and_levels = filtered.apply(lambda row: risk_levels(row["mid_price"], row["label"]), axis=1)
        filtered["stop_loss"] = [x[0] if x else None for x in labels_and_levels]
        filtered["take_profit"] = [x[1] if x else None for x in labels_and_levels]

        # Entry notes
        def entry_note(row):
            if row["label"] == "Conservative":
                return f"{row['label']} {trade_bias.title()} - tighter liquidity, delta closer to 0.55"
            elif row["label"] == "Aggressive":
                return f"{row['label']} {trade_bias.title()} - lower premium, slightly more directional"
            return f"{row['label']} {trade_bias.title()} - balanced setup"

        filtered["entry_label"] = filtered.apply(entry_note, axis=1)

        # Ranking score
        setup_component = safe_float(stock_setup_score, 0.0) * 8.0
        same_week_bonus = filtered["same_week_expiry"].apply(lambda x: 60.0 if x else 0.0)
        label_bonus = filtered["label"].map({"Conservative": 30.0, "Balanced": 20.0, "Aggressive": 10.0}).fillna(0.0)

        filtered["score"] = (
            filtered["volume"] * 0.25
            + filtered["openInterest"] * 0.25
            + (100.0 - filtered["spread_pct"]) * 18.0
            + (3.0 - (filtered["oi_volume_ratio"] - 1.5).abs()) * 90.0
            + (1.0 - filtered["delta_distance"]) * 220.0
            + same_week_bonus
            + label_bonus
            + setup_component
        )

        all_contracts.append(filtered)

    if not all_contracts:
        return pd.DataFrame()

    final_df = pd.concat(all_contracts, ignore_index=True)

    keep_cols = [
        "ticker", "trade_bias", "expiry", "dte_days", "same_week_expiry",
        "contractSymbol", "strike", "underlying_price", "strike_distance",
        "expected_move", "max_allowed_by_em", "max_allowed_by_atr",
        "lastPrice", "bid", "ask", "mid_price",
        "spread_pct", "volume", "openInterest", "oi_volume_ratio",
        "impliedVolatility", "delta", "label", "entry_label",
        "stop_loss", "take_profit", "score"
    ]
    keep_cols = [c for c in keep_cols if c in final_df.columns]

    final_df = final_df[keep_cols].sort_values(by="score", ascending=False).reset_index(drop=True)
    return final_df


# =========================================================
# EXCEL OUTPUT
# =========================================================
def make_excel_safe(df):
    df = df.copy()
    for col in df.columns:
        if pd.api.types.is_datetime64tz_dtype(df[col]):
            df[col] = df[col].dt.tz_localize(None)
        elif df[col].dtype == "object":
            try:
                converted = pd.to_datetime(df[col], errors="ignore")
                if pd.api.types.is_datetime64tz_dtype(converted):
                    df[col] = converted.dt.tz_localize(None)
            except Exception:
                pass
    return df


def write_best_wining_scan_excel(writer: pd.ExcelWriter, stock_df: pd.DataFrame, contract_df: pd.DataFrame) -> None:
    stock_df = make_excel_safe(stock_df.copy())
    contract_df = make_excel_safe(contract_df.copy())

    stock_df.to_excel(writer, sheet_name="StockSetups", index=False)
    contract_df.to_excel(writer, sheet_name="BestContracts", index=False)

    for sheet_name in ["StockSetups", "BestContracts"]:
        ws = writer.sheets[sheet_name]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for column_cells in ws.columns:
            max_length = 0
            col_letter = column_cells[0].column_letter
            for cell in column_cells:
                try:
                    max_length = max(max_length, len(str(cell.value)))
                except Exception:
                    pass
            ws.column_dimensions[col_letter].width = min(max_length + 2, 35)


def save_results_to_excel(stock_df, contract_df, filename):
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        write_best_wining_scan_excel(writer, stock_df, contract_df)


def best_wining_results_to_excel_bytes(stock_df: pd.DataFrame, contract_df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        write_best_wining_scan_excel(writer, stock_df, contract_df)
    buf.seek(0)
    return buf.getvalue()


def format_best_wining_display_df(df: pd.DataFrame) -> pd.DataFrame:
    """Valid stock rows first (by setup_score); rows with ``error`` last."""
    if df.empty:
        return df
    out = df.copy()
    if "error" not in out.columns:
        out["error"] = pd.NA
    if "setup_score" not in out.columns:
        return out
    ok = out["error"].isna()
    good = out[ok].copy()
    bad = out[~ok].copy()
    if not good.empty:
        good = good.sort_values(
            by=["setup_score", "bullish_votes", "bearish_votes"],
            ascending=[False, False, False],
        ).reset_index(drop=True)
    if not bad.empty:
        bad = bad.reset_index(drop=True)
    return pd.concat([good, bad], ignore_index=True) if not bad.empty else good


def run_best_wining_full_scan(
    symbols: list[str],
    *,
    top_stocks_for_options: int | None = None,
) -> tuple[dict, pd.DataFrame, pd.DataFrame]:
    """
    Stock setup + best option contracts (this module's logic).
    Returns (market_context, stock_df, contract_df).
    """
    top_n = top_stocks_for_options if top_stocks_for_options is not None else TOP_STOCKS_TO_SCAN_OPTIONS
    clean = list(dict.fromkeys(s.strip().upper() for s in symbols if s and str(s).strip()))
    market_context = get_sector_context()
    stock_rows = [analyze_ticker(sym, market_context) for sym in clean]
    stock_df = pd.DataFrame(stock_rows)

    valid = stock_df[stock_df["error"].isna()].copy() if "error" in stock_df.columns else stock_df.copy()
    actionable = valid[valid["trade_bias"].isin(["CALL", "PUT"])].head(top_n).copy()

    contract_frames: list[pd.DataFrame] = []
    for _, row in actionable.iterrows():
        contracts = get_best_contracts_for_ticker(
            ticker_symbol=row["ticker"],
            trade_bias=row["trade_bias"],
            stock_setup_score=safe_float(row.get("setup_score"), 0.0),
            atr_value=safe_float(row.get("atr"), 0.0),
        )
        if not contracts.empty:
            contracts = contracts.copy()
            contracts["stock_setup_score"] = row.get("setup_score")
            contracts["stock_setup_tags"] = row.get("setup_tags", "")
            contract_frames.append(contracts)

    if contract_frames:
        contract_df = pd.concat(contract_frames, ignore_index=True)
        contract_df["final_rank_score"] = contract_df["score"]
        contract_df = contract_df.sort_values(
            by=["final_rank_score", "stock_setup_score"],
            ascending=[False, False],
        ).reset_index(drop=True)
    else:
        contract_df = pd.DataFrame()

    return market_context, stock_df, contract_df


# =========================================================
# MAIN
# =========================================================
def main():
    print("Analyzing market context...")
    market_context = get_sector_context()
    print(
        f"Market Direction: {market_context['market_direction']} | "
        f"SPY: {market_context['SPY_direction']} ({market_context['SPY_daily_change_pct']}%) | "
        f"QQQ: {market_context['QQQ_direction']} ({market_context['QQQ_daily_change_pct']}%)"
    )

    stock_rows = []
    for symbol in TICKERS:
        print(f"Scanning stock setup: {symbol}")
        result = analyze_ticker(symbol, market_context)
        stock_rows.append(result)

    stock_df = pd.DataFrame(stock_rows)

    if "error" in stock_df.columns:
        stock_df = stock_df[stock_df["error"].isna()].copy() if stock_df["error"].notna().any() else stock_df

    if not stock_df.empty:
        stock_df = stock_df.sort_values(
            by=["setup_score", "bullish_votes", "bearish_votes"],
            ascending=[False, False, False]
        ).reset_index(drop=True)

    print("\nTop stock setups:")
    if not stock_df.empty:
        show_cols = [
            "ticker", "trade_bias", "close", "atr", "setup_score",
            "daily_change_pct", "after_hours_change_pct", "setup_tags"
        ]
        show_cols = [c for c in show_cols if c in stock_df.columns]
        print(stock_df[show_cols].head(10).to_string(index=False))
    else:
        print("No stock setups found.")

    actionable = stock_df[stock_df["trade_bias"].isin(["CALL", "PUT"])].head(TOP_STOCKS_TO_SCAN_OPTIONS).copy()

    contract_frames = []
    for _, row in actionable.iterrows():
        symbol = row["ticker"]
        bias = row["trade_bias"]
        atr_value = safe_float(row.get("atr"), 0.0)
        setup_score = safe_float(row.get("setup_score"), 0.0)

        print(f"Scanning options: {symbol} | Bias={bias}")
        contracts = get_best_contracts_for_ticker(
            ticker_symbol=symbol,
            trade_bias=bias,
            stock_setup_score=setup_score,
            atr_value=atr_value
        )

        if not contracts.empty:
            contracts["stock_setup_score"] = setup_score
            contracts["stock_setup_tags"] = row.get("setup_tags", "")
            contract_frames.append(contracts)

    if contract_frames:
        contract_df = pd.concat(contract_frames, ignore_index=True)
        contract_df["final_rank_score"] = contract_df["score"]
        contract_df = contract_df.sort_values(
            by=["final_rank_score", "stock_setup_score"],
            ascending=[False, False]
        ).reset_index(drop=True)
    else:
        contract_df = pd.DataFrame()

    print("\nBest contracts:")
    if not contract_df.empty:
        display_cols = [
            "ticker", "trade_bias", "expiry", "dte_days", "contractSymbol",
            "strike", "delta", "label", "mid_price", "stop_loss", "take_profit",
            "spread_pct", "volume", "openInterest", "oi_volume_ratio",
            "expected_move", "strike_distance", "final_rank_score"
        ]
        display_cols = [c for c in display_cols if c in contract_df.columns]
        print(contract_df[display_cols].head(20).to_string(index=False))
    else:
        print("No contracts matched all filters.")

    save_results_to_excel(stock_df, contract_df, OUTPUT_FILE)
    print(f"\nSaved to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()