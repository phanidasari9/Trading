from __future__ import annotations

import math
import os
import sys
import webbrowser
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd
import yfinance as yf


# ============================================================
# CONFIG
# ============================================================

@dataclass
class AnalyzerConfig:
    tickers: list[str]
    target_expiry: Optional[str] = None          # e.g. "2026-04-08" or None
    max_contracts_per_ticker: int = 10
    top_cluster_count: int = 6
    min_volume: int = 1000
    min_vol_oi_ratio: float = 3.0
    max_spread_pct: float = 0.18                 # optional liquidity filter
    same_day_expiry_first: bool = True
    atm_strike_window_pct: float = 0.12          # +/-12% around spot for ATM-focused scan
    export_png: bool = True
    html_file: str = "flow_dashboard_v2.html"
    png_file: str = "flow_dashboard_v2.png"
    auto_open_html: bool = True


# ============================================================
# HELPERS
# ============================================================

def safe_float(x, default=0.0) -> float:
    try:
        if x is None or (isinstance(x, float) and math.isnan(x)):
            return default
        return float(x)
    except Exception:
        return default


def safe_int(x, default=0) -> int:
    try:
        if x is None or (isinstance(x, float) and math.isnan(x)):
            return default
        return int(x)
    except Exception:
        return default


def fmt_money(x: float) -> str:
    x = safe_float(x)
    sign = "-" if x < 0 else ""
    x = abs(x)
    if x >= 1_000_000_000:
        return f"{sign}${x / 1_000_000_000:.2f}B"
    if x >= 1_000_000:
        return f"{sign}${x / 1_000_000:.2f}M"
    if x >= 1_000:
        return f"{sign}${x / 1_000:.1f}K"
    return f"{sign}${x:.2f}"


def fmt_pct(x: float) -> str:
    return f"{safe_float(x):.2f}%"


def zscore(series: pd.Series) -> pd.Series:
    s = pd.to_numeric(series, errors="coerce").fillna(0.0)
    std = s.std(ddof=0)
    if std == 0 or np.isnan(std):
        return pd.Series(np.zeros(len(s)), index=s.index)
    return (s - s.mean()) / std


# ============================================================
# CORE ANALYZER
# ============================================================

class MarketFlowAnalyzerV2:
    def __init__(self, config: AnalyzerConfig):
        self.config = config

    def get_price_data(self, ticker: str) -> pd.DataFrame:
        hist = yf.Ticker(ticker).history(period="6mo", interval="1d", auto_adjust=True)
        if hist.empty:
            raise ValueError(f"No price history returned for {ticker}.")
        return hist

    def choose_expiry(self, ticker_obj: yf.Ticker, target_expiry: Optional[str]) -> str:
        expiries = list(ticker_obj.options)
        if not expiries:
            raise ValueError("No expiries returned.")

        if target_expiry:
            if target_expiry not in expiries:
                raise ValueError(
                    f"Requested expiry {target_expiry} not found. Available expiries: {expiries}"
                )
            return target_expiry

        today = datetime.now().date()

        parsed = []
        for e in expiries:
            try:
                d = datetime.strptime(e, "%Y-%m-%d").date()
                parsed.append((e, d))
            except Exception:
                continue

        if not parsed:
            return expiries[0]

        # Same-day expiry priority, else nearest future expiry
        if self.config.same_day_expiry_first:
            same_day = [e for e, d in parsed if d == today]
            if same_day:
                return same_day[0]

        future = [(e, d) for e, d in parsed if d >= today]
        if future:
            future.sort(key=lambda x: x[1])
            return future[0][0]

        parsed.sort(key=lambda x: x[1], reverse=True)
        return parsed[0][0]

    def get_option_chain(self, ticker: str, expiry: str, spot: float) -> pd.DataFrame:
        tk = yf.Ticker(ticker)
        chain = tk.option_chain(expiry)

        calls = chain.calls.copy()
        puts = chain.puts.copy()

        if calls.empty and puts.empty:
            raise ValueError(f"No chain data found for {ticker} {expiry}")

        calls["side"] = "CALL"
        puts["side"] = "PUT"

        df = pd.concat([calls, puts], ignore_index=True)

        numeric_cols = [
            "strike", "lastPrice", "bid", "ask", "change", "percentChange",
            "volume", "openInterest", "impliedVolatility"
        ]
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")

        df["ticker"] = ticker
        df["expiry"] = expiry
        df["volume"] = df["volume"].fillna(0)
        df["openInterest"] = df["openInterest"].fillna(0)
        df["bid"] = df["bid"].fillna(0.0)
        df["ask"] = df["ask"].fillna(0.0)
        df["lastPrice"] = df["lastPrice"].fillna(0.0)
        df["impliedVolatility"] = df["impliedVolatility"].fillna(0.0)

        df["mid"] = np.where(
            (df["bid"] > 0) & (df["ask"] > 0),
            (df["bid"] + df["ask"]) / 2.0,
            df["lastPrice"]
        )
        df["mid"] = df["mid"].replace([np.inf, -np.inf], np.nan).fillna(0.0)

        df["spread"] = np.maximum(df["ask"] - df["bid"], 0.0)
        df["spread_pct"] = np.where(df["mid"] > 0, df["spread"] / df["mid"], 999.0)

        df["vol_oi_ratio"] = np.where(
            df["openInterest"] > 0,
            df["volume"] / df["openInterest"],
            np.where(df["volume"] > 0, 999.0, 0.0)
        )

        df["premium_traded"] = df["mid"] * df["volume"] * 100.0
        df["notional_oi"] = df["mid"] * df["openInterest"] * 100.0

        df["distance_from_spot_pct"] = np.where(
            spot > 0,
            np.abs(df["strike"] - spot) / spot,
            999.0
        )
        df["moneyness_signed_pct"] = np.where(
            spot > 0,
            (df["strike"] - spot) / spot,
            0.0
        )

        # Strong unusual filters
        df["passes_unusual_filter"] = (
            (df["volume"] >= self.config.min_volume) &
            (df["vol_oi_ratio"] >= self.config.min_vol_oi_ratio) &
            (df["spread_pct"] <= self.config.max_spread_pct)
        )

        # Heuristic "sweep-like"
        df["is_sweep_like"] = (
            df["passes_unusual_filter"] &
            (df["premium_traded"] >= 100_000)
        )

        # ATM focus
        df["within_atm_window"] = (
            df["distance_from_spot_pct"] <= self.config.atm_strike_window_pct
        )

        # Scoring
        df["z_volume"] = zscore(df["volume"])
        df["z_ratio"] = zscore(df["vol_oi_ratio"].replace(999.0, np.nan).fillna(df["vol_oi_ratio"].median()))
        df["z_premium"] = zscore(df["premium_traded"])
        df["z_iv"] = zscore(df["impliedVolatility"])
        df["atm_bonus"] = np.where(df["within_atm_window"], 1.0, -0.5)
        df["spread_penalty"] = np.clip(df["spread_pct"], 0, 1.0) * 2.0

        df["unusual_score"] = (
            df["z_volume"] * 2.5 +
            df["z_ratio"] * 2.8 +
            df["z_premium"] * 2.2 +
            df["atm_bonus"] -
            df["spread_penalty"]
        )

        # Keep active only
        df = df[df["volume"] > 0].copy()
        return df

    def compute_stock_trend(self, hist: pd.DataFrame) -> dict:
        close = hist["Close"].copy()
        last = safe_float(close.iloc[-1])

        sma_10 = safe_float(close.rolling(10).mean().iloc[-1]) if len(close) >= 10 else last
        sma_20 = safe_float(close.rolling(20).mean().iloc[-1]) if len(close) >= 20 else last
        sma_50 = safe_float(close.rolling(50).mean().iloc[-1]) if len(close) >= 50 else last

        ret_5d = ((close.iloc[-1] / close.iloc[-6]) - 1) * 100 if len(close) >= 6 else 0.0
        ret_20d = ((close.iloc[-1] / close.iloc[-21]) - 1) * 100 if len(close) >= 21 else 0.0

        trend_score = 0
        trend_score += 1 if last > sma_10 else 0
        trend_score += 1 if last > sma_20 else 0
        trend_score += 1 if last > sma_50 else 0
        trend_score += 1 if ret_5d > 0 else 0
        trend_score += 1 if ret_20d > 0 else 0

        return {
            "last_price": last,
            "sma_10": sma_10,
            "sma_20": sma_20,
            "sma_50": sma_50,
            "ret_5d": safe_float(ret_5d),
            "ret_20d": safe_float(ret_20d),
            "trend_score": trend_score,  # 0..5
        }

    def build_atm_clusters(self, df: pd.DataFrame, spot: float) -> pd.DataFrame:
        if df.empty:
            return pd.DataFrame()

        work = df[df["within_atm_window"]].copy()
        if work.empty:
            work = df.copy()

        # Cluster by exact strike
        grouped = (
            work.groupby(["ticker", "expiry", "strike"], as_index=False)
            .agg(
                total_volume=("volume", "sum"),
                total_oi=("openInterest", "sum"),
                total_premium=("premium_traded", "sum"),
                avg_score=("unusual_score", "mean"),
                call_volume=("side", lambda s: 0),  # temp
            )
        )

        side_pivot = (
            work.pivot_table(
                index=["ticker", "expiry", "strike"],
                columns="side",
                values="volume",
                aggfunc="sum",
                fill_value=0
            )
            .reset_index()
        )

        if "CALL" not in side_pivot.columns:
            side_pivot["CALL"] = 0
        if "PUT" not in side_pivot.columns:
            side_pivot["PUT"] = 0

        grouped = grouped.merge(side_pivot, on=["ticker", "expiry", "strike"], how="left")
        grouped["distance_pct"] = np.where(spot > 0, np.abs(grouped["strike"] - spot) / spot, 999.0)
        grouped["cluster_score"] = (
            np.log1p(grouped["total_volume"]) * 2.0 +
            np.log1p(grouped["total_premium"] / 1000.0) * 1.8 +
            grouped["avg_score"] * 1.5 -
            grouped["distance_pct"] * 8.0
        )

        grouped["dominant_side"] = np.where(grouped["CALL"] >= grouped["PUT"], "CALL", "PUT")

        grouped = grouped.sort_values(
            ["cluster_score", "total_premium", "total_volume"],
            ascending=False
        ).head(self.config.top_cluster_count)

        return grouped

    def compute_summary(self, ticker: str, df: pd.DataFrame, trend: dict) -> dict:
        calls = df[df["side"] == "CALL"].copy()
        puts = df[df["side"] == "PUT"].copy()

        call_volume = safe_int(calls["volume"].sum())
        put_volume = safe_int(puts["volume"].sum())
        call_oi = safe_int(calls["openInterest"].sum())
        put_oi = safe_int(puts["openInterest"].sum())

        call_premium = safe_float(calls["premium_traded"].sum())
        put_premium = safe_float(puts["premium_traded"].sum())

        unusual_calls = calls[calls["passes_unusual_filter"]]
        unusual_puts = puts[puts["passes_unusual_filter"]]

        call_unusual = len(unusual_calls)
        put_unusual = len(unusual_puts)

        total_premium = call_premium + put_premium
        premium_bias = 0.0 if total_premium <= 0 else (call_premium - put_premium) / total_premium

        total_volume = call_volume + put_volume
        volume_bias = 0.0 if total_volume <= 0 else (call_volume - put_volume) / total_volume

        unusual_total = call_unusual + put_unusual
        unusual_bias = 0.0 if unusual_total == 0 else (call_unusual - put_unusual) / unusual_total

        trend_bias = (trend["trend_score"] - 2.5) / 2.5

        raw_bias = (
            0.38 * premium_bias +
            0.22 * volume_bias +
            0.24 * unusual_bias +
            0.16 * trend_bias
        )

        score = max(0, min(100, round(50 + raw_bias * 50)))

        if score >= 60:
            bias_label = "Bullish"
        elif score <= 40:
            bias_label = "Bearish"
        else:
            bias_label = "Neutral"

        distance = abs(score - 50)
        if distance >= 24:
            confidence = "HIGH Conf"
        elif distance >= 14:
            confidence = "MED Conf"
        else:
            confidence = "LOW Conf"

        sweep_count = safe_int(df["is_sweep_like"].sum())
        sweep_flag = "SWEEP" if sweep_count >= 3 else "MIXED"

        pcr_vol = (put_volume / call_volume) if call_volume > 0 else np.nan
        pcr_premium = (put_premium / call_premium) if call_premium > 0 else np.nan

        return {
            "ticker": ticker,
            "score": score,
            "bias_label": bias_label,
            "confidence": confidence,
            "sweep_flag": sweep_flag,
            "call_volume": call_volume,
            "put_volume": put_volume,
            "call_oi": call_oi,
            "put_oi": put_oi,
            "call_premium": call_premium,
            "put_premium": put_premium,
            "call_unusual": call_unusual,
            "put_unusual": put_unusual,
            "pcr_vol": pcr_vol,
            "pcr_premium": pcr_premium,
            "trend_5d": trend["ret_5d"],
            "trend_20d": trend["ret_20d"],
            "last_price": trend["last_price"],
            "same_day_or_nearest_logic": True,
        }

    def get_top_unusual(self, df: pd.DataFrame) -> pd.DataFrame:
        work = df.copy()

        # Prioritize unusual filter, ATM, then score
        work["rank_bucket"] = np.where(work["passes_unusual_filter"], 1, 0)
        work["atm_bucket"] = np.where(work["within_atm_window"], 1, 0)

        work = work.sort_values(
            ["rank_bucket", "atm_bucket", "unusual_score", "premium_traded", "volume"],
            ascending=False
        )

        work = work.head(self.config.max_contracts_per_ticker).copy()
        return work

    def analyze_ticker(self, ticker: str) -> dict:
        ticker = ticker.upper().strip()
        tk = yf.Ticker(ticker)

        hist = self.get_price_data(ticker)
        trend = self.compute_stock_trend(hist)
        spot = trend["last_price"]

        expiry = self.choose_expiry(tk, self.config.target_expiry)
        chain = self.get_option_chain(ticker, expiry, spot)
        chain = chain[chain["within_atm_window"] | chain["passes_unusual_filter"]].copy()

        if chain.empty:
            # fallback to raw chain without ATM filter
            chain = self.get_option_chain(ticker, expiry, spot)

        summary = self.compute_summary(ticker, chain, trend)
        top_unusual = self.get_top_unusual(chain)
        atm_clusters = self.build_atm_clusters(chain, spot)

        return {
            "ticker": ticker,
            "expiry": expiry,
            "spot": spot,
            "trend": trend,
            "summary": summary,
            "top_unusual": top_unusual,
            "atm_clusters": atm_clusters,
            "full_chain": chain,
        }

    def run_all(self) -> list[dict]:
        results = []
        for t in self.config.tickers:
            try:
                results.append(self.analyze_ticker(t))
            except Exception as e:
                results.append({
                    "ticker": t.upper(),
                    "error": str(e)
                })
        return results


# ============================================================
# HTML RENDERER
# ============================================================

class DashboardRendererV2:
    def __init__(self, results: list[dict], config: AnalyzerConfig):
        self.results = results
        self.config = config

    @staticmethod
    def bias_color(label: str) -> str:
        if label == "Bullish":
            return "#22c55e"
        if label == "Bearish":
            return "#15803d"
        return "#f59e0b"

    @staticmethod
    def html_escape(text) -> str:
        s = str(text)
        return (
            s.replace("&", "&amp;")
             .replace("<", "&lt;")
             .replace(">", "&gt;")
             .replace('"', "&quot;")
        )

    def render_premium_bar(self, call_premium: float, put_premium: float) -> str:
        total = max(call_premium + put_premium, 1.0)
        call_pct = (call_premium / total) * 100
        put_pct = (put_premium / total) * 100
        return f"""
        <div class="premium-bar-wrap">
            <div class="premium-bar">
                <div class="premium-call" style="width:{call_pct:.2f}%"></div>
                <div class="premium-put" style="width:{put_pct:.2f}%"></div>
            </div>
            <div class="premium-bar-labels">
                <span class="green">CALL Premium {fmt_money(call_premium)}</span>
                <span class="red">PUT Premium {fmt_money(put_premium)}</span>
            </div>
        </div>
        """

    def render_summary_table(self, s: dict) -> str:
        def safe_ratio(x):
            if x is None or (isinstance(x, float) and np.isnan(x)):
                return "—"
            return f"{x:.2f}"

        return f"""
        <table class="cp-table">
            <thead>
                <tr>
                    <th>Metric</th>
                    <th class="green">CALLS</th>
                    <th class="red">PUTS</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>Volume</td>
                    <td class="green">{s["call_volume"]:,}</td>
                    <td class="red">{s["put_volume"]:,}</td>
                </tr>
                <tr>
                    <td>Open Interest</td>
                    <td class="green">{s["call_oi"]:,}</td>
                    <td class="red">{s["put_oi"]:,}</td>
                </tr>
                <tr>
                    <td>Premium Traded</td>
                    <td class="green">{fmt_money(s["call_premium"])}</td>
                    <td class="red">{fmt_money(s["put_premium"])}</td>
                </tr>
                <tr>
                    <td>Unusual Count</td>
                    <td class="green">{s["call_unusual"]:,}</td>
                    <td class="red">{s["put_unusual"]:,}</td>
                </tr>
                <tr>
                    <td>P/C Volume Ratio</td>
                    <td colspan="2">{safe_ratio(s["pcr_vol"])}</td>
                </tr>
                <tr>
                    <td>P/C Premium Ratio</td>
                    <td colspan="2">{safe_ratio(s["pcr_premium"])}</td>
                </tr>
            </tbody>
        </table>
        """

    def render_clusters(self, clusters: pd.DataFrame, spot: float) -> str:
        if clusters is None or clusters.empty:
            return '<div class="muted">No ATM strike clusters found.</div>'

        rows = []
        for _, r in clusters.iterrows():
            side_class = "green" if r["dominant_side"] == "CALL" else "red"
            rows.append(f"""
            <tr>
                <td>${safe_float(r["strike"]):.1f}</td>
                <td>{safe_int(r["total_volume"]):,}</td>
                <td>{fmt_money(safe_float(r["total_premium"]))}</td>
                <td class="{side_class}">{self.html_escape(r["dominant_side"])}</td>
                <td>{safe_float(r["distance_pct"]) * 100:.2f}%</td>
            </tr>
            """)

        return f"""
        <div class="subcard">
            <div class="subcard-title">Nearest ATM Strike Clustering <span class="muted">(Spot ${spot:.2f})</span></div>
            <table class="cluster-table">
                <thead>
                    <tr>
                        <th>Strike</th>
                        <th>Total Vol</th>
                        <th>Premium</th>
                        <th>Dominant</th>
                        <th>Distance</th>
                    </tr>
                </thead>
                <tbody>
                    {''.join(rows)}
                </tbody>
            </table>
        </div>
        """

    def render_unusual_list(self, ticker: str, expiry: str, df: pd.DataFrame) -> str:
        if df is None or df.empty:
            return '<div class="muted">No unusual contracts matched the current filters.</div>'

        max_prem = max(float(df["premium_traded"].max()), 1.0)

        items = []
        for _, row in df.iterrows():
            is_call = row["side"] == "CALL"
            icon = "📈" if is_call else "📉"
            color_class = "green" if is_call else "red"
            bar_pct = (safe_float(row["premium_traded"]) / max_prem) * 100
            sweep = '<span class="tag sweep">SWEEP</span>' if bool(row["is_sweep_like"]) else ""
            unusual = '<span class="tag unusual">UNUSUAL</span>' if bool(row["passes_unusual_filter"]) else ""
            atm = '<span class="tag atm">ATM</span>' if bool(row["within_atm_window"]) else ""

            items.append(f"""
            <div class="flow-item">
                <div class="flow-head">
                    <div class="flow-line">
                        <span class="flow-icon">{icon}</span>
                        <span class="flow-symbol">{ticker}</span>
                        <span>— {row["side"]} ${safe_float(row["strike"]):.1f} exp {expiry}</span>
                    </div>
                    <div class="flow-tags">{unusual}{sweep}{atm}</div>
                </div>

                <div class="flow-bar-track">
                    <div class="flow-bar {color_class}" style="width:{bar_pct:.2f}%"></div>
                </div>

                <div class="flow-meta">
                    <span>Vol: <span class="{color_class}">{safe_int(row["volume"]):,}</span></span>
                    <span>OI: {safe_int(row["openInterest"]):,}</span>
                    <span>Vol/OI: {safe_float(row["vol_oi_ratio"]):.1f}x</span>
                    <span>Premium: {fmt_money(safe_float(row["premium_traded"]))}</span>
                    <span>Spread: {safe_float(row["spread_pct"]) * 100:.1f}%</span>
                    <span>Score: {safe_float(row["unusual_score"]):.2f}</span>
                </div>
            </div>
            """)

        return "\n".join(items)

    def render_card(self, result: dict) -> str:
        if "error" in result:
            return f"""
            <section class="ticker-card error-card">
                <div class="ticker-header">
                    <div class="ticker-title">{self.html_escape(result["ticker"])}</div>
                    <div class="error-text">{self.html_escape(result["error"])}</div>
                </div>
            </section>
            """

        ticker = result["ticker"]
        expiry = result["expiry"]
        summary = result["summary"]
        trend = result["trend"]
        top_unusual = result["top_unusual"]
        clusters = result["atm_clusters"]

        bias_color = self.bias_color(summary["bias_label"])

        return f"""
        <section class="ticker-card">
            <div class="section-title">📊 Detailed Flow Analysis</div>

            <div class="summary-card">
                <div class="summary-main">
                    <div class="summary-pill-row">
                        <span class="dot"></span>
                        <span class="summary-symbol">{ticker}</span>
                        <span>— Score: <span class="accent" style="color:{bias_color}">{summary["score"]}</span></span>
                        <span>— <span class="accent" style="color:{bias_color}">{summary["bias_label"]}</span></span>
                        <span>— {summary["confidence"]}</span>
                        <span class="badge badge-sweep">{summary["sweep_flag"]}</span>
                    </div>
                    <div class="summary-sub">
                        Last: ${summary["last_price"]:.2f} |
                        Expiry: {expiry} |
                        5D: {fmt_pct(summary["trend_5d"])} |
                        20D: {fmt_pct(summary["trend_20d"])}
                    </div>
                </div>

                {self.render_premium_bar(summary["call_premium"], summary["put_premium"])}

                <div class="grid-two">
                    <div class="subcard">
                        <div class="subcard-title">Call / Put Summary</div>
                        {self.render_summary_table(summary)}
                    </div>
                    {self.render_clusters(clusters, result["spot"])}
                </div>
            </div>

            <div class="section-title flame">🔥 Top Unusual Activity</div>
            <div class="unusual-list">
                {self.render_unusual_list(ticker, expiry, top_unusual)}
            </div>
        </section>
        """

    def build_html(self) -> str:
        generated = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cards = "\n".join(self.render_card(r) for r in self.results)

        return f"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Market Flow Dashboard V2</title>
<style>
    * {{
        box-sizing: border-box;
    }}
    body {{
        margin: 0;
        font-family: Inter, Arial, Helvetica, sans-serif;
        background:
            radial-gradient(circle at top left, rgba(31,55,99,.25), transparent 28%),
            linear-gradient(180deg, #07101c 0%, #030712 100%);
        color: #f5f7fb;
        padding: 24px;
    }}
    .page {{
        max-width: 1180px;
        margin: 0 auto;
    }}
    .page-title {{
        font-size: 28px;
        font-weight: 800;
        margin-bottom: 6px;
    }}
    .page-sub {{
        color: #94a3b8;
        margin-bottom: 22px;
        font-size: 14px;
    }}
    .dashboard {{
        display: grid;
        grid-template-columns: 1fr;
        gap: 20px;
    }}
    .ticker-card {{
        background: rgba(7, 14, 27, 0.92);
        border: 1px solid #1e2d44;
        border-radius: 20px;
        box-shadow: 0 12px 28px rgba(0,0,0,0.32);
        padding: 20px;
    }}
    .error-card {{
        border-color: #5a1f2a;
    }}
    .error-text {{
        color: #fda4af;
        margin-top: 8px;
    }}
    .section-title {{
        font-size: 17px;
        font-weight: 800;
        margin-bottom: 14px;
        display: flex;
        align-items: center;
        gap: 8px;
    }}
    .flame {{
        margin-top: 18px;
    }}
    .summary-card {{
        border: 1px solid #243551;
        border-radius: 16px;
        background: linear-gradient(180deg, #0a1425 0%, #08111f 100%);
        padding: 18px;
    }}
    .summary-main {{
        margin-bottom: 14px;
    }}
    .summary-pill-row {{
        display: flex;
        flex-wrap: wrap;
        align-items: center;
        gap: 8px;
        font-size: 18px;
        font-weight: 750;
    }}
    .summary-sub {{
        margin-top: 8px;
        color: #9fb0c8;
        font-size: 13px;
    }}
    .dot {{
        width: 14px;
        height: 14px;
        border-radius: 50%;
        display: inline-block;
        background: linear-gradient(180deg, #ffd76a 0%, #d6a92c 100%);
        box-shadow: 0 0 12px rgba(255, 214, 86, 0.35);
    }}
    .summary-symbol {{
        font-weight: 900;
    }}
    .badge {{
        display: inline-flex;
        align-items: center;
        border-radius: 999px;
        padding: 5px 10px;
        font-size: 12px;
        font-weight: 800;
        letter-spacing: 0.03em;
        border: 1px solid #334766;
        background: #101b31;
    }}
    .badge-sweep {{
        color: #f8fafc;
    }}
    .premium-bar-wrap {{
        margin: 14px 0 18px 0;
    }}
    .premium-bar {{
        width: 100%;
        height: 18px;
        border-radius: 999px;
        overflow: hidden;
        display: flex;
        background: #101b31;
        border: 1px solid #2a3c5b;
    }}
    .premium-call {{
        background: linear-gradient(90deg, #16a34a 0%, #22c55e 100%);
        height: 100%;
    }}
    .premium-put {{
        background: linear-gradient(90deg, #047857 0%, #10b981 100%);
        height: 100%;
    }}
    .premium-bar-labels {{
        display: flex;
        justify-content: space-between;
        gap: 12px;
        margin-top: 8px;
        font-size: 13px;
        color: #cbd5e1;
    }}
    .grid-two {{
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 16px;
    }}
    .subcard {{
        background: #08111f;
        border: 1px solid #1e2d44;
        border-radius: 14px;
        padding: 14px;
    }}
    .subcard-title {{
        font-size: 14px;
        font-weight: 800;
        margin-bottom: 10px;
    }}
    .cp-table, .cluster-table {{
        width: 100%;
        border-collapse: collapse;
        font-size: 13px;
    }}
    .cp-table th, .cp-table td,
    .cluster-table th, .cluster-table td {{
        padding: 9px 8px;
        border-bottom: 1px solid #17253a;
        text-align: left;
    }}
    .cp-table th, .cluster-table th {{
        color: #cbd5e1;
        font-weight: 800;
        font-size: 12px;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }}
    .green {{
        color: #22c55e;
    }}
    .red {{
        color: #16a34a;
    }}
    .muted {{
        color: #94a3b8;
    }}
    .unusual-list {{
        display: flex;
        flex-direction: column;
        gap: 12px;
    }}
    .flow-item {{
        background: linear-gradient(180deg, #07101c 0%, #050c17 100%);
        border: 1px solid #16253a;
        border-radius: 14px;
        padding: 14px;
    }}
    .flow-head {{
        display: flex;
        justify-content: space-between;
        align-items: start;
        gap: 12px;
        margin-bottom: 10px;
    }}
    .flow-line {{
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
        align-items: center;
        font-size: 18px;
        font-weight: 700;
    }}
    .flow-icon {{
        font-size: 18px;
    }}
    .flow-symbol {{
        font-weight: 900;
    }}
    .flow-tags {{
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
    }}
    .tag {{
        border-radius: 999px;
        padding: 4px 9px;
        font-size: 11px;
        font-weight: 800;
        border: 1px solid transparent;
    }}
    .tag.unusual {{
        color: #f8fafc;
        background: #1d4ed8;
        border-color: #295dcf;
    }}
    .tag.sweep {{
        color: #fff7ed;
        background: #b45309;
        border-color: #c26a17;
    }}
    .tag.atm {{
        color: #ecfdf5;
        background: #166534;
        border-color: #1f7a42;
    }}
    .flow-bar-track {{
        width: 100%;
        height: 12px;
        border-radius: 999px;
        overflow: hidden;
        background: #0e1728;
        border: 1px solid #1e2d44;
        margin-bottom: 10px;
    }}
    .flow-bar {{
        height: 100%;
        border-radius: 999px;
    }}
    .flow-bar.green {{
        background: linear-gradient(90deg, #15803d 0%, #22c55e 100%);
    }}
    .flow-bar.red {{
        background: linear-gradient(90deg, #047857 0%, #34d399 100%);
    }}
    .flow-meta {{
        display: flex;
        flex-wrap: wrap;
        gap: 14px;
        color: #d5ddeb;
        font-size: 13px;
    }}
    .footer {{
        margin-top: 20px;
        color: #8fa0b8;
        font-size: 12px;
        line-height: 1.6;
    }}
    @media (max-width: 900px) {{
        .grid-two {{
            grid-template-columns: 1fr;
        }}
        .flow-line {{
            font-size: 16px;
        }}
    }}
</style>
</head>
<body>
    <div class="page">
        <div class="page-title">Market Flow Dashboard V2</div>
        <div class="page-sub">
            Multi-ticker unusual options dashboard • Same-day expiry priority • ATM clustering • Generated {generated}
        </div>

        <div class="dashboard">
            {cards}
        </div>

        <div class="footer">
            Filters: volume &gt;= {self.config.min_volume:,}, vol/OI &gt;= {self.config.min_vol_oi_ratio:.1f}, spread_pct &lt;= {self.config.max_spread_pct:.2f}, ATM window +/- {self.config.atm_strike_window_pct * 100:.1f}%.
        </div>
    </div>
</body>
</html>
        """

    def save_html(self, path: str) -> str:
        html = self.build_html()
        out = Path(path).resolve()
        out.write_text(html, encoding="utf-8")
        return str(out)


# ============================================================
# PNG EXPORT VIA PLAYWRIGHT
# ============================================================

def export_html_to_png(html_path: str, png_path: str, width: int = 1440, height: int = 2200) -> bool:
    """
    Exports the generated HTML report to PNG using Playwright.
    Requires:
        pip install playwright
        python -m playwright install chromium
    """
    try:
        from playwright.sync_api import sync_playwright
    except Exception:
        print("Playwright is not installed. Skipping PNG export.")
        return False

    html_uri = Path(html_path).resolve().as_uri()
    png_abs = str(Path(png_path).resolve())

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page(viewport={"width": width, "height": height}, device_scale_factor=1.5)
            page.goto(html_uri, wait_until="networkidle")
            page.screenshot(path=png_abs, full_page=True)
            browser.close()
        return True
    except Exception as e:
        print(f"PNG export failed: {e}")
        return False


# ============================================================
# MAIN
# ============================================================

def main():
    config = AnalyzerConfig(
        tickers=["TSLA", "AAPL", "NVDA", "SPY"],
        target_expiry=None,                 # None => same-day first, else nearest future
        max_contracts_per_ticker=10,
        top_cluster_count=6,
        min_volume=1000,
        min_vol_oi_ratio=3.0,
        max_spread_pct=0.18,
        same_day_expiry_first=True,
        atm_strike_window_pct=0.12,
        export_png=True,
        html_file="flow_dashboard_v2.html",
        png_file="flow_dashboard_v2.png",
        auto_open_html=True,
    )

    analyzer = MarketFlowAnalyzerV2(config)
    results = analyzer.run_all()

    renderer = DashboardRendererV2(results, config)
    html_path = renderer.save_html(config.html_file)

    print("=" * 90)
    print("MARKET FLOW DASHBOARD V2")
    print("=" * 90)
    for r in results:
        if "error" in r:
            print(f"{r['ticker']}: ERROR -> {r['error']}")
            continue
        s = r["summary"]
        print(
            f"{r['ticker']} | Exp {r['expiry']} | Score {s['score']} | "
            f"{s['bias_label']} | {s['confidence']} | {s['sweep_flag']} | "
            f"CallPrem {fmt_money(s['call_premium'])} | PutPrem {fmt_money(s['put_premium'])}"
        )

    print(f"\nHTML saved to: {html_path}")

    if config.export_png:
        ok = export_html_to_png(config.html_file, config.png_file)
        if ok:
            print(f"PNG saved to: {Path(config.png_file).resolve()}")
        else:
            print("PNG export skipped or failed.")

    if config.auto_open_html:
        webbrowser.open(Path(config.html_file).resolve().as_uri())


if __name__ == "__main__":
    main()