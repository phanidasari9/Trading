"""
Market movers analysis — logic aligned with Analysis.txt (yfinance, movers, volume spikes).
"""
from __future__ import annotations

from datetime import datetime
from io import BytesIO
from typing import Any

import pandas as pd
import yfinance as yf

DEFAULT_TICKERS = [
    "AAPL", "MSFT", "NVDA", "AMZN", "TSLA", "META", "GOOGL", "GOOG",
    "AMD", "NFLX", "PLTR", "MU", "AVGO", "COST", "WMT", "INTC",
    "AAL", "QQQ", "SPY", "UEC", "BAC", "JPM", "DIS", "PYPL", "UBER",
]

DEFAULT_HISTORY_PERIOD = "3mo"
DEFAULT_MIN_PRICE = 5
DEFAULT_MIN_AVG_VOLUME = 500_000
DEFAULT_TOP_N = 10
DEFAULT_VOLUME_SPIKE_THRESHOLD = 2.0


def fetch_stock_data(
    tickers: list[str],
    *,
    history_period: str = DEFAULT_HISTORY_PERIOD,
    min_price: float = DEFAULT_MIN_PRICE,
    min_avg_volume: float = DEFAULT_MIN_AVG_VOLUME,
) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []

    for ticker in tickers:
        try:
            stock = yf.Ticker(ticker)
            hist = stock.history(period=history_period, auto_adjust=False)

            if hist.empty or len(hist) < 2:
                continue

            latest = hist.iloc[-1]
            prev = hist.iloc[-2]

            close_price = float(latest["Close"])
            prev_close = float(prev["Close"])
            current_volume = float(latest["Volume"])

            avg_volume_20 = (
                float(hist["Volume"].tail(20).mean())
                if len(hist) >= 20
                else float(hist["Volume"].mean())
            )

            if prev_close == 0 or avg_volume_20 == 0:
                continue

            pct_change = ((close_price - prev_close) / prev_close) * 100
            volume_spike = current_volume / avg_volume_20

            if close_price < min_price:
                continue
            if avg_volume_20 < min_avg_volume:
                continue

            rows.append(
                {
                    "Ticker": ticker,
                    "Close": round(close_price, 2),
                    "Prev Close": round(prev_close, 2),
                    "% Change": round(pct_change, 2),
                    "Volume": int(current_volume),
                    "Avg Volume 20D": int(avg_volume_20),
                    "Volume Spike": round(volume_spike, 2),
                }
            )

        except Exception as exc:  # noqa: BLE001 — per-ticker resilience like Analysis.txt
            print(f"Error processing {ticker}: {exc}")

    return pd.DataFrame(rows)


def get_market_movers(
    df: pd.DataFrame,
    top_n: int = DEFAULT_TOP_N,
    spike_threshold: float = DEFAULT_VOLUME_SPIKE_THRESHOLD,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    if df.empty:
        empty = pd.DataFrame()
        return empty, empty, empty, empty

    most_active = df.sort_values(by="Volume", ascending=False).head(top_n)
    top_gainers = df.sort_values(by="% Change", ascending=False).head(top_n)
    top_losers = df.sort_values(by="% Change", ascending=True).head(top_n)
    volume_spikes = (
        df[df["Volume Spike"] >= spike_threshold]
        .sort_values(by="Volume Spike", ascending=False)
        .head(top_n)
    )

    return most_active, top_gainers, top_losers, volume_spikes


def _write_mover_sheets(
    writer: pd.ExcelWriter,
    most_active: pd.DataFrame,
    top_gainers: pd.DataFrame,
    top_losers: pd.DataFrame,
    volume_spikes: pd.DataFrame,
) -> None:
    most_active.to_excel(writer, sheet_name="Most Active", index=False)
    top_gainers.to_excel(writer, sheet_name="Top Gainers", index=False)
    top_losers.to_excel(writer, sheet_name="Top Losers", index=False)
    volume_spikes.to_excel(writer, sheet_name="Volume Spikes", index=False)


def save_results_to_excel(
    most_active: pd.DataFrame,
    top_gainers: pd.DataFrame,
    top_losers: pd.DataFrame,
    volume_spikes: pd.DataFrame,
    file_name: str,
) -> None:
    with pd.ExcelWriter(file_name, engine="openpyxl") as writer:
        _write_mover_sheets(writer, most_active, top_gainers, top_losers, volume_spikes)


def market_movers_excel_bytes(
    most_active: pd.DataFrame,
    top_gainers: pd.DataFrame,
    top_losers: pd.DataFrame,
    volume_spikes: pd.DataFrame,
) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        _write_mover_sheets(writer, most_active, top_gainers, top_losers, volume_spikes)
    buf.seek(0)
    return buf.getvalue()


def default_excel_filename() -> str:
    return f"market_movers_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
