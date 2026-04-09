"""
Trading dashboard — Streamlit UI over market_analysis (see Analysis.txt).
Run: streamlit run dashboard.py
"""
from __future__ import annotations

import io
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
import streamlit.components.v1 as components
from BestWiningOptionsv1 import (
    best_wining_results_to_excel_bytes,
    format_best_wining_display_df,
    run_best_wining_full_scan,
)
from Flow import money_flow_excel_bytes, run_money_flow_scan
from OptionSwing import format_swing_display_df, run_swing_full_scan, swing_results_to_excel_bytes
from Options import (
    AnalyzerConfig,
    DashboardRendererV2,
    MarketFlowAnalyzerV2,
    export_html_to_png,
    fmt_money,
)

from market_analysis import (
    DEFAULT_MIN_AVG_VOLUME,
    DEFAULT_MIN_PRICE,
    DEFAULT_TICKERS,
    DEFAULT_TOP_N,
    DEFAULT_VOLUME_SPIKE_THRESHOLD,
    default_excel_filename,
    fetch_stock_data,
    get_market_movers,
    market_movers_excel_bytes,
)
from watchlists_store import (
    delete_watchlist,
    ensure_default_universe,
    list_names,
    load_store,
    upsert_watchlist,
    validate_list_name,
)
from trading_ui import (
    CHART_BIAS,
    CHART_PREMIUM,
    CHART_SCALE_GAIN,
    CHART_SCALE_LOSS,
    CHART_SCALE_SPIKE,
    apply_trading_theme,
)

st.set_page_config(
    page_title="Trading Command Center",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

apply_trading_theme()
st.title("Trading Command Center")
st.caption("Market movers, options flow, swing setups, and best contracts.")


def parse_tickers(text: str) -> list[str]:
    parts: list[str] = []
    for line in text.replace(",", "\n").splitlines():
        t = line.strip().upper()
        if t:
            parts.append(t)
    return list(dict.fromkeys(parts))


def tickers_for_run(tickers_raw: str) -> list[str]:
    """If sidebar *Single ticker* is set, return only that symbol; else use the editor list."""
    box = str(st.session_state.get("single_ticker_override") or "").strip().upper()
    if box:
        token = box.replace(",", " ").split()[0].strip().lstrip("$")
        if token:
            return [token]
    return parse_tickers(tickers_raw)


def _on_watchlist_change() -> None:
    store_lists = load_store()["lists"]
    name = st.session_state.wl_selector
    st.session_state.tickers_area = "\n".join(store_lists.get(name, []))


lists = ensure_default_universe(DEFAULT_TICKERS)
names = list_names(lists)

if "wl_selector" not in st.session_state:
    st.session_state.wl_selector = "Default" if "Default" in lists else names[0]
if st.session_state.wl_selector not in lists:
    st.session_state.wl_selector = names[0]

if "tickers_area" not in st.session_state:
    st.session_state.tickers_area = "\n".join(lists[st.session_state.wl_selector])

with st.sidebar:
    st.header("Watchlists")
    st.selectbox(
        "Active watchlist",
        options=names,
        key="wl_selector",
        on_change=_on_watchlist_change,
        help="Choosing a list loads its tickers into the editor below.",
    )
    tickers_raw = st.text_area(
        "Tickers (one per line or comma-separated)",
        key="tickers_area",
        height=200,
        help="Edits are in memory until you save to the active watchlist.",
    )
    st.text_input(
        "Single ticker (optional)",
        key="single_ticker_override",
        placeholder="e.g. AAPL",
        help="When filled, **Refresh data** and all scans use only this symbol. Clear the field to use the list above.",
    )
    c_save, c_del = st.columns(2)
    with c_save:
        if st.button("Save → list", use_container_width=True, help="Overwrite active watchlist with editor"):
            tickers = parse_tickers(st.session_state.tickers_area)
            if not tickers:
                st.warning("Add at least one ticker before saving.")
            else:
                fresh = load_store()["lists"]
                upsert_watchlist(fresh, st.session_state.wl_selector, tickers)
                st.success(f"Saved **{st.session_state.wl_selector}** ({len(tickers)} tickers).")
    with c_del:
        if st.button("Delete list", use_container_width=True):
            fresh = load_store()["lists"]
            new_lists = delete_watchlist(fresh, st.session_state.wl_selector)
            if new_lists is None:
                st.warning("Keep at least one watchlist.")
            else:
                st.session_state.wl_selector = list_names(new_lists)[0]
                st.session_state.tickers_area = "\n".join(new_lists[st.session_state.wl_selector])
                st.rerun()

    st.text_input("New list name", key="new_wl_name", placeholder="e.g. Mag 7")
    if st.button("Create from editor", use_container_width=True):
        name = str(st.session_state.get("new_wl_name") or "").strip()
        err = validate_list_name(name)
        if err:
            st.error(err)
        else:
            fresh = load_store()["lists"]
            if name in fresh:
                st.error("A list with that name already exists.")
            else:
                tickers = parse_tickers(st.session_state.tickers_area)
                if not tickers:
                    st.warning("Add tickers in the editor before creating a list.")
                else:
                    upsert_watchlist(fresh, name, tickers)
                    st.session_state.wl_selector = name
                    st.rerun()

    st.divider()
    st.header("Market movers")
    history_period = st.selectbox(
        "History period (for avg volume)",
        options=["1mo", "3mo", "6mo", "1y", "ytd", "2y", "5y", "max"],
        index=1,
    )
    min_price = st.number_input("Min price ($)", min_value=0.0, value=float(DEFAULT_MIN_PRICE), step=1.0)
    min_avg_volume = st.number_input(
        "Min 20D avg volume",
        min_value=0,
        value=DEFAULT_MIN_AVG_VOLUME,
        step=50_000,
    )
    top_n = st.slider("Top N per category", min_value=3, max_value=25, value=DEFAULT_TOP_N)
    volume_spike_threshold = st.slider(
        "Volume spike threshold (× avg)",
        min_value=1.0,
        max_value=5.0,
        value=float(DEFAULT_VOLUME_SPIKE_THRESHOLD),
        step=0.1,
    )
    run = st.button("Refresh data", type="primary", use_container_width=True)


if "df" not in st.session_state:
    st.session_state.df = pd.DataFrame()
if "last_error" not in st.session_state:
    st.session_state.last_error = None
if "options_results" not in st.session_state:
    st.session_state.options_results = None
if "options_config" not in st.session_state:
    st.session_state.options_config = None
if "options_html_report" not in st.session_state:
    st.session_state.options_html_report = None
if "options_png_bytes" not in st.session_state:
    st.session_state.options_png_bytes = None
if "swing_market_context" not in st.session_state:
    st.session_state.swing_market_context = None
if "swing_df" not in st.session_state:
    st.session_state.swing_df = None
if "swing_contract_df" not in st.session_state:
    st.session_state.swing_contract_df = None
if "best_wining_mc" not in st.session_state:
    st.session_state.best_wining_mc = None
if "best_wining_stock_df" not in st.session_state:
    st.session_state.best_wining_stock_df = None
if "best_wining_contract_df" not in st.session_state:
    st.session_state.best_wining_contract_df = None
if "flow_master_df" not in st.session_state:
    st.session_state.flow_master_df = None
if "flow_bull_df" not in st.session_state:
    st.session_state.flow_bull_df = None
if "flow_bear_df" not in st.session_state:
    st.session_state.flow_bear_df = None
if "flow_options_map" not in st.session_state:
    st.session_state.flow_options_map = None

if run or st.session_state.df.empty:
    with st.spinner("Fetching quotes via yfinance…"):
        tickers = tickers_for_run(tickers_raw)
        if not tickers:
            st.warning("Enter **Single ticker** or add at least one symbol in the editor.")
        else:
            try:
                st.session_state.df = fetch_stock_data(
                    tickers,
                    history_period=history_period,
                    min_price=min_price,
                    min_avg_volume=float(min_avg_volume),
                )
                st.session_state.last_error = None
            except Exception as e:  # noqa: BLE001
                st.session_state.last_error = str(e)
                st.error(f"Fetch failed: {e}")

df = st.session_state.df

tab_movers, tab_options, tab_swing, tab_best, tab_flow = st.tabs(
    ["Market movers", "Options", "Swing options", "BestOptions", "Flow"]
)

with tab_movers:
    if st.session_state.last_error and df.empty:
        st.error(f"Last equity fetch failed: {st.session_state.last_error}")
    if df.empty:
        st.info("Click **Refresh data** in the sidebar to load equity movers.")
    else:
        most_active, top_gainers, top_losers, volume_spikes = get_market_movers(
            df,
            top_n=top_n,
            spike_threshold=volume_spike_threshold,
        )

        _single = str(st.session_state.get("single_ticker_override") or "").strip()
        if _single:
            st.caption(f"**Single:** `{tickers_for_run(tickers_raw)[0]}` — watchlist list ignored for this run.")
        else:
            st.caption(
                f"Universe: **{st.session_state.wl_selector}** "
                f"({len(parse_tickers(tickers_raw))} tickers in editor)."
            )

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Universe (after filters)", len(df))
        k2.metric("Top gainer", f"{top_gainers.iloc[0]['Ticker']}" if not top_gainers.empty else "—")
        k3.metric("Top loser", f"{top_losers.iloc[0]['Ticker']}" if not top_losers.empty else "—")
        k4.metric("Spike names ≥ threshold", len(volume_spikes) if not volume_spikes.empty else 0)

        tab_ma, tab_g, tab_l, tab_vs, tab_all = st.tabs(
            ["Most active", "Top gainers", "Top losers", "Volume spikes", "Full universe"]
        )

        with tab_ma:
            st.subheader("Highest volume")
            st.dataframe(most_active, use_container_width=True, hide_index=True)
            if not most_active.empty:
                fig = px.bar(most_active, x="Ticker", y="Volume", title="Volume by ticker")
                st.plotly_chart(fig, use_container_width=True)

        with tab_g:
            st.subheader("Largest % change up")
            st.dataframe(top_gainers, use_container_width=True, hide_index=True)
            if not top_gainers.empty:
                fig = px.bar(
                    top_gainers,
                    x="Ticker",
                    y="% Change",
                    title="Daily % change",
                    color="% Change",
                    color_continuous_scale=CHART_SCALE_GAIN,
                )
                st.plotly_chart(fig, use_container_width=True)

        with tab_l:
            st.subheader("Largest % change down")
            st.dataframe(top_losers, use_container_width=True, hide_index=True)
            if not top_losers.empty:
                fig = px.bar(
                    top_losers,
                    x="Ticker",
                    y="% Change",
                    title="Daily % change",
                    color="% Change",
                    color_continuous_scale=CHART_SCALE_LOSS,
                )
                st.plotly_chart(fig, use_container_width=True)

        with tab_vs:
            st.subheader(f"Volume spike ≥ {volume_spike_threshold}× 20D average")
            if volume_spikes.empty:
                st.write("No names met the threshold.")
            else:
                st.dataframe(volume_spikes, use_container_width=True, hide_index=True)
                fig = px.bar(
                    volume_spikes,
                    x="Ticker",
                    y="Volume Spike",
                    title="Volume spike multiple",
                    color="Volume Spike",
                    color_continuous_scale=CHART_SCALE_SPIKE,
                )
                st.plotly_chart(fig, use_container_width=True)

        with tab_all:
            st.subheader("All filtered tickers")
            sort_by = st.selectbox(
                "Sort by", options=list(df.columns), index=list(df.columns).index("% Change")
            )
            ascending = st.toggle("Ascending", value=False)
            st.dataframe(
                df.sort_values(by=sort_by, ascending=ascending), use_container_width=True, hide_index=True
            )

        st.divider()
        c_dl, c_xl = st.columns(2)
        with c_dl:
            csv_buf = io.StringIO()
            df.to_csv(csv_buf, index=False)
            st.download_button(
                "Download full universe (CSV)",
                data=csv_buf.getvalue(),
                file_name="universe.csv",
                mime="text/csv",
            )
        with c_xl:
            st.download_button(
                "Download Excel (same sheets as Analysis.txt)",
                data=market_movers_excel_bytes(most_active, top_gainers, top_losers, volume_spikes),
                file_name=default_excel_filename(),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

with tab_options:
    st.subheader("Options")
    st.caption(
        "Merged from `Options.py`: `MarketFlowAnalyzerV2` (chains, unusual activity, ATM clusters, "
        "bias score), `DashboardRendererV2` (HTML report), and optional PNG export. "
        "Tickers use the sidebar editor (same universe as movers)."
    )

    with st.expander("Analyzer settings", expanded=True):
        target_expiry_in = st.text_input(
            "Target expiry (YYYY-MM-DD)",
            value="",
            help="Leave empty to use same-day (if listed) or nearest future expiry.",
        )
        te = target_expiry_in.strip() or None
        max_tickers_scan = st.slider("Max tickers to scan", min_value=1, max_value=30, value=8)
        max_contracts = st.slider("Max contracts per ticker (top unusual)", 3, 25, 10)
        top_clusters = st.slider("ATM cluster rows per ticker", 3, 12, 6)
        min_oi_vol = st.number_input("Min volume", min_value=0, value=1000, step=100)
        min_vol_oi = st.number_input("Min vol/OI ratio", min_value=0.0, value=3.0, step=0.1)
        max_spread_pct = st.slider("Max bid/ask spread (fraction of mid)", 0.05, 0.50, 0.18, 0.01)
        atm_pct = st.slider("ATM strike window (±% of spot)", 5, 25, 12, 1) / 100.0
        same_day_first = st.checkbox("Prefer same-day expiry when available", value=True)

    run_opt = st.button("Run options scan", type="primary", use_container_width=True)

    if run_opt:
        opt_tickers = tickers_for_run(tickers_raw)[:max_tickers_scan]
        if not opt_tickers:
            st.warning("Enter **Single ticker** or add symbols in the sidebar editor.")
        else:
            cfg = AnalyzerConfig(
                tickers=opt_tickers,
                target_expiry=te,
                max_contracts_per_ticker=max_contracts,
                top_cluster_count=top_clusters,
                min_volume=int(min_oi_vol),
                min_vol_oi_ratio=float(min_vol_oi),
                max_spread_pct=float(max_spread_pct),
                same_day_expiry_first=same_day_first,
                atm_strike_window_pct=float(atm_pct),
                export_png=False,
                auto_open_html=False,
            )
            with st.spinner("Fetching option chains (yfinance)…"):
                analyzer = MarketFlowAnalyzerV2(cfg)
                st.session_state.options_results = analyzer.run_all()
                st.session_state.options_config = cfg
                st.session_state.options_png_bytes = None
                html_built = DashboardRendererV2(st.session_state.options_results, cfg).build_html()
                st.session_state.options_html_report = html_built

    results = st.session_state.options_results
    cfg_o = st.session_state.options_config
    html_report = st.session_state.options_html_report

    if results is None:
        st.info("Configure settings above and click **Run options scan**.")
    elif cfg_o is not None and html_report is None:
        st.session_state.options_html_report = DashboardRendererV2(results, cfg_o).build_html()
        html_report = st.session_state.options_html_report

    if results is not None and cfg_o is not None and html_report is not None:
        sub_data, sub_report = st.tabs(["Data & detail", "Visual report (HTML)"])

        unusual_cols = [
            "side",
            "strike",
            "volume",
            "openInterest",
            "vol_oi_ratio",
            "premium_traded",
            "spread_pct",
            "unusual_score",
            "passes_unusual_filter",
            "is_sweep_like",
            "within_atm_window",
        ]
        cluster_cols = [
            "strike",
            "total_volume",
            "total_premium",
            "dominant_side",
            "distance_pct",
            "cluster_score",
        ]

        with sub_data:
            summary_rows: list[dict] = []
            cli_lines: list[str] = []
            for r in results:
                if "error" in r:
                    summary_rows.append(
                        {
                            "Ticker": r.get("ticker", "—"),
                            "Expiry": "—",
                            "Score": None,
                            "Bias": "—",
                            "Confidence": "—",
                            "Sweep": "—",
                            "Call prem": "—",
                            "Put prem": "—",
                            "Error": r["error"],
                        }
                    )
                    cli_lines.append(f"{r.get('ticker', '?')}: ERROR -> {r['error']}")
                    continue
                s = r["summary"]
                summary_rows.append(
                    {
                        "Ticker": r["ticker"],
                        "Expiry": r["expiry"],
                        "Score": s["score"],
                        "Bias": s["bias_label"],
                        "Confidence": s["confidence"],
                        "Sweep": s["sweep_flag"],
                        "Call prem": fmt_money(s["call_premium"]),
                        "Put prem": fmt_money(s["put_premium"]),
                        "Error": "",
                    }
                )
                cli_lines.append(
                    f"{r['ticker']} | Exp {r['expiry']} | Score {s['score']} | "
                    f"{s['bias_label']} | {s['confidence']} | {s['sweep_flag']} | "
                    f"CallPrem {fmt_money(s['call_premium'])} | PutPrem {fmt_money(s['put_premium'])}"
                )

            st.markdown("**Summary**")
            st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

            prem_parts: list[dict] = []
            for r in results:
                if "error" in r:
                    continue
                s = r["summary"]
                prem_parts.append(
                    {"Ticker": r["ticker"], "Side": "Call", "Premium $": s["call_premium"]}
                )
                prem_parts.append({"Ticker": r["ticker"], "Side": "Put", "Premium $": s["put_premium"]})
            if prem_parts:
                figp = px.bar(
                    pd.DataFrame(prem_parts),
                    x="Ticker",
                    y="Premium $",
                    color="Side",
                    barmode="group",
                    title="Premium traded by side (from Options.py summary)",
                    color_discrete_map=CHART_PREMIUM,
                )
                st.plotly_chart(figp, use_container_width=True)

            with st.expander("CLI-style log (same lines as `Options.py` main)", expanded=False):
                st.code(
                    "=" * 60 + "\nMARKET FLOW DASHBOARD V2\n" + "=" * 60 + "\n" + "\n".join(cli_lines),
                    language="text",
                )

            for r in results:
                t = r.get("ticker", "?")
                if "error" in r:
                    with st.expander(f"{t} — error", expanded=False):
                        st.error(r["error"])
                    continue
                s = r["summary"]
                with st.expander(
                    f"{t} — score {s['score']} {s['bias_label']} · exp {r['expiry']} · spot ${r['spot']:.2f}",
                    expanded=False,
                ):
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("Bias score", s["score"])
                    m2.metric("5D trend", f"{s['trend_5d']:.1f}%")
                    m3.metric("20D trend", f"{s['trend_20d']:.1f}%")
                    m4.metric("Unusual", f"{s['call_unusual']}C / {s['put_unusual']}P")
                    tu = r["top_unusual"]
                    if tu is not None and not tu.empty:
                        show_u = [c for c in unusual_cols if c in tu.columns]
                        st.markdown("**Top unusual contracts**")
                        st.dataframe(tu[show_u], use_container_width=True, hide_index=True)
                    else:
                        st.write("No unusual rows after filters.")
                    cl = r["atm_clusters"]
                    if cl is not None and not cl.empty:
                        show_c = [c for c in cluster_cols if c in cl.columns]
                        st.markdown("**ATM strike clusters**")
                        st.dataframe(cl[show_c], use_container_width=True, hide_index=True)

            st.divider()
            dl1, dl2, dl3 = st.columns(3)
            with dl1:
                st.download_button(
                    "Download HTML",
                    data=html_report,
                    file_name="flow_dashboard_v2.html",
                    mime="text/html",
                    use_container_width=True,
                )
            with dl2:
                if st.button("Render PNG (Playwright)", use_container_width=True):
                    with tempfile.TemporaryDirectory() as td:
                        hp = Path(td) / "flow.html"
                        pp = Path(td) / "flow.png"
                        hp.write_text(html_report, encoding="utf-8")
                        ok = export_html_to_png(str(hp), str(pp))
                        if ok and pp.exists():
                            st.session_state.options_png_bytes = pp.read_bytes()
                            st.success("PNG ready — use Download below.")
                        else:
                            st.warning(
                                "PNG export failed. Install: pip install playwright && python -m playwright install chromium"
                            )
            with dl3:
                if st.session_state.options_png_bytes:
                    st.download_button(
                        "Download PNG",
                        data=st.session_state.options_png_bytes,
                        file_name="flow_dashboard_v2.png",
                        mime="image/png",
                        use_container_width=True,
                    )

        with sub_report:
            st.caption("Embedded `DashboardRendererV2` output (Market Flow Dashboard V2). Scroll inside the frame.")
            report_h = st.slider("Report frame height (px)", 800, 5000, 2800, 100)
            components.html(html_report, height=int(report_h), scrolling=True)
            st.download_button(
                "Download this HTML report",
                data=html_report,
                file_name="flow_dashboard_v2.html",
                mime="text/html",
                key="dl_html_subreport",
                use_container_width=True,
            )

with tab_swing:
    st.subheader("Swing options")
    st.caption(
        "Upgraded `OptionSwing.py`: stock setup (**trade_bias**, **setup_score**, ATR) plus **best option "
        "contracts** (liquidity, delta band, expected-move / ATR strike filters, Black–Scholes delta). "
        "Excel matches CLI: sheets **StockSetups** + **BestContracts**."
    )

    max_swing = st.slider("Max tickers to scan (stocks)", min_value=1, max_value=30, value=12, key="swing_max_tickers")
    top_opt = st.slider(
        "Top stocks for option chain scan",
        min_value=1,
        max_value=15,
        value=5,
        key="swing_top_options",
        help="Only CALL/PUT names are scanned; ranked by setup_score like the script.",
    )
    run_swing = st.button("Run swing scan", type="primary", use_container_width=True, key="run_swing_btn")

    if run_swing:
        syms = tickers_for_run(tickers_raw)[:max_swing]
        if not syms:
            st.warning("Enter **Single ticker** or add symbols in the sidebar editor.")
        else:
            try:
                with st.spinner("OptionSwing: stocks + option chains (yfinance — can take a minute)…"):
                    mc, sdf, cdf = run_swing_full_scan(syms, top_stocks_for_options=top_opt)
                    st.session_state.swing_market_context = mc
                    st.session_state.swing_df = sdf
                    st.session_state.swing_contract_df = cdf
            except Exception as e:  # noqa: BLE001
                st.session_state.swing_df = None
                st.session_state.swing_contract_df = None
                st.session_state.swing_market_context = None
                st.error(f"Swing scan failed: {e}")

    if st.session_state.swing_df is None:
        st.info("Click **Run swing scan** to load `OptionSwing` results.")
    else:
        mc = st.session_state.swing_market_context or {}
        sdf = st.session_state.swing_df
        cdf = st.session_state.swing_contract_df
        if cdf is None:
            cdf = pd.DataFrame()
        display_df = format_swing_display_df(sdf)

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Market (composite)", mc.get("market_direction", "—"))
        c2.metric("SPY", f"{mc.get('SPY_direction', '—')} ({mc.get('SPY_daily_change_pct', 0)}%)")
        c3.metric("QQQ", f"{mc.get('QQQ_direction', '—')} ({mc.get('QQQ_daily_change_pct', 0)}%)")
        err_n = int(sdf["error"].notna().sum()) if "error" in sdf.columns else 0
        c4.metric("Stocks", f"{len(sdf)}" + (f" ({err_n} err)" if err_n else ""))
        c5.metric("Contract rows", str(len(cdf)) if not cdf.empty else "0")

        st.markdown("**Stock setups** (sorted; errors last)")
        st.dataframe(display_df, use_container_width=True, hide_index=True)

        chart_df = display_df.copy()
        if "error" in chart_df.columns:
            chart_df = chart_df[chart_df["error"].isna()]
        if not chart_df.empty and "setup_score" in chart_df.columns and "ticker" in chart_df.columns:
            fig_s = px.bar(
                chart_df,
                x="ticker",
                y="setup_score",
                color="trade_bias" if "trade_bias" in chart_df.columns else None,
                title="Setup score by symbol",
                color_discrete_map=CHART_BIAS,
            )
            st.plotly_chart(fig_s, use_container_width=True)

        st.markdown("**Best contracts** (top filtered strikes per script logic)")
        if cdf.empty:
            st.write("No contracts passed filters for the top CALL/PUT names (see `OptionSwing.py` thresholds).")
        else:
            st.dataframe(cdf, use_container_width=True, hide_index=True)
            show_c = [c for c in ("ticker", "trade_bias", "expiry", "strike", "delta", "label", "mid_price", "score") if c in cdf.columns]
            if show_c and len(cdf) <= 50:
                fig_c = px.scatter(
                    cdf.head(40),
                    x="strike" if "strike" in cdf.columns else show_c[0],
                    y="score" if "score" in cdf.columns else show_c[-1],
                    color="ticker" if "ticker" in cdf.columns else None,
                    size="volume" if "volume" in cdf.columns else None,
                    hover_data=show_c,
                    title="Contract rank score (sample)",
                    color_discrete_sequence=px.colors.qualitative.Plotly,
                )
                st.plotly_chart(fig_c, use_container_width=True)

        xname = f"next_day_options_upgraded_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.download_button(
            "Download Excel (StockSetups + BestContracts)",
            data=swing_results_to_excel_bytes(display_df, cdf),
            file_name=xname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_swing_xlsx",
        )

with tab_best:
    st.subheader("BestOptions")
    st.caption(
        "Runs **`BestWiningOptionsv1.py`** (separate from `OptionSwing.py`): stock setup scores, then ranked "
        "**BestContracts** with delta / liquidity / expected-move filters. Tune this file independently."
    )

    max_best = st.slider(
        "Max tickers to scan (stocks)",
        min_value=1,
        max_value=30,
        value=12,
        key="best_max_tickers",
    )
    top_best = st.slider(
        "Top stocks for option chain scan",
        min_value=1,
        max_value=15,
        value=5,
        key="best_top_options",
    )
    run_best = st.button("Run BestOptions scan", type="primary", use_container_width=True, key="run_best_btn")

    if run_best:
        syms_b = tickers_for_run(tickers_raw)[:max_best]
        if not syms_b:
            st.warning("Enter **Single ticker** or add symbols in the sidebar editor.")
        else:
            try:
                with st.spinner("BestWiningOptionsv1: stocks + option chains…"):
                    bmc, bsdf, bcdf = run_best_wining_full_scan(syms_b, top_stocks_for_options=top_best)
                    st.session_state.best_wining_mc = bmc
                    st.session_state.best_wining_stock_df = bsdf
                    st.session_state.best_wining_contract_df = bcdf
            except Exception as e:  # noqa: BLE001
                st.session_state.best_wining_stock_df = None
                st.session_state.best_wining_contract_df = None
                st.session_state.best_wining_mc = None
                st.error(f"BestOptions scan failed: {e}")

    if st.session_state.best_wining_stock_df is None:
        st.info("Click **Run BestOptions scan** to load `BestWiningOptionsv1` results.")
    else:
        bmc = st.session_state.best_wining_mc or {}
        bsdf = st.session_state.best_wining_stock_df
        bcdf = st.session_state.best_wining_contract_df
        if bcdf is None:
            bcdf = pd.DataFrame()
        bdisplay = format_best_wining_display_df(bsdf)

        b1, b2, b3, b4, b5 = st.columns(5)
        b1.metric("Market (composite)", bmc.get("market_direction", "—"))
        b2.metric("SPY", f"{bmc.get('SPY_direction', '—')} ({bmc.get('SPY_daily_change_pct', 0)}%)")
        b3.metric("QQQ", f"{bmc.get('QQQ_direction', '—')} ({bmc.get('QQQ_daily_change_pct', 0)}%)")
        berr = int(bsdf["error"].notna().sum()) if "error" in bsdf.columns else 0
        b4.metric("Stocks", f"{len(bsdf)}" + (f" ({berr} err)" if berr else ""))
        b5.metric("Contract rows", str(len(bcdf)) if not bcdf.empty else "0")

        st.markdown("**Stock setups**")
        st.dataframe(bdisplay, use_container_width=True, hide_index=True)

        bchart = bdisplay.copy()
        if "error" in bchart.columns:
            bchart = bchart[bchart["error"].isna()]
        if not bchart.empty and "setup_score" in bchart.columns and "ticker" in bchart.columns:
            fig_b = px.bar(
                bchart,
                x="ticker",
                y="setup_score",
                color="trade_bias" if "trade_bias" in bchart.columns else None,
                title="Setup score by symbol (BestWiningOptionsv1)",
                color_discrete_map=CHART_BIAS,
            )
            st.plotly_chart(fig_b, use_container_width=True)

        st.markdown("**Best contracts**")
        if bcdf.empty:
            st.write("No contracts passed filters (see `BestWiningOptionsv1.py`).")
        else:
            st.dataframe(bcdf, use_container_width=True, hide_index=True)
            bshow = [
                c
                for c in (
                    "ticker",
                    "trade_bias",
                    "expiry",
                    "strike",
                    "delta",
                    "label",
                    "mid_price",
                    "score",
                )
                if c in bcdf.columns
            ]
            if bshow and len(bcdf) <= 50:
                fig_bc = px.scatter(
                    bcdf.head(40),
                    x="strike" if "strike" in bcdf.columns else bshow[0],
                    y="score" if "score" in bcdf.columns else bshow[-1],
                    color="ticker" if "ticker" in bcdf.columns else None,
                    size="volume" if "volume" in bcdf.columns else None,
                    hover_data=bshow,
                    title="Contract rank score (BestWiningOptionsv1)",
                    color_discrete_sequence=px.colors.qualitative.Plotly,
                )
                st.plotly_chart(fig_bc, use_container_width=True)

        bxname = f"best_wining_options_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.download_button(
            "Download Excel (StockSetups + BestContracts)",
            data=best_wining_results_to_excel_bytes(bdisplay, bcdf),
            file_name=bxname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_best_wining_xlsx",
        )

with tab_flow:
    st.subheader("Money flow")
    st.caption(
        "From **`Flow.py`**: daily money-flow (CMF, OBV slope, rel vol), 5m intraday flow score, and "
        "nearest-expiry **options premium** call vs put; **Composite_Flow_Score** ranks symbols. "
        "Respects **Single ticker** or the sidebar list."
    )

    max_flow = st.slider(
        "Max tickers to scan",
        min_value=1,
        max_value=25,
        value=15,
        key="flow_max_tickers",
    )
    top_flow_n = st.slider(
        "Top N (bullish / bearish tables)",
        min_value=3,
        max_value=25,
        value=10,
        key="flow_top_n",
    )
    run_flow = st.button("Run money flow scan", type="primary", use_container_width=True, key="run_flow_btn")

    if run_flow:
        syms_flow = tickers_for_run(tickers_raw)[:max_flow]
        if not syms_flow:
            st.warning("Enter **Single ticker** or add symbols in the sidebar editor.")
        else:
            try:
                with st.spinner("Flow scan (yfinance — can take a minute)…"):
                    mdf, bdf, brdf, omap = run_money_flow_scan(syms_flow, top_n=top_flow_n)
                    st.session_state.flow_master_df = mdf
                    st.session_state.flow_bull_df = bdf
                    st.session_state.flow_bear_df = brdf
                    st.session_state.flow_options_map = omap
            except Exception as e:  # noqa: BLE001
                st.session_state.flow_master_df = None
                st.session_state.flow_bull_df = None
                st.session_state.flow_bear_df = None
                st.session_state.flow_options_map = None
                st.error(f"Flow scan failed: {e}")

    if st.session_state.flow_master_df is None:
        st.info("Click **Run money flow scan** to load `Flow.py` results.")
    else:
        f_master = st.session_state.flow_master_df
        f_bull = st.session_state.flow_bull_df
        f_bear = st.session_state.flow_bear_df
        f_opt = st.session_state.flow_options_map or {}

        if f_master.empty:
            st.warning(
                "No symbols passed Flow filters (see `Flow.py`: min price, min 20D dollar volume, "
                "enough history)."
            )
        else:
            _single_f = str(st.session_state.get("single_ticker_override") or "").strip()
            if _single_f:
                st.caption(f"**Single:** `{tickers_for_run(tickers_raw)[0]}` — list ignored for this run.")
            else:
                st.caption(
                    f"Universe: **{st.session_state.wl_selector}** "
                    f"(up to {max_flow} symbols scanned)."
                )

            f1, f2, f3 = st.columns(3)
            f1.metric("Symbols ranked", len(f_master))
            f2.metric("Bullish table rows", len(f_bull) if f_bull is not None else 0)
            f3.metric("Bearish table rows", len(f_bear) if f_bear is not None else 0)

            t_all, t_bull, t_bear, t_opt = st.tabs(
                ["All flows", "Top bullish inflow", "Top bearish outflow", "Options detail"]
            )
            with t_all:
                st.dataframe(f_master, use_container_width=True, hide_index=True)
                if not f_master.empty and "Composite_Flow_Score" in f_master.columns:
                    fig_f = px.bar(
                        f_master.sort_values("Composite_Flow_Score", ascending=False).head(20),
                        x="Ticker",
                        y="Composite_Flow_Score",
                        color="Direction",
                        title="Composite flow score (top 20)",
                    )
                    st.plotly_chart(fig_f, use_container_width=True)
            with t_bull:
                if f_bull is None or f_bull.empty:
                    st.write("No rows with composite score > 0.")
                else:
                    st.dataframe(f_bull, use_container_width=True, hide_index=True)
            with t_bear:
                if f_bear is None or f_bear.empty:
                    st.write("No rows with composite score < 0.")
                else:
                    st.dataframe(f_bear, use_container_width=True, hide_index=True)
            with t_opt:
                tickers_with_opt = [t for t, d in f_opt.items() if d is not None and not d.empty]
                if not tickers_with_opt:
                    st.write("No per-ticker option detail tables (chains empty or filtered out).")
                else:
                    pick = st.selectbox("Ticker", options=tickers_with_opt, key="flow_opt_pick")
                    st.dataframe(f_opt[pick], use_container_width=True, hide_index=True)

            fxname = f"money_flow_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            st.download_button(
                "Download Excel (All_Flows + Bullish + Bearish + per-ticker options)",
                data=money_flow_excel_bytes(
                    f_master,
                    f_bull if f_bull is not None else pd.DataFrame(),
                    f_bear if f_bear is not None else pd.DataFrame(),
                    f_opt,
                ),
                file_name=fxname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_flow_xlsx",
            )
