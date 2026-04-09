"""Plotly + Streamlit styling (Robinhood-inspired: green up, red down)."""
from __future__ import annotations

import plotly.io as pio
import streamlit as st

# Robinhood-style brand approximations (not affiliated)
RH_GREEN = "#00C805"
RH_RED = "#FF5000"
RH_MUTED = "#6B7280"

_ROBINHOOD_PRIMARY_CSS = """
<style>
/* White label on Robinhood-green primary buttons */
.stButton > button[kind="primary"] {
  color: #FFFFFF !important;
}
</style>
"""

# Plotly scales (green up / red down)
CHART_SCALE_GAIN = [[0, "#E8FCE9"], [0.5, "#7FE57F"], [1, RH_GREEN]]
CHART_SCALE_LOSS = [[0, "#FFF0EB"], [0.5, "#FF9B7A"], [1, RH_RED]]
CHART_SCALE_SPIKE = [[0, "#F6F7F8"], [1, RH_GREEN]]
CHART_BIAS = {"CALL": RH_GREEN, "PUT": RH_RED, "NEUTRAL": RH_MUTED}
CHART_PREMIUM = {"Call": RH_GREEN, "Put": RH_RED}


def apply_trading_theme() -> None:
    pio.templates.default = "plotly_white"
    st.markdown(_ROBINHOOD_PRIMARY_CSS, unsafe_allow_html=True)
