"""Default Plotly template for the trading dashboard."""
from __future__ import annotations

import plotly.io as pio
import streamlit as st

# Matches `.streamlit/config.toml` [theme] primaryColor (#86efac) — light green highlights

_LIGHT_GREEN_HIGHLIGHT_CSS = """
<style>
/* Dark green text on light-green primary buttons (readable on #86efac) */
.stButton > button[kind="primary"] {
  color: #14532d !important;
}
</style>
"""

# Simple chart palettes (Plotly-friendly)
CHART_SCALE_GAIN = "Greens"
CHART_SCALE_LOSS = "Greens"
CHART_SCALE_SPIKE = "Blues"
CHART_BIAS = {"CALL": "#2ca02c", "PUT": "#15803d", "NEUTRAL": "#7f7f7f"}
CHART_PREMIUM = {"Call": "#1f77b4", "Put": "#059669"}


def apply_trading_theme() -> None:
    pio.templates.default = "plotly"
    st.markdown(_LIGHT_GREEN_HIGHLIGHT_CSS, unsafe_allow_html=True)
