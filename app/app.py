"""
Pharma Launch Forecast – Multi-Page Application
================================================
Three use cases accessible via sidebar navigation.
"""

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import streamlit as st

st.set_page_config(
    page_title="Pharma Launch Forecast",
    page_icon="\U0001f48a",
    layout="wide",
    initial_sidebar_state="expanded",
)

from main import show as eliquis_show
from glp1 import show as glp1_show
from rx_otc import show as rx_otc_show

# ─── Pages ──────────────────────────────────────────────────────────────
eliquis_page = st.Page(eliquis_show, title="Eliquis Generic Entry", icon="\U0001f4c9", url_path="eliquis", default=True)
glp1_page = st.Page(glp1_show, title="GLP-1 Brand Competition", icon="\U0001f489", url_path="glp1")
rx_otc_page = st.Page(rx_otc_show, title="Rx-to-OTC Switch", icon="\U0001f6d2", url_path="rx-otc")

nav = st.navigation(
    {
        "Launch Forecasts": [eliquis_page, glp1_page, rx_otc_page],
    }
)
nav.run()
