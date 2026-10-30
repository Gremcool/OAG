import streamlit as st

import pbc

st.set_page_config(
    page_title="OAG Preliminary Information Request",
    page_icon="📑",
    layout="wide",
    initial_sidebar_state="collapsed",
)

pbc.render_page()
