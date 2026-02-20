try:
    import pandas as pd
except ModuleNotFoundError as exc:
    raise SystemExit("Missing dependency: pandas. Install with `pip install pandas openpyxl rapidfuzz`.") from exc

try:
    import streamlit as st
except ModuleNotFoundError as exc:
    raise SystemExit("Missing dependency: streamlit. Install with `pip install streamlit`.") from exc

from excel_merger.app_page import render_app
from excel_merger.state import ensure_streamlit_context


if __name__ == "__main__":
    ensure_streamlit_context()

render_app()
