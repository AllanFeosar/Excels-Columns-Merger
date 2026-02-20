@echo off
set STREAMLIT_BROWSER_GATHER_USAGE_STATS=false
python -m streamlit run app.py %*
