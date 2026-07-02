@echo off
title MCPI-IRCP-I v5 editable
cd /d "%~dp0"
py -m pip install -r requirements.txt
if %errorlevel% neq 0 (
    python -m pip install -r requirements.txt
)
py -m streamlit run app.py
if %errorlevel% neq 0 (
    python -m streamlit run app.py
)
pause
