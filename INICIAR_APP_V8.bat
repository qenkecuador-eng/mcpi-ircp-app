@echo off
title MCPI-IRCP-I v8 motor judicial
cd /d "%~dp0"
py -m pip install -r requirements.txt
py -m streamlit run app.py
pause
