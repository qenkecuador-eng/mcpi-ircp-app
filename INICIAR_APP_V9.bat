@echo off
title MCPI-IRCP-I v9 ampliada
cd /d "%~dp0"
py -m pip install -r requirements.txt
py -m streamlit run app.py
pause
