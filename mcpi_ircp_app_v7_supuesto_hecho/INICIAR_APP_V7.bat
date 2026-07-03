@echo off
title MCPI-IRCP-I v7 supuesto de hecho
cd /d "%~dp0"
py -m pip install -r requirements.txt
py -m streamlit run app.py
pause
