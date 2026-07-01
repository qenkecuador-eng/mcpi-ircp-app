@echo off
title MCPI-IRCP-I Matriz
cd /d "%~dp0"
echo Instalando librerias necesarias...
py -m pip install -r requirements.txt
if %errorlevel% neq 0 (
    python -m pip install -r requirements.txt
)
echo Abriendo app en modo matriz...
echo Si aparece Email, presione ENTER sin escribir nada.
echo Si no abre el navegador, use http://localhost:8501
py -m streamlit run app.py
if %errorlevel% neq 0 (
    python -m streamlit run app.py
)
pause
