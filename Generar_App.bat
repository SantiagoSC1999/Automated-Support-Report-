@echo off
echo ==================================================
echo      APP - GENERADOR DE REPORTES MENSUALES
echo ==================================================
echo.
echo Abriendo la app en tu navegador...
echo.

call .venv\Scripts\activate.bat
streamlit run app.py

pause
