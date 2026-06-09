@echo off
echo ==================================================
echo         GENERADOR DE REPORTES MENSUALES
echo ==================================================
echo.
echo Iniciando el procesamiento del Excel...
echo.

call .venv\Scripts\activate.bat 
python fromexceltoword.py

echo.
echo ==================================================
echo Proceso finalizado. 
pause
