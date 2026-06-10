@echo off
:: La siguiente línea fuerza al CMD a situarse en la carpeta donde está el .bat
cd /d "%~dp0"

title Robot Tagetik - Ejecutando...
echo ===========================================
echo   INICIANDO AUTOMATIZACION DE REPORTES
echo ===========================================
echo.
"C:\Program Files\Python313\python.exe" "C:\Users\bmorenoa\Desktop\Reports Automation\main.py"
echo.
echo ===========================================
echo   PROCESO FINALIZADO
echo ===========================================
pause