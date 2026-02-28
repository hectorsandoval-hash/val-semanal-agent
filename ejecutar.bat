@echo off
REM =============================================
REM EJECUCION - AGENTE VALORIZACION SEMANAL
REM Pipeline: Gmail -> Excel -> Reporte -> Drive -> Telegram
REM =============================================

cd /d "%~dp0"

echo [%date% %time%] Iniciando val-semanal-agent... >> logs\ejecucion.log

REM Ejecutar pipeline completo
python main.py >> logs\ejecucion.log 2>&1

echo [%date% %time%] Ejecucion completada. >> logs\ejecucion.log
echo ------------------------------------------ >> logs\ejecucion.log
