@echo off
setlocal

set GCLOUD=gcloud
set SRC=%~dp0cloud_function

echo === Desplegando Bot Telegram en Cloud Functions ===

%GCLOUD% functions deploy val-semanal-bot ^
  --gen2=false ^
  --runtime python312 ^
  --trigger-http ^
  --allow-unauthenticated ^
  --entry-point webhook ^
  --source "%SRC%" ^
  --set-env-vars "TELEGRAM_BOT_TOKEN=%TELEGRAM_BOT_TOKEN%,GITHUB_RAW_URL=%GITHUB_RAW_URL%" ^
  --region us-central1 ^
  --memory 256MB ^
  --timeout 30s

echo.
echo === Deploy completado ===
pause
