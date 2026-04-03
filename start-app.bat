@echo off
setlocal
cd /d "%~dp0"
title Wan 2.7 Image Studio

if not exist "generated-images" mkdir "generated-images"
if not exist ".npm-cache" mkdir ".npm-cache"

where node >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Node.js is not installed or not in PATH.
  pause
  exit /b 1
)

where npm >nul 2>nul
if errorlevel 1 (
  echo [ERROR] npm is not installed or not in PATH.
  pause
  exit /b 1
)

if not exist "node_modules\express" (
  echo [1/3] Installing dependencies...
  call npm install --cache .npm-cache
  if errorlevel 1 (
    echo [ERROR] npm install failed.
    pause
    exit /b 1
  )
) else (
  echo [1/3] Dependencies already installed.
)

echo [2/3] Opening browser...
start "" powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Sleep -Seconds 2; Start-Process 'http://localhost:3000'"

echo [3/3] Starting local server...
call npm start

endlocal
