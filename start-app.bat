@echo off
setlocal
cd /d "%~dp0"
title Nano Banana PPT Studio

if not "%~1"=="" (
  set "APP_PORT=%~1"
) else (
  if defined PORT (
    set "APP_PORT=%PORT%"
  ) else (
    set "APP_PORT=3000"
  )
)

set "REQUESTED_PORT=%APP_PORT%"
for /f %%P in ('powershell -NoProfile -ExecutionPolicy Bypass -Command "$port = [int]$env:REQUESTED_PORT; while (Get-NetTCPConnection -LocalPort $port -ErrorAction SilentlyContinue) { $port++ }; Write-Output $port"') do set "APP_PORT=%%P"
set "PORT=%APP_PORT%"

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

if not "%APP_PORT%"=="%REQUESTED_PORT%" (
  echo [INFO] Port %REQUESTED_PORT% is already in use. Switched to %APP_PORT%.
)

echo [2/3] Opening browser...
start "" powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Sleep -Seconds 2; Start-Process 'http://localhost:%APP_PORT%'"

echo [3/3] Starting local server on port %APP_PORT%...
call npm start

endlocal
