@echo off
setlocal
cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
  echo Node.js is required to run the localhost version.
  echo Install Node.js from https://nodejs.org and run this file again.
  pause
  exit /b 1
)

if not exist "node_modules\vite\bin\vite.js" (
  echo Preparing LIC Policy Tracker for first use...
  call npm install
  if errorlevel 1 (
    echo Setup failed. Check your internet connection and try again.
    pause
    exit /b 1
  )
)

echo Starting LIC Policy Tracker at http://localhost:5173
start "" powershell.exe -NoProfile -WindowStyle Hidden -Command "Start-Sleep -Seconds 2; Start-Process 'http://localhost:5173'"
call npm run dev -- --host localhost --port 5173 --strictPort

echo.
echo The localhost server has stopped.
pause
