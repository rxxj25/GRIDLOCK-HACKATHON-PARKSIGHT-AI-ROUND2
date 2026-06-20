@echo off
setlocal

cd /d "%~dp0"

echo.
echo Starting ParkSight AI...
echo Project: %CD%
echo.

where node >nul 2>nul
if errorlevel 1 (
  echo ERROR: Node.js is not installed or not available in PATH.
  echo Install Node.js, then run this file again.
  pause
  exit /b 1
)

where npm >nul 2>nul
if errorlevel 1 (
  echo ERROR: npm is not installed or not available in PATH.
  pause
  exit /b 1
)

if not exist "node_modules" (
  echo Installing dependencies...
  call npm install
  if errorlevel 1 (
    echo ERROR: npm install failed.
    pause
    exit /b 1
  )
)

echo Checking PostgreSQL schema...
call npm run db:setup
if errorlevel 1 (
  echo.
  echo WARNING: Database setup failed. The app can still open, but DB-backed login/upload storage may not work.
  echo Check PostgreSQL and .env.local if this keeps happening.
  echo.
)

echo Opening http://localhost:8000 ...
start "" "http://localhost:8000"

echo.
echo Starting API and Vite dev server.
echo Keep this window open while using the app.
echo Press Ctrl+C to stop.
echo.

call npm run dev -- --port 8000

endlocal
