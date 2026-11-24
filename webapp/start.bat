@echo off
rem Starts the Grid agent web app. First time: copy .env.example to .env and choose a password (GRID_APP_PASSWORD) in it.
cd /d "%~dp0"
where node >nul 2>nul || (echo Node.js 24 or newer is needed: https://nodejs.org & pause & exit /b 1)
if not exist node_modules (call npm install --omit=dev || (pause & exit /b 1))
if not exist .env if "%GRID_APP_PASSWORD%"=="" if not "%GRID_APP_INSECURE%"=="1" (
  echo No password is set. Copy .env.example to .env and put a password of 8 or more characters after GRID_APP_PASSWORD=
  pause & exit /b 1
)
echo Starting on http://localhost:3000  (Ctrl+C to stop)
call npm start
pause
