@echo off
REM One-time setup: makes a private Python environment (.venv) in this folder and installs what the scripts need.
REM Run it again if you ever move this folder (the environment remembers its old location).
cd /d "%~dp0"
set PY=
for %%P in ("%LocalAppData%\Python\pythoncore-3.14-64\python.exe" "%LocalAppData%\Programs\Python\Python314\python.exe" "%LocalAppData%\Programs\Python\Python313\python.exe" "%LocalAppData%\Programs\Python\Python312\python.exe") do (
  if exist %%P if not defined PY set PY=%%~P
)
if not defined PY (
  where python >nul 2>nul && set PY=python
)
if not defined PY (
  echo Could not find Python. Install it from https://www.python.org/downloads/ and run this again.
  exit /b 1
)
echo Using %PY%
if exist .venv rmdir /s /q .venv
"%PY%" -m venv .venv || exit /b 1
".venv\Scripts\python.exe" -m pip install --upgrade pip || exit /b 1
".venv\Scripts\python.exe" -m pip install -r requirements.txt || exit /b 1
echo.
echo Done. In VS Code open this folder, press F5 and pick "Liberty" or "Magma".
