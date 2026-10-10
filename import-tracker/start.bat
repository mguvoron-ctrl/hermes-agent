@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Import Tracker

echo ============================================
echo    Import Tracker - starting
echo ============================================
echo.

rem --- find Python (prefer stable 3.13 / 3.12) ---
set "PY="
py -3.13 --version >nul 2>nul && set "PY=py -3.13"
if not defined PY ( py -3.12 --version >nul 2>nul && set "PY=py -3.12" )
if not defined PY ( where py >nul 2>nul && set "PY=py" )
if not defined PY ( where python >nul 2>nul && set "PY=python" )
if not defined PY (
  echo [!] Python not found.
  echo     Install it from https://www.python.org/downloads/
  echo     then run start.bat again.
  pause
  exit /b 1
)

rem --- first run: create venv and install dependencies ---
if not exist ".venv\Scripts\activate.bat" (
  echo First run: setting up, please wait a minute...
  %PY% -m venv .venv
  call ".venv\Scripts\activate.bat"
  python -m pip install --upgrade pip
  pip install -r requirements.txt
) else (
  call ".venv\Scripts\activate.bat"
)

rem --- start the server ---
echo.
echo Open in your browser:  http://127.0.0.1:8000
echo Login: admin / admin
echo (to stop the server - just close this window)
echo.
python -m uvicorn app:app --host 0.0.0.0 --port 8000
echo.
echo Server stopped.
pause
