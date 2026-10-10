@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Трекер поставок
echo ============================================
echo    Трекер поставок — запуск
echo ============================================
echo.

rem --- 1) Ищем Python ---
set "PY="
where py >nul 2>nul && set "PY=py"
if not defined PY (
  where python >nul 2>nul && set "PY=python"
)
if not defined PY (
  echo [!] Python не найден на этом компьютере.
  echo.
  echo     1^) Скачай Python: https://www.python.org/downloads/
  echo     2^) При установке ПОСТАВЬ галочку "Add Python to PATH"
  echo     3^) Запусти этот файл start.bat снова
  echo.
  pause
  exit /b 1
)

rem --- 2) Первый запуск: создаём окружение и ставим зависимости ---
if not exist ".venv\Scripts\activate.bat" (
  echo Первый запуск: настраиваю окружение, подожди минутку...
  %PY% -m venv .venv
  call ".venv\Scripts\activate.bat"
  python -m pip install --upgrade pip
  pip install -r requirements.txt
) else (
  call ".venv\Scripts\activate.bat"
)

rem --- 3) Запуск сервера ---
echo.
python -m uvicorn app:app --host 0.0.0.0 --port 8000
echo.
echo Сервер остановлен.
pause
