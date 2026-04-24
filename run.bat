@echo off
cd /d %~dp0
if not exist .venv (
  echo Creating virtual environment...
  py -3 -m venv .venv 2>nul || python -m venv .venv
)
if not exist .venv\Scripts\python.exe (
  echo Virtual environment creation failed.
  pause
  exit /b 1
)
echo Installing requirements...
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo Package install failed.
  echo If reportlab fails on your PC, tell me and I will give you a no-PDF version.
  pause
  exit /b 1
)
echo Starting server at http://127.0.0.1:8000
start http://127.0.0.1:8000
.venv\Scripts\python.exe app_flask.py
pause
