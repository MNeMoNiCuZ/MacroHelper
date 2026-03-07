@echo off
setlocal
cd /d "%~dp0"

if not exist venv\Scripts\python.exe (
  echo [INFO] Creating virtual environment...
  py -m venv venv
)

call venv\Scripts\activate.bat

echo [INFO] Installing requirements...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller

echo [INFO] Building executable...
pyinstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --name sc2_alarm_app ^
  --icon icon.ico ^
  --add-data "icon.ico;." ^
  app.py

if errorlevel 1 (
  echo [ERROR] Build failed.
  exit /b 1
)

echo [OK] Build complete: dist\MacroHelper.exe
endlocal
