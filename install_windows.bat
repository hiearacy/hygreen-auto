@echo off
cd /d "%~dp0"
py -3 -m pip install -r requirements.txt
if errorlevel 1 (
  echo Failed. Check that Python 3 is installed.
  pause
  exit /b 1
)
echo Installation complete.
pause
