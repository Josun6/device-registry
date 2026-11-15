@echo off
cd /d "%~dp0"
py -3 device_registry.py
if errorlevel 1 (
  echo.
  echo Could not start with py -3. Trying python...
  python device_registry.py
)
pause
