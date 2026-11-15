@echo off
setlocal
cd /d "%~dp0"
set "ENV=.venv-build-standard"
if not exist "%ENV%\Scripts\python.exe" (
    py -3 -m venv "%ENV%" || exit /b 1
)
"%ENV%\Scripts\python.exe" -m pip install -r requirements-build-standard.txt || exit /b 1
"%ENV%\Scripts\python.exe" -m PyInstaller --noconfirm --clean --onefile --windowed --optimize 2 ^
    --name DeviceRegistry --icon "assets\icon.ico" ^
    --add-data "assets\icon.png;assets" --add-data "assets\icon.ico;assets" ^
    device_registry.py
if errorlevel 1 exit /b 1
echo Standard build ready: dist\DeviceRegistry.exe
endlocal
pause
