@echo off
REM ------------------------------------------------------------------
REM  Builds BulkFileGenerator.exe (Windows). Run this from the folder
REM  that contains BulkFileGenerator.py, bulk_engine.py, ui.html,
REM  make_bundle.py and config.json. Needs Python 3.9+ on PATH.
REM
REM  The project list (config.json) and the page (ui.html) are packed
REM  INSIDE the .exe, so people using it see no code or settings files.
REM  To change the projects: edit config.json here and run this again.
REM ------------------------------------------------------------------
cd /d "%~dp0"

echo Installing build requirements...
python -m pip install --upgrade -r requirements.txt pyinstaller || goto :fail

echo Packing the project list and the page into the app...
python make_bundle.py || goto :fail

echo Building the .exe...
python -m PyInstaller --noconfirm --clean --onefile ^
  --name BulkFileGenerator ^
  --hidden-import _bundled ^
  BulkFileGenerator.py || goto :fail
del /q _bundled.py 2>nul
if exist dist\config.json del /q dist\config.json

echo.
echo Done! Your app is at:  %~dp0dist\BulkFileGenerator.exe
echo Share only the .exe. It creates Tracker\ and Output\ next to itself on first run.
pause
exit /b 0

:fail
del /q _bundled.py 2>nul
echo.
echo Build failed - see the messages above.
pause
exit /b 1
