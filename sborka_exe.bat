@echo off
chcp 1251 >nul
cd /d "%~dp0"

echo.
echo === Sborka AutoMaticaSMK v EXE ===
echo.

if not exist "main_gui.py" goto net_file

set "PY=python"
if exist ".venv\Scripts\python.exe" set "PY=.venv\Scripts\python.exe"
echo Python: %PY%

echo.
echo --- Ustanovka zavisimostey ---
"%PY%" -m pip install --upgrade pip
"%PY%" -m pip install pyinstaller pandas openpyxl
if errorlevel 1 goto err_dep

if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "main_gui.spec" del /q "main_gui.spec"

echo.
echo --- Sborka EXE (paru minut) ---
"%PY%" -m PyInstaller --noconfirm --onefile --windowed --name AutoMaticaSMK --hidden-import pandas --hidden-import openpyxl --collect-submodules pandas main_gui.py
if errorlevel 1 goto err_build

if exist "baza" xcopy /e /i /y "baza" "dist\baza" >nul
if exist "zadaniya" xcopy /e /i /y "zadaniya" "dist\zadaniya" >nul

echo.
echo ============================================================
echo  GOTOVO. EXE zdes:  dist\AutoMaticaSMK.exe
echo  Ryadom s exe dolzhna lezhat papka baza s tovary.csv
echo ============================================================
echo.
pause
exit /b 0

:net_file
echo [OSHIBKA] Net main_gui.py ryadom s etim bat-faylom.
pause
exit /b 1

:err_dep
echo [OSHIBKA] Ne udalos ustanovit zavisimosti.
pause
exit /b 1

:err_build
echo [OSHIBKA] Sborka ne udalas. Smotrite soobsheniya vyshe.
pause
exit /b 1
