@echo off
title Sborka EXE (v0.5)
cd /d "%~dp0"

echo ============================================================
echo   Sborka auto_vvod.exe  (v0.5: cikl + baza + istoriya)
echo ============================================================
echo.
echo Tekushaya papka: %CD%
echo.

REM --- [1/5] Proverka Python ---
echo [1/5] Proverka Python...
python --version
if %errorlevel% neq 0 goto NO_PYTHON
echo OK
echo.

REM --- [2/5] Poisk skripta auto_vvod*.py ---
echo [2/5] Poisk skripta auto_vvod*.py...
set "SCRIPT="
for %%F in (auto_vvod*.py) do if not defined SCRIPT set "SCRIPT=%%F"
if not defined SCRIPT goto NO_SCRIPT
echo Nayden: %SCRIPT%
if not exist "oformlenie_slovar.py" goto NO_MODULE
echo Modul oformlenie_slovar.py - na meste.
echo.

REM --- [3/5] Zavisimosti ---
echo [3/5] Ustanovka pyinstaller, pyautogui, thefuzz...
python -m pip install --upgrade pyinstaller pyautogui thefuzz
if %errorlevel% neq 0 goto NO_DEPS
echo.

REM --- [4/5] Sborka exe ---
REM oformlenie_slovar.py podtyanetsya avtomaticheski (import v skripte).
REM thefuzz dobavlyaem v hidden-import na vsyakiy sluchay.
echo [4/5] Sborka exe...
python -m PyInstaller --onefile --console --clean --name auto_vvod ^
  --hidden-import thefuzz --hidden-import oformlenie_slovar "%SCRIPT%"
if %errorlevel% neq 0 goto NO_BUILD
echo.

REM --- [5/5] Podgotovka papok ryadom s exe (v papke dist) ---
echo [5/5] Podgotovka struktury v papke dist...
if not exist "dist\baza" mkdir "dist\baza"
if not exist "dist\zadaniya" mkdir "dist\zadaniya"
REM Kopiruem bazu i zadaniya, esli oni est ryadom s batnikom
if exist "baza\tovary.csv" copy /Y "baza\tovary.csv" "dist\baza\tovary.csv" >nul
if exist "baza\istoriya.csv" copy /Y "baza\istoriya.csv" "dist\baza\istoriya.csv" >nul
if exist "zadaniya\*.txt" copy /Y "zadaniya\*.txt" "dist\zadaniya\" >nul

echo.
echo ============================================================
echo   GOTOVO!
echo ============================================================
echo.
echo V papke dist sobrana gotovaya struktura dlya fleshki:
echo    dist\auto_vvod.exe
echo    dist\baza\tovary.csv
echo    dist\zadaniya\
echo.
echo Skopiruyte VSYU papku dist na fleshku (ili eyo soderzhimoe).
echo Vazhno: baza i zadaniya dolzhny lezhat RYADOM s exe.
echo.
if not exist "dist\baza\tovary.csv" echo [VNIMANIE] Net baza\tovary.csv - sozdayte ego do zapuska exe!
goto END

:NO_PYTHON
echo.
echo [ERROR] Python ne nayden ili oshibka zapuska.
echo Ustanovite Python s python.org, galochka "Add Python to PATH".
goto END

:NO_SCRIPT
echo.
echo [ERROR] Ryadom net fayla auto_vvod*.py
echo Polozhite etot .bat v tu zhe papku, gde skript.
goto END

:NO_MODULE
echo.
echo [ERROR] Ryadom net fayla oformlenie_slovar.py
echo On nuzhen dlya raboty poiska. Polozhite ego ryadom s auto_vvod.py.
goto END

:NO_DEPS
echo.
echo [ERROR] Ne udalos ustanovit zavisimosti. Proverte internet.
goto END

:NO_BUILD
echo.
echo [ERROR] Sborka ne udalas (smotrite soobsheniya vyshe).
goto END

:END
echo.
echo --- Okno ne zakroetsya. Nazhmite lyubuyu klavishu. ---
pause >nul
