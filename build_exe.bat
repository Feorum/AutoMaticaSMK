@echo off
chcp 65001 >nul
echo ============================================================
echo  Сборка main_gui_v4.exe + ocr_scanner.exe
echo ============================================================

:: ── Проверка Python ──────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo ОШИБКА: python не найден в PATH.
    echo Установите Python 3.10+ и добавьте в PATH.
    pause & exit /b 1
)

:: ── Установка зависимостей ───────────────────────────────────
echo.
echo [1/4] Установка зависимостей...
pip install pyinstaller pandas openpyxl pytesseract pillow --quiet
if errorlevel 1 (
    echo ОШИБКА при установке зависимостей.
    pause & exit /b 1
)

:: ── Создание папки dist если нет ────────────────────────────
if not exist dist mkdir dist

:: ── Сборка main_gui_v4.exe ───────────────────────────────────
echo.
echo [2/4] Сборка main_gui_v4.exe...
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "main_gui_v4" ^
    --add-data "ocr_scanner.py;." ^
    --hidden-import "pytesseract" ^
    --hidden-import "PIL" ^
    --hidden-import "PIL.Image" ^
    --hidden-import "PIL.ImageFilter" ^
    --hidden-import "PIL.ImageEnhance" ^
    --hidden-import "pandas" ^
    --hidden-import "openpyxl" ^
    --hidden-import "csv" ^
    --hidden-import "tkinter" ^
    --hidden-import "tkinter.ttk" ^
    --hidden-import "tkinter.filedialog" ^
    --hidden-import "tkinter.messagebox" ^
    --collect-submodules "PIL" ^
    main_gui_v4.py

if errorlevel 1 (
    echo ОШИБКА при сборке main_gui_v4.exe
    pause & exit /b 1
)
echo OK: dist\main_gui_v4.exe

:: ── Сборка ocr_scanner.exe (CLI-утилита) ─────────────────────
echo.
echo [3/4] Сборка ocr_scanner.exe (CLI)...
pyinstaller ^
    --onefile ^
    --console ^
    --name "ocr_scanner" ^
    --hidden-import "pytesseract" ^
    --hidden-import "PIL" ^
    --hidden-import "PIL.Image" ^
    --hidden-import "PIL.ImageFilter" ^
    --hidden-import "PIL.ImageEnhance" ^
    --collect-submodules "PIL" ^
    ocr_scanner.py

if errorlevel 1 (
    echo ОШИБКА при сборке ocr_scanner.exe
    pause & exit /b 1
)
echo OK: dist\ocr_scanner.exe

:: ── Копирование папки baza рядом с exe ───────────────────────
echo.
echo [4/4] Подготовка папки dist...
if exist baza (
    xcopy baza dist\baza /E /I /Y /Q
    echo OK: папка baza скопирована в dist\
) else (
    echo ВНИМАНИЕ: папка baza не найдена рядом с батником.
    echo Создайте dist\baza\ вручную и положите туда tovary.csv и matching.csv.
)

if exist zakazy (
    xcopy zakazy dist\zakazy /E /I /Y /Q
)
if exist zadaniya (
    xcopy zadaniya dist\zadaniya /E /I /Y /Q
)

:: ── Результат ────────────────────────────────────────────────
echo.
echo ============================================================
echo  ГОТОВО!
echo.
echo  dist\main_gui_v4.exe  — основная программа (двойной клик)
echo  dist\ocr_scanner.exe  — OCR-утилита (командная строка)
echo.
echo  Структура рядом с exe:
echo    baza\tovary.csv
echo    baza\matching.csv
echo    baza\istoriya.csv  (создаётся автоматически)
echo    zakazy\            (сохранённые заказы)
echo    zadaniya\          (входящие задания)
echo.
echo  ВАЖНО: Tesseract OCR нужен на рабочем компьютере отдельно!
echo  Скачать: https://github.com/UB-Mannheim/tesseract/wiki
echo  При установке выбрать: Russian language pack
echo  После установки путь будет:
echo    C:\Program Files\Tesseract-OCR\tesseract.exe
echo ============================================================
pause
