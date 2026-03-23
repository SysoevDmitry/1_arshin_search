@echo off
chcp 65001 >nul
REM Скрипт сборки портативной версии ФГИС АРШИН для Windows
REM Использование: build_portable.bat

echo ========================================
echo Сборка портативной версии: ФГИС АРШИН v2.6
echo ========================================

REM Переход в директорию проекта
cd /d "%~dp0\.."

REM Проверка Python
python --version >nul 2>&1
if errorlevel 1 (
    echo Ошибка: Python не найден
    echo Установите Python 3.9+ с https://python.org
    pause
    exit /b 1
)

echo [1/5] Проверка Python...
python --version

REM Создание venv если нет
if not exist "venv" (
    echo [2/5] Создание виртуального окружения...
    python -m venv venv
)

call venv\Scripts\activate.bat

REM Установка зависимостей
echo [3/5] Установка зависимостей...
pip install --upgrade pip -q
pip install -r requirements.txt -q
pip install pyinstaller -q

REM Создание spec файла
echo [4/5] Создание спецификации...
(
echo from PyInstaller.utils.hooks import collect_submodules
echo.
echo a = Analysis(
echo     ['arshin_app.py'],
echo     pathex=[],
echo     binaries=[],
echo     datas=[
echo         ('config', 'config'),
echo         ('requirements.txt', 'requirements.txt'),
echo     ],
echo     hiddenimports=[
echo         'tkinter',
echo         'pandas',
echo         'openpyxl',
echo         'aiohttp',
echo         'requests',
echo         'sqlite3',
echo     ] + collect_submodules('pandas'),
echo     hookspath=[],
echo     runtime_hooks=[],
echo     excludes=[],
echo     win_no_prefer_redirects=False,
echo     win_private_assemblies=False,
echo     cipher=None,
echo     noarchive=False,
echo ^)
echo.
echo pyz = PYZ(a.pure, a.zipped_data, cipher=None)
echo.
echo exe = EXE(
echo     pyz,
echo     a.scripts,
echo     a.binaries,
echo     a.zipfiles,
echo     a.datas,
echo     [],
echo     name='FGIS_ARSHIN',
echo     debug=False,
echo     bootloader_ignore_signals=False,
echo     strip=False,
echo     upx=True,
echo     console=False,
echo     icon='arshin.ico',
echo ^)
) > arshin.spec

REM Сборка EXE
echo [5/5] Сборка EXE файла...
echo Это может занять несколько минут...
pyinstaller arshin.spec

REM Создание портативной папки
echo.
echo Создание портативной версии...
if not exist "App_Windows\ARSHIN_Portable" mkdir App_Windows\ARSHIN_Portable

if exist "dist\FGIS_ARSHIN.exe" (
    copy /Y "dist\FGIS_ARSHIN.exe" "App_Windows\ARSHIN_Portable\"
    copy /Y "run_arshin.bat" "App_Windows\ARSHIN_Portable\"
    
    if not exist "App_Windows\ARSHIN_Portable\config" mkdir App_Windows\ARSHIN_Portable\config
    copy /Y "config\*.csv" "App_Windows\ARSHIN_Portable\config\" 2>nul
    
    echo.
    echo ========================================
    echo ✅ Сборка завершена!
    echo ========================================
    echo.
    echo Портативная версия: App_Windows\ARSHIN_Portable\
    echo EXE файл: dist\FGIS_ARSHIN.exe
    echo.
    echo Для запуска:
    echo   App_Windows\ARSHIN_Portable\run_arshin.bat
    echo   или
    echo   dist\FGIS_ARSHIN.exe
    echo.
) else (
    echo.
    echo ⚠️  Ошибка сборки! Проверьте логи выше.
    echo.
)

pause
