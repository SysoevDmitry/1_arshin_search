@echo off
chcp 65001 >nul 2>&1
REM ФГИС АРШИН - Konsol_Excel - Скрипт запуска для Windows
REM Проверка виртуального окружения и зависимостей

setlocal enabledelayedexpansion

set "SCRIPT_DIR=%~dp0"
set "VENV_DIR=%SCRIPT_DIR%..\venv"
set "VENV_ACTIVATE=%SCRIPT_DIR%..\venv\Scripts\activate.bat"
set "REQUIREMENTS_FILE=%SCRIPT_DIR%..\requirements.txt"
set "PYTHON="

echo ======================================================================
echo 🔍 ФГИС АРШИН - Konsol_Excel v6.1
echo ======================================================================

REM Проверка виртуального окружения
if exist "%VENV_ACTIVATE%" (
    echo ✅ Виртуальное окружение найдено: %VENV_DIR%
    call "%VENV_ACTIVATE%"
    set "PYTHON=python"
) else (
    echo ⚠️  Виртуальное окружение не найдено: %VENV_DIR%
    echo 📦 Попытка использовать системный Python...
    
    where python >nul 2>&1
    if %errorlevel% equ 0 (
        set "PYTHON=python"
    ) else (
        where python3 >nul 2>&1
        if %errorlevel% equ 0 (
            set "PYTHON=python3"
        ) else (
            echo ❌ Python не найден. Установите Python 3.8+
            exit /b 1
        )
    )
    
    REM Проверка зависимостей
    echo 🔍 Проверка зависимостей...
    %PYTHON% -c "import pandas" >nul 2>&1
    if %errorlevel% neq 0 (
        echo ⚠️  Зависимости не установлены.
        echo 📦 Установка: pip install pandas openpyxl aiohttp tqdm
        exit /b 1
    )
)

REM Проверка зависимостей в venv
echo 🔍 Проверка зависимостей...
set "MISSING_DEPS="

for %%d in (pandas openpyxl aiohttp tqdm) do (
    %PYTHON% -c "import %%d" >nul 2>&1
    if %errorlevel% neq 0 (
        if "!MISSING_DEPS!"=="" (
            set "MISSING_DEPS=%%d"
        ) else (
            set "MISSING_DEPS=!MISSING_DEPS! %%d"
        )
    )
)

if not "!MISSING_DEPS!"=="" (
    echo ⚠️  Отсутствуют зависимости: !MISSING_DEPS!
    echo 📦 Установка зависимостей...
    
    if exist "%REQUIREMENTS_FILE%" (
        %PYTHON% -m pip install -r "%REQUIREMENTS_FILE%" --quiet
    ) else (
        %PYTHON% -m pip install pandas openpyxl aiohttp tqdm --quiet
    )
    
    echo ✅ Зависимости установлены
) else (
    echo ✅ Все зависимости установлены
)

REM Запуск приложения
echo 🚀 Запуск приложения...
echo ======================================================================

cd /d "%SCRIPT_DIR%"
%PYTHON% main.py %*
