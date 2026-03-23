@echo off
chcp 65001 >nul
REM Скрипт запуска ФГИС АРШИН для Windows
REM Использование: run_arshin.bat

cd /d "%~dp0"

echo ========================================
echo ФГИС АРШИН - Запуск
echo ========================================

REM Проверка Python
python --version >nul 2>&1
if errorlevel 1 (
    echo Ошибка: Python не найден
    echo Установите Python 3.9+ с https://python.org
    pause
    exit /b 1
)

echo ✓ Python: 
python --version

REM Создание venv если нет
if not exist "venv" (
    echo.
    echo Создание виртуального окружения...
    python -m venv venv
)

REM Активация
call venv\Scripts\activate.bat

REM Установка зависимостей
echo.
echo Проверка зависимостей...
pip install -q -r requirements.txt

REM Запуск приложения
echo.
echo Запуск приложения...
echo ========================================
python arshin_app.py

REM Деактивация
deactivate 2>nul

echo.
echo Приложение закрыто.
pause
