#!/bin/bash
# ФГИС АРШИН - Скрипт запуска GUI приложения для Armbian (ARM)
# Использует системный Python, без виртуального окружения
# Тяжёлые пакеты ставит через apt (предсобранные для ARM), лёгкие — через pip

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PYTHON=""

echo "======================================================================"
echo "🔍 ФГИС АРШИН - GUI Приложение v3.0 (Armbian/ARM, системный Python)"
echo "======================================================================"

# Проверка архитектуры
ARCH=$(uname -m)
echo "🖥  Архитектура: $ARCH"

# Проверка свободного места (минимум 300 МБ)
AVAIL_KB=$(df --output=avail "$SCRIPT_DIR" 2>/dev/null | tail -1)
AVAIL_MB=$((AVAIL_KB / 1024))
echo "💾 Свободно на диске: ${AVAIL_MB} МБ"
if [ "$AVAIL_MB" -lt 200 ]; then
    echo "❌ Недостаточно свободного места (нужно >200 МБ). Освободите место и повторите."
    exit 1
fi

# Установка системных зависимостей через apt
echo "📦 Проверка и установка системных зависимостей..."
APT_DEPS="python3-pip python3-tk"

APT_PY_DEPS=()

if ! python3 -c "import pandas" 2>/dev/null; then
    APT_PY_DEPS+=("python3-pandas")
fi

if ! python3 -c "import openpyxl" 2>/dev/null; then
    APT_PY_DEPS+=("python3-openpyxl")
fi

if ! python3 -c "import requests" 2>/dev/null; then
    APT_PY_DEPS+=("python3-requests")
fi

if ! python3 -c "import numpy" 2>/dev/null; then
    APT_PY_DEPS+=("python3-numpy")
fi

ALL_APT="$APT_DEPS ${APT_PY_DEPS[*]}"
if [ -n "$APT_PY_DEPS" ] || ! dpkg -s python3-pip python3-tk &>/dev/null; then
    echo "⚠️  Установка через apt: $ALL_APT"
    sudo apt-get update -qq
    sudo apt-get install -y $ALL_APT
    echo "✅ Пакеты apt установлены"
else
    echo "✅ Все apt-пакеты установлены"
fi

# Поиск Python
if command -v python3 &>/dev/null; then
    PYTHON="python3"
elif command -v python &>/dev/null; then
    PYTHON="python"
else
    echo "❌ Python не найден. Установка..."
    sudo apt-get install -y python3
    PYTHON="python3"
fi

echo "🐍 Используется: $($PYTHON --version)"

# Проверка tkinter (требуется для GUI)
echo "🔍 Проверка tkinter..."
if ! $PYTHON -c "import tkinter" 2>/dev/null; then
    echo "❌ tkinter не доступен."
    echo "📦 Установка: sudo apt-get install python3-tk"
    exit 1
fi
echo "✅ tkinter доступен"

# Установка через pip оставшихся зависимостей
echo "🔍 Проверка оставшихся зависимостей..."
MISSING_DEPS=()

for dep in pandas openpyxl requests aiohttp ttkbootstrap; do
    if ! $PYTHON -c "import $dep" 2>/dev/null; then
        MISSING_DEPS+=("$dep")
    fi
done

if [ ${#MISSING_DEPS[@]} -gt 0 ]; then
    echo "⚠️  Отсутствуют зависимости: ${MISSING_DEPS[*]}"
    echo "📦 Установка через pip (--user)..."
    $PYTHON -m pip install --user --upgrade pip --quiet 2>/dev/null || true
    for dep in "${MISSING_DEPS[@]}"; do
        echo "   Установка $dep..."
        $PYTHON -m pip install --user "$dep" --quiet || {
            echo "❌ Не удалось установить $dep"
            exit 1
        }
    done
    echo "✅ Зависимости установлены"
else
    echo "✅ Все зависимости установлены"
fi

# Запуск приложения
echo "🚀 Запуск приложения..."
echo "======================================================================"

cd "$SCRIPT_DIR"
exec $PYTHON arshin_app.py "$@"
