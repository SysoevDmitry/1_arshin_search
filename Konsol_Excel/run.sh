#!/bin/bash
# ФГИС АРШИН - Konsol_Excel - Скрипт запуска для Linux
# Проверка виртуального окружения и зависимостей

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="$SCRIPT_DIR/../venv"
VENV_PYTHON="$VENV_DIR/bin/python"
REQUIREMENTS_FILE="$SCRIPT_DIR/../requirements.txt"
PYTHON=""

echo "======================================================================"
echo "🔍 ФГИС АРШИН - Konsol_Excel v6.1"
echo "======================================================================"

# Поиск системного Python
SYSTEM_PYTHON=""
if command -v python3 &> /dev/null; then
    SYSTEM_PYTHON="python3"
elif command -v python &> /dev/null; then
    SYSTEM_PYTHON="python"
else
    echo "❌ Python не найден. Установите Python 3.8+"
    exit 1
fi

# Проверка и создание виртуального окружения
if [ -d "$VENV_DIR" ] && [ -f "$VENV_PYTHON" ]; then
    echo "✅ Виртуальное окружение найдено: $VENV_DIR"
else
    echo "⚠️  Виртуальное окружение не найдено: $VENV_DIR"
    echo "📦 Создание виртуального окружения..."
    $SYSTEM_PYTHON -m venv "$VENV_DIR"
    echo "✅ Виртуальное окружение создано"
fi

# Активация виртуального окружения
PYTHON="$VENV_PYTHON"
source "$VENV_DIR/bin/activate"
echo "✅ Виртуальное окружение активировано"

# Проверка зависимостей в venv
echo "🔍 Проверка зависимостей..."
MISSING_DEPS=()

for dep in pandas openpyxl aiohttp tqdm; do
    if ! $PYTHON -c "import $dep" 2>/dev/null; then
        MISSING_DEPS+=("$dep")
    fi
done

if [ ${#MISSING_DEPS[@]} -gt 0 ]; then
    echo "⚠️  Отсутствуют зависимости: ${MISSING_DEPS[*]}"
    echo "📦 Установка зависимостей..."
    
    if [ -f "$REQUIREMENTS_FILE" ]; then
        $PYTHON -m pip install -r "$REQUIREMENTS_FILE" --quiet
    else
        $PYTHON -m pip install pandas openpyxl aiohttp tqdm --quiet
    fi
    
    echo "✅ Зависимости установлены"
else
    echo "✅ Все зависимости установлены"
fi

# Запуск приложения
echo "🚀 Запуск приложения..."
echo "======================================================================"

cd "$SCRIPT_DIR"
exec $PYTHON main.py "$@"
