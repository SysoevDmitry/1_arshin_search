#!/bin/bash
# Скрипт запуска ФГИС АРШИН для Linux
# Использование: ./run_arshin.sh

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "========================================"
echo "ФГИС АРШИН - Запуск"
echo "========================================"

# Проверка Python
if ! command -v python3 &> /dev/null; then
    echo "❌ Ошибка: Python3 не найден"
    echo "   Установите Python3: sudo apt install python3"
    exit 1
fi

echo "✓ Python: $(python3 --version)"

# Создание/проверка виртуального окружения
if [ ! -d "venv" ]; then
    echo ""
    echo "📦 Создание виртуального окружения..."
    python3 -m venv venv
fi

# Активация
source venv/bin/activate

# Установка/обновление зависимостей
echo ""
echo "📋 Проверка зависимостей..."
pip install -q -r requirements.txt

# Запуск приложения
echo ""
echo "🚀 Запуск приложения..."
echo "========================================"
python3 arshin_app.py

# Деактивация после завершения
deactivate 2>/dev/null || true
