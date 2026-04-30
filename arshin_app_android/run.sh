#!/bin/bash
# -*- coding: utf-8 -*-
# Запуск Android-приложения на ПК (для отладки)
# Требуется: pip install kivy requests pandas openpyxl

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

# Активация виртуального окружения (если есть)
if [ -f "../venv/bin/activate" ]; then
    source "../venv/bin/activate"
fi

# Проверка Kivy
python3 -c "import kivy" 2>/dev/null || {
    echo "Установка Kivy..."
    pip install kivy
}

echo "Запуск приложения ФГИС АРШИН (Android-версия, отладка на ПК)..."
python3 main.py
