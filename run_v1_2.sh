#!/bin/bash
# Запуск arshin_app_v1_2.py с правильным Python из venv
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
exec "$SCRIPT_DIR/venv/bin/python3" "$SCRIPT_DIR/arshin_app_v1_2.py" "$@"
