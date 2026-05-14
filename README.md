# 🔍 ФГИС АРШИН — Монорепозиторий

**Дата:** 14 мая 2026 г.
**Платформы:** Linux (x64/ARM), Windows

---

## 📖 ОПИСАНИЕ

Монорепозиторий инструментов для работы с **ФГИС АРШИН** — федеральной информационной системой Росстандарта (единый реестр средств измерений РФ). Содержит две версии приложения и общую конфигурацию.

## 📁 СТРУКТУРА

```
arshin_search/
├── README.md                       # Этот файл
├── .gitignore
│
├── arshin_app/                     # GUI приложение v3.0 (Tkinter)
│   ├── arshin_app.py               # Точка входа
│   ├── README.md                   # Документация GUI
│   ├── requirements.txt            # Зависимости
│   ├── run.sh / run.bat            # Запуск (Linux/Windows)
│   └── run_armbian.sh              # Запуск (ARM)
│
├── Konsol_Excel/                   # Консольная версия v6.3 (Excel + SQLite)
│   ├── main.py                     # Точка входа
│   ├── api_client.py               # aiohttp-клиент API
│   ├── collector.py                # Сборщик данных
│   ├── database.py                 # SQLite (UPSERT + прогресс)
│   ├── excel_handler.py            # Чтение Excel
│   ├── models.py                   # Модели данных
│   ├── config.py                   # Конфигурация
│   ├── README.md                   # Документация
│   ├── KONSOL_APP.md               # Полное руководство
│   ├── run.sh / run.sh             # Запуск (Linux)
│   ├── run.ps1 / run_bg.ps1        # Запуск (Windows)
│   ├── run_armbian.sh              # Запуск (ARM)
│   └── config/                     # Копия конфигов
│
├── config/                         # Общие конфигурационные файлы
│   ├── README.md
│   ├── exact_queries.csv            # Точные запросы (электросчётчики)
│   └── manufacturers.csv            # Словарь производителей
│
├── Docs/                            # Документация
│   ├── ПРЕЗЕНТАЦИЯ.md              # Презентация проекта
│   └── Аршин интерфейсы v.2.2.md   # Спецификация API ФГИС АРШИН
│
└── arshin_app_android/              # Android-версия (локально, не в репо)
```

## 🚀 БЫСТРЫЙ СТАРТ

### GUI (графический интерфейс)

```bash
cd arshin_app
source ../venv/bin/activate
pip install -r requirements.txt
./run.sh          # Linux
```

### Консольная версия (терминал/SSH)

```bash
cd Konsol_Excel
./run.sh -f запрос.xlsx -y 2020-2025 -o результат.csv
./run.sh --stats                   # Статистика БД
./run.sh --resume                  # Продолжить после сбоя
```

## 📊 СРАВНЕНИЕ ВЕРСИЙ

| Характеристика | GUI (arshin_app) | Консоль (Konsol_Excel) |
|---|---|---|
| Версия | v3.0 | v6.3 |
| Интерфейс | Tkinter GUI | Терминал (SSH/tmux/screen) |
| Поиск | Одиночный + пакетный | Пакетный из Excel |
| Параллельность | Нет | До 10 запросов (asyncio) |
| БД | SQLite | SQLite (UPSERT + прогресс) |
| Экспорт | CSV, Excel | CSV (21 поле) |
| Автоматизация | Нет | cron, tmux, screen |
| Возобновление | Нет | DB-прогресс + --resume |
| Прогноз времени | Нет | Часы/дни в прогресс-баре |
| Windows | ✅ (GUI + консоль) | ✅ (PowerShell) |
| ARM (Armbian) | ✅ | ✅ (run_armbian.sh) |

## 📚 ДОКУМЕНТАЦИЯ

| Файл | Описание |
|------|----------|
| `arshin_app/README.md` | Документация GUI-версии |
| `Konsol_Excel/README.md` | Краткое руководство по консольной версии |
| `Konsol_Excel/KONSOL_APP.md` | Полное руководство по консольной версии |
| `Docs/ПРЕЗЕНТАЦИЯ.md` | Презентация проекта |
| `Docs/Аршин интерфейсы v.2.2.md` | Спецификация API ФГИС АРШИН |
| `config/README.md` | Описание конфигурационных файлов |

## 🔧 КОНФИГУРАЦИЯ

Общие конфигурационные файлы в `config/`, копия в `Konsol_Excel/config/`:

- `exact_queries.csv` — точные поисковые запросы (электросчётчики)
- `manufacturers.csv` — словарь производителей

---

**Разработчик:** АРШИН Проект
