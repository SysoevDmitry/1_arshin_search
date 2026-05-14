# AGENTS.md — ФГИС АРШИН

> КРИТИЧНО: Все внутренние рассуждения (thinking) — ТОЛЬКО на русском языке. НЕЛЬЗЯ использовать английский.

## Правила проекта

- **Документация** — все новые/обновлённые документы складывать в `Doc/`. Актуализировать документацию при изменении кода.
- **Локаль** — эталонная локаль русская (ru). Весь код, комментарии, названия полей, UI, CSV-заголовки — на русском. Новый код тоже на русском.
- **Венв** — общий для всего репозитория: `venv/` в корне. Активировать: `source venv/bin/activate`.
- **Работа с путями** — в репо есть кириллические пути (папки, имена файлов). Всегда экранировать/закавычивать.

## Структура репозитория

Два приложения, общий venv и конфиги:

```
arshin_search/
├── arshin_app.py          # GUI-приложение (Tkinter), ~3200 строк
├── Konsol_Excel/          # Консольная версия (пакетный поиск из Excel)
│   ├── main.py            # Точка входа (asyncio + argparse)
│   ├── api_client.py      # aiohttp-клиент к API АРШИН
│   ├── collector.py       # Логика: поиск → фильтр → сохранение
│   ├── database.py        # SQLite с проверкой полной уникальности
│   ├── models.py          # Dataclass VerificationRecord
│   ├── excel_handler.py   # Чтение/запись Excel, автоопределение колонок
│   ├── config.py          # Конфигурация (загружает CSV из Konsol_Excel/config/)
│   └── config/            # Копия корневого config/
├── config/                # Общие CSV-конфиги (exact_queries.csv, manufacturers.csv)
├── Doc/                   # Документация
├── тесты/                 # Тестовые Excel-файлы (не автотесты)
├── App_Linux/             # Портативная сборка (AppImage)
├── App_Windows/           # Портативная сборка (EXE)
└── requirements.txt       # pandas, openpyxl, aiohttp, tqdm, requests
```

### Важно: дублирование конфигов
Файлы в `config/` (корень) и `Konsol_Excel/config/` — **копии**. Изменять нужно **оба**.
- `arshin_app.py` читает `config/` относительно корня репо (через `run.sh` выставляет `pwd`)
- `Konsol_Excel/config.py` читает `Konsol_Excel/config/` относительно своей папки (через `Konsol_Excel/run.sh`)

## Запуск

```bash
source venv/bin/activate

# GUI
./run.sh                                          # или: python arshin_app.py

# Konsol_Excel
cd Konsol_Excel && ./run.sh --stats
cd Konsol_Excel && python main.py -f запрос.xlsx -y 2020-2025 -o результат.csv
cd Konsol_Excel && python main.py --template шаблон.xlsx
cd Konsol_Excel && python main.py -f запрос.xlsx --resume  # продолжить после сбоя
cd Konsol_Excel && python main.py --export-only результат.csv  # только экспорт БД
```

Скрипты `run.sh`/`run.bat` автоматически проверяют и устанавливают зависимости.

## Ключевая архитектура (Konsol_Excel)

- `main.py` → `ExcelHandler.read_queries()` читает Excel → `ExcelCollector.process_queries_batch()` обрабатывает запросы → `ParallelAPIClient.search_vri()` ходит в API → `Database.save_records_batch()` сохраняет с проверкой дублей → `Database.export_to_csv()` выгружает
- Проверка уникальности — по **13 полям** (`vri_id` + все основные + `search_query` + `row_index` + `id_pu`). Одна API-запись может быть сохранена под разными Id_ПУ
- Фильтрация результатов API: только электросчётчики (функция `is_electric_meter_from_queries`), точное совпадение серийного номера (`match_serial_number`)
- Параллелизм через `asyncio.Semaphore`, по умолчанию 5, меняется через `--concurrent N` (3–10)
- Retry на ошибки 429/502/408 с экспоненциальной задержкой
- **v6.2**: Сохранение прогресса после каждого запроса в таблице `search_progress`. При повторном запуске с тем же файлом — диалог продолжения (пропуск обработанных строк).

## Продолжение после сбоя (v6.2)

- Прогресс сохраняется в таблице `search_progress` (SQLite) после каждого успешно обработанного запроса
- При запуске с тем же Excel-файлом программа обнаруживает прогресс и предлагает продолжить
- Флаги:
  - `--resume` — принудительно продолжить, без диалога
  - `--no-resume` — игнорировать прогресс, начать заново
  - `--export-only CSV` — только экспорт существующей БД в CSV (без поиска)
- При успешном завершении прогресс автоматически очищается
- Если файл изменился (другое количество запросов) — выводится предупреждение

## API

- URL: `https://fgis.gost.ru/fundmetrology/eapi/vri`
- Параметры: `search`, `year`, `verification_date`, `start`, `rows`, + атрибутивные фильтры (`mi_number`, `mit_number`, ...)
- v6.1 поддерживает атрибутивный поиск (`--attribute-search`) и `--verification-date`
- v6.2 добавляет сохранение прогресса (`--resume`, `--no-resume`, `--export-only`)

## База данных

- GUI: `arshin_data.db` (корень)
- Konsol_Excel: `arshin_excel.db` (внутри `Konsol_Excel/`)
- Обе исключены из git (`.gitignore`), создаются автоматически при первом запуске

## Тестирование

Автотестов нет. Тестовые Excel-файлы в `тесты/` для ручной проверки.
Перед изменениями проверять работоспособность вручную на тестовом файле:
```bash
cd Konsol_Excel && python main.py -f ../тесты/Тестовый\ запрос_3.xlsx -y 2025 -o /tmp/test.csv
```

## Зависимости, не описанные в requirements.txt

- **tkinter** — нужен для GUI (`arshin_app.py`). Системный пакет: `python3-tk` (Ubuntu) или `python3-tkinter` (RHEL)
- **ttkbootstrap** — нужен для GUI (тема оформления). Проверяется в `run.sh`, но **отсутствует** в `requirements.txt`. Устанавливается вручную скриптами запуска.
- **SQLite3** — стандартный модуль Python, но на ALT Linux нужен системный пакет `python3-modules-sqlite3`

## Известные проблемы

- При ошибках API 429 уменьшить `--concurrent` до 3
- Если колонки Excel не определяются — проверить названия заголовков (поддерживаются: «Заводской номер», «Серийный номер», «Id_ПУ» и т.д.)
- `Konsol_Excel` не имеет `run.bat` (только `run.sh`). На Windows запускать `python main.py ...`
