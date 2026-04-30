#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ФГИС АРШИН - Excel версия с параллельными запросами
Версия 6.2 - Поддержка продолжения после сбоя и экспорт БД

Возможности:
- Пакетная обработка Excel файлов
- Параллельные запросы к API
- Сохранение связей (Id_ПУ, № строки, данные из файла)
- Проверка на ПОЛНУЮ уникальность (как в arshin_app.py)
- Автоматическое определение колонок
- Экспорт с русскими заголовками
- АТРИБУТИВНЫЙ ПОИСК (рекомендуется спецификацией)
- Поддержка verification_date (конкретная дата поверки)
- Получение записей по vri_id
- Обработка 408 Request Timeout
- Детализированное логирование 5XX ошибок
- ПРОДОЛЖЕНИЕ ПОСЛЕ СБОЯ (сохранение прогресса)
- ЭКСПОРТ БД без повторного поиска (--export-only)

Запуск:
    python main.py -f запрос.xlsx -y 2020-2025 -o результат.csv
    python main.py -f запрос.xlsx --verification-date 2024-05-15 -o результат.csv
    python main.py --template шаблон.xlsx
    python main.py --stats
    python main.py --get-record vri_id
    python main.py --export-only результат.csv      # экспорт без поиска
    python main.py -f запрос.xlsx --resume           # продолжить после сбоя
"""

import os
import sys
import argparse
import logging
import asyncio
from datetime import datetime
from typing import List, Optional

from config import (
    Config, EXACT_QUERIES, MANUFACTURERS_RULES,
    TQDM_AVAILABLE, PANDAS_AVAILABLE
)
from database import Database
from excel_handler import ExcelHandler
from collector import ExcelCollector
from api_client import ParallelAPIClient

# Логирование
LOG_DIR = "Logs"
os.makedirs(LOG_DIR, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(
            os.path.join(LOG_DIR, f"arshin_excel_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"),
            encoding='utf-8'
        )
    ]
)
logger = logging.getLogger(__name__)


def parse_years(years_str: str) -> List[int]:
    """Парсинг годов: 2020,2021,2022 или 2020-2025"""
    years = []
    for y in years_str.split(','):
        y = y.strip()
        if '-' in y:
            start, end = map(int, y.split('-'))
            years.extend(range(start, end + 1))
        else:
            years.append(int(y))
    return sorted(set(years))


def print_stats(stats: dict, db_stats: dict):
    """Вывод статистики обработки"""
    print("\n" + "="*70)
    print("📊 СТАТИСТИКА ОБРАБОТКИ")
    print("="*70)
    print(f"   Запросов обработано: {stats.get('queries_processed', 0)}")
    print(f"   Найдено записей: {stats.get('records_found', 0)}")
    print(f"   Сохранено: {stats.get('records_saved', 0)}")
    print(f"   Пропущено дублей: {stats.get('duplicates', 0)}")
    print(f"   Отфильтровано (не электросчетчики): {stats.get('filtered', 0)}")
    
    print("\n" + "="*70)
    print("📊 СТАТИСТИКА БАЗЫ ДАННЫХ")
    print("="*70)
    print(f"   Всего записей: {db_stats.get('total', 0)}")
    
    if db_stats.get('total', 0) > 0:
        print(f"   Пригодных: {db_stats.get('applicable', 0)} "
              f"({db_stats['applicable']/db_stats['total']*100:.1f}%)")
        print(f"   Непригодных: {db_stats.get('inapplicable', 0)}")
        
        if db_stats.get('by_year'):
            print("\n   По годам:")
            for year, count in list(db_stats['by_year'].items())[:10]:
                print(f"   - {year}: {count}")
        
        if db_stats.get('by_manufacturer'):
            print("\n   Производители (ТОП-10):")
            for mfg, count in list(db_stats['by_manufacturer'].items())[:10]:
                pct = (count / db_stats['total'] * 100) if db_stats['total'] else 0
                print(f"   - {mfg}: {count} ({pct:.1f}%)")
    
    print("="*70 + "\n")


def prepare_output_path(output: str) -> str:
    """
    Формирование имени выходного файла с датой и номером.
    Результат сохраняется в папку exports/ (создаётся автоматически).
    Если файл существует — добавляет дату впереди и/или номер по порядку.
    
    Примеры:
        результат.csv → exports/результат.csv (если не существует)
        результат.csv → exports/20260429_результат.csv (если существует)
        20260429_результат.csv → exports/20260429_результат_1.csv (если существует)
    """
    import os
    from datetime import datetime
    
    # Если путь без папки — сохраняем в exports/
    directory = os.path.dirname(output)
    if not directory:
        directory = "exports"
        output = os.path.join(directory, output)
    
    os.makedirs(directory, exist_ok=True)
    
    if not os.path.exists(output):
        return output
    
    # Разбираем путь
    basename = os.path.basename(output)
    name, ext = os.path.splitext(basename)
    today = datetime.now().strftime('%Y%m%d')
    
    # Если имя уже начинается с даты — добавляем номер
    if name.startswith(today):
        counter = 1
        while True:
            new_name = f"{name}_{counter}{ext}"
            new_path = os.path.join(directory, new_name)
            if not os.path.exists(new_path):
                return new_path
            counter += 1
    else:
        # Добавляем дату в начало
        new_name = f"{today}_{name}{ext}"
        new_path = os.path.join(directory, new_name)
        if not os.path.exists(new_path):
            return new_path
        # Если и с датой существует — добавляем номер
        counter = 1
        while True:
            numbered_name = f"{today}_{name}_{counter}{ext}"
            numbered_path = os.path.join(directory, numbered_name)
            if not os.path.exists(numbered_path):
                return numbered_path
            counter += 1


async def run_excel_search(filename: str, years: List[int], output: Optional[str],
                            verification_date: str = None,
                            use_attribute_search: bool = True,
                            resume: Optional[bool] = None):
    """Поиск по Excel файлу

    Args:
        filename: Путь к Excel файлу
        years: Список годов для поиска
        output: Путь для экспорта CSV
        verification_date: Конкретная дата поверки (вместо year)
        use_attribute_search: Использовать атрибутивный поиск
        resume: True — принудительно продолжить, False — начать заново, None — спросить
    """
    print(f"\n📂 Чтение файла: {filename}")

    if not os.path.exists(filename):
        print(f"❌ Файл не найден: {filename}")
        return

    # Чтение запросов
    queries = ExcelHandler.read_queries(filename)

    if not queries:
        print("❌ Не найдено запросов в файле")
        return

    print(f"✅ Загружено {len(queries)} запросов")

    if verification_date:
        print(f"📅 Дата поверки: {verification_date}")
    else:
        print(f"📅 Годы поиска: {', '.join(map(str, years))}")

    print(f"🚀 Параллельных запросов: {Config.MAX_CONCURRENT_REQUESTS}")
    print(f"🔍 Атрибутивный поиск: {'включен' if use_attribute_search else 'выключен'}")

    # Обработка
    db = Database()
    start_from = 0

    # Проверка прогресса предыдущего запуска
    progress = db.get_progress(filename)
    if progress:
        last_idx = progress['last_processed_index']
        prev_total = progress['total_queries']
        prev_params = progress.get('params', '')

        print(f"\n📌 Обнаружен прогресс предыдущего запуска:")
        print(f"   Обработано: {last_idx + 1} из {prev_total} запросов")
        print(f"   Параметры: {prev_params}")
        print(f"   Всего запросов сейчас: {len(queries)}")

        if prev_total != len(queries):
            print("⚠️  Количество запросов изменилось с прошлого раза!")
            if resume is None:
                print("   Продолжение с изменённым файлом небезопасно")
                answer = input("   Продолжить с последнего индекса? (y/n): ").strip().lower()
                if answer in ('y', 'yes', 'да', '1'):
                    start_from = last_idx + 1
                else:
                    db.clear_progress(filename)
        else:
            if resume is None:
                # Диалог с пользователем
                print(f"\n🔄 Продолжить с индекса {last_idx + 1}?")
                print("   y/yes/да  — продолжить (пропустить обработанные запросы)")
                print("   n/no/нет  — начать заново (удалить прогресс)")
                answer = input("   Ваш выбор: ").strip().lower()
                if answer in ('y', 'yes', 'да', '1', ''):
                    start_from = last_idx + 1
                    print(f"✅ Продолжаем с индекса {start_from}")
                else:
                    db.clear_progress(filename)
                    print("🔄 Начинаем заново")
            elif resume:
                start_from = last_idx + 1
                print(f"✅ Принудительное продолжение с индекса {start_from}")
            else:
                db.clear_progress(filename)
                print("🔄 Принудительный старт заново")

    async with ExcelCollector(db, input_file=filename) as collector:
        stats = await collector.process_queries_batch(
            queries, years,
            verification_date=verification_date,
            use_attribute_search=use_attribute_search,
            start_from=start_from
        )
        db_stats = db.get_stats()

        print_stats(stats, db_stats)

        # Экспорт
        if output:
            output = prepare_output_path(output)
            print(f"📤 Экспорт в {output}...")
            count = db.export_to_csv(output)
            print(f"✅ Экспортировано {count} записей")


def main():
    parser = argparse.ArgumentParser(
        description='ФГИС АРШИН - Excel версия v6.1',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Примеры:
  # Поиск по Excel файлу (по годам)
  %(prog)s -f запрос.xlsx -y 2020,2021,2022 -o результат.csv

  # Поиск по диапазону годов
  %(prog)s -f данные.xlsx -y 2020-2025 -o export.csv

  # Поиск по конкретной дате поверки
  %(prog)s -f данные.xlsx --verification-date 2024-05-15 -o export.csv

  # Атрибутивный поиск (рекомендуется спецификацией)
  %(prog)s -f запрос.xlsx -y 2024 --attribute-search

  # Получить запись по vri_id
  %(prog)s --get-record 2-162158132

  # Создать шаблон Excel
  %(prog)s --template шаблон.xlsx

  # Статистика БД
  %(prog)s --stats

  # Очистить БД
  %(prog)s --clear
        '''
    )

    parser.add_argument('-f', '--file', type=str,
                        help='Excel файл с запросами')
    parser.add_argument('-y', '--years', type=str, default='2020-2025',
                        help='Годы поиска (2020,2021 или 2020-2025)')
    parser.add_argument('-o', '--output', type=str,
                        help='Экспорт в CSV')

    parser.add_argument('--verification-date', type=str,
                        help='Дата поверки (yyyy-MM-dd), приоритет над --years')
    parser.add_argument('--attribute-search', action='store_true',
                        help='Использовать атрибутивный поиск (рекомендуется)')
    parser.add_argument('--get-record', type=str, metavar='VRI_ID',
                        help='Получить запись по vri_id')

    parser.add_argument('--template', type=str,
                        help='Создать шаблон Excel')
    parser.add_argument('--stats', action='store_true',
                        help='Статистика БД')
    parser.add_argument('--clear', action='store_true',
                        help='Очистить БД')

    parser.add_argument('--concurrent', type=int, default=5,
                        help='Кол-во одновременных запросов (3-10)')

    parser.add_argument('--resume', action='store_true',
                        help='Принудительно продолжить с места предыдущего сбоя')
    parser.add_argument('--no-resume', action='store_true',
                        help='Игнорировать прогресс, начать заново')
    parser.add_argument('--export-only', type=str, metavar='CSV',
                        help='Только экспорт существующей БД в CSV (без поиска)')

    args = parser.parse_args()

    # Заголовок
    print("\n" + "="*70)
    print("🔍 ФГИС АРШИН - Excel версия v6.1")
    print("="*70)
    print(f"📅 Дата: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"🚀 Параллельных запросов: {args.concurrent}")
    print(f"📦 Точных запросов: {len(EXACT_QUERIES)}")
    print(f"🏭 Производителей: {len(MANUFACTURERS_RULES)}")
    if args.verification_date:
        print(f"📅 Дата поверки: {args.verification_date}")
    if args.attribute_search:
        print(f"🔍 Атрибутивный поиск: включен")
    print("="*70)

    # Проверка библиотек
    if not PANDAS_AVAILABLE:
        print("❌ pandas не установлен: pip install pandas openpyxl")
        return

    # Настройка параллелизма
    Config.MAX_CONCURRENT_REQUESTS = min(max(args.concurrent, 1), 10)

    # Получение записи по vri_id
    if args.get_record:
        async def get_record():
            async with ParallelAPIClient() as client:
                result = await client.get_vri_record_with_retry(args.get_record)
                if result.get('error') is None:
                    print(f"\n✅ Запись получена:")
                    import json
                    print(json.dumps(result.get('item', {}), indent=2, ensure_ascii=False))
                else:
                    print(f"\n❌ Ошибка: {result.get('error')} - {result.get('message', '')}")
        asyncio.run(get_record())
        return

    # Статистика
    if args.stats:
        db = Database()
        stats = db.get_stats()
        print_stats({}, stats)
        return

    # Очистка
    if args.clear:
        if input("⚠️  Очистить БД? (yes/no): ").lower() == 'yes':
            db = Database()
            db.clear()
            print("✅ БД очищена")
        return

    # Создание шаблона
    if args.template:
        filename = ExcelHandler.create_template(args.template)
        if filename:
            print(f"✅ Шаблон создан: {filename}")
        return

    # Только экспорт БД (без поиска)
    if args.export_only:
        db = Database()
        stats = db.get_stats()
        print(f"\n📊 В базе данных {stats['total']} записей")
        if stats['total'] > 0:
            output = prepare_output_path(args.export_only)
            print(f"📤 Экспорт в {output}...")
            count = db.export_to_csv(output)
            print(f"✅ Экспортировано {count} записей")
        else:
            print("❌ Нет данных для экспорта")
        return

    # Файл обязателен
    if not args.file:
        parser.print_help()
        return

    # Парсинг годов
    years = parse_years(args.years)
    print(f"📋 Выбрано годов: {len(years)} ({min(years)}-{max(years)})")

    # Вопрос об очистке БД перед поиском
    db_check = Database()
    existing_count = db_check.get_stats().get('total', 0)
    resume_flag = args.resume

    # Если есть прогресс, но флаг --no-resume — очищаем старый прогресс
    if args.no_resume:
        db_check.clear_progress(args.file)
        print("🔄 Прогресс сброшен, начинаем заново")

    if existing_count > 0:
        print(f"\n📊 В базе уже есть {existing_count} записей")
        # Если --resume или обнаружен прогресс — не предлагаем очистку
        progress = db_check.get_progress(args.file) if not args.no_resume else None
        if progress or args.resume:
            print("📥 Режим продолжения — существующие записи сохраняются")
        else:
            print("  очистить / yes — удалить все записи и начать заново")
            print("  добавить / no  — добавить новые записи к существующим")
            answer = input("🗑  Ваш выбор: ").strip().lower()
            if answer in ('очистить', 'clear', 'y', 'yes', 'да', '1'):
                db_check.clear()
                print("✅ БД очищена")
            else:
                print("📥 Новые записи будут добавлены к существующим")

    # Запуск поиска
    if args.resume:
        resume_flag = True
    elif args.no_resume:
        resume_flag = False
    else:
        resume_flag = None

    asyncio.run(run_excel_search(
        args.file, years, args.output,
        verification_date=args.verification_date,
        use_attribute_search=args.attribute_search,
        resume=resume_flag
    ))


if __name__ == '__main__':
    main()
