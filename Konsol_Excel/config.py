#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Конфигурация для Excel версии ФГИС АРШИН
"""

import os
import csv
import io
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class Config:
    """Настройки приложения"""
    API_BASE_URL = "https://fgis.gost.ru/fundmetrology/eapi"
    API_VRI = f"{API_BASE_URL}/vri"
    
    # Параметры запросов
    MAX_ROWS_PER_REQUEST = 100
    REQUEST_TIMEOUT = 30
    
    # Параллелизм
    MAX_CONCURRENT_REQUESTS = 5  # Более консервативно для точных запросов
    REQUEST_DELAY = 0.2  # Пауза между запросами (сек)
    
    # База данных
    DB_PATH = "arshin_excel.db"
    
    # Файлы конфигурации
    CONFIG_DIR = "config"
    EXACT_QUERIES_FILE = os.path.join(CONFIG_DIR, "exact_queries.csv")
    MANUFACTURERS_FILE = os.path.join(CONFIG_DIR, "manufacturers.csv")
    
    # Экспорт
    EXPORT_DIR = "exports"

    @classmethod
    def load_exact_queries(cls) -> list:
        queries = []
        if os.path.exists(cls.EXACT_QUERIES_FILE):
            try:
                with open(cls.EXACT_QUERIES_FILE, 'r', encoding='utf-8') as f:
                    lines = [l for l in f if l.strip() and not l.strip().startswith('#')]
                csv_content = 'query,category,priority\n' + ''.join(lines)
                reader = csv.DictReader(io.StringIO(csv_content))
                for row in reader:
                    if 'query' in row and row['query'].strip():
                        queries.append(row['query'].strip())
                logger.info(f"✅ Загружено {len(queries)} точных запросов")
            except Exception as e:
                logger.error(f"Ошибка: {e}")
        return queries

    @classmethod
    def load_manufacturers(cls) -> dict:
        manufacturers = {}
        if os.path.exists(cls.MANUFACTURERS_FILE):
            try:
                with open(cls.MANUFACTURERS_FILE, 'r', encoding='utf-8') as f:
                    lines = [l for l in f if l.strip() and not l.strip().startswith('#')]
                csv_content = 'keyword,manufacturer,category,priority\n' + ''.join(lines)
                reader = csv.DictReader(io.StringIO(csv_content))
                for row in reader:
                    if 'keyword' in row and 'manufacturer' in row:
                        manufacturers[row['keyword'].lower().strip()] = row['manufacturer'].strip()
                logger.info(f"✅ Загружено {len(manufacturers)} производителей")
            except Exception as e:
                logger.error(f"Ошибка: {e}")
        return manufacturers


# Глобальные данные
EXACT_QUERIES = Config.load_exact_queries()
MANUFACTURERS_RULES = Config.load_manufacturers()

# Проверка библиотек
try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False
    logger.warning("⚠️  pandas не установлен")

try:
    from tqdm import tqdm
    TQDM_AVAILABLE = True
except ImportError:
    TQDM_AVAILABLE = False
    logger.warning("⚠️  tqdm не установлен")
