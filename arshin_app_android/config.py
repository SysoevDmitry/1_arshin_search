# -*- coding: utf-8 -*-
"""
Конфигурация Android-приложения ФГИС АРШИН
"""

import os
import csv
import io
import logging

logger = logging.getLogger(__name__)


class Config:
    """Настройки приложения"""
    API_BASE_URL = "https://fgis.gost.ru/fundmetrology/eapi"
    API_VRI = f"{API_BASE_URL}/vri"

    MAX_ROWS_PER_REQUEST = 100
    REQUEST_TIMEOUT = 30
    REQUEST_DELAY = 0.3

    DB_PATH = "arshin_android.db"
    EXPORT_DIR = "exports"

    @classmethod
    def get_config_dir(cls) -> str:
        """Путь к папке config/ — в APK или рядом с main.py"""
        script_dir = os.path.dirname(os.path.abspath(__file__))
        config_dir = os.path.join(script_dir, "config")
        if os.path.isdir(config_dir):
            return config_dir

        from kivy.app import App
        app = App.get_running_app()
        if app:
            path = os.path.join(app.user_data_dir, "config")
            if os.path.isdir(path):
                return path

        return script_dir

    @classmethod
    def get_db_path(cls) -> str:
        """Путь к БД в user_data_dir (Android) или рядом с main.py"""
        from kivy.app import App
        app = App.get_running_app()
        if app:
            return os.path.join(app.user_data_dir, cls.DB_PATH)
        return cls.DB_PATH

    @classmethod
    def load_exact_queries(cls) -> list:
        """Загрузка точных запросов из CSV"""
        config_dir = cls.get_config_dir()
        file_path = os.path.join(config_dir, "exact_queries.csv")
        queries = []
        if os.path.exists(file_path):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    lines = [l for l in f if l.strip() and not l.strip().startswith('#')]
                csv_content = 'query,category,priority\n' + ''.join(lines)
                reader = csv.DictReader(io.StringIO(csv_content))
                for row in reader:
                    if 'query' in row and row['query'].strip():
                        queries.append(row['query'].strip())
                logger.info(f"Загружено {len(queries)} точных запросов")
            except Exception as e:
                logger.error(f"Ошибка загрузки exact_queries.csv: {e}")
        return queries

    @classmethod
    def load_manufacturers(cls) -> dict:
        """Загрузка словаря производителей из CSV"""
        config_dir = cls.get_config_dir()
        file_path = os.path.join(config_dir, "manufacturers.csv")
        manufacturers = {}
        if os.path.exists(file_path):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    lines = [l for l in f if l.strip() and not l.strip().startswith('#')]
                csv_content = 'keyword,manufacturer,category,priority\n' + ''.join(lines)
                reader = csv.DictReader(io.StringIO(csv_content))
                for row in reader:
                    if 'keyword' in row and 'manufacturer' in row:
                        manufacturers[row['keyword'].lower().strip()] = row['manufacturer'].strip()
                logger.info(f"Загружено {len(manufacturers)} производителей")
            except Exception as e:
                logger.error(f"Ошибка загрузки manufacturers.csv: {e}")
        return manufacturers


EXACT_QUERIES = Config.load_exact_queries()
MANUFACTURERS_RULES = Config.load_manufacturers()
