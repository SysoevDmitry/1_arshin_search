#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
База данных для Excel версии
С ПРОВЕРКОЙ НА ПОЛНУЮ УНИКАЛЬНОСТЬ как в arshin_app.py
"""

import sqlite3
import logging
from typing import List, Set, Dict, Optional
from models import VerificationRecord

logger = logging.getLogger(__name__)


class Database:
    """
    SQLite база с проверкой на полную уникальность
    Адаптировано из arshin_app.py
    """
    
    def __init__(self, db_path: str = "arshin_excel.db"):
        self.db_path = db_path
        self._init_db()
    
    def _init_db(self):
        """Инициализация таблиц БД с полной структурой как в arshin_app.py"""
        with sqlite3.connect(self.db_path) as conn:
            # Таблица прогресса поиска (для возможности продолжения после сбоя)
            conn.execute('''
                CREATE TABLE IF NOT EXISTS search_progress (
                    input_file TEXT PRIMARY KEY,
                    last_processed_index INTEGER,
                    total_queries INTEGER,
                    params TEXT,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            conn.execute('''
                CREATE TABLE IF NOT EXISTS verification_records (
                    -- Основные данные из API
                    vri_id TEXT PRIMARY KEY,
                    mit_number TEXT,
                    mit_title TEXT,
                    mit_notation TEXT,
                    mi_modification TEXT,
                    mi_number TEXT,
                    verification_date TEXT,
                    valid_date TEXT,
                    applicability INTEGER,
                    org_title TEXT,
                    result_docnum TEXT,
                    
                    -- Служебные поля
                    manufacturer TEXT,
                    collected_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    
                    -- Связи (для пакетного поиска)
                    search_query TEXT,
                    row_index INTEGER,
                    id_pu TEXT,
                    
                    -- Данные из Excel файла
                    contract_number TEXT,
                    edo_code TEXT,
                    balance_owner TEXT,
                    operation_responsibility TEXT,
                    mpi TEXT
                )
            ''')
            
            # Индексы для ускорения поиска
            conn.execute('CREATE INDEX IF NOT EXISTS idx_manufacturer ON verification_records(manufacturer)')
            conn.execute('CREATE INDEX IF NOT EXISTS idx_verification_date ON verification_records(verification_date)')
            conn.execute('CREATE INDEX IF NOT EXISTS idx_search_query ON verification_records(search_query)')
            conn.execute('CREATE INDEX IF NOT EXISTS idx_mi_number ON verification_records(mi_number)')
            conn.execute('CREATE INDEX IF NOT EXISTS idx_id_pu ON verification_records(id_pu)')
            conn.execute('CREATE INDEX IF NOT EXISTS idx_row_index ON verification_records(row_index)')
            
            conn.commit()
        
        logger.info("📄 База данных инициализирована")
    
    def get_existing_ids(self) -> Set[str]:
        """Получить все сохраненные vri_id"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.execute('SELECT vri_id FROM verification_records')
            ids = {row[0] for row in cursor.fetchall()}
        return ids
    
    def is_duplicate(self, record: VerificationRecord, conn: sqlite3.Connection) -> bool:
        """
        Проверка на ПОЛНЫЙ дубль (все основные + служебные поля)
        Адаптировано из arshin_app.py
        """
        cursor = conn.cursor()
        cursor.execute('''
            SELECT COUNT(*) FROM verification_records
            WHERE vri_id = ?
              AND mit_number = ?
              AND mit_title = ?
              AND mi_number = ?
              AND verification_date = ?
              AND valid_date = ?
              AND applicability = ?
              AND org_title = ?
              AND result_docnum = ?
              AND manufacturer = ?
              AND search_query = ?
              AND row_index = ?
              AND id_pu = ?
        ''', (
            record.vri_id, record.mit_number, record.mit_title,
            record.mi_number, record.verification_date, record.valid_date,
            1 if record.applicability else 0,
            record.org_title, record.result_docnum,
            record.manufacturer,
            record.search_query, record.row_index, record.id_pu
        ))
        
        return cursor.fetchone()[0] > 0
    
    def save_records_batch(self, records: List[VerificationRecord]) -> dict:
        """
        Пакетное сохранение с проверкой на полную уникальность
        Адаптировано из arshin_app.py
        
        Returns:
            dict со статистикой: {'saved': int, 'duplicates': int, 'errors': int}
        """
        if not records:
            return {'saved': 0, 'duplicates': 0, 'errors': 0}
        
        saved = 0
        duplicates = 0
        errors = 0
        
        conn = sqlite3.connect(self.db_path)
        try:
            for record in records:
                try:
                    # Проверка на полный дубль
                    if self.is_duplicate(record, conn):
                        duplicates += 1
                        logger.debug(f"⚠️  Пропущен дубль: vri_id={record.vri_id}, "
                                   f"query={record.search_query}, row={record.row_index}")
                    else:
                        # Вставка новой записи
                        conn.execute('''
                            INSERT OR IGNORE INTO verification_records
                            (vri_id, mit_number, mit_title, mit_notation, mi_modification,
                             mi_number, verification_date, valid_date, applicability,
                             org_title, result_docnum, manufacturer,
                             search_query, row_index, id_pu,
                             contract_number, edo_code, balance_owner,
                             operation_responsibility, mpi)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            record.vri_id, record.mit_number, record.mit_title,
                            record.mit_notation, record.mi_modification, record.mi_number,
                            record.verification_date, record.valid_date,
                            1 if record.applicability else 0,
                            record.org_title, record.result_docnum,
                            record.manufacturer,
                            record.search_query, record.row_index, record.id_pu,
                            record.contract_number, record.edo_code, record.balance_owner,
                            record.operation_responsibility, record.mpi
                        ))
                        
                        if conn.total_changes > 0:
                            saved += 1
                            
                except Exception as e:
                    errors += 1
                    logger.error(f"❌ Ошибка сохранения {record.vri_id}: {e}")
            
            conn.commit()
            
        finally:
            conn.close()
        
        logger.info(f"💾 Сохранено: {saved} новых, {duplicates} дубликатов, {errors} ошибок")
        return {'saved': saved, 'duplicates': duplicates, 'errors': errors}
    
    def get_stats(self) -> Dict:
        """Получить детальную статистику"""
        with sqlite3.connect(self.db_path) as conn:
            total = conn.execute('SELECT COUNT(*) FROM verification_records').fetchone()[0]
            
            if total == 0:
                return {
                    'total': 0,
                    'by_year': {},
                    'by_manufacturer': {},
                    'applicable': 0,
                    'inapplicable': 0
                }
            
            # Статистика по годам
            # Дата в формате DD.MM.YYYY, извлекаем год через substr
            by_year = dict(conn.execute(
                "SELECT substr(verification_date, 7, 4) as year, COUNT(*) "
                "FROM verification_records WHERE verification_date != '' "
                "GROUP BY year ORDER BY year"
            ).fetchall())
            
            # Статистика по производителям
            by_manufacturer = dict(conn.execute(
                'SELECT manufacturer, COUNT(*) FROM verification_records '
                'GROUP BY manufacturer ORDER BY COUNT(*) DESC LIMIT 10'
            ).fetchall())
            
            # Статистика по пригодности
            applicable = conn.execute(
                'SELECT COUNT(*) FROM verification_records WHERE applicability = 1'
            ).fetchone()[0]
            
            return {
                'total': total,
                'by_year': by_year,
                'by_manufacturer': by_manufacturer,
                'applicable': applicable,
                'inapplicable': total - applicable
            }
    
    def get_records_by_query(self, search_query: str, row_index: int) -> List[Dict]:
        """Получить записи по поисковому запросу и индексу строки"""
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.execute('''
                SELECT * FROM verification_records
                WHERE search_query = ? AND row_index = ?
                ORDER BY verification_date DESC
            ''', (search_query, row_index))
            return [dict(row) for row in cursor.fetchall()]
    
    def export_to_csv(self, filename: str, include_service_fields: bool = True) -> int:
        """Экспорт в CSV с русскими заголовками"""
        import csv
        
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.execute('SELECT * FROM verification_records')
            columns = [desc[0] for desc in cursor.description]
            rows = cursor.fetchall()
            
            if not rows:
                logger.warning("⚠️  Нет данных для экспорта")
                return 0
            
            # Русские заголовки
            field_mapping = {
                'vri_id': 'ID',
                'mit_number': '№ реестра',
                'mit_title': 'Наименование',
                'mit_notation': 'Обозначение типа',
                'mi_modification': 'Модификация',
                'mi_number': 'Зав. номер',
                'verification_date': 'Дата поверки',
                'valid_date': 'Действ. до',
                'applicability': 'Пригоден',
                'org_title': 'Поверитель',
                'result_docnum': '№ документа',
                'manufacturer': 'Производитель',
                'search_query': 'Поисковый запрос',
                'row_index': '№ строки',
                'id_pu': 'Id_ПУ',
                'contract_number': 'Номер договора',
                'edo_code': 'Код ЭДО',
                'balance_owner': 'Балансовая принадлежность',
                'operation_responsibility': 'Эксплуатационная ответственность',
                'mpi': 'МПИ',
                'collected_at': 'Время сохранения',
                'record_url': 'URL записи'
            }
            
            with open(filename, 'w', encoding='utf-8-sig', newline='') as f:
                # QUOTE_ALL - все поля в кавычках для корректной обработки
                # Но переносы строк внутри полей заменяем на пробелы
                writer = csv.writer(f, delimiter=';', quoting=csv.QUOTE_ALL, lineterminator='\n')
                headers = [field_mapping.get(col, col) for col in columns]
                writer.writerow(headers)
                
                # Обработка строк: замена переносов на пробелы
                for row in rows:
                    cleaned_row = [str(v).replace('\n', ' ').replace('\r', ' ') if v else '' for v in row]
                    writer.writerow(cleaned_row)

            logger.info(f"📤 Экспортировано {len(rows)} записей в {filename}")
            return len(rows)
    
    def save_progress(self, input_file: str, last_index: int, total_queries: int, params: str = ""):
        """Сохранить прогресс обработки файла"""
        import os
        abs_path = os.path.abspath(input_file)
        with sqlite3.connect(self.db_path) as conn:
            conn.execute('''
                INSERT OR REPLACE INTO search_progress
                (input_file, last_processed_index, total_queries, params, updated_at)
                VALUES (?, ?, ?, ?, datetime('now', 'localtime'))
            ''', (abs_path, last_index, total_queries, params))
            conn.commit()

    def get_progress(self, input_file: str) -> dict:
        """Получить прогресс обработки файла. Возвращает None если прогресса нет."""
        import os
        abs_path = os.path.abspath(input_file)
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.execute(
                'SELECT * FROM search_progress WHERE input_file = ?',
                (abs_path,)
            )
            row = cursor.fetchone()
            if row:
                return dict(row)
            return None

    def clear_progress(self, input_file: str = None):
        """Очистить прогресс обработки (конкретного файла или весь)"""
        with sqlite3.connect(self.db_path) as conn:
            if input_file:
                import os
                abs_path = os.path.abspath(input_file)
                conn.execute('DELETE FROM search_progress WHERE input_file = ?', (abs_path,))
            else:
                conn.execute('DELETE FROM search_progress')
            conn.commit()

    def clear(self):
        """Очистка БД"""
        with sqlite3.connect(self.db_path) as conn:
            conn.execute('DELETE FROM verification_records')
            conn.execute('DELETE FROM search_progress')
            conn.commit()
        logger.info("🧹 База данных очищена")
