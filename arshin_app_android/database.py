# -*- coding: utf-8 -*-
"""
База данных Android-приложения
SQLite с проверкой уникальности записей
Адаптировано из Konsol_Excel/database.py
"""

import sqlite3
import logging
from typing import List, Dict, Set
from models import VerificationRecord
from config import Config

logger = logging.getLogger(__name__)


class Database:
    """SQLite база с проверкой на полную уникальность"""

    def __init__(self, db_path: str = None):
        self.db_path = db_path or Config.get_db_path()
        self._init_db()

    def _init_db(self):
        """Инициализация таблиц"""
        with sqlite3.connect(self.db_path) as conn:
            conn.execute('''
                CREATE TABLE IF NOT EXISTS verification_records (
                    vri_id TEXT,
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
                    manufacturer TEXT,
                    collected_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    search_query TEXT,
                    row_index INTEGER,
                    id_pu TEXT,
                    contract_number TEXT,
                    edo_code TEXT,
                    balance_owner TEXT,
                    operation_responsibility TEXT,
                    mpi TEXT,
                    PRIMARY KEY (vri_id, search_query, row_index, id_pu)
                )
            ''')

            conn.execute(
                'CREATE INDEX IF NOT EXISTS idx_manufacturer '
                'ON verification_records(manufacturer)'
            )
            conn.execute(
                'CREATE INDEX IF NOT EXISTS idx_verification_date '
                'ON verification_records(verification_date)'
            )
            conn.execute(
                'CREATE INDEX IF NOT EXISTS idx_search_query '
                'ON verification_records(search_query)'
            )
            conn.execute(
                'CREATE INDEX IF NOT EXISTS idx_verification_records_mi_number '
                'ON verification_records(mi_number)'
            )
            conn.commit()

        logger.info("База данных инициализирована")

    def get_existing_ids(self) -> Set[str]:
        """Получить все сохранённые vri_id"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.execute('SELECT vri_id FROM verification_records')
            ids = {row[0] for row in cursor.fetchall()}
        return ids

    def is_duplicate(self, record: VerificationRecord, conn: sqlite3.Connection) -> bool:
        """Проверка на полный дубль по всем ключевым полям"""
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
        """Пакетное сохранение с проверкой уникальности"""
        if not records:
            return {'saved': 0, 'duplicates': 0, 'errors': 0}

        saved = 0
        duplicates = 0
        errors = 0

        conn = sqlite3.connect(self.db_path)
        try:
            for record in records:
                try:
                    if self.is_duplicate(record, conn):
                        duplicates += 1
                        logger.debug(f"Пропущен дубль: vri_id={record.vri_id}")
                    else:
                        conn.execute('''
                            INSERT INTO verification_records
                            (vri_id, mit_number, mit_title, mit_notation,
                             mi_modification, mi_number, verification_date,
                             valid_date, applicability, org_title, result_docnum,
                             manufacturer, search_query, row_index, id_pu,
                             contract_number, edo_code, balance_owner,
                             operation_responsibility, mpi)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            record.vri_id, record.mit_number, record.mit_title,
                            record.mit_notation, record.mi_modification,
                            record.mi_number, record.verification_date,
                            record.valid_date, 1 if record.applicability else 0,
                            record.org_title, record.result_docnum,
                            record.manufacturer,
                            record.search_query, record.row_index, record.id_pu,
                            record.contract_number, record.edo_code,
                            record.balance_owner,
                            record.operation_responsibility, record.mpi
                        ))
                        saved += 1

                except Exception as e:
                    errors += 1
                    logger.error(f"Ошибка сохранения {record.vri_id}: {e}")

            conn.commit()
        finally:
            conn.close()

        return {'saved': saved, 'duplicates': duplicates, 'errors': errors}

    def get_stats(self) -> Dict:
        """Статистика БД"""
        with sqlite3.connect(self.db_path) as conn:
            total = conn.execute(
                'SELECT COUNT(*) FROM verification_records'
            ).fetchone()[0]

            if total == 0:
                return {
                    'total': 0,
                    'by_year': {},
                    'by_manufacturer': {},
                    'applicable': 0,
                    'inapplicable': 0
                }

            by_year = dict(conn.execute(
                "SELECT substr(verification_date, 7, 4) as year, COUNT(*) "
                "FROM verification_records WHERE verification_date != '' "
                "GROUP BY year ORDER BY year"
            ).fetchall())

            by_manufacturer = dict(conn.execute(
                'SELECT manufacturer, COUNT(*) FROM verification_records '
                'GROUP BY manufacturer ORDER BY COUNT(*) DESC LIMIT 10'
            ).fetchall())

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

    def get_all_records(self, limit: int = 500) -> List[Dict]:
        """Получить все записи (для отображения истории)"""
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.execute('''
                SELECT * FROM verification_records
                ORDER BY collected_at DESC LIMIT ?
            ''', (limit,))
            return [dict(row) for row in cursor.fetchall()]

    def get_search_history(self) -> List[str]:
        """Уникальные поисковые запросы (для автодополнения)"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.execute('''
                SELECT DISTINCT search_query FROM verification_records
                WHERE search_query != ''
                ORDER BY search_query
            ''')
            return [row[0] for row in cursor.fetchall()]

    def export_to_csv(self, filename: str) -> int:
        """Экспорт в CSV с русскими заголовками"""
        import csv

        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.execute('SELECT * FROM verification_records')
            columns = [desc[0] for desc in cursor.description]
            rows = cursor.fetchall()

            if not rows:
                logger.warning("Нет данных для экспорта")
                return 0

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
            }

            with open(filename, 'w', encoding='utf-8-sig', newline='') as f:
                writer = csv.writer(f, delimiter=';', quoting=csv.QUOTE_MINIMAL)
                headers = [field_mapping.get(col, col) for col in columns]
                writer.writerow(headers)

                for row in rows:
                    cleaned_row = [
                        str(v).replace('\n', ' ').replace('\r', ' ') if v else ''
                        for v in row
                    ]
                    writer.writerow(cleaned_row)

            logger.info(f"Экспортировано {len(rows)} записей в {filename}")
            return len(rows)

    def clear(self):
        """Очистка БД"""
        with sqlite3.connect(self.db_path) as conn:
            conn.execute('DELETE FROM verification_records')
            conn.commit()
        logger.info("База данных очищена")
