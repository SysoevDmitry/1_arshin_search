#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Обработчик Excel файлов
Адаптировано из arshin_app.py с автоматическим определением колонок
"""

import logging
from typing import Dict, List, Tuple, Any
from config import PANDAS_AVAILABLE

logger = logging.getLogger(__name__)


class ExcelHandler:
    """
    Работа с Excel файлами для пакетного поиска
    Адаптировано из arshin_app.py
    """
    
    @staticmethod
    def detect_columns(df) -> Dict[str, Tuple[str, int]]:
        if not PANDAS_AVAILABLE:
            logger.error("pandas не установлен")
            return {}

        import pandas as pd
        
        column_mapping = {}

        possible_columns = {
            'mi_number': [
                'серийный номер', 'заводской номер', 'номер пу', 'серийный номер пу',
                'зав. номер', 'зав №', 'serial number', 'device_number',
                'серийный', 'заводской', 'serial', 'pu_number', 'id_пу',
                'номер пу', 'прибор', 'счетчик', 'измеритель', 'id',
                'номер прибора', 'серийник'
            ],
            'mit_title': [
                'модель', 'model', 'наименование типа', 'тип прибора', 'тип пу',
                'наименование', 'название', 'name', 'тип си', 'устройство',
                'описание', 'прибор', 'наименование си', 'тип', 'модель прибора',
                'наименование прибора', 'тип си', 'название прибора'
            ],
            'mit_number': [
                'номер в реестре', 'реестровый номер', 'рег. номер', 'регистр',
                'номер типа', 'mit_number', 'vri_id', 'реестр', 'тип си номер'
            ],
            'org_title': [
                'организация поверитель', 'поверитель', 'организация', 'org_title',
                'org', 'кто поверял', 'поверка', 'company', 'предприятие'
            ],
            'vri_id': [
                'vri_id', 'vri', 'идентификатор', 'уникальный', 'uuid',
                'ключ', 'key', 'id записи'
            ],
            'search_term': [
                'поисковый запрос', 'ключевое слово', 'поиск', 'query', 'запрос',
                'search', 'термин', 'ключ'
            ],
            'year': [
                'год поверки', 'поверка год', 'год поверки си', 'verification year',
                'год', 'year', 'дата поверки'
            ],
            'manufacture_year': [
                'год выпуска', 'год производства', 'год изготовления', 'дата выпуска',
                'выпуск', 'manufacture', 'production year'
            ],
            'id_pu': ['id_пу', 'id пу', 'идентификатор пу'],
            'contract_number': ['номер договора', 'договор', 'contract', 'номер док'],
            'edo_code': ['код эдо', 'эдо', 'edo', 'edo код'],
            'balance_owner': ['балансовая принадлежность', 'баланс', 'владелец', 'баланс归属'],
            'operation_responsibility': ['эксплуатационная ответственность', 'эксплуатация', 'ответственность'],
            'mpi': ['мпи', 'межповерочный интервал', 'интервал', 'мпИ']
        }

        df_columns = []
        for col in df.columns:
            col_str = str(col).lower().strip()
            col_str = ' '.join(col_str.split())
            df_columns.append(col_str)

        logger.debug(f"Определение колонок: {list(df.columns)}")

        candidates = []
        for field, keywords in possible_columns.items():
            for idx, col_lower in enumerate(df_columns):
                score = 0
                original_col = df.columns[idx]

                if col_lower == field:
                    score = 100
                else:
                    for priority, keyword in enumerate(keywords):
                        if keyword == col_lower:
                            score = max(score, 95 - priority)
                        elif col_lower.startswith(keyword):
                            score = max(score, 80 - priority)
                        elif keyword in col_lower:
                            score = max(score, 60 - priority)
                        elif len(keyword) > 4 and keyword[:4] in col_lower:
                            score = max(score, 40 - priority)

                if score > 40:
                    candidates.append((score, field, original_col, idx))

        candidates.sort(key=lambda x: x[0], reverse=True)

        used_fields = set()
        used_cols = set()
        for score, field, col_name, idx in candidates:
            if field in used_fields or col_name in used_cols:
                continue
            column_mapping[field] = (col_name, idx)
            used_fields.add(field)
            used_cols.add(col_name)
            logger.debug(f"✅ Найдено: {field} -> {col_name} (score={score})")

        if 'mi_number' not in column_mapping:
            for idx, col_lower in enumerate(df_columns):
                if any(k in col_lower for k in ['серийн', 'завод', 'номер', 'serial']):
                    column_mapping['mi_number'] = (df.columns[idx], idx)
                    logger.debug(f"✅ mi_number найдена по ключу: {df.columns[idx]}")
                    break
        
        if 'mit_title' not in column_mapping:
            for idx, col_lower in enumerate(df_columns):
                if any(k in col_lower for k in ['модель', 'тип', 'наимен', 'model']):
                    column_mapping['mit_title'] = (df.columns[idx], idx)
                    logger.debug(f"✅ mit_title найдена по ключу: {df.columns[idx]}")
                    break

        logger.info(f"Автоматически определено колонок: {len(column_mapping)}: "
                    f"{', '.join(f'{f}→{c}' for f, (c, _) in column_mapping.items())}")
        return column_mapping
    
    @staticmethod
    def read_queries(filename: str) -> List[Dict]:
        if not PANDAS_AVAILABLE:
            logger.error("pandas не установлен")
            raise ImportError("pandas не установлен")

        import pandas as pd

        logger.info(f"📂 Чтение Excel файла: {filename}")
        
        df_check = pd.read_excel(filename, nrows=1, dtype=str)
        
        first_col = str(df_check.columns[0]) if len(df_check.columns) > 0 else ""
        if first_col.startswith('Unnamed'):
            logger.info("ℹ️  Заголовки найдены во второй строке (пропуск первой строки)")
            df = pd.read_excel(filename, header=1, dtype=str)
        else:
            df = pd.read_excel(filename, dtype=str)
        
        logger.debug(f"Прочитано {len(df)} строк, колонки: {list(df.columns)}")
        
        column_mapping = ExcelHandler.detect_columns(df)
        
        if not column_mapping:
            logger.warning("⚠️  Колонки не определены, используется первая колонка")
            column_mapping = {'search_term': (df.columns[0], 0)}
        
        queries = []
        for idx, row in df.iterrows():
            query = {
                'row_index': idx + 1,
                'original_data': {},
            }
            
            for col in df.columns:
                value = row[col]
                if value and str(value).lower() != 'nan':
                    query['original_data'][col] = str(value)
            
            for field, (col_name, _) in column_mapping.items():
                value = row[col_name]
                if value and str(value).lower() != 'nan':
                    query[field] = str(value)
                    # Дублируем в original_data под именем поля
                    # (from_api ищет по field-имени, а не по сырому имени колонки)
                    query['original_data'][field] = str(value)
            
            if 'search_term' in query or 'mi_number' in query:
                queries.append(query)
        
        logger.info(f"✅ Загружено {len(queries)} запросов из Excel")
        return queries
    
    @staticmethod
    def create_template(filename: str = None) -> str:
        if not PANDAS_AVAILABLE:
            logger.warning("pandas не установлен")
            return ""
        
        import pandas as pd
        import os
        from datetime import datetime
        
        if not filename:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"exports/batch_template_{timestamp}.xlsx"
        
        os.makedirs(os.path.dirname(filename) if os.path.dirname(filename) else '.', exist_ok=True)
        
        data = {
            'Заводской номер': ['12345678', '87654321', ''],
            'Наименование': ['Счетчик электрической энергии Меркурий 201', '', ''],
            'Номер в реестре': ['77310-20', '', ''],
            'Id_ПУ': ['ПУ-001', 'ПУ-002', ''],
            'Номер договора': ['Д-001/2023', 'Д-002/2023', ''],
            'Код ЭДО': ['123456789', '987654321', ''],
            'МПИ': ['16', '12', ''],
            'Примечание': ['Первый запрос', 'Второй запрос', '']
        }
        
        df = pd.DataFrame(data)
        
        try:
            from openpyxl import Workbook
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Шаблон')
                
                worksheet = writer.sheets['Шаблон']
                worksheet.column_dimensions['A'].width = 20
                worksheet.column_dimensions['B'].width = 45
                worksheet.column_dimensions['C'].width = 15
                worksheet.column_dimensions['D'].width = 15
                worksheet.column_dimensions['E'].width = 20
                worksheet.column_dimensions['F'].width = 15
                worksheet.column_dimensions['G'].width = 10
                worksheet.column_dimensions['H'].width = 25
            
            logger.info(f"📄 Шаблон создан: {filename}")
        except ImportError:
            logger.warning("openpyxl не установлен, создаю CSV шаблон")
            csv_filename = filename.replace('.xlsx', '.csv')
            df.to_csv(csv_filename, index=False, sep=';', encoding='utf-8-sig')
            logger.info(f"📄 CSV шаблон создан: {csv_filename}")
            filename = csv_filename
        
        return filename
