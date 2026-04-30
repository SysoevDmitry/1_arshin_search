# -*- coding: utf-8 -*-
"""
Обработчик Excel для Android-приложения
Чтение файлов .xlsx/.xls с телефона
Адаптировано из Konsol_Excel/excel_handler.py
"""

import logging
from typing import Dict, List, Tuple, Any

logger = logging.getLogger(__name__)


class ExcelHandler:
    """Работа с Excel файлами для пакетного поиска"""

    @staticmethod
    def detect_columns(df) -> Dict:
        """
        Автоматическое определение колонок по заголовкам

        Returns:
            Dict[field_name: (column_name, column_index)]
        """
        import pandas as pd

        column_mapping = {}

        possible_columns = {
            'mi_number': [
                'серийный номер', 'заводской номер', 'номер пу',
                'серийный номер пу', 'зав. номер', 'зав №',
                'serial number', 'device_number', 'серийный',
                'заводской', 'serial', 'pu_number', 'id_пу',
                'номер пу', 'прибор', 'счетчик', 'измеритель',
                'номер прибора', 'серийник', 'заводской №',
            ],
            'mit_title': [
                'модель', 'model', 'наименование типа', 'тип прибора',
                'тип пу', 'наименование', 'название', 'name', 'тип си',
                'устройство', 'описание', 'прибор', 'наименование си',
                'модель прибора', 'наименование прибора', 'тип си',
            ],
            'mit_number': [
                'номер в реестре', 'реестровый номер', 'рег. номер',
                'регистр', 'номер типа', 'mit_number', 'реестр',
            ],
            'org_title': [
                'организация поверитель', 'поверитель', 'организация',
                'org_title', 'org', 'кто поверял', 'поверка',
            ],
            'manufacture_year': [
                'год выпуска', 'год производства', 'год изготовления',
                'дата выпуска', 'выпуск', 'manufacture',
            ],
            'id_pu': ['id_пу', 'id пу', 'идентификатор пу'],
            'contract_number': ['номер договора', 'договор', 'contract'],
            'edo_code': ['код эдо', 'эдо', 'edo'],
            'balance_owner': ['балансовая принадлежность', 'баланс', 'владелец'],
            'operation_responsibility': [
                'эксплуатационная ответственность', 'эксплуатация',
                'ответственность',
            ],
            'mpi': ['мпи', 'межповерочный интервал', 'интервал'],
        }

        df_columns = []
        for col in df.columns:
            col_str = str(col).lower().strip()
            col_str = ' '.join(col_str.split())
            df_columns.append(col_str)

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
                        elif (
                            len(keyword) > 4
                            and keyword[:4] in col_lower
                        ):
                            score = max(score, 40 - priority)

                if score > 40:
                    candidates.append(
                        (score, field, original_col, idx)
                    )

        candidates.sort(key=lambda x: x[0], reverse=True)

        used_fields = set()
        used_cols = set()
        for score, field, col_name, idx in candidates:
            if field in used_fields or col_name in used_cols:
                continue
            column_mapping[field] = (col_name, idx)
            used_fields.add(field)
            used_cols.add(col_name)

        if 'mi_number' not in column_mapping:
            for idx, col_lower in enumerate(df_columns):
                if any(
                    k in col_lower
                    for k in ['серийн', 'завод', 'номер', 'serial']
                ):
                    column_mapping['mi_number'] = (df.columns[idx], idx)
                    break

        if 'mit_title' not in column_mapping:
            for idx, col_lower in enumerate(df_columns):
                if any(
                    k in col_lower
                    for k in ['модель', 'тип', 'наимен', 'model']
                ):
                    column_mapping['mit_title'] = (df.columns[idx], idx)
                    break

        logger.info(
            f"Определено колонок: {len(column_mapping)}: "
            f"{', '.join(f'{f}→{c}' for f, (c, _) in column_mapping.items())}"
        )
        return column_mapping

    @staticmethod
    def read_queries(filepath: str) -> List[Dict]:
        """
        Чтение запросов из Excel с сохранением исходных связей

        Returns:
            Список словарей с полями + row_index + original_data
        """
        import pandas as pd

        logger.info(f"Чтение Excel файла: {filepath}")

        df_check = pd.read_excel(filepath, nrows=1, dtype=str)

        first_col = (
            str(df_check.columns[0])
            if len(df_check.columns) > 0
            else ""
        )
        if first_col.startswith('Unnamed'):
            logger.info("Заголовки найдены во второй строке")
            df = pd.read_excel(filepath, header=1, dtype=str)
        else:
            df = pd.read_excel(filepath, dtype=str)

        logger.debug(
            f"Прочитано {len(df)} строк, колонки: {list(df.columns)}"
        )

        column_mapping = ExcelHandler.detect_columns(df)

        if not column_mapping:
            logger.warning("Колонки не определены, используется первая")
            column_mapping = {
                'search_term': (df.columns[0], 0)
            }

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

            if 'search_term' in query or 'mi_number' in query:
                queries.append(query)

        logger.info(f"Загружено {len(queries)} запросов из Excel")
        return queries
