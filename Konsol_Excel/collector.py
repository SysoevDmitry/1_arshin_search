#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Сборщик данных для Excel версии
Поиск по серийному номеру с фильтрацией по электросчетчикам

Обновлено: поддержка verification_date и атрибутивных фильтров
"""

import asyncio
import logging
import re
from typing import List, Dict, Optional, Set
from config import Config, TQDM_AVAILABLE, EXACT_QUERIES, MANUFACTURERS_RULES
from database import Database
from models import VerificationRecord
from api_client import ParallelAPIClient

if TQDM_AVAILABLE:
    from tqdm import tqdm

logger = logging.getLogger(__name__)


def normalize_serial_number(serial: str) -> str:
    """Нормализация серийного номера для сравнения"""
    if not serial:
        return ""
    # Удаляем пробелы, дефисы, приводим к верхнему регистру
    normalized = re.sub(r'[\s\-]', '', serial.upper())
    return normalized


def match_serial_number(record_serial: str, target_serial: str) -> bool:
    """
    Проверка точного совпадения серийного номера
    
    Args:
        record_serial: Серийный номер из записи API
        target_serial: Целевой серийный номер из Excel
    
    Returns:
        True если точно совпадает
    """
    if not target_serial or not record_serial:
        return False
    
    norm_record = normalize_serial_number(record_serial)
    norm_target = normalize_serial_number(target_serial)
    
    if norm_record == norm_target:
        return True
    
    # Совпадение без ведущих нулей
    # (Excel/API могут различаться: "023241147596" vs "23241147596")
    return norm_record.lstrip('0') == norm_target.lstrip('0')


def load_electric_phrases_from_queries() -> tuple:
    """
    Загрузка фраз и слов для фильтрации электросчетчиков из exact_queries.csv

    Returns:
        (include_phrases, exclude_phrases) - кортеж из двух списков
    """
    include_phrases = []
    exclude_phrases = []

    # Фразы для исключения (не электросчетчики)
    exclude_patterns = [
        # Вода
        'воды', 'водомер', 'водосчетчик', 'холодной воды', 'горячей воды',
        'счетчик воды', 'крыльчатые',
        # Газ
        'газа', 'газ', 'газосчетчик', 'газовые', 'бытовые газовые', 'диафрагменные',
        'счетчик газа',
        # Тепло
        'тепл', 'теплопотребление', 'теплопотребления', 'тепловычислитель',
        'вычислитель количества теплоты', 'распределения теплопотребления',
        # Медицинские приборы
        'термометры медицинские', 'термометр медицинский', 'медицинские максимальные',
        'ртутные стеклянные медицинские', 'максимальные стеклянные ртутные',
        # Весы
        'весы', 'взвешивания', 'для новорожденных', 'почтовые электронные',
        'рычажные настольные',
        # Давление
        'манометры', 'вакуумметры', 'мановакуумметры', 'напоромеры', 'тягомеры',
        'тягонапоромеры', 'обм',
        # Температура (не медицинские)
        'термометры ртутные', 'термометры стеклянные', 'термометры лабораторные',
        'термопреобразователи', 'термометры сопротивления',
        # Газовые сигнализаторы
        'сигнализаторы', 'метана', 'с внешним сенсором',
        # Прочие измерительные приборы
        'штангенглубиномеры', 'нутромеры', 'ареометры', 'гигрометры',
        'психрометрические', 'мегаомметры', 'приемники-ловушки',
        'счетчики времени наработки', 'бутиромеры', 'поверка си массы',
        'расхода электромагнитные', 'преобразователи расхода',
        'устройства для распределения', 'для статического взвешивания',
        # Комплексы и приборы для мониторинга радиации
        'комплексы измерительные', 'мониторинга радона', 'мониторинга торона',
        'дочерних продуктов', 'альфарад',
        # Приборы щитовые
        'приборы щитовые', 'щитовые цифровые', 'электроизмерительные многофункциональные',
        'параметров и показателей качества', 'щмк',
    ]

    for phrase in EXACT_QUERIES:
        phrase_lower = phrase.lower()
        # Проверяем, что это категория electric
        include_phrases.append(phrase_lower)

    # Добавляем ключевые слова-маркеры из фраз
    electric_markers = [
        'электрической энергии', 'электроэнергии', 'электросчетчик',
        'счетчик электрической', 'счетчик электроэнергии',
        'активной электроэнергии', 'реактивной электроэнергии',
        'однофазный', 'трехфазный', 'статический', 'индукционный',
        'меркурий', 'нева', 'энергомера', 'матрица', 'альфа', 'милур',
        'псч', 'сэт', 'цэ', 'се 101', 'се 102', 'се 201', 'се 301',
        'меркурий 20', 'меркурий 23', 'нева 10', 'нева 30',
        'трансформатор тока измерительный', 'трансформатор тока 0,4 кв',
        'трансформатор тока 0.4 кв',
    ]
    include_phrases.extend(electric_markers)

    exclude_phrases = exclude_patterns

    return include_phrases, exclude_phrases


# Глобальные списки для фильтрации
ELECTRIC_INCLUDE_PHRASES, ELECTRIC_EXCLUDE_PHRASES = load_electric_phrases_from_queries()


def is_electric_meter_from_queries(title: str, notation: str = "", modification: str = "") -> bool:
    """
    Проверка: является ли запись электросчетчиком
    Использует фразы и слова из exact_queries.csv

    Логика:
    1. Сначала проверяем исключения (вода, газ, тепло, медицинские и т.д.)
    2. Затем проверяем включения по фразам из exact_queries.csv

    Args:
        title: Наименование типа СИ
        notation: Обозначение типа
        modification: Модификация

    Returns:
        True если это электросчетчик
    """
    combined = f"{title} {notation} {modification}".lower()

    if not combined:
        return False

    # Шаг 1: Проверка исключений (не электросчетчики)
    for keyword in ELECTRIC_EXCLUDE_PHRASES:
        if keyword in combined:
            logger.debug(f"Исключено (не электросчетчик): {title}")
            return False

    # Шаг 2: Проверка включений по фразам из exact_queries.csv
    for phrase in ELECTRIC_INCLUDE_PHRASES:
        if phrase in combined:
            logger.debug(f"Электросчетчик найден по фразе '{phrase}': {title}")
            return True

    return False


class ExcelCollector:
    """
    Сборщик данных из Excel файла
    
    Логика работы:
    1. Поиск по серийному номеру через API
    2. Фильтрация результатов: только электросчетчики
    3. Фильтрация: точное совпадение серийного номера
    4. Сохранение с привязкой к данным из Excel (Id_ПУ, № строки)
    """

    def __init__(self, db: Database = None, input_file: str = None):
        self.db = db or Database()
        self.api: Optional[ParallelAPIClient] = None
        self.pbar = None
        self.input_file = input_file
        self.save_progress_params = ""
        self.stats = {
            'queries_processed': 0,
            'records_found': 0,
            'records_saved': 0,
            'duplicates': 0,
            'filtered_not_electric': 0,
            'filtered_serial_mismatch': 0,
            'no_match': 0
        }

    async def __aenter__(self):
        self.api = await ParallelAPIClient().__aenter__()
        return self

    async def __aexit__(self, *args):
        await self.api.__aexit__(*args)
        if self.pbar:
            self.pbar.close()

    @staticmethod
    def _adjust_years_for_query(query: Dict, years: List[int]) -> List[int]:
        """
        Корректировка диапазона годов по году выпуска прибора.
        Если в строке Excel указан год выпуска — поиск идёт от года выпуска
        до последнего года из диапазона (или до текущего).
        """
        manufacture_year_str = query.get('manufacture_year', '')
        if not manufacture_year_str or manufacture_year_str.lower() == 'nan':
            return years

        try:
            mfg_year = int(float(manufacture_year_str))
        except (ValueError, TypeError):
            return years

        from datetime import datetime
        current_year = datetime.now().year
        max_year = max(years) if years else current_year

        start_year = max(mfg_year, min(years) if years else mfg_year)
        if start_year > max_year:
            return [start_year]

        return [y for y in years if y >= start_year]

    async def process_query(self, serial_number: str, years: List[int],
                            row_index: int = 0, id_pu: str = "",
                            original_data: dict = None,
                            verification_date: str = None,
                            use_attribute_search: bool = True) -> int:
        """
        Обработка одного запроса из Excel

        Логика:
        1. Поиск по серийному номеру за каждый год (или по verification_date)
        2. Фильтрация: только электросчетчики (по ключевым словам)
        3. Фильтрация: точное совпадение серийного номера

        Args:
            serial_number: Серийный номер для поиска
            years: Список годов для поиска
            row_index: Индекс строки в Excel (1-based)
            id_pu: Id_ПУ из исходного файла
            original_data: Исходные данные строки
            verification_date: Конкретная дата поверки (вместо year)
            use_attribute_search: Использовать атрибутивный поиск (рекомендуется)

        Returns:
            Количество найденных записей
        """
        records_buffer: List[VerificationRecord] = []
        found_count = 0

        if not serial_number or serial_number.lower() == 'nan':
            logger.debug(f"⏭  Пропущено (нет серийного номера), строка {row_index}")
            return 0

        logger.debug(f"🔍 Поиск по серийному номеру: '{serial_number}'")

        # Атрибутивный поиск по mi_number (рекомендуется спецификацией)
        if use_attribute_search:
            # Используем фильтр по mi_number вместо search
            filters = {'mi_number': serial_number}

            # Поиск с атрибутивным фильтром
            if verification_date:
                # Поиск по конкретной дате
                result = await self.api.search_with_retry(
                    search_term=None,  # Не используем search
                    filters=filters,
                    verification_date=verification_date
                )
            else:
                # Поиск по годам
                for year in years:
                    result = await self.api.search_with_retry(
                        search_term=None,  # Не используем search
                        year=year,
                        filters=filters
                    )
                    items = result.get('items', [])
                    if items:
                        logger.debug(f"  Год {year}: найдено {len(items)} записей (атрибутивный поиск)")

                        # Обработка результатов
                        for item in items:
                            await self._process_api_item(item, serial_number, records_buffer,
                                                        row_index, id_pu, original_data)

                    # Пауза между запросами
                    await asyncio.sleep(Config.REQUEST_DELAY)

                # Сохранение в БД
                if records_buffer:
                    save_result = self.db.save_records_batch(records_buffer)
                    self.stats['records_saved'] += save_result['saved']
                    self.stats['duplicates'] += save_result['duplicates']
                    logger.info(f"  💾 Сохранено: {save_result['saved']} новых, {save_result['duplicates']} дублей")
                elif found_count == 0:
                    self.stats['no_match'] += 1

                self.stats['records_found'] += len(records_buffer)
                self.stats['queries_processed'] += 1

                return len(records_buffer)

        # Резервный вариант: обычный поиск по search параметру
        for year in (years if not verification_date else [None]):
            if verification_date:
                result = await self.api.search_with_retry(
                    search_term=serial_number,
                    verification_date=verification_date
                )
            else:
                result = await self.api.search_with_retry(serial_number, year)

            items = result.get('items', [])

            if not items:
                if verification_date:
                    logger.debug(f"  Дата {verification_date}: ничего не найдено")
                else:
                    logger.debug(f"  Год {year}: ничего не найдено")
                continue

            if verification_date:
                logger.debug(f"  Дата {verification_date}: найдено {len(items)} записей")
            else:
                logger.debug(f"  Год {year}: найдено {len(items)} записей")

            # Обработка результатов
            for item in items:
                await self._process_api_item(item, serial_number, records_buffer,
                                            row_index, id_pu, original_data)

            # Пауза между запросами
            await asyncio.sleep(Config.REQUEST_DELAY)

        # Сохранение в БД
        if records_buffer:
            save_result = self.db.save_records_batch(records_buffer)
            self.stats['records_saved'] += save_result['saved']
            self.stats['duplicates'] += save_result['duplicates']
            logger.info(f"  💾 Сохранено: {save_result['saved']} новых, {save_result['duplicates']} дублей")
        elif found_count == 0:
            self.stats['no_match'] += 1

        self.stats['records_found'] += found_count
        self.stats['queries_processed'] += 1

        return found_count

    async def _process_api_item(self, item: Dict, serial_number: str,
                                 records_buffer: List, row_index: int,
                                 id_pu: str, original_data: dict) -> bool:
        """
        Обработка одной записи из API

        Args:
            item: Данные из API
            serial_number: Целевой серийный номер
            records_buffer: Буфер для сохранения
            row_index: Индекс строки
            id_pu: Id_ПУ
            original_data: Исходные данные

        Returns:
            True если запись добавлена в буфер
        """
        mit_title = item.get('mit_title', '')
        mit_notation = item.get('mit_notation', '')
        mi_modification = item.get('mi_modification', '')
        record_serial = item.get('mi_number', '')

        # Фильтр 1: только электросчетчики
        if not is_electric_meter_from_queries(mit_title, mit_notation, mi_modification):
            self.stats['filtered_not_electric'] += 1
            logger.debug(f"  ❌ Исключено (не электросчетчик): {mit_title}")
            return False

        # Фильтр 2: точное совпадение серийного номера
        if not match_serial_number(record_serial, serial_number):
            self.stats['filtered_serial_mismatch'] += 1
            logger.debug(f"  ❌ Несовпадение серийного: {record_serial} != {serial_number}")
            return False

        # Создание записи с сохранением связей
        record = VerificationRecord.from_api(
            item,
            search_query=serial_number,
            row_index=row_index,
            id_pu=id_pu,
            original_data=original_data
        )
        records_buffer.append(record)
        logger.debug(f"  ✅ Найдено: {mit_title}, сер.№ {record_serial}")
        return True

    async def process_queries_batch(self, queries: List[Dict], years: List[int],
                                     batch_size: int = 5,
                                     verification_date: str = None,
                                     use_attribute_search: bool = True,
                                     start_from: int = 0) -> Dict:
        """
        Пакетная обработка запросов

        Args:
            queries: Список запросов из Excel
            years: Годы для поиска
            batch_size: Размер пакета для параллельной обработки
            verification_date: Дата поверки (вместо year)
            use_attribute_search: Использовать атрибутивный поиск
            start_from: Индекс, с которого начать обработку (для продолжения после сбоя)

        Returns:
            Статистика обработки
        """
        total_queries = len(queries)

        if verification_date:
            logger.info(f"🔍 Обработка {total_queries} запросов на дату {verification_date}"
                       + (f" (с индекса {start_from})" if start_from > 0 else ""))
        else:
            logger.info(f"🔍 Обработка {total_queries} запросов за {len(years)} годов"
                       + (f" (с индекса {start_from})" if start_from > 0 else ""))

        # Progress bar
        if TQDM_AVAILABLE:
            self.pbar = tqdm(total=total_queries, initial=start_from,
                           desc="Запросы", unit="запрос")

        # Обработка запросов
        for i, query in enumerate(queries):
            # Пропуск уже обработанных запросов (при продолжении после сбоя)
            if i < start_from:
                if self.pbar:
                    self.pbar.update(1)
                continue

            # Серийный номер - основное поле для поиска
            serial_number = query.get('mi_number', '')

            if not serial_number or serial_number.lower() == 'nan':
                logger.debug(f"⏭  Пропущено (нет серийного номера), строка {query.get('row_index', i+1)}")
                # Сохраняем прогресс даже для пропущенных запросов
                if self.input_file:
                    self.save_progress_params = self._build_progress_params(
                        years, verification_date, use_attribute_search
                    )
                    self.db.save_progress(
                        self.input_file, i, total_queries,
                        self.save_progress_params
                    )
                if self.pbar:
                    self.pbar.update(1)
                continue

            row_index = query.get('row_index', i + 1)
            id_pu = query.get('id_pu', '')
            original_data = query.get('original_data', {})

            # Корректировка годов поиска по году выпуска прибора
            query_years = self._adjust_years_for_query(query, years)

            if query_years != years:
                mfg_year = query.get('manufacture_year', '')
                logger.info(f"  📅 Строка {row_index}: год выпуска {mfg_year}, "
                           f"поиск {min(query_years)}–{max(query_years)} "
                           f"(из диапазона {min(years)}–{max(years)})")

            # Обработка запроса с новыми параметрами
            found = await self.process_query(
                serial_number, query_years,
                row_index=row_index,
                id_pu=id_pu,
                original_data=original_data,
                verification_date=verification_date,
                use_attribute_search=use_attribute_search
            )

            serial_short = serial_number[:20] + '...' if len(serial_number) > 20 else serial_number
            year_range = f"{min(query_years)}–{max(query_years)}" if not verification_date else verification_date
            logger.info(f"[{i+1}/{total_queries}] '{serial_short}' ({year_range}) → {found} записей")

            # Сохранение прогресса после каждого успешно обработанного запроса
            if self.input_file:
                self.save_progress_params = self._build_progress_params(
                    years, verification_date, use_attribute_search
                )
                self.db.save_progress(
                    self.input_file, i, total_queries,
                    self.save_progress_params
                )

            if self.pbar:
                self.pbar.update(1)
                self.pbar.set_postfix({
                    "найдено": self.stats['records_found'],
                    "сохранено": self.stats['records_saved']
                })

        if self.pbar:
            self.pbar.close()

        # Очистка прогресса после успешного завершения
        if self.input_file:
            self.db.clear_progress(self.input_file)

        return self.stats

    @staticmethod
    def _build_progress_params(years: List[int], verification_date: str,
                               use_attribute_search: bool) -> str:
        """Формирование строки параметров для сохранения в прогрессе"""
        if verification_date:
            return f"date={verification_date},attr={use_attribute_search}"
        return f"years={min(years)}-{max(years)},attr={use_attribute_search}"

    def get_stats(self) -> Dict:
        """Получить статистику обработки"""
        db_stats = self.db.get_stats()
        return {
            **self.stats,
            'db_total': db_stats['total'],
            'db_applicable': db_stats['applicable']
        }
