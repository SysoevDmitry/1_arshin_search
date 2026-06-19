#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Сборщик данных для Excel версии
Поиск по серийному номеру с фильтрацией по электросчетчикам

Объединённая версия v6.3:
- DB-прогресс (search_progress) — надёжное продолжение после сбоя
- UPSERT (INSERT OR REPLACE) — перезапись существующих записей
- Атрибутивный поиск + verification_date
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
    if not serial:
        return ""
    normalized = re.sub(r'[\s\-]', '', serial.upper())
    return normalized


def match_serial_number(record_serial: str, target_serial: str) -> bool:
    if not target_serial or not record_serial:
        return False
    
    norm_record = normalize_serial_number(record_serial)
    norm_target = normalize_serial_number(target_serial)
    
    if norm_record == norm_target:
        return True
    
    return norm_record.lstrip('0') == norm_target.lstrip('0')


def load_electric_phrases_from_queries() -> tuple:
    include_phrases = []
    exclude_phrases = []

    exclude_patterns = [
        'воды', 'водомер', 'водосчетчик', 'холодной воды', 'горячей воды',
        'счетчик воды', 'крыльчатые',
        'газа', 'газ', 'газосчетчик', 'газовые', 'бытовые газовые', 'диафрагменные',
        'счетчик газа',
        'тепл', 'теплопотребление', 'теплопотребления', 'тепловычислитель',
        'вычислитель количества теплоты', 'распределения теплопотребления',
        'термометры медицинские', 'термометр медицинский', 'медицинские максимальные',
        'ртутные стеклянные медицинские', 'максимальные стеклянные ртутные',
        'весы', 'взвешивания', 'для новорожденных', 'почтовые электронные',
        'рычажные настольные',
        'манометры', 'вакуумметры', 'мановакуумметры', 'напоромеры', 'тягомеры',
        'тягонапоромеры', 'обм',
        'термометры ртутные', 'термометры стеклянные', 'термометры лабораторные',
        'термопреобразователи', 'термометры сопротивления',
        'сигнализаторы', 'метана', 'с внешним сенсором',
        'штангенглубиномеры', 'нутромеры', 'ареометры', 'гигрометры',
        'психрометрические', 'мегаомметры', 'приемники-ловушки',
        'счетчики времени наработки', 'бутиромеры', 'поверка си массы',
        'расхода электромагнитные', 'преобразователи расхода',
        'устройства для распределения', 'для статического взвешивания',
        'комплексы измерительные', 'мониторинга радона', 'мониторинга торона',
        'дочерних продуктов', 'альфарад',
        'приборы щитовые', 'щитовые цифровые', 'электроизмерительные многофункциональные',
        'параметров и показателей качества', 'щмк',
    ]

    for phrase in EXACT_QUERIES:
        phrase_lower = phrase.lower()
        include_phrases.append(phrase_lower)

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


ELECTRIC_INCLUDE_PHRASES, ELECTRIC_EXCLUDE_PHRASES = load_electric_phrases_from_queries()


def is_electric_meter_from_queries(title: str, notation: str = "", modification: str = "") -> bool:
    combined = f"{title} {notation} {modification}".lower()

    if not combined:
        return False

    for keyword in ELECTRIC_EXCLUDE_PHRASES:
        if keyword in combined:
            logger.debug(f"Исключено (не электросчетчик): {title}")
            return False

    for phrase in ELECTRIC_INCLUDE_PHRASES:
        if phrase in combined:
            logger.debug(f"Электросчетчик найден по фразе '{phrase}': {title}")
            return True

    return False


class ExcelCollector:
    """
    Сборщик данных из Excel файла
    
    Логика работы:
    1. Поиск по серийному номеру через API (атрибутивный)
    2. Фильтрация результатов: только электросчетчики
    3. Фильтрация: точное совпадение серийного номера
    4. Сохранение с привязкой к данным из Excel (Id_ПУ, № строки)
    5. UPSERT: перезапись существующих записей свежими данными
    6. DB-прогресс: атомарное сохранение позиции после каждого запроса
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
            'updated': 0,
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
    def _extract_year_from_date(date_str: str) -> Optional[int]:
        if not date_str or str(date_str).lower() == 'nan':
            return None
        import re
        match = re.search(r'(\d{4})', str(date_str))
        if match:
            year = int(match.group(1))
            if 1900 <= year <= 2100:
                return year
        return None

    @staticmethod
    def _compute_base_year_from_gosverka(queries: List[Dict]) -> Optional[int]:
        from collections import Counter
        years_list = []
        for q in queries:
            gd = q.get('gosverka_date', '')
            year = ExcelCollector._extract_year_from_date(gd)
            if year:
                years_list.append(year)
        if not years_list:
            return None
        counter = Counter(years_list)
        most_common, cnt = counter.most_common(1)[0]
        logger.info(f"📅 Мода «Дата госповерки»: {most_common} (встречается {cnt} раз(а) из {len(years_list)} значений)")
        return most_common

    @staticmethod
    def _adjust_years_for_query(query: Dict, years: List[int],
                                fallback_year: int = None,
                                auto_range: bool = False) -> List[int]:
        if auto_range:
            ver_year_str = query.get('verification_year', '')
            if not ver_year_str or str(ver_year_str).lower() == 'nan':
                ver_year_str = query.get('year', '')

            if ver_year_str and str(ver_year_str).lower() != 'nan':
                try:
                    vy = int(float(ver_year_str))
                    nachalo = vy - 1
                    konec = vy + 2
                    logger.info(f"  📅 Строка {query.get('row_index', '?')}: год поверки {vy} → диапазон {nachalo}–{konec}")
                    return list(range(nachalo, konec + 1))
                except (ValueError, TypeError):
                    pass

            manufacture_year_str = query.get('manufacture_year', '')
            if manufacture_year_str and str(manufacture_year_str).lower() != 'nan':
                try:
                    my = int(float(manufacture_year_str))
                    nachalo = my - 1
                    konec = my + 2
                    logger.info(f"  📅 Строка {query.get('row_index', '?')}: год выпуска {my} → диапазон {nachalo}–{konec}")
                    return list(range(nachalo, konec + 1))
                except (ValueError, TypeError):
                    pass

            if fallback_year is not None:
                nachalo = fallback_year - 1
                konec = fallback_year + 2
                logger.info(f"  📅 Строка {query.get('row_index', '?')}: базовый год {fallback_year} (мода «Дата госповерки») → диапазон {nachalo}–{konec}")
                return list(range(nachalo, konec + 1))

            if years:
                logger.info(f"  📅 Строка {query.get('row_index', '?')}: глобальный диапазон {min(years)}–{max(years)}")
                return years

            from datetime import datetime
            tek = datetime.now().year
            nachalo = tek - 6
            konec = tek
            logger.info(f"  📅 Строка {query.get('row_index', '?')}: ничего не найдено → последние 7 лет ({nachalo}–{konec})")
            return list(range(nachalo, konec + 1))

        # Старый режим (без --auto-range): manufacture_year сужает глобальный диапазон
        manufacture_year_str = query.get('manufacture_year', '')
        if not manufacture_year_str or str(manufacture_year_str).lower() == 'nan':
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
        records_buffer: List[VerificationRecord] = []
        found_count = 0

        if not serial_number or serial_number.lower() == 'nan':
            logger.debug(f"⏭  Пропущено (нет серийного номера), строка {row_index}")
            return 0

        logger.debug(f"🔍 Поиск по серийному номеру: '{serial_number}'")

        if use_attribute_search:
            filters = {'mi_number': serial_number}

            if verification_date:
                result = await self.api.search_with_retry(
                    search_term=None,
                    filters=filters,
                    verification_date=verification_date
                )
            else:
                for year in years:
                    result = await self.api.search_with_retry(
                        search_term=None,
                        year=year,
                        filters=filters
                    )
                    items = result.get('items', [])
                    if items:
                        logger.debug(f"  Год {year}: найдено {len(items)} записей (атрибутивный поиск)")

                        for item in items:
                            await self._process_api_item(item, serial_number, records_buffer,
                                                        row_index, id_pu, original_data)

                    await asyncio.sleep(Config.REQUEST_DELAY)

                if records_buffer:
                    save_result = self.db.save_records_batch(records_buffer)
                    self.stats['records_saved'] += save_result['saved']
                    self.stats['updated'] += save_result.get('updated', 0)
                    logger.info(f"  💾 Сохранено: {save_result['saved']} новых, {save_result.get('updated', 0)} обновлено")
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

            for item in items:
                await self._process_api_item(item, serial_number, records_buffer,
                                            row_index, id_pu, original_data)

            await asyncio.sleep(Config.REQUEST_DELAY)

        if records_buffer:
            save_result = self.db.save_records_batch(records_buffer)
            self.stats['records_saved'] += save_result['saved']
            self.stats['updated'] += save_result.get('updated', 0)
            logger.info(f"  💾 Сохранено: {save_result['saved']} новых, {save_result.get('updated', 0)} обновлено")
        elif found_count == 0:
            self.stats['no_match'] += 1

        self.stats['records_found'] += found_count
        self.stats['queries_processed'] += 1

        return found_count

    async def _process_api_item(self, item: Dict, serial_number: str,
                                 records_buffer: List, row_index: int,
                                 id_pu: str, original_data: dict) -> bool:
        mit_title = item.get('mit_title', '')
        mit_notation = item.get('mit_notation', '')
        mi_modification = item.get('mi_modification', '')
        record_serial = item.get('mi_number', '')

        if not is_electric_meter_from_queries(mit_title, mit_notation, mi_modification):
            self.stats['filtered_not_electric'] += 1
            logger.debug(f"  ❌ Исключено (не электросчетчик): {mit_title}")
            return False

        if not match_serial_number(record_serial, serial_number):
            self.stats['filtered_serial_mismatch'] += 1
            logger.debug(f"  ❌ Несовпадение серийного: {record_serial} != {serial_number}")
            return False

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
                                      start_from: int = 0,
                                      fallback_year: int = None,
                                      auto_range: bool = False) -> Dict:
        total_queries = len(queries)

        if verification_date:
            logger.info(f"🔍 Обработка {total_queries} запросов на дату {verification_date}"
                       + (f" (с индекса {start_from})" if start_from > 0 else ""))
        else:
            if auto_range:
                opis = "АВТО (год−1…год+2 от «Дата поверки. Год» / «Год выпуска» / мода «Дата госповерки»)"
            elif years:
                opis = f"за {len(years)} лет ({min(years)}–{max(years)})"
            else:
                opis = "без диапазона"
            logger.info(f"🔍 Обработка {total_queries} запросов, {opis}"
                       + (f" (с индекса {start_from})" if start_from > 0 else ""))

        if TQDM_AVAILABLE:
            self.pbar = tqdm(total=total_queries, initial=start_from,
                           desc="Запросы", unit="запрос")

        for i, query in enumerate(queries):
            if i < start_from:
                continue

            serial_number = query.get('mi_number', '')

            if not serial_number or serial_number.lower() == 'nan':
                logger.debug(f"⏭  Пропущено (нет серийного номера), строка {query.get('row_index', i+1)}")
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

            query_years = self._adjust_years_for_query(
                query, years, fallback_year=fallback_year, auto_range=auto_range
            )

            if query_years and years and query_years != years:
                mfg_year = query.get('manufacture_year', '')
                logger.info(f"  📅 Строка {row_index}: год выпуска {mfg_year}, "
                           f"поиск {min(query_years)}–{max(query_years)} "
                           f"(из диапазона {min(years)}–{max(years)})")

            found = await self.process_query(
                serial_number, query_years,
                row_index=row_index,
                id_pu=id_pu,
                original_data=original_data,
                verification_date=verification_date,
                use_attribute_search=use_attribute_search
            )

            serial_short = serial_number[:20] + '...' if len(serial_number) > 20 else serial_number
            if verification_date:
                year_range = verification_date
            elif query_years:
                year_range = f"{min(query_years)}–{max(query_years)}"
            else:
                year_range = "без диапазона"
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
                postfix = {
                    "найдено": self.stats['records_found'],
                    "сохранено": self.stats['records_saved']
                }
                try:
                    remaining_s = self.pbar.format_dict.get('remaining_s')
                    if remaining_s and remaining_s > 0:
                        remaining_h = remaining_s / 3600
                        if remaining_h >= 24:
                            postfix["осталось"] = f"{remaining_h:.1f}ч ({remaining_h/24:.1f} дн.)"
                        else:
                            postfix["осталось"] = f"{remaining_h:.1f}ч"
                except Exception:
                    pass
                self.pbar.set_postfix(postfix)

        if self.pbar:
            self.pbar.close()

        # Очистка прогресса после успешного завершения
        if self.input_file:
            self.db.clear_progress(self.input_file)

        return self.stats

    @staticmethod
    def _build_progress_params(years: List[int], verification_date: str,
                                use_attribute_search: bool) -> str:
        if verification_date:
            return f"date={verification_date},attr={use_attribute_search}"
        if years:
            return f"years={min(years)}-{max(years)},attr={use_attribute_search}"
        return f"years=auto,attr={use_attribute_search}"

    def get_stats(self) -> Dict:
        db_stats = self.db.get_stats()
        return {
            **self.stats,
            'db_total': db_stats['total'],
            'db_applicable': db_stats['applicable']
        }
