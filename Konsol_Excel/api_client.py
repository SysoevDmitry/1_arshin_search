#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Асинхронный API клиент с ПАРАЛЛЕЛЬНЫМИ запросами
Для Excel версии

Доработки:
- Обработка 408 Request Timeout
- Получение записи по vri_id
- Параметр verification_date (взаимозаменяемость с year)
- Атрибутивный поиск (фильтры по полям)
- Детализированное логирование 5XX ошибок
"""

import asyncio
import aiohttp
import logging
from typing import List, Dict, Optional
from urllib.parse import quote
from datetime import datetime
from config import Config

logger = logging.getLogger(__name__)

_CONNECTION_ERRORS = (
    'Cannot connect to host', 'Temporary failure in name resolution',
    'Connection', 'connect', 'Name or service not known',
    'No address', 'getaddrinfo'
)
_RETRYABLE_ERRORS = {'429', '502', 'timeout', '408'}
_CONN_RETRY_BASE = 3  # базовая пауза для сетевых ошибок (сек)


class ParallelAPIClient:
    """
    Клиент с параллельными запросами через asyncio.Semaphore

    Производительность:
    - MAX_CONCURRENT_REQUESTS = 5 → ~5 запросов/сек
    - Обработка 50 страниц за ~10-15 секунд
    """

    def __init__(self, max_concurrent: int = None):
        self.max_concurrent = max_concurrent or Config.MAX_CONCURRENT_REQUESTS
        self.base_url = Config.API_VRI
        self.semaphore = None
        self.session = None
        self.stats = {'requests': 0, 'errors': 0, 'found': 0}
        self.error_log = []
        self._connection_error_count = 0
        self._last_conn_error_time = None
        self._conn_cooldown_until = None
    
    async def __aenter__(self):
        self.semaphore = asyncio.Semaphore(self.max_concurrent)
        self.session = aiohttp.ClientSession(
            headers={
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                'Accept': 'application/json',
                'Accept-Language': 'ru-RU,ru;q=0.9',
            },
            timeout=aiohttp.ClientTimeout(total=Config.REQUEST_TIMEOUT)
        )
        logger.info(f"🚀 APIClient запущен (параллельно: {self.max_concurrent})")
        return self
    
    async def __aexit__(self, *args):
        if self.session:
            await self.session.close()
        logger.debug("APIClient остановлен")
    
    async def search_vri(self, search_term: str = None, year: int = None,
                         start: int = 0, rows: int = None,
                         verification_date: str = None,
                         filters: Dict[str, str] = None) -> Dict:
        if rows is None:
            rows = Config.MAX_ROWS_PER_REQUEST

        rows = min(rows, Config.MAX_ROWS_PER_REQUEST)
        rows = max(1, rows)
        start = max(0, start)

        params = {}

        if search_term:
            params['search'] = quote(search_term, safe='*')

        if verification_date:
            params['verification_date'] = verification_date
        elif year:
            params['year'] = year

        params['start'] = start
        params['rows'] = rows

        if filters:
            for key, value in filters.items():
                params[key] = quote(value, safe='*?')

        async with self.semaphore:
            self.stats['requests'] += 1

            try:
                async with self.session.get(self.base_url, params=params) as resp:
                    if resp.status == 200:
                        data = await resp.json()
                        items = data.get('result', {}).get('items', [])
                        total = data.get('result', {}).get('count', 0)
                        self.stats['found'] += len(items)
                        self._connection_error_count = 0
                        self._conn_cooldown_until = None
                        return {'items': items, 'count': total, 'error': None}

                    elif resp.status == 429:
                        logger.warning(f"⚠️  API 429: пауза 3с")
                        await asyncio.sleep(3)
                        return {'items': [], 'count': 0, 'error': '429'}

                    elif resp.status == 502:
                        logger.warning(f"⚠️  API 502: пауза 2с")
                        await asyncio.sleep(2)
                        return {'items': [], 'count': 0, 'error': '502'}

                    elif resp.status == 408:
                        logger.warning(f"⚠️  API 408: превышен лимит времени поиска")
                        return {'items': [], 'count': 0, 'error': '408'}

                    elif resp.status == 409:
                        logger.warning(f"⚠️  API 409: лимит страниц")
                        return {'items': [], 'count': 0, 'error': '409'}

                    elif resp.status == 400:
                        logger.error(f"❌ API 400: некорректный запрос")
                        return {'items': [], 'count': 0, 'error': '400'}

                    elif resp.status >= 500:
                        error_time = datetime.now().isoformat()
                        error_info = {
                            'status': resp.status,
                            'time': error_time,
                            'url': str(resp.url),
                            'params': params
                        }
                        self.error_log.append(error_info)
                        logger.error(f"❌ API {resp.status} (5XX): ошибка сервера")
                        logger.error(f"   Время: {error_time}")
                        logger.error(f"   URL: {resp.url}")
                        logger.error(f"   Параметры: {params}")
                        logger.error(f"   Для отправки в поддержку: fgis2@rst.gov.ru")
                        return {'items': [], 'count': 0, 'error': f'{resp.status}'}

                    else:
                        self.stats['errors'] += 1
                        logger.debug(f"API статус: {resp.status}")
                        return {'items': [], 'count': 0, 'error': str(resp.status)}

            except asyncio.TimeoutError:
                self.stats['errors'] += 1
                logger.error(f"❌ Таймаут запроса")
                return {'items': [], 'count': 0, 'error': 'timeout'}

            except Exception as e:
                self.stats['errors'] += 1
                err_msg = str(e)
                is_conn_error = any(phrase in err_msg for phrase in _CONNECTION_ERRORS)
                
                if is_conn_error:
                    self._connection_error_count += 1
                    now = datetime.now()
                    if self._last_conn_error_time is None or (now - self._last_conn_error_time).total_seconds() > 30:
                        logger.error(f"❌ Ошибка соединения: {err_msg}")
                    self._last_conn_error_time = now
                    if self._connection_error_count % 20 == 0:
                        logger.error(f"❌ Накоплено {self._connection_error_count} ошибок соединения, пауза 10с...")
                        self._conn_cooldown_until = now
                    return {'items': [], 'count': 0, 'error': 'connection', 'error_msg': err_msg}
                else:
                    logger.error(f"❌ Ошибка: {err_msg}")
                    return {'items': [], 'count': 0, 'error': str(e)}
    
    async def search_with_retry(self, search_term: str = None, year: int = None,
                                 start: int = 0, rows: int = None,
                                 verification_date: str = None,
                                 filters: Dict[str, str] = None,
                                 max_retries: int = 3) -> Dict:
        for attempt in range(max_retries):
            if self._conn_cooldown_until:
                wait = (self._conn_cooldown_until - datetime.now()).total_seconds()
                if wait > 0:
                    await asyncio.sleep(min(wait, 10))
                    self._conn_cooldown_until = None

            result = await self.search_vri(
                search_term=search_term,
                year=year,
                start=start,
                rows=rows,
                verification_date=verification_date,
                filters=filters
            )

            if result.get('error') is None:
                return result

            error = result.get('error', '')
            if error in _RETRYABLE_ERRORS or error == 'connection':
                base_delay = _CONN_RETRY_BASE if error == 'connection' else Config.REQUEST_DELAY
                delay = base_delay * (2 ** attempt)
                logger.debug(f"Повторная попытка {attempt+1}/{max_retries} через {delay}с")
                await asyncio.sleep(delay)
            else:
                break

        return result

    async def get_vri_record(self, vri_id: str) -> Dict:
        url = f"{self.base_url}/{vri_id}"

        try:
            async with self.session.get(url) as resp:
                if resp.status == 200:
                    data = await resp.json()
                    result = data.get('result', {})
                    pub_status = result.get('publication', {}).get('status', '')

                    if pub_status == "Запись была модифицирована":
                        logger.warning(f"⚠️  Запись {vri_id} неактуальна (publication.status)")
                        return {
                            'item': result,
                            'error': 'not_actual',
                            'message': 'Запись была модифицирована - выполните повторный поиск'
                        }

                    logger.info(f"✅ Получена запись: {vri_id}")
                    return {'item': result, 'error': None}

                elif resp.status == 404:
                    logger.warning(f"⚠️  Запись не найдена: {vri_id}")
                    return {'item': None, 'error': '404', 'message': 'Запись не найдена'}

                elif resp.status == 429:
                    logger.warning(f"⚠️  API 429: пауза 3с")
                    await asyncio.sleep(3)
                    return {'item': None, 'error': '429'}

                elif resp.status >= 500:
                    error_time = datetime.now().isoformat()
                    error_info = {
                        'status': resp.status,
                        'time': error_time,
                        'url': str(url),
                        'vri_id': vri_id
                    }
                    self.error_log.append(error_info)
                    logger.error(f"❌ API {resp.status} (5XX) при получении записи: {vri_id}")
                    logger.error(f"   Время: {error_time}")
                    logger.error(f"   Для отправки в поддержку: fgis2@rst.gov.ru")
                    return {'item': None, 'error': f'{resp.status}'}

                else:
                    logger.error(f"❌ API статус: {resp.status}")
                    return {'item': None, 'error': str(resp.status)}

        except asyncio.TimeoutError:
            logger.error(f"❌ Таймаут запроса записи: {vri_id}")
            return {'item': None, 'error': 'timeout'}

        except Exception as e:
            err_msg = str(e)
            is_conn_error = any(phrase in err_msg for phrase in _CONNECTION_ERRORS)
            if is_conn_error:
                self._connection_error_count += 1
                now = datetime.now()
                if self._last_conn_error_time is None or (now - self._last_conn_error_time).total_seconds() > 30:
                    logger.error(f"❌ Ошибка соединения: {err_msg}")
                self._last_conn_error_time = now
                return {'item': None, 'error': 'connection', 'error_msg': err_msg}
            else:
                logger.error(f"❌ Ошибка получения записи {vri_id}: {e}")
                return {'item': None, 'error': str(e)}

    async def get_vri_record_with_retry(self, vri_id: str, max_retries: int = 3) -> Dict:
        for attempt in range(max_retries):
            if self._conn_cooldown_until:
                wait = (self._conn_cooldown_until - datetime.now()).total_seconds()
                if wait > 0:
                    await asyncio.sleep(min(wait, 10))
                    self._conn_cooldown_until = None

            result = await self.get_vri_record(vri_id)

            if result.get('error') is None:
                return result

            error = result.get('error', '')
            if error in _RETRYABLE_ERRORS or error == 'connection' or error in ['502', '503', '504']:
                delay = Config.REQUEST_DELAY * (2 ** attempt)
                logger.debug(f"Повторная попытка {attempt+1}/{max_retries} для {vri_id}")
                await asyncio.sleep(delay)
            else:
                break

        return result
    
    async def fetch_parallel(self, search_term: str = None, year: int = None,
                              pages: List[int] = None,
                              verification_date: str = None,
                              filters: Dict[str, str] = None) -> List[Dict]:
        if pages is None:
            pages = [0]

        tasks = [
            self.search_with_retry(
                search_term=search_term,
                year=year,
                start=page * Config.MAX_ROWS_PER_REQUEST,
                verification_date=verification_date,
                filters=filters
            )
            for page in pages
        ]
        results = await asyncio.gather(*tasks, return_exceptions=True)
        return results

    def get_error_log(self) -> List[Dict]:
        return self.error_log.copy()
