# -*- coding: utf-8 -*-
"""
API-клиент для Android-приложения
Синхронный (на основе requests) — надёжно работает на Android/Kivy
"""

import logging
import time
import requests
from typing import Dict, List, Optional
from urllib.parse import quote
from config import Config

logger = logging.getLogger(__name__)


class APIClient:
    """
    Синхронный клиент API ФГИС АРШИН

    Использует библиотеку requests вместо aiohttp —
    проще и надёжнее на Android.
    """

    BASE_URL = Config.API_VRI
    TIMEOUT = Config.REQUEST_TIMEOUT
    MAX_ROWS = Config.MAX_ROWS_PER_REQUEST

    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Linux; Android 13; Pixel 7) AppleWebKit/537.36',
            'Accept': 'application/json',
            'Accept-Language': 'ru-RU,ru;q=0.9',
        })
        self.stats = {'requests': 0, 'errors': 0, 'found': 0}

    def _request(self, params: dict) -> Dict:
        """Выполнение одного запроса с возвратом структурированного ответа"""
        self.stats['requests'] += 1

        try:
            resp = self.session.get(self.BASE_URL, params=params, timeout=self.TIMEOUT)

            if resp.status_code == 200:
                data = resp.json()
                items = data.get('result', {}).get('items', [])
                total = data.get('result', {}).get('count', 0)
                self.stats['found'] += len(items)
                return {'items': items, 'count': total, 'error': None}

            elif resp.status_code == 429:
                logger.warning("API 429: пауза 3с")
                time.sleep(3)
                return {'items': [], 'count': 0, 'error': '429'}

            elif resp.status_code == 502:
                logger.warning("API 502: пауза 2с")
                time.sleep(2)
                return {'items': [], 'count': 0, 'error': '502'}

            elif resp.status_code == 408:
                logger.warning("API 408: таймаут на стороне сервера")
                return {'items': [], 'count': 0, 'error': '408'}

            elif resp.status_code == 409:
                logger.warning("API 409: лимит страниц")
                return {'items': [], 'count': 0, 'error': '409'}

            elif resp.status_code >= 500:
                logger.error(f"API {resp.status_code}: ошибка сервера")
                return {'items': [], 'count': 0, 'error': str(resp.status_code)}

            else:
                self.stats['errors'] += 1
                logger.debug(f"API статус: {resp.status_code}")
                return {'items': [], 'count': 0, 'error': str(resp.status_code)}

        except requests.exceptions.Timeout:
            self.stats['errors'] += 1
            logger.error("Таймаут запроса")
            return {'items': [], 'count': 0, 'error': 'timeout'}

        except requests.exceptions.ConnectionError:
            self.stats['errors'] += 1
            logger.error("Ошибка соединения. Проверьте интернет.")
            return {'items': [], 'count': 0, 'error': 'connection'}

        except Exception as e:
            self.stats['errors'] += 1
            logger.error(f"Ошибка: {e}")
            return {'items': [], 'count': 0, 'error': str(e)}

    def search_vri(self,
                   search_term: str = None,
                   year: int = None,
                   start: int = 0,
                   rows: int = None,
                   verification_date: str = None,
                   filters: Dict[str, str] = None) -> Dict:
        """
        Поиск в реестре VRI

        Args:
            search_term: Поисковая строка
            year: Год поверки
            start: Начальная позиция
            rows: Количество записей (макс. 100)
            verification_date: Конкретная дата поверки (yyyy-MM-dd)
            filters: Атрибутивные фильтры {'mi_number': '...'}
        """
        if rows is None:
            rows = self.MAX_ROWS
        rows = min(rows, self.MAX_ROWS)
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

        return self._request(params)

    def search_with_retry(self,
                          search_term: str = None,
                          year: int = None,
                          verification_date: str = None,
                          filters: Dict[str, str] = None,
                          max_retries: int = 3) -> Dict:
        """Поиск с повторными попытками при ошибках"""
        for attempt in range(max_retries):
            result = self.search_vri(
                search_term=search_term,
                year=year,
                verification_date=verification_date,
                filters=filters
            )

            if result.get('error') is None:
                return result

            if result.get('error') in ['429', '502', 'timeout', '408']:
                delay = Config.REQUEST_DELAY * (2 ** attempt)
                logger.debug(f"Повтор {attempt+1}/{max_retries} через {delay}с")
                time.sleep(delay)
            else:
                break

        return result

    def get_vri_record(self, vri_id: str) -> Dict:
        """Получение отдельной записи по vri_id"""
        url = f"{self.BASE_URL}/{vri_id}"

        try:
            resp = self.session.get(url, timeout=self.TIMEOUT)

            if resp.status_code == 200:
                data = resp.json()
                result = data.get('result', {})
                pub_status = result.get('publication', {}).get('status', '')

                if pub_status == "Запись была модифицирована":
                    return {
                        'item': result,
                        'error': 'not_actual',
                        'message': 'Запись была модифицирована'
                    }

                return {'item': result, 'error': None}

            elif resp.status_code == 404:
                return {'item': None, 'error': '404', 'message': 'Запись не найдена'}

            elif resp.status_code >= 500:
                return {'item': None, 'error': str(resp.status_code)}

            else:
                return {'item': None, 'error': str(resp.status_code)}

        except Exception as e:
            return {'item': None, 'error': str(e)}

    def close(self):
        """Закрытие сессии"""
        if self.session:
            self.session.close()
