#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модели данных для Excel версии
С полной проверкой на уникальность как в arshin_app.py
"""

from dataclasses import dataclass, asdict
from datetime import datetime
from typing import Dict, Optional
from config import MANUFACTURERS_RULES


@dataclass
class VerificationRecord:
    """
    Запись о поверке СИ с ПОЛНОЙ служебной информацией
    Адаптировано из arshin_app.py
    """
    # Основные данные из API
    vri_id: str = ""
    mit_number: str = ""
    mit_title: str = ""
    mit_notation: str = ""
    mi_modification: str = ""
    mi_number: str = ""
    verification_date: str = ""
    valid_date: str = ""
    applicability: bool = True
    org_title: str = ""
    result_docnum: str = ""
    
    # Служебные поля (добавляются приложением)
    manufacturer: str = ""
    collected_at: str = ""
    
    # Связи (для пакетного поиска из Excel)
    search_query: str = ""
    row_index: int = 0  # Индекс строки в Excel (1-based)
    id_pu: str = ""  # Id_ПУ из исходного файла
    
    # Данные из исходного Excel файла (сохраняются для экспорта)
    contract_number: str = ""  # Номер договора
    edo_code: str = ""  # Код ЭДО
    balance_owner: str = ""  # Балансовая принадлежность
    operation_responsibility: str = ""  # Эксплуатационная ответственность
    mpi: str = ""  # МПИ (межповерочный интервал)

    @staticmethod
    def generate_url(vri_id: str) -> str:
        """Генерация URL записи в ФГИС АРШИН"""
        return f"https://fgis.gost.ru/fundmetrology/cm/erts/?id={vri_id}"

    @property
    def record_url(self) -> str:
        """URL записи"""
        return self.generate_url(self.vri_id)

    @staticmethod
    def is_electric_meter(title: str, notation: str = "", modification: str = "") -> bool:
        """
        Проверка: электросчетчик
        Возвращает True только для счетчиков электроэнергии
        """
        combined = f"{title} {notation} {modification}".lower()
        
        if not combined:
            return False
        
        # Исключения (не электрические приборы)
        exclude_keywords = [
            # Вода
            'воды', 'водомер', 'водосчетчик', 'холодной воды', 'горячей воды',
            'счетчик воды', 'крыльчатые',
            # Газ
            'газа', 'газ', 'газосчетчик', 'газовые', 'бытовые газовые',
            'счетчик газа',
            # Тепло
            'тепл', 'теплопотребление', 'тепловычислитель',
            # Медицинские приборы
            'термометры медицинские', 'термометр медицинский',
            # Весы
            'весы', 'взвешивания',
            # Давление
            'манометры', 'вакуумметры', 'мановакуумметры',
            # Температура
            'термометры ртутные', 'термометры стеклянные',
            # Прочие
            'штангенглубиномеры', 'нутромеры', 'ареометры', 'гигрометры',
            'мегаомметры', 'расхода электромагнитные', 'сигнализаторы',
        ]
        
        for keyword in exclude_keywords:
            if keyword in combined:
                return False
        
        # Включения (электрическая энергия)
        include_keywords = [
            'электрической энергии', 'электроэнергии', 'электросчетчик',
            'счетчик электрической', 'счетчик электроэнергии',
            'активной электроэнергии', 'реактивной электроэнергии',
            'меркурий', 'нева', 'энергомера', 'матрица', 'альфа', 'милур',
            'псч', 'сэт', 'цэ', 'се 101', 'се 102', 'се 201',
        ]
        
        for keyword in include_keywords:
            if keyword in combined:
                return True
        
        return False

    @staticmethod
    def detect_manufacturer(title: str, notation: str = "", modification: str = "") -> str:
        """Определение производителя по названию"""
        combined = f"{title} {notation} {modification}".lower()
        
        if not combined or combined.strip() == 'нет данных':
            return 'Другие'
        
        # Проверка по словарю
        for keyword, manufacturer in MANUFACTURERS_RULES.items():
            if keyword in combined:
                return manufacturer
        
        # Резервные правила
        fallback = {
            'меркурий': 'Меркурий',
            'нева': 'Нева',
            'энергомера': 'Энергомера',
        }
        for key, value in fallback.items():
            if key in combined:
                return value
        
        return 'Другие'

    @classmethod
    def from_api(cls, data: dict, search_query: str = "",
                 row_index: int = 0, id_pu: str = "",
                 original_data: dict = None) -> 'VerificationRecord':
        """
        Создание записи из ответа API с сохранением связей
        Адаптировано из arshin_app.py
        """
        original = original_data or {}
        
        record = cls(
            vri_id=data.get('vri_id', ''),
            mit_number=data.get('mit_number', ''),
            mit_title=data.get('mit_title', ''),
            mit_notation=data.get('mit_notation', ''),
            mi_modification=data.get('mi_modification', ''),
            mi_number=data.get('mi_number', ''),
            verification_date=data.get('verification_date', ''),
            valid_date=data.get('valid_date', ''),
            applicability=data.get('applicability', True),
            org_title=data.get('org_title', ''),
            result_docnum=data.get('result_docnum', ''),
            
            # Служебные поля
            manufacturer=cls.detect_manufacturer(
                data.get('mit_title', ''),
                data.get('mit_notation', ''),
                data.get('mi_modification', '')
            ),
            collected_at=datetime.now().isoformat(),
            
            # Связи
            search_query=search_query,
            row_index=row_index,
            id_pu=id_pu,
            
            # Данные из Excel файла
            contract_number=original.get('Номер договора', original.get('contract_number', '')),
            edo_code=original.get('Код ЭДО', original.get('edo_code', '')),
            balance_owner=original.get('Балансовая принадлежность', original.get('balance_owner', '')),
            operation_responsibility=original.get('Эксплуатационная ответственность', original.get('operation_responsibility', '')),
            mpi=original.get('МПИ', original.get('mpi', ''))
        )
        
        return record

    def to_dict(self) -> dict:
        """Преобразование в словарь с полной служебной информацией"""
        result = asdict(self)
        result['record_url'] = self.record_url
        return result

    def get_unique_key(self) -> tuple:
        """
        Ключ для проверки на ПОЛНУЮ уникальность
        Адаптировано из arshin_app.py - проверка по всем основным + служебным полям
        """
        return (
            self.vri_id,
            self.mit_number,
            self.mit_title,
            self.mi_number,
            self.verification_date,
            self.valid_date,
            self.applicability,
            self.org_title,
            self.result_docnum,
            self.manufacturer,
            self.search_query,
            self.row_index,
            self.id_pu
        )
