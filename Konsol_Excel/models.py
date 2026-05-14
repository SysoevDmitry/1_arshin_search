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
    row_index: int = 0
    id_pu: str = ""
    
    # Данные из исходного Excel файла (сохраняются для экспорта)
    contract_number: str = ""
    edo_code: str = ""
    balance_owner: str = ""
    operation_responsibility: str = ""
    mpi: str = ""

    @staticmethod
    def generate_url(vri_id: str) -> str:
        return f"https://fgis.gost.ru/fundmetrology/cm/erts/?id={vri_id}"

    @property
    def record_url(self) -> str:
        return self.generate_url(self.vri_id)

    @staticmethod
    def is_electric_meter(title: str, notation: str = "", modification: str = "") -> bool:
        combined = f"{title} {notation} {modification}".lower()
        
        if not combined:
            return False
        
        exclude_keywords = [
            'воды', 'водомер', 'водосчетчик', 'холодной воды', 'горячей воды',
            'счетчик воды', 'крыльчатые',
            'газа', 'газ', 'газосчетчик', 'газовые', 'бытовые газовые',
            'счетчик газа',
            'тепл', 'теплопотребление', 'тепловычислитель',
            'термометры медицинские', 'термометр медицинский',
            'весы', 'взвешивания',
            'манометры', 'вакуумметры', 'мановакуумметры',
            'термометры ртутные', 'термометры стеклянные',
            'штангенглубиномеры', 'нутромеры', 'ареометры', 'гигрометры',
            'мегаомметры', 'расхода электромагнитные', 'сигнализаторы',
        ]
        
        for keyword in exclude_keywords:
            if keyword in combined:
                return False
        
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
        combined = f"{title} {notation} {modification}".lower()
        
        if not combined or combined.strip() == 'нет данных':
            return 'Другие'
        
        for keyword, manufacturer in MANUFACTURERS_RULES.items():
            if keyword in combined:
                return manufacturer
        
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
            
            manufacturer=cls.detect_manufacturer(
                data.get('mit_title', ''),
                data.get('mit_notation', ''),
                data.get('mi_modification', '')
            ),
            collected_at=datetime.now().isoformat(),
            
            search_query=search_query,
            row_index=row_index,
            id_pu=id_pu,
            
            contract_number=original.get('Номер договора', original.get('contract_number', '')),
            edo_code=original.get('Код ЭДО', original.get('edo_code', '')),
            balance_owner=original.get('Балансовая принадлежность', original.get('balance_owner', '')),
            operation_responsibility=original.get('Эксплуатационная ответственность', original.get('operation_responsibility', '')),
            mpi=original.get('МПИ', original.get('mpi', ''))
        )
        
        return record

    def to_dict(self) -> dict:
        result = asdict(self)
        result['record_url'] = self.record_url
        return result

    def get_unique_key(self) -> tuple:
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
