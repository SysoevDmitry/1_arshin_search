# -*- coding: utf-8 -*-
"""
ФГИС АРШИН — Android приложение
Поиск поверок электросчётчиков через API ФГИС АРШИН
(полностью программный UI, без KV)
"""

import os
import re
import threading
import logging
from datetime import datetime
from typing import List, Dict

from kivy.app import App
from kivy.clock import mainthread
from kivy.core.window import Window
from kivy.metrics import dp
from kivy.properties import (
    StringProperty, BooleanProperty, ListProperty
)
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.checkbox import CheckBox
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.recycleview import RecycleView
from kivy.uix.recycleboxlayout import RecycleBoxLayout
from kivy.uix.scrollview import ScrollView
from kivy.uix.spinner import Spinner
from kivy.uix.tabbedpanel import TabbedPanel, TabbedPanelItem
from kivy.uix.textinput import TextInput
from kivy.utils import platform

from config import Config, EXACT_QUERIES, MANUFACTURERS_RULES
from models import VerificationRecord
from database import Database
from api_client import APIClient
from excel_handler import ExcelHandler

# Логирование
LOG_DIR = "Logs"
os.makedirs(LOG_DIR, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
)
logger = logging.getLogger('arshin_android')

# ---------------------------------------------------------------------------
# Фильтрация
# ---------------------------------------------------------------------------

def normalize_serial_number(serial: str) -> str:
    if not serial:
        return ""
    return re.sub(r'[\s\-]', '', serial.upper())


def match_serial_number(record_serial: str, target_serial: str) -> bool:
    if not target_serial or not record_serial:
        return False
    norm_record = normalize_serial_number(record_serial)
    norm_target = normalize_serial_number(target_serial)
    if norm_record == norm_target:
        return True
    return norm_record.lstrip('0') == norm_target.lstrip('0')


def load_electric_filters() -> tuple:
    include_phrases = []
    for phrase in EXACT_QUERIES:
        include_phrases.append(phrase.lower())

    electric_markers = [
        'электрической энергии', 'электроэнергии', 'электросчетчик',
        'счетчик электрической', 'счетчик электроэнергии',
        'активной электроэнергии', 'реактивной электроэнергии',
        'однофазный', 'трехфазный', 'статический', 'индукционный',
        'меркурий', 'нева', 'энергомера', 'матрица', 'альфа', 'милур',
        'псч', 'сэт', 'цэ', 'се 101', 'се 102', 'се 201', 'се 301',
        'меркурий 20', 'меркурий 23', 'нева 10', 'нева 30',
        'трансформатор тока измерительный',
    ]
    include_phrases.extend(electric_markers)

    exclude_phrases = [
        'воды', 'водомер', 'водосчетчик', 'холодной воды', 'горячей воды',
        'счетчик воды', 'крыльчатые',
        'газа', 'газ', 'газосчетчик', 'газовые', 'бытовые газовые',
        'диафрагменные', 'счетчик газа',
        'тепл', 'теплопотребление', 'тепловычислитель',
        'термометры медицинские', 'весы', 'взвешивания',
        'манометры', 'вакуумметры', 'сигнализаторы',
    ]
    return include_phrases, exclude_phrases


ELECTRIC_INCLUDE, ELECTRIC_EXCLUDE = load_electric_filters()


def is_electric_meter_from_queries(
    title: str, notation: str = "", modification: str = ""
) -> bool:
    combined = f"{title} {notation} {modification}".lower()
    if not combined:
        return False
    for keyword in ELECTRIC_EXCLUDE:
        if keyword in combined:
            return False
    for phrase in ELECTRIC_INCLUDE:
        if phrase in combined:
            return True
    return False


# ---------------------------------------------------------------------------
# Карточка результата
# ---------------------------------------------------------------------------

class ResultCard(BoxLayout):
    title_text = StringProperty('')
    date_text = StringProperty('')
    serial_text = StringProperty('')
    manufacturer_text = StringProperty('')
    status_text = StringProperty('')
    applicative = BooleanProperty(True)
    verifier_text = StringProperty('')

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.size_hint_y = None
        self.height = dp(105)
        self.padding = dp(6)
        self.spacing = dp(1)

        row0 = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(22))
        self._title_label = Label(
            font_size='13sp', bold=True, color=(0.1, 0.5, 0.1, 1),
            size_hint_x=0.75, halign='left', shorten=True
        )
        self.bind(title_text=self._title_label.setter('text'))
        self._date_label = Label(
            font_size='12sp', color=(0.3, 0.3, 0.3, 1),
            size_hint_x=0.25, halign='right'
        )
        self.bind(date_text=self._date_label.setter('text'))
        row0.add_widget(self._title_label)
        row0.add_widget(self._date_label)
        self.add_widget(row0)

        self._serial_label = Label(
            font_size='12sp', color=(0.2, 0.2, 0.2, 1),
            size_hint_y=None, height=dp(20), halign='left'
        )
        self.bind(serial_text=self._serial_label.setter('text'))
        self.add_widget(self._serial_label)

        row2 = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(20))
        self._mfg_label = Label(
            font_size='11sp', color=(0.4, 0.4, 0.4, 1),
            size_hint_x=0.6, halign='left', shorten=True
        )
        self.bind(manufacturer_text=self._mfg_label.setter('text'))
        self._status_label = Label(
            font_size='11sp', bold=True,
            size_hint_x=0.4, halign='right'
        )
        self.bind(status_text=self._status_label.setter('text'))
        self.bind(applicative=self._update_status_color)
        row2.add_widget(self._mfg_label)
        row2.add_widget(self._status_label)
        self.add_widget(row2)

        self._verifier_label = Label(
            font_size='10sp', color=(0.5, 0.5, 0.5, 1),
            size_hint_y=None, height=dp(18), halign='left', shorten=True
        )
        self.bind(verifier_text=self._verifier_label.setter('text'))
        self.add_widget(self._verifier_label)

    def _update_status_color(self, *args):
        if self.applicable:
            self._status_label.color = (0.2, 0.7, 0.2, 1)
        else:
            self._status_label.color = (0.8, 0.2, 0.2, 1)


# ---------------------------------------------------------------------------
# Вкладка «Поиск»
# ---------------------------------------------------------------------------

class SearchTab(TabbedPanelItem):
    def __init__(self, app, **kwargs):
        super().__init__(text='Поиск', **kwargs)
        self.app = app

        scroll = ScrollView(do_scroll_x=False)
        root = BoxLayout(
            orientation='vertical',
            size_hint_y=None,
            padding=dp(10),
            spacing=dp(6),
        )
        root.bind(minimum_height=root.setter('height'))

        # Заголовок
        root.add_widget(Label(
            text='Поиск поверок электросчётчиков',
            font_size='16sp', bold=True,
            size_hint_y=None, height=dp(30), halign='center',
        ))

        # Серийный номер
        root.add_widget(Label(
            text='Серийный (заводской) номер:',
            font_size='13sp', size_hint_y=None, height=dp(22),
        ))
        self.serial_input = TextInput(
            hint_text='Введите серийный номер...',
            multiline=False, font_size='16sp',
            size_hint_y=None, height=dp(44),
        )
        root.add_widget(self.serial_input)

        # Годы
        year_row = BoxLayout(
            orientation='horizontal',
            size_hint_y=None, height=dp(50), spacing=dp(8),
        )
        yf_box = BoxLayout(orientation='vertical', size_hint_x=0.5)
        yf_box.add_widget(Label(
            text='Год от:', font_size='12sp', size_hint_y=None, height=dp(18),
        ))
        self.year_from = Spinner(
            text='2020', font_size='14sp',
            values=['2018','2019','2020','2021','2022','2023','2024','2025','2026'],
        )
        yf_box.add_widget(self.year_from)
        year_row.add_widget(yf_box)

        yt_box = BoxLayout(orientation='vertical', size_hint_x=0.5)
        yt_box.add_widget(Label(
            text='Год до:', font_size='12sp', size_hint_y=None, height=dp(18),
        ))
        self.year_to = Spinner(
            text='2026', font_size='14sp',
            values=['2018','2019','2020','2021','2022','2023','2024','2025','2026'],
        )
        yt_box.add_widget(self.year_to)
        year_row.add_widget(yt_box)
        root.add_widget(year_row)

        # Атрибутивный поиск
        attr_row = BoxLayout(
            orientation='horizontal',
            size_hint_y=None, height=dp(36), spacing=dp(8),
        )
        self.attr_check = CheckBox(active=True, size_hint_x=None, width=dp(36))
        attr_row.add_widget(self.attr_check)
        attr_row.add_widget(Label(
            text='Атрибутивный поиск (mi_number)', font_size='12sp',
        ))
        root.add_widget(attr_row)

        # Кнопка поиска
        search_btn = Button(
            text='НАЙТИ',
            font_size='16sp', bold=True,
            size_hint_y=None, height=dp(48),
            background_color=(0.2, 0.6, 0.2, 1),
            color=(1, 1, 1, 1),
        )
        search_btn.bind(on_press=self.do_search)
        root.add_widget(search_btn)

        # Статус
        self.status_label = Label(
            text='',
            font_size='13sp', color=(0.4, 0.4, 0.4, 1),
            size_hint_y=None, height=dp(22), halign='center',
        )
        root.add_widget(self.status_label)

        # Заголовок результатов
        root.add_widget(Label(
            text='Результаты:',
            font_size='14sp', bold=True,
            size_hint_y=None, height=dp(24),
        ))

        # Список результатов
        self.results_list = RecycleView(size_hint_y=None, height=dp(400))
        self.results_box = RecycleBoxLayout(
            orientation='vertical',
            size_hint_y=None,
            default_size=(None, dp(105)),
            default_size_hint=(1, None),
        )
        self.results_box.bind(minimum_height=self.results_box.setter('height'))
        self.results_list.add_widget(self.results_box)
        root.add_widget(self.results_list)

        scroll.add_widget(root)
        self.add_widget(scroll)

    def do_search(self, *args):
        serial = self.serial_input.text.strip()
        if not serial:
            self.status_label.text = 'Введите серийный номер!'
            return

        year_from = int(self.year_from.text)
        year_to = int(self.year_to.text)
        if year_from > year_to:
            year_from, year_to = year_to, year_from

        years = list(range(year_from, year_to + 1))
        use_attr = self.attr_check.active

        self.status_label.text = 'Поиск...'
        self.results_box.clear_widgets()

        thread = threading.Thread(
            target=self.app._search_thread,
            args=(self, serial, years, use_attr),
            daemon=True,
        )
        thread.start()


# ---------------------------------------------------------------------------
# Вкладка «Excel»
# ---------------------------------------------------------------------------

class ExcelTab(TabbedPanelItem):
    def __init__(self, app, **kwargs):
        super().__init__(text='Excel', **kwargs)
        self.app = app
        self.excel_path = ''

        scroll = ScrollView(do_scroll_x=False)
        root = BoxLayout(
            orientation='vertical',
            size_hint_y=None,
            padding=dp(10),
            spacing=dp(6),
        )
        root.bind(minimum_height=root.setter('height'))

        root.add_widget(Label(
            text='Пакетный поиск из Excel',
            font_size='16sp', bold=True,
            size_hint_y=None, height=dp(30), halign='center',
        ))

        select_btn = Button(
            text='ВЫБРАТЬ EXCEL ФАЙЛ',
            font_size='14sp',
            size_hint_y=None, height=dp(44),
            background_color=(0.3, 0.5, 0.8, 1),
            color=(1, 1, 1, 1),
        )
        select_btn.bind(on_press=self.select_file)
        root.add_widget(select_btn)

        self.file_label = Label(
            text='Файл не выбран',
            font_size='12sp', color=(0.5, 0.5, 0.5, 1),
            size_hint_y=None, height=dp(22), halign='center',
        )
        root.add_widget(self.file_label)

        # Годы
        year_row = BoxLayout(
            orientation='horizontal',
            size_hint_y=None, height=dp(50), spacing=dp(8),
        )
        yf_box = BoxLayout(orientation='vertical', size_hint_x=0.5)
        yf_box.add_widget(Label(
            text='Год от:', font_size='12sp', size_hint_y=None, height=dp(18),
        ))
        self.ey_from = Spinner(
            text='2020', font_size='14sp',
            values=['2018','2019','2020','2021','2022','2023','2024','2025','2026'],
        )
        yf_box.add_widget(self.ey_from)
        year_row.add_widget(yf_box)

        yt_box = BoxLayout(orientation='vertical', size_hint_x=0.5)
        yt_box.add_widget(Label(
            text='Год до:', font_size='12sp', size_hint_y=None, height=dp(18),
        ))
        self.ey_to = Spinner(
            text='2026', font_size='14sp',
            values=['2018','2019','2020','2021','2022','2023','2024','2025','2026'],
        )
        yt_box.add_widget(self.ey_to)
        year_row.add_widget(yt_box)
        root.add_widget(year_row)

        self.run_btn = Button(
            text='ЗАПУСТИТЬ ПОИСК',
            font_size='14sp', bold=True,
            size_hint_y=None, height=dp(44),
            background_color=(0.2, 0.6, 0.2, 1),
            color=(1, 1, 1, 1),
            disabled=True,
        )
        self.run_btn.bind(on_press=self.run_batch)
        root.add_widget(self.run_btn)

        self.ex_status = Label(
            text='',
            font_size='12sp', color=(0.4, 0.4, 0.4, 1),
            size_hint_y=None, height=dp(20), halign='center',
        )
        root.add_widget(self.ex_status)

        self.ex_results = Label(
            text='',
            font_size='12sp',
            size_hint_y=None, height=dp(60), halign='left',
        )
        root.add_widget(self.ex_results)

        scroll.add_widget(root)
        self.add_widget(scroll)

    def select_file(self, *args):
        content = BoxLayout(orientation='vertical', spacing=dp(4))
        fc = FileChooserListView(
            filters=['*.xlsx', '*.xls'],
            path='/home/dmitry',
        )
        content.add_widget(fc)

        btn_box = BoxLayout(
            orientation='horizontal',
            size_hint_y=None, height=dp(44), spacing=dp(8),
        )

        popup = Popup(title='Выберите Excel файл', content=content, size_hint=(0.95, 0.8))

        def on_select(instance):
            sel = fc.selection
            if sel:
                self.excel_path = sel[0]
                self.file_label.text = os.path.basename(self.excel_path)
                self.run_btn.disabled = False
            popup.dismiss()

        btn_box.add_widget(Button(
            text='Отмена', size_hint_y=None, height=dp(40),
            on_press=popup.dismiss,
        ))
        btn_box.add_widget(Button(
            text='Выбрать', size_hint_y=None, height=dp(40),
            on_press=on_select,
        ))
        content.add_widget(btn_box)
        popup.open()

    def run_batch(self, *args):
        if not self.excel_path:
            return

        yf = int(self.ey_from.text)
        yt = int(self.ey_to.text)
        if yf > yt:
            yf, yt = yt, yf
        years = list(range(yf, yt + 1))

        self.ex_status.text = 'Чтение Excel...'

        thread = threading.Thread(
            target=self.app._batch_thread,
            args=(self, self.excel_path, years),
            daemon=True,
        )
        thread.start()


# ---------------------------------------------------------------------------
# Вкладка «История»
# ---------------------------------------------------------------------------

class HistoryTab(TabbedPanelItem):
    def __init__(self, app, **kwargs):
        super().__init__(text='История', **kwargs)
        self.app = app

        root = BoxLayout(orientation='vertical', padding=dp(10), spacing=dp(6))

        # Фильтр
        filter_row = BoxLayout(
            orientation='horizontal',
            size_hint_y=None, height=dp(40), spacing=dp(6),
        )
        self.filter_input = TextInput(
            hint_text='Фильтр по серийному номеру...',
            multiline=False, font_size='14sp',
            size_hint_x=0.7,
        )
        filter_row.add_widget(self.filter_input)

        filter_btn = Button(
            text='Фильтр', font_size='12sp', size_hint_x=0.3,
        )
        filter_btn.bind(on_press=lambda x: self.load_history(self.filter_input.text))
        filter_row.add_widget(filter_btn)
        root.add_widget(filter_row)

        # Счётчик
        self.count_label = Label(
            text='Всего записей: 0',
            font_size='12sp', color=(0.5, 0.5, 0.5, 1),
            size_hint_y=None, height=dp(20),
        )
        root.add_widget(self.count_label)

        # Список истории
        self.history_list = RecycleView()
        self.history_box = RecycleBoxLayout(
            orientation='vertical',
            size_hint_y=None,
            default_size=(None, dp(105)),
            default_size_hint=(1, None),
        )
        self.history_box.bind(minimum_height=self.history_box.setter('height'))
        self.history_list.add_widget(self.history_box)
        root.add_widget(self.history_list)

        # Кнопки действий
        action_row = BoxLayout(
            orientation='horizontal',
            size_hint_y=None, height=dp(42), spacing=dp(6),
        )
        export_btn = Button(
            text='ЭКСПОРТ CSV', font_size='12sp', size_hint_x=0.5,
            background_color=(0.3, 0.5, 0.8, 1), color=(1, 1, 1, 1),
        )
        export_btn.bind(on_press=lambda x: self.app.export_csv())
        action_row.add_widget(export_btn)

        clear_btn = Button(
            text='ОЧИСТИТЬ', font_size='12sp', size_hint_x=0.5,
            background_color=(0.8, 0.3, 0.3, 1), color=(1, 1, 1, 1),
        )
        clear_btn.bind(on_press=lambda x: self.app.clear_history(self))
        action_row.add_widget(clear_btn)
        root.add_widget(action_row)

        self.add_widget(root)

    def load_history(self, filter_text=''):
        records = self.app.db.get_all_records(limit=500)

        if filter_text:
            ft = filter_text.lower()
            records = [
                r for r in records
                if ft in str(r.get('mi_number', '')).lower()
                or ft in str(r.get('search_query', '')).lower()
            ]

        self.history_box.clear_widgets()
        for r in records:
            card = ResultCard()
            card.title_text = r.get('mit_title') or '—'
            card.date_text = r.get('verification_date') or ''
            card.serial_text = (
                f"Зав. № {r.get('mi_number', '')}"
                if r.get('mi_number') else ''
            )
            card.manufacturer_text = r.get('manufacturer') or ''
            card.applicable = bool(r.get('applicability', True))
            card.status_text = 'Пригоден' if card.applicable else 'Непригоден'
            card.verifier_text = r.get('org_title') or ''
            self.history_box.add_widget(card)

        self.count_label.text = f'Всего записей: {len(records)}'


# ---------------------------------------------------------------------------
# Главное приложение
# ---------------------------------------------------------------------------

class ArshinApp(App):
    title = 'ФГИС АРШИН'

    def build(self):
        Window.softinput_mode = 'below_target'
        Window.clearcolor = (0.97, 0.97, 0.97, 1)
        Window.size = (420, 750)

        self.db = Database()
        self.api = APIClient()

        panel = TabbedPanel(do_default_tab=False)
        panel.tab_pos = 'top'
        panel.background_color = (0.97, 0.97, 0.97, 1)

        self.search_tab = SearchTab(self)
        self.excel_tab = ExcelTab(self)
        self.history_tab = HistoryTab(self)

        panel.add_widget(self.search_tab)
        panel.add_widget(self.excel_tab)
        panel.add_widget(self.history_tab)
        panel.default_tab = self.search_tab

        panel.bind(current_tab=self._on_tab_switch)

        from kivy.clock import Clock
        Clock.schedule_once(lambda dt: self._init_history(), 0.3)

        return panel

    def _init_history(self):
        if hasattr(self, 'history_tab'):
            self.history_tab.load_history()

    def _on_tab_switch(self, instance, tab):
        if tab and tab.text == 'История':
            self.history_tab.load_history()

    def on_stop(self):
        self.api.close()
        logger.info("Приложение остановлено")

    # -------------------------------------------------------------------
    # Логика поиска (в фоновом потоке)
    # -------------------------------------------------------------------

    def _search_thread(self, tab, serial, years, use_attr):
        all_records = []
        total_found = 0

        try:
            for year in years:
                if use_attr:
                    result = self.api.search_with_retry(
                        search_term=None, year=year,
                        filters={'mi_number': serial},
                    )
                else:
                    result = self.api.search_with_retry(
                        search_term=serial, year=year,
                    )

                items = result.get('items', [])
                for item in items:
                    mit_title = item.get('mit_title', '')
                    mit_notation = item.get('mit_notation', '')
                    mi_modification = item.get('mi_modification', '')
                    record_serial = item.get('mi_number', '')

                    if not is_electric_meter_from_queries(
                        mit_title, mit_notation, mi_modification
                    ):
                        continue

                    if not match_serial_number(record_serial, serial):
                        continue

                    record = VerificationRecord.from_api(item, search_query=serial)
                    all_records.append(record)
                    total_found += 1

            if all_records:
                self.db.save_records_batch(all_records)

            self._on_search_done(tab, serial, total_found, all_records)

        except Exception as e:
            logger.error(f"Ошибка поиска: {e}")
            self._on_search_error(tab, str(e))

    @mainthread
    def _on_search_done(self, tab, serial, total, records):
        if total == 0:
            tab.status_label.text = f'По запросу «{serial}» ничего не найдено'
            return

        tab.status_label.text = f'Найдено: {total} по «{serial}»'
        tab.results_box.clear_widgets()

        for r in records:
            card = ResultCard()
            card.title_text = r.mit_title or '—'
            card.date_text = r.verification_date or ''
            card.serial_text = (
                f"Зав. № {r.mi_number}" if r.mi_number else ''
            )
            card.manufacturer_text = r.manufacturer or ''
            card.applicable = r.applicability
            card.status_text = 'Пригоден' if r.applicability else 'Непригоден'
            card.verifier_text = r.org_title or ''
            tab.results_box.add_widget(card)

    @mainthread
    def _on_search_error(self, tab, msg):
        tab.status_label.text = f'Ошибка: {msg}'

    # -------------------------------------------------------------------
    # Пакетный поиск
    # -------------------------------------------------------------------

    def _batch_thread(self, tab, filepath, years):
        try:
            queries = ExcelHandler.read_queries(filepath)
        except Exception as e:
            self._on_batch_status(tab, f"Ошибка чтения: {e}")
            return

        if not queries:
            self._on_batch_status(tab, "Нет запросов в файле")
            return

        self._on_batch_status(tab, f"Загружено {len(queries)} запросов. Поиск...")

        total_found = total_saved = total_dupes = processed = 0
        api = APIClient()

        try:
            for query in queries:
                serial = query.get('mi_number', '')
                if not serial or str(serial).lower() == 'nan':
                    processed += 1
                    continue

                row_index = query.get('row_index', 0)
                id_pu = query.get('id_pu', '')
                original = query.get('original_data', {})
                buffer = []

                for year in years:
                    result = api.search_with_retry(
                        search_term=None, year=year,
                        filters={'mi_number': serial},
                    )
                    for item in result.get('items', []):
                        mit_title = item.get('mit_title', '')
                        mit_notation = item.get('mit_notation', '')
                        mi_modification = item.get('mi_modification', '')
                        record_serial = item.get('mi_number', '')

                        if not is_electric_meter_from_queries(
                            mit_title, mit_notation, mi_modification
                        ):
                            continue
                        if not match_serial_number(record_serial, serial):
                            continue

                        record = VerificationRecord.from_api(
                            item, search_query=serial,
                            row_index=row_index, id_pu=id_pu,
                            original_data=original,
                        )
                        buffer.append(record)
                        total_found += 1

                if buffer:
                    sv = self.db.save_records_batch(buffer)
                    total_saved += sv['saved']
                    total_dupes += sv['duplicates']

                processed += 1
                short = serial[:20] + '...' if len(serial) > 20 else serial
                self._on_batch_status(
                    tab,
                    f"[{processed}/{len(queries)}] «{short}» → {len(buffer)} зап."
                )

        except Exception as e:
            logger.error(f"Ошибка пакетного поиска: {e}")
            self._on_batch_status(tab, f"Ошибка: {e}")
        finally:
            api.close()

        self._on_batch_done(tab, processed, total_found, total_saved, total_dupes)

    @mainthread
    def _on_batch_status(self, tab, text):
        tab.ex_status.text = text

    @mainthread
    def _on_batch_done(self, tab, processed, found, saved, dupes):
        tab.ex_status.text = 'Готово!'
        tab.ex_results.text = (
            f"Обработано: {processed}\n"
            f"Найдено: {found}\n"
            f"Сохранено: {saved} (дублей: {dupes})"
        )

    # -------------------------------------------------------------------
    # Экспорт / очистка
    # -------------------------------------------------------------------

    def export_csv(self):
        export_dir = Config.EXPORT_DIR
        os.makedirs(export_dir, exist_ok=True)
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        path = os.path.join(export_dir, f'arshin_export_{ts}.csv')
        count = self.db.export_to_csv(path)

        if count == 0:
            self._show_popup('Экспорт', 'Нет данных для экспорта')
            return

        msg = f'Экспортировано {count} записей\n{path}'

        if platform == 'android':
            try:
                from android.storage import primary_external_storage_path
                import shutil
                dest = os.path.join(
                    primary_external_storage_path(),
                    'Download', f'arshin_export_{ts}.csv',
                )
                shutil.copy2(path, dest)
                msg += f'\nТакже: {dest}'
            except Exception as e:
                logger.warning(f"Копирование в Downloads: {e}")

        self._show_popup('Экспорт', msg)

    def clear_history(self, history_tab):
        def do_clear(instance):
            self.db.clear()
            history_tab.load_history()
            popup.dismiss()

        content = BoxLayout(orientation='vertical', spacing=dp(8))
        content.add_widget(Label(
            text='Удалить ВСЕ сохранённые записи?', font_size='14sp',
        ))

        btn_row = BoxLayout(
            orientation='horizontal',
            size_hint_y=None, height=dp(40), spacing=dp(8),
        )
        popup = Popup(
            title='Очистка истории', content=content,
            size_hint=(0.8, 0.35),
        )
        btn_row.add_widget(Button(
            text='Отмена', on_press=popup.dismiss,
        ))
        btn_row.add_widget(Button(
            text='Удалить',
            background_color=(0.8, 0.3, 0.3, 1),
            color=(1, 1, 1, 1),
            on_press=do_clear,
        ))
        content.add_widget(btn_row)
        popup.open()

    def _show_popup(self, title, message):
        content = BoxLayout(orientation='vertical', spacing=dp(8))
        content.add_widget(Label(text=message, font_size='13sp'))
        popup = Popup(title=title, content=content, size_hint=(0.8, 0.4))
        btn = Button(text='OK', size_hint_y=None, height=dp(40))
        btn.bind(on_press=popup.dismiss)
        content.add_widget(btn)
        popup.open()


if __name__ == '__main__':
    ArshinApp().run()
