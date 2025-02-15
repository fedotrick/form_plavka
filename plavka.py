import sys
import os
import re
import logging
from datetime import datetime, date, time
from typing import Optional, List, Dict, Set, Any
from dataclasses import dataclass
from enum import Enum

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QMessageBox, QLabel, QFrame, QDateEdit, QComboBox,
    QTableWidget, QTableWidgetItem, QDialog, QFileDialog, QGroupBox,
    QGridLayout, QTabWidget, QTextEdit, QScrollArea, QLineEdit
)
from PySide6.QtCore import Qt, QDate, QTime
from PySide6.QtGui import QColor, QShortcut, QKeySequence
from PySide6.QtWidgets import QGraphicsDropShadowEffect

from database import Database
from models import MeltRecord, SectorData
from constants import (
    TIME_FORMAT, DATE_FORMAT, TEMPERATURE_RANGE, MAX_PARTICIPANTS,
    PARTICIPANTS, CASTING_NAMES, ExperimentType, SectorName
)

# Nord цветовая схема
NORD_COLORS = {
    'POLAR_NIGHT': {
        'nord0': '#2E3440',
        'nord1': '#3B4252',
        'nord2': '#434C5E',
        'nord3': '#4C566A'
    },
    'SNOW_STORM': {
        'nord4': '#D8DEE9',
        'nord5': '#E5E9F0',
        'nord6': '#ECEFF4'
    },
    'FROST': {
        'nord7': '#8FBCBB',
        'nord8': '#88C0D0',
        'nord9': '#81A1C1',
        'nord10': '#5E81AC'
    },
    'AURORA': {
        'nord11': '#BF616A',
        'nord12': '#D08770',
        'nord13': '#EBCB8B',
        'nord14': '#A3BE8C',
        'nord15': '#B48EAD'
    }
}

# Стили для виджетов
NORD_STYLESHEET = f"""
QMainWindow {{
    background-color: {NORD_COLORS['POLAR_NIGHT']['nord0']};
}}

QWidget {{
    background-color: {NORD_COLORS['POLAR_NIGHT']['nord1']};
    color: {NORD_COLORS['SNOW_STORM']['nord4']};
}}

QLabel {{
    color: {NORD_COLORS['SNOW_STORM']['nord4']};
    font-size: 12px;
}}

QLineEdit, QTextEdit, QDateEdit, QComboBox {{
    background-color: {NORD_COLORS['POLAR_NIGHT']['nord2']};
    color: {NORD_COLORS['SNOW_STORM']['nord6']};
    border: 1px solid {NORD_COLORS['FROST']['nord9']};
    border-radius: 4px;
    padding: 5px;
    font-size: 12px;
}}

QLineEdit:focus, QTextEdit:focus, QDateEdit:focus, QComboBox:focus {{
    border: 2px solid {NORD_COLORS['FROST']['nord8']};
}}

QPushButton {{
    background-color: {NORD_COLORS['FROST']['nord9']};
    color: {NORD_COLORS['SNOW_STORM']['nord6']};
    border: none;
    border-radius: 4px;
    padding: 8px 16px;
    font-size: 12px;
}}

QPushButton:hover {{
    background-color: {NORD_COLORS['FROST']['nord10']};
}}

QPushButton:pressed {{
    background-color: {NORD_COLORS['FROST']['nord8']};
}}

QGroupBox {{
    border: 2px solid {NORD_COLORS['FROST']['nord9']};
    border-radius: 6px;
    margin-top: 1em;
    padding-top: 10px;
    font-size: 13px;
    font-weight: bold;
}}

QGroupBox::title {{
    color: {NORD_COLORS['FROST']['nord8']};
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 5px 0 5px;
}}

QComboBox::drop-down {{
    border: none;
}}

QComboBox::down-arrow {{
    image: none;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 5px solid {NORD_COLORS['FROST']['nord8']};
    width: 0;
    height: 0;
    margin-right: 5px;
}}

QTableWidget {{
    background-color: {NORD_COLORS['POLAR_NIGHT']['nord2']};
    alternate-background-color: {NORD_COLORS['POLAR_NIGHT']['nord1']};
    gridline-color: {NORD_COLORS['FROST']['nord9']};
}}

QTableWidget::item {{
    padding: 5px;
}}

QHeaderView::section {{
    background-color: {NORD_COLORS['POLAR_NIGHT']['nord3']};
    color: {NORD_COLORS['SNOW_STORM']['nord4']};
    padding: 5px;
    border: none;
}}

QScrollBar:vertical {{
    border: none;
    background-color: {NORD_COLORS['POLAR_NIGHT']['nord1']};
    width: 10px;
    margin: 0;
}}

QScrollBar::handle:vertical {{
    background-color: {NORD_COLORS['FROST']['nord9']};
    border-radius: 5px;
    min-height: 20px;
}}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    border: none;
    background: none;
}}
"""

# Настройка логирования
logging.basicConfig(
    filename='plavka.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Константы
TIME_FORMAT = "HH:mm"
TEMPERATURE_RANGE = (500, 2000)

class AutoCompleteComboBox(QComboBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setEditable(True)
        self.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        completer = self.completer()
        completer.setFilterMode(Qt.MatchFlag.MatchContains)
        completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)

    def keyPressEvent(self, event):
        """Обработка нажатий клавиш"""
        super().keyPressEvent(event)
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            self.addItem(self.currentText())
            self.setCurrentText("")

class PreviewDialog(QDialog):
    """
    Диалог предварительного просмотра данных перед сохранением.
    
    Показывает все введенные данные в удобном для просмотра формате,
    позволяя пользователю проверить корректность перед сохранением.
    """
    def __init__(self, record: MeltRecord, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.record = record
        self.setWindowTitle("Предпросмотр данных")
        self.setMinimumWidth(600)
        self.setup_ui()
    
    def setup_ui(self) -> None:
        """Настройка пользовательского интерфейса"""
        layout = QVBoxLayout(self)
        
        # Создаем текстовый браузер для отображения данных
        preview = QTextEdit()
        preview.setReadOnly(True)
        
        # Формируем текст предпросмотра
        text = f"""
        <h2>Основная информация</h2>
        <p><b>Дата плавки:</b> {self.record.date}</p>
        <p><b>Номер плавки:</b> {self.record.plavka_number}</p>
        <p><b>Номер кластера:</b> {self.record.cluster_number}</p>
        
        <h2>Участники</h2>
        <p><b>Старший смены:</b> {self.record.senior_shift}</p>
        <p><b>Первый участник:</b> {self.record.participant1}</p>
        <p><b>Второй участник:</b> {self.record.participant2}</p>
        <p><b>Третий участник:</b> {self.record.participant3}</p>
        <p><b>Четвертый участник:</b> {self.record.participant4}</p>
        
        <h2>Параметры плавки</h2>
        <p><b>Наименование отливки:</b> {self.record.casting_name}</p>
        <p><b>Тип эксперимента:</b> {self.record.experiment_type}</p>
        
        <h2>Сектор A</h2>
        <p><b>Номер сектора:</b> {self.record.sector_a.sector_number}</p>
        <p><b>Время прогрева ковша:</b> {self.record.sector_a.heating_time}</p>
        <p><b>Время перемещения:</b> {self.record.sector_a.movement_time}</p>
        <p><b>Время заливки:</b> {self.record.sector_a.pouring_time}</p>
        <p><b>Температура заливки:</b> {self.record.sector_a.temperature}°C</p>
        
        <h2>Сектор B</h2>
        <p><b>Номер сектора:</b> {self.record.sector_b.sector_number}</p>
        <p><b>Время прогрева ковша:</b> {self.record.sector_b.heating_time}</p>
        <p><b>Время перемещения:</b> {self.record.sector_b.movement_time}</p>
        <p><b>Время заливки:</b> {self.record.sector_b.pouring_time}</p>
        <p><b>Температура заливки:</b> {self.record.sector_b.temperature}°C</p>
        
        <h2>Сектор C</h2>
        <p><b>Номер сектора:</b> {self.record.sector_c.sector_number}</p>
        <p><b>Время прогрева ковша:</b> {self.record.sector_c.heating_time}</p>
        <p><b>Время перемещения:</b> {self.record.sector_c.movement_time}</p>
        <p><b>Время заливки:</b> {self.record.sector_c.pouring_time}</p>
        <p><b>Температура заливки:</b> {self.record.sector_c.temperature}°C</p>
        
        <h2>Сектор D</h2>
        <p><b>Номер сектора:</b> {self.record.sector_d.sector_number}</p>
        <p><b>Время прогрева ковша:</b> {self.record.sector_d.heating_time}</p>
        <p><b>Время перемещения:</b> {self.record.sector_d.movement_time}</p>
        <p><b>Время заливки:</b> {self.record.sector_d.pouring_time}</p>
        <p><b>Температура заливки:</b> {self.record.sector_d.temperature}°C</p>
        
        <h2>Комментарий</h2>
        <p>{self.record.comment}</p>
        """
        
        preview.setHtml(text)
        layout.addWidget(preview)
        
        # Кнопки
        button_layout = QHBoxLayout()
        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.accept)
        button_layout.addWidget(ok_button)
        
        layout.addLayout(button_layout)

class MainWindow(QMainWindow):
    """
    Главное окно приложения для учета плавок.
    
    Предоставляет интерфейс для создания и редактирования записей о плавках,
    с поддержкой горячих клавиш и автодополнения.
    """
    def __init__(self):
        super().__init__()
        self.db = Database()
        
        # Константы
        self.TEMPERATURE_RANGE = (1300, 1800)  # Допустимый диапазон температур
        
        # Создаем словарь для хранения последних использованных значений
        self.recent_values = {
            'casting_names': set(),
            'experiment_types': set(),
            'participants': set()
        }
        
        # Инициализация UI
        self.init_ui()
        
        # Генерируем начальный номер плавки
        self.record_id = None  # ID записи при редактировании
        self.generate_plavka_number()
        
        # Настройка компонентов
        self.setup_shortcuts()
        self.setup_field_connections()
        self.setup_tab_order()
        self.load_recent_values()
        
    def init_ui(self):
        """Инициализация пользовательского интерфейса"""
        # Применяем стиль Nord
        self.setStyleSheet(NORD_STYLESHEET)
        
        self.setWindowTitle("Электронный журнал плавки")
        self.setMinimumWidth(800)
        
        # Создаем центральный виджет и основной layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # Создаем группы для разных секций формы
        main_info_group = QGroupBox("Основная информация")
        participants_group = QGroupBox("Участники")
        melt_params_group = QGroupBox("Параметры плавки")
        sectors_group = QGroupBox("Секторы")
        comment_group = QGroupBox("Комментарий")
        
        # Добавляем группы в основной layout
        main_layout.addWidget(main_info_group)
        main_layout.addWidget(participants_group)
        main_layout.addWidget(melt_params_group)
        main_layout.addWidget(sectors_group)
        main_layout.addWidget(comment_group)
        
        # Создаем layout'ы для каждой группы
        main_info_layout = QGridLayout(main_info_group)
        participants_layout = QGridLayout(participants_group)
        melt_params_layout = QGridLayout(melt_params_group)
        sectors_layout = QGridLayout(sectors_group)
        comment_layout = QVBoxLayout(comment_group)
        
        # Основная информация
        self.Плавка_дата = QDateEdit()
        self.Плавка_дата.setCalendarPopup(True)
        self.Плавка_дата.setDate(QDate.currentDate())
        
        self.Номер_плавки = QLineEdit()
        self.Номер_кластера = QLineEdit()
        self.Учетный_номер = QLineEdit()
        self.Учетный_номер.setReadOnly(True)
        
        main_info_layout.addWidget(QLabel("Дата:"), 0, 0)
        main_info_layout.addWidget(self.Плавка_дата, 0, 1)
        main_info_layout.addWidget(QLabel("Номер плавки:"), 0, 2)
        main_info_layout.addWidget(self.Номер_плавки, 0, 3)
        main_info_layout.addWidget(QLabel("Номер кластера:"), 1, 0)
        main_info_layout.addWidget(self.Номер_кластера, 1, 1)
        main_info_layout.addWidget(QLabel("Учетный номер:"), 1, 2)
        main_info_layout.addWidget(self.Учетный_номер, 1, 3)
        
        # Участники
        self.Старший_смены_плавки = AutoCompleteComboBox()
        self.Первый_участник_смены_плавки = AutoCompleteComboBox()
        self.Второй_участник_смены_плавки = AutoCompleteComboBox()
        self.Третий_участник_смены_плавки = AutoCompleteComboBox()
        self.Четвертый_участник_смены_плавки = AutoCompleteComboBox()
        
        participants_layout.addWidget(QLabel("Старший смены:"), 0, 0)
        participants_layout.addWidget(self.Старший_смены_плавки, 0, 1)
        participants_layout.addWidget(QLabel("Участник 1:"), 1, 0)
        participants_layout.addWidget(self.Первый_участник_смены_плавки, 1, 1)
        participants_layout.addWidget(QLabel("Участник 2:"), 1, 2)
        participants_layout.addWidget(self.Второй_участник_смены_плавки, 1, 3)
        participants_layout.addWidget(QLabel("Участник 3:"), 2, 0)
        participants_layout.addWidget(self.Третий_участник_смены_плавки, 2, 1)
        participants_layout.addWidget(QLabel("Участник 4:"), 2, 2)
        participants_layout.addWidget(self.Четвертый_участник_смены_плавки, 2, 3)
        
        # Параметры плавки
        self.Наименование_отливки = AutoCompleteComboBox()
        self.Тип_эксперемента = AutoCompleteComboBox()
        
        melt_params_layout.addWidget(QLabel("Наименование отливки:"), 0, 0)
        melt_params_layout.addWidget(self.Наименование_отливки, 0, 1)
        melt_params_layout.addWidget(QLabel("Тип эксперимента:"), 0, 2)
        melt_params_layout.addWidget(self.Тип_эксперемента, 0, 3)
        
        # Секторы
        # Создаем виджеты для секторов
        self.create_sector_widgets()
        
        # Добавляем секторы в layout
        sectors_layout.addWidget(QLabel("Сектор"), 0, 0)
        sectors_layout.addWidget(QLabel("Номер опоки"), 0, 1)
        sectors_layout.addWidget(QLabel("Время прогрева"), 0, 2)
        sectors_layout.addWidget(QLabel("Время перемещения"), 0, 3)
        sectors_layout.addWidget(QLabel("Время заливки"), 0, 4)
        sectors_layout.addWidget(QLabel("Температура"), 0, 5)
        
        # Сектор A
        sectors_layout.addWidget(QLabel("A"), 1, 0)
        sectors_layout.addWidget(self.Сектор_A_опоки, 1, 1)
        sectors_layout.addWidget(self.Плавка_время_прогрева_ковша_A, 1, 2)
        sectors_layout.addWidget(self.Плавка_время_перемещения_A, 1, 3)
        sectors_layout.addWidget(self.Плавка_время_заливки_A, 1, 4)
        sectors_layout.addWidget(self.Плавка_температура_заливки_A, 1, 5)
        
        # Сектор B
        sectors_layout.addWidget(QLabel("B"), 2, 0)
        sectors_layout.addWidget(self.Сектор_B_опоки, 2, 1)
        sectors_layout.addWidget(self.Плавка_время_прогрева_ковша_B, 2, 2)
        sectors_layout.addWidget(self.Плавка_время_перемещения_B, 2, 3)
        sectors_layout.addWidget(self.Плавка_время_заливки_B, 2, 4)
        sectors_layout.addWidget(self.Плавка_температура_заливки_B, 2, 5)
        
        # Сектор C
        sectors_layout.addWidget(QLabel("C"), 3, 0)
        sectors_layout.addWidget(self.Сектор_C_опоки, 3, 1)
        sectors_layout.addWidget(self.Плавка_время_прогрева_ковша_C, 3, 2)
        sectors_layout.addWidget(self.Плавка_время_перемещения_C, 3, 3)
        sectors_layout.addWidget(self.Плавка_время_заливки_C, 3, 4)
        sectors_layout.addWidget(self.Плавка_температура_заливки_C, 3, 5)
        
        # Сектор D
        sectors_layout.addWidget(QLabel("D"), 4, 0)
        sectors_layout.addWidget(self.Сектор_D_опоки, 4, 1)
        sectors_layout.addWidget(self.Плавка_время_прогрева_ковша_D, 4, 2)
        sectors_layout.addWidget(self.Плавка_время_перемещения_D, 4, 3)
        sectors_layout.addWidget(self.Плавка_время_заливки_D, 4, 4)
        sectors_layout.addWidget(self.Плавка_температура_заливки_D, 4, 5)
        
        # Комментарий
        self.Комментарий = QTextEdit()
        comment_layout.addWidget(self.Комментарий)
        
        # Кнопки
        button_layout = QHBoxLayout()
        self.save_button = QPushButton("Сохранить")
        self.save_button.clicked.connect(self.save_data)
        self.search_button = QPushButton("Поиск")
        self.search_button.clicked.connect(self.show_search_dialog)
        
        button_layout.addWidget(self.save_button)
        button_layout.addWidget(self.search_button)
        main_layout.addLayout(button_layout)
        
        # Настраиваем размеры и отступы
        self.setMinimumSize(1000, 800)
        for layout in [main_info_layout, participants_layout, melt_params_layout, sectors_layout]:
            layout.setSpacing(10)
            layout.setContentsMargins(10, 10, 10, 10)

    def create_sector_widgets(self):
        """Создание виджетов для секторов"""
        # Сектор A
        self.Сектор_A_опоки = QLineEdit()
        self.Плавка_время_прогрева_ковша_A = QLineEdit()
        self.Плавка_время_прогрева_ковша_A.setInputMask("99:99")
        self.Плавка_время_прогрева_ковша_A.setProperty("time", "true")
        self.Плавка_время_перемещения_A = QLineEdit()
        self.Плавка_время_перемещения_A.setInputMask("99:99")
        self.Плавка_время_перемещения_A.setProperty("time", "true")
        self.Плавка_время_заливки_A = QLineEdit()
        self.Плавка_время_заливки_A.setInputMask("99:99")
        self.Плавка_время_заливки_A.setProperty("time", "true")
        self.Плавка_температура_заливки_A = QLineEdit()
        self.Плавка_температура_заливки_A.setProperty("temperature", "true")
        
        # Сектор B
        self.Сектор_B_опоки = QLineEdit()
        self.Плавка_время_прогрева_ковша_B = QLineEdit()
        self.Плавка_время_прогрева_ковша_B.setInputMask("99:99")
        self.Плавка_время_прогрева_ковша_B.setProperty("time", "true")
        self.Плавка_время_перемещения_B = QLineEdit()
        self.Плавка_время_перемещения_B.setInputMask("99:99")
        self.Плавка_время_перемещения_B.setProperty("time", "true")
        self.Плавка_время_заливки_B = QLineEdit()
        self.Плавка_время_заливки_B.setInputMask("99:99")
        self.Плавка_время_заливки_B.setProperty("time", "true")
        self.Плавка_температура_заливки_B = QLineEdit()
        self.Плавка_температура_заливки_B.setProperty("temperature", "true")
        
        # Сектор C
        self.Сектор_C_опоки = QLineEdit()
        self.Плавка_время_прогрева_ковша_C = QLineEdit()
        self.Плавка_время_прогрева_ковша_C.setInputMask("99:99")
        self.Плавка_время_прогрева_ковша_C.setProperty("time", "true")
        self.Плавка_время_перемещения_C = QLineEdit()
        self.Плавка_время_перемещения_C.setInputMask("99:99")
        self.Плавка_время_перемещения_C.setProperty("time", "true")
        self.Плавка_время_заливки_C = QLineEdit()
        self.Плавка_время_заливки_C.setInputMask("99:99")
        self.Плавка_время_заливки_C.setProperty("time", "true")
        self.Плавка_температура_заливки_C = QLineEdit()
        self.Плавка_температура_заливки_C.setProperty("temperature", "true")
        
        # Сектор D
        self.Сектор_D_опоки = QLineEdit()
        self.Плавка_время_прогрева_ковша_D = QLineEdit()
        self.Плавка_время_прогрева_ковша_D.setInputMask("99:99")
        self.Плавка_время_прогрева_ковша_D.setProperty("time", "true")
        self.Плавка_время_перемещения_D = QLineEdit()
        self.Плавка_время_перемещения_D.setInputMask("99:99")
        self.Плавка_время_перемещения_D.setProperty("time", "true")
        self.Плавка_время_заливки_D = QLineEdit()
        self.Плавка_время_заливки_D.setInputMask("99:99")
        self.Плавка_время_заливки_D.setProperty("time", "true")
        self.Плавка_температура_заливки_D = QLineEdit()
        self.Плавка_температура_заливки_D.setProperty("temperature", "true")
        
    def generate_plavka_number(self):
        """Генерирует номер плавки в формате месяц-номер"""
        try:
            current_date = self.Плавка_дата.date()
            current_month = current_date.month()
            current_year = current_date.year()
            
            next_number = 1  # По умолчанию начинаем с 1
            
            if os.path.exists('plavka.db'):
                try:
                    records = self.db.get_records()
                    if records:
                        # Фильтруем записи только текущего месяца и года
                        current_month_records = [
                            record for record in records 
                            if record['date'] and record['date'].startswith(f"{current_year}-{current_month:02d}")
                        ]
                        
                        if current_month_records:
                            # Ищем последний номер для текущего месяца
                            last_numbers = []
                            for record in current_month_records:
                                if not record['plavka_number']:
                                    continue
                                    
                                try:
                                    month, number = record['plavka_number'].split('-')
                                    if month == str(current_month) and number.isdigit():
                                        last_numbers.append(int(number))
                                except (ValueError, AttributeError):
                                    logging.warning(f"Некорректный формат номера плавки в БД: {record['plavka_number']}")
                                    continue
                            
                            if last_numbers:
                                next_number = max(last_numbers) + 1
                except Exception as db_error:
                    logging.error(f"Ошибка при чтении из базы данных: {str(db_error)}")
            
            # Форматируем номер плавки: месяц-номер(с ведущими нулями)
            new_plavka_number = f"{current_month}-{str(next_number).zfill(3)}"
            self.Номер_плавки.setText(new_plavka_number)
            
            # Обновляем учетный номер после генерации номера плавки
            self.update_uchet_number()
            
            return new_plavka_number
            
        except Exception as e:
            logging.error(f"Ошибка при генерации номера плавки: {str(e)}")
            self.Номер_плавки.clear()
            self.update_uchet_number()  # Очистит учетный номер, так как номер плавки пустой
            return None

    def update_uchet_number(self):
        """Обновляет учетный номер на основе номера плавки"""
        try:
            plavka_number = self.Номер_плавки.text().strip()
            if not plavka_number:
                logging.warning("Номер плавки пустой")
                self.Учетный_номер.clear()
                return None

            # Проверяем формат номера плавки (должен быть месяц-номер)
            if not re.match(r'^\d+-\d+$', plavka_number):
                logging.warning(f"Неверный формат номера плавки: {plavka_number}")
                self.Учетный_номер.clear()
                return None

            year = str(self.Плавка_дата.date().year())[-2:]  # Последние 2 цифры года
            uchet_number = f"{plavka_number}/{year}"
            self.Учетный_номер.setText(uchet_number)
            return uchet_number
        
        except Exception as e:
            logging.error(f"Ошибка при обновлении учетного номера: {str(e)}")
            self.Учетный_номер.clear()
            return None

    def generate_id(self, Плавка_дата, Номер_плавки):
        """Генерирует ID в формате год + номер_плавки"""
        try:
            if not Номер_плавки or not Плавка_дата:
                logging.warning("Не указана дата или номер плавки")
                return None
                
            year = Плавка_дата.year()
            номер_плавки = Номер_плавки.strip()
            
            # Проверяем формат номера плавки (должен быть месяц-номер)
            if not re.match(r'^\d+-\d+$', номер_плавки):
                logging.warning(f"Неверный формат номера плавки: {номер_плавки}")
                return None
                
            # Заменяем '-' на '.' для формата ID
            номер_плавки = номер_плавки.replace('-', '.')
            
            id_number = f"{year}{номер_плавки}"
            
            # Проверяем, что ID уникален
            if self.check_duplicate_id(id_number):
                logging.error(f"ID уже существует в базе данных: {id_number}")
                return None
                
            return id_number
            
        except Exception as e:
            logging.error(f"Ошибка при генерации ID: {str(e)}")
            return None
    
    def generate_учетный_номер(self, Плавка_дата, Номер_плавки):
        # Получаем последние две цифры года
        last_two_digits_year = str(Плавка_дата.year())[-2:]
        
        # Проверяем, что номер плавки не пустой
        if Номер_плавки:  
            return f"{Номер_плавки}/{last_two_digits_year}"
        else:
            QMessageBox.warning(self, "Ошибка")
        
        return None    

    def validate_time(self, time_str):
        """Проверка корректности ввода времени в формате ЧЧ:ММ"""
        try:
            hours, minutes = map(int, time_str.split(':'))
            if 0 <= hours < 24 and 0 <= minutes < 60:
                return True
        except ValueError:
            return False
        return False

    def check_duplicate_id(self, id_number):
        """Проверка существования ID в plavka.db"""
        try:
            if not os.path.exists('plavka.db'):
                return False
            
            return self.db.check_duplicate_id(id_number)
        except Exception as e:
            logging.error(f"Ошибка при проверке дубликата ID: {str(e)}")
            return False

    def validate_fields(self):
        # Проверка обязательных полей
        if not self.Номер_плавки.text().strip():
            QMessageBox.warning(self, "Ошибка", "Номер плавки обязателен")
            return False
        
        # Проверка температур заливки
        try:
            temp_A = float(self.Плавка_температура_заливки_A.text())
            temp_B = float(self.Плавка_температура_заливки_B.text())
            temp_C = float(self.Плавка_температура_заливки_C.text())
            temp_D = float(self.Плавка_температура_заливки_D.text())
            if not (500 <= temp_A <= 2000) or not (500 <= temp_B <= 2000) or not (500 <= temp_C <= 2000) or not (500 <= temp_D <= 2000):  # примерный диапазон
                QMessageBox.warning(self, "Ошибка", "Недопустимая температурa заливки")
                return False
        except ValueError:
            QMessageBox.warning(self, "Ошибка", "Температурa должна быть числом")
            return False
        
        return True

    def format_temperature(self, temp_str):
        """Форматирование температур в нужный формат"""
        try:
            temp = float(temp_str)
            return f"{temp:.1f}°C"
        except ValueError:
            return temp_str

    def save_data(self) -> bool:
        """
        Сохраняет данные формы в базу данных
        
        Returns:
            bool: True если сохранение успешно, False в случае ошибки
        """
        try:
            if not self.validate_fields():
                return False

            # Создаем объекты секторов
            sectors = {}
            for sector in SectorName:
                sector_lower = sector.name.lower()
                heating_time = getattr(self, f'Плавка_время_прогрева_ковша_{sector.name}').text()
                movement_time = getattr(self, f'Плавка_время_перемещения_{sector.name}').text()
                pouring_time = getattr(self, f'Плавка_время_заливки_{sector.name}').text()
                temperature = getattr(self, f'Плавка_температура_заливки_{sector.name}').text()
                
                if any([heating_time, movement_time, pouring_time, temperature]):
                    sectors[f'sector_{sector_lower}'] = SectorData(
                        sector_number=sector.name,
                        heating_time=datetime.strptime(heating_time, '%H:%M').time() if heating_time else None,
                        movement_time=datetime.strptime(movement_time, '%H:%M').time() if movement_time else None,
                        pouring_time=datetime.strptime(pouring_time, '%H:%M').time() if pouring_time else None,
                        temperature=float(temperature) if temperature else None
                    )

            # Создаем запись о плавке
            record = MeltRecord(
                id=self.generate_id(self.Плавка_дата.date(), self.Номер_плавки.text()),
                uchet_number=self.Учетный_номер.text(),
                date=self.Плавка_дата.date().toPython(),
                plavka_number=self.Номер_плавки.text(),
                cluster_number=self.Номер_кластера.text(),
                senior_shift=self.Старший_смены_плавки.currentText(),
                participant1=self.Первый_участник_смены_плавки.currentText(),
                participant2=self.Второй_участник_смены_плавки.currentText(),
                participant3=self.Третий_участник_смены_плавки.currentText(),
                participant4=self.Четвертый_участник_смены_плавки.currentText(),
                casting_name=self.Наименование_отливки.currentText(),
                experiment_type=ExperimentType(self.Тип_эксперемента.currentText()) if self.Тип_эксперемента.currentText() else None,
                comment=self.Комментарий.toPlainText(),
                **sectors
            )

            # Показываем предпросмотр данных
            preview = PreviewDialog(record, self)
            if preview.exec() == QDialog.DialogCode.Accepted:
                # Сохраняем в базу данных
                if self.db.save_plavka(record):
                    QMessageBox.information(self, "Успех", "Данные успешно сохранены")
                    self.clear_form()
                    return True
                else:
                    QMessageBox.critical(self, "Ошибка", "Не удалось сохранить данные")
                    return False

        except Exception as e:
            logging.error(f"Ошибка при сохранении данных: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка при сохранении: {str(e)}")
            return False

    def validate_fields(self) -> bool:
        """
        Проверяет корректность заполнения полей формы
        
        Returns:
            bool: True если все поля заполнены корректно, False в противном случае
        """
        # Проверяем обязательные поля
        if not self.Номер_плавки.text():
            QMessageBox.warning(self, "Ошибка", "Не указан номер плавки")
            return False

        # Проверяем формат времени для всех временных полей
        time_fields = [
            widget for widget in self.findChildren(QLineEdit)
            if widget.property("time")
        ]
        for field in time_fields:
            if field.text() and not self.validate_time(field.text()):
                QMessageBox.warning(self, "Ошибка", f"Неверный формат времени: {field.text()}")
                return False

        # Проверяем температурy
        temp_fields = [
            widget for widget in self.findChildren(QLineEdit)
            if widget.property("temperature")
        ]
        for field in temp_fields:
            if field.text():
                temp = self.format_temperature(field.text())
                if temp is None or not (TEMPERATURE_RANGE[0] <= temp <= TEMPERATURE_RANGE[1]):
                    QMessageBox.warning(
                        self, 
                        "Ошибка", 
                        f"Температурa должна быть в диапазоне от {TEMPERATURE_RANGE[0]} до {TEMPERATURE_RANGE[1]}"
                    )
                    return False

        return True

    def validate_time(self, time_str: str) -> bool:
        """
        Проверка корректности ввода времени в формате ЧЧ:ММ
        
        Args:
            time_str: Строка со временем для проверки
            
        Returns:
            bool: True если время в корректном формате, False в противном случае
        """
        if not time_str:
            return True
        try:
            datetime.strptime(time_str, '%H:%M')
            return True
        except ValueError:
            return False

    def format_temperature(self, temp_str: str) -> Optional[float]:
        """
        Форматирование строки температурy в число
        
        Args:
            temp_str: Строка с температурой для форматирования
            
        Returns:
            Optional[float]: Отформатированная температурa или None в случае ошибки
        """
        if not temp_str:
            return None
        try:
            # Удаляем все нечисловые символы, кроме точки и минуса
            temp_str = ''.join(c for c in temp_str if c.isdigit() or c in '.-')
            return float(temp_str)
        except ValueError:
            return None

    def clear_form(self):
        """Очистка всех полей формы"""
        # Очищаем основную информацию
        self.Номер_плавки.clear()
        self.Номер_кластера.clear()
        self.Учетный_номер.clear()
        
        # Очищаем участников
        self.Старший_смены_плавки.clearEditText()
        self.Первый_участник_смены_плавки.clearEditText()
        self.Второй_участник_смены_плавки.clearEditText()
        self.Третий_участник_смены_плавки.clearEditText()
        self.Четвертый_участник_смены_плавки.clearEditText()
        
        # Очищаем параметры плавки
        self.Наименование_отливки.clearEditText()
        self.Тип_эксперемента.clearEditText()
        
        # Очищаем секторы опок
        self.Сектор_A_опоки.clear()
        self.Сектор_B_опоки.clear()
        self.Сектор_C_опоки.clear()
        self.Сектор_D_опоки.clear()
        
        # Очищаем временные параметры
        # Сектор A
        self.Плавка_время_прогрева_ковша_A.clear()
        self.Плавка_время_перемещения_A.clear()
        self.Плавка_время_заливки_A.clear()
        self.Плавка_температура_заливки_A.clear()
        
        # Сектор B
        self.Плавка_время_прогрева_ковша_B.clear()
        self.Плавка_время_перемещения_B.clear()
        self.Плавка_время_заливки_B.clear()
        self.Плавка_температура_заливки_B.clear()
        
        # Сектор C
        self.Плавка_время_прогрева_ковша_C.clear()
        self.Плавка_время_перемещения_C.clear()
        self.Плавка_время_заливки_C.clear()
        self.Плавка_температура_заливки_C.clear()
        
        # Сектор D
        self.Плавка_время_прогрева_ковша_D.clear()
        self.Плавка_время_перемещения_D.clear()
        self.Плавка_время_заливки_D.clear()
        self.Плавка_температура_заливки_D.clear()
        
        # Очищаем комментарий
        self.Комментарий.clear()
        
        # Устанавливаем текущую дату
        self.Плавка_дата.setDate(QDate.currentDate())
        
        # Генерируем новый номер плавки
        self.generate_plavka_number()

    def setup_shortcuts(self):
        """Настройка горячих клавиш"""
        # Ctrl+S для сохранения
        save_shortcut = QShortcut(QKeySequence("Ctrl+S"), self)
        save_shortcut.activated.connect(self.save_data)
        
        # Ctrl+N для очистки формы
        new_shortcut = QShortcut(QKeySequence("Ctrl+N"), self)
        new_shortcut.activated.connect(self.clear_form)
        
        # Ctrl+F для поиска
        search_shortcut = QShortcut(QKeySequence("Ctrl+F"), self)
        search_shortcut.activated.connect(self.show_search_dialog)
        
        # F5 для генерации номера плавки
        generate_shortcut = QShortcut(QKeySequence("F5"), self)
        generate_shortcut.activated.connect(self.generate_plavka_number)

    def show_search_dialog(self):
        dialog = SearchDialog(self.db, parent=self)
        dialog.exec()

    def setup_field_connections(self):
        """Настройка связей между полями для автоматической валидации"""
        # Валидация времени при изменении
        time_fields = [
            widget for widget in self.findChildren(QLineEdit)
            if widget.property("time")
        ]
        for field in time_fields:
            field.textChanged.connect(lambda text, f=field: self.validate_time_field(f))
            
        # Валидация температур при изменении
        temp_fields = [
            widget for widget in self.findChildren(QLineEdit)
            if widget.property("temperature")
        ]
        for field in temp_fields:
            field.textChanged.connect(lambda text, f=field: self.validate_temperature_field(f))

    def setup_tab_order(self):
        """Настройка порядка перехода между полями по Tab"""
        # Основная информация
        self.setTabOrder(self.Плавка_дата, self.Номер_плавки)
        self.setTabOrder(self.Номер_плавки, self.Номер_кластера)
        self.setTabOrder(self.Номер_кластера, self.Учетный_номер)
        
        # Участники
        self.setTabOrder(self.Учетный_номер, self.Старший_смены_плавки)
        self.setTabOrder(self.Старший_смены_плавки, self.Первый_участник_смены_плавки)
        self.setTabOrder(self.Первый_участник_смены_плавки, self.Второй_участник_смены_плавки)
        self.setTabOrder(self.Второй_участник_смены_плавки, self.Третий_участник_смены_плавки)
        self.setTabOrder(self.Третий_участник_смены_плавки, self.Четвертый_участник_смены_плавки)
        
        # Параметры плавки
        self.setTabOrder(self.Четвертый_участник_смены_плавки, self.Наименование_отливки)
        self.setTabOrder(self.Наименование_отливки, self.Тип_эксперемента)
        
        # Секторы
        self.setTabOrder(self.Тип_эксперемента, self.Сектор_A_опоки)
        self.setTabOrder(self.Сектор_A_опоки, self.Сектор_B_опоки)
        self.setTabOrder(self.Сектор_B_опоки, self.Сектор_C_опоки)
        self.setTabOrder(self.Сектор_C_опоки, self.Сектор_D_опоки)
        
        # Параметры секторов
        # Сектор A
        self.setTabOrder(self.Сектор_D_опоки, self.Плавка_время_прогрева_ковша_A)
        self.setTabOrder(self.Плавка_время_прогрева_ковша_A, self.Плавка_время_перемещения_A)
        self.setTabOrder(self.Плавка_время_перемещения_A, self.Плавка_время_заливки_A)
        self.setTabOrder(self.Плавка_время_заливки_A, self.Плавка_температура_заливки_A)
        
        # Сектор B
        self.setTabOrder(self.Плавка_температура_заливки_A, self.Плавка_время_прогрева_ковша_B)
        self.setTabOrder(self.Плавка_время_прогрева_ковша_B, self.Плавка_время_перемещения_B)
        self.setTabOrder(self.Плавка_время_перемещения_B, self.Плавка_время_заливки_B)
        self.setTabOrder(self.Плавка_время_заливки_B, self.Плавка_температура_заливки_B)
        
        # Сектор C
        self.setTabOrder(self.Плавка_температура_заливки_B, self.Плавка_время_прогрева_ковша_C)
        self.setTabOrder(self.Плавка_время_прогрева_ковша_C, self.Плавка_время_перемещения_C)
        self.setTabOrder(self.Плавка_время_перемещения_C, self.Плавка_время_заливки_C)
        self.setTabOrder(self.Плавка_время_заливки_C, self.Плавка_температура_заливки_C)
        
        # Сектор D
        self.setTabOrder(self.Плавка_температура_заливки_C, self.Плавка_время_прогрева_ковша_D)
        self.setTabOrder(self.Плавка_время_прогрева_ковша_D, self.Плавка_время_перемещения_D)
        self.setTabOrder(self.Плавка_время_перемещения_D, self.Плавка_время_заливки_D)
        self.setTabOrder(self.Плавка_время_заливки_D, self.Плавка_температура_заливки_D)
        
        # Комментарий
        self.setTabOrder(self.Плавка_температура_заливки_D, self.Комментарий)

    def validate_time(self, time_str: str) -> bool:
        """
        Проверка корректности формата времени
        
        Args:
            time_str: Строка со временем для проверки
            
        Returns:
            bool: True если формат корректный, False иначе
        """
        if not time_str:
            return True
            
        try:
            datetime.strptime(time_str, '%H:%M')
            return True
        except ValueError:
            return False

    def validate_temperature(self, temp_str: str) -> bool:
        """
        Проверка корректности значения температурy
        
        Args:
            temp_str: Строка с температурой для проверки
            
        Returns:
            bool: True если значение корректное, False иначе
        """
        if not temp_str:
            return True
            
        try:
            temp = float(temp_str)
            return self.TEMPERATURE_RANGE[0] <= temp <= self.TEMPERATURE_RANGE[1]
        except ValueError:
            return False

    def format_time(self, time_str: str) -> Optional[time]:
        """
        Форматирование строки времени в объект time
        
        Args:
            time_str: Строка со временем для форматирования
            
        Returns:
            Optional[time]: Отформатированное время или None в случае ошибки
        """
        if not time_str:
            return None
            
        try:
            return datetime.strptime(time_str, '%H:%M').time()
        except ValueError:
            return None

    def format_temperature(self, temp_str: str) -> Optional[float]:
        """
        Форматирование строки температурy в число
        
        Args:
            temp_str: Строка с температурой для форматирования
            
        Returns:
            Optional[float]: Отформатированная температурa или None в случае ошибки
        """
        if not temp_str:
            return None
            
        try:
            temp = float(temp_str)
            if self.TEMPERATURE_RANGE[0] <= temp <= self.TEMPERATURE_RANGE[1]:
                return temp
            return None
        except ValueError:
            return None

    def load_recent_values(self):
        """Загрузка последних использованных значений из базы данных"""
        try:
            # Получаем последние записи из базы данных
            records = self.db.get_records()
            
            if not records:
                return
                
            # Собираем уникальные значения
            for record in records:
                # Участники
                if record.senior_shift:
                    self.recent_values['participants'].add(record.senior_shift)
                if record.participant1:
                    self.recent_values['participants'].add(record.participant1)
                if record.participant2:
                    self.recent_values['participants'].add(record.participant2)
                if record.participant3:
                    self.recent_values['participants'].add(record.participant3)
                if record.participant4:
                    self.recent_values['participants'].add(record.participant4)
                
                # Наименования отливок
                if record.casting_name:
                    self.recent_values['casting_names'].add(record.casting_name)
                
                # Типы экспериментов
                if record.experiment_type:
                    self.recent_values['experiment_types'].add(record.experiment_type)
            
            # Добавляем значения в комбобоксы
            # Участники
            self.Старший_смены_плавки.addItems(sorted(self.recent_values['participants']))
            self.Первый_участник_смены_плавки.addItems(sorted(self.recent_values['participants']))
            self.Второй_участник_смены_плавки.addItems(sorted(self.recent_values['participants']))
            self.Третий_участник_смены_плавки.addItems(sorted(self.recent_values['participants']))
            self.Четвертый_участник_смены_плавки.addItems(sorted(self.recent_values['participants']))
            
            # Наименования отливок
            self.Наименование_отливки.addItems(sorted(self.recent_values['casting_names']))
            
            # Типы экспериментов
            self.Тип_эксперемента.addItems([e.value for e in ExperimentType])
            
        except Exception as e:
            logging.error(f"Ошибка при загрузке последних значений: {str(e)}")
            QMessageBox.warning(
                self,
                "Предупреждение",
                "Не удалось загрузить последние использованные значения. "
                "Автодополнение может работать некорректно."
            )

    def validate_time_field(self, field: QLineEdit) -> None:
        """
        Валидация поля времени с визуальной обратной связью
        
        Args:
            field: Поле для валидации
        """
        text = field.text()
        if not text:  # Пустое поле допустимо
            field.setStyleSheet("")
            return
            
        if self.validate_time(text):
            field.setStyleSheet("QLineEdit { background-color: #e8f5e9; }")  # Светло-зеленый
        else:
            field.setStyleSheet("QLineEdit { background-color: #ffebee; }")  # Светло-красный

    def validate_temperature_field(self, field: QLineEdit) -> None:
        """
        Валидация поля температур с визуальной обратной связью
        
        Args:
            field: Поле для валидации
        """
        text = field.text()
        if not text:  # Пустое поле допустимо
            field.setStyleSheet("")
            return
            
        try:
            temp = float(text)
            if self.TEMPERATURE_RANGE[0] <= temp <= self.TEMPERATURE_RANGE[1]:
                field.setStyleSheet("QLineEdit { background-color: #e8f5e9; }")  # Светло-зеленый
            else:
                field.setStyleSheet("QLineEdit { background-color: #ffebee; }")  # Светло-красный
        except ValueError:
            field.setStyleSheet("QLineEdit { background-color: #ffebee; }")  # Светло-красный

class SearchDialog(QDialog):
    def __init__(self, db: Database, parent: Optional[QWidget] = None) -> None:
        """
        Инициализация диалога поиска
        
        Args:
            db: Объект базы данных
            parent: Родительский виджет
        """
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Поиск записей")
        self.setMinimumSize(1000, 700)
        
        # Добавляем тень для окна
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(20)
        shadow.setXOffset(0)
        shadow.setYOffset(0)
        shadow.setColor(QColor(0, 0, 0, 60))
        self.setGraphicsEffect(shadow)
        self.setup_ui()
        
    def setup_ui(self) -> None:
        """
        Initialize and setup the user interface components.
        
        This method creates all the input fields, labels, and buttons needed for editing
        a melt record. It organizes them in a scrollable layout for better usability.
        """
        layout = QVBoxLayout(self)
        
        # Добавляем фильтры
        filter_group = QGroupBox("Фильтры")
        filter_layout = QGridLayout()
        
        # Фильтр по дате
        self.date_from = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setDate(QDate.currentDate())
        
        filter_layout.addWidget(QLabel("Дата с:"), 0, 0)
        filter_layout.addWidget(self.date_from, 0, 1)
        filter_layout.addWidget(QLabel("по:"), 0, 2)
        filter_layout.addWidget(self.date_to, 0, 3)
        
        # Фильтр по типу отливки
        self.filter_casting = QComboBox()
        self.filter_casting.addItems(["Все"] + CASTING_NAMES)
        filter_layout.addWidget(QLabel("Тип отливки:"), 1, 0)
        filter_layout.addWidget(self.filter_casting, 1, 1)
        
        # Фильтр по температуре
        self.temp_from = QLineEdit()
        self.temp_to = QLineEdit()
        filter_layout.addWidget(QLabel("Температура от:"), 2, 0)
        filter_layout.addWidget(self.temp_from, 2, 1)
        filter_layout.addWidget(QLabel("до:"), 2, 2)
        filter_layout.addWidget(self.temp_to, 2, 3)
        
        filter_group.setLayout(filter_layout)
        layout.addWidget(filter_group)
        
        # Добавляем вкладки для результатов и статистики
        self.tab_widget = QTabWidget()
        
        # Вкладка результатов поиска
        search_tab = QWidget()
        search_layout = QVBoxLayout(search_tab)
        
        # Таблица результатов
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(10)
        self.results_table.setHorizontalHeaderLabels([
            "Дата", "Номер плавки", "Учетный номер", "Кластер",
            "Старший смены", "Отливка", "Тип эксперимента",
            "Температура A", "Температура B", "Температура C", "Температура D"
        ])
        search_layout.addWidget(self.results_table)
        
        self.tab_widget.addTab(search_tab, "Результаты поиска")
        
        # Вкладка статистики
        stats_tab = QWidget()
        stats_layout = QVBoxLayout(stats_tab)
        
        self.stats_text = QTextEdit()
        self.stats_text.setReadOnly(True)
        stats_layout.addWidget(self.stats_text)
        
        self.tab_widget.addTab(stats_tab, "Статистика")
        
        layout.addWidget(self.tab_widget)
        
        # Кнопки
        button_layout = QHBoxLayout()
        self.search_button = QPushButton("Поиск")
        self.search_button.clicked.connect(self.search_records)
        
        self.edit_button = QPushButton("Редактировать")
        self.edit_button.clicked.connect(self.edit_selected)
        
        self.export_button = QPushButton("Экспорт")
        self.export_button.clicked.connect(self.export_results)
        
        self.stats_button = QPushButton("Обновить статистику")
        self.stats_button.clicked.connect(self.update_statistics)
        
        self.backup_button = QPushButton("Создать резервную копию")
        self.backup_button.clicked.connect(self.create_backup)
        
        button_layout.addWidget(self.search_button)
        button_layout.addWidget(self.edit_button)
        button_layout.addWidget(self.export_button)
        button_layout.addWidget(self.stats_button)
        button_layout.addWidget(self.backup_button)
        
        layout.addLayout(button_layout)

    def search_records(self) -> None:
        """Поиск записей с применением фильтров"""
        try:
            # Получаем все записи
            records = self.db.get_records()
            
            # Применяем фильтры
            filtered_records = []
            for record in records:
                if self.apply_filters(record):
                    filtered_records.append(record)
            
            # Очищаем таблицу
            self.results_table.setRowCount(0)
            
            # Заполняем таблицу отфильтрованными данными
            for row, record in enumerate(filtered_records):
                self.results_table.insertRow(row)
                
                # Заполняем основные данные
                self.results_table.setItem(row, 0, QTableWidgetItem(record.date.strftime('%d.%m.%Y')))
                self.results_table.setItem(row, 1, QTableWidgetItem(record.plavka_number))
                self.results_table.setItem(row, 2, QTableWidgetItem(record.uchet_number))
                self.results_table.setItem(row, 3, QTableWidgetItem(record.cluster_number))
                self.results_table.setItem(row, 4, QTableWidgetItem(record.senior_shift))
                self.results_table.setItem(row, 5, QTableWidgetItem(record.casting_name))
                self.results_table.setItem(row, 6, QTableWidgetItem(
                    record.experiment_type.value if record.experiment_type else ""
                ))
                
                # Заполняем температурy секторов
                for i, sector in enumerate(['a', 'b', 'c', 'd']):
                    sector_data = getattr(record, f'sector_{sector}')
                    temp = str(sector_data.temperature) if sector_data and sector_data.temperature else ""
                    self.results_table.setItem(row, 7 + i, QTableWidgetItem(temp))
            
            # Обновляем статистику
            self.update_statistics()
            
        except Exception as e:
            logging.error(f"Ошибка при поиске записей: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка при поиске: {str(e)}")

    def apply_filters(self, record: MeltRecord) -> bool:
        """
        Применяет фильтры к записи
        
        Args:
            record: Запись для проверки
            
        Returns:
            bool: True если запись соответствует фильтрам, False в противном случае
        """
        # Фильтр по дате
        record_date = record.date
        if record_date:
            from_date = self.date_from.date().toPython()
            to_date = self.date_to.date().toPython()
            if not (from_date <= record_date <= to_date):
                return False
        
        # Фильтр по типу отливки
        selected_casting = self.filter_casting.currentText()
        if selected_casting != "Все" and record.casting_name != selected_casting:
            return False
        
        # Фильтр по температуре
        temp_from = self.temp_from.text()
        temp_to = self.temp_to.text()
        if temp_from or temp_to:
            try:
                temp_from = float(temp_from) if temp_from else float('-inf')
                temp_to = float(temp_to) if temp_to else float('inf')
                
                # Проверяем температурy во всех секторах
                record_temps = []
                for sector in ['a', 'b', 'c', 'd']:
                    sector_data = getattr(record, f'sector_{sector}')
                    if sector_data and sector_data.temperature:
                        record_temps.append(sector_data.temperature)
                
                # Если нет ни одной температурy в диапазоне, пропускаем запись
                if not any(temp_from <= temp <= temp_to for temp in record_temps):
                    return False
                    
            except ValueError:
                logging.warning("Некорректный формат температурy в фильтре")
                return False
        
        return True

    def update_statistics(self) -> None:
        """Обновляет статистику по данным"""
        try:
            records = self.db.get_records()
            
            # Собираем статистику
            total_records = len(records)
            casting_stats = {}
            temp_stats = {
                'min': float('inf'),
                'max': float('-inf'),
                'avg': 0.0,
                'count': 0
            }
            
            for record in records:
                # Статистика по отливкам
                if record.casting_name:
                    casting_stats[record.casting_name] = casting_stats.get(record.casting_name, 0) + 1
                
                # Статистика по температуре
                for sector in ['a', 'b', 'c', 'd']:
                    sector_data = getattr(record, f'sector_{sector}')
                    if sector_data and sector_data.temperature:
                        temp = sector_data.temperature
                        temp_stats['min'] = min(temp_stats['min'], temp)
                        temp_stats['max'] = max(temp_stats['max'], temp)
                        temp_stats['avg'] += temp
                        temp_stats['count'] += 1
            
            # Вычисляем среднюю температурy
            if temp_stats['count'] > 0:
                temp_stats['avg'] /= temp_stats['count']
            
            # Формируем текст статистики
            stats_text = f"Общее количество записей: {total_records}\n\n"
            
            stats_text += "Статистика по отливкам:\n"
            for casting, count in casting_stats.items():
                stats_text += f"{casting}: {count} ({count/total_records*100:.1f}%)\n"
            
            stats_text += f"\nСтатистика по температуре:\n"
            if temp_stats['count'] > 0:
                stats_text += f"Минимальная: {temp_stats['min']:.1f}°C\n"
                stats_text += f"Максимальная: {temp_stats['max']:.1f}°C\n"
                stats_text += f"Средняя: {temp_stats['avg']:.1f}°C\n"
            else:
                stats_text += "Нет данных о температуре\n"
            
            self.stats_text.setText(stats_text)
            
        except Exception as e:
            logging.error(f"Ошибка при обновлении статистики: {str(e)}")
            self.stats_text.setText(f"Ошибка при обновлении статистики: {str(e)}")

    def edit_selected(self) -> None:
        """Открывает диалог редактирования для выбранной записи"""
        selected_items = self.results_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Предупреждение", "Выберите запись для редактирования")
            return
        
        # Получаем ID записи из выбранной строки
        row = selected_items[0].row()
        record_id = self.results_table.item(row, 0).text()  # Предполагаем, что ID в первой колонке
        
        dialog = EditRecordDialog(record_id, self)
        if dialog.exec() == QDialog.Accepted:
            # Обновляем таблицу результатов
            self.search_records()

    def export_results(self) -> None:
        """Экспортирует результаты поиска в Excel файл"""
        try:
            filename, _ = QFileDialog.getSaveFileName(
                self, "Сохранить результаты", "", "Excel Files (*.xlsx)"
            )
            if not filename:
                return
                
            if not filename.endswith('.xlsx'):
                filename += '.xlsx'
            
            # Получаем данные из таблицы
            data = []
            headers = []
            for j in range(self.results_table.columnCount()):
                headers.append(self.results_table.horizontalHeaderItem(j).text())
            
            for i in range(self.results_table.rowCount()):
                row = []
                for j in range(self.results_table.columnCount()):
                    item = self.results_table.item(i, j)
                    row.append(item.text() if item else "")
                data.append(row)
            
            # Создаем DataFrame и сохраняем в Excel
            import pandas as pd
            df = pd.DataFrame(data, columns=headers)
            df.to_excel(filename, index=False)
            
            QMessageBox.information(self, "Успех", "Данные успешно экспортированы")
            
        except Exception as e:
            logging.error(f"Ошибка при экспорте данных: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка при экспорте: {str(e)}")

    def create_backup(self) -> None:
        """Создает резервную копию базы данных"""
        try:
            backup_dir = "backups"
            if not os.path.exists(backup_dir):
                os.makedirs(backup_dir)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = os.path.join(backup_dir, f"plavka_backup_{timestamp}.db")
            
            import shutil
            shutil.copy2("plavka.db", backup_path)
            
            QMessageBox.information(
                self, 
                "Успех", 
                f"Резервная копия создана:\n{backup_path}"
            )
            
        except Exception as e:
            logging.error(f"Ошибка при создании резервной копии: {str(e)}")
            QMessageBox.critical(
                self, 
                "Ошибка", 
                f"Произошла ошибка при создании резервной копии: {str(e)}"
            )

class EditRecordDialog(QDialog):
    """
    Dialog for editing an existing melt record.
    
    This dialog allows users to modify all fields of a melt record including
    sector data, participant information, and timing details.
    
    Args:
        record_id (str): The ID of the record to edit
        parent (Optional[QWidget]): The parent widget, defaults to None
    """
    def __init__(self, record_id: str, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.record_id = record_id
        self.db = Database()
        self.setWindowTitle(f"Редактирование записи {record_id}")
        self.setup_ui()
        self.load_record_data()
        
    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        # Создаем область прокрутки
        scroll_area = QScrollArea(self)
        scroll_area.setWidgetResizable(True)
        scroll_content = QFrame()
        content_layout = QVBoxLayout(scroll_content)
        
        # Список участников
        participants = [
            "Белков", "Карасев", "Ермаков", "Рабинович",
            "Валиулин", "Волков", "Семенов", "Левин",
            "Исмаилов", "Беляев", "Политов", "Кокшин",
            "Терентьев"
        ]
        participants.sort()
        
        # Список наименований отливок
        naimenovanie_otlivok = [
            "Вороток", "Ригель", "Ригель optima", "Блок-картер", "Колесо РИТМ",
            "Накладка резьб", "Блок цилиндров", "Диагональ optima"
        ]
        naimenovanie_otlivok.sort()
        
        # Список типов эксперементов
        types = ["Бумага", "Волокно"]
        types.sort()
        
        # Создаем поля ввода
        self.Плавка_дата = QDateEdit(self)
        self.Плавка_дата.setDisplayFormat("dd.MM.yyyy")
        self.Плавка_дата.setCalendarPopup(True)
        content_layout.addWidget(QLabel("Дата плавки:"))
        content_layout.addWidget(self.Плавка_дата)

        self.Номер_плавки = QLineEdit(self)
        content_layout.addWidget(QLabel("Номер плавки:"))
        content_layout.addWidget(self.Номер_плавки)

        self.Номер_кластера = QLineEdit(self)
        content_layout.addWidget(QLabel("Номер кластера:"))
        content_layout.addWidget(self.Номер_кластера)

        # Комбобоксы для участников
        self.Старший_смены_плавки = QComboBox(self)
        self.Старший_смены_плавки.addItems(participants)
        content_layout.addWidget(QLabel("Старший смены:"))
        content_layout.addWidget(self.Старший_смены_плавки)

        self.Первый_участник_смены_плавки = QComboBox(self)
        self.Первый_участник_смены_плавки.addItems(participants)
        content_layout.addWidget(QLabel("Первый участник:"))
        content_layout.addWidget(self.Первый_участник_смены_плавки)

        self.Второй_участник_смены_плавки = QComboBox(self)
        self.Второй_участник_смены_плавки.addItems(participants)
        content_layout.addWidget(QLabel("Второй участник:"))
        content_layout.addWidget(self.Второй_участник_смены_плавки)

        self.Третий_участник_смены_плавки = QComboBox(self)
        self.Третий_участник_смены_плавки.addItems(participants)
        content_layout.addWidget(QLabel("Третий участник:"))
        content_layout.addWidget(self.Третий_участник_смены_плавки)

        self.Четвертый_участник_смены_плавки = QComboBox(self)
        self.Четвертый_участник_смены_плавки.addItems(participants)
        content_layout.addWidget(QLabel("Четвертый участник:"))
        content_layout.addWidget(self.Четвертый_участник_смены_плавки)

        self.Наименование_отливки = QComboBox(self)
        self.Наименование_отливки.addItems(naimenovanie_otlivok)
        content_layout.addWidget(QLabel("Наименование отливки:"))
        content_layout.addWidget(self.Наименование_отливки)

        self.Тип_эксперемента = QComboBox(self)
        self.Тип_эксперемента.addItems(types)
        content_layout.addWidget(QLabel("Тип эксперимента:"))
        content_layout.addWidget(self.Тип_эксперемента)

        # Создаем поля для секторов опок
        self.Сектор_A_опоки = QLineEdit(self)
        content_layout.addWidget(QLabel("Сектор A опоки:"))
        content_layout.addWidget(self.Сектор_A_опоки)

        self.Сектор_B_опоки = QLineEdit(self)
        content_layout.addWidget(QLabel("Сектор B опоки:"))
        content_layout.addWidget(self.Сектор_B_опоки)

        self.Сектор_C_опоки = QLineEdit(self)
        content_layout.addWidget(QLabel("Сектор C опоки:"))
        content_layout.addWidget(self.Сектор_C_опоки)

        self.Сектор_D_опоки = QLineEdit(self)
        content_layout.addWidget(QLabel("Сектор D опоки:"))
        content_layout.addWidget(self.Сектор_D_опоки)

        # Создаем поля для временных параметров сектора A
        self.Плавка_время_прогрева_ковша_A = QLineEdit(self)
        self.Плавка_время_прогрева_ковша_A.setInputMask("99:99")
        self.Плавка_время_прогрева_ковша_A.setProperty("time", "true")
        content_layout.addWidget(QLabel("Время прогрева ковша (ЧЧ:ММ):"))
        content_layout.addWidget(self.Плавка_время_прогрева_ковша_A)

        self.Плавка_время_перемещения_A = QLineEdit(self)
        self.Плавка_время_перемещения_A.setInputMask("99:99")
        self.Плавка_время_перемещения_A.setProperty("time", "true")
        content_layout.addWidget(QLabel("Время перемещения (ЧЧ:ММ):"))
        content_layout.addWidget(self.Плавка_время_перемещения_A)

        self.Плавка_время_заливки_A = QLineEdit(self)
        self.Плавка_время_заливки_A.setInputMask("99:99")
        self.Плавка_время_заливки_A.setProperty("time", "true")
        content_layout.addWidget(QLabel("Время заливки (ЧЧ:ММ):"))
        content_layout.addWidget(self.Плавка_время_заливки_A)

        self.Плавка_температура_заливки_A = QLineEdit(self)
        self.Плавка_температура_заливки_A.setProperty("temperature", "true")
        content_layout.addWidget(QLabel("Температура заливки:"))
        content_layout.addWidget(self.Плавка_температура_заливки_A)

        # Создаем поля для временных параметров сектора B
        self.Плавка_время_прогрева_ковша_B = QLineEdit(self)
        self.Плавка_время_прогрева_ковша_B.setInputMask("99:99")
        self.Плавка_время_прогрева_ковша_B.setProperty("time", "true")
        content_layout.addWidget(QLabel("Время прогрева ковша (ЧЧ:ММ):"))
        content_layout.addWidget(self.Плавка_время_прогрева_ковша_B)

        self.Плавка_время_перемещения_B = QLineEdit(self)
        self.Плавка_время_перемещения_B.setInputMask("99:99")
        self.Плавка_время_перемещения_B.setProperty("time", "true")
        content_layout.addWidget(QLabel("Время перемещения (ЧЧ:ММ):"))
        content_layout.addWidget(self.Плавка_время_перемещения_B)

        self.Плавка_время_заливки_B = QLineEdit(self)
        self.Плавка_время_заливки_B.setInputMask("99:99")
        self.Плавка_время_заливки_B.setProperty("time", "true")
        content_layout.addWidget(QLabel("Время заливки (ЧЧ:ММ):"))
        content_layout.addWidget(self.Плавка_время_заливки_B)

        self.Плавка_температура_заливки_B = QLineEdit(self)
        self.Плавка_температура_заливки_B.setProperty("temperature", "true")
        content_layout.addWidget(QLabel("Температура заливки:"))
        content_layout.addWidget(self.Плавка_температура_заливки_B)

        # Создаем поля для временных параметров сектора C
        self.Плавка_время_прогрева_ковша_C = QLineEdit(self)
        self.Плавка_время_прогрева_ковша_C.setInputMask("99:99")
        self.Плавка_время_прогрева_ковша_C.setProperty("time", "true")
        content_layout.addWidget(QLabel("Время прогрева ковша (ЧЧ:ММ):"))
        content_layout.addWidget(self.Плавка_время_прогрева_ковша_C)

        self.Плавка_время_перемещения_C = QLineEdit(self)
        self.Плавка_время_перемещения_C.setInputMask("99:99")
        self.Плавка_время_перемещения_C.setProperty("time", "true")
        content_layout.addWidget(QLabel("Время перемещения (ЧЧ:ММ):"))
        content_layout.addWidget(self.Плавка_время_перемещения_C)

        self.Плавка_время_заливки_C = QLineEdit(self)
        self.Плавка_время_заливки_C.setInputMask("99:99")
        self.Плавка_время_заливки_C.setProperty("time", "true")
        content_layout.addWidget(QLabel("Время заливки (ЧЧ:ММ):"))
        content_layout.addWidget(self.Плавка_время_заливки_C)

        self.Плавка_температура_заливки_C = QLineEdit(self)
        self.Плавка_температура_заливки_C.setProperty("temperature", "true")
        content_layout.addWidget(QLabel("Температура заливки:"))
        content_layout.addWidget(self.Плавка_температура_заливки_C)

        # Создаем поля для временных параметров сектора D
        self.Плавка_время_прогрева_ковша_D = QLineEdit(self)
        self.Плавка_время_прогрева_ковша_D.setInputMask("99:99")
        self.Плавка_время_прогрева_ковша_D.setProperty("time", "true")
        content_layout.addWidget(QLabel("Время прогрева ковша (ЧЧ:ММ):"))
        content_layout.addWidget(self.Плавка_время_прогрева_ковша_D)

        self.Плавка_время_перемещения_D = QLineEdit(self)
        self.Плавка_время_перемещения_D.setInputMask("99:99")
        self.Плавка_время_перемещения_D.setProperty("time", "true")
        content_layout.addWidget(QLabel("Время перемещения (ЧЧ:ММ):"))
        content_layout.addWidget(self.Плавка_время_перемещения_D)

        self.Плавка_время_заливки_D = QLineEdit(self)
        self.Плавка_время_заливки_D.setInputMask("99:99")
        self.Плавка_время_заливки_D.setProperty("time", "true")
        content_layout.addWidget(QLabel("Время заливки (ЧЧ:ММ):"))
        content_layout.addWidget(self.Плавка_время_заливки_D)

        self.Плавка_температура_заливки_D = QLineEdit(self)
        self.Плавка_температура_заливки_D.setProperty("temperature", "true")
        content_layout.addWidget(QLabel("Температура заливки:"))
        content_layout.addWidget(self.Плавка_температура_заливки_D)

        # Создаем поле для комментария
        self.Комментарий = QTextEdit(self)
        self.Комментарий.setPlaceholderText("Введите комментарий...")
        content_layout.addWidget(QLabel("Комментарий:"))
        content_layout.addWidget(self.Комментарий)

        # Кнопки
        button_layout = QHBoxLayout()
        save_button = QPushButton("Сохранить изменения")
        cancel_button = QPushButton("Отмена")
        
        save_button.clicked.connect(self.save_changes)
        cancel_button.clicked.connect(self.reject)
        
        button_layout.addWidget(save_button)
        button_layout.addWidget(cancel_button)
        
        # Устанавливаем виджеты
        scroll_area.setWidget(scroll_content)
        layout.addWidget(scroll_area)
        layout.addLayout(button_layout)

    def load_record_data(self) -> None:
        """
        Load the record data from the database.
        
        This method retrieves the record from the database and fills the form fields
        with the record data. If an error occurs during loading, it shows an error message.
        """
        try:
            record = self.db.get_record(self.record_id)
            if record:
                self.fill_fields(record)
            else:
                raise ValueError(f"Record with ID {self.record_id} not found")
            
        except Exception as e:
            logging.error(f"Ошибка при загрузке записи: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Ошибка при загрузке записи: {str(e)}")

    def fill_fields(self, record: MeltRecord) -> None:
        """
        Fill the form fields with data from a MeltRecord.
        
        Args:
            record: The MeltRecord instance containing the data to display
            
        Raises:
            Exception: If there is an error filling any of the fields
        """
        try:
            # Заполняем основные поля
            self.Плавка_дата.setDate(QDate.fromString(record.date, "yyyy-MM-dd"))
            self.Номер_плавки.setText(record.plavka_number)
            self.Номер_кластера.setText(record.cluster_number)
            
            # Устанавливаем значения комбобоксов для участников
            self.Старший_смены_плавки.setCurrentText(record.senior_shift)
            self.Первый_участник_смены_плавки.setCurrentText(record.participant1)
            self.Второй_участник_смены_плавки.setCurrentText(record.participant2)
            self.Третий_участник_смены_плавки.setCurrentText(record.participant3)
            self.Четвертый_участник_смены_плавки.setCurrentText(record.participant4)
            
            self.Наименование_отливки.setCurrentText(record.casting_name)
            self.Тип_эксперемента.setCurrentText(record.experiment_type)
            
            # Заполняем секторы опоки
            self.Сектор_A_опоки.setText(record.sector_a.sector_number)
            self.Сектор_B_опоки.setText(record.sector_b.sector_number)
            self.Сектор_C_опоки.setText(record.sector_c.sector_number)
            self.Сектор_D_опоки.setText(record.sector_d.sector_number)
            
            # Заполняем данные сектора A
            self.Плавка_время_прогрева_ковша_A.setText(record.sector_a.heating_time)
            self.Плавка_время_перемещения_A.setText(record.sector_a.movement_time)
            self.Плавка_время_заливки_A.setText(record.sector_a.pouring_time)
            self.Плавка_температура_заливки_A.setText(str(record.sector_a.temperature))

            # Заполняем данные сектора B
            self.Плавка_время_прогрева_ковша_B.setText(record.sector_b.heating_time)
            self.Плавка_время_перемещения_B.setText(record.sector_b.movement_time)
            self.Плавка_время_заливки_B.setText(record.sector_b.pouring_time)
            self.Плавка_температура_заливки_B.setText(str(record.sector_b.temperature))

            # Заполняем данные сектора C
            self.Плавка_время_прогрева_ковша_C.setText(record.sector_c.heating_time)
            self.Плавка_время_перемещения_C.setText(record.sector_c.movement_time)
            self.Плавка_время_заливки_C.setText(record.sector_c.pouring_time)
            self.Плавка_температура_заливки_C.setText(str(record.sector_c.temperature))

            # Заполняем данные сектора D
            self.Плавка_время_прогрева_ковша_D.setText(record.sector_d.heating_time)
            self.Плавка_время_перемещения_D.setText(record.sector_d.movement_time)
            self.Плавка_время_заливки_D.setText(record.sector_d.pouring_time)
            self.Плавка_температура_заливки_D.setText(str(record.sector_d.temperature))

            self.Комментарий.setText(record.comment)
            
        except Exception as e:
            logging.error(f"Ошибка при заполнении полей: {str(e)}")
            raise

    def save_changes(self) -> None:
        """
        Save the changes made to the record in the database.
        
        This method creates a new MeltRecord instance from the form data and updates
        the database. It shows a success message if the update is successful, or an
        error message if something goes wrong.
        """
        try:
            # Создаем объекты для секторов
            sector_a = SectorData(
                sector_number=self.Сектор_A_опоки.text(),
                heating_time=datetime.strptime(self.Плавка_время_прогрева_ковша_A.text(), '%H:%M').time() if self.Плавка_время_прогрева_ковша_A.text() else None,
                movement_time=datetime.strptime(self.Плавка_время_перемещения_A.text(), '%H:%M').time() if self.Плавка_время_перемещения_A.text() else None,
                pouring_time=datetime.strptime(self.Плавка_время_заливки_A.text(), '%H:%M').time() if self.Плавка_время_заливки_A.text() else None,
                temperature=float(self.Плавка_температура_заливки_A.text()) if self.Плавка_температура_заливки_A.text() else None
            )
            
            sector_b = SectorData(
                sector_number=self.Сектор_B_опоки.text(),
                heating_time=datetime.strptime(self.Плавка_время_прогрева_ковша_B.text(), '%H:%M').time() if self.Плавка_время_прогрева_ковша_B.text() else None,
                movement_time=datetime.strptime(self.Плавка_время_перемещения_B.text(), '%H:%M').time() if self.Плавка_время_перемещения_B.text() else None,
                pouring_time=datetime.strptime(self.Плавка_время_заливки_B.text(), '%H:%M').time() if self.Плавка_время_заливки_B.text() else None,
                temperature=float(self.Плавка_температура_заливки_B.text()) if self.Плавка_температура_заливки_B.text() else None
            )
            
            sector_c = SectorData(
                sector_number=self.Сектор_C_опоки.text(),
                heating_time=datetime.strptime(self.Плавка_время_прогрева_ковша_C.text(), '%H:%M').time() if self.Плавка_время_прогрева_ковша_C.text() else None,
                movement_time=datetime.strptime(self.Плавка_время_перемещения_C.text(), '%H:%M').time() if self.Плавка_время_перемещения_C.text() else None,
                pouring_time=datetime.strptime(self.Плавка_время_заливки_C.text(), '%H:%M').time() if self.Плавка_время_заливки_C.text() else None,
                temperature=float(self.Плавка_температура_заливки_C.text()) if self.Плавка_температура_заливки_C.text() else None
            )
            
            sector_d = SectorData(
                sector_number=self.Сектор_D_опоки.text(),
                heating_time=datetime.strptime(self.Плавка_время_прогрева_ковша_D.text(), '%H:%M').time() if self.Плавка_время_прогрева_ковша_D.text() else None,
                movement_time=datetime.strptime(self.Плавка_время_перемещения_D.text(), '%H:%M').time() if self.Плавка_время_перемещения_D.text() else None,
                pouring_time=datetime.strptime(self.Плавка_время_заливки_D.text(), '%H:%M').time() if self.Плавка_время_заливки_D.text() else None,
                temperature=float(self.Плавка_температура_заливки_D.text()) if self.Плавка_температура_заливки_D.text() else None
            )

            # Создаем запись о плавке
            record = MeltRecord(
                id=self.record_id,
                uchet_number=self.Учетный_номер.text(),
                date=self.Плавка_дата.date().toPython(),
                plavka_number=self.Номер_плавки.text(),
                cluster_number=self.Номер_кластера.text(),
                senior_shift=self.Старший_смены_плавки.currentText(),
                participant1=self.Первый_участник_смены_плавки.currentText(),
                participant2=self.Второй_участник_смены_плавки.currentText(),
                participant3=self.Третий_участник_смены_плавки.currentText(),
                participant4=self.Четвертый_участник_смены_плавки.currentText(),
                casting_name=self.Наименование_отливки.currentText(),
                experiment_type=self.Тип_эксперемента.currentText(),
                sector_a=sector_a,
                sector_b=sector_b,
                sector_c=sector_c,
                sector_d=sector_d,
                comment=self.Комментарий.toPlainText()
            )

            # Обновляем запись в базе данных
            self.db.update_record(record)
            
            QMessageBox.information(self, "Успех", "Запись успешно обновлена")
            self.accept()
            
        except Exception as e:
            logging.error(f"Ошибка при сохранении изменений: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Ошибка при сохранении изменений: {str(e)}")

    def setup_field_connections(self) -> None:
        """Настройка связей между полями для автоматической валидации"""
        # Валидация времени при изменении
        time_fields = [
            self.Плавка_время_прогрева_ковша_A,
            self.Плавка_время_перемещения_A,
            self.Плавка_время_заливки_A,
            self.Плавка_время_прогрева_ковша_B,
            self.Плавка_время_перемещения_B,
            self.Плавка_время_заливки_B,
            self.Плавка_время_прогрева_ковша_C,
            self.Плавка_время_перемещения_C,
            self.Плавка_время_заливки_C,
            self.Плавка_время_прогрева_ковша_D,
            self.Плавка_время_перемещения_D,
            self.Плавка_время_заливки_D
        ]
        
        for field in time_fields:
            field.textChanged.connect(lambda text, f=field: self.validate_time_field(f))
            
        # Валидация температур при изменении
        temp_fields = [
            self.Плавка_температура_заливки_A,
            self.Плавка_температура_заливки_B,
            self.Плавка_температура_заливки_C,
            self.Плавка_температура_заливки_D
        ]
        
        for field in temp_fields:
            field.textChanged.connect(lambda text, f=field: self.validate_temperature_field(f))

    def setup_tab_order(self) -> None:
        """Настройка порядка перехода между полями по Tab"""
        # Основная информация
        self.setTabOrder(self.Плавка_дата, self.Номер_плавки)
        self.setTabOrder(self.Номер_плавки, self.Номер_кластера)
        
        # Участники
        self.setTabOrder(self.Номер_кластера, self.Старший_смены_плавки)
        self.setTabOrder(self.Старший_смены_плавки, self.Первый_участник_смены_плавки)
        self.setTabOrder(self.Первый_участник_смены_плавки, self.Второй_участник_смены_плавки)
        self.setTabOrder(self.Второй_участник_смены_плавки, self.Третий_участник_смены_плавки)
        self.setTabOrder(self.Третий_участник_смены_плавки, self.Четвертый_участник_смены_плавки)
        
        # Параметры плавки
        self.setTabOrder(self.Четвертый_участник_смены_плавки, self.Наименование_отливки)
        self.setTabOrder(self.Наименование_отливки, self.Тип_эксперемента)
        
        # Сектор A
        self.setTabOrder(self.Тип_эксперемента, self.Сектор_A_опоки)
        self.setTabOrder(self.Сектор_A_опоки, self.Плавка_время_прогрева_ковша_A)
        self.setTabOrder(self.Плавка_время_прогрева_ковша_A, self.Плавка_время_перемещения_A)
        self.setTabOrder(self.Плавка_время_перемещения_A, self.Плавка_время_заливки_A)
        self.setTabOrder(self.Плавка_время_заливки_A, self.Плавка_температура_заливки_A)
        
        # Сектор B
        self.setTabOrder(self.Плавка_температура_заливки_A, self.Сектор_B_опоки)
        self.setTabOrder(self.Сектор_B_опоки, self.Плавка_время_прогрева_ковша_B)
        self.setTabOrder(self.Плавка_время_прогрева_ковша_B, self.Плавка_время_перемещения_B)
        self.setTabOrder(self.Плавка_время_перемещения_B, self.Плавка_время_заливки_B)
        self.setTabOrder(self.Плавка_время_заливки_B, self.Плавка_температура_заливки_B)
        
        # Сектор C
        self.setTabOrder(self.Плавка_температура_заливки_B, self.Сектор_C_опоки)
        self.setTabOrder(self.Сектор_C_опоки, self.Плавка_время_прогрева_ковша_C)
        self.setTabOrder(self.Плавка_время_прогрева_ковша_C, self.Плавка_время_перемещения_C)
        self.setTabOrder(self.Плавка_время_перемещения_C, self.Плавка_время_заливки_C)
        self.setTabOrder(self.Плавка_время_заливки_C, self.Плавка_температура_заливки_C)
        
        # Сектор D
        self.setTabOrder(self.Плавка_температура_заливки_C, self.Сектор_D_опоки)
        self.setTabOrder(self.Сектор_D_опоки, self.Плавка_время_прогрева_ковша_D)
        self.setTabOrder(self.Плавка_время_прогрева_ковша_D, self.Плавка_время_перемещения_D)
        self.setTabOrder(self.Плавка_время_перемещения_D, self.Плавка_время_заливки_D)
        self.setTabOrder(self.Плавка_время_заливки_D, self.Плавка_температура_заливки_D)
        
        # Комментарий и кнопки
        self.setTabOrder(self.Плавка_температура_заливки_D, self.Комментарий)
        self.setTabOrder(self.Комментарий, self.save_button)
        self.setTabOrder(self.save_button, self.search_button)

    def validate_time_field(self, field: QLineEdit) -> None:
        """
        Валидация поля времени с визуальной обратной связью
        
        Args:
            field: Поле для валидации
        """
        text = field.text()
        if not text:  # Пустое поле допустимо
            field.setStyleSheet("")
            return
            
        if self.validate_time(text):
            field.setStyleSheet("QLineEdit { background-color: #e8f5e9; }")  # Светло-зеленый
        else:
            field.setStyleSheet("QLineEdit { background-color: #ffebee; }")  # Светло-красный

    def validate_temperature_field(self, field: QLineEdit) -> None:
        """
        Валидация поля температур с визуальной обратной связью
        
        Args:
            field: Поле для валидации
        """
        text = field.text()
        if not text:  # Пустое поле допустимо
            field.setStyleSheet("")
            return
            
        try:
            temp = float(text)
            if TEMPERATURE_RANGE[0] <= temp <= TEMPERATURE_RANGE[1]:
                field.setStyleSheet("QLineEdit { background-color: #e8f5e9; }")  # Светло-зеленый
            else:
                field.setStyleSheet("QLineEdit { background-color: #ffebee; }")  # Светло-красный
        except ValueError:
            field.setStyleSheet("QLineEdit { background-color: #ffebee; }")  # Светло-красный

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
