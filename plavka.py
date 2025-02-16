import sys
import os
import re
import logging
import shutil
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLineEdit,
    QPushButton, QMessageBox, QLabel, QScrollArea, QFrame,
    QDateEdit, QComboBox, QTableWidget, QTableWidgetItem,
    QHBoxLayout, QDialog, QFileDialog, QGroupBox, QGridLayout,
    QTabWidget, QTextEdit
)
from PySide6.QtCore import Qt, QDate
from PySide6 import QtGui
from openpyxl import Workbook, load_workbook
from datetime import datetime, timedelta
import pandas as pd
from PySide6.QtGui import QColor
from PySide6.QtWidgets import QGraphicsDropShadowEffect

# В начале файла добавить настройку логирования
logging.basicConfig(
    filename='plavka.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Вынести настройки в отдельные константы
EXCEL_FILENAME = 'plavka.xlsx'
TEMPERATURE_RANGE = (500, 2000)
TIME_FORMAT = "HH:mm"

# Добавляем новые константы
SEARCH_FIELDS = ['ID', 'Учетный_номер', 'Номер_плавки', 'Наименование_отливки']
EXPORT_FORMATS = {
    'CSV': '*.csv',
    'PDF': '*.pdf',
    'Excel': '*.xlsx'
}

# Добавляем новые константы
BACKUP_DIR = 'backups'
MAX_BACKUPS = 5  # Максимальное количество резервных копий

# Функция для сохранения данных в Excel
def save_to_excel(ID, Учетный_номер, Плавка_дата, Номер_плавки, Номер_кластера,
                  Старший_смены_плавки, Первый_участник_смены_плавки,
                  Второй_участник_смены_плавки, Третий_участник_смены_плавки,
                  Четвертый_участник_смены_плавки, Наименование_отливки,
                  Тип_эксперемента, Сектор_A_опоки, Сектор_B_опоки,
                  Сектор_C_опоки, Сектор_D_опоки, 
                  Плавка_время_прогрева_ковша_A, Плавка_время_перемещения_A, Плавка_время_заливки_A, Плавка_температура_заливки_A,
                  Плавка_время_прогрева_ковша_B, Плавка_время_перемещения_B, Плавка_время_заливки_B, Плавка_температура_заливки_B,
                  Плавка_время_прогрева_ковша_C, Плавка_время_перемещения_C, Плавка_время_заливки_C, Плавка_температура_заливки_C,
                  Плавка_время_прогрева_ковша_D, Плавка_время_перемещения_D, Плавка_время_заливки_D, Плавка_температура_заливки_D,
                  Комментарий
                  ):
    file_name = 'plavka.xlsx'

    if not os.path.exists(file_name):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Records"
        headers = ["ID", "Учетный_номер", "Плавка_дата", "Номер_плавки", "Номер_кластера",
                  "Старший_смены_плавки", "Первый_участник_смены_плавки",
                  "Второй_участник_смены_плавки", "Третий_участник_смены_плавки",
                  "Четвертый_участник_смены_плавки", "Наименование_отливки",
                  "Тип_эксперемента", "Сектор_A_опоки", "Сектор_B_опоки",
                  "Сектор_C_опоки", "Сектор_D_опоки",
                  "Плавка_время_прогрева_ковша_A", "Плавка_время_перемещения_A", "Плавка_время_заливки_A", "Плавка_температура_заливки_A",
                  "Плавка_время_прогрева_ковша_B", "Плавка_время_перемещения_B", "Плавка_время_заливки_B", "Плавка_температура_заливки_B",
                  "Плавка_время_прогрева_ковша_C", "Плавка_время_перемещения_C", "Плавка_время_заливки_C", "Плавка_температура_заливки_C",
                  "Плавка_время_прогрева_ковша_D", "Плавка_время_перемещения_D", "Плавка_время_заливки_D", "Плавка_температура_заливки_D",
                  "Комментарий"]
        sheet.append(headers)
    else:
        workbook = load_workbook(file_name)
        sheet = workbook.active

    # Создаем список данных в том же порядке, что и заголовки
    data = [
        ID, Учетный_номер, Плавка_дата, Номер_плавки, Номер_кластера,
        Старший_смены_плавки, Первый_участник_смены_плавки,
        Второй_участник_смены_плавки, Третий_участник_смены_плавки,
        Четвертый_участник_смены_плавки, Наименование_отливки,
        Тип_эксперемента, Сектор_A_опоки, Сектор_B_опоки,
        Сектор_C_опоки, Сектор_D_опоки,
        Плавка_время_прогрева_ковша_A, Плавка_время_перемещения_A, Плавка_время_заливки_A, Плавка_температура_заливки_A,
        Плавка_время_прогрева_ковша_B, Плавка_время_перемещения_B, Плавка_время_заливки_B, Плавка_температура_заливки_B,
        Плавка_время_прогрева_ковша_C, Плавка_время_перемещения_C, Плавка_время_заливки_C, Плавка_температура_заливки_C,
        Плавка_время_прогрева_ковша_D, Плавка_время_перемещения_D, Плавка_время_заливки_D, Плавка_температура_заливки_D,
        Комментарий
    ]

    # Добавляем данные в таблицу
    sheet.append(data)

    # Автоматически регулируем ширину столбцов
    for column in sheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        sheet.column_dimensions[column_letter].width = adjusted_width

    # Сохраняем файл
    try:
        workbook.save(file_name)
        return True
    except Exception as e:
        logging.error(f"Ошибка при сохранении в Excel: {str(e)}")
        return False

# Основное окно приложения
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Электронный журнал плавки")
        
        # Устанавливаем светлый фон в стиле Nord
        self.setStyleSheet("""
            QWidget {
                background-color: #eceff4;
                color: #2e3440;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            QPushButton {
                background-color: #5e81ac;
                color: #ffffff;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-size: 14px;
                font-weight: bold;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #81a1c1;
            }
            QPushButton:pressed {
                background-color: #4c566a;
            }
            QLineEdit, QDateEdit, QComboBox, QTextEdit {
                background-color: #ffffff;
                color: #2e3440;
                border: 2px solid #d8dee9;
                border-radius: 4px;
                padding: 6px;
                min-width: 150px;
                font-size: 13px;
            }
            QLineEdit:focus, QDateEdit:focus, QComboBox:focus, QTextEdit:focus {
                border: 2px solid #5e81ac;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid #4c566a;
                width: 0;
                height: 0;
                margin-right: 5px;
            }
            QGroupBox {
                border: 2px solid #d8dee9;
                border-radius: 6px;
                margin-top: 1em;
                padding: 15px;
                font-size: 14px;
                font-weight: bold;
                background-color: #e5e9f0;
            }
            QGroupBox::title {
                color: #5e81ac;
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
            QLabel {
                color: #2e3440;
                font-size: 13px;
                min-width: 120px;
            }
            QTextEdit {
                min-height: 80px;
            }
            /* Стили для полей с температурой */
            QLineEdit[temperature="true"] {
                color: #bf616a;
                font-weight: bold;
                background-color: #fff0f0;
            }
            /* Стили для полей со временем */
            QLineEdit[time="true"] {
                color: #2e7d32;
                background-color: #f0fff0;
            }
            /* Стили для заголовков секторов */
            QGroupBox[sector="true"] {
                background-color: #e5e9f0;
            }
            QGroupBox[sector="true"]::title {
                color: #5e81ac;
                font-size: 15px;
            }
            /* Скроллбары */
            QScrollBar:vertical {
                border: none;
                background-color: #e5e9f0;
                width: 10px;
                margin: 0;
            }
            QScrollBar::handle:vertical {
                background-color: #81a1c1;
                border-radius: 5px;
                min-height: 20px;
            }
            QScrollBar::handle:vertical:hover {
                background-color: #5e81ac;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0;
                background: none;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: none;
            }
        """)
        
        # Создаем все виджеты
        self.create_widgets()
        
        # Создаем основной layout
        main_layout = QHBoxLayout()  # Используем горизонтальный layout
        
        # Создаем левую колонку
        left_column = QVBoxLayout()
        left_column.setSpacing(10)
        
        # Создаем правую колонку
        right_column = QVBoxLayout()
        right_column.setSpacing(10)
        
        # Создаем группы для логического разделения элементов
        basic_info_group = QGroupBox("Основная информация")
        participants_group = QGroupBox("Участники")
        casting_group = QGroupBox("Параметры отливки")
        time_group = QGroupBox("Временные параметры")
        comment_group = QGroupBox("Комментарий")
        
        # Создаем grid layouts для каждой группы
        basic_grid = QGridLayout()
        basic_grid.setSpacing(10)
        participants_grid = QGridLayout()
        participants_grid.setSpacing(10)
        casting_grid = QGridLayout()
        casting_grid.setSpacing(10)
        time_grid = QGridLayout()
        time_grid.setSpacing(10)
        
        # Основная информация
        basic_grid.addWidget(QLabel("Дата:"), 0, 0)
        basic_grid.addWidget(self.Плавка_дата, 0, 1)
        basic_grid.addWidget(QLabel("Номер плавки:"), 1, 0)
        basic_grid.addWidget(self.Номер_плавки, 1, 1)
        basic_grid.addWidget(QLabel("Номер кластера:"), 2, 0)
        basic_grid.addWidget(self.Номер_кластера, 2, 1)
        basic_info_group.setLayout(basic_grid)
        
        # Участники
        participants_grid.addWidget(QLabel("Старший смены:"), 0, 0)
        participants_grid.addWidget(self.Старший_смены_плавки, 0, 1)
        participants_grid.addWidget(QLabel("Участник 1:"), 1, 0)
        participants_grid.addWidget(self.Первый_участник_смены_плавки, 1, 1)
        participants_grid.addWidget(QLabel("Участник 2:"), 2, 0)
        participants_grid.addWidget(self.Второй_участник_смены_плавки, 2, 1)
        participants_grid.addWidget(QLabel("Участник 3:"), 3, 0)
        participants_grid.addWidget(self.Третий_участник_смены_плавки, 3, 1)
        participants_grid.addWidget(QLabel("Участник 4:"), 4, 0)
        participants_grid.addWidget(self.Четвертый_участник_смены_плавки, 4, 1)
        participants_group.setLayout(participants_grid)
        
        # Параметры отливки
        casting_grid.addWidget(QLabel("Наименование:"), 0, 0)
        casting_grid.addWidget(self.Наименование_отливки, 0, 1, 1, 3)
        casting_grid.addWidget(QLabel("Тип эксперимента:"), 1, 0)
        casting_grid.addWidget(self.Тип_эксперемента, 1, 1, 1, 3)
        
        # Добавляем секторы опок в сетку
        casting_grid.addWidget(QLabel("Секторы опок:"), 2, 0)
        sectors_grid = QGridLayout()
        sectors_grid.addWidget(QLabel("A:"), 0, 0)
        sectors_grid.addWidget(self.Сектор_A_опоки, 0, 1)
        sectors_grid.addWidget(QLabel("B:"), 0, 2)
        sectors_grid.addWidget(self.Сектор_B_опоки, 0, 3)
        sectors_grid.addWidget(QLabel("C:"), 1, 0)
        sectors_grid.addWidget(self.Сектор_C_опоки, 1, 1)
        sectors_grid.addWidget(QLabel("D:"), 1, 2)
        sectors_grid.addWidget(self.Сектор_D_опоки, 1, 3)
        sectors_widget = QWidget()
        sectors_widget.setLayout(sectors_grid)
        casting_grid.addWidget(sectors_widget, 2, 1, 1, 3)
        casting_group.setLayout(casting_grid)
        
        # Временные параметры в сетку 2x2
        time_params_layout = QGridLayout()
        
        # Сектор A
        sector_a_group = QGroupBox("Сектор A")
        sector_a_group.setProperty("sector", "true")
        sector_a_layout = QGridLayout()
        sector_a_layout.addWidget(QLabel("Время прогрева:"), 0, 0)
        sector_a_layout.addWidget(self.Плавка_время_прогрева_ковша_A, 0, 1)
        sector_a_layout.addWidget(QLabel("Время перемещения:"), 1, 0)
        sector_a_layout.addWidget(self.Плавка_время_перемещения_A, 1, 1)
        sector_a_layout.addWidget(QLabel("Время заливки:"), 2, 0)
        sector_a_layout.addWidget(self.Плавка_время_заливки_A, 2, 1)
        sector_a_layout.addWidget(QLabel("Температура:"), 3, 0)
        sector_a_layout.addWidget(self.Плавка_температура_заливки_A, 3, 1)
        sector_a_group.setLayout(sector_a_layout)
        time_params_layout.addWidget(sector_a_group, 0, 0)

        # Сектор B
        sector_b_group = QGroupBox("Сектор B")
        sector_b_group.setProperty("sector", "true")
        sector_b_layout = QGridLayout()
        sector_b_layout.addWidget(QLabel("Время прогрева:"), 0, 0)
        sector_b_layout.addWidget(self.Плавка_время_прогрева_ковша_B, 0, 1)
        sector_b_layout.addWidget(QLabel("Время перемещения:"), 1, 0)
        sector_b_layout.addWidget(self.Плавка_время_перемещения_B, 1, 1)
        sector_b_layout.addWidget(QLabel("Время заливки:"), 2, 0)
        sector_b_layout.addWidget(self.Плавка_время_заливки_B, 2, 1)
        sector_b_layout.addWidget(QLabel("Температура:"), 3, 0)
        sector_b_layout.addWidget(self.Плавка_температура_заливки_B, 3, 1)
        sector_b_group.setLayout(sector_b_layout)
        time_params_layout.addWidget(sector_b_group, 0, 1)

        # Сектор C
        sector_c_group = QGroupBox("Сектор C")
        sector_c_group.setProperty("sector", "true")
        sector_c_layout = QGridLayout()
        sector_c_layout.addWidget(QLabel("Время прогрева:"), 0, 0)
        sector_c_layout.addWidget(self.Плавка_время_прогрева_ковша_C, 0, 1)
        sector_c_layout.addWidget(QLabel("Время перемещения:"), 1, 0)
        sector_c_layout.addWidget(self.Плавка_время_перемещения_C, 1, 1)
        sector_c_layout.addWidget(QLabel("Время заливки:"), 2, 0)
        sector_c_layout.addWidget(self.Плавка_время_заливки_C, 2, 1)
        sector_c_layout.addWidget(QLabel("Температура:"), 3, 0)
        sector_c_layout.addWidget(self.Плавка_температура_заливки_C, 3, 1)
        sector_c_group.setLayout(sector_c_layout)
        time_params_layout.addWidget(sector_c_group, 1, 0)

        # Сектор D
        sector_d_group = QGroupBox("Сектор D")
        sector_d_group.setProperty("sector", "true")
        sector_d_layout = QGridLayout()
        sector_d_layout.addWidget(QLabel("Время прогрева:"), 0, 0)
        sector_d_layout.addWidget(self.Плавка_время_прогрева_ковша_D, 0, 1)
        sector_d_layout.addWidget(QLabel("Время перемещения:"), 1, 0)
        sector_d_layout.addWidget(self.Плавка_время_перемещения_D, 1, 1)
        sector_d_layout.addWidget(QLabel("Время заливки:"), 2, 0)
        sector_d_layout.addWidget(self.Плавка_время_заливки_D, 2, 1)
        sector_d_layout.addWidget(QLabel("Температура:"), 3, 0)
        sector_d_layout.addWidget(self.Плавка_температура_заливки_D, 3, 1)
        sector_d_group.setLayout(sector_d_layout)
        time_params_layout.addWidget(sector_d_group, 1, 1)

        time_group.setLayout(time_params_layout)
        
        # Добавляем поле для комментария
        comment_layout = QVBoxLayout()
        comment_layout.addWidget(self.Комментарий)
        comment_group.setLayout(comment_layout)
        
        # Кнопки управления
        buttons_layout = QHBoxLayout()
        buttons_layout.addWidget(self.save_button)
        buttons_layout.addWidget(self.search_button)
        
        # Добавляем группы в колонки
        left_column.addWidget(basic_info_group)
        left_column.addWidget(participants_group)
        left_column.addWidget(casting_group)
        
        right_column.addWidget(time_group)
        right_column.addWidget(comment_group)
        right_column.addLayout(buttons_layout)
        
        # Добавляем колонки в основной layout
        left_widget = QWidget()
        left_widget.setLayout(left_column)
        right_widget = QWidget()
        right_widget.setLayout(right_column)
        
        main_layout.addWidget(left_widget)
        main_layout.addWidget(right_widget)
        
        # Устанавливаем основной layout
        self.setLayout(main_layout)
        
        # Устанавливаем размер окна
        self.setMinimumSize(1600, 850)

    def create_widgets(self):
        """Создание всех виджетов формы"""
        # Создаем основные поля ввода
        self.Плавка_дата = QDateEdit(self)
        self.Плавка_дата.setDisplayFormat("dd.MM.yyyy")
        self.Плавка_дата.setCalendarPopup(True)
        self.Плавка_дата.setDate(QDate.currentDate().addDays(-1))
        
        self.Номер_плавки = QLineEdit(self)
        self.Номер_плавки.setReadOnly(True)
        
        self.Номер_кластера = QLineEdit(self)
        
        # Создаем комбобоксы для участников
        self.Старший_смены_плавки = QComboBox(self)
        self.Первый_участник_смены_плавки = QComboBox(self)
        self.Второй_участник_смены_плавки = QComboBox(self)
        self.Третий_участник_смены_плавки = QComboBox(self)
        self.Четвертый_участник_смены_плавки = QComboBox(self)
        
        # Добавляем участников в комбобоксы
        participants = [
            "Белков", "Карасев", "Ермаков", "Рабинович",
            "Валиулин", "Волков", "Семенов", "Левин",
            "Исмаилов", "Беляев", "Политов", "Кокшин",
            "Терентьев", "отсутствует"
        ]
        participants.sort()
        
        for combo in [self.Старший_смены_плавки, self.Первый_участник_смены_плавки,
                     self.Второй_участник_смены_плавки, self.Третий_участник_смены_плавки,
                     self.Четвертый_участник_смены_плавки]:
            combo.addItems(participants)
            combo.setCurrentIndex(-1)
        
        # Создаем остальные поля
        self.Наименование_отливки = QComboBox(self)
        self.Наименование_отливки.addItems([
            "Вороток", "Ригель", "Ригель optima", "Блок-картер", "Колесо РИТМ",
            "Накладка резьб", "Блок цилиндров", "Диагональ optima", "Кольцо"
        ])
        self.Наименование_отливки.setCurrentIndex(-1)
        
        self.Тип_эксперемента = QComboBox(self)
        self.Тип_эксперемента.addItems(["Бумага", "Волокно"])
        self.Тип_эксперемента.setCurrentIndex(-1)
        
        # Создаем поля для секторов опок
        self.Сектор_A_опоки = QLineEdit(self)
        self.Сектор_B_опоки = QLineEdit(self)
        self.Сектор_C_опоки = QLineEdit(self)
        self.Сектор_D_опоки = QLineEdit(self)
        
        # Создаем поля для временных параметров сектора A
        self.Плавка_время_прогрева_ковша_A = QLineEdit(self)
        self.Плавка_время_прогрева_ковша_A.setInputMask("23:59")
        self.Плавка_время_прогрева_ковша_A.setProperty("time", "true")
        self.Плавка_время_перемещения_A = QLineEdit(self)
        self.Плавка_время_перемещения_A.setInputMask("23:59")
        self.Плавка_время_перемещения_A.setProperty("time", "true")
        self.Плавка_время_заливки_A = QLineEdit(self)
        self.Плавка_время_заливки_A.setInputMask("23:59")
        self.Плавка_время_заливки_A.setProperty("time", "true")
        self.Плавка_температура_заливки_A = QLineEdit(self)
        self.Плавка_температура_заливки_A.setProperty("temperature", "true")

        # Создаем поля для временных параметров сектора B
        self.Плавка_время_прогрева_ковша_B = QLineEdit(self)
        self.Плавка_время_прогрева_ковша_B.setInputMask("23:59")
        self.Плавка_время_прогрева_ковша_B.setProperty("time", "true")
        self.Плавка_время_перемещения_B = QLineEdit(self)
        self.Плавка_время_перемещения_B.setInputMask("23:59")
        self.Плавка_время_перемещения_B.setProperty("time", "true")
        self.Плавка_время_заливки_B = QLineEdit(self)
        self.Плавка_время_заливки_B.setInputMask("23:59")
        self.Плавка_время_заливки_B.setProperty("time", "true")
        self.Плавка_температура_заливки_B = QLineEdit(self)
        self.Плавка_температура_заливки_B.setProperty("temperature", "true")

        # Создаем поля для временных параметров сектора C
        self.Плавка_время_прогрева_ковша_C = QLineEdit(self)
        self.Плавка_время_прогрева_ковша_C.setInputMask("23:59")
        self.Плавка_время_прогрева_ковша_C.setProperty("time", "true")
        self.Плавка_время_перемещения_C = QLineEdit(self)
        self.Плавка_время_перемещения_C.setInputMask("23:59")
        self.Плавка_время_перемещения_C.setProperty("time", "true")
        self.Плавка_время_заливки_C = QLineEdit(self)
        self.Плавка_время_заливки_C.setInputMask("23:59")
        self.Плавка_время_заливки_C.setProperty("time", "true")
        self.Плавка_температура_заливки_C = QLineEdit(self)
        self.Плавка_температура_заливки_C.setProperty("temperature", "true")

        # Создаем поля для временных параметров сектора D
        self.Плавка_время_прогрева_ковша_D = QLineEdit(self)
        self.Плавка_время_прогрева_ковша_D.setInputMask("23:59")
        self.Плавка_время_прогрева_ковша_D.setProperty("time", "true")
        self.Плавка_время_перемещения_D = QLineEdit(self)
        self.Плавка_время_перемещения_D.setInputMask("23:59")
        self.Плавка_время_перемещения_D.setProperty("time", "true")
        self.Плавка_время_заливки_D = QLineEdit(self)
        self.Плавка_время_заливки_D.setInputMask("23:59")
        self.Плавка_время_заливки_D.setProperty("time", "true")
        self.Плавка_температура_заливки_D = QLineEdit(self)
        self.Плавка_температура_заливки_D.setProperty("temperature", "true")

        # Создаем поле для комментария
        self.Комментарий = QTextEdit(self)
        self.Комментарий.setPlaceholderText("Введите комментарий...")
        
        # Создаем кнопки
        self.save_button = QPushButton("Сохранить", self)
        self.save_button.clicked.connect(self.save_data)
        
        self.search_button = QPushButton("Поиск", self)
        self.search_button.clicked.connect(self.show_search_dialog)
        
        # Добавляем обработчик изменения даты
        self.Плавка_дата.dateChanged.connect(self.generate_plavka_number)
        
        # Генерируем начальный номер плавки
        self.generate_plavka_number()

    def generate_plavka_number(self):
        try:
            current_month = self.Плавка_дата.date().month()
            
            if os.path.exists('plavka.xlsx'):
                df = pd.read_excel('plavka.xlsx')
                if not df.empty:
                    # Конвертируем даты в datetime
                    df['Плавка_дата'] = pd.to_datetime(df['Плавка_дата'], format='%d.%m.%Y')
                    
                    # Фильтруем записи только текущего месяца и года
                    current_year = self.Плавка_дата.date().year()
                    current_month_records = df[
                        (df['Плавка_дата'].dt.month == current_month) & 
                        (df['Плавка_дата'].dt.year == current_year)
                    ]
                    
                    if not current_month_records.empty:
                        # Ищем последний номер для текущего месяца
                        last_numbers = []
                        for num in current_month_records['Номер_плавки']:
                            try:
                                if isinstance(num, str) and '-' in num:
                                    month, number = num.split('-')
                                    if month == str(current_month):
                                        last_numbers.append(int(number))
                            except (ValueError, TypeError):
                                continue
                        
                        next_number = max(last_numbers) + 1 if last_numbers else 1
            else:
                next_number = 1
            
            # Форматируем номер плавки: месяц-номер(с ведущими нулями)
            new_plavka_number = f"{current_month}-{str(next_number).zfill(3)}"
            self.Номер_плавки.setText(new_plavka_number)
            
            # Обновляем учетный номер после генерации номера плавки
            self.update_uchet_number()
            
        except Exception as e:
            logging.error(f"Ошибка при генерации номера плавки: {str(e)}")
            self.Номер_плавки.setText("")

    def update_uchet_number(self):
        """Обновляет учетный номер на основе номера плавки"""
        try:
            plavka_number = self.Номер_плавки.text()
            if plavka_number:
                year = str(self.Плавка_дата.date().year())[-2:]  # Последние 2 цифры года
                uchet_number = f"{plavka_number}/{year}"
                return uchet_number
        except Exception as e:
            logging.error(f"Ошибка при обновлении учетного номера: {str(e)}")
        return None

    def generate_id(self, Плавка_дата, Номер_плавки):
        year = Плавка_дата.year()
        month = Плавка_дата.month()
        
        # Извлекаем число после дефиса
        match = re.search(r'-(\d+)', Номер_плавки)
        if match:
            number = match.group(1)
            # Преобразуем число в трехзначный формат (например, 1 -> 001)
            number = number.zfill(3)
            if len(number) <= 3:  # Проверяем, что число не превышает 999
                return f"{year}{month:02d}{number}"
            else:
                QMessageBox.warning(self, "Ошибка", "Номер плавки после дефиса не должен превышать 999.")
        else:
            QMessageBox.warning(self, "Ошибка", "Неверный формат номера плавки. Требуется формат с дефисом (например: xxx-123).")
        
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
        """Проверка существования ID в plavka.xlsx"""
        try:
            if not os.path.exists('plavka.xlsx'):
                return False
            
            wb = load_workbook('plavka.xlsx', read_only=True)
            ws = wb.active
            
            # Преобразуем проверяемый ID в строку для сравнения
            id_to_check = str(id_number).strip()
            
            # Проверяем первый столбец (ID)
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0] and str(row[0]).strip() == id_to_check:
                    wb.close()
                    return True
            
            wb.close()
            return False
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
                QMessageBox.warning(self, "Ошибка", "Недопустимая температура заливки")
                return False
        except ValueError:
            QMessageBox.warning(self, "Ошибка", "Температура должна быть числом")
            return False
        
        return True

    def format_temperature(self, temp_str):
        """Форматирование температур в нужный формат"""
        try:
            temp = float(temp_str)
            return f"{temp:.1f}°C"
        except ValueError:
            return temp_str

    def save_data(self):
        try:
            logging.info(f"Начало сохранения данных плавки {self.Номер_плавки.text()}")
            # Получаем ID из поля ввода
            id_number = self.generate_id(self.Плавка_дата.date(), self.Номер_плавки.text())
            
            # Проверяем, не пустой ли ID
            if not id_number:
                QMessageBox.warning(self, "Ошибка", "Введите ID плавки!")
                return
            
            # Проверяем на дубликат
            if self.check_duplicate_id(id_number):
                QMessageBox.warning(self, "Ошибка", 
                    f"Плавка с ID {id_number} уже существует в базе данных!")
                return
            
            # Если проверки пройдены, продолжаем сохранение
            Плавка_дата = self.Плавка_дата.date()
            formatted_date = Плавка_дата.toString("dd.MM.yyyy")
            Номер_плавки = self.Номер_плавки.text()

            Учетный_номер = self.update_uchet_number()
            if Учетный_номер is None:
                return

            Номер_кластера = self.Номер_кластера.text()
            Старший_смены_плавки = self.Старший_смены_плавки.currentText()  # Изменено на currentText()
            Первый_участник_смены_плавки = self.Первый_участник_смены_плавки.currentText()  # Изменено на currentText()
            Второй_участник_смены_плавки = self.Второй_участник_смены_плавки.currentText()  # Изменено на currentText()
            Третий_участник_смены_плавки = self.Третий_участник_смены_плавки.currentText()  # Изменено на currentText()
            Четвертый_участник_смены_плавки = self.Четвертый_участник_смены_плавки.currentText()  # Изменено на currentText()
            Наименование_отливки = self.Наименование_отливки.currentText()  # Изменено на currentText()
            Тип_эксперемента = self.Тип_эксперемента.currentText()  # Изменено на currentText()
            Сектор_A_опоки = self.Сектор_A_опоки.text()
            Сектор_B_опоки = self.Сектор_B_опоки.text()
            Сектор_C_опоки = self.Сектор_C_опоки.text()
            Сектор_D_опоки = self.Сектор_D_опоки.text()

            Плавка_время_прогрева_ковша_A = self.Плавка_время_прогрева_ковша_A.text()
            Плавка_время_перемещения_A = self.Плавка_время_перемещения_A.text()
            Плавка_время_заливки_A = self.Плавка_время_заливки_A.text()

            Плавка_время_прогрева_ковша_B = self.Плавка_время_прогрева_ковша_B.text()
            Плавка_время_перемещения_B = self.Плавка_время_перемещения_B.text()
            Плавка_время_заливки_B = self.Плавка_время_заливки_B.text()

            Плавка_время_прогрева_ковша_C = self.Плавка_время_прогрева_ковша_C.text()
            Плавка_время_перемещения_C = self.Плавка_время_перемещения_C.text()
            Плавка_время_заливки_C = self.Плавка_время_заливки_C.text()

            Плавка_время_прогрева_ковша_D = self.Плавка_время_прогрева_ковша_D.text()
            Плавка_время_перемещения_D = self.Плавка_время_перемещения_D.text()
            Плавка_время_заливки_D = self.Плавка_время_заливки_D.text()

            Плавка_температура_заливки_A = self.Плавка_температура_заливки_A.text()
            Плавка_температура_заливки_B = self.Плавка_температура_заливки_B.text()
            Плавка_температура_заливки_C = self.Плавка_температура_заливки_C.text()
            Плавка_температура_заливки_D = self.Плавка_температура_заливки_D.text()

            # Валидация времени
            if not (self.validate_time(Плавка_время_заливки_A) and self.validate_time(Плавка_время_прогрева_ковша_A) and self.validate_time(Плавка_время_перемещения_A) and
                    self.validate_time(Плавка_время_заливки_B) and self.validate_time(Плавка_время_прогрева_ковша_B) and self.validate_time(Плавка_время_перемещения_B) and
                    self.validate_time(Плавка_время_заливки_C) and self.validate_time(Плавка_время_прогрева_ковша_C) and self.validate_time(Плавка_время_перемещения_C) and
                    self.validate_time(Плавка_время_заливки_D) and self.validate_time(Плавка_время_прогрева_ковша_D) and self.validate_time(Плавка_время_перемещения_D)):
                QMessageBox.warning(self, "Ошибка", "Некорректный ввод времени. Используйте формат ЧЧ:ММ.")
                return

            Комментарий = self.Комментарий.toPlainText()

            save_to_excel(id_number, Учетный_номер, formatted_date, Номер_плавки, Номер_кластера,
                           Старший_смены_плавки, Первый_участник_смены_плавки,
                           Второй_участник_смены_плавки, Третий_участник_смены_плавки,
                           Четвертый_участник_смены_плавки, Наименование_отливки,
                           Тип_эксперемента, Сектор_A_опоки, Сектор_B_опоки,
                           Сектор_C_опоки, Сектор_D_опоки, 
                           Плавка_время_прогрева_ковша_A, Плавка_время_перемещения_A, Плавка_время_заливки_A, Плавка_температура_заливки_A,
                           Плавка_время_прогрева_ковша_B, Плавка_время_перемещения_B, Плавка_время_заливки_B, Плавка_температура_заливки_B,
                           Плавка_время_прогрева_ковша_C, Плавка_время_перемещения_C, Плавка_время_заливки_C, Плавка_температура_заливки_C,
                           Плавка_время_прогрева_ковша_D, Плавка_время_перемещения_D, Плавка_время_заливки_D, Плавка_температура_заливки_D,
                           Комментарий)

            QMessageBox.information(self, "Успех", "Данные сохранены в Excel!")

            # Очистка полей ввода
            self.clear_fields()
            logging.info("Данные успешно сохранены")
            
            # После успешного сохранения генерируем новый номер
            self.generate_plavka_number()
            
        except Exception as e:
            logging.error(f"Ошибка при сохранении данных: {str(e)}")
            QMessageBox.critical(self, "Ошибка", str(e))

    def clear_fields(self):
        self.Плавка_дата.setDate(QDate.currentDate().addDays(-1))
        self.Номер_плавки.clear()
        self.Номер_кластера.clear()
        self.Старший_смены_плавки.setCurrentIndex(-1)  # Сброс выбора
        self.Первый_участник_смены_плавки.setCurrentIndex(-1)  # Сброс выбора
        self.Второй_участник_смены_плавки.setCurrentIndex(-1)  # Сброс выбора
        self.Третий_участник_смены_плавки.setCurrentIndex(-1)  # Сброс выбора
        self.Четвертый_участник_смены_плавки.setCurrentIndex(-1)  # Сброс выбора
        self.Наименование_отливки.setCurrentIndex(-1)  # Сброс выбора
        self.Тип_эксперемента.setCurrentIndex(-1)  # Сброс выбора
        self.Сектор_A_опоки.clear()
        self.Сектор_B_опоки.clear()
        self.Сектор_C_опоки.clear()
        self.Сектор_D_опоки.clear()
        self.Плавка_время_прогрева_ковша_A.clear()
        self.Плавка_время_перемещения_A.clear()
        self.Плавка_время_заливки_A.clear()
        self.Плавка_температура_заливки_A.clear()
        self.Плавка_время_прогрева_ковша_B.clear()
        self.Плавка_время_перемещения_B.clear()
        self.Плавка_время_заливки_B.clear()
        self.Плавка_температура_заливки_B.clear()
        self.Плавка_время_прогрева_ковша_C.clear()
        self.Плавка_время_перемещения_C.clear()
        self.Плавка_время_заливки_C.clear()
        self.Плавка_температура_заливки_C.clear()
        self.Плавка_время_прогрева_ковша_D.clear()
        self.Плавка_время_перемещения_D.clear()
        self.Плавка_время_заливки_D.clear()
        self.Плавка_температура_заливки_D.clear()
        self.Комментарий.clear()

    def show_search_dialog(self):
        dialog = SearchDialog(self)
        dialog.exec_()

class StatisticsWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        
    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        # Создаем таблицу для отображения данных
        self.data_table = QTableWidget()
        self.data_table.setColumnCount(2)
        self.data_table.setHorizontalHeaderLabels(["Дата", "Температура"])
        layout.addWidget(self.data_table)
        
        # Кнопки для разных типов отображения
        buttons_layout = QHBoxLayout()
        
        temp_button = QPushButton("Температуры")
        temp_button.clicked.connect(lambda: self.show_data('temperature'))
        
        casting_button = QPushButton("Статистика отливок")
        casting_button.clicked.connect(lambda: self.show_data('castings'))
        
        time_button = QPushButton("Временной анализ")
        time_button.clicked.connect(lambda: self.show_data('time'))
        
        buttons_layout.addWidget(temp_button)
        buttons_layout.addWidget(casting_button)
        buttons_layout.addWidget(time_button)
        
        layout.addLayout(buttons_layout)
    
    def show_data(self, data_type):
        self.data_table.setRowCount(0)
        
        try:
            wb = load_workbook(EXCEL_FILENAME, read_only=True)
            ws = wb.active
            headers = [cell.value for cell in ws[1]]
            
            if data_type == 'temperature':
                self._show_temperature(ws, headers)
            elif data_type == 'castings':
                self._show_castings(ws, headers)
            elif data_type == 'time':
                self._show_time_analysis(ws, headers)
                
            wb.close()
            
        except Exception as e:
            logging.error(f"Ошибка при отображении данных: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Ошибка при отображении данных: {str(e)}")
    
    def _show_temperature(self, ws, headers):
        self.data_table.setHorizontalHeaderLabels(["Дата", "Температура"])
        
        for row in ws.iter_rows(min_row=2, values_only=True):
            data = dict(zip(headers, row))
            try:
                temp_A = float(data['Плавка_температура_заливки_A'])
                temp_B = float(data['Плавка_температура_заливки_B'])
                temp_C = float(data['Плавка_температура_заливки_C'])
                temp_D = float(data['Плавка_температура_заливки_D'])
                date = data['Плавка_дата']
                
                row_position = self.data_table.rowCount()
                self.data_table.insertRow(row_position)
                
                self.data_table.setItem(row_position, 0, QTableWidgetItem(date))
                self.data_table.setItem(row_position, 1, QTableWidgetItem(f"{temp_A}°C"))
                
                row_position = self.data_table.rowCount()
                self.data_table.insertRow(row_position)
                
                self.data_table.setItem(row_position, 0, QTableWidgetItem(date))
                self.data_table.setItem(row_position, 1, QTableWidgetItem(f"{temp_B}°C"))
                
                row_position = self.data_table.rowCount()
                self.data_table.insertRow(row_position)
                
                self.data_table.setItem(row_position, 0, QTableWidgetItem(date))
                self.data_table.setItem(row_position, 1, QTableWidgetItem(f"{temp_C}°C"))
                
                row_position = self.data_table.rowCount()
                self.data_table.insertRow(row_position)
                
                self.data_table.setItem(row_position, 0, QTableWidgetItem(date))
                self.data_table.setItem(row_position, 1, QTableWidgetItem(f"{temp_D}°C"))
                
            except (ValueError, TypeError):
                continue
        
        self.data_table.resizeColumnsToContents()

class SearchDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
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
        
    def setup_ui(self):
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
        self.filter_casting.addItems(["Все"] + [
            "Вороток", "Ригель", "Ригель optima", "Блок-картер",
            "Накладка резьб", "Блок цилиндров", "Диагональ optima"
        ])
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
        
        # Существующие виджеты поиска
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Введите текст для поиска...")
        search_layout.addWidget(self.search_input)
        
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(len(SEARCH_FIELDS))
        self.results_table.setHorizontalHeaderLabels(SEARCH_FIELDS)
        search_layout.addWidget(self.results_table)
        
        self.tab_widget.addTab(search_tab, "Результаты поиска")
        
        # Вкладка статистики
        stats_tab = QWidget()
        stats_layout = QVBoxLayout(stats_tab)
        
        self.stats_text = QTextEdit()
        self.stats_text.setReadOnly(True)
        stats_layout.addWidget(self.stats_text)
        
        self.tab_widget.addTab(stats_tab, "Статистика")
        
        # Добавляем вкладку визуализации
        viz_tab = StatisticsWidget()
        self.tab_widget.addTab(viz_tab, "Визуализация")
        
        layout.addWidget(self.tab_widget)
        
        # Кнопки
        button_layout = QHBoxLayout()
        self.search_button = QPushButton("Поиск")
        self.edit_button = QPushButton("Редактировать")
        self.export_button = QPushButton("Экспорт")
        self.stats_button = QPushButton("Обновить статистику")
        self.backup_button = QPushButton("Создать резервную копию")
        
        button_layout.addWidget(self.search_button)
        button_layout.addWidget(self.edit_button)
        button_layout.addWidget(self.export_button)
        button_layout.addWidget(self.stats_button)
        button_layout.addWidget(self.backup_button)
        layout.addLayout(button_layout)
        
        # Подключаем обработчики
        self.search_button.clicked.connect(self.search_records)
        self.edit_button.clicked.connect(self.edit_selected)
        self.export_button.clicked.connect(self.export_results)
        self.stats_button.clicked.connect(self.update_statistics)
        self.backup_button.clicked.connect(self.create_backup)

    def apply_filters(self, row, headers):
        """Применяет фильтры к записи"""
        try:
            data = dict(zip(headers, row))
            
            # Фильтр по дате
            record_date = QDate.fromString(data['Плавка_дата'], "dd.MM.yyyy")
            if not (self.date_from.date() <= record_date <= self.date_to.date()):
                return False
            
            # Фильтр по типу отливки
            if self.filter_casting.currentText() != "Все" and \
               data['Наименование_отливки'] != self.filter_casting.currentText():
                return False
            
            # Фильтр по температуре
            if self.temp_from.text() and self.temp_to.text():
                try:
                    temp_A = float(data['Плавка_температура_заливки_A'])
                    temp_B = float(data['Плавка_температура_заливки_B'])
                    temp_C = float(data['Плавка_температура_заливки_C'])
                    temp_D = float(data['Плавка_температура_заливки_D'])
                    temp_from = float(self.temp_from.text())
                    temp_to = float(self.temp_to.text())
                    if not (temp_from <= temp_A <= temp_to) and not (temp_from <= temp_B <= temp_to) and not (temp_from <= temp_C <= temp_to) and not (temp_from <= temp_D <= temp_to):
                        return False
                except ValueError:
                    pass
            
            return True
        except Exception as e:
            logging.error(f"Ошибка при применении фильтров: {str(e)}")
            return False

    def update_statistics(self):
        """Обновляет статистику по данным"""
        try:
            wb = load_workbook(EXCEL_FILENAME, read_only=True)
            ws = wb.active
            headers = [cell.value for cell in ws[1]]
            
            stats = {
                'total_records': 0,
                'avg_temp': [],
                'casting_types': {},
                'participants': set(),
                'min_temp': float('inf'),
                'max_temp': float('-inf')
            }
            
            for row in ws.iter_rows(min_row=2, values_only=True):
                # Проверяем фильтры
                if not self.apply_filters(row, headers):
                    continue
                    
                # Ищем совпадения
                data = dict(zip(headers, row))
                stats['total_records'] += 1
                
                # Температура
                try:
                    temp_A = float(data['Плавка_температура_заливки_A'])
                    temp_B = float(data['Плавка_температура_заливки_B'])
                    temp_C = float(data['Плавка_температура_заливки_C'])
                    temp_D = float(data['Плавка_температура_заливки_D'])
                    stats['avg_temp'].append(temp_A)
                    stats['avg_temp'].append(temp_B)
                    stats['avg_temp'].append(temp_C)
                    stats['avg_temp'].append(temp_D)
                    stats['min_temp'] = min(stats['min_temp'], temp_A, temp_B, temp_C, temp_D)
                    stats['max_temp'] = max(stats['max_temp'], temp_A, temp_B, temp_C, temp_D)
                except (ValueError, TypeError):
                    pass
                
                # Типы отливок
                casting = data['Наименование_отливки']
                stats['casting_types'][casting] = stats['casting_types'].get(casting, 0) + 1
                
                # Участники
                stats['participants'].add(data['Старший_смены_плавки'])
                stats['participants'].add(data['Первый_участник_смены_плавки'])
                stats['participants'].add(data['Второй_участник_смены_плавки'])
                stats['participants'].add(data['Третий_участник_смены_плавки'])
                stats['participants'].add(data['Четвертый_участник_смены_плавки'])
            
            # Формируем отчет
            report = [
                "=== Общая статистика ===",
                f"Всего записей: {stats['total_records']}",
                f"Количество участников: {len(stats['participants'])}",
                "",
                "=== Температура заливки ===",
                f"Средняя: {sum(stats['avg_temp'])/len(stats['avg_temp']):.1f}°C" if stats['avg_temp'] else "Нет данных",
                f"Минимальная: {stats['min_temp']}°C" if stats['min_temp'] != float('inf') else "Нет данных",
                f"Максимальная: {stats['max_temp']}°C" if stats['max_temp'] != float('-inf') else "Нет данных",
                "",
                "=== Распределение по типам отливок ===",
            ]
            
            for casting, count in sorted(stats['casting_types'].items()):
                report.append(f"{casting}: {count} ({count/stats['total_records']*100:.1f}%)")
            
            self.stats_text.setText("\n".join(report))
            wb.close()
            
        except Exception as e:
            logging.error(f"Ошибка при обновлении статистики: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Ошибка при обновлении статистики: {str(e)}")

    def search_records(self):
        search_text = self.search_input.text().lower()
        try:
            wb = load_workbook(EXCEL_FILENAME, read_only=True)
            ws = wb.active
            headers = [cell.value for cell in ws[1]]
            
            self.results_table.setRowCount(0)
            
            for row in ws.iter_rows(min_row=2, values_only=True):
                # Проверяем фильтры
                if not self.apply_filters(row, headers):
                    continue
                    
                # Ищем совпадения
                for cell in row:
                    if cell and str(cell).lower().find(search_text) != -1:
                        row_position = self.results_table.rowCount()
                        self.results_table.insertRow(row_position)
                        
                        for col, field in enumerate(SEARCH_FIELDS):
                            field_index = headers.index(field)
                            self.results_table.setItem(
                                row_position, col, 
                                QTableWidgetItem(str(row[field_index])))
                        break
            
            wb.close()
            
        except Exception as e:
            logging.error(f"Ошибка при поиске: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Ошибка при поиске: {str(e)}")

    def edit_selected(self):
        current_row = self.results_table.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, "Предупреждение", "Выберите запись для редактирования")
            return
            
        # Получаем ID выбранной записи
        record_id = self.results_table.item(current_row, 0).text()
        
        # Создаем диалог редактирования
        edit_dialog = EditRecordDialog(record_id, self)
        if edit_dialog.exec_() == QDialog.Accepted:
            # Обновляем таблицу после редактирования
            self.search_records()

    def export_results(self):
        try:
            format_str = "Excel files (*.xlsx);;CSV files (*.csv);;PDF files (*.pdf)"
            file_name, selected_format = QFileDialog.getSaveFileName(
                self, "Экспорт данных", "", format_str
            )
            
            if file_name:
                # Создаем DataFrame из данных таблицы
                data = []
                for row in range(self.results_table.rowCount()):
                    row_data = []
                    for col in range(self.results_table.columnCount()):
                        item = self.results_table.item(row, col)
                        row_data.append(item.text() if item else "")
                    data.append(row_data)
                
                df = pd.DataFrame(data, columns=SEARCH_FIELDS)
                
                if "xlsx" in selected_format:
                    df.to_excel(file_name, index=False)
                elif "csv" in selected_format:
                    df.to_csv(file_name, index=False)
                elif "pdf" in selected_format:
                    # Для PDF потребуется дополнительная настройка
                    df.to_html(file_name.replace('.pdf', '.html'))
                    # Конвертация HTML в PDF (требуется дополнительная библиотека)
                
                QMessageBox.information(self, "Успех", "Данные успешно экспортированы")
                
        except Exception as e:
            logging.error(f"Ошибка при экспорте: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Ошибка при экспорте: {str(e)}")

    def create_backup(self):
        try:
            # Создаем директорию для резервных копий если её нет
            if not os.path.exists(BACKUP_DIR):
                os.makedirs(BACKUP_DIR)
            
            # Формируем имя файла резервной копии
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_file = os.path.join(BACKUP_DIR, f'plavka_backup_{timestamp}.xlsx')
            
            # Копируем файл
            shutil.copy2(EXCEL_FILENAME, backup_file)
            
            # Удаляем старые резервные копии если их больше MAX_BACKUPS
            backups = sorted([os.path.join(BACKUP_DIR, f) for f in os.listdir(BACKUP_DIR)])
            while len(backups) > MAX_BACKUPS:
                os.remove(backups[0])
                backups.pop(0)
            
            QMessageBox.information(self, "Успех", 
                f"Резервная копия создана:\n{backup_file}")
            
        except Exception as e:
            logging.error(f"Ошибка при создании резервной копии: {str(e)}")
            QMessageBox.critical(self, "Ошибка", 
                f"Ошибка при создании резервной копии: {str(e)}")

class EditRecordDialog(QDialog):
    def __init__(self, record_id, parent=None):
        super().__init__(parent)
        self.record_id = record_id
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

        # Создаем поля для секторов опоки
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

    def load_record_data(self):
        try:
            wb = load_workbook(EXCEL_FILENAME)
            ws = wb.active
            
            headers = [cell.value for cell in ws[1]]
            
            for row in ws.iter_rows(min_row=2, values_only=True):
                if str(row[0]) == self.record_id:
                    # Заполняем поля данными
                    self.fill_fields(row, headers)
                    break
                    
            wb.close()
            
        except Exception as e:
            logging.error(f"Ошибка при загрузке записи: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Ошибка при загрузке записи: {str(e)}")

    def fill_fields(self, row, headers):
        """Заполняет поля формы данными из записи"""
        try:
            # Создаем словарь с данными
            data = dict(zip(headers, row))
            
            # Заполняем поля
            self.Плавка_дата.setDate(QDate.fromString(data['Плавка_дата'], "dd.MM.yyyy"))
            self.Номер_плавки.setText(str(data['Номер_плавки']))
            self.Номер_кластера.setText(str(data['Номер_кластера']))
            
            # Устанавливаем значения комбобоксов
            self.Старший_смены_плавки.setCurrentText(str(data['Старший_смены_плавки']))
            self.Первый_участник_смены_плавки.setCurrentText(str(data['Первый_участник_смены_плавки']))
            self.Второй_участник_смены_плавки.setCurrentText(str(data['Второй_участник_смены_плавки']))
            self.Третий_участник_смены_плавки.setCurrentText(str(data['Третий_участник_смены_плавки']))
            self.Четвертый_участник_смены_плавки.setCurrentText(str(data['Четвертый_участник_смены_плавки']))
            
            self.Наименование_отливки.setCurrentText(str(data['Наименование_отливки']))
            self.Тип_эксперемента.setCurrentText(str(data['Тип_эксперемента']))
            
            # Заполняем секторы опоки
            self.Сектор_A_опоки.setText(str(data['Сектор_A_опоки']))
            self.Сектор_B_опоки.setText(str(data['Сектор_B_опоки']))
            self.Сектор_C_опоки.setText(str(data['Сектор_C_опоки']))
            self.Сектор_D_опоки.setText(str(data['Сектор_D_опоки']))
            
            # Заполняем время и температуру
            self.Плавка_время_прогрева_ковша_A.setText(str(data['Плавка_время_прогрева_ковша_A']))
            self.Плавка_время_перемещения_A.setText(str(data['Плавка_время_перемещения_A']))
            self.Плавка_время_заливки_A.setText(str(data['Плавка_время_заливки_A']))
            self.Плавка_температура_заливки_A.setText(str(data['Плавка_температура_заливки_A']))

            self.Плавка_время_прогрева_ковша_B.setText(str(data['Плавка_время_прогрева_ковша_B']))
            self.Плавка_время_перемещения_B.setText(str(data['Плавка_время_перемещения_B']))
            self.Плавка_время_заливки_B.setText(str(data['Плавка_время_заливки_B']))
            self.Плавка_температура_заливки_B.setText(str(data['Плавка_температура_заливки_B']))

            self.Плавка_время_прогрева_ковша_C.setText(str(data['Плавка_время_прогрева_ковша_C']))
            self.Плавка_время_перемещения_C.setText(str(data['Плавка_время_перемещения_C']))
            self.Плавка_время_заливки_C.setText(str(data['Плавка_время_заливки_C']))
            self.Плавка_температура_заливки_C.setText(str(data['Плавка_температура_заливки_C']))

            self.Плавка_время_прогрева_ковша_D.setText(str(data['Плавка_время_прогрева_ковша_D']))
            self.Плавка_время_перемещения_D.setText(str(data['Плавка_время_перемещения_D']))
            self.Плавка_время_заливки_D.setText(str(data['Плавка_время_заливки_D']))
            self.Плавка_температура_заливки_D.setText(str(data['Плавка_температура_заливки_D']))

            self.Комментарий.setText(str(data['Комментарий']))
            
        except Exception as e:
            logging.error(f"Ошибка при заполнении полей: {str(e)}")
            raise

    def save_changes(self):
        """Сохраняет изменения в Excel файл"""
        try:
            wb = load_workbook(EXCEL_FILENAME)
            ws = wb.active
            
            # Находим строку с нужным ID
            row_index = None
            for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True)):
                if str(row[0]) == self.record_id:
                    row_index = idx + 2
                    break
            
            if row_index:
                # Обновляем данные в строке
                ws.cell(row=row_index, column=3).value = self.Плавка_дата.date().toString("dd.MM.yyyy")
                ws.cell(row=row_index, column=4).value = self.Номер_плавки.text()
                ws.cell(row=row_index, column=5).value = self.Номер_кластера.text()
                ws.cell(row=row_index, column=6).value = self.Старший_смены_плавки.currentText()
                ws.cell(row=row_index, column=7).value = self.Первый_участник_смены_плавки.currentText()
                ws.cell(row=row_index, column=8).value = self.Второй_участник_смены_плавки.currentText()
                ws.cell(row=row_index, column=9).value = self.Третий_участник_смены_плавки.currentText()
                ws.cell(row=row_index, column=10).value = self.Четвертый_участник_смены_плавки.currentText()
                ws.cell(row=row_index, column=11).value = self.Наименование_отливки.currentText()
                ws.cell(row=row_index, column=12).value = self.Тип_эксперемента.currentText()
                ws.cell(row=row_index, column=13).value = self.Сектор_A_опоки.text()
                ws.cell(row=row_index, column=14).value = self.Сектор_B_опоки.text()
                ws.cell(row=row_index, column=15).value = self.Сектор_C_опоки.text()
                ws.cell(row=row_index, column=16).value = self.Сектор_D_опоки.text()
                ws.cell(row=row_index, column=17).value = self.Плавка_время_прогрева_ковша_A.text()
                ws.cell(row=row_index, column=18).value = self.Плавка_время_перемещения_A.text()
                ws.cell(row=row_index, column=19).value = self.Плавка_время_заливки_A.text()
                ws.cell(row=row_index, column=20).value = self.Плавка_температура_заливки_A.text()
                ws.cell(row=row_index, column=21).value = self.Плавка_время_прогрева_ковша_B.text()
                ws.cell(row=row_index, column=22).value = self.Плавка_время_перемещения_B.text()
                ws.cell(row=row_index, column=23).value = self.Плавка_время_заливки_B.text()
                ws.cell(row=row_index, column=24).value = self.Плавка_температура_заливки_B.text()
                ws.cell(row=row_index, column=25).value = self.Плавка_время_прогрева_ковша_C.text()
                ws.cell(row=row_index, column=26).value = self.Плавка_время_перемещения_C.text()
                ws.cell(row=row_index, column=27).value = self.Плавка_время_заливки_C.text()
                ws.cell(row=row_index, column=28).value = self.Плавка_температура_заливки_C.text()
                ws.cell(row=row_index, column=29).value = self.Плавка_время_прогрева_ковша_D.text()
                ws.cell(row=row_index, column=30).value = self.Плавка_время_перемещения_D.text()
                ws.cell(row=row_index, column=31).value = self.Плавка_время_заливки_D.text()
                ws.cell(row=row_index, column=32).value = self.Плавка_температура_заливки_D.text()
                ws.cell(row=row_index, column=33).value = self.Комментарий.toPlainText()
                
                wb.save(EXCEL_FILENAME)
                QMessageBox.information(self, "Успех", "Изменения сохранены")
                self.accept()
            
        except Exception as e:
            logging.error(f"Ошибка при сохранении изменений: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Ошибка при сохранении изменений: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
