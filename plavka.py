import sys
import os
import re
import logging
from datetime import datetime
from typing import Optional, Dict, Any, List
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLineEdit,
    QPushButton, QMessageBox, QLabel, QFrame,
    QDateEdit, QComboBox, QTableWidget, QTableWidgetItem,
    QHBoxLayout, QDialog, QFileDialog, QGroupBox, QGridLayout,
    QTabWidget, QTextEdit, QScrollArea
)
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QColor, QShortcut, QKeySequence
from PySide6.QtWidgets import QGraphicsDropShadowEffect
from database import Database
from models import MeltRecord, SectorData
from constants import (
    TIME_FORMAT, DATE_FORMAT, TEMPERATURE_RANGE, MAX_PARTICIPANTS,
    PARTICIPANTS, CASTING_NAMES, ExperimentType, SectorName
)

# Настройка логирования
logging.basicConfig(
    filename='plavka.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        logging.info("Инициализация главного окна")
        try:
            self.db = Database()
            logging.info("База данных инициализирована")
        except Exception as e:
            logging.error(f"Ошибка при инициализации базы данных: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Не удалось инициализировать базу данных: {str(e)}")
        self.init_ui()

    def init_ui(self):
        """Инициализация пользовательского интерфейса"""
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
        
        # Подключаем сигнал изменения даты после создания виджетов
        self.Плавка_дата.dateChanged.connect(self.generate_plavka_number)
        
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
        basic_grid.addWidget(QLabel("Учетный номер:"), 3, 0)
        basic_grid.addWidget(self.Учетный_номер, 3, 1)
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
        casting_grid.addWidget(self.Тип_эксперимента, 1, 1, 1, 3)
        
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
        
        self.Номер_плавки = QLineEdit(self)
        self.Номер_плавки.setReadOnly(True)
        
        # Добавляем поле учетного номера
        self.Учетный_номер = QLineEdit(self)
        self.Учетный_номер.setReadOnly(True)
        
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
        
        self.Тип_эксперимента = QComboBox(self)
        self.Тип_эксперимента.addItems(["Бумага", "Волокно"])
        self.Тип_эксперимента.setCurrentIndex(-1)
        
        # Создаем поля для секторов опок
        self.Сектор_A_опоки = QLineEdit(self)
        self.Сектор_B_опоки = QLineEdit(self)
        self.Сектор_C_опоки = QLineEdit(self)
        self.Сектор_D_опоки = QLineEdit(self)
        
        # Создаем поля для временных параметров сектора A
        self.Плавка_время_прогрева_ковша_A = QLineEdit(self)
        self.Плавка_время_прогрева_ковша_A.setInputMask("99:99")
        self.Плавка_время_прогрева_ковша_A.setProperty("time", "true")
        self.Плавка_время_перемещения_A = QLineEdit(self)
        self.Плавка_время_перемещения_A.setInputMask("99:99")
        self.Плавка_время_перемещения_A.setProperty("time", "true")
        self.Плавка_время_заливки_A = QLineEdit(self)
        self.Плавка_время_заливки_A.setInputMask("99:99")
        self.Плавка_время_заливки_A.setProperty("time", "true")
        self.Плавка_температура_заливки_A = QLineEdit(self)
        self.Плавка_температура_заливки_A.setProperty("temperature", "true")

        # Создаем поля для временных параметров сектора B
        self.Плавка_время_прогрева_ковша_B = QLineEdit(self)
        self.Плавка_время_прогрева_ковша_B.setInputMask("99:99")
        self.Плавка_время_прогрева_ковша_B.setProperty("time", "true")
        self.Плавка_время_перемещения_B = QLineEdit(self)
        self.Плавка_время_перемещения_B.setInputMask("99:99")
        self.Плавка_время_перемещения_B.setProperty("time", "true")
        self.Плавка_время_заливки_B = QLineEdit(self)
        self.Плавка_время_заливки_B.setInputMask("99:99")
        self.Плавка_время_заливки_B.setProperty("time", "true")
        self.Плавка_температура_заливки_B = QLineEdit(self)
        self.Плавка_температура_заливки_B.setProperty("temperature", "true")

        # Создаем поля для временных параметров сектора C
        self.Плавка_время_прогрева_ковша_C = QLineEdit(self)
        self.Плавка_время_прогрева_ковша_C.setInputMask("99:99")
        self.Плавка_время_прогрева_ковша_C.setProperty("time", "true")
        self.Плавка_время_перемещения_C = QLineEdit(self)
        self.Плавка_время_перемещения_C.setInputMask("99:99")
        self.Плавка_время_перемещения_C.setProperty("time", "true")
        self.Плавка_время_заливки_C = QLineEdit(self)
        self.Плавка_время_заливки_C.setInputMask("99:99")
        self.Плавка_время_заливки_C.setProperty("time", "true")
        self.Плавка_температура_заливки_C = QLineEdit(self)
        self.Плавка_температура_заливки_C.setProperty("temperature", "true")

        # Создаем поля для временных параметров сектора D
        self.Плавка_время_прогрева_ковша_D = QLineEdit(self)
        self.Плавка_время_прогрева_ковша_D.setInputMask("99:99")
        self.Плавка_время_прогрева_ковша_D.setProperty("time", "true")
        self.Плавка_время_перемещения_D = QLineEdit(self)
        self.Плавка_время_перемещения_D.setInputMask("99:99")
        self.Плавка_время_перемещения_D.setProperty("time", "true")
        self.Плавка_время_заливки_D = QLineEdit(self)
        self.Плавка_время_заливки_D.setInputMask("99:99")
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
        
        # Устанавливаем дату после создания всех виджетов
        self.Плавка_дата.setDate(QDate.currentDate())
        
        # Генерируем номер плавки
        self.generate_plavka_number()

    def generate_plavka_number(self):
        """Генерирует номер плавки в формате месяц-номер"""
        try:
            current_date = self.Плавка_дата.date()
            current_month = current_date.month()
            logging.info(f"Генерация номера плавки для месяца: {current_month}, дата: {current_date.toString('dd.MM.yyyy')}")
            
            next_number = 1  # По умолчанию начинаем с 1
            
            if os.path.exists('plavka.db'):
                try:
                    records = self.db.get_records()  # Записи уже отсортированы по дате в порядке убывания
                    logging.info(f"Получено записей из БД: {len(records)}")
                    
                    # Получаем все номера плавок текущего месяца
                    last_numbers = set()  # Используем множество для исключения дубликатов
                    for record in records:
                        if not record or not record.date or not record.plavka_number:
                            continue
                            
                        logging.info(f"Обработка записи: дата={record.date.strftime('%d.%m.%Y')}, номер={record.plavka_number}")
                        
                        try:
                            month, number = record.plavka_number.split('-')
                            record_month = int(month)
                            
                            if record_month == current_month:
                                try:
                                    num = int(number)
                                    last_numbers.add(num)  # Добавляем в множество
                                    logging.info(f"Добавлен номер {num} в список для месяца {current_month}")
                                except ValueError:
                                    logging.warning(f"Некорректный номер в номере плавки: {number}")
                            else:
                                logging.info(f"Пропуск номера - не текущий месяц ({record_month} != {current_month})")
                        except (ValueError, AttributeError) as e:
                            logging.warning(f"Некорректный формат номера плавки в БД: {record.plavka_number}. Ошибка: {str(e)}")
                            continue
                    
                    if last_numbers:
                        next_number = max(last_numbers) + 1
                        logging.info(f"Найдены уникальные номера для месяца {current_month}: {sorted(last_numbers)}")
                        logging.info(f"Максимальный номер: {max(last_numbers)}, следующий будет {next_number}")
                    else:
                        logging.info(f"Не найдено номеров плавок для месяца {current_month}, начинаем с 1")
                except Exception as db_error:
                    logging.error(f"Ошибка при чтении из базы данных: {str(db_error)}")
                    
            # Форматируем номер плавки: месяц-номер(с ведущими нулями)
            new_plavka_number = f"{current_month}-{str(next_number).zfill(3)}"
            logging.info(f"Сгенерирован новый номер плавки: {new_plavka_number}")
            
            # Проверяем, что номер плавки действительно изменился
            old_number = self.Номер_плавки.text()
            if old_number != new_plavka_number:
                logging.info(f"Номер плавки изменился с {old_number} на {new_plavka_number}")
            else:
                logging.info(f"Номер плавки не изменился: {old_number}")
                
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
        """
        Генерирует учетный номер в формате номер_плавки/год
        Например: 2-157/25
        """
        try:
            # Получаем последние две цифры года
            last_two_digits_year = str(Плавка_дата.year())[-2:]
            
            # Проверяем, что номер плавки не пустой и имеет правильный формат
            if not Номер_плавки:
                logging.warning("Номер плавки не указан")
                QMessageBox.warning(self, "Ошибка", "Номер плавки не указан")
                return None
            
            номер_плавки = Номер_плавки.strip()
            if not re.match(r'^\d+-\d+$', номер_плавки):
                logging.warning(f"Неверный формат номера плавки: {номер_плавки}")
                QMessageBox.warning(self, "Ошибка", "Неверный формат номера плавки")
                return None
            
            return f"{номер_плавки}/{last_two_digits_year}"
            
        except Exception as e:
            logging.error(f"Ошибка при генерации учетного номера: {str(e)}")
            QMessageBox.warning(self, "Ошибка", "Не удалось сгенерировать учетный номер")
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
                QMessageBox.warning(self, "Ошибка", "Недопустимая температура")
                return False
                
        except ValueError:
            QMessageBox.warning(self, "Ошибка", "Недопустимая температура")
            return False
        
        return True

    def save_data(self):
        """Сохраняет данные формы в базу данных"""
        logging.info("Попытка сохранения данных")
        
        try:
            # Проверяем валидность данных перед сохранением
            if not self.validate_temperatures():
                logging.warning("Валидация температур не пройдена")
                return
                
            # Создаем объект записи
            record = MeltRecord()
            
            # Основные данные плавки
            record.date = self.Плавка_дата.date().toPython()
            record.plavka_number = self.Номер_плавки.text()
            record.uchet_number = self.Учетный_номер.text()
            record.experiment_type = self.Тип_эксперимента.currentText()
            
            # Температуры заливки
            record.casting_temperature_a = float(self.Плавка_температура_заливки_A.text() or 0)
            record.casting_temperature_b = float(self.Плавка_температура_заливки_B.text() or 0)
            record.casting_temperature_c = float(self.Плавка_температура_заливки_C.text() or 0)
            record.casting_temperature_d = float(self.Плавка_температура_заливки_D.text() or 0)
            
            # Данные по секторам
            record.sector_data = {}
            for sector in SectorName:
                sector_data = SectorData()
                sector_data.casting_start = getattr(self, f"{sector.value}_время_начала_заливки").time().toString(TIME_FORMAT)
                sector_data.casting_end = getattr(self, f"{sector.value}_время_конца_заливки").time().toString(TIME_FORMAT)
                sector_data.casting_name = getattr(self, f"{sector.value}_заливщик").currentText()
                record.sector_data[sector.value] = sector_data
            
            # Сохраняем в базу данных
            self.db.save_record(record)
            logging.info(f"Запись успешно сохранена: {record.plavka_number}")
            
            QMessageBox.information(self, "Успех", "Данные успешно сохранены")
            
            # Очищаем форму после успешного сохранения
            self.clear_form()
            
        except Exception as e:
            logging.error(f"Ошибка при сохранении данных: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить данные: {str(e)}")

    def show_search_dialog(self):
        """Показывает диалог поиска записей"""
        logging.info("Открытие диалога поиска")
        try:
            records = self.db.get_records()
            if not records:
                QMessageBox.information(self, "Информация", "В базе данных нет записей")
                return
                
            dialog = QDialog(self)
            dialog.setWindowTitle("Поиск записей")
            dialog.setModal(True)
            
            layout = QVBoxLayout()
            
            # Создаем таблицу для отображения записей
            table = QTableWidget()
            table.setColumnCount(7)
            table.setHorizontalHeaderLabels([
                "Дата", "Номер плавки", "Учетный номер", 
                "Тип эксперимента", "Температура A", "Температура B",
                "Температура C", "Температура D"
            ])
            
            # Заполняем таблицу данными
            for record in records:
                row = table.rowCount()
                table.insertRow(row)
                
                date_str = record.date.strftime(DATE_FORMAT) if record.date else ""
                table.setItem(row, 0, QTableWidgetItem(date_str))
                table.setItem(row, 1, QTableWidgetItem(record.plavka_number))
                table.setItem(row, 2, QTableWidgetItem(record.uchet_number))
                table.setItem(row, 3, QTableWidgetItem(record.experiment_type))
                table.setItem(row, 4, QTableWidgetItem(str(record.casting_temperature_a)))
                table.setItem(row, 5, QTableWidgetItem(str(record.casting_temperature_b)))
                table.setItem(row, 6, QTableWidgetItem(str(record.casting_temperature_c)))
                table.setItem(row, 7, QTableWidgetItem(str(record.casting_temperature_d)))
            
            # Добавляем таблицу в диалог
            layout.addWidget(table)
            
            # Кнопка закрытия
            close_button = QPushButton("Закрыть")
            close_button.clicked.connect(dialog.close)
            layout.addWidget(close_button)
            
            dialog.setLayout(layout)
            dialog.resize(800, 600)
            dialog.exec()
            
        except Exception as e:
            logging.error(f"Ошибка при открытии диалога поиска: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Не удалось открыть диалог поиска: {str(e)}")

if __name__ == "__main__":
    try:
        logging.info("Запуск приложения")
        app = QApplication(sys.argv)
        window = MainWindow()
        window.show()
        logging.info("Главное окно отображено")
        sys.exit(app.exec())
    except Exception as e:
        logging.error(f"Критическая ошибка при запуске приложения: {str(e)}")
        print(f"Критическая ошибка: {str(e)}")
