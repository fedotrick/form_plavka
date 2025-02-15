import sys
import os
import re
import logging
from datetime import datetime
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLineEdit,
    QPushButton, QMessageBox, QLabel, QFrame,
    QDateEdit, QComboBox, QTableWidget, QTableWidgetItem,
    QHBoxLayout, QDialog, QFileDialog, QGroupBox, QGridLayout,
    QTabWidget, QTextEdit
)
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QColor
from PySide6.QtWidgets import QGraphicsDropShadowEffect
from database import Database

# Настройка логирования
logging.basicConfig(
    filename='plavka.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Константы
TIME_FORMAT = "HH:mm"
TEMPERATURE_RANGE = (500, 2000)

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.db = Database()
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
        
        # Добавляем обработчик изменения даты
        self.Плавка_дата.dateChanged.connect(self.generate_plavka_number)
        
        # Генерируем начальный номер плавки
        self.generate_plavka_number()

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

            self.db.save_plavka({
                'id': id_number,
                'uchet_number': Учетный_номер,
                'date': Плавка_дата.toString("yyyy-MM-dd"),
                'plavka_number': Номер_плавки,
                'cluster_number': Номер_кластера,
                'senior_shift': Старший_смены_плавки,
                'participant1': Первый_участник_смены_плавки,
                'participant2': Второй_участник_смены_плавки,
                'participant3': Третий_участник_смены_плавки,
                'participant4': Четвертый_участник_смены_плавки,
                'casting_name': Наименование_отливки,
                'experiment_type': Тип_эксперемента,
                'sector_A': Сектор_A_опоки,
                'heating_time_A': Плавка_время_прогрева_ковша_A,
                'movement_time_A': Плавка_время_перемещения_A,
                'pouring_time_A': Плавка_время_заливки_A,
                'temperature_A': Плавка_температура_заливки_A,
                'sector_B': Сектор_B_опоки,
                'heating_time_B': Плавка_время_прогрева_ковша_B,
                'movement_time_B': Плавка_время_перемещения_B,
                'pouring_time_B': Плавка_время_заливки_B,
                'temperature_B': Плавка_температура_заливки_B,
                'sector_C': Сектор_C_опоки,
                'heating_time_C': Плавка_время_прогрева_ковша_C,
                'movement_time_C': Плавка_время_перемещения_C,
                'pouring_time_C': Плавка_время_заливки_C,
                'temperature_C': Плавка_температура_заливки_C,
                'sector_D': Сектор_D_опоки,
                'heating_time_D': Плавка_время_прогрева_ковша_D,
                'movement_time_D': Плавка_время_перемещения_D,
                'pouring_time_D': Плавка_время_заливки_D,
                'temperature_D': Плавка_температура_заливки_D,
                'comment': Комментарий
            })

            QMessageBox.information(self, "Успех", "Данные успешно сохранены!")
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
        dialog = SearchDialog(self.db, parent=self)
        dialog.exec()

class SearchDialog(QDialog):
    def __init__(self, db, parent=None):
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
        self.results_table.setColumnCount(2)
        self.results_table.setHorizontalHeaderLabels(["Дата", "Температура"])
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
            record_date = QDate.fromString(data['date'], "yyyy-MM-dd")
            if not (self.date_from.date() <= record_date <= self.date_to.date()):
                return False
            
            # Фильтр по типу отливки
            if self.filter_casting.currentText() != "Все" and \
               data['casting_name'] != self.filter_casting.currentText():
                return False
            
            # Фильтр по температуре
            if self.temp_from.text() and self.temp_to.text():
                try:
                    temp_A = float(data['temperature_A'])
                    temp_B = float(data['temperature_B'])
                    temp_C = float(data['temperature_C'])
                    temp_D = float(data['temperature_D'])
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
            records = self.db.get_records()
            
            stats = {
                'total_records': 0,
                'avg_temp': [],
                'casting_types': {},
                'participants': set(),
                'min_temp': float('inf'),
                'max_temp': float('-inf')
            }
            
            for record in records:
                # Проверяем фильтры
                if not self.apply_filters(record, self.db.get_headers()):
                    continue
                    
                # Ищем совпадения
                stats['total_records'] += 1
                
                # Температура
                try:
                    temp_A = float(record['temperature_A'])
                    temp_B = float(record['temperature_B'])
                    temp_C = float(record['temperature_C'])
                    temp_D = float(record['temperature_D'])
                    stats['avg_temp'].append(temp_A)
                    stats['avg_temp'].append(temp_B)
                    stats['avg_temp'].append(temp_C)
                    stats['avg_temp'].append(temp_D)
                    stats['min_temp'] = min(stats['min_temp'], temp_A, temp_B, temp_C, temp_D)
                    stats['max_temp'] = max(stats['max_temp'], temp_A, temp_B, temp_C, temp_D)
                except (ValueError, TypeError):
                    pass
                
                # Типы отливок
                casting = record['casting_name']
                stats['casting_types'][casting] = stats['casting_types'].get(casting, 0) + 1
                
                # Участники
                stats['participants'].add(record['senior_shift'])
                stats['participants'].add(record['participant1'])
                stats['participants'].add(record['participant2'])
                stats['participants'].add(record['participant3'])
                stats['participants'].add(record['participant4'])
            
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
            
        except Exception as e:
            logging.error(f"Ошибка при обновлении статистики: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Ошибка при обновлении статистики: {str(e)}")

    def search_records(self):
        search_text = self.search_input.text().lower()
        try:
            records = self.db.get_records()
            
            self.results_table.setRowCount(0)
            
            for row in records:
                # Проверяем фильтры
                if not self.apply_filters(row, self.db.get_headers()):
                    continue
                    
                # Ищем совпадения
                for value in row.values():
                    if value and str(value).lower().find(search_text) != -1:
                        row_position = self.results_table.rowCount()
                        self.results_table.insertRow(row_position)
                        
                        for col, field in enumerate(self.db.get_headers()):
                            field_value = row[field]
                            self.results_table.setItem(
                                row_position, col, 
                                QTableWidgetItem(str(field_value) if field_value is not None else ""))
                        break
            
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
                
                df = pd.DataFrame(data, columns=self.db.get_headers())
                
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
            if not os.path.exists('backups'):
                os.makedirs('backups')
            
            # Формируем имя файла резервной копии
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_file = os.path.join('backups', f'plavka_backup_{timestamp}.db')
            
            # Копируем файл
            shutil.copy2('plavka.db', backup_file)
            
            # Удаляем старые резервные копии если их больше MAX_BACKUPS
            backups = sorted([os.path.join('backups', f) for f in os.listdir('backups')])
            while len(backups) > 5:
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
            record = self.db.get_record(self.record_id)
            
            # Заполняем поля данными
            self.fill_fields(record)
            
        except Exception as e:
            logging.error(f"Ошибка при загрузке записи: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Ошибка при загрузке записи: {str(e)}")

    def fill_fields(self, record):
        """Заполняет поля формы данными из записи"""
        try:
            # Заполняем поля
            self.Плавка_дата.setDate(QDate.fromString(record['date'], "yyyy-MM-dd"))
            self.Номер_плавки.setText(record['plavka_number'])
            self.Номер_кластера.setText(record['cluster_number'])
            
            # Устанавливаем значения комбобоксов
            self.Старший_смены_плавки.setCurrentText(record['senior_shift'])
            self.Первый_участник_смены_плавки.setCurrentText(record['participant1'])
            self.Второй_участник_смены_плавки.setCurrentText(record['participant2'])
            self.Третий_участник_смены_плавки.setCurrentText(record['participant3'])
            self.Четвертый_участник_смены_плавки.setCurrentText(record['participant4'])
            
            self.Наименование_отливки.setCurrentText(record['casting_name'])
            self.Тип_эксперемента.setCurrentText(record['experiment_type'])
            
            # Заполняем секторы опоки
            self.Сектор_A_опоки.setText(record['sector_A'])
            self.Сектор_B_опоки.setText(record['sector_B'])
            self.Сектор_C_опоки.setText(record['sector_C'])
            self.Сектор_D_опоки.setText(record['sector_D'])
            
            # Заполняем время и температуру
            self.Плавка_время_прогрева_ковша_A.setText(record['heating_time_A'])
            self.Плавка_время_перемещения_A.setText(record['movement_time_A'])
            self.Плавка_время_заливки_A.setText(record['pouring_time_A'])
            self.Плавка_температура_заливки_A.setText(record['temperature_A'])

            self.Плавка_время_прогрева_ковша_B.setText(record['heating_time_B'])
            self.Плавка_время_перемещения_B.setText(record['movement_time_B'])
            self.Плавка_время_заливки_B.setText(record['pouring_time_B'])
            self.Плавка_температура_заливки_B.setText(record['temperature_B'])

            self.Плавка_время_прогрева_ковша_C.setText(record['heating_time_C'])
            self.Плавка_время_перемещения_C.setText(record['movement_time_C'])
            self.Плавка_время_заливки_C.setText(record['pouring_time_C'])
            self.Плавка_температура_заливки_C.setText(record['temperature_C'])

            self.Плавка_время_прогрева_ковша_D.setText(record['heating_time_D'])
            self.Плавка_время_перемещения_D.setText(record['movement_time_D'])
            self.Плавка_время_заливки_D.setText(record['pouring_time_D'])
            self.Плавка_температура_заливки_D.setText(record['temperature_D'])

            self.Комментарий.setText(record['comment'])
            
        except Exception as e:
            logging.error(f"Ошибка при заполнении полей: {str(e)}")
            raise

    def save_changes(self):
        """Сохраняет изменения в базу"""
        try:
            # Обновляем данные в базе
            self.db.update_record(self.record_id, {
                'date': self.Плавка_дата.date().toString("yyyy-MM-dd"),
                'plavka_number': self.Номер_плавки.text(),
                'cluster_number': self.Номер_кластера.text(),
                'senior_shift': self.Старший_смены_плавки.currentText(),
                'participant1': self.Первый_участник_смены_плавки.currentText(),
                'participant2': self.Второй_участник_смены_плавки.currentText(),
                'participant3': self.Третий_участник_смены_плавки.currentText(),
                'participant4': self.Четвертый_участник_смены_плавки.currentText(),
                'casting_name': self.Наименование_отливки.currentText(),
                'experiment_type': self.Тип_эксперемента.currentText(),
                'sector_A': self.Сектор_A_опоки.text(),
                'heating_time_A': self.Плавка_время_прогрева_ковша_A.text(),
                'movement_time_A': self.Плавка_время_перемещения_A.text(),
                'pouring_time_A': self.Плавка_время_заливки_A.text(),
                'temperature_A': self.Плавка_температура_заливки_A.text(),
                'sector_B': self.Сектор_B_опоки.text(),
                'heating_time_B': self.Плавка_время_прогрева_ковша_B.text(),
                'movement_time_B': self.Плавка_время_перемещения_B.text(),
                'pouring_time_B': self.Плавка_время_заливки_B.text(),
                'temperature_B': self.Плавка_температура_заливки_B.text(),
                'sector_C': self.Сектор_C_опоки.text(),
                'heating_time_C': self.Плавка_время_прогрева_ковша_C.text(),
                'movement_time_C': self.Плавка_время_перемещения_C.text(),
                'pouring_time_C': self.Плавка_время_заливки_C.text(),
                'temperature_C': self.Плавка_температура_заливки_C.text(),
                'sector_D': self.Сектор_D_опоки.text(),
                'heating_time_D': self.Плавка_время_прогрева_ковша_D.text(),
                'movement_time_D': self.Плавка_время_перемещения_D.text(),
                'pouring_time_D': self.Плавка_время_заливки_D.text(),
                'temperature_D': self.Плавка_температура_заливки_D.text(),
                'comment': self.Комментарий.toPlainText()
            })

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
