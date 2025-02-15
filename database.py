import sqlite3
import logging
from datetime import datetime
from typing import Optional, List, Dict, Any
from models import MeltRecord

class Database:
    """Класс для работы с базой данных SQLite"""
    
    def __init__(self, db_name: str = 'plavka.db') -> None:
        """
        Инициализирует подключение к базе данных
        
        Args:
            db_name: Имя файла базы данных
        """
        self.db_name = db_name
        self.init_database()

    def init_database(self) -> None:
        """Инициализация базы данных и создание таблиц"""
        try:
            with sqlite3.connect(self.db_name) as conn:
                cursor = conn.cursor()
                
                # Создаем таблицу для основной информации о плавке
                cursor.execute('''
                    CREATE TABLE IF NOT EXISTS plavki (
                        id TEXT PRIMARY KEY,
                        uchet_number TEXT,
                        date TEXT,
                        plavka_number TEXT NOT NULL,
                        cluster_number TEXT,
                        senior_shift TEXT,
                        participant1 TEXT,
                        participant2 TEXT,
                        participant3 TEXT,
                        participant4 TEXT,
                        casting_name TEXT,
                        experiment_type TEXT,
                        comment TEXT,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                ''')

                # Создаем таблицу для секторов
                cursor.execute('''
                    CREATE TABLE IF NOT EXISTS sectors (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        plavka_id TEXT,
                        sector_name TEXT,
                        sector_number TEXT,
                        heating_time TEXT,
                        movement_time TEXT,
                        pouring_time TEXT,
                        temperature REAL,
                        FOREIGN KEY (plavka_id) REFERENCES plavki (id),
                        UNIQUE(plavka_id, sector_name)
                    )
                ''')

                conn.commit()
                logging.info("База данных успешно инициализирована")

        except sqlite3.Error as e:
            logging.error(f"Ошибка при инициализации базы данных: {e}")
            raise

    def save_plavka(self, record: MeltRecord) -> bool:
        """
        Сохранение данных о плавке
        
        Args:
            record: Объект записи о плавке
            
        Returns:
            bool: True если сохранение успешно, False в случае ошибки
        """
        try:
            with sqlite3.connect(self.db_name) as conn:
                cursor = conn.cursor()
                data = record.to_dict()
                
                # Сохраняем основную информацию о плавке
                cursor.execute('''
                    INSERT INTO plavki (
                        id, uchet_number, date, plavka_number, cluster_number,
                        senior_shift, participant1, participant2, participant3, participant4,
                        casting_name, experiment_type, comment
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    data['id'], data['uchet_number'], data['date'], data['plavka_number'],
                    data['cluster_number'], data['senior_shift'], data['participant1'],
                    data['participant2'], data['participant3'], data['participant4'],
                    data['casting_name'], data['experiment_type'], data['comment']
                ))

                # Сохраняем данные по секторам
                for sector in ['A', 'B', 'C', 'D']:
                    sector_data = getattr(record, f'sector_{sector.lower()}')
                    if sector_data:
                        cursor.execute('''
                            INSERT INTO sectors (
                                plavka_id, sector_name, sector_number,
                                heating_time, movement_time, pouring_time, temperature
                            ) VALUES (?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            data['id'], 
                            sector,
                            sector_data.sector_number,
                            data[f'heating_time_{sector.lower()}'],
                            data[f'movement_time_{sector.lower()}'],
                            data[f'pouring_time_{sector.lower()}'],
                            sector_data.temperature
                        ))

                conn.commit()
                return True

        except sqlite3.Error as e:
            logging.error(f"Ошибка при сохранении данных: {e}")
            return False

    def check_duplicate_id(self, id_number: str) -> bool:
        """
        Проверка существования ID в базе данных
        
        Args:
            id_number: ID для проверки
            
        Returns:
            bool: True если ID существует, False если нет
        """
        try:
            with sqlite3.connect(self.db_name) as conn:
                cursor = conn.cursor()
                cursor.execute('SELECT id FROM plavki WHERE id = ?', (id_number,))
                return cursor.fetchone() is not None
        except sqlite3.Error as e:
            logging.error(f"Ошибка при проверке ID: {e}")
            return False

    def get_records(self) -> List[MeltRecord]:
        """
        Получение всех записей из базы данных
        
        Returns:
            List[MeltRecord]: Список записей о плавках
        """
        try:
            with sqlite3.connect(self.db_name) as conn:
                cursor = conn.cursor()
                
                query = '''
                    SELECT p.*, 
                           s1.sector_number as sector_a, s1.heating_time as heating_a, s1.movement_time as movement_a, s1.pouring_time as pouring_a, s1.temperature as temp_a,
                           s2.sector_number as sector_b, s2.heating_time as heating_b, s2.movement_time as movement_b, s2.pouring_time as pouring_b, s2.temperature as temp_b,
                           s3.sector_number as sector_c, s3.heating_time as heating_c, s3.movement_time as movement_c, s3.pouring_time as pouring_c, s3.temperature as temp_c,
                           s4.sector_number as sector_d, s4.heating_time as heating_d, s4.movement_time as movement_d, s4.pouring_time as pouring_d, s4.temperature as temp_d
                    FROM plavki p
                    LEFT JOIN sectors s1 ON p.id = s1.plavka_id AND s1.sector_name = 'A'
                    LEFT JOIN sectors s2 ON p.id = s2.plavka_id AND s2.sector_name = 'B'
                    LEFT JOIN sectors s3 ON p.id = s3.plavka_id AND s3.sector_name = 'C'
                    LEFT JOIN sectors s4 ON p.id = s4.plavka_id AND s4.sector_name = 'D'
                    ORDER BY p.date DESC
                '''
                
                cursor.execute(query)
                columns = [description[0] for description in cursor.description]
                records = []
                
                for row in cursor.fetchall():
                    record_dict = dict(zip(columns, row))
                    records.append(MeltRecord.from_dict(record_dict))
                    
                return records

        except sqlite3.Error as e:
            logging.error(f"Ошибка при получении записей: {e}")
            return []

    def get_record_by_id(self, record_id: str) -> Optional[MeltRecord]:
        """
        Получение записи по ID
        
        Args:
            record_id: ID записи
            
        Returns:
            Optional[MeltRecord]: Запись о плавке или None если запись не найдена
        """
        try:
            with sqlite3.connect(self.db_name) as conn:
                cursor = conn.cursor()
                
                query = '''
                    SELECT p.*, 
                           s1.sector_number as sector_a, s1.heating_time as heating_a, s1.movement_time as movement_a, s1.pouring_time as pouring_a, s1.temperature as temp_a,
                           s2.sector_number as sector_b, s2.heating_time as heating_b, s2.movement_time as movement_b, s2.pouring_time as pouring_b, s2.temperature as temp_b,
                           s3.sector_number as sector_c, s3.heating_time as heating_c, s3.movement_time as movement_c, s3.pouring_time as pouring_c, s3.temperature as temp_c,
                           s4.sector_number as sector_d, s4.heating_time as heating_d, s4.movement_time as movement_d, s4.pouring_time as pouring_d, s4.temperature as temp_d
                    FROM plavki p
                    LEFT JOIN sectors s1 ON p.id = s1.plavka_id AND s1.sector_name = 'A'
                    LEFT JOIN sectors s2 ON p.id = s2.plavka_id AND s2.sector_name = 'B'
                    LEFT JOIN sectors s3 ON p.id = s3.plavka_id AND s3.sector_name = 'C'
                    LEFT JOIN sectors s4 ON p.id = s4.plavka_id AND s4.sector_name = 'D'
                    WHERE p.id = ?
                '''
                
                cursor.execute(query, (record_id,))
                row = cursor.fetchone()
                if row:
                    columns = [description[0] for description in cursor.description]
                    record_dict = dict(zip(columns, row))
                    return MeltRecord.from_dict(record_dict)
                return None

        except sqlite3.Error as e:
            logging.error(f"Ошибка при получении записи: {e}")
            return None

    def update_record(self, record: MeltRecord) -> bool:
        """
        Обновление существующей записи
        
        Args:
            record: Объект записи о плавке
            
        Returns:
            bool: True если обновление успешно, False в случае ошибки
        """
        try:
            with sqlite3.connect(self.db_name) as conn:
                cursor = conn.cursor()
                data = record.to_dict()
                
                # Обновляем основную информацию
                cursor.execute('''
                    UPDATE plavki SET
                        uchet_number = ?, date = ?, plavka_number = ?, cluster_number = ?,
                        senior_shift = ?, participant1 = ?, participant2 = ?, participant3 = ?,
                        participant4 = ?, casting_name = ?, experiment_type = ?, comment = ?
                    WHERE id = ?
                ''', (
                    data['uchet_number'], data['date'], data['plavka_number'],
                    data['cluster_number'], data['senior_shift'], data['participant1'],
                    data['participant2'], data['participant3'], data['participant4'],
                    data['casting_name'], data['experiment_type'], data['comment'],
                    data['id']
                ))

                # Обновляем данные секторов
                for sector in ['A', 'B', 'C', 'D']:
                    sector_data = getattr(record, f'sector_{sector.lower()}')
                    if sector_data:
                        cursor.execute('''
                            UPDATE sectors SET
                                sector_number = ?, heating_time = ?, movement_time = ?,
                                pouring_time = ?, temperature = ?
                            WHERE plavka_id = ? AND sector_name = ?
                        ''', (
                            sector_data.sector_number,
                            data[f'heating_time_{sector.lower()}'],
                            data[f'movement_time_{sector.lower()}'],
                            data[f'pouring_time_{sector.lower()}'],
                            sector_data.temperature,
                            data['id'],
                            sector
                        ))

                conn.commit()
                return True

        except sqlite3.Error as e:
            logging.error(f"Ошибка при обновлении записи: {e}")
            return False
