import sqlite3
import logging
from datetime import datetime

class Database:
    def __init__(self, db_name='plavka.db'):
        self.db_name = db_name
        self.init_database()

    def init_database(self):
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

    def save_plavka(self, data):
        """Сохранение данных о плавке"""
        try:
            with sqlite3.connect(self.db_name) as conn:
                cursor = conn.cursor()
                
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
                sectors = ['A', 'B', 'C', 'D']
                for sector in sectors:
                    cursor.execute('''
                        INSERT INTO sectors (
                            plavka_id, sector_name, sector_number,
                            heating_time, movement_time, pouring_time, temperature
                        ) VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        data['id'], 
                        sector,
                        data[f'sector_{sector}'],
                        data[f'heating_time_{sector}'],
                        data[f'movement_time_{sector}'],
                        data[f'pouring_time_{sector}'],
                        data[f'temperature_{sector}']
                    ))

                conn.commit()
                return True

        except sqlite3.Error as e:
            logging.error(f"Ошибка при сохранении данных: {e}")
            return False

    def check_duplicate_id(self, id_number):
        """Проверка существования ID в базе данных"""
        try:
            with sqlite3.connect(self.db_name) as conn:
                cursor = conn.cursor()
                cursor.execute('SELECT id FROM plavki WHERE id = ?', (id_number,))
                return cursor.fetchone() is not None
        except sqlite3.Error as e:
            logging.error(f"Ошибка при проверке ID: {e}")
            return False

    def search_records(self, filters=None):
        """Поиск записей с применением фильтров"""
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
                '''
                
                where_clauses = []
                params = []
                
                if filters:
                    if filters.get('date_from'):
                        where_clauses.append('p.date >= ?')
                        params.append(filters['date_from'])
                    if filters.get('date_to'):
                        where_clauses.append('p.date <= ?')
                        params.append(filters['date_to'])
                    if filters.get('plavka_number'):
                        where_clauses.append('p.plavka_number LIKE ?')
                        params.append(f"%{filters['plavka_number']}%")
                    if filters.get('casting_name'):
                        where_clauses.append('p.casting_name LIKE ?')
                        params.append(f"%{filters['casting_name']}%")

                if where_clauses:
                    query += ' WHERE ' + ' AND '.join(where_clauses)

                query += ' ORDER BY p.date DESC'
                
                cursor.execute(query, params)
                return cursor.fetchall()

        except sqlite3.Error as e:
            logging.error(f"Ошибка при поиске записей: {e}")
            return []

    def get_record_by_id(self, record_id):
        """Получение записи по ID"""
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
                return cursor.fetchone()

        except sqlite3.Error as e:
            logging.error(f"Ошибка при получении записи: {e}")
            return None

    def get_records(self):
        """Получение всех записей из базы данных"""
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
                    record = dict(zip(columns, row))
                    records.append(record)
                
                return records

        except sqlite3.Error as e:
            logging.error(f"Ошибка при получении записей: {e}")
            return []

    def update_record(self, data):
        """Обновление существующей записи"""
        try:
            with sqlite3.connect(self.db_name) as conn:
                cursor = conn.cursor()
                
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
                sectors = ['A', 'B', 'C', 'D']
                for sector in sectors:
                    cursor.execute('''
                        UPDATE sectors SET
                            sector_number = ?, heating_time = ?,
                            movement_time = ?, pouring_time = ?, temperature = ?
                        WHERE plavka_id = ? AND sector_name = ?
                    ''', (
                        data[f'sector_{sector}'],
                        data[f'heating_time_{sector}'],
                        data[f'movement_time_{sector}'],
                        data[f'pouring_time_{sector}'],
                        data[f'temperature_{sector}'],
                        data['id'],
                        sector
                    ))

                conn.commit()
                return True

        except sqlite3.Error as e:
            logging.error(f"Ошибка при обновлении записи: {e}")
            return False
