import sqlite3
import shutil
import os
from datetime import datetime
import logging

logging.basicConfig(level=logging.INFO)

def recreate_database():
    # Создаем резервную копию
    if os.path.exists('plavka.db'):
        backup_name = f'plavka_backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.db'
        shutil.copy2('plavka.db', backup_name)
        logging.info(f"Создана резервная копия: {backup_name}")
        
    try:
        # Подключаемся к базе данных
        with sqlite3.connect('plavka.db') as conn:
            cursor = conn.cursor()
            
            # Получаем структуру таблицы
            cursor.execute('PRAGMA table_info(plavki)')
            columns_info = cursor.fetchall()
            
            # Создаем SQL для временной таблицы
            create_sql = 'CREATE TABLE plavki_temp (\n'
            create_sql += ',\n'.join([f"{col[1]} {col[2]}" + 
                                    (" PRIMARY KEY" if col[5] == 1 else "") + 
                                    (" NOT NULL" if col[3] == 1 else "") + 
                                    (f" DEFAULT {col[4]}" if col[4] is not None else "")
                                    for col in columns_info])
            create_sql += ')'
            
            # Создаем временную таблицу
            cursor.execute('DROP TABLE IF EXISTS plavki_temp')
            cursor.execute(create_sql)
            
            # Получаем все записи
            cursor.execute('SELECT * FROM plavki')
            all_records = cursor.fetchall()
            
            # Получаем имена столбцов
            columns = [col[1] for col in columns_info]
            
            # Копируем только правильные записи
            for record in all_records:
                record_dict = dict(zip(columns, record))
                
                # Пропускаем записи февраля с номером больше 157
                if record_dict['date']:
                    try:
                        date = datetime.strptime(record_dict['date'], '%Y-%m-%d')
                        if date.month == 2 and date.year == 2025:
                            plavka_number = record_dict['plavka_number']
                            if plavka_number:
                                try:
                                    month, number = plavka_number.split('-')
                                    if int(month) == 2 and int(number) > 157:
                                        logging.info(f"Пропускаем запись с номером {plavka_number}")
                                        continue
                                except ValueError:
                                    continue
                    except ValueError:
                        continue
                
                # Форматируем номер плавки
                if record_dict['plavka_number']:
                    try:
                        month, number = record_dict['plavka_number'].split('-')
                        record_dict['plavka_number'] = f"{int(month)}-{int(number):03d}"
                    except ValueError:
                        continue
                
                # Вставляем запись
                placeholders = ','.join(['?' for _ in columns])
                cursor.execute(
                    f'INSERT INTO plavki_temp ({",".join(columns)}) VALUES ({placeholders})',
                    [record_dict[col] for col in columns]
                )
            
            # Заменяем старую таблицу на новую
            cursor.execute('DROP TABLE plavki')
            cursor.execute('ALTER TABLE plavki_temp RENAME TO plavki')
            
            # Проверяем результат
            cursor.execute('''
                SELECT date, plavka_number 
                FROM plavki 
                WHERE strftime('%m', date) = '02' AND strftime('%Y', date) = '2025'
                ORDER BY date DESC, plavka_number DESC 
                LIMIT 5
            ''')
            
            print("\nПоследние 5 записей за февраль 2025:")
            print("Дата\t\tНомер плавки")
            print("-" * 30)
            
            for row in cursor.fetchall():
                date_str = row[0]
                plavka_number = row[1]
                print(f"{date_str}\t{plavka_number}")
                
    except sqlite3.Error as e:
        print(f"Ошибка при работе с базой данных: {e}")

if __name__ == "__main__":
    recreate_database()
