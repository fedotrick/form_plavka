import sqlite3
import logging
from datetime import datetime

logging.basicConfig(level=logging.INFO)

def fix_database():
    try:
        with sqlite3.connect('plavka.db') as conn:
            cursor = conn.cursor()
            
            # Получаем все записи за февраль 2025 года
            cursor.execute('''
                SELECT id, date, plavka_number 
                FROM plavki 
                WHERE strftime('%m', date) = '02' AND strftime('%Y', date) = '2025'
                ORDER BY date, plavka_number
            ''')
            
            records = cursor.fetchall()
            
            # Удаляем все записи с номером больше 157 за февраль
            cursor.execute('''
                DELETE FROM plavki 
                WHERE strftime('%m', date) = '02' 
                AND strftime('%Y', date) = '2025'
                AND CAST(SUBSTR(plavka_number, 3) AS INTEGER) > 157
            ''')
            
            # Исправляем формат номеров (добавляем ведущие нули)
            cursor.execute('''
                SELECT id, plavka_number 
                FROM plavki 
                WHERE strftime('%m', date) = '02'
                AND strftime('%Y', date) = '2025'
            ''')
            
            for row in cursor.fetchall():
                id_num, plavka_number = row
                if '-' in plavka_number:
                    month, number = plavka_number.split('-')
                    new_number = f"{month}-{int(number):03d}"
                    if new_number != plavka_number:
                        cursor.execute('''
                            UPDATE plavki 
                            SET plavka_number = ? 
                            WHERE id = ?
                        ''', (new_number, id_num))
            
            conn.commit()
            
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
    fix_database()
