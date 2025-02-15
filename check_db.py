import sqlite3
import logging
from datetime import datetime

logging.basicConfig(level=logging.INFO)

def check_database():
    try:
        with sqlite3.connect('plavka.db') as conn:
            cursor = conn.cursor()
            
            # Получаем последние 10 записей, отсортированные по дате
            cursor.execute('''
                SELECT date, plavka_number 
                FROM plavki 
                ORDER BY date DESC 
                LIMIT 10
            ''')
            
            print("\nПоследние 10 записей:")
            print("Дата\t\tНомер плавки")
            print("-" * 30)
            
            for row in cursor.fetchall():
                date_str = row[0]
                plavka_number = row[1]
                print(f"{date_str}\t{plavka_number}")
            
            # Получаем все уникальные номера плавок для февраля
            cursor.execute('''
                SELECT DISTINCT plavka_number 
                FROM plavki 
                WHERE strftime('%m', date) = '02'
                ORDER BY plavka_number
            ''')
            
            print("\nВсе номера плавок за февраль:")
            numbers = cursor.fetchall()
            print(", ".join(row[0] for row in numbers))
            
            # Подсчитываем количество записей по месяцам
            cursor.execute('''
                SELECT strftime('%m', date) as month, COUNT(*) as count 
                FROM plavki 
                GROUP BY month 
                ORDER BY month
            ''')
            
            print("\nКоличество записей по месяцам:")
            print("Месяц\tКоличество")
            print("-" * 20)
            
            for row in cursor.fetchall():
                month = row[0]
                count = row[1]
                print(f"{month}\t{count}")
                
    except sqlite3.Error as e:
        print(f"Ошибка при работе с базой данных: {e}")

if __name__ == "__main__":
    check_database()
