import pandas as pd
import sqlite3
import os
import datetime

EXCEL_FILE = 'plavka.xlsx'
SQLITE_DB = 'plavka.db'
TABLE_NAME = 'Плавки'

# Чтение данных из Excel
if not os.path.exists(EXCEL_FILE):
    raise FileNotFoundError(f"Файл {EXCEL_FILE} не найден.")

df = pd.read_excel(EXCEL_FILE, engine='openpyxl')

# Копировать 'Плавка_дата' как строку из Excel в формате DD.MM.YYYY
if 'Плавка_дата' in df.columns:
    df['Плавка_дата'] = df['Плавка_дата'].apply(lambda x: x.strftime('%d.%m.%Y') if isinstance(x, (datetime.datetime, pd.Timestamp)) else x)

# Приведение типов (температуры к int, если возможно)
for col in [
    'Плавка_температура_заливки_A',
    'Плавка_температура_заливки_B',
    'Плавка_температура_заливки_C',
    'Плавка_температура_заливки_D']:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce').astype('Int64')

# Приведение номеров опок к строке без .0
for col in ['Сектор_A_опоки', 'Сектор_B_опоки', 'Сектор_C_опоки', 'Сектор_D_опоки']:
    if col in df.columns:
        df[col] = df[col].apply(lambda x: str(int(x)) if pd.notnull(x) and isinstance(x, float) and x.is_integer() else str(x) if pd.notnull(x) else '')

# Преобразуем все значения типа datetime.time к строке формата 'HH:MM'
for col in df.columns:
    if df[col].dtype == 'object':
        df[col] = df[col].apply(lambda x: x.strftime('%H:%M') if hasattr(x, 'strftime') else x)

# Подключение к SQLite
conn = sqlite3.connect(SQLITE_DB)
cursor = conn.cursor()

# Создание таблицы, если не существует
with open('db_schema.sql', encoding='utf-8') as f:
    schema = f.read()
cursor.executescript(schema)

# Запись данных в таблицу
# (Если id уже есть, не вставлять его, иначе - вставлять все)
insert_cols = [col for col in df.columns if col != 'id']
placeholders = ','.join(['?'] * len(insert_cols))
insert_sql = f"INSERT INTO {TABLE_NAME} ({','.join(insert_cols)}) VALUES ({placeholders})"

for _, row in df.iterrows():
    values = [row[col] if pd.notnull(row[col]) else None for col in insert_cols]
    cursor.execute(insert_sql, values)

conn.commit()
conn.close()

print(f"Миграция завершена. Перенесено записей: {len(df)}") 