import pandas as pd
import os
from openpyxl import Workbook

# Пути к файлам
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(SCRIPT_DIR, 'data')
OLD_EXCEL = os.path.join(DATA_DIR, 'plavka.xlsx')
NEW_EXCEL = os.path.join(DATA_DIR, 'plavka_new.xlsx')

# Новый порядок столбцов
new_columns = ['id_plavka', 'Учетный_номер', 'Плавка_дата', 
              'Номер_плавки', 'Номер_кластера', 'Старший_смены_плавки', 
              'Первый_участник_смены_плавки', 'Второй_участник_смены_плавки', 
              'Третий_участник_смены_плавки', 'Четвертый_участник_смены_плавки', 
              'Наименование_отливки', 'Тип_эксперемента', 
              'Сектор_A_опоки', 'Сектор_B_опоки', 'Сектор_C_опоки', 'Сектор_D_опоки',
              'Плавка_время_прогрева_ковша_A', 'Плавка_время_перемещения_A', 
              'Плавка_время_заливки_A', 'Плавка_температура_заливки_A',
              'Плавка_время_прогрева_ковша_B', 'Плавка_время_перемещения_B', 
              'Плавка_время_заливки_B', 'Плавка_температура_заливки_B',
              'Плавка_время_прогрева_ковша_C', 'Плавка_время_перемещения_C', 
              'Плавка_время_заливки_C', 'Плавка_температура_заливки_C',
              'Плавка_время_прогрева_ковша_D', 'Плавка_время_перемещения_D', 
              'Плавка_время_заливки_D', 'Плавка_температура_заливки_D',
              'Комментарий', 'Плавка_время_заливки', 'id']

try:
    # Проверяем существование старого файла
    if not os.path.exists(OLD_EXCEL):
        print("Старый файл не найден. Создаем новый файл с правильной структурой.")
        wb = Workbook()
        ws = wb.active
        ws.append(new_columns)
        wb.save(NEW_EXCEL)
    else:
        # Читаем данные из старого файла
        df = pd.read_excel(OLD_EXCEL)
        
        # Создаем копию с новым порядком столбцов
        # Если какие-то столбцы отсутствуют, они будут созданы с пустыми значениями
        df_new = pd.DataFrame(columns=new_columns)
        
        # Копируем данные из существующих столбцов
        for col in new_columns:
            if col in df.columns:
                df_new[col] = df[col]
        
        # Сохраняем в новый файл
        df_new.to_excel(NEW_EXCEL, index=False)
        
        print(f"Данные успешно перенесены в новый файл: {NEW_EXCEL}")
        
        # Создаем бэкап старого файла
        backup_file = OLD_EXCEL + '.bak'
        if os.path.exists(OLD_EXCEL):
            os.rename(OLD_EXCEL, backup_file)
            print(f"Создан бэкап старого файла: {backup_file}")
        
        # Переименовываем новый файл
        os.rename(NEW_EXCEL, OLD_EXCEL)
        print(f"Новый файл переименован в: {OLD_EXCEL}")
        
except Exception as e:
    print(f"Произошла ошибка: {str(e)}")
