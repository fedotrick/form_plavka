import streamlit as st
import pandas as pd
import sqlite3
from io import BytesIO
import plotly.express as px
import re
import os
from openpyxl import load_workbook

DB_FILE = 'plavka.db'
TABLE_NAME = 'Плавки'
EXCEL_FILE = 'plavka.xlsx'

st.set_page_config(page_title="Электронный журнал плавки", layout="wide")
st.title("Электронный журнал плавки")

# Функция для загрузки данных из SQLite
@st.cache_data
def load_data():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql(f'SELECT * FROM {TABLE_NAME}', conn)
    conn.close()
    return df

# Функция для добавления новой записи в Excel
def add_to_excel(new_row):
    try:
        if not os.path.exists(EXCEL_FILE):
            # Если файл не существует, создаем новый с заголовками
            df = pd.DataFrame([new_row])
            df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')
            st.success(f"Данные сохранены в новый файл {EXCEL_FILE}")
            return True
        
        # Загружаем существующий файл
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
        
        # Проверяем, совпадают ли заголовки
        headers = [cell.value for cell in ws[1]]
        
        # Добавляем новую строку
        row_values = []
        for header in headers:
            if header in new_row:
                row_values.append(new_row[header])
            else:
                row_values.append(None)
        
        ws.append(row_values)
        
        # Сохраняем изменения
        wb.save(EXCEL_FILE)
        st.success(f"Данные также сохранены в Excel файл {EXCEL_FILE}")
        return True
    except Exception as e:
        st.error(f"Ошибка при сохранении в Excel: {str(e)}")
        return False

def save_to_excel(df):
    # Приведение номеров опок к строке без .0
    for col in ['Сектор_A_опоки', 'Сектор_B_опоки', 'Сектор_C_опоки', 'Сектор_D_опоки']:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: str(int(x)) if pd.notnull(x) and isinstance(x, float) and x.is_integer() else str(x) if pd.notnull(x) else '')
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output

# Функция для проверки формата времени чч:мм
TIME_PATTERN = re.compile(r'^([01]?\d|2[0-3]):[0-5]\d$')
def is_valid_time(val):
    if not val: return True
    return bool(TIME_PATTERN.match(str(val)))

def parse_plavka_number(num):
    # num: строка вида '5-102' или '05-102'
    try:
        parts = str(num).split('-')
        if len(parts) == 2:
            nnn = parts[1].zfill(3)
        else:
            nnn = str(num).zfill(3)
        return nnn
    except Exception:
        return str(num).zfill(3)

def generate_id_plavka(date, num):
    nnn = parse_plavka_number(num)
    return f"{date.year}{date.month:02d}{nnn}"

def generate_uchet_number(date, num):
    nnn = parse_plavka_number(num)
    return f"{date.month:02d}-{nnn}/{str(date.year)[-2:]}"

def parse_import_message(text):
    # Парсим дату смены из шапки
    m = re.search(r'📅 ([0-9]{2}\.[0-9]{2}\.[0-9]{4})', text)
    plavka_date = m.group(1) if m else ''
    # Парсим старшего смены и участников из шапки
    m = re.search(r'Старший: ([^\n]+)', text)
    starshiy = m.group(1).strip() if m else ''
    m = re.search(r'Участники: ([^\n]+)', text)
    uchastniki = [x.strip() for x in m.group(1).split(',')] if m else []
    # Разбиваем на блоки по плавкам
    blocks = re.split(r'📊 Плавка \d+/\d+', text)
    results = []
    for block in blocks:
        if 'Маршрутная карта' not in block:
            continue
        data = {}
        # Дата смены для каждой плавки
        data['Плавка_дата'] = plavka_date
        # Старший смены и участники
        data['Старший_смены_плавки'] = starshiy
        for i, field in enumerate(['Первый_участник_смены_плавки', 'Второй_участник_смены_плавки', 'Третий_участник_смены_плавки', 'Четвертый_участник_смены_плавки']):
            data[field] = uchastniki[i] if i < len(uchastniki) else ''
        # Маршрутная карта
        m = re.search(r'Маршрутная карта: (\d+)', block)
        if m:
            data['Маршрутная_карта'] = m.group(1)
        # Учетный номер
        m = re.search(r'Учетный номер: ([0-9]{2}-[0-9]{3}/[0-9]{2})', block)
        if m:
            data['Учетный_номер'] = m.group(1)
        # Кластер
        m = re.search(r'Кластер: ([^\n]+)', block)
        if m:
            data['Номер_кластера'] = m.group(1)
        # Отливка
        m = re.search(r'Отливка: ([^\n]+)', block)
        if m:
            data['Наименование_отливки'] = m.group(1)
        # Литниковая система
        m = re.search(r'Литниковая система: ([^\n]+)', block)
        if m:
            data['Тип_эксперемента'] = m.group(1)
        # Опоки
        m = re.search(r'Опоки: ([^\n]+)', block)
        opoki = []
        if m:
            opoki = [x.strip().replace('Опока №', '') for x in m.group(1).split(',')]
            opoki = [str(int(o)) if o.isdigit() or (o.replace('.','',1).isdigit() and float(o).is_integer()) else o for o in opoki]
        # Температура
        m = re.search(r'Температура: ([0-9]+[.,]?[0-9]*)', block)
        temp = float(m.group(1).replace(',', '.')) if m else None
        # Время слива
        m = re.search(r'Время слива: ([0-9]{2}:[0-9]{2})', block)
        time_val = m.group(1) if m else ''
        # Заполняем сектора (A, B), остальные пустые
        for i, sector in enumerate(['A', 'B', 'C', 'D']):
            if i < len(opoki):
                data[f'Сектор_{sector}_опоки'] = opoki[i]
                data[f'Плавка_температура_заливки_{sector}'] = temp
                data[f'Плавка_время_заливки_{sector}'] = time_val
            else:
                data[f'Сектор_{sector}_опоки'] = ''
                data[f'Плавка_температура_заливки_{sector}'] = None
                data[f'Плавка_время_заливки_{sector}'] = ''
        results.append(data)
    return results

# Основные вкладки
menu = st.sidebar.radio("Навигация", ["Таблица", "Добавить запись", "Экспорт в Excel", "Аналитика и статистика", "Импорт из сообщения", "О программе"])

# Кнопка обновления данных в сайдбаре
if st.sidebar.button("🔄 Обновить данные"):
    st.cache_data.clear()
    st.rerun()

df = load_data()

if menu == "Таблица":
    st.subheader("Все записи журнала")
    # Расширенный фильтр по всем полям
    with st.expander("Расширенный поиск по всем полям"):
        search_col = st.selectbox("Поле для поиска", df.columns, index=0)
        search_val = st.text_input("Значение для поиска (точное или часть)")
        multi_search = st.checkbox("Комбинировать с фильтрами ниже")
    # Базовые фильтры
    col1, col2, col3 = st.columns(3)
    with col1:
        date_filter = st.text_input("Фильтр по дате (например, 2024-05-01)")
    with col2:
        participant_filter = st.text_input("Фильтр по участнику")
    with col3:
        name_filter = st.text_input("Фильтр по наименованию отливки")
    filtered = df.copy()
    # Применяем расширенный фильтр
    if search_val:
        filtered = filtered[filtered[search_col].astype(str).str.contains(search_val, case=False, na=False)]
    # Применяем базовые фильтры, если multi_search или search_val не задан
    if (not search_val) or multi_search:
        if date_filter:
            filtered = filtered[filtered['Плавка_дата'].astype(str).str.contains(date_filter)]
        if participant_filter:
            filtered = filtered[
                filtered['Старший_смены_плавки'].astype(str).str.contains(participant_filter) |
                filtered['Первый_участник_смены_плавки'].astype(str).str.contains(participant_filter) |
                filtered['Второй_участник_смены_плавки'].astype(str).str.contains(participant_filter) |
                filtered['Третий_участник_смены_плавки'].astype(str).str.contains(participant_filter) |
                filtered['Четвертый_участник_смены_плавки'].astype(str).str.contains(participant_filter)
            ]
        if name_filter:
            filtered = filtered[filtered['Наименование_отливки'].astype(str).str.contains(name_filter)]
    st.dataframe(filtered, use_container_width=True)

elif menu == "Добавить запись":
    st.subheader("Добавить новую запись")
    with st.form("add_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            Плавка_дата = st.date_input("Дата", format="DD.MM.YYYY")
            Номер_плавки = st.text_input("Номер плавки")
            Номер_кластера = st.text_input("Номер кластера")
            Плавка_время_заливки = st.text_input("Время слива (чч:мм)")
            Старший_смены_плавки = st.selectbox("Старший смены", ["", "Белков", "Карасев", "Ермаков", "Рабинович", "Валиулин", "Волков", "Семенов", "Левин", "Исмаилов", "Беляев", "Политов", "Кокшин", "Терентьев", "отсутствует"])
            Первый_участник_смены_плавки = st.selectbox("Участник 1", ["", "Белков", "Карасев", "Ермаков", "Рабинович", "Валиулин", "Волков", "Семенов", "Левин", "Исмаилов", "Беляев", "Политов", "Кокшин", "Терентьев", "отсутствует"])
            Второй_участник_смены_плавки = st.selectbox("Участник 2", ["", "Белков", "Карасев", "Ермаков", "Рабинович", "Валиулин", "Волков", "Семенов", "Левин", "Исмаилов", "Беляев", "Политов", "Кокшин", "Терентьев", "отсутствует"])
            Третий_участник_смены_плавки = st.selectbox("Участник 3", ["", "Белков", "Карасев", "Ермаков", "Рабинович", "Валиулин", "Волков", "Семенов", "Левин", "Исмаилов", "Беляев", "Политов", "Кокшин", "Терентьев", "отсутствует"])
            Четвертый_участник_смены_плавки = st.selectbox("Участник 4", ["", "Белков", "Карасев", "Ермаков", "Рабинович", "Валиулин", "Волков", "Семенов", "Левин", "Исмаилов", "Беляев", "Политов", "Кокшин", "Терентьев", "отсутствует"])
            Наименование_отливки = st.selectbox("Наименование отливки", ["", "Вороток", "Ригель", "Ригель optima", "Блок-картер", "Колесо РИТМ", "Накладка резьб", "Блок цилиндров", "Диагональ optima", "Кольцо"])
            Тип_эксперемента = st.selectbox("Тип эксперимента", ["", "Бумага", "Волокно"])
            Сектор_A_опоки = st.text_input("Сектор A опоки")
            Сектор_B_опоки = st.text_input("Сектор B опоки")
            Сектор_C_опоки = st.text_input("Сектор C опоки")
            Сектор_D_опоки = st.text_input("Сектор D опоки")
        with col2:
            Плавка_время_прогрева_ковша_A = st.text_input("Время прогрева ковша A (чч:мм)")
            Плавка_время_перемещения_A = st.text_input("Время перемещения A (чч:мм)")
            Плавка_время_заливки_A = st.text_input("Время заливки A (чч:мм)")
            Плавка_температура_заливки_A = st.number_input("Температура заливки A", min_value=500, max_value=2000, step=1, format="%d")
            Плавка_время_прогрева_ковша_B = st.text_input("Время прогрева ковша B (чч:мм)")
            Плавка_время_перемещения_B = st.text_input("Время перемещения B (чч:мм)")
            Плавка_время_заливки_B = st.text_input("Время заливки B (чч:мм)")
            Плавка_температура_заливки_B = st.number_input("Температура заливки B", min_value=500, max_value=2000, step=1, format="%d")
            Плавка_время_прогрева_ковша_C = st.text_input("Время прогрева ковша C (чч:мм)")
            Плавка_время_перемещения_C = st.text_input("Время перемещения C (чч:мм)")
            Плавка_время_заливки_C = st.text_input("Время заливки C (чч:мм)")
            Плавка_температура_заливки_C = st.number_input("Температура заливки C", min_value=500, max_value=2000, step=1, format="%d")
            Плавка_время_прогрева_ковша_D = st.text_input("Время прогрева ковша D (чч:мм)")
            Плавка_время_перемещения_D = st.text_input("Время перемещения D (чч:мм)")
            Плавка_время_заливки_D = st.text_input("Время заливки D (чч:мм)")
            Плавка_температура_заливки_D = st.number_input("Температура заливки D", min_value=500, max_value=2000, step=1, format="%d")
            Комментарий = st.text_area("Комментарий")
        submitted = st.form_submit_button("Добавить запись")
    if submitted:
        # Валидация
        errors = []
        # Проверка обязательных полей
        if not str(Плавка_дата) or not Номер_плавки:
            errors.append("Поля 'Дата' и 'Номер плавки' обязательны!")
        # Проверка формата времени
        for t in [Плавка_время_заливки, Плавка_время_прогрева_ковша_A, Плавка_время_перемещения_A, Плавка_время_заливки_A, Плавка_время_прогрева_ковша_B, Плавка_время_перемещения_B, Плавка_время_заливки_B, Плавка_время_прогрева_ковша_C, Плавка_время_перемещения_C, Плавка_время_заливки_C, Плавка_время_прогрева_ковша_D, Плавка_время_перемещения_D, Плавка_время_заливки_D]:
            if t and not is_valid_time(t):
                errors.append(f"Некорректный формат времени: {t}")
        # Проверка уникальности номера плавки на дату
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute(f"SELECT COUNT(*) FROM {TABLE_NAME} WHERE Плавка_дата=? AND Номер_плавки=?", (str(Плавка_дата), Номер_плавки))
        if cur.fetchone()[0] > 0:
            errors.append("Запись с таким номером плавки на эту дату уже существует!")
        conn.close()
        if errors:
            for err in errors:
                st.error(err)
        else:
            # Генерация служебных полей по новым правилам
            id_plavka = generate_id_plavka(Плавка_дата, Номер_плавки)
            uchet_number = generate_uchet_number(Плавка_дата, Номер_плавки)
            
            # Форматируем дату в DD.MM.YYYY
            formatted_date = Плавка_дата.strftime('%d.%m.%Y')
            
            new_row = {
                'id_plavka': id_plavka,
                'Учетный_номер': uchet_number,
                'Плавка_дата': formatted_date,
                'Номер_плавки': Номер_плавки,
                'Номер_кластера': Номер_кластера,
                'Старший_смены_плавки': Старший_смены_плавки,
                'Первый_участник_смены_плавки': Первый_участник_смены_плавки,
                'Второй_участник_смены_плавки': Второй_участник_смены_плавки,
                'Третий_участник_смены_плавки': Третий_участник_смены_плавки,
                'Четвертый_участник_смены_плавки': Четвертый_участник_смены_плавки,
                'Наименование_отливки': Наименование_отливки,
                'Тип_эксперемента': Тип_эксперемента,
                'Сектор_A_опоки': Сектор_A_опоки,
                'Сектор_B_опоки': Сектор_B_опоки,
                'Сектор_C_опоки': Сектор_C_опоки,
                'Сектор_D_опоки': Сектор_D_опоки,
                'Плавка_время_прогрева_ковша_A': Плавка_время_прогрева_ковша_A,
                'Плавка_время_перемещения_A': Плавка_время_перемещения_A,
                'Плавка_время_заливки_A': Плавка_время_заливки_A,
                'Плавка_температура_заливки_A': int(Плавка_температура_заливки_A),
                'Плавка_время_прогрева_ковша_B': Плавка_время_прогрева_ковша_B,
                'Плавка_время_перемещения_B': Плавка_время_перемещения_B,
                'Плавка_время_заливки_B': Плавка_время_заливки_B,
                'Плавка_температура_заливки_B': int(Плавка_температура_заливки_B),
                'Плавка_время_прогрева_ковша_C': Плавка_время_прогрева_ковша_C,
                'Плавка_время_перемещения_C': Плавка_время_перемещения_C,
                'Плавка_время_заливки_C': Плавка_время_заливки_C,
                'Плавка_температура_заливки_C': int(Плавка_температура_заливки_C),
                'Плавка_время_прогрева_ковша_D': Плавка_время_прогрева_ковша_D,
                'Плавка_время_перемещения_D': Плавка_время_перемещения_D,
                'Плавка_время_заливки_D': Плавка_время_заливки_D,
                'Плавка_температура_заливки_D': int(Плавка_температура_заливки_D),
                'Комментарий': Комментарий,
                'Плавка_время_заливки': Плавка_время_заливки
            }
            
            # Сохраняем в SQLite
            conn = sqlite3.connect(DB_FILE)
            columns = ','.join(new_row.keys())
            placeholders = ','.join(['?'] * len(new_row))
            sql = f"INSERT INTO {TABLE_NAME} ({columns}) VALUES ({placeholders})"
            conn.execute(sql, list(new_row.values()))
            conn.commit()
            conn.close()
            
            # Сохраняем в Excel
            add_to_excel(new_row)
            
            # Очищаем кэш данных и перезагружаем страницу
            st.cache_data.clear()
            st.success("Запись успешно добавлена!")
            
            # Добавляем кнопку для просмотра добавленной записи
            if st.button("Показать все записи"):
                menu = "Таблица"
                st.rerun()

elif menu == "Экспорт в Excel":
    st.subheader("Экспортировать все данные в Excel")
    excel_data = save_to_excel(df)
    st.download_button(
        label="Скачать Excel-файл",
        data=excel_data,
        file_name="plavka_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

elif menu == "Аналитика и статистика":
    st.subheader("Аналитика и статистика журнала")
    if df.empty:
        st.info("Нет данных для анализа.")
    else:
        tab1, tab2, tab3, tab4 = st.tabs([
            "Температуры по секторам",
            "Плавки по датам",
            "Наименования отливок",
            "Участники"
        ])
        with tab1:
            st.markdown("**Распределение температур заливки по секторам**")
            for sector in ['A', 'B', 'C', 'D']:
                col = f'Плавка_температура_заливки_{sector}'
                if col in df.columns:
                    fig = px.histogram(df, x=col, nbins=20, title=f"Температура заливки сектор {sector}")
                    st.plotly_chart(fig, use_container_width=True)
        with tab2:
            st.markdown("**Количество плавок по датам**")
            if 'Плавка_дата' in df.columns:
                df['Плавка_дата'] = pd.to_datetime(df['Плавка_дата'], errors='coerce')
                date_counts = df['Плавка_дата'].value_counts().sort_index()
                fig = px.bar(x=date_counts.index, y=date_counts.values, labels={'x': 'Дата', 'y': 'Количество плавок'}, title="Плавки по датам")
                st.plotly_chart(fig, use_container_width=True)
        with tab3:
            st.markdown("**Распределение по наименованиям отливок**")
            if 'Наименование_отливки' in df.columns:
                name_counts = df['Наименование_отливки'].value_counts()
                fig = px.pie(values=name_counts.values, names=name_counts.index, title="Наименования отливок")
                st.plotly_chart(fig, use_container_width=True)
        with tab4:
            st.markdown("**Распределение по участникам**")
            participants = []
            for col in ['Старший_смены_плавки', 'Первый_участник_смены_плавки', 'Второй_участник_смены_плавки', 'Третий_участник_смены_плавки', 'Четвертый_участник_смены_плавки']:
                if col in df.columns:
                    participants.extend(df[col].dropna().astype(str).tolist())
            part_series = pd.Series(participants)
            part_counts = part_series.value_counts()
            fig = px.bar(x=part_counts.index, y=part_counts.values, labels={'x': 'Участник', 'y': 'Количество участий'}, title="Участники во всех плавках")
            st.plotly_chart(fig, use_container_width=True)

elif menu == "Импорт из сообщения":
    st.subheader("Импорт из текстового сообщения")
    msg = st.text_area("Вставьте сообщение со статистикой по смене:", height=400)
    if st.button("Импортировать"):
        parsed = parse_import_message(msg)
        added, errors = 0, 0
        for rec in parsed:
            try:
                # Получаем дату из распарсенного сообщения (уже в формате DD.MM.YYYY)
                plavka_date = rec.get('Плавка_дата', '')
                
                # Генерация служебных полей
                uchet = rec.get('Учетный_номер', '')
                
                # Если формат учетного номера правильный, извлекаем из него компоненты для id_plavka
                if re.match(r'[0-9]{2}-[0-9]{3}/[0-9]{2}', uchet):
                    mm, nnn_yy = uchet.split('-')
                    nnn, yy = nnn_yy.split('/')
                    year = int('20' + yy)
                    month = int(mm)
                    id_plavka = f"{year}{month:02d}{nnn}"
                    номер_плавки = f"{month}-{nnn}"
                else:
                    # Если учетного номера нет, генерируем компоненты из даты
                    if plavka_date and re.match(r'[0-9]{2}\.[0-9]{2}\.[0-9]{4}', plavka_date):
                        day, month, year = map(int, plavka_date.split('.'))
                        # Генерируем случайный номер плавки, если его нет
                        nnn = "001"  # Можно заменить на логику из вашего приложения
                        id_plavka = f"{year}{month:02d}{nnn}"
                        номер_плавки = f"{month}-{nnn}"
                    else:
                        id_plavka = ""
                        номер_плавки = ""
                
                new_row = {
                    'id_plavka': id_plavka,
                    'Учетный_номер': uchet,
                    'Плавка_дата': plavka_date,  # Используем дату из сообщения в формате DD.MM.YYYY
                    'Номер_плавки': номер_плавки,
                    'Номер_кластера': rec.get('Номер_кластера', ''),
                    'Старший_смены_плавки': rec.get('Старший_смены_плавки', ''),
                    'Первый_участник_смены_плавки': rec.get('Первый_участник_смены_плавки', ''),
                    'Второй_участник_смены_плавки': rec.get('Второй_участник_смены_плавки', ''),
                    'Третий_участник_смены_плавки': rec.get('Третий_участник_смены_плавки', ''),
                    'Четвертый_участник_смены_плавки': rec.get('Четвертый_участник_смены_плавки', ''),
                    'Наименование_отливки': rec.get('Наименование_отливки', ''),
                    'Тип_эксперемента': rec.get('Тип_эксперемента', ''),
                    'Сектор_A_опоки': rec.get('Сектор_A_опоки', ''),
                    'Сектор_B_опоки': rec.get('Сектор_B_опоки', ''),
                    'Сектор_C_опоки': rec.get('Сектор_C_опоки', ''),
                    'Сектор_D_опоки': rec.get('Сектор_D_опоки', ''),
                    'Плавка_температура_заливки_A': rec.get('Плавка_температура_заливки_A', None),
                    'Плавка_температура_заливки_B': rec.get('Плавка_температура_заливки_B', None),
                    'Плавка_температура_заливки_C': rec.get('Плавка_температура_заливки_C', None),
                    'Плавка_температура_заливки_D': rec.get('Плавка_температура_заливки_D', None),
                    'Плавка_время_заливки_A': rec.get('Плавка_время_заливки_A', ''),
                    'Плавка_время_заливки_B': rec.get('Плавка_время_заливки_B', ''),
                    'Плавка_время_заливки_C': rec.get('Плавка_время_заливки_C', ''),
                    'Плавка_время_заливки_D': rec.get('Плавка_время_заливки_D', ''),
                    'Плавка_время_заливки': rec.get('Плавка_время_заливки_A', ''),
                    'Маршрутная_карта': rec.get('Маршрутная_карта', ''),
                }
                
                # Сохраняем в SQLite
                conn = sqlite3.connect(DB_FILE)
                columns = ','.join(new_row.keys())
                placeholders = ','.join(['?'] * len(new_row))
                sql = f"INSERT INTO {TABLE_NAME} ({columns}) VALUES ({placeholders})"
                conn.execute(sql, list(new_row.values()))
                conn.commit()
                conn.close()
                
                # Сохраняем в Excel
                add_to_excel(new_row)
                
                added += 1
            except Exception as e:
                errors += 1
                st.error(f"Ошибка при добавлении записи с Учетным номером {uchet}: {e}")
        
        st.success(f"Импортировано записей: {added}. Ошибок: {errors}.")
        
        # Очищаем кэш данных
        st.cache_data.clear()
        
        # Добавляем кнопку для просмотра добавленных записей
        if added > 0 and st.button("Показать импортированные записи"):
            menu = "Таблица"
            st.rerun()

elif menu == "О программе":
    st.markdown("""
    ### Электронный журнал плавки
    - Веб-приложение на Streamlit
    - Данные хранятся в SQLite и Excel
    - Возможности: просмотр, фильтрация, экспорт, добавление записей
    
    **Особенности новой версии:**
    - Автоматическое сохранение данных в Excel и SQLite
    - Улучшенный интерфейс
    - Возможность быстрого просмотра добавленных записей
    - Кнопка обновления данных
    """) 