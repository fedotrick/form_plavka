import os
import subprocess
import sys

def run_streamlit_app():
    # Получаем путь к текущей директории
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Путь к streamlit_app.py
    app_path = os.path.join(current_dir, "streamlit_app.py")
    
    # Проверяем существование файла приложения
    if not os.path.exists(app_path):
        print("Ошибка: Файл streamlit_app.py не найден!")
        input("Нажмите Enter для выхода...")
        sys.exit(1)
    
    # Команда для запуска Streamlit
    cmd = ["streamlit", "run", app_path, "--browser.serverAddress=localhost"]
    
    try:
        # Запускаем Streamlit
        subprocess.run(cmd)
    except Exception as e:
        print(f"Ошибка при запуске приложения: {e}")
        input("Нажмите Enter для выхода...")
        sys.exit(1)

if __name__ == "__main__":
    run_streamlit_app() 