import os
import subprocess
import sys
import glob
import shutil

def install_requirements():
    """Устанавливает необходимые зависимости"""
    print("Устанавливаем необходимые пакеты...")
    packages = ["pyinstaller", "streamlit", "pandas", "plotly", "openpyxl", "xlsxwriter"]
    
    for package in packages:
        print(f"Устанавливаем {package}...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        except Exception as e:
            print(f"Ошибка при установке {package}: {e}")
            if package == "pyinstaller":  # PyInstaller обязателен
                return False
    return True

def check_required_files():
    """Проверяет наличие всех необходимых файлов"""
    required_files = ["streamlit_app.py", "launcher.py"]
    missing_files = []
    
    for file in required_files:
        if not os.path.exists(file):
            missing_files.append(file)
    
    if missing_files:
        print(f"Ошибка: следующие необходимые файлы отсутствуют: {', '.join(missing_files)}")
        return False
    
    # Проверяем наличие файлов данных
    data_files = ["plavka.db"]
    missing_data = []
    
    for file in data_files:
        if not os.path.exists(file):
            missing_data.append(file)
    
    if missing_data:
        print(f"Предупреждение: следующие файлы данных отсутствуют: {', '.join(missing_data)}")
        print("EXE-файл будет создан, но для полноценной работы приложения эти файлы нужно будет добавить вручную.")
    
    return True

def ensure_requirements_file():
    """Проверяет и создает файл requirements.txt, если он отсутствует"""
    if not os.path.exists("requirements.txt"):
        print("Создаем файл requirements.txt...")
        with open("requirements.txt", "w") as f:
            f.write("streamlit\npandas\nplotly\nopenpyxl\nxlsxwriter\n")
        print("Файл requirements.txt создан")
    return True

def build_exe():
    """Создаёт EXE-файл с помощью PyInstaller"""
    print("Создаём EXE-файл...")
    
    # Определяем все файлы данных, которые нужно включить
    data_files = []
    
    # Основные файлы приложения
    if os.path.exists("streamlit_app.py"):
        data_files.append("--add-data=streamlit_app.py;.")
    
    # Файлы базы данных
    if os.path.exists("plavka.db"):
        data_files.append("--add-data=plavka.db;.")
    
    # Excel файлы
    for excel_file in glob.glob("*.xlsx"):
        data_files.append(f"--add-data={excel_file};.")
    
    # SQL файлы схемы
    for sql_file in glob.glob("*.sql"):
        data_files.append(f"--add-data={sql_file};.")
    
    # Другие Python файлы
    for py_file in glob.glob("*.py"):
        if py_file not in ["launcher.py", "build_exe.py"]:
            data_files.append(f"--add-data={py_file};.")
    
    # Файл requirements.txt
    if os.path.exists("requirements.txt"):
        data_files.append("--add-data=requirements.txt;.")
    
    # Команда для запуска PyInstaller
    cmd = [
        "pyinstaller",
        "--name=Журнал_плавки",
        "--onefile",
        "--windowed",
        "--icon=NONE",
    ] + data_files + ["launcher.py"]
    
    try:
        subprocess.check_call(cmd)
        print("\nEXE-файл успешно создан!")
        print("Вы найдёте его в папке 'dist'")
        
        # Создаём README в папке dist
        if not os.path.exists("dist"):
            os.makedirs("dist")
        
        with open("dist/README.txt", "w", encoding="utf-8") as f:
            f.write("ЭЛЕКТРОННЫЙ ЖУРНАЛ ПЛАВКИ\n")
            f.write("=======================\n\n")
            f.write("Для запуска программы:\n")
            f.write("1. Дважды щелкните на файле 'Журнал_плавки.exe'\n")
            f.write("2. Приложение откроется в вашем браузере\n\n")
            f.write("При первом запуске может потребоваться некоторое время для распаковки файлов\n")
        
        return True
    except Exception as e:
        print(f"Ошибка при создании EXE-файла: {e}")
        return False

def copy_important_files():
    """Копирует важные файлы в папку dist"""
    if not os.path.exists("dist"):
        return
    
    # Копируем BAT-файл для запуска
    if os.path.exists("run_journal.bat"):
        shutil.copy2("run_journal.bat", "dist/")
    
    # Копируем README.md
    if os.path.exists("README.md"):
        shutil.copy2("README.md", "dist/")

if __name__ == "__main__":
    print("=== Создание EXE-файла для Журнала плавки ===")
    
    if not check_required_files():
        print("Отсутствуют необходимые файлы")
        input("Нажмите Enter для выхода...")
        sys.exit(1)
    
    ensure_requirements_file()
    
    if not install_requirements():
        print("Ошибка при установке зависимостей")
        input("Нажмите Enter для выхода...")
        sys.exit(1)
    
    if not build_exe():
        print("Ошибка при создании EXE-файла")
        input("Нажмите Enter для выхода...")
        sys.exit(1)
    
    copy_important_files()
    
    print("\nПроцесс завершён успешно!")
    print("Все необходимые файлы находятся в папке 'dist'")
    input("Нажмите Enter для выхода...") 