@echo off
echo ===================================
echo = Создание EXE-файла для запуска  =
echo = приложения "Журнал плавки"      =
echo ===================================
echo.

REM Проверяем, установлен ли Python
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Ошибка: Python не установлен!
    echo Пожалуйста, установите Python с сайта python.org
    echo.
    pause
    exit /b 1
)

echo Python успешно обнаружен
echo.

REM Запускаем скрипт создания EXE
echo Запускаем процесс создания EXE-файла...
echo Это может занять несколько минут...
echo.

python build_exe.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Произошла ошибка при создании EXE-файла
    echo.
    pause
    exit /b 1
)

echo.
echo Процесс завершен! Файл создан в папке dist/
echo.
pause 