@echo off
title Журнал плавки - Запуск
echo Запуск журнала плавки...
echo Приложение откроется в браузере автоматически...
echo.
echo Пожалуйста, подождите...

REM Проверяем наличие Python
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    color 0C
    echo ОШИБКА: Python не установлен!
    echo.
    echo Для работы программы необходимо установить Python
    echo Скачайте и установите Python с сайта: python.org
    echo.
    pause
    exit /b 1
)

REM Проверяем Streamlit
streamlit --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Установка необходимых компонентов...
    pip install streamlit pandas plotly openpyxl xlsxwriter >nul 2>&1
    if %ERRORLEVEL% NEQ 0 (
        color 0C
        echo ОШИБКА: Не удалось установить необходимые компоненты!
        echo.
        pause
        exit /b 1
    )
)

start "" streamlit run streamlit_app.py --browser.serverAddress=localhost

exit 