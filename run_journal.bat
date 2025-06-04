@echo off
echo ===================================
echo =      Запуск Журнала плавки      =
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

REM Проверяем, установлен ли Streamlit
streamlit --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Streamlit не установлен. Устанавливаем...
    pip install streamlit pandas plotly openpyxl xlsxwriter
    if %ERRORLEVEL% NEQ 0 (
        echo Ошибка при установке пакетов!
        pause
        exit /b 1
    )
)

echo Запускаем журнал плавки...
echo Приложение откроется в браузере...
echo Для завершения работы закройте это окно.
echo.

streamlit run streamlit_app.py --browser.serverAddress=localhost

pause 