@echo off
setlocal EnableExtensions
chcp 65001 >nul 2>&1

pushd "%~dp0" >nul 2>&1
if errorlevel 1 (
    echo [ОШИБКА] Не удалось открыть папку приложения: %~dp0
    goto :fail_without_popd
)

if not exist "app.py" (
    echo [ОШИБКА] Рядом с файлом запуска не найден app.py.
    goto :fail
)
if not exist "requirements.txt" (
    echo [ОШИБКА] Рядом с файлом запуска не найден requirements.txt.
    goto :fail
)

set "PYTHON_CMD="
where py >nul 2>&1
if not errorlevel 1 (
    py -3 -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>&1
    if not errorlevel 1 set "PYTHON_CMD=py -3"
)
if not defined PYTHON_CMD (
    where python >nul 2>&1
    if not errorlevel 1 (
        python -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>&1
        if not errorlevel 1 set "PYTHON_CMD=python"
    )
)
if not defined PYTHON_CMD (
    echo [ОШИБКА] Не найден Python 3.10 или новее.
    echo Установите Python с https://www.python.org/downloads/ и включите Add Python to PATH.
    goto :fail
)

set "VENV_DIR=.venv-windows"
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"
if not exist "%VENV_PY%" (
    echo Создаю виртуальное окружение для Windows...
    %PYTHON_CMD% -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo [ОШИБКА] Не удалось создать виртуальное окружение.
        goto :fail
    )
)

"%VENV_PY%" -c "import sys" >nul 2>&1
if errorlevel 1 (
    echo [ОШИБКА] Виртуальное окружение повреждено.
    echo Удалите папку %VENV_DIR% и запустите этот файл снова.
    goto :fail
)

echo Проверяю и устанавливаю зависимости...
"%VENV_PY%" -m pip install --disable-pip-version-check -r "requirements.txt"
if errorlevel 1 (
    echo [ОШИБКА] Не удалось установить зависимости.
    echo Проверьте интернет, настройки прокси и совместимость версии Python.
    goto :fail
)

echo.
echo Приложение запускается. Не закрывайте это окно, пока оно работает.
echo Если браузер не открылся, откройте адрес Local URL из этого окна.
echo.
"%VENV_PY%" -m streamlit run "app.py" --browser.gatherUsageStats=false
set "APP_EXIT_CODE=%ERRORLEVEL%"
if "%APP_EXIT_CODE%"=="0" goto :success
if "%APP_EXIT_CODE%"=="130" goto :success
if "%APP_EXIT_CODE%"=="-1073741510" goto :success

echo [ОШИБКА] Streamlit завершился с кодом %APP_EXIT_CODE%.
goto :fail

:success
popd >nul 2>&1
endlocal
exit /b 0

:fail
popd >nul 2>&1
:fail_without_popd
echo.
echo Нажмите любую клавишу, чтобы закрыть окно...
pause >nul
endlocal
exit /b 1
