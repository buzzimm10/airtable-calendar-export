@echo off
title Disney YouTube Calendar Tool
cls

echo =======================================================
echo INICIANDO CALENDAR TOOL
echo =======================================================
echo.

:: 1. Comprobar si el comando 'python' funciona
python --version >nul 2>&1
if %errorlevel% == 0 (
    set "PYTHON_CMD=python"
    goto python_detectado
)

:: 2. Si no funciona, comprobar si el comando corto 'py' funciona
py --version >nul 2>&1
if %errorlevel% == 0 (
    set "PYTHON_CMD=py"
    goto python_detectado
)

:: 3. Si ninguno funciona, dar error
echo ERROR: Python no esta detectado en este equipo.
echo.
echo Por favor, desinstale Python y vuelvalo a instalar
echo ASEGURANDOSE de marcar la casilla "Add python.exe to PATH"
echo al principio del instalador.
echo.
pause
exit

:python_detectado
echo Python detectado correctamente (Comando: %PYTHON_CMD%)
echo.

:: 4. Verificar o crear el Entorno Virtual (.venv)
if exist ".venv" goto activar_entorno

echo Creando entorno virtual local (.venv) por primera vez...
%PYTHON_CMD% -m venv .venv
if %errorlevel% neq 0 (
    echo Error al crear el entorno virtual.
    pause
    exit
)

:activar_entorno
echo Sincronizando librerias del sistema...
call .venv\Scripts\activate
python -m pip install --upgrade pip --quiet
pip install -r requirements.txt --quiet

echo.
echo =======================================================
echo TODO LISTO. Abriendo la aplicacion en tu navegador...
echo =======================================================
echo.
streamlit run src/ui/streamlit_app.py

pause