@echo off
title Disney YouTube Calendar Tool
clear

echo =======================================================
echo ✨ ESPERA UN MOMENTO - INICIANDO CALENDAR TOOL ✨
echo =======================================================
echo.

:: 1. Verificar si Python está instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ ERROR: Python no esta instalado en este equipo.
    echo Por favor, instala Python desde la Microsoft Store o python.org antes de continuar.
    echo.
    pause
    exit
}

:: 2. Verificar o crear el Entorno Virtual (.venv) local
if not exist ".venv" (
    echo 📦 Creando entorno virtual local por primera vez...
    python -m venv .venv
    if %errorlevel% neq 0 (
        echo ❌ Error al crear el entorno virtual.
        pause
        exit
    )
)

:: 3. Activar el entorno virtual e instalar/actualizar librerias
echo 🔄 Sincronizando librerias del sistema...
call .venv\Scripts\activate
python -m pip install --upgrade pip >nul
pip install -r requirements.txt --quiet

:: 4. Lanzar la aplicacion de Streamlit
echo.
echo =======================================================
echo 🚀 TODO LISTO. Abriendo la aplicacion en tu navegador...
echo =======================================================
echo.
streamlit run src/ui/streamlit_app.py

pause