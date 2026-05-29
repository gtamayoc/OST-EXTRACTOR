@echo off
echo =======================================================
echo    Iniciando el Extractor de Correos a Obsidian
echo =======================================================

:: Verificar si Python esta instalado
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Python no esta instalado en el sistema o no esta en la variable PATH.
    echo Por favor, instala Python e intentalo de nuevo.
    pause
    exit /b 1
)

:: Definir la ruta del entorno virtual
set VENV_DIR=%~dp0.venv

:: Crear entorno virtual si no existe
if not exist "%VENV_DIR%" (
    echo [INFO] No se encontro un entorno virtual. Creando uno en %VENV_DIR%...
    python -m venv "%VENV_DIR%"
    if %errorlevel% neq 0 (
        echo [ERROR] No se pudo crear el entorno virtual de Python.
        pause
        exit /b 1
    )
    echo [INFO] Entorno virtual creado con exito.
)

:: Activar el entorno virtual
echo [INFO] Activando el entorno virtual...
call "%VENV_DIR%\Scripts\activate.bat"
if %errorlevel% neq 0 (
    echo [ERROR] No se pudo activar el entorno virtual.
    pause
    exit /b 1
)

:: Instalar dependencias si existe requirements.txt
if exist "%~dp0requirements.txt" (
    echo [INFO] Instalando/Actualizando dependencias desde requirements.txt...
    pip install -r "%~dp0requirements.txt"
    if %errorlevel% neq 0 (
        echo [ERROR] Hubo un problema al instalar las dependencias.
        pause
        exit /b 1
    )
) else (
    echo [WARNING] No se encontro el archivo requirements.txt. Intentando ejecutar de todas formas.
)

:: Ejecutar el script extractor
echo [INFO] Iniciando extractor.py...
echo -------------------------------------------------------
python "%~dp0extractor.py"
echo -------------------------------------------------------

echo [INFO] Proceso finalizado.
pause
