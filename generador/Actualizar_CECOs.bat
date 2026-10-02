@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON=py"
) else (
    where python >nul 2>nul
    if not %errorlevel%==0 (
        echo No se encontro Python instalado.
        echo Instale Python desde https://www.python.org/downloads/
        echo Durante la instalacion marque: Add Python to PATH.
        pause
        exit /b 1
    )
    set "PYTHON=python"
)

%PYTHON% -c "import openpyxl" >nul 2>nul
if not %errorlevel%==0 (
    echo Instalando el componente necesario para leer Excel...
    %PYTHON% -m pip install -r requirements.txt
    if not %errorlevel%==0 (
        echo No se pudo instalar openpyxl.
        pause
        exit /b 1
    )
)

if not exist "entradas" mkdir "entradas"
echo Generando archivos. Este proceso puede tardar varios minutos...
%PYTHON% generar_datos.py
if not %errorlevel%==0 (
    echo.
    echo La generacion termino con error. Revise el mensaje anterior.
    pause
    exit /b 1
)

echo.
echo Archivos listos en la carpeta SALIDA.
pause
