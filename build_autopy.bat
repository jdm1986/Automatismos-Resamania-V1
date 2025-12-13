@echo off
echo ======================================
echo  Compilando AUTOMATISMOS RESAMANIA
echo ======================================

REM Ir a la carpeta del proyecto
cd /d "%~dp0"

REM Crear venv si no existe y activarlo
if not exist ".venv\Scripts\python.exe" (
    echo Creando entorno virtual .venv...
    py -3 -m venv .venv
)
call .venv\Scripts\activate.bat

REM Asegurar dependencias basicas de build
python -m pip install --upgrade pip >nul 2>&1
pip install pyinstaller pandas pillow >nul 2>&1

REM Construir ejecutable unico
pyinstaller ^
  --noconfirm ^
  --onefile ^
  --windowed ^
  --icon "%~dp0favicon.ico" ^
  --add-data "%~dp0logodeveloper.png;." ^
  --add-data "%~dp0config.json;." ^
  --name "AUTOMASMOS_RESAMANIA" ^
  "%~dp0main.py"

echo.
echo Listo. Ejecutable en dist\AUTOMASMOS_RESAMANIA.exe
pause
