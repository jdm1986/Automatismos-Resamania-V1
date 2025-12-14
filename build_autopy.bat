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
pip install pyinstaller pandas pillow pywin32 >nul 2>&1

REM Construir ejecutable unico
pyinstaller ^
  --noconfirm ^
  --onefile ^
  --windowed ^
  --icon "%~dp0favicon.ico" ^
  --add-data "%~dp0logodeveloper.png;." ^
  --add-data "%~dp0feliz_cumpleanos.png;." ^
  --add-data "%~dp0config.json;." ^
  --hidden-import=win32com ^
  --hidden-import=win32com.client ^
  --hidden-import=pythoncom ^
  --hidden-import=pywintypes ^
  --name "AUTOMATISMOS_RESAMANIA" ^
  "%~dp0main.py"

echo.
echo Listo. Ejecutable en dist\AUTOMATISMOS_RESAMANIA.exe
pause
