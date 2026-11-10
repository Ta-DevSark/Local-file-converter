@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if errorlevel 1 (
  echo Python launcher introuvable. Installez Python 3.11/3.12 64 bits.
  pause
  exit /b 1
)

py -3.12 -m venv .venv 2>nul
if errorlevel 1 py -3.11 -m venv .venv
if errorlevel 1 (
  echo Impossible de creer l'environnement Python.
  pause
  exit /b 1
)

call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install pymupdf python-docx pyinstaller
if errorlevel 1 (
  echo Installation des dependances echouee.
  pause
  exit /b 1
)

pyinstaller --noconfirm --clean --onefile --windowed --name ConvertisseurLocal --collect-all fitz --collect-all pymupdf app.py
if errorlevel 1 (
  echo Compilation echouee.
  pause
  exit /b 1
)

echo.
echo EXE cree dans dist\ConvertisseurLocal.exe
echo LibreOffice doit rester installe sur ce PC pour la conversion vers PDF.
pause
