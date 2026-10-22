@echo off
REM Erzeugt dist\Freiformschmieden-Rechner.exe
REM Voraussetzung: Python 3.10 oder neuer von python.org (mit tcl/tk, Standardinstallation)
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (set PY=py -3) else (set PY=python)
%PY% -m pip install --upgrade pyinstaller openpyxl
%PY% -m PyInstaller --noconfirm --onefile --windowed --name Freiformschmieden-Rechner --icon icon.ico --version-file version_info.txt --add-data "icon.ico;." --add-data "qform_runner.py;." --hidden-import qform_runner freiformschmieden.py
echo.
echo Fertig: dist\Freiformschmieden-Rechner.exe
pause
