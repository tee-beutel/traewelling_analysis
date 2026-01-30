@echo off
SETLOCAL
title Traewelling Analysis

echo 1. Pruefe Python
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 goto NO_PYTHON

if not exist venv (
    echo 2. Erstelle venv
    python -m venv venv
)

echo 3. Installiere Pakete
call venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt

echo.
echo Starte Analyse...
python traewelling_analysis_v6.0.py
goto END

:NO_PYTHON
echo FEHLER: Python wurde nicht gefunden!
echo Bitte installiere Python und setze den Haken bei "Add to PATH".
pause
goto END

:END
echo.
echo Fertig.
pause