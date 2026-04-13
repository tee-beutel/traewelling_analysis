@echo off
SETLOCAL
title Traewelling Analysis

:START
echo 1. Pruefe Python...
:: Prüft explizit auf Version 3.12 (optional, hier einfach allgemein python)
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 goto INSTALL_PYTHON

:VENV_CHECK
if not exist venv (
    echo 2. Erstelle venv...
    python -m venv venv
)

echo 3. Installiere Pakete...
call venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt

echo.
echo Starte Analyse...
python traewelling_analysis_v6.1.py
goto END

:INSTALL_PYTHON
echo Python wurde nicht gefunden. Versuche Installation via winget...
:: -e stellt sicher, dass das exakte Paket genutzt wird, --id ist eindeutig
winget install -e --id Python.Python.3.12 --scope machine

if %ERRORLEVEL% EQU 0 (
    echo.
    echo [ERFOLG] Python 3.12 wurde installiert.
    echo BITTE SCHLIESSE DIESES FENSTER UND STARTE DAS SKRIPT NEU,
    echo damit die Systemvariablen übernommen werden.
    pause
    exit
) else (
    echo.
    echo [FEHLER] Die automatische Installation ist fehlgeschlagen.
    echo Bitte installiere Python 3.12 manuell von python.org
    pause
    goto END
)

:END
echo.
echo Fertig.
pause