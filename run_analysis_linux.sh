#!/bin/bash

# Stoppt bei Fehlern (set -e)
set -e

# Farben für die Ausgabe
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m'

echo -e "${GREEN}[1/4] Suche passende Python-Version...${NC}"

# Suche nach Python (Priorität: 3.13 > 3.12 > python3)
if command -v python3.13 &> /dev/null; then
    PYTHON_EXE="python3.13"
elif command -v python3.12 &> /dev/null; then
    PYTHON_EXE="python3.12"
else
    PYTHON_EXE="python3"
fi

# Prüfung der Version direkt über Python (macht 'bc' überflüssig)
if ! $PYTHON_EXE -c "import sys; exit(0 if sys.version_info >= (3, 12) else 1)" &> /dev/null; then
    echo -e "${RED}Fehler: Deine Python-Version ist zu alt.${NC}"
    echo "Das Skript benötigt mindestens Python 3.12."
    $PYTHON_EXE --version
    exit 1
fi

PY_VERSION=$($PYTHON_EXE -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")')
echo -e "${YELLOW}Nutze $PYTHON_EXE (Version $PY_VERSION)${NC}"

# Virtuelle Umgebung (venv) erstellen
if [ ! -f "venv/bin/pip" ]; then
    echo -e "${GREEN}[2/4] Erstelle virtuelle Umgebung (frisch)...${NC}"
    rm -rf venv  # Falls ein kaputter Rest da ist, weg damit
    $PYTHON_EXE -m venv venv
fi

echo -e "${GREEN}[3/4] Installiere Anforderungen...${NC}"
# In der venv nutzen wir direkt den Pfad zum Python-Binary, das ist in Skripten sicherer als 'source'
./venv/bin/python -m pip install --upgrade pip
./venv/bin/pip install -r requirements.txt

echo -e "\n${GREEN}[4/4] Starte Analyse...${NC}"
echo "--------------------------------------------------"
./venv/bin/python traewelling_analysis_v6.1.py

echo -e "\n--------------------------------------------------"
echo -e "${GREEN}Analyse war erfolgreich! Excel wurde erstellt.${NC}"
