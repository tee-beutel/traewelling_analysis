#!/bin/bash
set -e # Stoppt bei jedem Fehler sofort

GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m' 

echo -e "${GREEN}[1/4] Suche passende Python-Version...${NC}"

# Suche nach der besten Python-Version (Priorität auf 3.12+)
if command -v python3.12 &> /dev/null; then
    PYTHON_EXE="python3.12"
elif command -v python3.13 &> /dev/null; then
    PYTHON_EXE="python3.13"
else
    PYTHON_EXE="python3"
fi

PY_VERSION=$($PYTHON_EXE -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")')
echo -e "${YELLOW}Nutze $PYTHON_EXE (Version $PY_VERSION)${NC}"

# Check ob Version >= 3.12 für die f-Strings benötigt wird
if [ "$(echo "$PY_VERSION < 3.12" | bc -l)" -eq 1 ]; then
    echo -e "${RED}Fehler: Deine Python-Version ($PY_VERSION) ist zu alt.${NC}"
    echo "Das Skript benötigt mindestens Python 3.12 für die Datenverarbeitung."
    exit 1
fi

# venv erstellen
if [ ! -d "venv" ]; then
    echo -e "${GREEN}[2/4] Erstelle Umgebung...${NC}"
    $PYTHON_EXE -m venv venv
fi

echo -e "${GREEN}[3/4] Installiere Anforderungen...${NC}"
source venv/bin/activate
# In der venv ist 'python' nun automatisch die richtige Version
python -m pip install --upgrade pip
pip install -r requirements.txt

echo -e "\n${GREEN}[4/4] Starte Analyse...${NC}"
echo "--------------------------------------------------"
python traewelling_analysis_v6.1.py

echo -e "\n--------------------------------------------------"
echo -e "${GREEN}Analyse war erfolgreich! Excel wurde erstellt.${NC}"