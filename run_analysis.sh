#!/bin/bash
set -e  # Stoppt das Skript sofort bei einem Fehler

GREEN='\033[0;32m'
RED='\033[0;31m'
NC='\033[0m' 

echo -e "${GREEN}[1/4] System-Check...${NC}"
OS_TYPE="$(uname)"

# Linux-spezifische Reparatur (falls nötig)
if [ "$OS_TYPE" == "Linux" ]; then
    if ! python3 -m venv --help &> /dev/null; then
        echo "Installiere fehlende System-Komponenten..."
        sudo apt update && sudo apt install -y python3-venv python3-pip
    fi
fi

# venv sauber erstellen oder nutzen
if [ ! -d "venv" ]; then
    echo -e "${GREEN}[2/4] Erstelle Umgebung...${NC}"
    python3 -m venv venv
fi

echo -e "${GREEN}[3/4] Installiere Anforderungen...${NC}"
source venv/bin/activate
python3 -m pip install --upgrade pip
# Hier werden alle Module (auch timezonefinder) installiert
pip install -r requirements.txt

echo -e "\n${GREEN}[4/4] Starte Python-Analyse...${NC}"
echo "--------------------------------------------------"

# Führe Python aus und fange Fehler ab
if python3 traewelling_analysis_v6.0.py; then
    echo -e "\n--------------------------------------------------"
    echo -e "${GREEN}Analyse war erfolgreich! Excel wurde erstellt.${NC}"
else
    echo -e "\n--------------------------------------------------"
    echo -e "${RED}FEHLER: Das Python-Programm ist abgestürzt.${NC}"
    echo "Bitte prüfe, ob JSON-Dateien im Ordner liegen."
    exit 1
fi