#!/bin/bash

# Farben für die Ausgabe
GREEN='\033[0;32m'
NC='\033[0m' # No Color

echo -e "${GREEN}[1/3] Prüfe Python Installation...${NC}"
if ! command -v python3 &> /dev/null
then
    echo "Fehler: python3 konnte nicht gefunden werden."
    exit
fi

if [ ! -d "venv" ]; then
    echo -e "${GREEN}[2/3] Erstelle virtuelle Umgebung (venv)...${NC}"
    python3 -m venv venv
fi

echo -e "${GREEN}[3/3] Installiere Anforderungen...${NC}"
source venv/bin/activate
pip install -r requirements.txt

echo -e "\n${GREEN}Starte Analyse...${NC}"
echo "--------------------------------------------------"
python3 traewelling_analysis_v6.0.py

echo -e "\nAnalyse beendet."