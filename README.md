# Traewelling Analysis v6.0 🚄📊

Willkommen beim ultimativen Tool für Statistik-Liebhaber und Bahn-Nerds! Dieses Skript analysiert deine **Träwelling-Exporte** (JSON-Dateien) und verwandelt sie in detaillierte Excel-Tabellen, CSV-Statistiken und eine Übersicht deiner meistbefahrenen Strecken und genutzten Fahrzeuge. Damit du sofort startklar bist habe ich dir auch meine .json's zur Verfügung gestellt.

---

## 🚀 Schnellstart

Um es dir so einfach wie möglich zu machen, gibt es automatisierte Start-Dateien. Du musst lediglich deine Träwelling-JSON-Exporte in den gleichen Ordner legen.

### Windows

Doppelklicke auf die Datei `run_analysis_win.bat`.

### macOS

Öffne das Terminal im Ordner und gib ein:

```bash
chmod +x run_analysis_mac.sh
./run_analysis.sh

```

---

## 🛠 Wie es funktioniert (Die Technik dahinter)

Damit das Skript stabil läuft, nutzt es zwei Standard-Konzepte der Python-Entwicklung: **Virtual Environments** und **Requirements-Listen**.

### 1. Virtual Environment (`venv`)

Eine "virtuelle Umgebung" ist wie ein isolierter Sandkasten für dein Projekt.

* **Warum?** Normalerweise werden Python-Bibliotheken global auf deinem PC installiert. Das kann zu Konflikten führen, wenn Projekt A eine alte Version von Pandas braucht und Projekt B eine neue.
* **Vorteil:** Das `venv` sorgt dafür, dass alle benötigten Bibliotheken nur in diesem einen Ordner existieren und dein restliches System sauber bleibt.

### 2. `requirements.txt`

Diese Datei ist die "Einkaufsliste" für dein Skript. Sie enthält alle Bibliotheken (z. B. `pandas`, `openpyxl`, `timezonefinder`), die das Skript zum Rechnen und Erstellen der Excel-Dateien benötigt.

* Wenn du die Automatisierungs-Skripte startest, wird automatisch der Befehl `pip install -r requirements.txt` ausgeführt. Dies gleicht die Liste mit deinem `venv` ab und installiert fehlende Teile.

### 3. Die Start-Skripte (`.bat` & `.sh`)

Diese Dateien sind kleine Roboter, die dir die Arbeit abnehmen. Sie führen folgende Schritte nacheinander aus:

1. Prüfen, ob Python installiert ist.
2. Prüfen, ob ein `venv` existiert (wenn nicht, wird eines erstellt).
3. Das `venv` aktivieren.
4. Alle Bibliotheken aus der `requirements.txt` installieren/aktualisieren.
5. Das eigentliche Python-Skript starten.

---

## 📋 Voraussetzungen

* **Python 3.9 oder höher:** Wichtig für das `zoneinfo`-Modul zur korrekten Zeitberechnung.
* **Träwelling-Daten:** Lade deine Check-ins bei [Traewelling.de](https://traewelling.de) als JSON herunter.

---

## 📂 Dateistruktur

Ein sauberer Ordner sollte so aussehen:

| Datei / Ordner | Beschreibung |
| --- | --- |
| `traewelling_analysis_v6.0.py` | Das Hauptgehirn (Python-Skript). |
| `requirements.txt` | Die Liste der benötigten Bibliotheken. |
| `*.json` | Deine Träwelling-Exporte (beliebig viele). |
| `run_analysis.bat` / `.sh` | Deine Startrampe für das Programm. |
| `venv/` | (Wird automatisch erstellt) Die isolierte Python-Umgebung. |

---

## 📊 Was erhältst du am Ende?

Nach dem Durchlauf findest du neue Dateien in deinem Ordner:

* **`Username's_data.xlsx`**: Die "Heilige Gral"-Datei mit Reitern für Betreiber-Statistiken, Pünktlichkeitsquoten, Fahrzeuglisten und Haltestellen-Rankings.
* **`gis_number_export.csv`**: Perfekt für Karten-Fans, die ihre Reisen in Programmen wie QGIS visualisieren wollen.

Weitere Funktionen werden laufend hinzugefügt...
---

> **Achtung:** Das Skript ist sehr hungrig auf Daten. Je mehr JSON-Exporte du im Ordner hast, desto umfassender (aber auch bisschen langsamer) wird die Analyse!

