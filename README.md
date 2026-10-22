# Freiformschmieden-Rechner

Vorberechnung für das Freiformschmieden von Aluminiumlegierungen, entwickelt am Institut für Metallformung (IMF) der TU Bergakademie Freiberg. Das Programm hat eine grafische Oberfläche (tkinter) und läuft als eigenständige Windows-exe oder direkt mit Python.

Aktuelle Version: **1.6**

## Funktionen

| Reiter | Inhalt |
|---|---|
| Endlänge nach dem Schmieden | Endlänge, Querschnitt und Verschmiedungsgrad für Rund, Vierkant und Achtkant (mit Kantenradius und Materialverlust) |
| Aufheizzeit im Ofen | Aufheizzeit des Rohblocks bis zur Schmiedetemperatur |
| Stichplan | Stichfolge mit Bissen, Drehlagen, Hüben, Presskraft und Temperaturverlauf (Kern, Oberfläche, Kontaktzone); Fließspannung nach Sellars-Tegart oder Hensel-Spittel |
| QForm-Simulation | Baut den Stichplan über die Python-API von QForm UK 12 als 3D-Projekt auf und rechnet ihn Stich für Stich |

Weitere Funktionen: Projektdateien (`.ffproj`), Export als Excel-Datei, CSV und Hubtabellen für QForm.

## Start

* **exe:** `Freiformschmieden-Rechner.exe` starten, Python wird nicht benötigt. Die exe liegt nicht im Repository; sie entsteht mit `build_exe.bat` (siehe unten) oder wird als Release-Datei bereitgestellt.
* **Python:** Python 3.10 oder neuer von python.org (mit tcl/tk), dann `pip install openpyxl` (für den Excel-Export) und `python freiformschmieden.py`.

## exe bauen

`build_exe.bat` unter Windows ausführen. Das Skript installiert PyInstaller und openpyxl und erzeugt `dist\Freiformschmieden-Rechner.exe`. `qform_runner.py` wird als Datendatei eingebunden (`--add-data "qform_runner.py;."`), weil der Rechner es in jeden Simulationsordner kopiert.

## QForm-Simulation

Voraussetzungen: QForm UK 12 mit Python-API (`<QForm-Verzeichnis>\API\App\Python\QFormAPI.py`) und eine Lizenz mit dem Merkmal *API initial data access*.

* Werkstoff des Werkstücks als `.qdat`-Datei oder Datenbankpfad; Geometrie aus Datei oder aus dem Rohblock erzeugt
* Sättel aus Datei, Standard `Flachsattel_Pahnke.qshape` (300 mm breit, ebene Arbeitsfläche 250 mm)
* Werkstück- und Satteltemperatur, Wiedererwärmung je Stich
* Reibung über die Schmierstoffe der QForm-Datenbank
* Hubtabelle (Blows) und Hubgeschwindigkeit aus dem Stichplan, hydraulische Presse (Standard 10 MN)
* Auswertung je Stich: Länge, Querschnitt in der Mitte, Profil über der Länge, größte Presskraft je Hub

Der Rechner schreibt je Simulation einen Ordner mit `simulation.json`, `qform_runner.py` und zwei Startdateien. Die Berechnung läuft als eigener Prozess und lässt sich auch ohne den Rechner starten:

```
python qform_runner.py simulation.json               alle Stiche aufbauen und rechnen
python qform_runner.py simulation.json --nur-aufbau  nur Stich 1 aufbauen, QForm bleibt offen
```

Nicht nach QForm übertragen werden Verweilzeit unter Last, Kontaktleitwerte, Formänderungswirkungsgrad und Reibungszahl des Rechners; Reibung und Wärmeübergang kommen aus dem gewählten Schmierstoff, die Fließspannung aus dem Werkstoff in QForm.

## Dateien

| Datei | Inhalt |
|---|---|
| `freiformschmieden.py` | Oberfläche und Berechnungen (Endlänge, Aufheizzeit), Projektdateien |
| `stichplan.py` | Stichplan: Geometrie, Bisse, Temperatur und Kraft je Stich |
| `xlsx_export.py` | Excel-Export |
| `qform_export.py` | Hubtabellen für QForm |
| `qform_sim.py` | Simulationskonfiguration aus dem Stichplan, Startdateien |
| `qform_db.py` | Werkstoffe und Schmierstoffe der QForm-Datenbank |
| `qform_runner.py` | Aufbau und Berechnung über die QForm-API (nur Standardbibliothek und QFormAPI.py) |
| `build_exe.bat`, `version_info.txt`, `icon.ico` | Build der exe |
| `docs/` | Dokumentation der Modelle und Annahmen (Stand 1.5) |
| `beispiele/` | Beispielprojekt EN AW-6082 und Excel-Ausgabe |

Änderungen: siehe [CHANGELOG.md](CHANGELOG.md).
