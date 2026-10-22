# Änderungen

## 1.6

Neuer Reiter **QForm-Simulation**: baut den Stichplan über die Python-API von QForm UK 12 als 3D-Projekt auf und rechnet ihn Stich für Stich in einem eigenen Prozess (läuft weiter, auch wenn der Rechner geschlossen wird).

* Werkstoff des Werkstücks als `.qdat`-Datei oder Datenbankpfad; Werkstückgeometrie aus Datei oder aus dem Rohblock erzeugt (rund und vierkant als Grundkörper, Kantenradius und Achtkant als DXF-Extrusion)
* Sättel aus Datei, Standard `Flachsattel_Pahnke.qshape`; Sattelwerkstoff aus der Datenbank
* Werkstück- und Satteltemperatur wie Stichplan oder eigener Wert
* Wiedererwärmung je Stich: Temperatur aus dem vorherigen Stich übernehmen oder auf einen Wert setzen
* Reibung über Schmierstoffe der QForm-Datenbank (Warmschmieden Al und Stahl, Siebel, ohne Reibung, eigener Pfad)
* Hubtabelle (Blows) aus dem Stichplan: Transport bzw. Pause an Luft, Nebenzeit auf dem Sattel, Vorschub, Drehung
* Hydraulische Presse, Standard 10 MN, Hubgeschwindigkeit aus dem Stichplan
* Bissfolge vom Manipulator weg oder zum Manipulator hin; das Ende, an dem die Bissfolge beginnt, ragt eine Bisslänge unter die ebene Arbeitsfläche des Untersattels (aus dem Netz des Sattels bestimmt), die ganze Länge wird geschmiedet
* Vorschub je Biss mit dem Anteil f der Längung zur Greifseite (Standard 0,5), damit die Bisse ohne ungeschmiedete Stege aneinander anschließen: vom Manipulator weg s·(1 + (λ−1)·(1−f)), zum Manipulator hin s·(1 + (λ−1)·f); f = 0 entspricht dem Vorschub des Stichplans
* Auswertung je Stich: Länge, Querschnitt in der Mitte, Profil über der Länge, größte Presskraft je Hub
* Schaltflächen: Dateien speichern, Stich 1 aufbauen (prüfen), Aufbauen und rechnen, Abbrechen; Fortschritt und Protokoll im Reiter
* Simulationsordner auch ohne den Rechner startbar (`simulation.json`, `qform_runner.py`, Startdateien)

Geänderte Standardwerte: Hubgeschwindigkeit 60 mm/s; Sattelbreite 250 mm (ebene Arbeitsfläche des Pahnke-Sattels, 300 mm mit R 25).

Getestet mit QForm UK 12.0.5: Aufbau für Rohblock rund, vierkant, vierkant mit Kantenradius und achtkant, beide Bissfolgen; Berechnung über zwei Stiche mit Wiedererwärmung.

## 1.5

* Nebenzeit je Hub wird in QForm als CoolingInTool übernommen
* Bisslänge wahlweise konstant für alle Stiche
* Dokumentation der Modelle und Annahmen (`docs/`)

## Frühere Versionen

* Stichplan mit Temperatur- und Kraftberechnung
* Projektdateien (`.ffproj`), Excel- und CSV-Export, Hubtabellen für QForm
* Endlänge nach dem Schmieden und Aufheizzeit im Ofen
