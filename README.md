# PDF-Stamper

Desktop-App in Python mit PySide6 und PyMuPDF für das automatische Stempeln von PDFs.

## Funktionen

- Formularfelder aus einem Stempel-PDF auslesen
- Formular einmalig befüllen
- befülltes Stempel-PDF speichern
- Ziel-PDFs stapelweise verarbeiten
- freie Fläche automatisch per Raster + Pixelanalyse suchen
- Stempel auf die Zielseiten platzieren

## Installation

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

## Technische Logik

1. Zielseite wird gerendert
2. aus dem Bild wird eine Belegungsmaske erzeugt
3. zusätzlich werden Text-, Bild- und Zeichnungsrechtecke aus dem PDF gelesen
4. alle möglichen Stempelpositionen werden rasterweise geprüft
5. die beste freie Position wird gewählt
6. der Stempel wird als PDF nativ eingebettet

## Bekannte Grenzen

- Formular-Flattening ist in PDFs nicht immer 100 % einheitlich
- bei stark gefüllten Plänen kann keine freie Position gefunden werden
- noch kein manueller Review-Modus
- noch keine visuelle Vorschau

## Sinnvolle nächste Ausbaustufen

- Vorschaufenster mit markierter Platzierung
- Titelblock-Erkennung unten rechts
- Regeln pro Planformat
- Ausschlusszonen
- Drag-and-drop
- Export eines CSV-Protokolls
- Verarbeitung in Worker-Thread, damit die GUI währenddessen responsiv bleibt
