# Notenverwaltung

Programm zur Verwaltung von Schülerinnennoten mit grafischer Oberfläche (Tkinter). Läuft unter Linux und Windows.

## Funktionen

- **Schuljahre & Klassen** verwalten
- **Schülerinnen** einzeln oder als Liste hinzufügen
- **Mündliche Noten** (konfigurierbare Gewichtung, Standard 60%) und **schriftliche Noten** (40%) eintragen
- **Klausuren** mit Punkte-System (Aufgaben mit Maximalpunkten, Punkte pro Schülerin)
- **Notenschlüssel**:
  - **BG** (Berufliches Gymnasium): Noten 0–15 Punkte
  - **IHK** (Berufsschule): Noten 1–6
- **Verschlüsselte Speicherung** der Daten (passwortgeschützt, `.ndat`-Datei)
- **Export** als Markdown (`.md`) oder CSV (`.csv`, Excel-kompatibel)
- **Gesamtnotenberechnung** pro Halbjahr und Jahresnote
- **Notenschlüssel-CSV** bearbeitbar und auf andere Klassen übertragbar
- Automatische **Migration** alter Markdown-Daten

## Voraussetzungen

- Python 3.8 oder neuer
- Keine zusätzlichen Pakete erforderlich (nur Standardbibliothek)

## Installation & Start

```bash
# Repository klonen
git clone git@github.com:deraal09/noten_verwaltung.git
cd noten_verwaltung

# Programm starten
python3 noten_verwaltung.py
```

Beim ersten Start wird ein Passwort vergeben, mit dem die Daten verschlüsselt gespeichert werden.

## Windows-EXE erstellen

Mit dem mitgelieferten Makefile kann eine Windows-`.exe` erstellt werden (erfordert Wine und Windows Python):

```bash
# Windows-Python unter Wine einrichten (einmalig)
make wine-setup

# EXE erstellen
make exe
```

Die fertige EXE liegt dann in `dist_windows/Notenverwaltung.exe`.

## Verwendung

### Schuljahr & Klasse anlegen

1. Schuljahr über das Dropdown oben auswählen oder mit **+** hinzufügen
2. Klasse mit **Hinzufügen** anlegen und Notenschlüssel (BG oder IHK) wählen

### Schülerinnen eintragen

- **Einzeln**: Button *Hinzufügen* → Nachname und Vorname eingeben
- **Als Liste**: Button *Liste* → Mehrere Schülerinnen pro Zeile im Format `Nachname, Vorname` eingeben (auch aus Excel kopierbar)

### Noten eintragen

1. Schuljahr, Halbjahr, Klasse und Schülerin auswählen
2. Im Tab **Noten** mündliche oder schriftliche Noten eintragen
3. Die Gesamtnote wird automatisch berechnet

### Klausuren anlegen

1. Im Tab **Klausuren** eine neue Klausur hinzufügen (Name, Anzahl Aufgaben, Maximalpunkte)
2. Mit **Punkte bearbeiten** die erreichten Punkte pro Schülerin eintragen
3. Die Klausurnoten werden automatisch aus dem Notenschlüssel abgeleitet

### Notenschlüssel anpassen

- Button **Notenschlüssel** im Klausuren-Tab öffnet den Editor
- Standardwerte für BG oder IHK können geladen werden
- Der angepasste Schlüssel kann auf andere Klassen übertragen werden

### Export

- **Datei → Export Markdown...**: Menschenlesbares Markdown-Dokument
- **Datei → Export CSV...**: Für Excel/LibreOffice Calc

## Dateien

| Datei | Beschreibung |
|-------|-------------|
| `noten_verwaltung.py` | Hauptprogramm |
| `noten.ndat` | Verschlüsselte Datendatei (wird automatisch erstellt) |
| `notenschluessel.txt` | Referenz-Tabelle der Notenschlüssel |
| `Makefile` | Build-Skript für Windows-EXE |

## Lizenz

Dieses Projekt ist privat genutzt. Siehe Repository für Details.