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
- **Verschlüsselte Speicherung** der Daten (passwortgeschützt, `.ndat`-Datei, siehe unten)
- **Export** als Markdown (`.md`) oder CSV (`.csv`, Excel-kompatibel)
- **Fächer** pro Klasse verwalten (Dropdown-Auswahl in der Leiste)
- **Klassen übertragen** in ein neues Schuljahr (Schüler + Fächer ohne Noten)
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

### Fächer verwalten

- Fach über das Dropdown **Fach** in der Leiste auswählen
- Mit **+** ein neues Fach hinzufügen, mit **−** löschen
- Noten und Klausuren werden pro Fach gespeichert

### Noten eintragen

1. Schuljahr, Halbjahr, Fach, Klasse und Schülerin auswählen
2. Im Tab **Noten** mündliche oder schriftliche Noten eintragen
3. Die Gesamtnote wird automatisch berechnet

### Klasse übertragen

- Klasse auswählen → Button **Übertragen** → Ziel-Schuljahr eingeben
- Schülerinnen und Fächer werden übernommen, Noten zurückgesetzt
- Neues Schuljahr wird bei Bedarf automatisch angelegt

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

## Verschlüsselungsverfahren

Die Daten werden passwortgeschützt in einer `.ndat`-Datei gespeichert. Das Verfahren basiert auf folgenden Schritten:

1. **Serialisierung**: Die Daten werden als JSON (UTF-8) formatiert
2. **Komprimierung**: Der JSON-String wird mit **zlib** komprimiert
3. **Salz generieren**: Ein zufälliger 16-Byte-Salt wird erzeugt (`os.urandom`)
4. **Schlüsselableitung**: Aus Passwort und Salt wird ein 32-Byte-Schlüssel mit **PBKDF2-HMAC-SHA256** (100.000 Iterationen) abgeleitet
5. **Verifikations-Hash**: SHA-256-Hash des abgeleiteten Schlüssels wird gespeichert (zur Überprüfung beim Entschlüsseln)
6. **Verschlüsselung**: Die komprimierten Daten werden per **XOR** mit dem abgeleiteten Schlüssel verschlüsselt
7. **Dateiformat**: `Base64(Salt [16 Byte] + Key-Hash [32 Byte] + XOR-verschlüsselte Daten)`

### Kurz gesagt

| Komponente | Algorithmus |
|------------|-------------|
| Schlüsselableitung | PBKDF2-HMAC-SHA256, 100.000 Iterationen |
| Verschlüsselung | XOR mit abgeleitetem Schlüssel |
| Komprimierung | zlib |
| Verifikation | SHA-256-Hash des Schlüssels |
| Dateiformat | Base64-kodiert |

> **Hinweis**: Dieses Verfahren bietet Passwortschutz und Datenintegrität, ist jedoch nicht als kryptografisch starke Verschlüsselung (z.B. AES) zu verstehen. Für den Einsatz im schulischen Kontext ist es ausreichend.

## Dateien

| Datei | Beschreibung |
|-------|-------------|
| `noten_verwaltung.py` | Hauptprogramm |
| `noten.ndat` | Verschlüsselte Datendatei (wird automatisch erstellt) |
| `notenschluessel.txt` | Referenz-Tabelle der Notenschlüssel |
| `Makefile` | Build-Skript für Windows-EXE |

## Lizenz

Dieses Projekt ist privat genutzt. Siehe Repository für Details.