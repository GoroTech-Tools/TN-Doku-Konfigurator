# Changelog – TN-Doku-Konfigurator
<!-- markdownlint-disable MD012 -->

Alle bemerkenswerten Änderungen dieses Projekts werden in dieser Datei dokumentiert.

Format basiert auf [Keep a Changelog](https://keepachangelog.com/) und folgt [Semantic Versioning](https://semver.org/).

---

## [Unreleased]

_Keine Änderungen._

---

## [1.2.5] – 2026-09-20

### Added (1.2.5)

- Die Schaltfläche **„Bearbeiten"** öffnet die konfigurierte Teilnehmer-CSV in einer eigenen Microsoft-Excel-Instanz.
- Nach dem Speichern und Schließen der CSV in Excel werden Teilnehmerdaten und Vorschau automatisch neu geladen.
- Die Excel-Erkennung berücksichtigt Benutzer- und Systemregistrierung sowie 32- und 64-Bit-Registryansichten.

### Changed

- Die Versionsnummer wird beim normalen Build automatisch nach Semantic Versioning als Patch-Bump erhöht.
- Patch-Versionen laufen unbegrenzt weiter (z. B. `1.2.9` → `1.2.10`); für fachliche Releases können Minor und Major gezielt gewählt werden.

### Fixed (1.2.5)

- Quiet-Builds werten PyInstaller-Statusmeldungen auf `stderr` nicht mehr fälschlich als Buildfehler.

---

## [1.2.4] – 2026-09-20

### Fixed (1.2.4)

- `python-docx` wird beim PyInstaller-Build vollständig in die Onefile-EXE eingebunden.
- Die portable EXE startet ohne `ModuleNotFoundError: No module named 'docx'`.

### Changed (1.2.4)

- Der Build verwendet bevorzugt die Projekt-`.venv` und prüft die erforderliche Word-Abhängigkeit vor dem Paketieren.
- Release-Dokumentation und ZIP-Beispiele auf das versionierte Format `TN-Doku-Konfigurator_v{version}.zip` aktualisiert.

---

## [1.0.0] – 2026-05-31

### Changed (1.0.0)

- Projektbezeichnung in README und Dokumentation auf `TN-Doku-Konfigurator` vereinheitlicht.
- Kontext präzisiert: Die Teilnehmer-Ablagesysteme gelten für die kaufmännische Qualifizierung von Personen im BFW Weser-Ems.
- Build-/Release-Artefakte auf `TN-Doku-Konfigurator` umgestellt (EXE, ZIP, Spec-Datei und Workflow-Pfade).
- Release-Version auf `1.0.0` gesetzt und Dokumentationsbeispiele aktualisiert.

---

## [3.0.6] – 2026-05-26

### Changed (3.0.6)

- Versionsreferenzen in der Dokumentation auf `3.0.6` aktualisiert (`DOKUMENTATION_ANWENDER`, `DOKUMENTATION_TECHNIK`, `INSTALLATION`).
- Build-/Release-Lauf für `v3.0.6` mit neuem ZIP-Artefakt (`TN-Doku-Konfigurator_3.0.6.zip`) durchgeführt.

### Fixed (3.0.6)

- Markdown-Bereinigung im Build-Prozess gehärtet, damit mehrere aufeinanderfolgende Leerzeilen (MD012) zuverlässig entfernt werden.
- Erzeugung von Release Notes stabilisiert: Inhalte werden vor dem Schreiben normalisiert, um MD012-Restfehler zu vermeiden.

---

## [3.0.5] – 2026-05-26

### Changed (3.0.5)

- Ausgabeziel vereinheitlicht: `Jahrgang XXXX` und `Anwesenheitsliste KFL XXXX.docx` werden beide direkt im Ordner `output` erzeugt.
- GUI vereinfacht: separates Eingabefeld für Anwesenheitslisten-Zielordner entfernt.
- Schaltfläche „Ergebnis-Ordner öffnen“ öffnet nun den konfigurierten Ausgabe-Ordner (`output`) statt eines `Jahrgang XXXX`-Unterordners.
- Ausgabe-Ordner wird bei Bedarf automatisch erstellt.

### Fixed (3.0.5)

- Statischer Analyzer-Hinweis zu `sys._MEIPASS` beseitigt durch robusten Fallback-Zugriff via `getattr(...)` in der Icon-Pfadauflösung.

---

## [3.0.4] – 2026-05-26

### Changed (3.0.4)

- Teilnehmer-Vorschau in der GUI auf maximal 8 sichtbare Einträge begrenzt; weitere Einträge bleiben per Scrollbar erreichbar.
- Doku-Schaltflächen umbenannt zu `Doku_Anwender` und `Doku_Technik`.
- Schaltfläche `CSV laden & Vorschau` entfernt; CSV wird bei Auswahl weiterhin direkt geladen.
- CSV wird zusätzlich beim Verlassen des CSV-Feldes (still) sowie bei `Enter` im Feld geladen.
- Vor dem Start der Verarbeitung wird eine geänderte CSV-Eingabe automatisch neu eingelesen.

---

## [3.0.3] – 2026-05-26

### Added (3.0.3)

- Neuer, separat konfigurierbarer Speicherort für die erzeugte Anwesenheitsliste in der GUI.

### Changed (3.0.3)

- Standardpfade in der GUI werden nun relativ zum EXE-Verzeichnis angezeigt (CSV, Ausgabe, Ablagesystem, Anwesenheitsliste).
- Bei manueller Auswahl über Dateidialoge werden die Pfade als absolute Pfade übernommen.
- Aktionsschaltflächen auf reine Textbeschriftung ohne Symbole umgestellt.
- Schaltfläche „Dokumentation erstellen“ visuell hervorgehoben (fett/farbig).
- GUI-Höhe und Protokollbereich vergrößert, damit mehr Log-Ausgabe sichtbar bleibt.

---

## [3.0.2] – 2026-05-26

### Changed (3.0.2)

- Verzeichnisstruktur konsolidiert: Build-/Setup-Einstieg auf `src/` vereinheitlicht (`src/build.ps1`, `src/setup.ps1`, `src/requirements.txt`).
- Root-`setup.ps1` dient als Kompatibilitäts-Entrypoint und delegiert an `src/setup.ps1`.

---

## [3.0.1] – 2026-05-18

### Added (3.0.1)

- GUI erweitert um Schnellzugriffe auf `DOKUMENTATION_ANWENDER.md` und `DOKUMENTATION_TECHNIK.md`.

### Changed (3.0.1)

- CSV-Vorschaufenster in der GUI auf kompaktere Höhe reduziert.
- Dokumentationsdateien auf konsistente Namen umgestellt (`DOKUMENTATION_ANWENDER.md`, `DOKUMENTATION_TECHNIK.md`).
- Build-/Release-Prozess gehärtet: `docs/` wird verpflichtend in das Distributionsverzeichnis kopiert, auf Pflichtdateien validiert und in ZIP-Releases mit aufgenommen.

### Fixed (3.0.1)

- Wiederkehrende Markdown-Lint-Fehler (u. a. MD060) durch zentrale Projektregelung entschärft.

---

## [3.0.0] – 2026-05-18

### Changed (3.0.0)

- Build-System auf **PyInstaller Onefile** umgestellt: Verteilung erfolgt jetzt mit einer einzelnen `TN-Doku-Konfigurator.exe` im Paketordner.
- `TN-Doku-Konfigurator.spec` auf Onefile-Layout angepasst (kein `COLLECT`-/`_internal`-Ordner mehr im Release-Verzeichnis).
- `build.ps1` für Onefile-Ausgabe überarbeitet (Onefile-EXE wird erzeugt und anschließend in den versionierten Distributionsordner übernommen).

### Docs (3.0.0)

- Installations-, Anwender- und technische Dokumentation auf Version `3.0.0` aktualisiert.

---

## [2.0.4] – 2026-05-18

### Fixed (2.0.4)

- Build-Skript validiert jetzt vor dem Packen explizit das Vorhandensein von `_internal/_tk_data`, damit keine fehlerhaften Tkinter-Pakete veröffentlicht werden.
- Das Setzen des Hidden-Attributs für `_internal` wurde entfernt, um Probleme bei Entpack-/Kopiervorgängen zu vermeiden.

---

## [2.0.3] – 2026-05-18

### Changed (2.0.3)

- Dokumentation sprachlich vereinheitlicht: `INN-tegrativ gGmbH` durch `Bildungseinrichtung` ersetzt.
- Lizenzdatei auf `LICENSE` (MIT) umgestellt und alle Referenzen von `LIZENZ.txt` aktualisiert.
- Build-Prozess und Projektdokumentation auf konsistente Lizenzreferenzen angepasst.

---

## [2.0.2] – 2026-04-26

### Fixed (2.0.2)

- App-Icon in GUI-Titelleiste: Pfadauflösung für `sys._MEIPASS` im eingefrorenen Modus korrigiert; Icon wird nun zuverlässig aus dem temporären PyInstaller-Verzeichnis geladen.

---

## [2.0.1] – 2026-04-26

### Added (2.0.1)

- App-Icon (`app_icon.ico`) in EXE-Icon und GUI-Titelleiste integriert.
- `TN-Doku-Konfigurator.spec` um Icon-Pfad erweitert.

---

## [2.0.0] – 2026-04-26

### Added (2.0.0)

- Vollständige Dokumentation: Anwenderanleitung, Installationsanleitung, FAQ (30+ Einträge), Technische Referenz, Changelog.
- `_internal`-Verzeichnis nach Build als Hidden markiert (saubereres Erscheinungsbild im Dateiexplorer).

---

## [1.0.1] – 2026-04-26

### Fixed

- **Anwesenheitsliste – Datum-Tab korrigiert:** Der Tab-Stop für das Feld „Datum:" wird jetzt dynamisch aus der tatsächlichen Inhaltsbreite berechnet, statt einen festen Wert zu nutzen. Dies verhindert, dass das Datum-Feld außerhalb des Seitenrands liegt.
- **Anwesenheitsliste – leere Vorlage-Zeile entfernt:** Die leere Zeile aus der Word-Vorlage wird vor der Befüllung gelöscht. Nur echte Teilnehmer-Zeilen werden eingefügt.
- **Anwesenheitsliste – 1-Seiten-Fit:** Zeilenhöhen werden automatisch berechnet, sodass die Anwesenheitsliste genau auf eine Seite passt. Wird am Ende Platz frei, vergrößert sich die Zeilenhöhe gleichmäßig.

### Technical Details

- **Tab-Stop-Berechnung:** `content_w_twips = int((page_width - left_margin - right_margin) * 20)`
- **Zeilenhöhen-Berechnung:** `row_h_pt = available_pt / num_participants` (Minimum: 14pt)
- **Vorlage-Zeilen-Entfernung:** `while len(table.rows) > 1: table._tbl.remove(table.rows[1]._tr)`

### Build Info

- **Version:** 1.0.1
- **Build-Datum:** 26.04.2026
- **EXE-Größe:** ~7.2 MB
- **ZIP-Größe:** ~35 MB

---

## [1.0.0] – 2026-04-26

### Added (Initial Release)

#### Kernfunktionalität

- **Ordnerstruktur anlegen:** Erstellt einen Jahrgangsordner mit je einem Unterordner pro Teilnehmer (Format: `Name - Maßnahmekürzel`)
- **Ablagestruktur kopieren:** Überträgt vordefinierte Ordner-Strukturen aus `_Listen` in jeden Teilnehmer-Ordner
- **Excel-Dateien vorbereiten:**
  - Benennt Musterdateien (XLSM) pro Teilnehmer um
  - Repariert Tabellenblätter (Sheet „Muster" → Teilnehmername)
  - Bewahrt VBA-Makros mit `openpyxl.load_workbook(..., keep_vba=True)`
- **Anwesenheitsliste erzeugen:** Erstellt Word-Datei (DOCX) mit Kopfzeile und befüllter Tabelle
- **Keine Office-Abhängigkeit:** Läuft ohne installiertes Microsoft Office (nutzt `openpyxl` + `python-docx`)

#### GUI (Tkinter)

- Eingabefelder für CSV-Datei, Ausgabe-Ordner, _Listen-Ordner
- Automatische Erkennung von `Teilnehmer_Beginn.CSV` und `_Listen` im App-Verzeichnis
- CSV-Vorschau in Treeview-Tabelle (Name, Maßnahme, Kürzel)
- Log-Ausgabe mit Fortschrittsanzeige (5 Schritte)
- Fehlerbehandlung mit aussagekräftigen Meldungen
- Threading für responsive GUI (Verarbeitung läuft im Hintergrund)

#### CSV-Import

- Automatische Encoding-Erkennung (Windows-1252, UTF-8-sig, UTF-8, Latin-1)
- Semikolon als Standard-Trennzeichen
- Validierung von Pflicht-Spalten (`Name`, `Maßnahme`, `Maßnahmekürzel`)

#### Build & Deployment

- **PyInstaller-Integration:** Erstellt Onedir-EXE mit allen Abhängigkeiten
- **build.ps1 Automatisierung:** Version-Bump, EXE-Erstellung, ZIP-Paketierung
- **Artifact-Handling:** `_Listen`, `Teilnehmer_Beginn.CSV`, `README.md`, `LICENSE` werden in Dist kopiert
- **Portable Ausführung:** Keine Installation erforderlich, kann auf USB-Stick kopiert werden

#### Projektstruktur

```text
src/
├── main.py             # GUI (Tkinter)
├── core.py             # 5-Schritte-Pipeline
├── csv_reader.py       # CSV-Import
└── build_info.py       # Versionierung

Unterstützungsdateien:
├── build.ps1           # Build-Automatisierung
├── TN-Doku-Konfigurator.spec  # PyInstaller-Config
├── setup.ps1           # Venv & Dependencies Setup
├── requirements.txt    # Python-Pakete
├── README.md           # Schnelleinstieg
├── LICENSE             # Lizenzinformation
└── _Listen/            # Vorlagen (in EXE enthalten)
```

#### Plattform

- Windows 10/11 (64-bit)
- Python 3.x (verpackt in EXE, keine Installation nötig)

#### Dokumentation

- Umfangreiche README.md mit Quickstart
- CSV-Format-Tabelle
- Projektstruktur-Übersicht
- Build-Anleitung für Entwickler

### Technical Specs

- **EXE-Größe:** ~7.2 MB (Onedir mit allen Libs)
- **Speicher:** ~50–100 MB während Verarbeitung
- **Performance:** ~5–10 Sekunden für 26 Teilnehmer
- **Python-Pakete:** openpyxl (3.1.0+), python-docx (1.1.0+), pyinstaller (6.17.0+)

### Known Limitations (v1.0.0)

- Nur Windows (Tkinter hat Plattform-Limits)
- XLSM-Makros werden bewahrt, aber nicht modifiziert
- DOCM (Word-Makros) nicht unterstützt (nur DOCX)
- CSV-Encoding auf UTF-8 / Windows-1252 beschränkt

### Git History

- Initial commit: `bb086ff` (Projektstruktur + Source)
- First build: `994e045` (EXE v1.0.1 – vor Bugfix)
- Bugfix release: `9f10069` (Anwesenheitsliste-Korrekturen)

---

## [Upcoming – v1.1.0]

### Planned Features (nicht in aktueller Version)

- CSV-Vorschau mit Zeilen- / Spaltensortierung
- Option: Jahrgang in Ausgabordner-Namen einfügen
- CLI-Variante (ohne GUI) für Batch-Verarbeitung
- Mehrsprachige UI (EN, DE)
- Template-Validierung (Prüfung auf erforderliche Dateien vor Verarbeitung)
- Detaillierte Error-Logs (optional speichern)

### Experimental (unter Evaluation)

- macOS/Linux-Port (würde Tkinter auf Qt/PyQt wechseln)
- Support für weitere Office-Formate (ODS, PPTX)
- Drag-and-Drop für Dateien/Ordner in der GUI

---

## Versionierungsschema

**TN-Doku-Konfigurator** folgt [Semantic Versioning](https://semver.org/):

- **MAJOR (X.0.0):** Breaking Changes, komplette Neuentwicklung
- **MINOR (0.Y.0):** Neue Features (rückwärts-kompatibel)
- **PATCH (0.0.Z):** Bugfixes, kleinere Verbesserungen

Beispiele:

- `1.0.0` – Initial Release
- `1.0.1` – Bugfixes für v1.0.0
- `1.1.0` – Neue Features
- `2.0.0` – Breaking Changes (z. B. neue API, neue Struktur)

---

## Wie man Änderungen einreicht

1. Fork des Repositories auf GitHub
2. Feature-Branch erstellen: `git checkout -b feature/neue-funktion`
3. Commits mit aussagekräftigen Nachrichten:

   ```text
   fix: Fehler bei CSV-Import behoben
   feat: Neue Option hinzugefügt
   docs: Dokumentation erweitert
   ```

4. Push zum Fork: `git push origin feature/neue-funktion`
5. Pull Request öffnen

---

## Release-Prozess

1. **Version aktualisieren** in `src/build_info.py`
2. **Changelog aktualisieren** (diese Datei)
3. **Build erstellen:**

   ```powershell
   .\src\build.ps1  # Auto-Version-Bump, EXE + ZIP
   ```

4. **Git-Commit:**

   ```powershell
   git add -A
   git commit -m "build: Release v1.0.1 – Anwesenheitsliste-Fixes"
   ```

5. **Git-Tag:**

   ```powershell
   git tag -a v1.0.1 -m "Release v1.0.1"
   git push origin main
   git push origin v1.0.1
   ```

6. **GitHub Release erstellen** (auf github.com):
   - Tag `v1.0.1` auswählen
   - ZIP aus `release/` hochladen
   - Release Notes kopieren

---

## Support & Lizenz

- **Lizenz:** Siehe [LICENSE](./LICENSE)
- **Repository:** [https://github.com/GoroTech-Tools/TN-Doku-Konfigurator](https://github.com/GoroTech-Tools/TN-Doku-Konfigurator)
- **Issues:** [https://github.com/GoroTech-Tools/TN-Doku-Konfigurator/issues](https://github.com/GoroTech-Tools/TN-Doku-Konfigurator/issues)
- **Dokumentation:** Siehe `docs/`-Ordner

---

## Danksagungen

Entwicklung für Bildungseinrichtung.

Dank an alle Tester und Nutzer für Feedback und Verbesserungsvorschläge!
