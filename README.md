<!-- markdownlint-disable MD012 -->

# TN-Doku-Konfigurator

Portables Windows-Tool zur automatischen Erstellung von Teilnehmer-Ablagesystemen
für neue Ausbildungsgruppen der Bildungseinrichtung.

Die Teilnehmer-Ablagesysteme gelten für die kaufmännische Qualifizierung von Personen im BFW Weser-Ems.

## Version

1.2.5 (Build: 20.09.2026)

## Funktionen

- **Ordnerstruktur anlegen** – erzeugt je Teilnehmer einen eigenen Ordner
  (`Name - Maßnahmekürzel`) unterhalb eines Jahrgangsordners.
- **Ablagestruktur kopieren** – überträgt vordefinierte Unterordner und Vorlagen
  aus `data/Ablagesystem` in jeden Teilnehmerordner.
- **Excel-Dateien vorbereiten** – benennt die Musterdatei um und passt das
  Tabellenblatt (Sheet „Muster" → Teilnehmername) an; **kein installiertes Office erforderlich**.
- **Anwesenheitsliste erzeugen** – erstellt eine befüllte Word-Datei auf Basis
  der Word-Vorlage; **kein installiertes Office erforderlich**.
- **Portabel** – kein Python, keine Installation nötig (Windows 10/11, 64-bit).
- **Vollständig gebündelt** – Word-Dokumente werden mit `python-docx` verarbeitet;
  die benötigte Bibliothek ist in der EXE enthalten.

## Voraussetzungen

- Windows 10 oder 11 (64-bit)
- Microsoft Office ist **nicht** erforderlich

## Quickstart

1. ZIP-Archiv entpacken.
2. Ordner `data/Ablagesystem` mit den aktuellen Vorlagen befüllen (sofern nicht bereits enthalten).
3. `data/Teilnehmer_Beginn.CSV` mit den Teilnehmerdaten der neuen Gruppe befüllen.
4. `TN-Doku-Konfigurator.exe` starten.
5. CSV-Datei und Ausgabe-Ordner prüfen (werden automatisch vorausgefüllt).
6. Auf **„Dokumentation erstellen"** klicken.
7. Ergebnisse im Ordner `output` prüfen (`Jahrgang XXXX` + `Anwesenheitsliste KFL XXXX.docx`).

**Erste Schritte:** Siehe [docs/INSTALLATION.md](docs/INSTALLATION.md) oder [docs/DOKUMENTATION_ANWENDER.md](docs/DOKUMENTATION_ANWENDER.md)

## CSV-Format

Die Datei `data/Teilnehmer_Beginn.CSV` muss semikolongetrennt sein und mindestens diese drei Spalten enthalten:

| Spalte           | Beschreibung                | Beispiel        |
| ---------------- | --------------------------- | --------------- |
| `Name`           | Nachname, Vorname           | `Müller, Klaus` |
| `Maßnahme`       | Vollständiger Maßnahme-Name | `KBM 2601`      |
| `Maßnahmekürzel` | Kürzel der Maßnahme         | `KBM`           |

Die letzten vier Zeichen von `Maßnahme` werden als Jahrgangs-Suffix verwendet
(z. B. `2601` → `Jahrgang 2601`).

## Dokumentation

Umfangreiche Dokumentation im Ordner `docs/`:

- [docs/DOKUMENTATION_ANWENDER.md](docs/DOKUMENTATION_ANWENDER.md): Schritt-für-Schritt Bedienung
- [docs/INSTALLATION.md](docs/INSTALLATION.md): Installation und Vorbereitung
- [docs/FAQ.md](docs/FAQ.md): Häufig gestellte Fragen und Troubleshooting
- [docs/DOKUMENTATION_TECHNIK.md](docs/DOKUMENTATION_TECHNIK.md): Architektur und Build-System
- [docs/CHANGELOG.md](docs/CHANGELOG.md): Versionsgeschichte und Release-Notes

## Projektstruktur (Quellcode)

```text
TN-Doku-Konfigurator/
├── src/
│   ├── main.py
│   ├── core.py
│   ├── csv_reader.py
│   └── build_info.py
├── docs/
│   ├── README.md
│   ├── DOKUMENTATION_ANWENDER.md
│   ├── INSTALLATION.md
│   ├── FAQ.md
│   ├── DOKUMENTATION_TECHNIK.md
│   └── CHANGELOG.md
├── data/
│   ├── Ablagesystem/
│   └── Teilnehmer_Beginn.CSV
├── build.ps1
├── setup.ps1
└── src/
  ├── TN-Doku-Konfigurator.spec
  ├── build.ps1
  ├── setup.ps1
  └── requirements.txt
```

## Entwicklung & Build

```powershell
# Einmalig: Virtuelle Umgebung anlegen und Pakete installieren
.\src\setup.ps1

# Anwendung direkt starten (Entwicklungsmodus)
.\.venv\Scripts\python.exe src/main.py

# EXE und ZIP erstellen (erhöht automatisch die Patch-Version)
.\src\build.ps1

# Ohne Versionserhöhung (z. B. nur für Tests)
.\src\build.ps1 -NoVersionBump

# Minor- bzw. Major-Version erhöhen
.\src\build.ps1 -VersionIncrement Minor
.\src\build.ps1 -VersionIncrement Major

# Nur EXE, kein ZIP
.\src\build.ps1 -SkipZip
```

## Changelog

Siehe [docs/CHANGELOG.md](docs/CHANGELOG.md) für die vollständige Versionsgeschichte.

## Support & Links

- **GitHub Repository:** [https://github.com/GoroTech-Tools/TN-Doku-Konfigurator](https://github.com/GoroTech-Tools/TN-Doku-Konfigurator)
- **Issues & Feedback:** [https://github.com/GoroTech-Tools/TN-Doku-Konfigurator/issues](https://github.com/GoroTech-Tools/TN-Doku-Konfigurator/issues)
- **Lizenz:** [LICENSE](docs/LICENSE)
- **Dokumentation:** [docs/](docs/)
