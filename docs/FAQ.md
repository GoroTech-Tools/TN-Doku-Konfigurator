<!-- markdownlint-disable MD009 MD012 MD022 MD031 MD032 MD040 -->

# FAQ – Häufig gestellte Fragen

## Allgemeines

### F: Brauche ich Microsoft Office installiert?
**A:** Nein! TN-Doku-Konfigurator funktioniert auch ohne Office. Excel- und Word-Dateien werden direkt in Python bearbeitet.

### F: Läuft die Anwendung auf macOS oder Linux?
**A:** Nein, momentan nur auf Windows (10/11, 64-bit). Ein Cross-Platform-Port ist geplant, aber nicht prioritär.

### F: Ist die Anwendung kostenlos?
**A:** Ja, TN-Doku-Konfigurator ist Open Source und kostenlos. Siehe [LICENSE](./LICENSE).

### F: Für welchen Einsatzbereich sind die Teilnehmer-Ablagesysteme gedacht?
**A:** Die Teilnehmer-Ablagesysteme gelten für die kaufmännische Qualifizierung von Personen im BFW Weser-Ems.

### F: Kann ich die Anwendung weitergeben?
**A:** Ja! Die ZIP-Datei kannst du ohne Probleme an andere Nutzer weitergeben. Sie funktioniert überall.

### F: Wo finde ich die Version?
**A:** Die Versionsnummer steht:
- Im GUI-Fenster (oben im Titel)
- In der README.md
- Im GitHub-Release

---

## Installation & Start

### F: Die EXE startet nicht – was kann ich machen?
**A:** Mögliche Ursachen:
1. **Antivirus blockiert die EXE** – In den Sicherheitseinstellungen freigeben
2. **Falsche Architektur** – Sicherstellen, dass du die 64-bit-Version hast
3. **Beschädigte Datei** – Nochmal herunterladen
4. **System-Dateien** – Windows Update durchführen

### F: Kann ich die Anwendung auf einen USB-Stick kopieren?
**A:** Ja! Einfach den kompletten entpackten Ordner auf den USB-Stick kopieren. Sie läuft überall.

### F: Darf ich die Anwendung umbenennen?
**A:** Ja, der Ordnername ist egal. Nur die EXE-Datei sollte `TN-Doku-Konfigurator.exe` heißen.

### F: Kann ich die Anwendung in ein bestehendes Verzeichnis mit `data/Ablagesystem` entpacken?
**A:** Ja! Falls `data/Ablagesystem` und `data/Teilnehmer_Beginn.CSV` bereits vorhanden sind, werden sie automatisch erkannt.

---

## CSV & Datenformat

### F: Welches Trennzeichen muss ich verwenden?
**A:** **Semikolon (`;`)** – nicht Komma!

Korrekt:
```
Name;Maßnahme;Maßnahmekürzel
```

Falsch:
```
Name,Maßnahme,Maßnahmekürzel
```

### F: Welche Zeichenkodierung muss die CSV haben?
**A:** Vorzugsweise:
- **Windows-1252** (Standard Windows) – Umlaute ä, ö, ü funktionieren
- **UTF-8** – Auch sicher
- **UTF-8 mit BOM** – Auch OK

Nicht empfohlen: ANSI (kann zu Problemen mit Umlauten führen)

### F: Wie erstelle ich eine CSV mit Excel?
**A:** 
1. Daten in Excel eintragen
2. **Datei → Speichern unter**
3. **Format wählen:** „CSV (Trennzeichen-getrennt) (.csv)"
4. **Dateiname:** `Teilnehmer_Beginn.CSV` (anschließend in `data/` ablegen)
5. Wenn Excel fragt „Möchten Sie die Datei im CSV-Format speichern?" → **Ja**
6. **Trennzeichen-Dialog:** Semikolon (`;`) auswählen

### F: Darf es Leerzeilen oder Kommentare in der CSV geben?
**A:** Nein, nur die Datenzeilen (Header + Teilnehmerdaten). Leerzeilen werden ignoriert, können aber zu Fehlern führen.

### F: Was ist der Unterschied zwischen „Maßnahme" und „Maßnahmekürzel"?
**A:** 
- **Maßnahme:** Vollständiger Name, z. B. `KBM 2601`
- **Maßnahmekürzel:** Abkürzung, z. B. `KBM`

Die **letzten 4 Zeichen** von `Maßnahme` werden als Jahrgangs-Suffix verwendet (z. B. `2601` → `Jahrgang 2601`).

### F: Was ist mit Teilnehmern mit Sonderzeichen in Namen?
**A:** Kein Problem! Umlaute (ä, ö, ü), Bindestriche, Leerzeichen funktionieren alle. Nur Schrägstriche (`/`) können problematisch sein (werden durch Unterstriche ersetzt).

---

## Anwendung & Verarbeitung

### F: Wie lange dauert die Verarbeitung?
**A:** Typischerweise:
- **10 Teilnehmer:** ~2 Sekunden
- **26 Teilnehmer:** ~5–10 Sekunden
- **50 Teilnehmer:** ~15–20 Sekunden

Größere Listen dauern länger, aber linear.

### F: Kann ich mehrere Gruppen hintereinander verarbeiten?
**A:** Ja! Nach Abschluss einer Gruppe:
1. Neue CSV eintragen (oder via Button auswählen)
2. **„Dokumentation erstellen"** erneut klicken
3. Ein neuer `Jahrgang XXXX`-Ordner wird erstellt

### F: Was passiert, wenn ein Teilnehmer-Ordner bereits existiert?
**A:** Er wird übersprungen (nicht überschrieben). Falls die Struktur schon existiert, wird nichts doppelt angelegt.

### F: Kann ich die Anwesenheitsliste nachträglich bearbeiten?
**A:** Ja! Die erzeugte `Anwesenheitsliste KFL XXXX.docx` ist eine normale Word-Datei. Du kannst sie öffnen und mit Word bearbeiten, ohne dass die Anwendung das sieht.

### F: Werden Makros in Excel/Word bewahrt?
**A:** 
- **Excel (XLSM):** Ja, Makros werden bewahrt mit `keep_vba=True`
- **Word (DOCX):** Nein, Word-Makros (DOCM) werden nicht unterstützt

---

## Ausgabe & Ergebnis

### F: Wo landen die erstellten Dateien?
**A:** Im Ausgabe-Ordner (standardmäßig `_Ausgabe`):
```
_Ausgabe/
└── Jahrgang 2601/
    ├── Anwesenheitsliste KFL 2601.docx
    ├── Müller, Klaus - KBM/
    ├── Schmidt, Maria - KBM/
    └── …
```

### F: Kann ich den Jahrgangsordner verschieben?
**A:** Ja! Den Ordner `Jahrgang XXXX` einfach mit dem Windows-Explorer verschieben oder kopieren.

### F: Was passiert, wenn ich einen Jahrgangsordner nochmal mit derselben CSV verarbeite?
**A:** 
- Ein neuer `Jahrgang XXXX`-Ordner wird erstellt (derselbe Name)
- Der alte Ordner wird **nicht überschrieben** – stattdessen bekommen neue Ordner einen Zähler: `Jahrgang 2601 (1)`, `Jahrgang 2601 (2)`, etc.

### F: Kann ich Ausgabe-Ordner und Ablagesystem-Ordner auf unterschiedliche Laufwerke legen?
**A:** Ja! Du kannst beliebige Pfade via Button auswählen. Sie müssen nicht im gleichen Verzeichnis sein.

---

## Fehlerbehandlung

### F: Fehler „CSV-Datei nicht gefunden"
**A:** 
1. **Prüfen:** Existiert die Datei `data/Teilnehmer_Beginn.CSV`?
2. **Prüfen:** Ist der Dateiiname exakt korrekt (Groß-/Kleinschreibung, Umlaute)?
3. **Lösungsweg:** Datei manuell via Button auswählen

### F: Fehler „Ablagesystem-Ordner nicht gefunden"
**A:**
1. **Prüfen:** Existiert der Ordner `data/Ablagesystem`?
2. **Prüfen:** Enthält er die erforderlichen Dateien (`Anwesenheitsliste.docx`, Excel-Musterdatei)?
3. **Lösungsweg:** Ordner erstellen und befüllen oder manuell auswählen

### F: Fehler „Maßnahme zu kurz"
**A:** Die Spalte `Maßnahme` muss mindestens 4 Zeichen enthalten (z. B. `KBM 2601`). Überprüfe die CSV.

### F: Fehler „Pflicht-Spalten fehlen"
**A:** Die CSV muss mindestens diese Spalten enthalten:
- `Name`
- `Maßnahme`
- `Maßnahmekürzel`

**Prüfe:** Sind die Spaltenköpfe exakt korrekt (Groß-/Kleinschreibung)?

### F: Fehler beim Lesen von Excel/Word
**A:** Die Vorlage ist wahrscheinlich beschädigt.
1. **Neue Vorlage-Datei besorgen** (von Kollege oder neu erstellen)
2. **In den Ordner `data/Ablagesystem` legen**
3. **Nochmal versuchen**

---

## Wartung & Updates

### F: Wie aktualisiere ich die Anwendung?
**A:** 
1. Neue ZIP-Datei von GitHub herunterladen (neuere Versionsnummer)
2. In einen neuen Ordner entpacken
3. `data/Ablagesystem` und `data/Teilnehmer_Beginn.CSV` aus der alten Version **übernehmen** oder **neu bereitstellen**
4. Die neue EXE starten

Alte und neue Version können nebeneinander existieren.

### F: Gibt es automatische Updates?
**A:** Momentan nein. Du musst die neue ZIP manuell herunterladen. In Zukunft könnte es einen Auto-Updater geben.

### F: Wo melde ich Fehler oder Bugs?
**A:** 
- GitHub Issues: [https://github.com/GoroTech-Tools/TN-Doku-Konfigurator/issues](https://github.com/GoroTech-Tools/TN-Doku-Konfigurator/issues)
- Direkt per E-Mail an die Entwicklung

### F: Kann ich Feature-Requests machen?
**A:** Ja! GitHub Issues oder E-Mail an die Entwicklung. Gute Ideen werden erwogen!

---

## Sicherheit & Datenschutz

### F: Werden meine Daten hochgeladen?
**A:** Nein! Die Anwendung arbeitet vollständig lokal. Es gibt keine Cloud-Verbindung.

### F: Werden Logs gespeichert?
**A:** Nein. Nur das Log-Fenster zeigt den Fortschritt. Nach Schließen der Anwendung ist alles weg.

### F: Ist die Quelle offen?
**A:** Ja! Der Code ist auf GitHub öffentlich verfügbar: [https://github.com/GoroTech-Tools/TN-Doku-Konfigurator](https://github.com/GoroTech-Tools/TN-Doku-Konfigurator)

---

## Spezialfälle

### F: Was ist, wenn zwei Teilnehmer den gleichen Namen haben?
**A:** Die Ordner bekommen automatisch den gleichen Namen, können sich aber in Teilordnern unterscheiden (z. B. unterschiedliche Maßnahme). Falls völlig identisch → Ordner-Zähler werden verwendet.

### F: Kann ich mit verschiedenen Maßnahmekürzel im gleichen Lauf arbeiten?
**A:** Ja! Das Tool erkennt den Suffix aus der **ersten** Zeile der CSV (z. B. `2601`). Alle weiteren Zeilen müssen den **gleichen Suffix** haben (aber andere Kürzel sind OK):

```
Name;Maßnahme;Maßnahmekürzel
Müller, Klaus;KBM 2601;KBM       ← Suffix: 2601
Schmidt, Maria;BFK 2601;BFK       ← Gleicher Suffix: 2601
Bauer, Tim;IBA 2601;IBA           ← Gleicher Suffix: 2601
```

### F: Was, wenn alle Teilnehmer eine unterschiedliche Maßnahme haben?
**A:** Das funktioniert nicht! Der Jahrgangs-Suffix wird aus der **ersten** Zeile extrahiert. Alle anderen Zeilen sollten den gleichen Suffix haben.

**Lösung:** Falls verschiedene Jahrgänge → Mehrere CSV-Dateien und Läufe durchführen.

---

## Technische Fragen (für Entwickler)

### F: Welche Python-Version wird verwendet?
**A:** Python 3.x (standardmäßig 3.11+). Siehe `src/requirements.txt`.

### F: Kann ich eigene Funktionen hinzufügen?
**A:** Ja! Der Code ist modular aufgebaut. Neue Module können in `src/` hinzugefügt werden. Danach ein neuer Build via `.\src\build.ps1`.

### F: Kann ich eine CLI-Version (ohne GUI) haben?
**A:** Ja! Der Code in `core.py` ist GUI-unabhängig und kann einfach direkt aufgerufen werden. Eine CLI-Variante ist machbar.

### F: Wie kann ich für das Projekt beitragen?
**A:** 
1. GitHub-Repo forken
2. Feature-Branch erstellen
3. Änderungen vornehmen
4. Pull Request stellen
5. Review abwarten

---

## Kontakt & Support

**Fragen nicht beantwortet?**
- Siehe [DOKUMENTATION_ANWENDER.md](DOKUMENTATION_ANWENDER.md)
- Oder [DOKUMENTATION_TECHNIK.md](DOKUMENTATION_TECHNIK.md) (für Entwickler)
- GitHub Issues: [https://github.com/GoroTech-Tools/TN-Doku-Konfigurator/issues](https://github.com/GoroTech-Tools/TN-Doku-Konfigurator/issues)
- Entwicklung kontaktieren
