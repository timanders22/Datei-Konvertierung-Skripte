# Datei-Konvertierung — Skripte

Achtzehn Skripte für Windows, die einen Dateibestand vor einer Migration
aufräumen: Altformate anheben, Schutz entfernen, Schriften vereinheitlichen,
PDFs durchsuchbar machen, Dateinamen für Google Drive tauglich machen.

Python 3 und Windows PowerShell 5.1. Sie greifen über COM auf die
installierten Office-Anwendungen zu und laufen deshalb nur unter Windows.

---

## Wozu

Ein gewachsener Dateibestand bringt vor einer Migration lauter Kleinigkeiten
mit, die einzeln harmlos sind und zusammen den Umzug anhalten: `.doc` aus dem
Kompatibilitätsmodus, geschützte Arbeitsmappen ohne bekanntes Kennwort,
Umlaute, die als Mojibake in Dateinamen stehen, PDFs ohne Textebene,
Pfadlängen über 260 Zeichen.

Die Skripte arbeiten diese Punkte der Reihe nach ab. Die Nummerierung ist die
empfohlene Reihenfolge, keine Zwangsfolge — jedes Skript läuft für sich.

## Die Reihenfolge

| | Skript | Aufgabe |
|---|---|---|
| 0 | `0_Vorab-Check.py` | Prüft auf dem Zielrechner, ob alle nötigen Programme, Python-Pakete, DLLs, Office-Einstellungen und PATH-Einträge vorhanden sind |
| 1 | `1_temp_dateien_entfernen.py` | Löscht temporäre und überflüssige Dateien, auf Wunsch auch leere Ordner. UNC und lange Pfade |
| 2a–c | `2a_entferne_schutz_word.ps1`, `2b_…excel.ps1`, `2c_…powerpoint.ps1` | Entfernt Dokument-, Schreib- und Abschnittsschutz; hebt `.doc`/`.dot` nach `.docx`/`.dotx` |
| 3a–c | `3a_doc_docx…`, `3b_xls_xlsx…`, `3c_ppt_pptx…` | Holt Dateien aus dem Kompatibilitätsmodus in das aktuelle Format |
| 4a–c | `4a_ersetze_font_in_word.py`, `4b_…excel.py`, `4c_…powerpoint.py` | Ersetzt Schriften bestandsweit, ohne die übrige Formatierung anzufassen |
| 5 | `5_OCR_PDF.py` | Texterkennung für PDFs (Tesseract, Ghostscript), auf Wunsch PDF/A mit veraPDF-Prüfung |
| 6 | `6_Excel_automatische_Berechnung.ps1` | Setzt eine versehentlich auf „manuell" gestellte Berechnung zurück auf automatisch |
| 7 | `7_Dateien_ohne_Makro_finden.ps1` | Wandelt `.docm`/`.xlsm`/`.pptm` in das makrofreie Format — **nur**, wenn wirklich kein VBA-Code enthalten ist |
| 8 | `8_verschieben_auf_Google_Drive.ps1` | Verschiebt Verzeichnisse nach Google Drive (rclone oder Drive for Desktop), mit Pfadlängenprüfung |
| 9 | `9_fehlerhafte_Dateien_finden.ps1` | Findet Office-Dateien, die sich nicht mehr öffnen lassen |
| 10 | `10_dateinamen_bereinigen.py` | Repariert Mojibake und entfernt Zeichen, die Google Drive nicht mag |
| 11 | `11_Typo_Dateinamen_korrigieren.ps1` | Rechtschreibprüfung für Dateinamen (Hunspell), mit interaktiver Bestätigung |

## Gemeinsame Teile

| Datei | Wozu |
|---|---|
| `_gemeinsam.py`, `_gemeinsam.psm1` | Grundfunktionen, die vorher in jedem Skript einzeln standen: Konsolen-Encoding, Einzelinstanz-Sperre, Langpfade, Zeitstempel, Office-Sitzungswarnung, CSV-Maskierung. Die Skripte binden sie **mit Rückfall** ein — fehlt die Datei, arbeiten sie mit ihrer eingebauten Kopie weiter, und die ps2exe-Übersetzung funktioniert unverändert. |
| `pfade.json` | Verzeichnis-Presets und Bezugsquellen an einer Stelle statt in acht Skripten. Fehlt sie, gelten die eingebauten Vorgaben. |
| `pruefe_syntax.ps1` | Parser-Lauf über alle `.ps1`. |
| `pruefe_alles.ps1` | Sammelprüfung über beide Sprachen: Syntax, `py_compile`, optional PSScriptAnalyzer und ruff — plus Musterprüfungen auf die Fallen, die in dieser Sammlung schon einmal aufgetreten sind. **Vor jedem Einsatz laufen lassen.** |
| `migration.jsonl` | Gemeinsames Laufprotokoll über alle Schritte, eine Zeile je Datei und Aktion. Ergänzt die skripteigenen Protokolle, ersetzt sie nicht. |

## Bevor Sie anfangen

**Erst `0_Vorab-Check.py`.** Es sagt Ihnen, was fehlt, bevor ein anderes
Skript auf halber Strecke stehenbleibt.

**Dann `pruefe_alles.ps1`.** Ein reiner Syntaxlauf hätte den
schwerwiegendsten Fehler dieser Sammlung nicht gefunden — der Ausdruck war
syntaktisch einwandfrei und tat trotzdem das Gegenteil des Gemeinten.

**Und dann ein Probelauf.** Die Skripte, die schreiben, haben einen
Trockenlauf; `10_dateinamen_bereinigen.py` etwa `--dry-run`. Auf einem
Bestand, den man nicht selbst angelegt hat, ist der erste Lauf immer ein
Probelauf.

**Sicherung.** Die Skripte arbeiten auf den Dateien selbst. Was sie ändern,
ändern sie richtig — aber eine Sicherung ersetzt das nicht.

## Pfade

In den Skripten stehen **Platzhalter**, keine echten Adressen:
`\\server\dfs`, `\\server\home`, `Q:\`, `R:\`, `G:\Geteilte Ablagen`,
`G:\Meine Ablage`. Tragen Sie dort Ihre eigenen ein. Wo ein Skript einen
Pfad erwartet, nimmt es ihn auch als Argument.

## Was sie bewusst nicht tun

**Kein Skript rät ein Kennwort.** `2a`–`2c` entfernen den Schutz dort, wo
Windows und Office ihn ohne Kennwort freigeben. Eine Datei mit echtem
Benutzerkennwort wird übersprungen und protokolliert — nicht angegriffen.
Dasselbe bei PDFs in `5_OCR_PDF.py`: geprüft wird auf ein *leeres*
Benutzerkennwort, sonst bleibt die Datei, wie sie ist.

**`7` wandelt nur, was wirklich makrofrei ist.** Nicht „sieht so aus" —
gezählt werden die Zeilen VBA-Code. Bei einer einzigen bleibt die Datei
unangetastet.

**Nichts wird stillschweigend zurechtgebogen.** Was nicht ins Muster passt,
steht hinterher im Protokoll.

## Herkunft

Entstanden bei der Vorbereitung einer Dateimigration nach Google Drive. Die
Skripte sind hier in allgemeiner Form abgelegt: ohne Server- und
Freigabenamen, ohne Benutzerkonten, ohne Zugangsdaten.

## Lizenz

MIT, siehe `LICENSE`. Die verwendeten Fremdprogramme — Tesseract,
Ghostscript, veraPDF, rclone, Hunspell — sind **nicht** enthalten und haben
ihre eigenen Lizenzen; die Skripte nennen die Bezugsquellen.
