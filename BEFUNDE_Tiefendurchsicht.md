# Befunde der Tiefendurchsicht

Mehr-Agenten-Durchsicht vom 14.08.2026. 18 Leser (16 im Hauptlauf, 2 in der
Nachlese), jeder Befund anschliessend von einem unabhaengigen Skeptiker
gegengeprueft, dessen Auftrag das Widerlegen war.

| | |
|---|---|
| Gemeldet | 91 |
| Bestaetigt | 87 |
| Widerlegt | 4 |
| davon kritisch | 5 |
| davon hoch | 33 |
| davon mittel | 36 |
| davon gering | 13 |
| Bereits behoben | 4 |

**Abdeckung vollstaendig.** Der im Hauptlauf ausgefallene Leser
(`5_OCR_PDF.py` ab Zeile 2600) wurde in der Nachlese in zwei Segmenten
nachgeholt; dort kamen 11 weitere Befunde hinzu.

## Zur Entstehung des Berichts

Die Bestaetigungsquote lag bei rund 96 Prozent. Vier Befunde wurden im
Hauptlauf stichprobenartig von Hand nachgeprueft. Beim Abarbeiten wurde
JEDER Befund einzeln am Quelltext verifiziert - das hat sich gelohnt: vier
Markierungen im Bericht waren falsch (siehe Uebergabe, Abschnitt 4c), und
zwei Befunde griffen zu kurz. Als Fehlalarm erwies sich keiner.

## Uebergabe: Korrekturlauf vom 14.08.2026

Alle 91 gemeldeten Befunde sind abgearbeitet (87 bestaetigte korrigiert,
4 widerlegte ohne Aenderung). Jeder Befund wurde vor der Korrektur einzeln
am Quelltext geprueft, ein grosser Teil zusaetzlich ausfuehrend nachgestellt.

| Stufe | Stand |
|---|---|
| KRITISCH | 5 / 5 |
| HOCH | 33 / 33 |
| MITTEL | 36 / 36 |
| GERING | 13 / 13 |
| zusaetzlich gefunden | `_apply_security_descriptor` (7 Skripte) |

Syntaktisch geprueft: alle 12 Python- und 10 PowerShell-Dateien.
Sicherung vor allen Aenderungen: `Claude Skripte_SICHERUNG_2026-08-13_233450`
auf dem Desktop.

---

### 1. Was NICHT getestet ist - vor dem produktiven Einsatz lesen

Die Aenderungen sind umfangreich. Ausfuehrend geprueft wurden Logik,
Dateisystem-, ACL- und Prozessverhalten. **Nicht getestet sind die
COM-Pfade mit Word und PowerPoint** - dort gab es keine Testumgebung.
Betroffen sind vor allem:

- `2a`, `2c` (Word/PowerPoint-Schutzentfernung)
- `3a`, `3c` (Word/PowerPoint-Konvertierung)
- `4a`, `4c` (Schriftartersetzung Word/PowerPoint)
- `9_fehlerhafte_Dateien_finden.ps1` (Smoke-Tests)

Excel-COM war verfuegbar und wurde benutzt (Kennwortverhalten von `SaveAs`,
`FormatCondition.Font`). Word und PowerPoint nicht.

> **Korrektur vom 15.08.2026:** Die Begruendung "keine Testumgebung" war
> falsch. Auf diesem Rechner ist **Microsoft Office Professional Plus 2024
> (Volume) 16.0.17932.20910** lokal installiert, und alle drei COM-Server
> lassen sich starten - nachgemessen: `Word.Application` Build 16.0.17932,
> `Excel.Application` und `PowerPoint.Application` Build 17932.
> **Die COM-Pfade wurden daraufhin getestet - siehe Abschnitt 9.**
> Stand jetzt: `2a`, `3a`, `4a`, `4c` bestanden, `2c` und `3c` mit Befund
> (einer davon behoben), `7` blockiert, `9` weiterhin ungetestet.
>
> **Stand 29.09.2026:** Jetzt sind auch `2b`, `3b`, `4b`, `6`, `7` und `9`
> gegen das echte Office gelaufen, jeweils Probelauf und Echtlauf. Dabei
> kamen neun weitere Fehler ans Licht, alle behoben - siehe Abschnitt 10.
> Am Abend folgte eine Durchsicht mit neun kritischen Pruefagenten
> (rund 95 Meldungen, nachgeprueft, behoben, alle Office-Skripte erneut am
> echten Office getestet) - siehe Abschnitt 11. In der Nacht wurden die
> dort offenen Punkte umgesetzt und gemessen - siehe Abschnitt 12, die
> Rueckmeldung des Anwenders dazu in Abschnitt 13 und 14.

**Empfehlung:** Erster Lauf je Skript mit `-WhatIf` bzw. `--dry-run` auf
einem Testbestand, nicht auf Q:/R:. Die Probelauf-Pfade wurden in dieser
Sitzung mehrfach korrigiert (siehe 3b/3686, 3c/2674) und sind jetzt
belastbar - vorher loeschten sie tatsaechlich Dateien.

---

### 2. Umgebungsfakten, die den Korrekturen zugrunde liegen

Diese Punkte wurden auf diesem Rechner gemessen, nicht angenommen:

- **PowerShell-Edition:** `FileInfo.GetAccessControl` existiert in
  Windows PowerShell 5.1.26100.9168, NICHT in PowerShell 7.6.4 (Core).
  `2a`, `2b`, `2c` und `7` brechen unter `pwsh` jetzt mit Exit-Code 2 ab.
  **Die Skripte muessen mit `powershell.exe` gestartet werden.**
- **`C:\OCR` ist fuer jeden authentifizierten Benutzer beschreibbar**
  (`Authentifizierte Benutzer:(I)(OI)(CI)(M)`, geerbt von der Standard-ACL
  unter `C:\`). Der Ordner enthaelt ausgefuehrte Programme
  (`jbig2.exe`, `verapdf.bat`, `pngquant.exe`) und die geladene
  `leptonica-1.76.0.dll`. `5_OCR_PDF.py` bricht bei erhoehten Rechten jetzt
  ab und nennt den Reparaturbefehl; `0_Vorab-Check.py` haertet den Ordner im
  Admin-Lauf. **Das sollte einmalig von Hand nachgezogen werden:**

      icacls "C:\OCR" /inheritance:r /grant *S-1-5-32-544:(OI)(CI)F /grant *S-1-5-18:(OI)(CI)F /grant *S-1-5-32-545:(OI)(CI)RX

- **ps2exe** liegt in `%USERPROFILE%\Documents\WindowsPowerShell\Modules\ps2exe`
  (1.0.17 und 1.0.18). Der erzeugte Host heisst `PSRunspace-Host`, nicht
  `ConsoleHost` - das war die Ursache von Befund 9/43.
- **Excel-COM:** `SaveAs(..., Password="")` entfernt die Verschluesselung
  (nicht "unveraendert lassen"). `FormatCondition.Font.Name` und `.Size`
  sind nicht setzbar, `Bold`/`Italic`/`Underline`/`Strikethrough` schon.
- **robocopy `/XD <blosser Name>`** schuetzt beim `/MIR`-Purge NUR
  Verzeichnisse direkt in der Wurzel; verschachtelte werden geloescht -
  auch mit Vollpfad-`/XD`. Auf der Kopierseite (`/E`) wirkt es dagegen auf
  jeder Ebene.
- **ocrmypdf `oversample`** ist eine DPI-UNTERgrenze ("to at least the
  specified DPI"), keine Reduktion.

---

### 3. Wo ich bewusst vom Vorschlag abgewichen bin

| Befund | Vorschlag | Stattdessen | Grund |
|---|---|---|---|
| `0/376` | Obergrenze fuer das Warten auf Stage 2 | Lebenszeichen im Minutentakt, kein Abbruch | Stage 1 laeuft auf mittlerer Integritaetsstufe und kann den elevierten Stage-2-Prozess nicht beenden - ein Abbruch wuerde ihn nur verwaisen lassen |
| `8/2392` | `UnauthorizedAccessException` gar nicht wiederholen | `MaxAttempts 2` | Nach dem Abraeumen des Schreibschutzes ist die Hauptursache weg; ein Versuch bleibt fuer kurzlebige AV-Sperren |
| `2b/614` | Backups mit `.bak_<PID>` kennzeichnen | `CreationTime` + `SetCreationTime` nach dem Kopieren | Der PID-Umbau beruehrt Rollback, Orphan-Regex und `Get-OriginalFromBackupPath`; die CreationTime loest den Parallelfall bereits |
| `5_OCR/1091` | `ocrmypdf.` aus der Whitelist streichen | Altersbedingung (vor Laufbeginn UND > 1 h) | Reste abgebrochener eigener Laeufe sind GB-gross und gehoeren geraeumt; nur aktive Fremdprozesse muessen verschont bleiben |
| `4b/2124` | Funktion entfernen ODER dokumentieren | Vorhandensein melden statt vergeblich setzen, plus Hinweis in der Umfangsanzeige | Der stille Fehlschlag liess das Blatt als vollstaendig umgestellt gelten |
| `4b/3231` | Aufraeumen ODER Wiederholung ausnehmen | Aufraeumen | So bleibt die Wiederholung erhalten, die bei voruebergehenden Fehlern hilft |
| `2b/1025` | 5.1 erzwingen ODER editionsneutral | 5.1 erzwingen | `FileSystemAclExtensions` ist in BEIDEN Editionen nicht geladen, `Get-Acl` haette das Langpfad-Problem |

---

### 4. Ueber den Bericht hinaus gefunden

**a) `_apply_security_descriptor` konnte den Zugriff ENTZIEHEN.**
Beim Verifizieren von `5_OCR/2036` aufgefallen. Nicht die DACL war die
Ursache, sondern das Setzen des EIGENTUEMERS: danach verlieren die geerbten
ACEs ihr `(I)`-Kennzeichen und eine geerbte `EIGENTUEMERRECHTE`-ACE
(S-1-3-4) bekommt `(IO)` - sie gilt dann fuer die Datei selbst nicht mehr.
Betrifft alle sieben schreibenden Skripte (3a, 3b, 3c, 4a, 4b, 4c,
5_OCR_PDF) in logisch identischer Form. Behoben: Owner, Gruppe und DACL in
EINEM `SetNamedSecurityInfo`-Aufruf, Eigentuemer nur bei Abweichung.
Eigener Berichtseintrag vor dem MITTEL-Block.

**b) Der Google-Drive-Befund (`8/2528`) griff zu kurz.** Der `/MIR`-Purge
selbst zerstoert verschachtelte ausgeschlossene Verzeichnisse. Beides
behoben; Einzelheiten beim Befund.

**c) Vier Befunde waren im Bericht falsch markiert - in beide Richtungen:**

- `5_OCR/5145`, `3c/2592`, `5_OCR/2020`: bereits behoben, aber nicht als
  solche gekennzeichnet.
- `2b/1867`: als behoben markiert, aber nur zur Haelfte erledigt (nur
  `New-Item`, nicht die `Remove-Item`-Aufraeumpfade).

Den `[BEHOBEN]`-Markierungen war also nicht zu trauen. Jeder Befund wurde
deshalb am Quelltext geprueft.

**d) Dieselben Fehler steckten haeufig in Schwesterskripten, ohne gemeldet
zu sein.** Mitkorrigiert wurden unter anderem:

- Passwort-Array (`(,$Passwords)`): auch in `2b` (dort HAERTER - ohne
  `[string]`-Cast entsteht eine ArrayList und der COM-Aufruf wirft)
- `-WhatIf`-Infrastruktur: auch in `2c` (dort zwingend, sonst waere das neu
  eingezogene `Confirm-Write` nie erreicht worden)
- PowerShell-Editionspruefung: auch in `2a`, `2c`, `7`
- `input()` ohne Konsole: auch in `3a`, `4c` und je vier weitere Stellen
- Probelauf loescht Dateien: auch in `3c`
- Verwaisten-Backup-Regeln: `2c` hatte BEIDE Fehler (Groesse und
  `LastWriteTime`)

---

### 5. Wiederkehrende Fehlermuster (fuer kuenftige Aenderungen)

1. **`except Exception` faengt kein `SystemExit`.** Der Signal-Handler
   beendet sich mit `sys.exit()`; Aufraeum- und Rollback-Zweige liefen
   deshalb bei Strg+C nicht. Betraf `4b`, `4c`, `3a`, `3b`, `3c`, `4a`.
   Beim Umstellen auf `except BaseException` muss der Abbruch nach der
   Rettung **weitergereicht** werden (`if not isinstance(e, Exception): raise`),
   sonst wird Strg+C verschluckt.
2. **`-WhatIf` gilt fuer den ganzen Skript-Scope.** Jedes
   ShouldProcess-faehige Cmdlet gehorcht ihm - auch `New-Item`,
   `Set-Content`, `Add-Content` und `Remove-Item` auf eigene Temp-Dateien.
   Infrastruktur gehoert provider-frei (`[System.IO.Directory]::CreateDirectory`)
   oder mit `-WhatIf:$false`.
3. **Namen vor ihrer Zuweisung benutzt.** `detail_logger` in frueh
   laufenden `except`-Zweigen: der Fehlerschlucker warf selbst einen
   `NameError`. Betraf `5_OCR`, `3b`, `3c`.
4. **Vorlaeufige Zustaende als dauerhaft vermerkt.** Gesperrte oder
   voruebergehend unzugaengliche Dateien landeten in der Resume-Liste und
   wurden fuer immer uebersprungen. Betraf `10`, `1_temp`, `3a`, `3c`.
5. **Probelauf und Echtlauf trafen unterschiedliche Entscheidungen.**
   Betraf `3b`, `3c`, `4b`, `5_OCR` - teils mit Loeschungen im Probelauf.

Ergaenzt am 29.09.2026 (Belege in Abschnitt 10):

6. **Cmdlet-Ausgaben nie ungewandelt an COM geben.** `Join-Path`,
   `Split-Path` & Co. liefern in `PSObject` verpackte Strings. Word
   `SaveAs2` blieb damit unbegrenzt haengen, `Documents.Open` vertrug es
   zufaellig. Pfade fuer Office immer als `[string](...)` uebergeben.
   Ausgaben von Skriptfunktionen sind nicht verpackt (gemessen).
7. **Muster 2 gilt auch fuer das Aufraeumen.** Unter `-WhatIf` wurden
   `Stop-Process` auf die EIGENEN Office-Instanzen, `Stop-Job`/`Remove-Job`
   und `Unblock-File` auf Temp-Kopien nur simuliert - Office-Prozesse und
   Jobs blieben stehen, in `7` fehlte sogar der Arbeitsordner. Jeder Aufruf,
   der nur skripteigene Dinge anfasst, braucht `-WhatIf:$false`. Code in
   `Start-Job`-Bloecken und eigenen Runspaces ist nicht betroffen.
8. **Excel-Kennwoerter:** `WriteResPassword` hat hoechstens 15 Zeichen -
   ein laengerer Wert laesst `Workbooks.Open` fuer JEDE Mappe scheitern.
   Ein leerer String gilt als "nicht angegeben" und oeffnet den
   Kennwortdialog. `Worksheet.Unprotect()` ohne Argument oeffnet ebenfalls
   einen Dialog; `Unprotect("")` wirft sofort. Word `Unprotect()` wirft
   dagegen sofort (gemessen).
9. **Dateien lesen, die Office offen haelt:** nur mit
   `FileShare.ReadWrite`. `ZipFile.OpenRead` teilt nur Lesezugriff und
   scheitert. `Add-Type System.IO.Compression.FileSystem` laedt
   `System.IO.Compression` (mit `ZipArchive`) NICHT mit.
10. **Kennwortdialoge unsichtbarer Office-Instanzen koennen auf dem
    Bildschirm erscheinen.** Wer dort klickt, verfaelscht Testergebnisse
    (siehe Abschnitt 10, Testbedingungen).

---

### 6. Offene Punkte

- **Nicht abschliessend geklaert:** Die Endung `.syd` in
  `1_temp_dateien_entfernen.py` (`_SUFFIX_RULES_BASE`). Ich konnte nicht
  belastbar feststellen, ob sie ein Temp-Artefakt oder ein Datenformat
  bezeichnet, und habe sie deshalb unveraendert gelassen. Falls die
  Herkunft des Eintrags nicht bekannt ist, gehoert sie vorsichtshalber in
  `_OPTIONAL_DEV_SUFFIXES` (dorthin sind `.sdf`, `.class`, `.user`, `.prv`
  bereits verschoben).
- **`C:\OCR`-ACL** einmalig von Hand haerten (Befehl siehe Abschnitt 2).
- **`C:\tmp_skripte` und `C:\tmp_skripte\bin`** haben dieselbe offene ACL.
  `bin` ist derzeit leer und wird von keinem Skript benutzt; `tmp_skripte`
  ist Arbeits- und Logordner und muss beschreibbar bleiben. Kein
  Handlungsbedarf, aber bekannt.
- **`.claude/settings.json`** wurde angelegt: `Desktop\Claude Skripte` als
  zusaetzliches Arbeitsverzeichnis plus zwoelf lesende Freigaben. Die
  bestehende `settings.local.json` (32 KB, fast nur exakte Einzelbefehle)
  wurde NICHT angefasst - sie ist der Grund, warum Rueckfragen trotz
  Allowlist wiederkehren, und koennte auf wenige generelle Muster
  eingedampft werden.

### 7. Nachtrag 14.08.2026: zweite externe Durchsicht

Ein weiterer Programmierer hat die Skripte durchgesehen. Von 16 Punkten
haben **vier** der Nachpruefung standgehalten und sind behoben, zwoelf nicht.

**Bestaetigt und behoben:**

- **Wettlauf im COM-Zeitwaechter** (`3a`, `3b`, `3c`, `4a`). `return call_fn()`
  berechnet den Rueckgabewert, ERST DANACH setzt das finally das
  done_event. Laeuft der Timeout dazwischen ab, toetet der Waechter Office,
  obwohl der Aufruf gelungen ist - der Aufrufer bekommt ein Ergebnis und
  arbeitet mit einer toten Instanz weiter. Nachgemessen an einer Nachbildung:
  43 von 300 Laeufen widerspruechlich.
  **Wichtig:** Der vorgeschlagene Fix (nur ein Schloss) haette NICHT gereicht -
  gemessen 30 von 300. Noetig ist zusaetzlich die unteilbare Pruefung auf dem
  Erfolgspfad: entweder der Waechter gewinnt (dann TimeoutError) oder der
  Aufruf (dann kein Kill). Mit beidem: 0 von 300. An der echten Funktion aus
  `3a` gegengeprueft (200 Laeufe am Timeout-Rand, kein Widerspruch).
  `4b/_run_com_with_watchdog` war bereits korrekt (nutzt `done.is_set()` unter
  dem Schloss) und blieb unveraendert.
- **`BrokenProcessPool` ohne `_terminate_pool_workers`** (`5_OCR_PDF`, beide
  Zweige). Bei einem Teil-Crash blieben ueberlebende Worker stehen und
  hielten das Arbeitsverzeichnis mit den unkomprimierten Seitenbildern.
- **`Read-Host` ohne Konsolenpruefung** (`_gemeinsam.psm1` und die vier
  lokalen Kopien in `2a`, `2b`, `2c`, `7`). Dabei zusaetzlich gefunden:
  `2a` und `2c` riefen `Show-OfficeRunningWarning` OHNE `-Silent` auf -
  anders als `2b` und `7`. Beides korrigiert.
- **Uneinheitliche Protokollstufe** in `Get-FileSecuritySnapshot`. Das war
  eine Folge der Korrektur zu `2b/1025` in dieser Sitzung: `2b` wurde auf
  WARN gezogen, `2a` und `2c` blieben auf DEBUG. Jetzt einheitlich WARN.

**Nicht bestaetigt** (jeweils am Quelltext geprueft): doppelte Deklaration von
`Add-LongPathPrefix` (`2c` hat genau eine; kein PS-Skript hat Doppel-
deklarationen), doppelte Deklaration von `_repair_mojibake` (eine Funktion
plus die bewusst getrennte `_repair_mojibake_core`), `UnboundLocalError` im
finally von `4c` (Initialisierung Z. 2097 liegt vor dem try Z. 2138),
Rekursionstiefe bei Diagrammen in `4b` (`_set_chart_fonts` ruft sich nicht
selbst auf), `$script:InteractiveLaunch` in `11` (Z. 1370 gesetzt, Z. 1401
benutzt), `_resolve_unique_path` in `3b` (die Funktion existiert dort nicht;
`3b` nutzt `_resolve_collision`), und `$LASTEXITCODE` hinter einer
robocopy-Pipeline - mit echtem robocopy in drei Varianten geprueft, darunter
mit Ausnahme im Scriptblock: der Exit-Code kommt in allen Faellen korrekt an.

**Berechtigt, aber ohne Laufzeitwirkung:** Die Typannotation von
`walk_and_clean_junk` in `1_temp_dateien_entfernen.py` lautet
`Generator[Tuple[str, str], ...]`, geliefert werden Dreier-Tupel. Korrigiert.
Die abweichende Parameterreihenfolge des Zeitwaechters zwischen `3a` und `4a`
ist real, aber beide Aufrufseiten passen jeweils zu ihrer Definition.

**Nicht geprueft** (Zeitgruende, geringe Schwere): `_MRU_AVAILABLE`-Meldung in
`3c`, `$script:ScriptCmdlet` bei `-WhatIf` in `6`, TOCTOU in
`Get-UniqueTargetPath` in `7`, Preflight bei nicht unterstuetzten Endungen
in `9`. **Inzwischen abgearbeitet - siehe Abschnitt 8.** Die Einschaetzung
"geringe Schwere" war fuer zwei der vier Punkte falsch.

---

### 8. Nachtrag 15.08.2026: die vier offenen Punkte aus Abschnitt 7

Alle vier abgearbeitet. Zwei hatten Laufzeitwirkung, zwei waren stille
Fallen ohne heutige Auswirkung. Die Einstufung "geringe Schwere" aus
Abschnitt 7 hielt bei `6` und `7` nicht stand.

**a) `-WhatIf` in `6_Excel_automatische_Berechnung.ps1` war vollstaendig
kaputt.** Nicht `$script:ScriptCmdlet` war die Ursache, sondern erneut
Fehlermuster 2 aus Abschnitt 5: `New-Item` legt den Tempordner an, gehorcht
aber `-WhatIf` und tat es deshalb nicht. Die Arbeitskopie lief danach bei
JEDER Datei auf `DirectoryNotFoundException`. Nachgestellt an einer xlsx mit
`calcMode="manual"`: 1 Datei gefunden, 0 Befunde, 1 Fehler - der dokumentierte
Probelauf (`.SYNOPSIS`, Zeile 18) meldete also nichts ausser Fehlern.
Zusaetzlich stand die Kopfzeile auf "LIVE" samt Warnung "Aenderungen werden
direkt geschrieben", weil `$ReadOnly` nur `-ReadOnlyMode` auswertete, nicht
`-WhatIf`. Behoben: Tempordner und Log-Verzeichnis provider-frei ueber
`[System.IO.Directory]::CreateDirectory`, das Aufraeumen des eigenen leeren
Tempordners mit `-WhatIf:$false`, `$ReadOnly` beruecksichtigt jetzt
`$WhatIfPreference`. Nachher gemessen: 1 gefunden, "Wuerden korrigiert
werden: 1", 0 Fehler, Datei-Hash unveraendert, keine Ordnerreste. Echtlauf
gegengeprueft: `calcMode`/`calcOnSave` entfernt, 0 Fehler.

Am gemeldeten Punkt selbst war trotzdem etwas dran: `Confirm-Write` fiel bei
fehlendem `$PSCmdlet` oder werfendem `ShouldProcess` hart auf `$true` zurueck -
also auf Schreiben, trotz `-WhatIf`. Jetzt Rueckfall ueber
`$WhatIfPreference`; dass die Variable im Skript-Scope gesetzt und in
Funktionen sichtbar ist, wurde nachgemessen.

**b) TOCTOU in `Get-UniqueTargetPath` (`7_Dateien_ohne_Makro_finden.ps1`)
bestaetigt, mit Datenverlust als Folge.** Zwischen der Freiheitspruefung
(Z. 1613) und dem Schreiben (Z. 1681) liegen die komplette
COM-Konvertierung, ein `Start-Sleep` und das Umbenennen des Originals mit bis
zu vier Wiederholungen. Das Schreiben war ein
`[System.IO.File]::Copy(..., overwrite=$true)` - belegte in diesem Fenster
jemand den Zielnamen, wurde die fremde Datei ersatzlos ueberschrieben
(nachgestellt: Fremdinhalt weg). Behoben zweistufig: Ablage unter einem
eigenen `.tmp_<8 Hex>`-Namen im Zielverzeichnis, dann `File.Move` an den
endgueltigen Platz. `Move` legt nicht ueber eine bestehende Datei, sondern
wirft `ERROR_ALREADY_EXISTS` (0x800700B7, gemessen) - das ist die unteilbare
Pruefung, die dem getrennten `Exists()` fehlte. Bei Kollision wird ein
Ausweichname gezogen statt ueberschrieben, `$wasRenamed` zieht mit, und der
Fehlerzweig raeumt die Zwischenkopie ab. Drei Szenarien nachgestellt (kein
Wettlauf / Ziel belegt / Ziel und erster Ausweichname belegt): Fremddatei
bleibt jedes Mal intakt, Ergebnis landet auf `_2` bzw. `_3`, keine Reste.

**c) `_MRU_AVAILABLE` in `3c` - der Hinweis stimmt, der Name log.** Die
Variable las sich wie das Ergebnis einer Faehigkeitspruefung ("wird auf False
gesetzt, sobald RecentFiles erstmals nicht erreichbar war"), eine Pruefung gab
es aber nie: `_remove_from_mru` fasst weder `ppt_app` noch `paths` an und
schaltet beim ersten Aufruf bedingungslos um. Sie bedeutet in Wahrheit "Hinweis
schon protokolliert". Umbenannt in `_MRU_HINWEIS_GEZEIGT`, Kommentar
richtiggestellt. Die inhaltliche Aussage wurde diesmal gemessen statt
angenommen: PowerPoint 16.0 exponiert `Application.RecentFiles` ueber COM
tatsaechlich nicht (`DISP_E_UNKNOWNNAME`, 0x80020006) - Word und Excel schon.
Es gibt dort also nichts zu bereinigen; keine Verhaltensaenderung.

**d) Nicht unterstuetzte Endungen in `9` - heute unerreichbar, aber ein
stiller Verlust.** `if (-not $targetApp) { continue }` (Z. 1688) kann derzeit
nicht greifen: die Suche filtert gegen genau die Vereinigung der drei
Endungslisten. Falls der Zweig je erreicht wird, verschwand die Datei
spurlos - sie zaehlte in "N Office-Dateien gefunden" mit, tauchte aber in
keiner Berichtszeile auf. Wer eine Endung nur der Suchmenge hinzufuegt oder
aus einer App-Liste entfernt, bekaeme einen Bericht, der Vollstaendigkeit
behauptet, ohne sie zu haben. Jetzt wird die Datei als Kategorie
"Nicht geprüft" aufgenommen und protokolliert. Kategorien sind im
Excel-Bericht freier Text, also ohne Nebenwirkung.

**Pruefstand:** Alle 10 PowerShell- und 12 Python-Dateien syntaktisch
geprueft. Ausfuehrend nachgestellt wurden a) und b) inklusive Gegenprobe im
Echtlauf; c) wurde an echtem PowerPoint-COM gemessen; d) ist als
unerreichbarer Zweig nur statisch geprueft.

---

### 9. COM-Testrunde 15.08.2026 (Office 2024 lokal)

Nachdem sich herausstellte, dass Office ProPlus 2024 lokal installiert ist
(Abschnitt 1), wurden die COM-Pfade erstmals ausgefuehrt. Grundlage ist ein
per Office-COM erzeugter Wegwerf-Bestand von 18 Dateien im Scratchpad -
nie Q:/R:. Der Bestand wurde unabhaengig vom Erzeuger nachgeprueft
(Dateisignatur, `documentProtection`, `writeProtection`, `modifyVerifier`,
`vbaProject.bin`), damit die Testvoraussetzungen belegt und nicht geglaubt
sind. Jedes Skript lief erst als Probelauf, dann echt; bewertet wurde am
Dateiinhalt, nicht am Skriptprotokoll.

| Skript | Probelauf | Echtlauf | Ergebnis |
|---|---|---|---|
| `2a` Word-Schutz | sauber | korrekt | **bestanden** |
| `2c` PowerPoint-Schutz | sauber | korrekt, aber Ueberreichweite | **Befund a** |
| `3a` doc→docx | sauber | korrekt | **bestanden** |
| `3c` ppt→pptx | **aenderte Dateien** | korrekt | **Befund b, behoben** |
| `4a` Schrift Word | sauber | korrekt | **bestanden** |
| `4c` Schrift PowerPoint | sauber | korrekt | **bestanden** |
| `7` Makrofrei-Umwandlung | sauber | **blockiert** | **Befund c** |
| `9` Defektsuche | - | - | **nicht getestet, Befund d** |

Belegt bestanden: `2a` entfernte Dokument- und Schreibschutz, wandelte
`.doc` nach `.docx`, liess die verschluesselte Datei korrekt unangetastet
(ohne Kennwort im `-NoInteractive`-Modus) und das VBA-Projekt intakt; die
vier ungeschuetzten Dateien blieben **byteidentisch**. `3a` konvertierte das
Altformat und liess die aktuelle Datei unberuehrt. `4a`/`4c` ersetzten
`Courier New` durch `Arial` bis in Master und Layouts, ihre Probelaeufe
liessen auch Altformate (`.doc`, `.ppt`) byteidentisch.

**a) `2c` schreibt JEDE PowerPoint-Datei um, auch voellig ungeschuetzte.**
Das Skript zaehlt `spLocks`-Attribute wie `noGrp` und `noChangeAspect` zum
Schutz und setzt sie auf `0` - in Folie, Master und allen elf Layouts.
Genau diese Attribute schreibt PowerPoint aber standardmaessig an jeden
Platzhalter. Gemessen an der ungeschuetzten Kontrolldatei: 34515 -> 34188
Bytes, `slideLayout1.xml` vorher fuenfmal `noGrp="1"`, nachher fuenfmal
`noGrp="0"` - und das in 14 XML-Teilen. Folgen: jede `.pptx` der Migration
wird angefasst (Zeitstempel, ACL, und bei Google Drive ein vollstaendiger
Neu-Upload), das Standardverhalten von Platzhaltern aendert sich still, und
die Bilanz meldete "Entsperrt: 6", obwohl nur EINE Datei echten
Anwenderschutz trug. `2a` hat diese Ueberreichweite nicht - dort blieben
ungeschuetzte Dateien byteidentisch. **Nicht geaendert:** welche Sperren als
Schutz gelten, ist eine fachliche Entscheidung, keine Fehlerkorrektur.
Empfehlung: `spLocks` nur anfassen, wenn die Datei daneben echten Schutz
traegt (`modifyVerifier`, `writeProtection`, Verschluesselung).
**-> Behoben am 29.09.2026** (Abschnitt 10): `noGrp` und `noChangeAspect`
zaehlen nicht mehr als Sperre, die ungeschuetzte Datei bleibt byteidentisch.

**b) Der Probelauf von `3c` veraenderte die Dateien - behoben.** Der
Probelauf sagt zu, nichts zu aendern. Gemessen an einer `.ppt`: erster
Probelauf 258560 -> 260608 Bytes, zweiter wieder 258560, jedes Mal mit
anderem Hash. Der Zeitstempel wird an anderer Stelle wiederhergestellt, die
Aenderung war deshalb **unsichtbar**. Ursache war nicht das Speichern,
sondern das Oeffnen: `Presentations.Open(..., ReadOnly=COM_FALSE)`
veranlasst PowerPoint, eine Datei im Altformat sofort neu zu schreiben. Der
Kommentar im Code behauptete das Gegenteil ("Aenderungen bestehen nur im
Arbeitsspeicher und gehen beim Schliessen verloren"). Drei Varianten
gemessen: blosses `Close()` -> geaendert, `Saved=True` vor `Close()` ->
**ebenfalls geaendert**, `ReadOnly=True` -> unveraendert. Behoben durch
schreibgeschuetztes Oeffnen im Probelauf; `remove_protection()` entfaellt
dort, weil dessen Ergebnis den Probelauf-Status nicht beeinflusst.
Gegenprobe: drei Probelaeufe hintereinander, alle Dateien byteidentisch,
Statistik unveraendert; Echtlauf weiterhin korrekt. `4a`/`4c` zeigen den
Fehler bei Altformaten nicht - es war kein allgemeines Muster.

**c) `7` kann unbegrenzt haengen: die Umwandlung hat keinen Zeitwaechter.**
Der Echtlauf blieb reproduzierbar bei der ersten `.docm` stehen - zweimal
ueber neun bzw. drei Minuten ohne Fortschritt, danach abgebrochen. Der
Probelauf verarbeitet dieselbe Datei in zwei Sekunden. Es gab keinen
modalen Dialog (alle Fenster der Office-Prozesse aufgezaehlt) und keinen
Umgebungsschaden: Words `DisabledItems` stammen aus dem Alltag des Nutzers,
und nach dem Loeschen der eigenen AutoWiederherstellen-Reste blieb das
Verhalten gleich. In einem 20-Zeilen-Skript ohne jede Beteiligung von `7`
haengt `Documents.SaveAs2(<ziel>.docx, 16)` auf einer `.docm` ebenfalls
unbegrenzt - die Blockade kommt also aus Word.
**Der Skriptbefund ist ein anderer:** `7` schuetzt nur den Smoke-Test mit
`Wait-Job -Timeout` (Z. 981). Die eigentliche Umwandlung in Z. 1640-1644
(`SaveAs2`/`SaveAs`) und das vorangehende `Open` laufen voellig ungesichert.
`2a` legt jeden COM-Aufruf in einen Job mit Timeout, `3a`/`3c` haben einen
Zeitwaechter - `7` hat an der entscheidenden Stelle keinen. Eine einzige
blockierende Datei haelt damit einen unbeaufsichtigten Lauf ueber Q:/R:
fuer immer an, ohne Meldung. Das gehoert vor dem Produktiveinsatz
nachgezogen; die Muster dafuer stehen im eigenen Skript (Z. 866-981).
**Offen:** ob Word die `.docm`-Umwandlung nur auf diesem Rechner
verweigert, ist nicht geklaert.
**-> Geklaert und behoben am 29.09.2026** (Abschnitt 10): Die Blockade kam
NICHT aus Word. Der Zielpfad stammte aus `Join-Path` und wurde als
verpacktes `PSObject` an `SaveAs2` uebergeben; mit `[string]` ist die Datei
nach 0,1 s geschrieben. Der 20-Zeilen-Nachbau oben hatte denselben Fehler.
Der fehlende Zeitwaechter ist ebenfalls nachgezogen.

**d) `9` liess sich nicht automatisiert testen - und das ist selbst ein
Befund.** Als einziges der elf Skripte hat `9` keinen unbeaufsichtigten
Modus: kein `param`-Block, kein `-Automated`/`-NoInteractive`, der Zielpfad
und zwei weitere Antworten kommen ausschliesslich aus `Read-Host`.
`Confirm-YesNo` liefert bei umgeleiteter Standardeingabe bewusst `$false`,
worauf Z. 757 den Lauf beendet - ein Test ueber die Standardeingabe ist
also konstruktionsbedingt unmoeglich, ein geplanter Task ebenso. `2a`, `2c`
und `7` koennen das alle. **Der COM-Pfad von `9` bleibt ungetestet.**
**-> Behoben am 29.09.2026** (Abschnitt 10): `-TargetPath` und
`-NoInteractive`. Der erste Lauf fand sofort einen schweren Fehler.

**Nebenbefund zur Testmethode:** Word blockiert beim Speichern gelegentlich
unsichtbar, wenn es mit `Visible = $false` laeuft. Wer diese Skripte
weiterentwickelt, sollte COM-Erzeugung pro Datei in einen eigenen Prozess
mit hartem Zeitlimit legen - genau das, was Befund c) fuer `7` fordert.
**-> Zurueckgezogen am 29.09.2026:** Die "gelegentliche" Blockade war der
`Join-Path`-Fehler aus Befund c), kein Word-Verhalten. Mit reinen Strings
blockierte Word in keinem der Laeufe vom 29.09.

---

### 10. Nachtrag 29.09.2026: Gesamtpruefung mit echtem Office

**Vorgehen.** Zuerst statisch: Syntax aller Dateien; fuer Python drei
geeichte Pruefungen (undefinierte Namen, Aufrufe gegen die Signatur,
Reihenfolge beim Import); fuer PowerShell unbekannte Befehle, unbekannte
Parameter, nie zugewiesene Variablen und Aufrufe vor der Definition. Jede
Pruefung wurde vorher an einer Datei mit absichtlich eingebauten Fehlern
geeicht. Ergebnis: Python ohne Befund, PowerShell ein echter Randfall (s. u.,
Nr. 10) und vier erklaerte Fehlalarme. Danach liefen **alle Office-Skripte
gegen Office 2024** auf Wegwerf-Bestaenden im Scratchpad (nie Q:/R:), je
Probelauf und Echtlauf. Bewertet wurde am Dateiinhalt (Hash, XML-Merkmale,
Zellinhalt), nicht am Protokoll des Skripts.

| Skript | Ergebnis 29.09. |
|---|---|
| `2a` Word-Schutz | bestanden (Regression nach Nr. 8) |
| `2b` Excel-Schutz | bestanden - erstmals gelaufen; Nr. 8 |
| `2c` PowerPoint-Schutz | Nr. 7 und 8 behoben, dann bestanden |
| `3b` xls->xlsx | bestanden - erstmals gelaufen |
| `4b` Schrift Excel | Nr. 9 behoben, dann bestanden |
| `6` Excel-Berechnung | bestanden - erstmals gelaufen; Nr. 8 |
| `7` Makrofrei-Umwandlung | Nr. 1-4 und 8 behoben, dann bestanden |
| `9` Defektsuche | Nr. 5 und 6 behoben, dann bestanden |

`3a`, `3c`, `4a`, `4c` sind seit August unveraendert und wurden nicht
erneut gefahren.

**Behobene Fehler (alle nachgestellt, Gegenprobe nach der Korrektur):**

1. **`7` hat keine einzige `.docm` umgewandelt.** `$localTarget` kam aus
   `Join-Path` und ging als verpacktes `PSObject` an `Word.SaveAs2` - der
   Aufruf kehrte nie zurueck, der Hauptthread drehte mit voller CPU-Last.
   Die veroeffentlichte Fassung hing im Nachbau reproduzierbar, mit
   `[string](Join-Path ...)` war die Datei nach 0,1 s geschrieben. Das ist
   die wahre Ursache von Abschnitt 9 c.
2. **`7` hat keine einzige `.xlsm` umgewandelt.** `Test-HasExcel4Macro`
   las das Paket per `ZipFile.OpenRead`, waehrend Excel die Mappe offen
   hielt - Freigabeverletzung, Rueckgabe "nicht pruefbar", Datei behalten.
   Nach Umstellung auf `FileShare.ReadWrite` war der Typ `ZipArchive`
   unbekannt, weil `System.IO.Compression` nie geladen wurde; auch das
   ergaenzt. Jetzt: makrofreie `.xlsm` umgewandelt, XLM-Mappe erkannt und
   behalten, VBA-Mappe behalten.
3. **`7` ohne Zeitwaechter (Abschnitt 9 c).** Neu: ein Hintergrund-Runspace
   beendet nach 180 s die EIGENE Office-Instanz (Name und Startzeit
   geprueft), die Datei bleibt unveraendert, die Instanz wird neu gestartet.
   Scharfschalten, Entschaerfen und Beenden laufen unter einem Schloss.
   Gemessen: haengender Excel-Aufruf nach 3,2 s befreit; Wettlauftest
   (Frist laeuft genau beim Entschaerfen ab) 120/120 stimmig - geeicht:
   ohne Schloss findet derselbe Test einen Bruch. Im Gesamtlauf wurden die
   kennwortgeschuetzte `.docm` und `.xlsm` nach Frist abgebrochen und der
   Lauf ging weiter. **Grenze:** Dreht der Hauptthread in PowerShell selbst
   (wie bei Nr. 1), hilft das Beenden von Office nicht - gemessen.
4. **`-WhatIf` war in `7` unbenutzbar.** Der Arbeitsordner wurde nur
   simuliert angelegt, der Smoke-Test scheiterte deshalb fuer alle drei
   Programme, der Lauf brach ab - und die drei eigenen Office-Prozesse
   blieben stehen (Stop-Process ebenfalls simuliert). Jetzt identisch mit
   `-ReadOnlyMode`, Bestand byteweise unveraendert.
5. **`9` meldete JEDE intakte Excel-Datei als "Office-Fehler".** Der
   Platzhalter fuer `WriteResPassword` war 40 Zeichen lang; Excel erlaubt
   15 (gemessen: 15 ok, 16 Fehler 0x800A03EC). Eigener 15-stelliger
   Platzhalter.
6. **`9` hat jetzt einen unbeaufsichtigten Modus** (`-TargetPath`,
   `-NoInteractive`): keine Rueckfragen, kein Tastendruck, Bericht wird
   gespeichert und Excel beendet. Laufen bereits Office-Sitzungen oder
   scheitert der Smoke-Test, bricht der Lauf mit Exitcode 1 ab;
   Hintergrundprozesse werden unbeaufsichtigt nie beendet. Ergebnis auf dem
   Pruefbestand: intakte Dateien und Schreibschutz ohne Meldung, defekte,
   abgeschnittene und kennwortgeschuetzte Dateien gemeldet, Bestand unveraendert, kein
   Office-Prozess bleibt zurueck.
7. **`2c` schrieb jede `.pptx` um (Abschnitt 9 a).** PowerPoint setzt
   `noGrp` von sich aus - gemessen 72x an Platzhaltern und am
   Tabellenrahmen einer frisch erzeugten Datei; `noChangeAspect` steht an
   eingefuegten Bildern. Beide zaehlen nicht mehr. Alt: alle drei
   Testdateien umgeschrieben; neu: ungeschuetzte Datei byteidentisch,
   Aenderungskennwort und Formsperre (`noMove`/`noResize`) weiterhin
   entfernt.
8. **`-WhatIf` legte in `2a`, `2b`, `2c`, `6`, `7` das eigene Aufraeumen
   lahm** (Fehlermuster 7). Beleg aus dem `2b`-Probelauf: "WhatIf: ...
   Stop-Process ... EXCEL", eine Excel-Instanz blieb stehen, `Remove-Job`
   nur simuliert. 51 Aufrufe ergaenzt (eigene Office-Prozesse, Jobs,
   Temp-/Log-Ordner, Temp-Kopien, alte Logs); Anwenderdateien bleiben
   simuliert. Nachher in allen Probelaeufen: Bestand unveraendert, kein
   Restprozess, simuliert werden nur noch die fachlichen Aenderungen.
9. **`4b`: `Unprotect()` ohne Kennwort.** Bei einem kennwortgeschuetzten
   Blatt oeffnet Excel dann einen Kennwortdialog - auch mit
   `Interactive=False` - und der Lauf steht (gemessen: ueber 60 s ohne
   Ende). Zusaetzlich meldete die Funktion "entsperrt", sobald der Aufruf
   ohne Ausnahme zurueckkam: im Testlauf fuer ein Blatt, das danach
   weiterhin geschuetzt war. Jetzt `Unprotect("")` (wirft nach 0,03 s)
   und Nachlesen von `ProtectContents`. Blattschutz-Dateien: 0,7 s statt
   18 s, Status "teilweise" statt falsch "erfolgreich".
10. **`2c`: `trap` ruft eine spaeter definierte Funktion.** Der `trap`
    gilt fuer das ganze Skript; faellt ein Fehler vor Zeile ~915, endet der
    Aufruf in `CommandNotFound` (nachgestellt). Praktisch nur bei
    PID-Wiederverwendung erreichbar; jetzt mit `Get-Command` abgesichert.

**Offene Beobachtungen (nicht geaendert):**

- **`4b` und verschluesselte bzw. schreibreservierte Mappen:** Das Oeffnen
  mit `Password=""` loest den Kennwortdialog aus; der 60-s-Waechter von
  `4b` beendet Excel zuverlaessig, jede solche Datei kostet aber rund eine
  Minute plus Neustart. Ein nicht leerer Platzhalter half im Kurztest
  nicht; nicht weiter untersucht. In der vorgesehenen Reihenfolge entfernt
  `2b` die Schreibreservierung vorher.
- **`9` klassifiziert grob:** eine abgeschnittene `.xlsx` erscheint als
  "Office-Fehler", eine `.pptx` aus Zufallsbytes als "Unbekannt". Gemeldet
  werden beide - kein Fehler, aber die Kategorie hilft wenig.
- **Excel-Standardschrift dieses Rechners ist "Courier New, 12"**
  (`HKCU\Software\Microsoft\Office\16.0\Excel\Options\Font`). Kein Skript
  setzt sie; woher sie stammt, ist nicht feststellbar. Neue Mappen tragen
  sie - fuer Tests von `4b` wichtig, sonst folgenlos. Nicht angefasst.

**Testbedingungen - fuer kuenftige Laeufe:** Kennwortdialoge unsichtbarer
Office-Instanzen koennen auf dem Desktop auftauchen. Im ersten `4b`-Lauf
kamen zwei blockierende Aufrufe nach 18 bzw. 2 s zurueck, im zweiten
blockierten dieselben bis zum Waechter - dazwischen hat offenbar jemand die
Dialoge beantwortet. Solche Tests nur laufen lassen, wenn am Rechner
niemand klickt, und Laufzeiten gegenpruefen.

**Anonymisierung fuer die Veroeffentlichung:** Hinweise auf die Branche
und ein 8.3-Profilname in einem Kommentar wurden neutralisiert (22 Stellen).

---

### 11. Nachtrag 29.09.2026 (abends): Durchsicht mit kritischen Pruefagenten

**Vorgehen.** Neun Pruefagenten haben alle 23 Dateien vollstaendig gelesen
und nur Fehler mit konkretem Ausloeser gemeldet (je Befund: Code, Ausloeser,
Folge, Beleg). Ergebnis: rund 95 Meldungen. Jede Meldung wurde vor der
Aenderung nachgeprueft - von Behebungs-Agenten, die sie nachstellen oder
begruendet widerlegen mussten (ohne Office), und von mir: jede Differenz
gelesen, alle statischen Pruefungen, danach **jedes Office- und
OCR-Skript am echten System** (Office 2024, ocrmypdf 17.4) auf
Wegwerf-Bestaenden, Probelauf und Echtlauf, bewertet am Dateiinhalt.
Sicherung des Stands davor: Scratchpad `vor_agentenfix`.

**Widerlegt:** 1 (9: Word-Schreibkennwort > 15 Zeichen - Word nimmt es an).
E1 (5_OCR) nur teilweise: der erste Lauf verlor Dokument-Anhaenge nicht,
wohl aber Datei-Anmerkungen und ZUGFeRD-Anhaenge im Folgelauf.

**Schwerste bestaetigte und behobene Fehler:**

| Skript | Fehler (Auswahl) |
|---|---|
| 2c, 3c, 4c | PowerPoint ist Einzelinstanz: eine offene Sitzung des Anwenders wurde uebernommen, minimiert, per `Quit()` geschlossen oder beim Timeout hart beendet (gemessen: `New-Object`/`DispatchEx` liefert die laufende Sitzung). Jetzt: Abbruch, solange PowerPoint laeuft (2c/3c Exit 3, 4c Exit 2); nachgetestet mit laufender Sitzung - sie bleibt unberuehrt. |
| 3a | Fehlerpfad loeschte eine fremde, schon vorhandene `.docx`; Bearbeitungsschutz ging beim Neuspeichern verloren (jetzt erhalten, am Office belegt); Word-Prozesse wurden nach Name statt eigener PID beendet. |
| 2b | Oeffnen-Kennwoerter wurden **nie** entfernt: `$wb.WriteResPassword` ist keine Eigenschaft (Ausnahme brach jeden Versuch ab), zusaetzlich verwarf die Rueckschreib-Bedingung das Ergebnis. Jetzt am Office belegt: Datei entschluesselt. |
| 5_OCR | Ergebnis ersetzte das Original ohne Vergleich (Anhaenge, Datei-Anmerkungen, Seiten); `repair_pdf` entkernte das Original vor der OCR; ocrmypdf-Rueckgaben 4/10 ignoriert; PDF/A-2b-Dateien wurden in JEDEM Lauf neu kodiert. Belegt: zweiter Lauf aendert nichts mehr. |
| 8 | 13x `WriteLine("..." -f a, b)`: das Komma trennt Methodenargumente, `-f` warf - die Pruefsummen-Verifikation meldete "0 Dateien bit-identisch" und gab keine Datei frei (belegt: vorher auch bei Abweichung "alles gut", jetzt korrekt 2 OK + 1 ABWEICHUNG). Ausserdem: Vorfahren geschuetzter Ordner (C:\Users) als Quelle erlaubt, abgebrochenes Loeschen als ERFOLG gemeldet, negative robocopy-Exitcodes nicht erkannt. |
| 4c | Scheiterte die Stage-2-Sicherung, wurde direkt ins Original gespeichert; gescheiterte .ppt-Konvertierung als Erfolg gezaehlt; Waechter ohne Schloss (Kill nach Erfolg). |
| 4a | `word.Hwnd` gibt es nicht - der Smoke-Test-Waechter konnte Word nie beenden; Metadaten-Typnummern falsch zugeordnet (entfernte @-Erwaehnungen/Aufgaben ohne Zustimmung). Nachtrag beim Office-Test: verschluesselte bzw. schreibkennwortgeschuetzte `.docx` hingen je 180 s und hinterliessen `~$`-Besitzerdateien auf der Ablage - jetzt vorab uebersprungen (Lauf 47 s statt 401 s). |
| 3b | Echtes Rueckschritt-Zeichen (0x08) im Regex machte den Namens-Pre-Clean wirkungslos; Benutzernamen "Kriterien"/"Extract" wurden geloescht (jetzt erhalten, belegt); Dateiattribute auch im Probelauf zurueckgesetzt (jetzt erhalten, belegt). |
| 3b, 4b, 7 | XLM-Funktionen in definierten Namen (GET.CELL, EVALUATE ...) gingen beim Speichern als .xlsx verloren - 3b/4b sichern jetzt als .xlsm, 7 erkennt sie und wandelt nicht mehr um. |
| 10 | "Plus-Heilung" benannte legitime Namen falsch um (`Antrag+§34` -> `Antragä34`, `C++Kurs ©` -> `CüKurs ©`); belegt an 336 Faellen ohne Verlust echter Heilungen. |
| 1, 10, 4c | NTFS-Junctions wurden betreten (Loeschen/Umbenennen ausserhalb des gewaehlten Baums, Zyklen). |
| 0 | Stage 2 (Admin) uebernahm Installer, Zielpfade und Python-Pfad ungeprueft aus einer vom Benutzer beschreibbaren Datei; C:\OCR-Haertung liess den Benutzer als Besitzer. |
| 2a, 2b, 2c | `add_CancelKeyPress`-Skriptblock liess powershell.exe bei Strg+C hart abstuerzen (mitten im Rueckschreiben). |
| 2c | `ppAlertsNone` stand auf 2 (= alle Warnungen); `noRot`/`noChangeArrowheads` stehen ebenfalls von PowerPoint aus in Vorlagen (Notizseiten, Bilder) - weiter unnoetiges Umschreiben; Backup-Aufraeumer loeschte Anwender-`.bak`. |

Dazu viele mittlere und geringe Fehler (Schreibschutz-Attribute, relative
Pfade, `-like '\\?\*'`-Platzhalter, Statistik-, Protokoll- und
Exitcode-Fehler, Probelaeufe, die Resume-Dateien loeschten ...). Die
einzelnen Belege liegen in den Kommentaren an den geaenderten Stellen.

**Verhaltensaenderungen, die man kennen muss:**

- 2c, 3c, 4c laufen nicht, solange PowerPoint offen ist.
- 5_OCR: bei PDF/A-Ziel werden Dateien mit Anhaengen (ZUGFeRD, XRechnung,
  Datei-Anmerkungen) uebersprungen (`SKIP_HAS_ATTACHMENTS`); ersetzt wird
  nur, wenn Seiten, Anhaenge, Anmerkungen und Formularfelder vollstaendig
  sind; kein Upgrade 2b -> 2u ohne veraPDF und `--pdfa-upgrade`; keine
  `.pdf.backup` mehr; auch mit einem Worker 60-Minuten-Limit je Datei.
- 1: `.dmp`, `.wbk`, `.xlk` nur noch mit der Rueckfrage wie `.bak`.
- 10: Namen mit "+" neben §, £, ½, «, ©, ° bleiben stehen.
- 0: Stage 2 akzeptiert nur Eintraege aus der Allowlist, fuehrt Installer
  aus einem privaten Admin-Ordner aus, setzt Administratoren als Besitzer
  von C:\OCR.
- 8: die Pruefsummen-Verifikation arbeitet erstmals - damit wird das
  Loeschen der Quelle nach bestandener Pruefung tatsaechlich erreicht.
  **Vor dem ersten echten Lauf an einem kleinen, unkritischen Ordner
  ausprobieren.**

**Offen / bewusst nicht geaendert:**

- **Entscheidungen fuer den Anwender:** 4b stellt auch Symbolschriften
  (Wingdings, Symbol) auf Arial um - aus Haekchen werden Buchstaben; 2c/6
  legen Arbeitskopien in "Dokumente" (bei OneDrive-Umleitung werden sie
  hochgeladen); 3a/3c leeren die Zuletzt-verwendet-Listen von Word und
  Windows komplett, nicht nur Skriptspuren.
- 4b und 4c: verschluesselte bzw. mit Aenderungskennwort geschuetzte
  Mappen/Praesentationen kosten je 60 bzw. 180 s bis zum Waechter (kein
  Haenger, keine Reste). In der vorgesehenen Reihenfolge entfernen 2b/2c die
  Kennwoerter vorher.
- 4b: wandelt es selbst eine `.xls` um, gehen Schreibschutz/Versteckt
  verloren (3b erhaelt sie; in der Reihenfolge laeuft 3b vorher).
- XLM-Namen: am echten Excel nicht gemessen (die deutsche Oberflaeche
  nahm GET.CELL-Namen per COM in keiner Variante an); belegt mit Attrappen
  und Paketdateien.
- 0 Stage 2: Installer-Inhalt und Pruefsumme stammen weiter aus dem
  Benutzerkontext (Schliessen nur mit fester Hash-Liste oder Signatur).
- 2c/3c/4c: oeffnet der Anwender WAEHREND des Laufs PowerPoint, landet die
  Datei in der Skript-Instanz; ein Waechter-Eingriff traefe sie mit.
- 5_OCR: `try_remove_empty_password` ersetzt das Original weiterhin sofort
  (verlustarm, nur ohne Verschluesselung).

Alle diese Punkte sind inzwischen erledigt - siehe Abschnitt 12.

---

### 12. Nachtrag 29.09.2026 (nachts): Entscheidungen und Restpunkte aus Abschnitt 11

Alle Punkte der Liste "Offen / bewusst nicht geaendert" aus Abschnitt 11
sind umgesetzt und am echten Office gemessen (Office 2024; Laeufe ausserhalb
der Claude-App gestartet, siehe "Messfalle" unten).

**Entscheidungen des Anwenders:**

| Punkt | Umsetzung | Beleg |
|---|---|---|
| Symbolschriften (4b, ebenso 4a und 4c betroffen) | Text in Wingdings, Symbol, Webdings usw. behaelt seine Schrift. Erkennung: feste Liste plus alle installierten Schriften mit Symbolzeichensatz (GDI). Word: formatgebundene Suche vor dem Umstellen, danach die vier Schriftnamen zurueck; Excel: Halbierung des Zellbereichs bis zu einheitlichen Teilen, Mischzellen zeichenweise; PowerPoint: abschnittsweise (Runs). | Symbolproben Word 13/13, Excel 13/13, PowerPoint 8/8 (alte 4a-Fassung: 7 Symbole umgestellt). Per "Symbol einfuegen" gesetzte Zeichen (w:sym) blieben schon vorher erhalten. |
| Arbeitskopien nicht in "Dokumente" (2c, 6 - und ebenso 2a, 2b, 3a-c, 4a-c, 5, 7) | Gemeinsamer Ordner `%LOCALAPPDATA%\Dateimigration-Arbeitskopien`; alte Reste in "Dokumente" raeumen die Skripte weiter ab. Nur 1, 9 und 10 lagen schon ausserhalb. **Der neue Ordner muss als vertrauenswuerdiger Speicherort eingetragen werden** (wie vorher "Dokumente"). | Alle zehn Office-Skripte: keine neuen Ordner in "Dokumente", Arbeitsordner nach dem Lauf leer. |
| Zuletzt verwendet (3a, 3c - und 3b, 4a-c) | Momentaufnahme des Windows-Ordners "Recent" beim Start; entfernt werden nur Verknuepfungen, die danach entstanden sind UND auf das bearbeitete Verzeichnis oder den Arbeitsordner zeigen. Word-Liste (3a, jetzt auch 4a): nur Eintraege in diesen Ordnern. `5 --clear-mru` leert weiterhin alles, der Hilfetext sagt das jetzt. | Echtes System: 3/6 eigene Verknuepfungen entfernt, Waechter-Verknuepfungen (alt, fremd) blieben. |

**Restpunkte:**

| Punkt | Umsetzung | Beleg |
|---|---|---|
| 4b/4c: Wartezeit bei Kennwortdateien | Vorab erkannt und uebersprungen: verschluesselt (CFB statt ZIP; .xls ueber msoffcrypto), Schreibreservierung (`fileSharing`, .xls FILESHARING-Satz), Aenderungskennwort (`p:modifyVerifier`). Ebenso in **7** (dort 180 s je Datei). | 4b: 17 Mappen in 20 s statt 2x60 s Waechter; 7: 26 s statt 407 s. |
| 4b: Attribute bei eigener .xls-Umwandlung | Schreibschutz/Versteckt/System/Archiv werden gelesen, bevor der Schreibschutz fuers Speichern faellt, und am Ende zurueckgeschrieben (Fehlerfall: aufs Original). Betraf auch direkt bearbeitete Mappen. | `attr.xls` -> `attr.xlsx` mit 0x23 (R+H+A). |
| XLM-Namen am echten Excel | Mappe mit `GET.CELL` in einem Namen auf Paketebene gebaut (die Oberflaeche nimmt solche Namen per COM nicht an), von Excel als .xls gespeichert. | 3b und 4b: Ergebnis `.xlsm`, Name `Zellfarbe = GET.CELL(63,Tabelle1!$A$1)` erhalten; 7 behaelt die `.xlsm`. |
| 0 Stage 2: Installer aus dem Benutzerkontext | Authenticode-Pruefung der privaten Kopie: gueltig + erwarteter Herausgeber (Ghostscript: Artifex; Java: Eclipse, Microsoft, Oracle, Azul, Amazon, BellSoft) -> ausfuehren; gebrochene Signatur -> Abbruch; unsigniert oder fremder Herausgeber -> nur nach Bestaetigung im Admin-Fenster (mit SHA-256), ohne Konsole Abbruch. | 7/7 Faelle mit echten Dateien (gueltig, fremd, manipuliert, unsigniert, bestaetigt). |
| 2c/3c/4c: PowerPoint waehrend des Laufs geoeffnet | 4c prueft jetzt wie 2c/3c vor jeder Datei und beendet den Lauf, wenn in seiner Instanz etwas offen ist; PowerPoint bleibt offen. | Echt nachgestellt (Doppelklick waehrend des Laufs): Lauf endet, Praesentation bleibt offen. |
| 5_OCR: Leer-Kennwort-Entschluesselung ersetzte das Original sofort | Jetzt Arbeitskopie wie bei der Reparatur; das Original wird nur mit dem geprueften Ergebnis ersetzt. | Funktion: Kopie unverschluesselt, Original byte-gleich. Regression tpdf: 6/6/0, danach 0/12/0 wie zuvor. |

**Nebenbei gefunden und behoben:**

- **4b hat Excel-Textfelder nie umgestellt**: `Shape.HasTextFrame` gibt es
  nur in PowerPoint; die Abfrage warf immer, der Fehler wurde verschluckt.
- **4c stellte ueber `TextEffect.FontName` ganze Textrahmen um** (jede Form
  mit Text hat ein TextEffect-Objekt) - jetzt nur noch bei klassischem WordArt.
- **3a stellte Words Anzeige "Zuletzt verwendet" dauerhaft auf 0**
  (`DisplayRecentFiles = False` ist eine Profileinstellung). Die Liste selbst
  blieb in der Registry erhalten, war aber unsichtbar. Die Zeile ist entfernt.
  **Auf Rechnern, auf denen 3a lief:** Word > Optionen > Erweitert >
  "Diese Anzahl zuletzt verwendeter Dokumente anzeigen" wieder setzen.
- **3a und 4a trugen Arbeitskopien trotz `AddToRecentFiles=False` in Words
  Liste ein** (in der Registry gesehen: sechs Eintraege von 4a, einer von 3a).
- **3b entfernte beim Umwandeln einer .xls deren Schreibreservierungs-
  Kennwort stillschweigend** - solche Dateien werden jetzt uebersprungen
  (erst 2b).
- **5_OCR: der Entschluesselungszweig griff praktisch nie** - PyMuPDF meldet
  eine PDF mit reinem Owner-Kennwort nicht als verschluesselt; ocrmypdf
  verarbeitet sie selbst (gemessen). Der Zweig bleibt fuer Sonderfaelle.
- Excel-COM in spaeter Bindung: `Range.Characters(i, n)` und
  `TextRange2.Runs(i)` scheitern ("Mitglied nicht gefunden" bzw.
  "Auflistung nicht unterstuetzt"); `GetCharacters`/`GetRuns` gehen.

**Messfalle dieser Umgebung:** Die Claude-Desktop-App ist ein MSIX-Paket.
Dateien, die aus ihr heraus unter `%LOCALAPPDATA%`/`%APPDATA%` (ausser Temp)
angelegt werden, landen virtualisiert im Paketordner - Word sieht sie nicht
("Datei nicht gefunden"). Die Office-Tests liefen deshalb ueber einen per WMI
ausserhalb des Pakets gestarteten Prozess, so wie der Anwender die Skripte
startet. Beim Anwender tritt die Falle nicht auf.

**Verbleibt (bewusst):**

- PowerPoint: oeffnet der Anwender eine Datei, WAEHREND ein Aufruf haengt,
  traefe der Waechter-Eingriff sie mit (zwischen zwei Dateien wird jetzt
  geprueft, waehrend eines haengenden Aufrufs geht das per COM nicht).
- 4b: sehr kleinteilig formatierte Blaetter (> 20 000 Teilbereiche) werden
  pauschal umgestellt; reine Symbolzellen kommen zurueck, Symbolzeichen
  INNERHALB gemischter Zellen nicht (Warnung im Protokoll).
- Nicht erkannt: Aenderungskennwort in .ppt, Kennwoerter in .xlsb.
- 2a/2b/2c (PowerShell) hinterlassen wie bisher Recent-Verknuepfungen auf
  ihre (geloeschten) Arbeitskopien.
- 0 Stage 2 mit `--yes`: die Bestaetigung fuer unsignierte Installer gilt
  als erteilt (Tesseract und veraPDF sind unsigniert).

Diese Punkte und die Entscheidung zu 3b sind in Abschnitt 13 nachgearbeitet.

---

### 13. Nachtrag 30.09.2026: Rueckmeldung des Anwenders zu Abschnitt 12

**Entscheidungen:**

- **Schreibkennwoerter werden entfernt.** 3b entfernt das Schreib-
  reservierungs-Kennwort einer .xls beim Umwandeln (so war es schon; das
  Ueberspringen aus Abschnitt 12 ist zurueckgenommen, der Wegfall steht jetzt
  im Protokoll). 3c entfernt Aenderungskennwoerter in .pptx & Co. ueber eine
  Arbeitskopie ohne `<p:modifyVerifier>` (wie 2c).
- **Stage 2 ohne Rueckfrage.** Ghostscript und Java muessen gueltig vom
  erwarteten Herausgeber signiert sein (sonst Abbruch - diese Installer sind
  ab Werk signiert); Tesseract und veraPDF laufen unsigniert, die Pruefsumme
  steht im Protokoll; eine gebrochene Signatur bricht immer ab. 8/8 Faelle.
- **Word "Zuletzt verwendet" wieder auf 50** gesetzt (per Word selbst; die
  Liste war vollstaendig erhalten). Die 28 Eintraege, die meine Testlaeufe am
  29.09. dort hinterlassen hatten (Arbeitskopien von 7, 4a, 3a), sind
  entfernt; nichts anderes.

**Behoben:**

| Punkt | Umsetzung | Beleg |
|---|---|---|
| PowerPoint-Waechter traefe eine Anwenderdatei | 2c/3c/4c pruefen vor dem Beenden die Fenster der Instanz: die Skript-Instanz hat nur den Rahmen "PowerPoint" (eigene Praesentationen fensterlos), eine Anwenderdatei macht daraus "x.pptx - PowerPoint". Dann wartet der Waechter mit Hinweis, bis die Datei geschlossen ist; kehrt der Aufruf vorher zurueck, entfaellt das Beenden ganz. Aufraeum-Kills am Laufende entfallen in dem Fall. | Echtes PowerPoint mit offener Anwenderdatei und kuenstlich haengendem Aufruf: 3c/4c 4/4, 2c 2/2. Gemessen dazu: waehrend eines modalen Dialogs (Kennwortabfrage) oeffnet PowerPoint gar keine Anwenderdatei. |
| 2a/2b/2c hinterliessen Recent-Eintraege | `_gemeinsam.psm1`: Momentaufnahme beim Start, am Ende nur neue Verknuepfungen auf eigene Pfade entfernen (auch 8.3-Kurzform von %TEMP%). Ebenso in **7** (13 je Lauf). | Regression: 2a/2b/2c/7 hinterlassen 0; Einzeltest 5/5. |
| 7 fuellte Words Liste | `Documents.Open`/`SaveAs2` ohne `AddToRecentFiles=False` (21 Eintraege gefunden). | Nach einem 7-Lauf 28 -> 28. |
| .ppt / .xlsb: Kennwoerter nicht vorab erkannt | .xlsb: Satz BrtFileSharingIso (0x2A4) bzw. 0x224 mit Hash in `xl/workbook.bin`; verschluesselt = CFB. .ppt: Oeffnungskennwort -> msoffcrypto; Aenderungskennwort ist in .ppt eine Verschluesselung mit festem Standardkennwort (entschluesselbar, daran erkannt). 2c (PowerShell) erkennt das Oeffnungskennwort am headerToken des CurrentUserAtom (0xF3D1C4DF). | 4b: .xlsb 4/4 richtig; 3c/4c: .ppt 4/4; 2c: pw.ppt sofort statt nach 180 s. |
| 3c: aendpw.pptx | Vorher kein Waechter, sondern dreimal "SaveAs : Failed". Jetzt Arbeitskopie ohne Kennwort -> umgewandelt, Kennwort weg. Oeffnungskennwort -> sofort uebersprungen. | tp2: aendpw.pptx ohne modifyVerifier, pw.ppt unveraendert. |

**Gemessen und nicht behebbar:** Eine **.ppt mit Aenderungskennwort** kann
PowerPoint ohne das Kennwort nicht speichern - schreibgeschuetzt oeffnen geht,
aber SaveAs, SaveCopyAs und das Leeren von WritePassword scheitern ("Presentation
cannot be modified"), auch mit `Untitled`. Die mit msoffcrypto entschluesselte
Datei lehnt PowerPoint als beschaedigt ab. 3c und 4c ueberspringen sie sofort
mit Hinweis, 2c meldet einen Fehler (ohne Wartezeit). Der gegenteilige
Kommentar in 2c ist berichtigt.

---

### 14. Nachtrag 30.09.2026: Schreibkennwoerter auch in 3a, 4a-4c und 7 entfernen

Entscheidung des Anwenders: Schreib-/Aenderungskennwoerter werden ueberall
entfernt, nicht uebersprungen (4a und 3a sind der Vollstaendigkeit halber
mit angeglichen - sie verhielten sich wie 4b bzw. umgekehrt wie 3b).

**Gemessen (Office 2024):** Excel und Word oeffnen solche Dateien
schreibgeschuetzt ohne Abfrage. Excel speichert sie per `SaveAs` mit leerem
`WriteResPassword` ohne Kennwort (.xlsx, .xls, .xlsb). Word dagegen traegt ein
Schreibkennwort aus einer .doc per `SaveAs2` in die neue .docx weiter,
`WritePassword=""` wird ignoriert - dort wird auf Dateiebene nachgezogen.
PowerPoint siehe Abschnitt 13 (.pptx ueber die Datei, .ppt nicht moeglich).

| Skript | Umsetzung |
|---|---|
| 3a | Nach dem Umwandeln `<w:writeProtection>` mit Hash aus der neuen .docx entfernen (Bearbeitungsschutz `w:documentProtection` bleibt, wie bisher). |
| 4a | .docx & Co.: Arbeitskopie ohne `<w:writeProtection>`, behandelt wie eine Langpfad-Kopie. .doc: Stage 1 oeffnet jetzt schreibgeschuetzt (vorher haette eine .doc mit Schreibkennwort bis zum Waechter gehangen) und zieht das Kennwort nach dem Speichern auf Dateiebene nach. |
| 4b | Schreibgeschuetzt oeffnen, per `SaveAs` ohne Schreibkennwort speichern (Stage 1 und 2); bei Langpfad-Kopien wird das Ergebnis nach dem Schliessen ueber die Kopie gelegt, die Schritt 8 zurueckschiebt. |
| 4c | .pptx & Co.: Arbeitskopie ohne `<p:modifyVerifier>` (wie 3c); .ppt weiter uebersprungen. |
| 7 | Kennwort in der lokalen Kopie entfernen (w:writeProtection, fileSharing, modifyVerifier); die umgewandelte Datei entsteht ohne. Behaelt 7 eine Datei (Makros), bleibt das Original wie immer unangetastet. |

**Beleg:** Bestaende tw2/tx6/tp2 und ein 7-Bestand mit Kennwortdateien, je
auch mit Pfad ueber 240 Zeichen: 3a 2/2, 4a 4/4, 3b 2/2, 4b 9/9, 3c 4/4,
4c 4/4, 2c 2/2, 7 3/3 - Schreibkennwort jeweils weg, Oeffnungskennwoerter
und .ppt mit Aenderungskennwort unveraendert. Regression aller Office-
Skripte, Symbolproben und 7 danach unveraendert gruen.

---


# KRITISCH

## `2b_entferne_schutz_excel.ps1` Zeile 1867  `[BEHOBEN]`

**New-Item kennt keinen Parameter -LiteralPath: der temporaere Arbeitsordner wird nie angelegt, das Skript verarbeitet keine einzige Datei.**

```
New-Item -LiteralPath $TempPath -ItemType Directory -Force | Out-Null
```

*Begruendung:* New-Item besitzt weder in Windows PowerShell 5.1 noch in PowerShell 7 einen Parameter -LiteralPath; das Parameterset ist Path, Name, ItemType, Value, Force, Credential. Nachgewiesen mit '(Get-Command New-Item).Parameters.Keys' unter 5.1.26100.9168 und 7.6.4 - '-LiteralPath' fehlt in beiden. Der Aufruf scheitert mit einer ParameterBindingException ('Es wurde kein Parameter gefunden, der dem Parameternamen "LiteralPath" entspricht'). Der Fehler ist nicht abbrechend, das Skript laeuft weiter - nur existiert $TempPath ab hier nicht. Zeile 1864-1866 hat einen eventuell vorhandenen Vorgaengerordner unmittelbar davor geloescht, es gibt also keinen Rueckfall. Eine AST-Pruefung aller Kommandoaufrufe der Datei meldet genau diese eine Fundstelle.

*Auswirkung:* Direkt danach laeuft der Trust-Center-Smoke-Test: New-MinimalXlsx ruft ZipFile::Open($TempPath\trustcheck_*.xlsx,'Create') auf und bekommt DirectoryNotFoundException ('Ein Teil des Pfades ... konnte nicht gefunden werden'). Test-ExcelTrustCenter liefert Ok=$false mit 'Test-XLSX konnte nicht erstellt werden'. Mit -NoInteractive (Aufgabenplanung) beendet sich das Skript an Zeile 1923 mit exit 2 - der geplante Task meldet taeglich Fehlschlag, ohne je eine Datei angefasst zu haben. Interaktiv wird dem Anwender die Falschdiagnose 'Dokumentenordner nicht als vertrauenswuerdiger Speicherort eingetragen' angezeigt; bestaetigt er mit 'j', scheitert fuer JEDE gefundene Datei die Arbeitskopie in Zeile 2132 ([System.IO.File]::Copy($srcLong,$tempFile,$true)) mit derselben DirectoryNotFoundException. Die Meldung 'Ein Teil des Pfades ... konnte nicht gefunden werden' passt auf keines der Transient-Muster in Invoke-WithRetry, und die Klausel '$_.Exception -is [System.IO.IOException]' greift nicht, weil PowerShell .NET-Methodenfehler in eine MethodInvocationException verpackt (nachgestellt: -is [IOException] ergibt False). Es wird also sofort geworfen, der aeussere catch in Zeile 2411 zaehlt stats.Errors hoch: 100 Prozent Fehlerquote, kein einziger Schutz entfernt.

*Vorschlag:* -LiteralPath durch -Path ersetzen (der Pfad enthaelt keine Wildcards) oder besser provider-frei anlegen: [System.IO.Directory]::CreateDirectory($TempPath) | Out-Null. Zusaetzlich sollte der Smoke-Test-Fehlerpfad die Existenz von $TempPath explizit pruefen und melden, statt das Trust Center zu beschuldigen.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `3c_ppt_pptx_auf_neueste_Version_aktualisieren.py` Zeile 2571  `[BEHOBEN]`

**Fehlerpfad stellt aus einem Backup wieder her, dessen Erzeugung gerade gescheitert ist - und ueberschreibt damit die noch unveraenderte Originaldatei mit einem Fragment.**

```
if backup_path and safe_exists(backup_path) and target_path:
            if robust_copy(backup_path, target_path):
                _apply_security_descriptor(target_path, orig_sd)
                safe_remove(backup_path)
                pbar.write("    → Backup wiederhergestellt")
```

*Begruendung:* backup_path wird bereits in Zeile 2462 gesetzt, also BEVOR die Backup-Kopie ueberhaupt versucht wird. Scheitert robust_copy(target_path, backup_path) in Zeile 2487, wirft der Code bewusst ab ('Abbruch zum Schutz der Zieldatei', Zeile 2488-2491). shutil.copy2 hinterlaesst bei einem Abbruch mitten in der Kopie aber eine TEILDATEI am Ziel - nachgestellt: nach einem Lesefehler nach 1 MB von 5 MB existiert die Zieldatei mit 1048576 Bytes und os.path.exists() liefert True. Der except-Zweig in Zeile 2571 prueft nur safe_exists(backup_path), nicht ob die Sicherung jemals gelungen ist und nicht, ob die Zieldatei ueberhaupt schon angetastet wurde (converted_ok/update_complete werden hier nicht abgefragt). Er kopiert das Fragment ueber die intakte Originaldatei und meldet 'Backup wiederhergestellt'.

*Auswirkung:* Ablage Q:\, Datei 'Jahresbericht.pptx' (40 MB). Die SMB-Verbindung bricht waehrend der Backup-Kopie weg; alle drei robust_copy-Versuche scheitern, 'Jahresbericht.pptx.bak' bleibt mit 12 MB liegen. Ausnahme -> Zeile 2571 -> robust_copy(.bak, Jahresbericht.pptx) trunkiert die noch voellig unveraenderte Originaldatei auf 12 MB, loescht anschliessend das .bak (Zeile 2574) und meldet Erfolg. Die einzige vollstaendige Fassung ist weg, obwohl das Skript diese Datei nie geoeffnet-gespeichert hatte.

*Vorschlag:* Wiederherstellung nur zulassen, wenn die Zieldatei tatsaechlich ersetzt wurde (Flag converted_ok aus Zeile 2503 abfragen) UND die Sicherung nachweislich gelungen ist. Dazu ein eigenes Flag backup_ok setzen, das erst nach erfolgreichem robust_copy in Zeile 2487 True wird; im Fehlerfall das Fragment mit safe_remove(backup_path) entfernen und backup_path auf None setzen, bevor die Ausnahme weitergereicht wird.


## `5_OCR_PDF.py` Zeile 102  `[BEHOBEN]`

**C:\OCR wird an den Anfang von PATH gesetzt und als Programm-/DLL-Fundort benutzt, obwohl das Verzeichnis auf Windows-Standardinstallationen fuer JEDEN authentifizierten Benutzer schreibbar ist - bei einem Lauf 'als Administrator' ist das eine lokale Rechteausweitung.**

```
os.environ["PATH"] = r"C:\OCR;" + r"C:\OCR\bin;" + os.environ.get("PATH", "")
```

*Begruendung:* Zeile 102 stellt C:\OCR und C:\OCR\bin VOR alle Systemverzeichnisse in PATH; der Wert wird ueber os.environ an jeden Subprozess (Tesseract, Ghostscript, veraPDF, unpaper, jbig2) und an die Worker (spawn) vererbt. Zusaetzlich stehen feste C:\OCR-Kandidaten ganz oben in den Suchlisten: Zeile 2380/2381 (C:\OCR\Tesseract-OCR\tesseract.exe, C:\OCR\tesseract.exe) VOR C:\Program Files, Zeile 2520/2521 (Ghostscript), 2565/2566 (jbig2.exe), 2747-2750 (verapdf.exe/.bat). Gefundene Treffer werden ungeprueft ausgefuehrt (_run, Zeile 2404/2534/2575/2762) - eine Signatur- oder Herkunftspruefung findet nicht statt. Findet sich in C:\OCR eine Datei nach den Leptonica-Mustern (Zeile 2467-2474), wird ihr Verzeichnis ueber _register_dll_dirs (Zeile 2497) sogar dauerhaft in den DLL-Suchpfad des Prozesses aufgenommen. Nachgestellt auf diesem Rechner mit 'icacls C:\ ' und 'icacls C:\OCR': C:\ traegt 'NT-AUTORITAET\Authentifizierte Benutzer:(OI)(CI)(IO)(M)', C:\OCR hat daraus geerbt 'NT-AUTORITAET\Authentifizierte Benutzer:(I)(OI)(CI)(M)' - also Aendern-Recht fuer jeden angemeldeten Benutzer. Das ist kein Sonderfall dieses Rechners, sondern die Standard-ACL jedes direkt unter C:\ angelegten Ordners. Das Skript legt einen Admin-Lauf selbst nahe (Hinweis in Zeile 980/981 zu Netzlaufwerken bei Start 'als Administrator').

*Auswirkung:* Ein Standardbenutzer (oder Schadcode in dessen Kontext) legt C:\OCR\tesseract.exe ab oder ersetzt eine vorhandene tesseract.exe/gswin64c.exe/verapdf.bat. Sobald ein Administrator das Skript ueber die Archivablage startet, wird ab Zeile 2404 (Systempruefung, noch vor der ersten PDF) diese Datei mit Administratorrechten ausgefuehrt. Gleiches gilt fuer eine untergeschobene leptonica-1.x.dll, deren Verzeichnis in Zeile 2497 in den DLL-Suchpfad gehaengt wird.

*Vorschlag:* C:\OCR beim Start auf sichere ACL pruefen (kein Schreibrecht fuer Users/Authenticated Users/CREATOR OWNER) und bei Verstoss abbrechen statt zu nutzen; alternativ die Werkzeuge nach %ProgramFiles%\OCR verlegen, das Verzeichnis mit expliziter ACL anlegen (icacls /inheritance:r und nur Administratoren/SYSTEM Schreibrecht) und C:\OCR aus Zeile 102, den Kandidatenlisten und EXTRA_DLL_DIRS (Zeile 594/595) entfernen. Falls C:\OCR bleiben muss: den Ordner beim Setup mit 'icacls C:\OCR /inheritance:r /grant Administratoren:(OI)(CI)F SYSTEM:(OI)(CI)F Benutzer:(OI)(CI)RX' absichern und diese Pruefung im Skript wiederholen.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `8_verschieben_auf_Google_Drive.ps1` Zeile 2528  `[BEHOBEN]`

**Der abschliessende Remove-Item -Recurse loescht genau die Verzeichnisse, die der /MIR-Purge per /XD bewusst verschont hat - also Daten, die nie kopiert wurden.**

```
Remove-Item -LiteralPath (Add-LongPathPrefix $SourcePath) -Recurse -Force -ErrorAction Stop
```

*Begruendung:* Der Kommentar in Zeile 74-80 legt als Invariante fest, dass $script:ExcludedDirNames ('$RECYCLE.BIN', 'System Volume Information', '~snapshot', '.snapshot') UEBERALL ausgeschlossen sind - 'beim Kopieren ..., bei Pruefung/Verifikation UND beim Loeschen'; '~snapshot'/'.snapshot' werden ausdruecklich als NAS-Schattenkopien bezeichnet, die 'nicht migriert werden duerfen'. Der Purge in Zeile 2516-2518 haelt diese Zusage ein: nachgestellt mit echtem robocopy bleibt ein '~snapshot'-Baum nach '<leer> <quelle> /MIR /XD ~snapshot ...' vollstaendig erhalten (Exit-Code 2, Nutzdatenordner geloescht, Snapshot-Datei noch da). Unmittelbar danach setzt Zeile 2523-2531 fuer jede Nicht-Root-Quelle ein 'Remove-Item -Recurse -Force' auf denselben Ordner an - und das kennt die Ausschlussliste nicht. Damit wird ausgerechnet der Bestand vernichtet, der nie ins Ziel kopiert wurde und deshalb auch in keiner Verifikation auftaucht. Robocopy wurde zuvor mit '/XD ~snapshot' aufgerufen (Zeile 2069), es existiert also nachweislich keine Zielkopie.

*Auswirkung:* Quelle 'Q:\Abteilung' (Nicht-Root) auf einer NetApp-Freigabe mit sichtbarem Snapshot-Verzeichnis, Methode [1] Robocopy, Verifikation [3] 'Keine', RoboExit < 4: Nachgestellt in %TEMP% -> nach dem /MIR-Purge ist 'Abteilung\~snapshot\taeglich_2026-08-01\altbestand.docx' noch vorhanden (True), nach dem Remove-Item ist sie weg (False), obwohl sie nie nach G: kopiert wurde. Auf einer echten NetApp ist der Snapshot-Ordner meist schreibgeschuetzt - dann scheitert Remove-Item stattdessen, der Quellordner bleibt stehen und der Lauf meldet 'Fehler beim Loeschen'. Beide Ausgaenge sind falsch.

*Vorschlag:* Den Quellordner nicht per Remove-Item -Recurse aufloesen, sondern - wie im selektiven Zweig (Zeile 2420-2450) - nur noch leere Verzeichnisse tiefenzuerst mit [System.IO.Directory]::Delete($dir, $false) entfernen und die Wurzel selbst nur loeschen, wenn sie danach leer ist. Verbleiben ausgeschlossene Verzeichnisse, ist der Ordner absichtlich nicht leer: das gehoert als Hinweis ins Protokoll ('X Systemordner absichtlich behalten'), nicht in ein rekursives Force-Delete.

*Vom Pruefer ausgefuehrt nachgestellt.*

*Bei der Verifikation zusaetzlich gefunden (14.08.2026):* Der Befund greift zu
kurz. Der /MIR-Purge in Zeile 2516-2518 haelt die Zusage NUR fuer
Verzeichnisse unmittelbar in der Quellwurzel. Nachgestellt mit echtem
robocopy: von `~snapshot`, `A\~snapshot`, `A\B\~snapshot`, `.snapshot`,
`A\.snapshot` und `A\B\System Volume Information` ueberlebten den Purge
ausschliesslich die beiden Eintraege in der Wurzel - alle verschachtelten
wurden geloescht. Auch das Nachreichen der VOLLPFADE an /XD aendert daran
nichts (ebenfalls nachgestellt): weil der Leerordner den Zwischenordner `A`
gar nicht enthaelt, gilt der komplette Ast als ueberzaehlig und wird
geloescht, ohne dass /XD je darunter zum Tragen kaeme. Die Kopierseite
(Zeile 2069, `/E`) verhaelt sich dagegen korrekt und schliesst auf JEDER
Ebene aus - nachgestellt. Damit gilt fuer verschachtelte Snapshots exakt
dieselbe Datenverlustklasse wie im Befund oben, nur ueber robocopy statt
ueber Remove-Item: nie kopiert, trotzdem geloescht.

*Behoben:* (a) Remove-Item -Recurse ersetzt durch tiefenzuerst nur leere
Verzeichnisse plus Bericht ueber absichtlich behaltene Ordner
(`Get-RemainingExcludedDir`). (b) Vor dem Purge wird geprueft, ob
ausgeschlossene Verzeichnisse UNTERHALB der Wurzel liegen; wenn ja, wird der
schnelle robocopy-Purge uebersprungen und ueber die skripteigene
Aufzaehlung dateiweise geloescht, die die Ausschlussliste auf jeder Ebene
beachtet. Beide Pfade nachgestellt: flacher und verschachtelter Snapshot
ueberleben, Nutzdaten werden entfernt, keine leeren Ordner bleiben zurueck;
im Normalfall ohne Ausschluesse verschwindet die Quellwurzel wie bisher.


## `9_fehlerhafte_Dateien_finden.ps1` Zeile 43  `[BEHOBEN]`

**Als .exe kompiliert (ps2exe) bricht das Skript sofort bei der ersten Rueckfrage ab - der Interaktivitaetstest erkennt den ps2exe-Host nicht**

```
$interactive = ($Host.Name -eq 'ConsoleHost')
```

*Begruendung:* Der Kopf der Datei (Zeilen 5-9) schreibt die ps2exe-Uebersetzung ausdruecklich als Auslieferungsweg vor ('Invoke-ps2exe ... -sta'). ps2exe stellt aber einen eigenen PSHost bereit, dessen Name NICHT 'ConsoleHost' ist. In der auf diesem Rechner installierten Version 1.0.18 liefert die Host-Implementierung 'PSRunspace-Host' (ps2exe.ps1, Zeile 2435: return "PSRunspace-Host";). Damit ist $interactive immer $false, Confirm-YesNo faellt in Zeile 53-56 in den Zweig 'Keine interaktive Konsole' und gibt $false zurueck - obwohl eine voll funktionsfaehige Konsole samt Read-Host da ist (das Skript wird ja gerade OHNE -noConsole kompiliert). Nachgestellt: identischer Codeausschnitt mit ps2exe 1.0.18 uebersetzt und ausgefuehrt, Ausgabe 'Host.Name = PSRunspace-Host / interactive = False'.

*Auswirkung:* 9_fehlerhafte_Dateien_finden.exe starten -> die erste Rueckfrage 'Konfiguration pruefen' (Zeile 662) liefert ohne Zutun $false -> Zeile 664-667 'Bitte konfigurieren Sie zuerst das Trust Center ... Abbruch.' -> Stop-Script. Es wird nie eine Datei geprueft, nie ein Bericht erzeugt. Betroffen sind alle drei Rueckfragen (Zeile 662, 936, 1348). Als .ps1 in powershell.exe gestartet faellt der Fehler nicht auf, in der ausgelieferten .exe laeuft das Skript gar nicht an.

*Vorschlag:* Nicht am Hostnamen festmachen, sondern an der tatsaechlichen Eingabefaehigkeit. Z.B.: $interactive = $true; try { if ([Console]::IsInputRedirected) { $interactive = $false } } catch { $interactive = ($Host.Name -eq 'ConsoleHost') }. Falls der Hostname weiter geprueft werden soll, ps2exe mit aufnehmen: ($Host.Name -in @('ConsoleHost','PSRunspace-Host')). Dieselbe Aufweichung ist in Zeile 152 (TreatControlCAsInput) sinnvoll, dort ist die Folge aber nur der fehlende Strg+C-Abbruch.

*Vom Pruefer ausgefuehrt nachgestellt.*



# HOCH

## `0_Vorab-Check.py` Zeile 376  `[BEHOBEN]`

**run_cmd() setzt seinen Timeout nicht durch: startet der aufgerufene Prozess einen Hintergrundprozess, der die Pipes erbt, blockiert subprocess.run() weit ueber den Timeout hinaus - und der eigentliche Installer wird nie beendet.**

```
p = subprocess.run(
            args,
            capture_output=True,
            text=False,
            timeout=timeout,
            shell=False,
            cwd=cwd,
            startupinfo=get_startupinfo(),
            env=env,
        )
```

*Begruendung:* subprocess.run() mit capture_output=True toetet bei Timeout nur den DIREKTEN Kindprozess und ruft danach communicate() OHNE Timeout auf. Diese zweite communicate()-Runde wartet auf EOF der Pipes. Ein vom Kind gestarteter Enkelprozess (bei .bat/cmd.exe/Installern der Normalfall: 'start ...', java-Launcher, nachgeladene Setup-Stufen) erbt die Pipe-Handles und haelt sie offen. Damit kehrt der Aufruf erst zurueck, wenn der Enkel von sich aus endet - der Timeout-Parameter ist wirkungslos. Zusaetzlich laeuft der eigentliche Installer nach dem 'Timeout' unbeaufsichtigt weiter, waehrend das Skript ihn als abgebrochen protokolliert und mit dem naechsten Schritt fortfaehrt.
Nachgestellt auf diesem Rechner (Python 3.14): eine .bat, die per 'start /b' einen 40-s-Enkel startet, mit subprocess.run(capture_output=True, timeout=5) aufgerufen -> TimeoutExpired kam erst nach 39,6 s statt nach 5 s.

*Auswirkung:* Konkret: Stage 2 ruft in Zeile 2366 run_cmd([bat, xml_file], timeout=900) auf die veraPDF-Installer-.bat auf (IzPack startet Java). Startet die .bat den Java-Prozess im Hintergrund, haengt Stage 2 beliebig lange - nicht 900 s. Stage 1 wartet in Zeile 624 mit WaitForSingleObject(handle, 0xFFFFFFFF) unbegrenzt auf Stage 2. Ergebnis: der komplette Zwei-Stufen-Lauf steht still, ohne Fehlermeldung, ohne Timeout, auf einem fremden Rechner nur per Task-Manager aufloesbar. Gleiches gilt fuer run_installer() (Zeile 1685/1688/1693, timeout=1200). Nach einem 'Timeout' meldet run_installer zudem 'Installer-Timeout' und faehrt fort, obwohl der Installer real noch laeuft - der naechste Installer kann parallel gestartet werden (msiexec-Sperre, halbfertige Installation).

*Vorschlag:* Nicht subprocess.run mit Pipes verwenden, wo ein Timeout garantiert sein muss. Stattdessen: Popen mit stdout/stderr auf temporaere DATEIEN (keine Pipes) umlenken, dann p.wait(timeout=...); bei TimeoutExpired den gesamten Prozessbaum beenden (subprocess.run(['taskkill','/PID',str(p.pid),'/T','/F'])) und erst danach die Temp-Dateien lesen. Alternativ das Kind in ein Windows-Job-Objekt mit JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE haengen. Zusaetzlich in relaunch_as_admin_and_wait() (Zeile 624) statt INFINITE eine Obergrenze setzen (z. B. 3600 s) und bei Ablauf sauber abbrechen.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `0_Vorab-Check.py` Zeile 2743  `[BEHOBEN]`

**Schlaegt die Pruefsummenbildung in Stage 1 fehl, wird sha256=None ins Manifest geschrieben - Stage 2 fuehrt die Datei dann per Warnung UNGEPRUEFT mit Adminrechten aus (fail open).**

```
entry["sha256"] = sha256_of(entry.get("installer"))
                if entry.get("dir"):
                    entry["files"] = hash_tree(entry["dir"])
```

*Begruendung:* sha256_of() (Zeile 457-465) faengt JEDE Exception ab und liefert dann None; hash_tree() (Zeile 467-492) liefert im Fehlerfall ein leeres dict. Beide Rueckgabewerte werden hier ohne Pruefung ins Manifest uebernommen. In Stage 2 gilt dann: verify_staged_entry() Zeile 743-744 -> 'if not expected: log_warn(... wird ungeprueft ausgefuehrt)' und faehrt fort; verify_staged_dir() Zeile 709-710 -> 'if not root or not expected: return True'. Damit degradiert der Schutzmechanismus, auf dem das gesamte Two-Stage-Modell beruht (siehe die eigenen Kommentare in Zeile 2603-2616 und 1723-1734), lautlos zu 'gar keine Pruefung' - und zwar genau in dem Moment, in dem etwas mit der Stage-Datei nicht stimmt. Ein Schutz, der bei Fehlern aufmacht statt zumacht, ist die falsche Richtung.

*Auswirkung:* Datei C:\tmp_skripte\stage\tesseract-ocr-w64-setup-5.5.0.exe ist beim Schreiben des Manifests kurz gesperrt (Virenscanner-Scan, Backup-Agent, geoeffnetes Handle) -> sha256_of() wirft, liefert None -> Manifest enthaelt "sha256": null -> Stage 2 gibt nur '⚠️ Keine Pruefsumme im Manifest ... wird ungeprueft ausgefuehrt' aus und startet die EXE mit Adminrechten. Wird die Datei zwischen Stage 1 und Stage 2 ausgetauscht (genau das Angriffsszenario, gegen das Zeile 2611-2616 argumentiert), faellt es nicht mehr auf. Beim veraPDF-Ordner reicht schon ein Fehler in hash_tree(), damit der ganze Baum ungeprueft bleibt.

*Vorschlag:* In Zeile 2743 den Rueckgabewert pruefen: digest = sha256_of(...); wenn None -> Eintrag NICHT ins Manifest aufnehmen, log_err + STATE['manual_actions'] setzen. Analog fuer hash_tree(): leeres Ergebnis bei gesetztem entry['dir'] als Fehler behandeln. Zusaetzlich in verify_staged_entry() Zeile 743 aus dem log_warn ein 'return False' machen (fehlende Pruefsumme = Abbruch), ebenso verify_staged_dir() Zeile 709 bei vorhandenem 'dir' aber leerer Hash-Liste.


## `10_dateinamen_bereinigen.py` Zeile 859  `[BEHOBEN]`

**Phase A ersetzt '+' dauerhaft durch '├', auch wenn die anschliessende Heilung fehlschlaegt - legitime Namen mit Akzentzeichen neben einem Plus werden verstuemmelt.**

```
chars[i] = "├"  # ├
```

*Begruendung:* Die Substitution ist laut Docstring nur eine Vorstufe fuer den Heiler ('Substitut fuer + als Box-Drawing-Substitut'). Sie wird aber unbedingt vorgenommen und nirgends zurueckgenommen, wenn _try_heal danach kein Ergebnis liefert. _is_mojibake_neighbor() haelt jedes Nicht-ASCII-Zeichen ausser aeoeueAEOEUEss fuer einen Mojibake-Indikator - also auch é, à, ©, €, °, µ, die in Archivbestaenden massenhaft legitim vorkommen. Das '├' bleibt dann als endgueltiger Bestandteil des neuen Namens stehen und wird auf der Platte umbenannt. Der Docstring behauptet ausdruecklich das Gegenteil ('sonst bleiben legitime C++_Tutorial.pdf und 1+1=2.txt unangetastet').

*Auswirkung:* Nachgestellt ueber sanitize_filename(): 'André+Partner.pdf' -> 'André├Partner.pdf'; 'Café+Restaurant Plan.pdf' -> 'Café├Restaurant Plan.pdf'; 'Résumé+Anhang.doc' -> 'Résumé├Anhang.doc'; 'Muster GmbH ©+Rechte.pdf' -> 'Muster GmbH ©├Rechte.pdf'; 'Größe 10€+MwSt.pdf' -> 'Größe 10€├MwSt.pdf'; 'Temperatur 20°+30° Meßreihe.txt' -> 'Temperatur 20°├30° Meßreihe.txt'. Jeweils echte Umbenennung auf der Platte.

*Vorschlag:* Die Substitution nur behalten, wenn der resultierende Cluster tatsaechlich heilt: _pre_substitute_plus spekulativ auf einer Kopie ausfuehren, den kompletten Heildurchlauf darauf anwenden und das Ergebnis nur uebernehmen, wenn kein '├' (bzw. kein Zeichen aus ─-▟) mehr im Ergebnis steht. Alternativ am Ende von _repair_mojibake jedes uebrig gebliebene, aus einem '+' entstandene '├' wieder zu '+' zuruecksetzen.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `10_dateinamen_bereinigen.py` Zeile 1150  `[BEHOBEN]`

**Erfolgreiche Umbenennung wird als ERROR gemeldet und die CSV-Ruecknahmezeile fehlt, weil log_rename() innerhalb des Rename-try-Blocks liegt und das CSV mit striktem UTF-8 geoeffnet ist.**

```
log_rename("UMBENANNT", is_dir, directory, basename, new_name)
```

*Begruendung:* _rename_log_handle wird in Zeile 229-231 mit encoding="utf-8-sig" (Fehlerbehandlung strict) geoeffnet; der Schreibvorgang in Zeile 293 ist nur mit 'except OSError' abgesichert. Ein Dateiname mit halber UTF-16-Surrogatpaar-Haelfte (\udc80) - genau die Sorte, deren Existenz das Skript in Zeile 1009-1012 selbst dokumentiert - erzeugt beim CSV-Schreiben einen UnicodeEncodeError. Der ist eine Unterklasse von ValueError, NICHT von OSError, laeuft also am except vorbei. Weil log_rename() in Zeile 1150 noch INNERHALB des try-Blocks ab Zeile 1144 steht, faengt ihn 'except Exception as e' in Zeile 1181 ab. Ergebnis: os.rename ist bereits gelaufen, die Datei traegt den neuen Namen, aber das Skript meldet 'Umbenennung fehlgeschlagen', zaehlt ERROR statt RENAMED und schreibt keine CSV-Zeile. Damit fehlt fuer genau diese Umbenennung die einzige Rueckabwicklungsquelle, die das Skript in Zeile 1540-1541 dem Anwender zusichert.

*Auswirkung:* Nachgestellt in einem Temp-Verzeichnis: Datei 'D─nderungen\udc80.txt' -> nach dem Lauf liegt auf der Platte 'DÄnderungen\udc80.txt' (Codepunkte geprueft: 0x44 0xC4 ...), die CSV enthaelt nur die Kopfzeile, die Statistik meldet 'Bereinigt: 0 / Fehler: 1', Exitcode 1. Eine Umbenennung ohne Protokolleintrag ist nicht mehr rueckgaengig zu machen.

*Vorschlag:* CSV in Zeile 229-231 mit errors="surrogatepass" (oder "replace") oeffnen, das except in Zeile 294 auf 'except Exception' erweitern und log_rename()/logger.info() aus dem try-Block heraus hinter _rename_with_retry legen, damit ein Protokollfehler nie eine ausgefuehrte Umbenennung als Fehler umdeutet. Analog fuer den Race-Zweig (Zeile 1173) und den logging.FileHandler in Zeile 308.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `1_temp_dateien_entfernen.py` Zeile 154  `[BEHOBEN]`

**Die Suffix-Regel '.sdf' loescht unbedingt und ohne Rueckfrage Datenbank- bzw. Geodaten-Dateien, die keine Temp-Dateien sind.**

```
(".sdf",          "dev"),
```

*Begruendung:* '.sdf' ist die Dateiendung von SQL-Server-Compact-Datenbanken und von Autodesk Spatial Data Files - beides vollwertige Nutzdaten, keine Zwischenstaende. Nur im Sonderfall eines Visual-Studio-Projektordners ist '<Projekt>.sdf' ein Cache. Die Regel greift jedoch ueber den gesamten Bestand: should_delete_file prueft ausschliesslich die Endung des Dateinamens, ohne Kontext. Dass in diesem Bestand Autodesk-Produkte im Einsatz sind, belegt die Regeltabelle selbst (.dwl, .dwl2, .sv$ ab Zeile 141). Die Datei macht ausserdem an anderer Stelle vor, wie mit riskanten Endungen umzugehen ist: .bak und .log stehen bewusst in _OPTIONAL_BAK_LOG_SUFFIXES und desktop.ini in _OPTIONAL_DESKTOP_INI, beide nur auf ausdruecklichen Wunsch (Kommentar Zeile 189-191). '.sdf' steht dagegen unbedingt in der Basisliste. Geloescht wird ueber os.remove (Zeile 360) - ohne Papierkorb, also unwiederbringlich. Dieselbe Pruefung verdienen die Nachbareintraege '.user' (Zeile 152), '.ncb' (153), '.class' (150), '.prv' (144, u. a. private Schluesseldateien) und '.syd' (137).

*Auswirkung:* Nachgestellt: should_delete_file('Sammlungsdatenbank.sdf', ...) liefert (True, 'Suffix: .sdf', 'dev'); ebenso 'Stadtplan_Karte.sdf'. Eine seit ueber 24 Stunden unveraenderte SQL-CE- oder Autodesk-SDF-Datei auf Q:\ wird im ECHT-Modus geloescht und ist weg. Ebenfalls bestaetigt: 'Inventar.user' -> True, 'Buchhaltung.ncb' -> True, 'Notizen.prv' -> True, 'Konfig.class' -> True.

*Vorschlag:* '.sdf' aus _SUFFIX_RULES_BASE entfernen. Wird der Visual-Studio-Cache wirklich gebraucht, praeziser fassen (Wildcard-Regel nur zusammen mit einer .sln/.vcxproj im selben Ordner) oder analog zu .bak/.log in eine optionale, abgefragte Liste verschieben. Die genannten Nachbareintraege in derselben Weise gegenpruefen.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `1_temp_dateien_entfernen.py` Zeile 768  `[BEHOBEN]`

**Ein Verzeichnis wird als 'abgeschlossen' vermerkt, sobald sein Zaehler kurzzeitig auf 0 faellt - also lange bevor der Walk den darunterliegenden Ast ueberhaupt betreten hat; die Wiederaufnahme ueberspringt danach den kompletten Teilbaum.**

```
rest = self.offen.get(verzeichnis, 1) - 1
            if rest <= 0:
                self.offen.pop(verzeichnis, None)
                schreiben = True
```

*Begruendung:* Der Zaehler 'offen' waechst erst beim Einreihen der einzelnen Datei (Zeile 1210) und faellt beim Fertigmelden (Zeile 1193). Er erreicht 0 immer dann, wenn gerade keine Datei dieses Verzeichnisses mehr in Arbeit ist - im Test bereits nach der ERSTEN Datei. Vor allem aber gilt: os.walk (topdown) liefert ein Verzeichnis IMMER vor seinen Unterverzeichnissen. Die letzte Datei eines Elternordners ist daher zwangslaeufig fertig, bevor der Walk die Unterordner ueberhaupt gelesen hat. Der Elternordner wird also strukturell - nicht nur bei unguenstigem Timing - zu frueh in die Fortschrittsdatei geschrieben. Beim naechsten Start greift dann Zeile 654-658 ('if done_dirs and schluessel in done_dirs: dirs[:] = []; continue') und schneidet den GESAMTEN Ast ab. Der Kommentar in Zeile 655/656 ('die Unterordner ... stehen ebenfalls im Vermerk') trifft nicht zu. Zusaetzlich wird der Vermerk am Ende des zweiten, scheinbar erfolgreichen Laufs durch fortschritt.aufraeumen() (Zeile 1256) geloescht - die Luecke ist danach nicht mehr erkennbar.

*Auswirkung:* Nachgestellt in C:\Users\BENUTZ~1\AppData\Local\Temp\claude\...\scratchpad\t_resume2.py: Baum mit 2 Dateien in der Wurzel und 60 Unterordnern a 3 Dateien (182 Dateien). Abbruch (Strg+C) beim Eintritt in sub40, ohne jede kuenstliche Verzoegerung. Die Fortschrittsdatei enthaelt danach 10 Eintraege, darunter das WURZELVERZEICHNIS. Zweiter Lauf (Wiederaufnahme): Meldung '1 abgeschlossene Verzeichnisse werden uebersprungen', 'Dateien geprueft: 0', 'Entfernt: 0' - der Walk bricht sofort ab. 153 der 182 Temp-Dateien bleiben unangetastet liegen, die Fortschrittsdatei wird geloescht, die Zusammenfassung meldet einen sauberen Durchlauf. Im ersten Test (t_resume.py) stand die Wurzel sogar dreimal, also nach jeder einzelnen Datei, im Vermerk.

*Vorschlag:* Der Vermerk darf nicht auf Verzeichnis-, sondern nur auf Teilbaum-Ebene gelten. Praktikabel: (a) 'erledigt' erst dann schreiben, wenn der Walk das Verzeichnis samt aller Unterverzeichnisse verlassen hat - dazu die Kinder je Verzeichnis mitzaehlen und den Elternzaehler erst freigeben, wenn alle Kinder abgeschlossen sind (bottom-up), oder (b) einfacher und robust: den Vermerk nur fuer BLATT-Verzeichnisse schreiben (dirs leer) und in Zeile 654 nicht 'dirs[:] = []' setzen, sondern lediglich die Dateien dieses Verzeichnisses ueberspringen und weiter absteigen. Variante (b) kostet nur den erneuten Verzeichnisdurchlauf, nie Dateien.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `2a_entferne_schutz_word.ps1` Zeile 996  `[BEHOBEN]`

**Die Verwaisten-Backup-Pruefung akzeptiert jedes Original mit Groesse > 0 Byte und loescht damit genau das Backup, das ein bei einem Abbruch TRUNKIERTES Original haette retten sollen.**

```
try { $origOk = ([System.IO.FileInfo]::new($orig)).Length -gt 0 } catch { $origOk = $false }
```

*Begruendung:* Regel 2 in Remove-OrphanedBackups prueft nur Existenz und Length > 0. Der Kommentar Z. 964-966 nennt als Schutzziel ausdruecklich den Fall, dass das Backup 'die EINZIGE intakte Kopie' ist, und Z. 2140-2142 beschreibt das ausloesende Ereignis woertlich: 'Crash mitten im Copy work->final trunkiert das Original'. [System.IO.File]::Copy($workFile, $finalDest, $true) (Z. 2158) truncated die Zieldatei beim Start und schreibt fortlaufend - ein harter Abbruch (Stromausfall, Kill, BSOD) hinterlaesst also typischerweise ein TEILWEISE geschriebenes Original mit Groesse > 0, nicht 0 Byte. Genau dieser Zustand besteht Regel 2. Regel 3 (Mindestalter 24 h) verzoegert das nur; beim naechsten Lauf am Folgetag greift der Loeschzweig Z. 1012-1016. Ein Groessen-/Inhaltsvergleich zwischen Backup und Original findet nirgends statt.

*Auswirkung:* Lauf am Montag bricht beim Zurueckschreiben von Q:\Vertraege\Vertrag.docx (2,4 MB) hart ab; auf der Ablage liegen danach Vertrag.docx mit 380 KB (abgeschnitten, aber > 0) und Vertrag.docx.bak mit den vollstaendigen 2,4 MB. Lauf am Mittwoch: Regel 1 ok, Regel 2 ok (380 KB > 0), Regel 3 ok (> 24 h) -> [System.IO.File]::Delete loescht das Backup. Die vollstaendige Fassung ist unwiederbringlich weg, gemeldet wird lediglich 'Backups: 1 entfernt'.

*Vorschlag:* Regel 2 verschaerfen: Backup nur loeschen, wenn das Original mindestens so gross ist wie das Backup (bei .doc->.docx-Konvertierung ist der Vergleich nicht 1:1 moeglich - dann konservativ behalten und als Orphan melden). Mindestens: bei ([System.IO.FileInfo]::new($orig)).Length -lt $fi.Length das Backup in die OrphanList aufnehmen und NICHT loeschen.


## `2a_entferne_schutz_word.ps1` Zeile 1207  `[BEHOBEN]`

**Das fuehrende Komma in der -ArgumentList verschachtelt das Passwort-Array; im Job werden alle Passwoerter zu EINEM leerzeichen-verbundenen String zusammengezogen, die einzelnen Passwoerter werden nie probiert.**

```
} -ArgumentList $SourcePath, $DestPathBase, (,$Passwords), $pidFile,
```

*Begruendung:* (,$Passwords) erzeugt ein 1-Element-Array, das das string[] umschliesst. Als Element einer aeusseren Array-Literal-Liste wird es NICHT entrollt, also empfaengt der Job-Parameter [object[]]$pws ein Array mit genau einem Element - dem inneren Array. Die Schleife 'foreach ($p in $pws) { $tryList.Add([string]$p) }' (Z. 1107-1109 bzw. 1278-1280) iteriert daher nur EINMAL, und der Cast [string] auf ein string[] liefert die leerzeichen-verbundene Darstellung. Nachgestellt mit exakt demselben Konstrukt: bei drei Passwoertern enthaelt tryList ['Ablage1 Archiv2 Depot3'] und ['']; ohne das fuehrende Komma korrekt ['Ablage1'],['Archiv2'],['Depot3']. Ironischerweise soll der Kommentar Z. 1095-1104 genau dieses Problem verhindern - der dort ergaenzte [string]-Cast unterdrueckt nur die Ausnahme und macht aus dem Absturz eine stille Fehlfunktion. Bei genau EINEM Passwort faellt der Fehler nicht auf ([string] auf ein 1-Element-Array ergibt zufaellig das richtige Passwort), deshalb ueberlebt der Fehler jeden Ein-Passwort-Test. Bei leerer Liste entsteht zusaetzlich ein doppelter Leerpasswort-Versuch (tryList = ['','']), also pro Datei bis zu 2 unnoetige COM-Open-Versuche (im Convert-Pfad wegen der Repair-Eskalation 4 statt 2).

*Auswirkung:* Der Anwender gibt im Startdialog wie vorgesehen 2 oder 3 Passwoerter ein. Eine mit 'Archiv2' verschluesselte Datei wird nur mit dem Phantom-Passwort 'Ablage1 Archiv2 Depot3' und mit '' geoeffnet - beides scheitert. Ergebnis: SKIP (Passwort) bzw. SKIP (Verschluesselt), $stats.Encrypted++, die Datei bleibt geschuetzt. Betroffen sind Convert-DocToDocx (Z. 1207) UND Remove-OpenPassword (Z. 1330), also der komplette Passwort-Zweig des Skripts, sobald mehr als ein Passwort eingegeben wird.

*Vorschlag:* Das fuehrende Komma an beiden Stellen entfernen: '-ArgumentList $SourcePath, $DestPathBase, $Passwords, $pidFile, ...' (Z. 1207) und '-ArgumentList $FilePath, $Passwords, $OriginalExtension, $pidFile' (Z. 1330). Nachgestellt: damit binden alle Passwoerter einzeln, und auch bei LEERER Passwortliste verschieben sich die Folgeparameter NICHT ($origExt und $pidFile kommen korrekt an) - die Sorge, die das Komma motiviert hat, trifft hier nicht zu. Alternativ das Wrapper-Array im Job aufloesen. 2b und 2c pruefen, dort steht dasselbe Konstrukt.

*Vom Pruefer ausgefuehrt nachgestellt.*

*Bei der Korrektur zusaetzlich erledigt:* Der Vorschlag verlangt, 2b und 2c
auf dasselbe Konstrukt zu pruefen. In `2c` gibt es keine Passwortbehandlung.
In `2b` steht es zweimal (Zeile 1389 und 1486) und wirkt dort sogar haerter:
weil der `[string]`-Cast fehlt, ist das erste Element von `$tryList` eine
ArrayList statt eines Strings - nachgestellt `Typen=[ArrayList, String]` statt
vier Strings. Das ist genau der COM-Fehler 'ArrayList kann nicht in Object
konvertiert werden', den der Kommentar in 2a Z. 1095-1104 beschreibt. Beide
Stellen in 2b sind mitkorrigiert.

Nachgestellt wurde ausserdem die im Befund genannte Gegenprobe: bei LEERER
Passwortliste verschieben sich die Folgeparameter NICHT - `$origExt` und
`$pidFile` kommen in beiden Fassungen korrekt an. Die Sorge, die das
fuehrende Komma motiviert hat, trifft also nicht zu.



## `2a_entferne_schutz_word.ps1` Zeile 1729  `[BEHOBEN]`

**Unter -WhatIf legt New-Item den Arbeitsordner nicht an; der dokumentierte Simulationslauf bricht deshalb ab (-NoInteractive) bzw. meldet jede einzelne Datei als Fehler und laeuft ~10 s pro Datei ins Leere.**

```
New-Item -Path $TempPath -ItemType Directory -Force | Out-Null
```

*Begruendung:* Das Skript deklariert [CmdletBinding(SupportsShouldProcess)]. Beim Aufruf mit -WhatIf setzt PowerShell $WhatIfPreference=$true fuer den gesamten Skript-Scope, und JEDES nachgelagerte ShouldProcess-faehige Cmdlet gehorcht ihm - also auch New-Item (Z. 1729), Set-Content (Z. 1735/1737/1739) und Add-Content in Write-Log/Write-DetailedLog/Write-CsvLog. Nachgestellt mit einem Minimalskript gleicher Bauart: 'TempPath existiert: False', 'Logdatei existiert: False'. Folge: Test-WordTrustCenter (Z. 1750) laesst New-MinimalDocx in den nicht existierenden Ordner schreiben; [System.IO.Compression.ZipFile]::Open in einem fehlenden Verzeichnis wirft DirectoryNotFoundException (nachgestellt) und liefert Ok=$false. Confirm-Write selbst funktioniert korrekt - es wird nur nie erreicht.

*Auswirkung:* Aufruf '2a_entferne_schutz_word.ps1 -TargetPath Q:\ -NoInteractive -WhatIf': Smoke-Test schlaegt fehl, Z. 1764-1768 beendet mit exit 2 - keine einzige Datei wird geprueft. Interaktiv mit 'j' auf die Smoke-Test-Warnung: pro Datei scheitert [System.IO.File]::Copy($srcLong, $tempFile, $true) (Z. 1970) mit DirectoryNotFoundException, Invoke-WithRetry wiederholt 7-mal (200+400+800+1600+3200+4000 ms ≈ 10 s), danach aeusserer catch -> $stats.Errors++. Ergebnis eines Simulationslaufs ueber 20.000 Dateien: 'Wuerde aendern: 0', 'Fehler: 20000', Laufzeit ~55 Stunden, und wegen der ebenfalls unterdrueckten Set-Content/Add-Content-Aufrufe existiert weder Log noch CSV zum Nachlesen. Genau der im Kopfkommentar Z. 77-79 empfohlene erste Lauf auf einer neuen Ablage liefert damit ein wertloses Ergebnis.

*Vorschlag:* Infrastruktur-Operationen vom WhatIf-Mechanismus ausnehmen: Arbeitsordner mit [System.IO.Directory]::CreateDirectory($TempPath) anlegen und die drei Protokollkoepfe mit [System.IO.File]::WriteAllText / AppendAllText schreiben (bzw. New-Item/Set-Content/Add-Content mit -WhatIf:$false aufrufen). Die eigentliche Freigabe der schreibenden Dateioperationen bleibt bei Confirm-Write.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `2a_entferne_schutz_word.ps1` Zeile 2091  `[BEHOBEN]`

**Ein erfolgreich entferntes Oeffnen-Passwort wird verworfen, wenn die Datei sonst keinen XML-Schutz enthaelt - das Original bleibt verschluesselt, wird aber als 'Kein Schutz' protokolliert.**

```
if (@($actions).Count -gt 0 -or $wasConverted) {
```

*Begruendung:* Die Rueckschreib-Bedingung fragt nur die XML-Aktionen und die Formatkonvertierung ab. Das Ergebnis der COM-Passwortentfernung ($pwRemoved, Z. 2072) geht NICHT in die Bedingung ein - die Variable wird nach Z. 2074 nirgends mehr verwendet (per Grep ueber die ganze Datei: nur Z. 2072/2073/2074). Remove-OpenPassword hat die Entschluesselung aber nur auf der Temp-Kopie ($workFile) durchgefuehrt ($doc.Password=''; $doc.Save(), Z. 1316-1318). Findet Remove-WordProtection danach keine documentProtection/writeProtection/formProt (der Normalfall: die Datei war ja nur mit einem Oeffnen-Passwort verschluesselt, nicht zusaetzlich formulargeschuetzt), ist $actions leer und $wasConverted ist fuer .docx/.docm ohnehin $false. Damit laeuft der else-Zweig Z. 2248-2251, die entschluesselte Temp-Datei wird im finally geloescht (Z. 2283-2285).

*Auswirkung:* Eine mit Oeffnen-Passwort verschluesselte Vertrag.docx, deren Passwort bekannt und korrekt eingegeben ist: Word oeffnet und entschluesselt die Temp-Kopie, Z. 2083 protokolliert 'Oeffnen-Passwort entfernt', danach wird die entschluesselte Kopie kommentarlos geloescht. Die Datei auf der Ablage bleibt verschluesselt, die CSV meldet Status NO_CHANGE / 'Kein Schutz gefunden' und die Endstatistik zaehlt sie unter Geprueft, nicht unter Entsperrt. Der Betreiber haelt die Ablage danach fuer passwortfrei, obwohl sie es nicht ist - der Fehler faellt erst im Folgeschritt (3a/4a) auf.

*Vorschlag:* Erfolgsflag mitfuehren und in die Bedingung aufnehmen, z. B. $pwRemoved vor dem try auf $false initialisieren und Z. 2091 zu 'if (@($actions).Count -gt 0 -or $wasConverted -or $pwRemoved)' erweitern; $actionStrPre entsprechend um 'Oeffnen-Passwort' ergaenzen (Z. 2093-2096), damit Log und CSV den Grund korrekt ausweisen. 2b/2c auf dieselbe Bedingung pruefen.


## `2b_entferne_schutz_excel.ps1` Zeile 614  `[BEHOBEN]`

**Die 24-Stunden-Schutzfrist fuer Backups greift nie: File.Copy uebernimmt die LastWriteTime des Originals, ein Sekunden altes .bak gilt sofort als verwaist und wird geloescht.**

```
if ($fi.LastWriteTime -gt $cutoff) {
```

*Begruendung:* Das Backup entsteht in Zeile 2323 mit [System.IO.File]::Copy($srcLong, $backupPath, $true). File.Copy uebertraegt die LastWriteTime der Quelle auf das Ziel - nur die CreationTime wird neu gesetzt. Remove-OrphanedBackups vergleicht aber genau diese uebernommene LastWriteTime gegen $cutoff = (Get-Date).AddHours(-24) aus Zeile 589. Bei Archiv-Ablagen ist die LastWriteTime der Dokumente typischerweise Jahre alt, das Backup erbt sie. Nachgestellt: Original auf LastWriteTime 08/14/2020 gesetzt, Backup um 08/14/2026 01:34:57 per File.Copy erzeugt -> 'Backup LastWriteTime : 08/14/2020', 'Backup CreationTime : 08/14/2026', Ergebnis der Skriptpruefung: 'WIRD GELOESCHT - obwohl das Backup Sekunden alt ist'. Die Frist ist damit wirkungslos; der Zweig $res.Kept ist praktisch toter Code. Verschaerfend ist die zweite Bedingung in Zeile 605: $origOk = Length -gt 0. Ein durch einen Abbruch mitten im Copy work->final trunkiertes Original ist groesser als 0 Byte und gilt damit als 'intakt' (nachgestellt: True) - obwohl der Kommentar in Zeile 2309-2312 genau diesen Fall als Grund nennt, ein altes .bak keinesfalls zu ueberschreiben, und der Funktionskopf in Zeile 580-581 verspricht: 'Fehlt das Original oder ist es leer, kann das Backup die einzige intakte Kopie sein'.

*Auswirkung:* Zwei belegbare Datenverlustpfade. (1) Paralleler Lauf: Instanz B startet und laeuft in Zeile 1955 Remove-OrphanedBackups ueber dieselbe Ablage, waehrend Instanz A zwischen Zeile 2323 (Backup erzeugt) und 2336 (Backup geloescht) steht. B loescht A's laufendes .bak, weil dessen LastWriteTime das Alter des Dokuments traegt. Scheitert danach A's Copy work->final, findet der Rollback in Zeile 2377 kein Backup mehr - die Bedingung [System.IO.File]::Exists($backupPath) ist falsch, es wird gar kein Rollback versucht, und das teilweise ueberschriebene Original bleibt beschaedigt zurueck. Parallelbetrieb ist vom Skript ausdruecklich vorgesehen (PID-eigener Temp-Ordner; Reserve-UniqueDestination begruendet die atomare Reservierung in Zeile 845-848 mit 'ein parallel laufender Durchgang'). (2) Einzellauf nach hartem Abbruch: Stromausfall mitten im Copy work->final hinterlaesst ein trunkiertes Original und ein intaktes .bak. Der naechste Lauf sieht Length > 0 (origOk) und ein 'aelteres' Backup - und loescht die einzige unversehrte Kopie in Zeile 621.

*Vorschlag:* Fuer die Altersfrist die CreationTime des Backups verwenden statt der LastWriteTime, da nur sie den echten Erzeugungszeitpunkt traegt (nachgestellt korrekt auf 'jetzt' gesetzt): if ($fi.CreationTime -gt $cutoff). Zusaetzlich das Backup gegen das Original abgleichen, statt nur Length -gt 0 zu pruefen - etwa Backup behalten, sobald seine Groesse von der des Originals abweicht, oder eine Pruefsumme vergleichen. Fuer den Parallelfall die Backups des laufenden Prozesses kenntlich machen (z. B. .bak_<PID>) und in Remove-OrphanedBackups Backups lebender PIDs ueberspringen, analog zu Remove-StaleTempFolders in Zeile 760-767.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `2b_entferne_schutz_excel.ps1` Zeile 1867  `[BEHOBEN]`

*Nachtrag 14.08.2026 - die Markierung [BEHOBEN] war nur zur Haelfte richtig.*
Ersetzt worden war lediglich `New-Item` durch
`[System.IO.Directory]::CreateDirectory` (Zeile 1923). Die im Befund
ausdruecklich mitgenannten `Remove-Item`-Aufraeumpfade waren weiterhin
ungeschuetzt und unter -WhatIf unterdrueckt, so dass Arbeitskopien saemtlicher
Excel-Dateien im Dokumentenordner liegen geblieben waeren. Alle zehn Stellen
sind jetzt mit `-WhatIf:$false -Confirm:$false` versehen. Dieselbe Luecke
bestand unbemerkt auch in `2c` (New-Item Zeile 1609 plus zehn Remove-Item) -
dort ebenfalls behoben, sonst waere der Smoke-Test im Simulationslauf
gescheitert und das neu eingezogene Confirm-Write aus Befund 2120 nie
erreicht worden.

**Zweiter, unabhaengiger Fehler in derselben Zeile: unter -WhatIf unterdrueckt ShouldProcess das Anlegen des Arbeitsordners, damit ist der dokumentierte Simulationslauf funktionsunfaehig.**

```
New-Item -LiteralPath $TempPath -ItemType Directory -Force | Out-Null
```

*Begruendung:* Das Skript deklariert [CmdletBinding(SupportsShouldProcess)]. Wird es mit -WhatIf aufgerufen, setzt PowerShell $WhatIfPreference im Skript-Scope auf $true, und dieser Wert vererbt sich an jedes aufgerufene Cmdlet, das ShouldProcess unterstuetzt - New-Item gehoert dazu. Nachgestellt mit einer minimalen Advanced Function: ohne -WhatIf 'TempPath angelegt? = True', mit -WhatIf 'WhatIf: Ausfuehren des Vorgangs "Verzeichnis erstellen" ...' und 'TempPath angelegt? = False'. Dieser Fehler bleibt bestehen, auch wenn -LiteralPath auf -Path korrigiert wird; er ist damit nicht dieselbe Ursache. Der Skriptkopf empfiehlt -WhatIf ausdruecklich: 'Empfohlen fuer den ersten Lauf auf einer neuen Ablage', und Confirm-Write wurde laut Kommentar in Zeile 279-285 eigens nachgeruestet, damit -WhatIf ueberhaupt wirkt.

*Auswirkung:* Ein Simulationslauf 'ps1 -TargetPath Q:\ -WhatIf' legt $TempPath nicht an. Jede Datei scheitert in Zeile 2132 an der Arbeitskopie (nachgestellt: 'Copy src->temp = FEHLER: Ein Teil des Pfades ... konnte nicht gefunden werden'), landet im catch in Zeile 2411 als 'Allg. Fehler' und wird in der CSV als ERR statt als WHATIF protokolliert. Die Statistik meldet am Ende 'Wuerde aendern: 0' und 'Fehler: <alle Dateien>'. Der Anwender, der vor dem Echtlauf auf einer 200.000-Dateien-Ablage pruefen will, bekommt damit exakt null verwertbare Information - die Simulation, die den Schaden verhindern soll, faellt aus. Zusaetzlich sind unter -WhatIf auch alle Remove-Item-Aufraeumpfade unterdrueckt (nachgestellt: 'TempPath nach Remove-Item noch da? = True'), betrifft Zeile 2185, 2421 und 2451 - Arbeitskopien saemtlicher Excel-Dateien bleiben im Dokumentenordner liegen.

*Vorschlag:* Die Infrastruktur-Operationen von der Simulation ausnehmen: [System.IO.Directory]::CreateDirectory($TempPath) verwenden (provider-frei, kein ShouldProcess) bzw. bei Cmdlets explizit -WhatIf:$false -Confirm:$false setzen. Gleiches gilt fuer die Remove-Item-Aufrufe auf $TempPath/$tempFile/$workFile in Zeile 1865, 1922, 1933, 2185, 2421 und 2451 - sie raeumen ausschliesslich eigene Temp-Artefakte auf und duerfen nie simuliert werden. Fachlich schreibende Stellen bleiben wie bisher ueber Confirm-Write gesteuert.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `2c_entferne_schutz_powerpoint.ps1` Zeile 2120  `[BEHOBEN]`

**Der Zone.Identifier-Zweig ersetzt das Original ohne Confirm-Write - bei -WhatIf wird trotz Simulation geloescht und ueberschrieben.**

```
try {
                    Invoke-WithRetry -Action { [System.IO.File]::Copy($srcLong, $zoneBak, $true) }
                    [System.IO.File]::Delete($srcLong)
                    Invoke-WithRetry -Action { [System.IO.File]::Copy($workFile, $srcLong, $true) }
```

*Begruendung:* Jede andere schreibende Stelle im Skript geht durch Confirm-Write (Z. 972 Backup loeschen, Z. 1779 Sperrdatei loeschen, Z. 1800 Schreibschutz-Attribut, Z. 1979 Rueckschreiben). Der else-Zweig ab Z. 2102 - er greift, wenn Remove-PptxProtection nichts gefunden hat UND nicht konvertiert wurde, die Temp-Kopie aber einen Zone.Identifier hatte - ruft Confirm-Write ueberhaupt nicht auf. Damit ist $PSCmdlet.ShouldProcess wirkungslos, und die von Z. 55-57 des Skriptkopfs ausdruecklich zugesicherte Eigenschaft ('-WhatIf zeigt an, welche Dateien geaendert wuerden, ohne sie anzufassen') gilt fuer diesen Zweig nicht. Zusaetzlich wird in Z. 2134 $stats.Unlocked++ statt $stats.WouldChange++ gezaehlt, die Statistik des Probelaufs ist also ebenfalls falsch. Nachgestellt mit einem Struktur-Nachbau (Confirm-Write + identischer Verzweigung) unter powershell.exe 5.1: Aufruf mit -WhatIf gab 'ZWEIG 2 (Zone.Identifier): Original ERSETZT' aus, der Dateiinhalt wechselte von 'ORIGINAL-INHALT' auf 'BEARBEITETE-KOPIE'.

*Auswirkung:* Empfohlener erster Probelauf auf einer fremden Archiv-Ablage: '2c_entferne_schutz_powerpoint.ps1 -TargetPath Q:\ -WhatIf'. Fuer jede .pptx ohne Schutz, die einmal aus dem Internet/Mail heruntergeladen wurde (Zone.Identifier-ADS ist auf solchen Ablagen sehr haeufig), wird das Original tatsaechlich geloescht und durch die Temp-Kopie ersetzt - waehrend der Anwender glaubt, nichts werde angefasst. Reisst die Kopie in Z. 2121 ab (Netzhaenger, AV), haengt die Rettung am Rollback ab Z. 2151; scheitert auch der, ist die Datei weg (Log: 'KRITISCH: Rollback fehlgeschlagen').

*Vorschlag:* Vor Z. 2118 dieselbe Freigabe wie im Hauptzweig einziehen, z. B.: if (-not (Confirm-Write $filePath 'Zone.Identifier entfernen')) { $stats.WouldChange++; Write-CsvLog -Status 'WHATIF' -Actions 'Wuerde Zone.Identifier entfernen' -Path $filePath; return }. Der Block ab Z. 2106 (Stale-.bak beiseite legen) muss ebenfalls hinter diese Pruefung, sonst benennt schon die Simulation ein vorhandenes .bak um.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `3b_xls_xlsx_auf_neueste_Version_aktualisieren.py` Zeile 575  `[BEHOBEN]`

**_kill_orphaned_excel() beendet jede Excel-Sitzung, die der Anwender waehrend des Laufs oeffnet - geschuetzt sind nur die beim Start bereits laufenden PIDs.**

```
if is_foreign_excel_pid(proc.info["pid"]):
                continue
            try:
                proc.kill()
```

*Begruendung:* _FOREIGN_EXCEL_PIDS wird genau einmal vor dem Start gefuellt (snapshot_foreign_excel_pids, Zeile 3789). Jede Excel-Instanz, die spaeter entsteht und nicht in diesem Schnappschuss steht, gilt als 'verwaist' und wird mit proc.kill() hart beendet - ohne Speichern-Rueckfrage. _kill_orphaned_excel() laeuft nicht nur einmal, sondern bei jedem periodischen Neustart (Zeile 3758, alle EXCEL_RESTART_EVERY=200 Dateien), bei jedem Pre-Clean-Neustart (Zeile 2799, 2884), im Signal-Handler (Zeile 685) und im finally (Zeile 4164). Bei einem Lauf ueber Q:\, der Stunden dauert, ist das faktisch ein Dauerzustand. Die Startwarnung sagt dem Anwender ausdruecklich das Gegenteil zu (Zeile 779: 'Ihre Sitzung wird vom Skript NICHT beendet') und warnt nirgends davor, waehrend des Laufs Excel zu oeffnen. Dabei ist die eigene Instanz eindeutig bekannt (excel_app_pid), ein pauschaler Kill also nicht noetig.

*Auswirkung:* Anwender startet den Lauf, oeffnet 20 Minuten spaeter eine Kalkulation und tippt darin. Beim naechsten periodischen Excel-Neustart (nach 200 Dateien) wird seine Sitzung per proc.kill() beendet; ungespeicherte Aenderungen sind weg - genau der Datenverlust, den der Kommentar in Zeile 519-526 abstellen wollte.

*Vorschlag:* Nur die eigene Instanz beenden (Kill ausschliesslich auf excel_app_pid, verifiziert ueber create_time()), oder _FOREIGN_EXCEL_PIDS vor jedem Kill-Zyklus um alle PIDs erweitern, die nicht die eigene sind. Zusaetzlich die Startwarnung um den Hinweis ergaenzen, waehrend des Laufs kein Excel zu oeffnen.


## `3b_xls_xlsx_auf_neueste_Version_aktualisieren.py` Zeile 607  `[BEHOBEN]`

**Die Stale-Lock-Erkennung erkennt einen laufenden PyInstaller-Build nicht als lebenden Prozess und loescht dessen Lock - parallele Laeufe auf demselben Verzeichnis werden erlaubt.**

```
if "EXCEL" in p.name().upper() or "PYTHON" in p.name().upper():
                        return None  # echter laufender Prozess
```

*Begruendung:* Der Prozessname des laufenden Skripts wird gegen die Muster 'EXCEL' und 'PYTHON' geprueft. Im dokumentierten Auslieferungsweg (Header Zeile 48-53: PyInstaller --onefile --name "3b_xls_xlsx_auf_neueste_Version_aktualisieren") heisst der Prozess '3B_XLS_XLSX_AUF_NEUESTE_VERSION_AKTUALISIEREN.EXE' und enthaelt weder 'EXCEL' noch 'PYTHON'. Der Zweig faellt daher durch auf os.remove(lock_path) (Zeile 612) und protokolliert 'Stale-Lock entfernt', obwohl der Lock-Halter lebt. Anschliessend schreibt der zweite Lauf sein eigenes Lock und startet. Zusaetzlich loescht der zuerst fertige Lauf im finally (Zeile 4184, _release_lock) die Lock-Datei des anderen.

*Auswirkung:* Zwei per Aufgabenplanung gestartete .exe-Laeufe auf Q:\ laufen gleichzeitig. Beide indizieren dieselben Dateien; Lauf A ersetzt gerade Bericht.xls -> Bericht.xlsx, waehrend Lauf B dieselbe Bericht.xls noch geoeffnet hat und sein eigenes Konvertat darueberschreibt bzw. das inzwischen entfernte Original nicht mehr findet. Ergebnis: Sharing-Violations, doppelte Ausweichnamen (_2, _3) und im ungeguenstigen Fall eine Zieldatei aus einer aelteren Quellversion.

*Vorschlag:* Nicht ueber den Prozessnamen pruefen, sondern ueber Prozess-Identitaet: PID + create_time() in die Lock-Datei schreiben und beim Stale-Check vergleichen; alternativ das Lock ueber einen exklusiv gehaltenen Dateihandle (msvcrt.locking / CreateFile ohne Share-Flags) fuer die gesamte Laufzeit halten.


## `3b_xls_xlsx_auf_neueste_Version_aktualisieren.py` Zeile 2522  `[BEHOBEN]`

**SaveAs mit Password="" entfernt die Oeffnungs-Verschluesselung: mit hinterlegtem Passwort geoeffnete Mappen werden unverschluesselt zurueckgeschrieben.**

```
Password             = "",
                WriteResPassword     = "",
```

*Begruendung:* Fuer verschluesselte Dateien wird passwords_to_try = passwords gesetzt (Zeile 2683) und die Mappe mit dem passenden Kennwort geoeffnet (Zeile 2746). Der anschliessende SaveAs setzt Password und WriteResPassword hart auf den Leerstring - das ist in Excel-COM kein 'unveraendert lassen', sondern 'ohne Kennwort speichern'. Die so entstandene Klartextdatei ersetzt danach das Original (robust_move, Zeile 3207; safe_remove des Originals, Zeile 3216). Nirgends wird der Anwender darauf hingewiesen: ask_passwords() (Zeile 1382-1405) spricht nur davon, dass Passwoerter 'der Reihe nach ausprobiert' werden, und die Konfigurationsausgabe behauptet sogar das Gegenteil - Zeile 4012: 'Passwortdateien:       Ueberspringen (als SKIPPED loggen)', direkt gefolgt von 'Passwoerter: N hinterlegt'. Der Header dokumentiert als destruktives Verhalten ausschliesslich Druckbereiche/Namen (Zeile 7-17), nicht den Verlust der Verschluesselung.

*Auswirkung:* Eine mit Oeffnungskennwort geschuetzte Personal-/Haushaltsmappe auf Q:\ wird beim Lauf mit --password (oder interaktiv hinterlegtem Passwort) entschluesselt und ersetzt das verschluesselte Original. Danach kann jeder mit Leserecht auf der Freigabe die Inhalte oeffnen - Vertraulichkeitsverlust, nicht rueckgaengig zu machen (das Backup wird bei Erfolg in Zeile 3233 geloescht).

*Vorschlag:* Das beim Oeffnen erfolgreiche Kennwort merken und beim SaveAs wieder mitgeben (Password=pw_erfolgreich), oder verschluesselte Dateien wie angekuendigt konsequent als SKIPPED behandeln. Mindestens: explizite Warnung + Bestaetigung vor dem Lauf und korrekte Konfigurationsausgabe.

*Bei der Korrektur mit echtem Excel nachgestellt (14.08.2026):* Die
Kernbehauptung haelt. Eine mit Kennwort erzeugte .xlsx wurde mit dem Kennwort
geoeffnet und dann zweimal gespeichert. Mit `SaveAs(Password="")` kam sie
UNVERSCHLUESSELT heraus und liess sich ohne Kennwort oeffnen (Inhalt im Klartext
lesbar); mit `SaveAs(Password=<Kennwort>)` blieb die Verschluesselung erhalten.
`Password=""` ist in Excel-COM also tatsaechlich 'ohne Kennwort speichern' und
nicht 'unveraendert lassen'.

*Behoben:* Das beim Oeffnen erfolgreiche Kennwort wird an allen drei
Oeffnungspfaden (Normal, Repair, Hail-Mary-Re-Open) festgehalten und an
`_save_as_workbook` durchgereicht. Zusaetzlich die widerspruechliche
Konfigurationsausgabe korrigiert - sie behauptete unabhaengig von der Lage
'Passwortdateien: Ueberspringen' und nannte in der Zeile darunter die Zahl der
hinterlegten Kennwoerter.



## `3b_xls_xlsx_auf_neueste_Version_aktualisieren.py` Zeile 3686  `[BEHOBEN]`

**Der Probelauf (--dry-run) loescht Dateien auf der Freigabe: clean_temp ist fest auf True verdrahtet und dry_run wird nicht durchgereicht.**

```
gen = file_generator(directory, clean_temp=True,
                         exclude_patterns=exclude_patterns)
```

*Begruendung:* process_directory ruft file_generator IMMER mit clean_temp=True auf, unabhaengig vom Parameter dry_run. In file_generator (Zeile 2210-2223) werden daraufhin alle Dateien, deren Name mit '~$' oder '._' beginnt, per safe_remove() geloescht - ohne Endungsfilter, also im gesamten Baum und fuer JEDEN Dateityp ('._Foto.jpg', '._Bericht.docx'). Auf NAS-/SMB-Freigaben mit Mac-Clients sind '._*'-Dateien AppleDouble-Container (Finder-Metadaten, Resource-Forks), keine Excel-Reste. Das widerspricht der an drei Stellen ausdruecklich gegebenen Zusage des Probelaufs: Zeile 3726 'Der Probelauf laesst das Dateisystem vollstaendig unangetastet', Zeile 3840 'SPEICHERT aber NICHTS ... ohne eine Datei zu aendern', Zeile 4190 'PROBELAUF ABGESCHLOSSEN (es wurde NICHTS geaendert)'. Der Probelauf ist genau der Modus, den ein Anwender vor dem ersten Echt-Lauf auf einer fremden Ablage waehlt (Zeile 3973).

*Auswirkung:* Aufruf 'python 3b_....py --dir Q:\ --dry-run' ueber eine Archiv-Freigabe mit Mac-Zugriff: alle nicht gesperrten '._*'-Dateien im gesamten Baum sind nach dem 'Probelauf' geloescht, Mac-Metadaten/Resource-Forks unwiederbringlich weg. Die Abschlussmeldung behauptet trotzdem 'es wurde NICHTS geaendert'.

*Vorschlag:* clean_temp=not dry_run uebergeben (bzw. dry_run bis in file_generator durchreichen) und den '._'-Praefix vom '~$'-Praefix trennen: '._' ist kein Office-Owner-File und sollte, wenn ueberhaupt, nur mit ausdruecklichem Schalter geloescht werden.


## `3c_ppt_pptx_auf_neueste_Version_aktualisieren.py` Zeile 2592  `[BEHOBEN]`

**Der Aufraeumzweig loescht als 'korruptes Fragment' eine fremde, vorbestehende Datei, die das Skript nie angelegt hat.**

```
else:
                safe_remove(target_path)
                pbar.write("    → Korruptes Dateifragment nach Abbruch sicher entfernt.")
```

*Begruendung:* In Zeile 2446 wird target_path auf den Namen mit neuer Endung gesetzt. Existiert dort bereits eine FREMDE Datei, soll Zeile 2448 (_resolve_unique_path) einen Ausweichnamen reservieren. Wirft diese Funktion (sie tut das explizit in Zeile 2117, wenn reserve_unique_path nach 999 + 10 Versuchen None liefert - alle os.open-Versuche scheitern z.B. bei vollem Zielvolume mit ENOSPC), wird Zeile 2454 nie erreicht: target_path zeigt weiterhin auf die fremde Datei. Im except-Zweig ist backup_path None (der Backup-Block ab 2461 wurde nie betreten) und update_complete False - also greift genau dieser else-Zweig und loescht die fremde Datei. Sie wurde vom Lauf weder erzeugt noch veraendert.

*Auswirkung:* Ordner enthaelt 'Fuehrung.ppt' und die davon unabhaengige 'Fuehrung.pptx'. Waehrend der Verarbeitung von Fuehrung.ppt laeuft das Zielvolume voll, alle Reservierungsversuche fuer 'Fuehrung_2.pptx' scheitern mit OSError, _resolve_unique_path wirft. Ergebnis: 'Fuehrung.pptx' wird geloescht (Loeschen funktioniert bei vollem Volume weiterhin), 'Fuehrung.ppt' bleibt unkonvertiert, Meldung 'Korruptes Dateifragment nach Abbruch sicher entfernt'. Datenverlust an einer Datei, die nie Gegenstand der Operation war.

*Vorschlag:* Ein Flag mitfuehren, das nur True wird, wenn target_path von diesem Lauf erzeugt bzw. reserviert wurde (target_was_reserved oder target_path == original_path oder converted_ok). Nur dann im except-Zweig safe_remove(target_path) ausfuehren; andernfalls die fremde Datei unangetastet lassen.

*Nachpruefung 14.08.2026:* Der Befund war echt, ist aber bereits durch die
Korrektur von Befund 2571 (KRITISCH) miterledigt - dort wurde ein Flag
`target_touched` eingefuehrt, das erst unmittelbar vor `robust_move` gesetzt
wird. In der hier beschriebenen Lage (Wurf aus `_resolve_unique_path`) ist es
`False`, der Aufraeumzweig greift also nicht mehr. Nachgestellt mit dem
Variablenstand genau dieser Lage: ALT 'LOESCHT target_path' (fremde Datei
weg), NEU 'nichts' (fremde Datei unangetastet). Gegenprobe mit einem ECHTEN
Fragment (`target_touched=True`): beide Fassungen raeumen weiterhin auf.
Keine zusaetzliche Aenderung noetig.



## `4a_ersetze_font_in_word.py` Zeile 3352  `[BEHOBEN]`

**Der Smoke-Test samt Rueckfrage laeuft auch im Modus --auto; ein geplanter Task haengt am Prompt oder beendet sich mit Exit-Code 0, ohne eine einzige Datei angefasst zu haben.**

```
smoke_ok, smoke_msg, smoke_cat = test_trust_center_smoke(timeout=25.0)
```

*Begruendung:* Der gesamte Smoke-Test-Block (Zeilen 3351-3394) ist NICHT durch 'if not auto_mode:' geschuetzt. Bei Fehlschlag laeuft Zeile 3386 'if not ask_yes_no("\nTrotzdem fortfahren? (Timeout-Risiko pro Datei!)", default_yes=False)' unbedingt. ask_yes_no() blockiert an input(). Die Schwesterskripte machen es anders: 4b_ersetze_font_in_excel.py kapselt denselben Block in Zeile 3534 mit 'if not auto_mode:', 3a_doc_docx... steigt in Zeile 3586-3589 im auto_mode explizit mit sys.exit(2) aus. Zusaetzlich widerspricht das der eigenen Hilfe in Zeile 3166: '--auto ... Trust-Center-Check und interaktive Rueckfragen ueberspringen'.

*Auswirkung:* Geplanter Task ruft '4a_ersetze_font_in_word.py --auto --dir Q:\'. Auf dem Server ist das Trust Center nicht konfiguriert -> Smoke-Test scheitert. Fall 1 (stdin am Taskplaner geschlossen): ask_yes_no faengt EOFError und liefert default_yes=False -> 'Abgebrochen.' + sys.exit(0) in Zeile 3390. Der Task meldet Erfolg (Exit 0), es wurde keine Datei verarbeitet, keine Statistik und keine Run-Summary geschrieben - der Fehler faellt tagelang nicht auf. Fall 2 (stdin an eine Konsole/Pipe gebunden, z.B. Start ueber einen cmd-Wrapper): der Task haengt unbegrenzt an der Rueckfrage.

*Vorschlag:* Block 3351-3394 analog 4b in 'if not auto_mode:' einfassen oder analog 3a vor Zeile 3386 einfuegen: 'if auto_mode: file_logger.error("Abbruch (auto_mode) nach fehlgeschlagenem Smoke-Test."); sys.exit(2)'. Exit-Code ungleich 0, damit der Taskplaner den Fehlschlag sieht.


## `4a_ersetze_font_in_word.py` Zeile 3497  `[BEHOBEN]`

**Das abschliessende 'Beliebige Taste'-input() ist als einziges Skript der Sammlung nicht gegen --auto abgesichert und faengt nur EOFError ab.**

```
input("\nBeliebige Taste drücken, um das Fenster zu schließen ...")
```

*Begruendung:* Zeile 3496-3499 laeuft unbedingt, auch mit --auto. Alle Schwesterskripte schuetzen dieselbe Stelle: 4b_ersetze_font_in_excel.py Zeile 3519 ('if not auto_mode:') und Zeile 3677 ('if _is_tty() and not auto_mode:'), 4c Zeile 2959, 3c Zeile 3251, 10_dateinamen_bereinigen.py Zeile 1552, die PS-Skripte ueber '-not $NoInteractive'. Dieselbe ungeschuetzte Stelle noch einmal in Zeile 3339 (Ausstieg nach dem Probelauf). Ausserdem wird nur EOFError abgefangen: startet der Task ohne Konsole (pythonw / Taskplaner 'Unabhaengig von der Benutzeranmeldung'), ist sys.stdin None und input() wirft RuntimeError, nicht EOFError.

*Auswirkung:* Nachgestellt: (a) 'python -c ... sys.stdin=None; input(...)' liefert 'RuntimeError: lost sys.stdin' - der except-EOFError-Zweig greift nicht, das Skript endet nach vollstaendig geleisteter Arbeit mit einem Traceback und Exit-Code 1. (b) Ist stdin an eine offene Pipe/Konsole gebunden, blockiert der Prozess dauerhaft. Weil die Einzelinstanz-Sperre (Zeile 3406-3411) erst per atexit freigegeben wird, blockiert der haengende Prozess zusaetzlich JEDEN weiteren geplanten Lauf ('bereits aktiv').

*Vorschlag:* Zeilen 3496-3499 und 3338-3341 in 'if not auto_mode:' einfassen und den except-Zweig auf '(EOFError, RuntimeError, OSError)' erweitern - wie in 4b Zeile 3677 ueber _is_tty() geloest.

*Vom Pruefer ausgefuehrt nachgestellt.*

*Bei der Korrektur zusaetzlich erledigt:* Der Befund nennt, dass nur EOFError
abgefangen wird. Dieselbe Schwaeche steckte in vier weiteren input()-Stellen
(Verzeichnisauswahl, Pfadeingabe, Schriftartabfrage, Zaehlmodus) - alle auf
`(EOFError, RuntimeError, OSError)` erweitert. Nachgestellt mit
`sys.stdin = None`: alte Fassung 'RuntimeError: lost sys.stdin' am
`except EOFError` vorbei, neue Fassung liefert sauber den Vorgabewert.
Die beiden 'Beliebige Taste'-Stellen laufen jetzt ueber einen gemeinsamen
Helfer `_warte_auf_taste(auto_mode)`, der zusaetzlich `_is_tty()` prueft -
wie es 4b vormacht.



## `4b_ersetze_font_in_excel.py` Zeile 2714  `[BEHOBEN]`

**Strg+C waehrend des Stage-2-Moves: Original bleibt als .bak liegen, die fertige Kopie wird im finally geloescht - der Wiederherstellungspfad wird uebersprungen.**

```
bak_orig = prepare_long_path(
                        original_path + f"_{uuid.uuid4().hex[:6]}.bak")
                    _av_safe_replace(safe_orig, bak_orig)
                _av_safe_move(temp_stage2_path, safe_orig)
```

*Begruendung:* _signal_handler (Zeile 497-531) endet mit sys.exit(130). Das loest SystemExit aus - SystemExit erbt von BaseException und wird von 'except Exception as e_s2_move' (Zeile 2728) NICHT gefangen, der dortige Restore aus dem .bak (Zeile 2733-2740) laeuft also nie. Der finally-Block laeuft dagegen sehr wohl und loescht in Zeile 3012-3016 die Stage-2-Datei mit den fertigen Aenderungen. Zusaetzlich raeumt der Signal-Handler ueber _safe_cleanup_temp() (Zeile 514) den gesamten TEMP_PROCESS_PATH per rmtree ab; die Stage-2-Datei heisst 'stage2_*' und steht weder unter RESCUE_ noch in _preserved_temp_files, ist also nicht geschuetzt. Das Zeitfenster ist nicht schmal: _av_safe_move wiederholt bei AV-Sperren bis zu 40 Mal (AV_MAX_RETRIES) mit bis zu 3 s Pause, also ueber 100 s pro Operation, zuzueglich der reinen Kopierzeit einer grossen Mappe auf die Freigabe. Der zurueckbleibende .bak-Name faellt ausserdem unter SKIP_FILE_SUFFIXES (Zeile 286), wird also auch beim naechsten Lauf nicht mehr angefasst, und kein Aufraeumpfad des Skripts kennt ihn. Nachgestellt mit einem Minimalskript (os.replace -> Signal -> sys.exit): 'except Exception' wurde nicht erreicht, das finally loeschte die neue Datei, uebrig blieb ausschliesslich Foo.xlsx_ab12.bak. Derselbe Mechanismus greift in Block 8 (Zeile 2858 _av_safe_replace, danach 2861 _av_safe_move).

*Auswirkung:* Anwender drueckt Strg+C, waehrend gerade Q:\Ablage\Jahresbericht.xlsx zurueckgeschrieben wird: danach existiert nur noch Q:\Ablage\Jahresbericht.xlsx_9f3a1c.bak, die bearbeitete Fassung im Temp-Ordner ist geloescht, die Datei taucht in keiner Excel-Suche mehr auf und wird von jedem Folgelauf ignoriert.

*Vorschlag:* Im Signal-Handler nur ein Abbruch-Flag setzen und an einer definierten Stelle der Hauptschleife beenden (nicht mitten in der Dateiverarbeitung), oder in Block 5.5 und Block 8 zusaetzlich 'except BaseException' bzw. ein try/finally mit Restore-aus-.bak verwenden und die Stage-2-/Temp-Datei bis zum bestaetigten Move in _preserved_temp_files halten.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `4b_ersetze_font_in_excel.py` Zeile 2823  `[BEHOBEN]`

**Block 7 loescht das .xls/.xlt-Original ohne jede Integritaetspruefung der neu geschriebenen Zieldatei - als einziger Ersetzungspfad der Datei.**

```
safe_orig = prepare_long_path(original_path)
            try:
                os.chmod(safe_orig, stat.S_IWRITE)
                _av_safe_remove(safe_orig)
                detail_logger.debug(f"Alte Datei gelöscht: {original_path}")
```

*Begruendung:* Alle uebrigen Ersetzungspfade pruefen vor bzw. nach dem Schreiben mit verify_saved_file(): Block 6 (Long-Path-Konvertierung) in Zeile 2786, Block 5.5 (Stage-2-Move) in Zeile 2716, Block 8 (Long-Path-Restore) in Zeile 2863. Der haeufigste Produktionsfall - kurzer Pfad, .xls -> .xlsx - laeuft ueber Block 7 und hat KEINEN verify_saved_file-Aufruf (per Grep bestaetigt: verify_saved_file kommt nur in 2716, 2786, 2863 vor). Geschrieben wurde new_path durch workbook.SaveAs (Zeile 2515) und danach workbook.Save (Zeile 2654) - beides direkt auf die Netzwerkfreigabe. Genau gegen dieses Risiko (Netzwerk-Drop/AV/COM-Fehler waehrend Save) argumentiert der Kommentar in Zeile 2641-2647 und baut dafuer den Stage-2-Mechanismus - fuer den Konvertierungszweig greift dieser Schutz aber nicht, weil save_in_place_safe (Zeile 2648) bei was_converted=True auf True steht.

*Auswirkung:* Q:\Archiv\Inventar.xls, Save() liefert wegen eines Netzwerk-Aussetzers eine trunkierte/kaputte Inventar.xlsx zurueck, ohne eine COM-Exception zu werfen. Block 7 loescht anschliessend Inventar.xls. Ergebnis: unbrauchbare .xlsx, Original weg, Statuszeile meldet SUCCESS.

*Vorschlag:* In Block 7 vor os.chmod/_av_safe_remove pruefen: if not verify_saved_file(prepare_long_path(new_path)): raise RuntimeError(...) - dann bleibt das Original stehen und die Datei landet als ERROR im Protokoll (analog Zeile 2786-2788).


## `4c_ersetze_font_in_powerpoint.py` Zeile 975  `[BEHOBEN]`

**Rollback in _replace_file_with_backup greift bei Strg+C nicht: 'except Exception' faengt das vom Signal-Handler ausgeloeste SystemExit nicht - das Original bleibt als .bak zurueck, am Originalnamen liegt eine abgeschnittene Datei.**

```
try:
        _safe_move(src, dst)
    except Exception:
        if bak_long and os.path.exists(bak_long):
```

*Begruendung:* Der in Zeile 3019 registrierte _signal_handler endet mit sys.exit(1) (Zeile 407). Bei Strg+C/SIGTERM waehrend des Kopiervorgangs wird dadurch SystemExit im Hauptthread an der Unterbrechungsstelle - also mitten in shutil.copy2 innerhalb von _safe_move (Zeile 947) - ausgeloest. SystemExit und KeyboardInterrupt leiten sich von BaseException ab, nicht von Exception; das except in Zeile 975 faengt sie nicht. Die Rueckrollung (os.replace(bak_long, dst_long), Zeile 980) wird uebersprungen, obwohl das Original in Zeile 972 bereits per os.replace auf den .bak-Namen umbenannt wurde. Nachgestellt mit einer 1:1-Nachbildung der Funktion (Ausgabe: 'ROLLBACK-ZWEIG ERREICHT' erscheint nicht, original.pptx hat 4 Bytes, original.pptx_d27458.bak 1400 Bytes). Verschaerfend: kein Aufraeumpfad des Skripts erfasst .bak-Dateien im Quellbestand (weder _safe_cleanup_temp, das nur TEMP_PROCESS_PATH kennt, noch _cleanup_windows_temp); zusaetzlich loescht der finally-Block in Zeile 2563 die Stage-2-Tempdatei, und der Signal-Handler ruft in Zeile 400 _safe_cleanup_temp() auf, das den ganzen Temp-Ordner entfernt.

*Auswirkung:* Strg+C (oder Task-Abbruch/SIGTERM) genau waehrend des finalen Zurueckschreibens einer Praesentation, z. B. bei einer 80-MB-.pptx ueber die Netzfreigabe (Kopierdauer mehrere Sekunden): 'Vortrag.pptx' liegt danach als 0-bis-wenige-Byte-Datei auf der Freigabe, der einzige vollstaendige Datenbestand steht unter 'Vortrag.pptx_a1b2c3.bak' - ein Name, den weder der Anwender noch ein spaeterer Lauf des Skripts findet oder aufraeumt. Fuer den Nutzer sieht die Datei zerstoert aus.

*Vorschlag:* In _replace_file_with_backup 'except Exception:' zu 'except BaseException:' aendern (Rollback ausfuehren, dann re-raise). Analog fuer _safe_move-Aufrufer. Zusaetzlich beim Skriptstart nach '*.bak'-Resten im Zielbestand suchen bzw. den .bak-Pfad zusaetzlich in die Fehler-Logdatei (file_logger) statt nur ins Detail-Log schreiben.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `4c_ersetze_font_in_powerpoint.py` Zeile 3008  `[BEHOBEN]`

**Smoke-Test und seine Rueckfrage laufen auch mit --auto, obwohl die Option laut Hilfetext genau das ueberspringen soll - der geplante Lauf bleibt an der Eingabeaufforderung stehen oder bricht still ab.**

```
if not ask_yes_no("\nTrotzdem fortfahren? (Timeout-Risiko pro Datei!)",
                          default_yes=False):
```

*Begruendung:* Der Hilfetext von --auto lautet 'Trust-Center-Check und interaktive Rückfragen überspringen' (Zeile 2772). Alle anderen Rueckfragen sind entsprechend geklammert (Zeile 2828, 2938, 2945, 2959). Der komplette Smoke-Test-Block ab Zeile 2973 ist es nicht - weder der Aufruf test_trust_center_smoke() noch die Rueckfrage in Zeile 3008. Im Schwesterskript 4b_ersetze_font_in_excel.py ist derselbe Block mit 'if not auto_mode:' geklammert (Zeile 3534). Zusaetzlich startet der Smoke-Test in jedem Automatiklauf eine zusaetzliche PowerPoint-Instanz und schreibt eine Testdatei in TEMP_PROCESS_PATH.

*Auswirkung:* Geplanter Lauf 'python 4c_... --auto --dir Q:\' auf einem Rechner, dessen Trust-Center-Eintrag fuer den Temp-Ordner fehlt: Der Smoke-Test schlaegt fehl, das Skript fragt 'Trotzdem fortfahren?'. Mit angehaengter Konsole blockiert input() unbegrenzt (Task haengt); ohne stdin liefert ask_yes_no per EOFError den Vorgabewert False (Zeile 1581-1582) und das Skript beendet sich mit sys.exit(0), also Erfolgs-Exitcode - der Aufgabenplaner meldet 'erfolgreich', obwohl keine einzige Datei verarbeitet wurde.

*Vorschlag:* Den Block Zeile 2973-3016 wie in 4b mit 'if not auto_mode:' klammern; falls der Smoke-Test auch automatisch laufen soll, im Automatikmodus statt der Rueckfrage direkt mit einem von 0 verschiedenen Exitcode abbrechen bzw. protokolliert fortfahren.


## `4c_ersetze_font_in_powerpoint.py` Zeile 3097  `[BEHOBEN]`

**Die Abschluss-Wartezeile 'Beliebige Taste druecken' laeuft auch im Modus --auto, und _cleanup_windows_temp() steht dahinter statt davor.**

```
try:
        input("Beliebige Taste drücken, um das Fenster zu schließen ...")
    except EOFError:
        pass
    _cleanup_windows_temp()
```

*Begruendung:* Zwei Fehler in fuenf Zeilen. (1) Kein auto_mode-Schutz: In derselben Datei ist die identische Wartezeile im Probelauf-Zweig mit 'if not auto_mode:' abgesichert (Zeile 2959-2963), das Schwesterskript 4b_ersetze_font_in_excel.py hat an derselben Stelle 'if _is_tty() and not auto_mode:' (Zeile 3677). Haengt an einem geplanten Task eine Konsole (schtasks mit interaktivem Konto, oder Aufruf ueber cmd /k), liefert stdin kein EOF - input() blockiert unbegrenzt, der Task laeuft nie zu Ende und haelt die Einzelinstanz-Sperre (Zeile 3033) sowie ggf. eine PowerPoint-Instanz. Ist stdin ganz abgeloest, wirft input() unter Umstaenden RuntimeError('input(): lost sys.stdin') oder OSError statt EOFError - beides wird hier nicht gefangen und beendet das Skript mit Traceback. (2) _cleanup_windows_temp() steht NACH dem input(). In 4a_ersetze_font_in_word.py steht der gleiche Aufruf davor (Zeile 3495 vor 3497). Schliesst der Anwender das Fenster ueber das Kreuz, statt Enter zu druecken, unterbleibt die Temp-Bereinigung vollstaendig.

*Auswirkung:* Naechtlicher Lauf per Aufgabenplanung mit '--auto --dir Q:\': Verarbeitung ist fertig, der Prozess bleibt aber an der Eingabeaufforderung stehen. Der Folgelauf am naechsten Tag findet die Einzelinstanz-Sperre belegt und bricht mit sys.exit(1) ab (Zeile 3036) - die Migration steht still, ohne dass ein Fehler im Log erscheint. Beim manuellen Lauf, den der Anwender ueber das Fensterkreuz beendet, bleiben ~$-, gen_py- und 4c-Reste in %TEMP% liegen.

*Vorschlag:* _cleanup_windows_temp() VOR den input()-Block ziehen und den input()-Block mit 'if not auto_mode:' klammern (analog Zeile 2959 und 4b Zeile 3677); zusaetzlich neben EOFError auch OSError/RuntimeError abfangen.


## `5_OCR_PDF.py` Zeile 301  `[BEHOBEN]`

**detail_logger wird in except-Zweigen benutzt, die beim Import noch VOR seiner Zuweisung (Zeile 684) laufen - der Fehlerschlucker wirft dann selbst einen NameError und das Skript startet gar nicht.**

```
detail_logger.debug(f"_get_known_folder: Exception verworfen: {_e!r}")
```

*Begruendung:* detail_logger existiert erst ab Zeile 684 (bzw. 687 im Worker); vorher gibt es den Namen im Modul nicht (durch grep ueber alle Vorkommen bestaetigt: einzige Zuweisungen in 684 und 687). Mehrere except-Zweige, die diesen Namen benutzen, sind aber schon frueher erreichbar: Zeile 407-410 baut TEMP_DIR und ruft dabei _get_documents_folder -> _get_known_folder auf, dessen except in Zeile 300/301 steht; ebenso Zeile 316 (_get_user_shell_folder_from_registry) und Zeile 298 (CoTaskMemFree). Noch eindeutiger ist Zeile 646: _ensure_bom wird in Zeile 653/654 von _setup_logging aufgerufen - also von genau der Funktion, deren Rueckgabewert detail_logger erst erzeugt; dort KANN der Name konstruktionsbedingt noch nicht existieren. Ein Handler, der einen Fehler 'verwerfen' soll, loest damit einen zweiten, nicht abgefangenen Fehler aus. Nachgestellt: die Zeilen 252-302 und 370-382 des Originals unveraendert in eine Testdatei uebernommen und _get_known_folder zum Scheitern gebracht - Ergebnis 'NameError: name detail_logger is not defined' aus Zeile 301, ausgeloest waehrend der Behandlung der urspruenglichen Ausnahme.

*Auswirkung:* Auf einem Rechner, auf dem SHGetKnownFolderPath fuer FOLDERID_DOCUMENTS scheitert (fehlende/umgeleitete Known-Folder-Registrierung, gekapptes Roaming-Profil, Terminalserver-Sitzung ohne Dokumente-Ordner), bricht der Import in Zeile 408 mit einem NameError ab, statt wie vorgesehen auf %USERPROFILE%\Documents auszuweichen. Ebenso stirbt der Start in Zeile 653, wenn das Anlegen der Logdatei fehlschlaegt (volle Platte, AV-Sperre). In beiden Faellen laeuft das Skript nicht an und die Meldung zeigt auf den falschen Fehler.

*Vorschlag:* detail_logger/error_logger vor der ersten Verwendung binden - z.B. direkt nach den Importen 'detail_logger = logging.getLogger("DetailLogger")' und 'error_logger = logging.getLogger("ErrorLogger")' setzen und in Zeile 684 nur noch die Handler ergaenzen; alternativ in den fruehen Helfern (Zeile 221, 245, 298, 301, 316, 646) statt detail_logger ein 'pass' bzw. ein globals()-gesichertes Logging verwenden.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `5_OCR_PDF.py` Zeile 2036  `[BEHOBEN]`

**set_ocr_marker ersetzt die Originaldatei, stellt aber weder Zeitstempel noch NTFS-Sicherheitsinfo wieder her - bei --verify-markers gehen Aenderungsdatum und Owner/DACL der Ablagedatei endgueltig verloren.**

```
if safe_replace_with_retry(tmp, p):
```

*Begruendung:* set_ocr_marker (Zeile 2016-2050) schreibt eine komplette Neufassung der PDF nach '<original>.pdf.tmp_marker' (Zeile 2035) und schiebt sie in Zeile 2036 per safe_replace_with_retry ueber das Original. Die Datei ist danach eine NEUE Datei: neuer Zeitstempel, Owner = ausfuehrendes Konto, nur die vererbbaren ACEs des Ordners. Genau davor warnt der Kommentarblock ab Zeile 1188-1194 ('verliert der Nutzer dadurch den Zugriff'), und try_remove_empty_password direkt darunter macht es richtig: Zeile 2083/2084 sichern _read_file_times und _get_security_descriptor, Zeile 2091/2093 spielen sie zurueck. set_ocr_marker tut das nicht - weder selbst noch beim Aufrufer. Der Aufruf in Zeile 3278 (Marker-Korrektur nach veraPDF unter --verify-markers) liegt VOR Zeile 3409 (orig_times = _read_file_times) und VOR Zeile 3414 (orig_sd = _get_security_descriptor); die betroffenen Zweige kehren in Zeile 3295 bzw. 3302 zurueck und erreichen die Wiederherstellung in Zeile 3706/3708 nie. Und selbst im Zweig, der weiterlaeuft (Zeile 3290 'pass'), lesen 3409/3414 nur noch die bereits zerstoerten Werte. Nachgestellt mit pikepdf: Original auf 2021 datiert, Kern der Funktion (Pdf.open -> save(tmp) -> os.replace) ausgefuehrt, mtime danach 'heute'.

*Auswirkung:* Lauf mit --verify-markers ueber die Ablage: Eine Datei traegt den Marker PDF/A-2u, veraPDF bestaetigt aber nur 2b (Zeile 3274/3277). Zeile 3278 schreibt die Datei neu; der Zweig endet in Zeile 3295 mit SKIPPED. Ergebnis: Aenderungsdatum steht auf dem Laufdatum (Archivinformation weg), Eigentuemer ist das ausfuehrende Admin-Konto und die explizite Benutzer-ACE des Originals ist fort - auf einer Ablage mit Owner-Mapping (NAS/Unix-Security-Style) verliert der Fachnutzer damit den Zugriff auf seine Datei, ohne dass irgendetwas protokolliert wird.

*Vorschlag:* In set_ocr_marker vor Zeile 2021 orig_times = _read_file_times(file_path) und orig_sd = _get_security_descriptor(file_path) sichern und nach erfolgreichem safe_replace_with_retry (Zeile 2036) _apply_security_descriptor(file_path, orig_sd) sowie _write_file_times(file_path, orig_times) aufrufen - analog zu try_remove_empty_password, Zeile 2083-2093. Dasselbe in _add_ocr_marker_legacy fuer den Nicht-Inkrementell-Zweig (Zeile 1998-2001).

*Vom Pruefer ausgefuehrt nachgestellt.*

*Behoben (14.08.2026):* Die Sicherung erfolgt jetzt IN `set_ocr_marker` bzw.
`_add_ocr_marker_legacy` selbst (Helfer `_sichere_datei_metadaten` /
`_stelle_datei_metadaten_her`), nicht beim Aufrufer - damit ist Befund 3278
miterledigt, denn der dortige Zweig kehrt vor jeder Wiederherstellung zurueck.
Nachgestellt mit echter PDF und echtem os.replace: Original auf 2011-03-04
datiert, Marker geschrieben - alte Fassung mtime 2026-08-14, neue Fassung
mtime 2011-03-04, Marker gesetzt, Datei lesbar.

*Bei der Verifikation zusaetzlich gefunden - NICHT behoben, siehe unten:*
`_apply_security_descriptor` kann in einem Sonderfall den Zugriff ENTZIEHEN.
Nachgestellt: Datei in einem mit `tempfile.mkdtemp()` (Modus 0700) erzeugten
Ordner, deren Zugriff allein an einer geerbten ACE `EIGENTUEMERRECHTE`
(S-1-3-4) haengt. Nach `_get_security_descriptor` + `_apply_security_descriptor`
auf DIESELBE Datei steht dort `EIGENTUEMERRECHTE:(I)(IO)(F)` statt
`EIGENTUEMERRECHTE:(I)(F)` - das ergaenzte Inherit-Only-Flag nimmt die ACE fuer
die Datei selbst ausser Kraft, die Datei ist anschliessend nicht mehr lesbar
(PermissionError). Bei Ordnern mit benannten Benutzer-ACEs (Desktop,
C:	mp_skripte, normaler Temp-Ordner) laeuft derselbe Umlauf sauber -
nachgeprueft. Fuer Ablagen ist der Normalfall damit unkritisch, der Mechanismus
wird aber in 3c, 4b, 4c und 5_OCR_PDF an vielen Stellen benutzt und gehoert
gesondert geprueft.



## `5_OCR_PDF.py` Zeile 3278  `[BEHOBEN]`

**Marker-Korrektur ersetzt das Original, bevor Zeitstempel/ACL gesichert sind – und stellt sie danach nie wieder her**

```
if actual_code == "PDFA_2B":
                    set_ocr_marker(file_path, _compose_subject(
                        MARKER_VALUE_PDFA_2B, str(pdf_info.get("subject") or "")))
```

*Begruendung:* set_ocr_marker() schreibt eine neue Datei (<name>.pdf.tmp_marker) und schiebt sie per safe_replace_with_retry/os.replace ueber das Original (Zeilen 2020-2036). Damit traegt die Datei danach die aktuelle mtime und nur noch die vom Ordner geerbten ACEs. Die Sicherung der Originalwerte erfolgt aber erst spaeter: 'orig_times = _read_file_times(file_path)' (Zeile 3409) und 'orig_sd = _get_security_descriptor(file_path)' (Zeile 3414) – sie lesen also bereits die durch die Marker-Korrektur veraenderten Werte. Zusaetzlich endet dieser Zweig in der Regel direkt danach mit 'return build_result("SKIPPED", "SKIP_MARKER_PDFA_2B", ...)' (Zeile 3295); die Wiederherstellung in Zeile 3705-3708 laeuft nur 'if ocr_successful' und wird nie erreicht. Nachgestellt: Datei mit mtime 2011-03-04 angelegt, Subject per pikepdf umgeschrieben und per os.replace zurueckgeschoben -> mtime danach 2026-08-14. Der Kommentarblock ab Zeile 1188 begruendet ausdruecklich, warum Owner/DACL bei Ersetzungen uebernommen werden muessen (Zugriffsverlust bei Lauf unter Admin-Konto); dieser Pfad umgeht den Mechanismus.

*Auswirkung:* Beim Lauf mit --verify-markers (auch interaktiv aktivierbar, Zeile 5041-5048) verlieren Archivdateien, die nur nachverifiziert und anschliessend als SKIPPED gemeldet werden, ihr Aenderungsdatum und ihre expliziten NTFS-Rechte. Unter einem Admin-Konto kann das den Nutzerzugriff auf die Datei kosten. Im Upgrade-Fall (Ziel pdfa-2u) wird spaeter zwar restauriert, aber der bereits verfaelschte Zeitstempel.

*Vorschlag:* orig_times/orig_sd (Zeilen 3409/3414) vor den Block der Marker-Pruefung (Zeile 3256) ziehen und in jedem Zweig, der file_path anfasst – also auch vor den SKIPPED-Returns in Zeile 3285/3295 – _write_file_times()/_apply_security_descriptor() aufrufen.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `5_OCR_PDF.py` Zeile 3959  `[BEHOBEN]`

**Probelauf meldet 'uebersprungen' fuer Dateien, die der echte Lauf ersetzt**

```
if info.get("has_text") or info.get("has_ocr_marker"):
```

*Begruendung:* Der Probelauf ueberspringt jede Datei mit Textebene oder mit irgendeinem OCR-Marker. Der echte Lauf tut das nicht: (a) has_text fuehrt in Zeile 3317-3318 nur zu einem Hinweis ('ocrmypdf entscheidet mit skip_text=True'); ocrmypdf ueberspringt mit mode=skip lediglich die OCR der Textseiten (_pipeline.py: is_ocr_required -> ProcessingMode.skip), erzeugt aber sehr wohl eine Ausgabedatei, die dann per safe_replace_with_retry ueber das Original geschoben wird. (b) has_ocr_marker ist auch fuer die Marker 'PDF'/'LEGACY'/'PDFA_2B' wahr; bei der Standard-Zielvorgabe DEFAULT_OUTPUT_TYPE='pdfa-2u' (Zeile 435) werden genau diese Dateien im echten Lauf als Upgrade neu geschrieben (Zeile 3306-3315). Zusaetzlich meldet der Probelauf verschluesselte Dateien pauschal als uebersprungen, waehrend der echte Lauf Owner-Only-Verschluesselung entfernt und die Datei ersetzt (Zeile 3193).

*Auswirkung:* Genau die Funktion, die laut Docstring (Zeile 3898-3904) den Erstlauf auf fremdem Bestand verantwortbar machen soll, unterschaetzt die Zahl der in-place ersetzten Dateien massiv – bei Bestaenden mit vielen digital erzeugten PDFs (Word-Export etc.) um Groessenordnungen. Der Bediener gibt einen Lauf frei, der weit mehr anfasst als angekuendigt.

*Vorschlag:* Im Probelauf dieselbe Entscheidungslogik wie in process_pdf_file verwenden: nur bei existing_marker == Zielcode ueberspringen (bzw. PDFA_2U bei Ziel pdfa-2u), has_text nicht als Skip werten sondern als Hinweis ausgeben, und Owner-Only-Verschluesselung als 'wuerde entschluesselt + ersetzt' melden.


## `5_OCR_PDF.py` Zeile 5145  `[BEHOBEN]`

**--dry-run wird im Watch-Modus stillschweigend ignoriert – der Probelauf ersetzt PDFs tatsaechlich, und zwar ohne Einzelinstanz-Sperre**

```
if args.watch:
        run_watch_mode(target_dir, config, temp_dir, workers=num_workers,
                       initial_scan=not args.no_initial_scan)
```

*Begruendung:* Die Verzweigung auf args.watch steht VOR der dry-run-Abfrage. args.dry_run wird nur im else-Zweig (Zeile 5154) ausgewertet. Es gibt keine argparse-Pruefung, die --dry-run und --watch/--daemon als unvereinbar zurueckweist. Verschaerfend: Zeile 4861 ('if gem is not None and not args.dry_run') laesst den Einzelinstanz-Schutz genau deshalb aus, weil ein Probelauf 'nichts schreibt' (Kommentar Zeile 4856-4859) – diese Annahme ist in der Kombination falsch. Nachgestellt mit einem argparse-Nachbau: Namespace(watch=True, dry_run=True) -> Sperre nicht belegt, Aufruf geht nach run_watch_mode.

*Auswirkung:* Wer den Hotfolder gefahrlos testen will ('python 5_OCR_PDF.py Q:\ --watch --dry-run'), startet in Wahrheit den vollen Schreibbetrieb: jede eingelegte oder beim Initial-Scan gefundene PDF wird per safe_replace_with_retry an Ort und Stelle ersetzt. Zusaetzlich laeuft dieser schreibende Lauf ohne Einzelinstanz-Sperre und kann damit parallel zu einem echten Lauf ueber denselben Bestand arbeiten – genau das, was der Kommentar in Zeile 4857 ('Temp-Kopien und Ersetzungen gegenseitig wegziehen') verhindern soll.

*Vorschlag:* Entweder vor der Verzweigung hart abbrechen (z.B. 'if args.watch and args.dry_run: parser.error("--dry-run und --watch schliessen sich aus")') oder die Reihenfolge umdrehen: 'if args.dry_run: stats_result = dry_run_directory(...)' zuerst pruefen, danach erst args.watch. Unabhaengig davon sollte die Einzelinstanz-Sperre auch im Probelauf belegt werden, sobald irgendein schreibender Pfad erreichbar bleibt.

*Vom Pruefer ausgefuehrt nachgestellt.*

*Nachpruefung 14.08.2026:* Dieser Befund war bereits behoben, nur nicht als
solcher markiert. Der Zweig `if args.watch and args.dry_run:` existiert samt
Begruendung und fuehrt statt des Hotfolder-Betriebs eine einmalige Pruefung
aus. `--daemon` ist ein Alias fuer `--watch` (dieselbe `args.watch`) und damit
mitabgedeckt. Keine Aenderung noetig.



## `8_verschieben_auf_Google_Drive.ps1` Zeile 2392  `[BEHOBEN]`

**Beim selektiven Loeschen wird das ReadOnly-Attribut nicht entfernt - schreibgeschuetzte Quelldateien werden nie geloescht und kosten je 15 Sekunden Leerlauf.**

```
[System.IO.File]::Delete($srcLong)
```

*Begruendung:* [System.IO.File]::Delete wirft auf einer Datei mit gesetztem ReadOnly-Attribut eine UnauthorizedAccessException - das ist dokumentiertes .NET-Verhalten und wurde nachgestellt. Invoke-WithRetry (Zeile 376-379) faengt genau diesen Typ ab und wiederholt ihn fuenfmal mit 1+2+4+8 Sekunden Pause, bevor er weitergeworfen wird; ein Attribut-Problem verschwindet durch Warten aber nicht. Test-FileIsLocked (Zeile 2381) schlaegt nicht an, weil es nur mit FileAccess::Read oeffnet - schreibgeschuetzte Dateien gelten dort korrekt als 'nicht gesperrt'. Dass das Skript das Problem grundsaetzlich kennt, zeigt Import-Timestamps: dort wird in Zeile 646-653 vor dem Setzen der Zeitstempel eigens das ReadOnly-Bit geloescht und in Zeile 668-670 wieder gesetzt. Im Loeschpfad fehlt dieser Schritt. Der Vollloesch-Zweig ueber robocopy /MIR ist davon nicht betroffen - nachgestellt: robocopy entfernt schreibgeschuetzte Zieldateien anstandslos. Beide Wege verhalten sich also unterschiedlich, obwohl beide 'Quelle loeschen' heissen.

*Auswirkung:* Quelle enthaelt schreibgeschuetzte Dateien (Archivgut, von CD/DVD uebernommene Bestaende - im Archivbestand der Normalfall), Verifikation [1] CRC oder [2] Light. Nachgestellt mit der Original-Kopie von Invoke-WithRetry: das Loeschen einer schreibgeschuetzten Datei dauert 15,1 s, endet mit 'Zugriff auf den Pfad ... wurde verweigert' und die Datei existiert danach noch. Jede betroffene Datei landet als 'Dateifehler' in $DeleteErrors, bleibt in der Quelle liegen und der Lauf meldet am Ende 'Beim Loeschen sind Fehler aufgetreten' - obwohl die Datei im Ziel nachweislich bit-identisch angekommen ist. Bei 5.000 schreibgeschuetzten Dateien sind das rund 21 Stunden reine Wartezeit, in denen nichts geloescht wird.

*Vorschlag:* Vor dem Delete das ReadOnly-Bit abraeumen, analog zu Zeile 646-653: Attribute per [System.IO.File]::GetAttributes($srcLong) lesen, bei gesetztem ReadOnly per SetAttributes($srcLong, $attr -band (-bnot [System.IO.FileAttributes]::ReadOnly)) entfernen und erst dann Delete aufrufen. Zusaetzlich sollte Invoke-WithRetry UnauthorizedAccessException hier nicht wiederholen - ein Rechte-/Attributfehler wird durch Warten nicht besser und blockiert nur den Lauf.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `9_fehlerhafte_Dateien_finden.ps1` Zeile 1096  `[BEHOBEN]`

**Der Smoke-Test ruft Quit() ungeprueft auf - bei laufender Word-/PowerPoint-Sitzung des Anwenders wird dessen Sitzung samt ungespeicherter Dokumente beendet**

```
try { $word.Quit() } catch {}
```

*Begruendung:* Die Datei dokumentiert die Gefahr in Zeile 907-917 selbst: 'New-Object -ComObject Word.Application startet KEINE neue Instanz, wenn Word bereits laeuft - COM haengt sich an die vorhandene' und '$word.Quit() beendet damit die Sitzung des Anwenders samt ungespeicherter Dokumente'. Der Aufraeumteil am Skriptende zieht daraus die richtige Konsequenz und beendet nur selbst gestartete Instanzen (Zeile 2013: 'if ($wordPid -gt 0) { $word.Quit() }', Zeile 2021 fuer PowerPoint). Der Smoke-Test-Job tut das nicht: er ermittelt in Zeile 1079-1083 zwar $myWordPid und weiss damit genau, ob eine eigene Instanz entstanden ist, benutzt diese Information aber nur fuer den Timeout-Kill (Zeile 1117-1123) - im finally wird Quit() bedingungslos aufgerufen. Dasselbe Muster in Zeile 1275 fuer PowerPoint. Verschaerfend: unmittelbar davor wurde $word.DisplayAlerts = 0 gesetzt (Zeile 1072), also verwirft Word ohne Rueckfrage. Der Ablauf ist der ungluecklichste denkbare: die Warnung in Abschnitt 8 sagt dem Anwender in Zeile 925 ausdruecklich zu 'Ihre Sitzung wird am Ende nicht beendet' - direkt danach laeuft in Abschnitt 8b der Smoke-Test, der genau das tut.

*Auswirkung:* Anwender hat Word mit ungespeicherten Dokumenten offen, startet das Skript, liest die Warnung in Zeile 922-935 und bestaetigt 'Trotzdem fortfahren' im Vertrauen auf die Zusage aus Zeile 925. Sekunden spaeter beendet der Word-Smoke-Test in Zeile 1096 seine Sitzung; wegen DisplayAlerts=0 ohne Speichern-Rueckfrage. Die ungespeicherte Arbeit ist weg, noch bevor die erste Datei geprueft wurde. Fuer PowerPoint identisch (Zeile 1275, DisplayAlerts=1 = ppAlertsNone).

*Vorschlag:* Im Job dieselbe Bedingung setzen wie im Aufraeumteil: 'if ($myWordPid) { try { $word.Quit() } catch {} }' bzw. 'if ($myPptPid) { try { $ppt.Quit() } catch {} }' - $myWordPid/$myPptPid sind an dieser Stelle bereits ermittelt. Bei uebernommener Fremdsitzung nur die COM-Referenz freigeben und $word.Visible wieder auf $true setzen. Sauberer waere zusaetzlich, den Smoke-Test bei erkannten laufenden Office-Sitzungen ganz zu ueberspringen; belastbar ist er dort ohnehin nicht.


## `9_fehlerhafte_Dateien_finden.ps1` Zeile 1511  `[BEHOBEN]`

**Die Bremse gegen die Endlos-Neustartschleife greift fuer PowerPoint nie: Schluessel 'ppt' wird gegen den Wert 'PowerPoint' geprueft**

```
if ($appType -and $script:AppFailed.ContainsKey($appType) -and $script:AppFailed[$appType]) {
```

*Begruendung:* $script:AppFailed wird in Zeile 1489 mit den Schluesseln 'word', 'excel', 'ppt' angelegt und in Zeile 1710 ueber $targetApp gesetzt ($targetApp ist 'word'/'excel'/'ppt', Zeile 1556-1558). Die Abfrage hier benutzt aber $appType aus Get-OfficeAppType (Zeile 276-283), das 'Word', 'Excel' oder 'PowerPoint' zurueckgibt. Fuer Word und Excel geht das gut, weil ein @{}-Hashtable in PowerShell case-insensitiv ist. 'PowerPoint' hat aber mit 'ppt' keine Aehnlichkeit. Nachgestellt in Windows PowerShell 5.1: ContainsKey('Word')=True, ContainsKey('Excel')=True, ContainsKey('PowerPoint')=False. Auch $script:AppFailed['PowerPoint'] waere $null, also falsy - die Bremse ist doppelt wirkungslos.

*Auswirkung:* Kaputte oder nicht mehr startbare PowerPoint-Installation (Lizenzproblem, voller RAM, zerstoerte Registrierung): Der Restart in Zeile 1702 wirft, Zeile 1710 setzt AppFailed['ppt']=$true und protokolliert 'Weitere ppt-Dateien werden bis Skriptende uebersprungen'. Genau das passiert dann NICHT. Fuer jede weitere .pptx/.ppt-Datei laeuft der volle Weg erneut: Zugriff auf $pptn (bereits per ReleaseComObject in Zeile 1701 freigegeben) wirft, Test-OfficeAppAlive schlaegt fehl, needsRestart, Stop-Process auf die veraltete PID, dann wieder New-Object -ComObject PowerPoint.Application. Bei 3.000 Praesentationen sind das 3.000 COM-Aktivierungsversuche, die je nach Defekt jeweils in einen langen DCOM-Timeout laufen und halbgestartete POWERPNT.EXE hinterlassen koennen - exakt die 'Endlos-Crash-Schleife', die der Kommentar in Zeile 1485-1488 verhindern soll. Der Lauf bleibt praktisch haengen.

*Vorschlag:* Gegen denselben Schluessel pruefen, der auch gesetzt wird. Entweder die Zuordnung $targetApp (Zeile 1556-1558) vor diese Abfrage ziehen und $script:AppFailed[$targetApp] pruefen, oder $script:AppFailed in Zeile 1489 mit den Get-OfficeAppType-Werten anlegen (@{ 'Word'=$false; 'Excel'=$false; 'PowerPoint'=$false }) und Zeile 1710 entsprechend auf $appType umstellen. Wichtig ist nur, dass Setzen und Pruefen dieselbe Schluesselmenge benutzen.

*Vom Pruefer ausgefuehrt nachgestellt.*




## NACHTRAG (nicht aus dem Hauptlauf): `_apply_security_descriptor` — alle sieben Skripte  `[BEHOBEN]`

*Dieser Befund stammt NICHT aus der Mehr-Agenten-Durchsicht, sondern fiel
beim Verifizieren von `5_OCR_PDF.py/2036` auf. Er zaehlt deshalb nicht in
die Zahlen der Kopftabelle. Schwere: entspricht HOCH (moeglicher
Zugriffsverlust auf Ablagedateien).*

**Das Setzen des Eigentuemers ordnet die Vererbung neu und kann der Datei
genau den Zugriff nehmen, den die Funktion erhalten soll.**

```
win32security.SetNamedSecurityInfo(
    p, win32security.SE_FILE_OBJECT, sec_flags, owner, group, None, None)
```

*Gefunden bei der Verifikation von Befund 5_OCR_PDF.py/2036, nicht im
Hauptlauf gemeldet.*

*Begruendung:* Die Funktion setzte in ZWEI getrennten Aufrufen erst die DACL,
dann Owner und Gruppe. Schrittweise eingegrenzt (jeder Teil einzeln auf eine
frische Datei angewandt): der DACL-Aufruf allein ist folgenlos, der
GROUP-Aufruf allein ebenfalls - der OWNER-Aufruf richtet den Schaden an. Nach
ihm verlieren die geerbten ACEs ihr `(I)`-Kennzeichen, und eine geerbte
`EIGENTUEMERRECHTE`-ACE (OWNER RIGHTS, S-1-3-4) bekommt zusaetzlich `(IO)` =
INHERIT_ONLY. Damit gilt sie fuer die Datei selbst nicht mehr. Wer seinen
Zugriff allein daraus bezieht, kann die Datei anschliessend nicht mehr oeffnen
(PermissionError) - das Gegenteil dessen, was die Funktion bezweckt.
Betroffen sind Objekte, deren Zugriff an einer geerbten OWNER-RIGHTS-ACE
haengt; bei Ordnern mit benannten Benutzer-ACEs (Desktop, C:	mp_skripte,
normaler Temp-Ordner) laeuft derselbe Umlauf sauber - gegengeprueft.
Die Funktion steckt in allen sieben schreibenden Skripten (3a, 3b, 3c, 4a, 4b,
4c, 5_OCR_PDF) in logisch identischer Form; die Fassungen unterscheiden sich
nur im Protokoll-Idiom.

*Auswirkung:* Nach jeder Ersetzung, die diese Funktion aufruft - also nach
Konvertierung, Schriftartersetzung, OCR und Marker-Korrektur - kann die Datei
fuer den Fachnutzer unzugaenglich werden, obwohl der Mechanismus eigens
eingebaut wurde, um genau das zu verhindern.

*Behoben:* Owner, Gruppe und DACL werden jetzt in EINEM
SetNamedSecurityInfo-Aufruf gesetzt; Windows berechnet die Rechte damit in
einem Zug und der schaedliche Zwischenzustand entsteht nicht. Zusaetzlich wird
der Eigentuemer nur gesetzt, wenn er tatsaechlich abweicht. Schlaegt der
gemeinsame Aufruf fehl (typisch: kein SeRestorePrivilege im Nutzer-Kontext),
wird die DACL wie bisher einzeln nachgezogen. Nachgeprueft ueber alle sieben
Skripte in beiden Ordnerlagen: Datei bleibt lesbar, kein `(IO)`-Flag - und die
Gegenprobe zeigt, dass eine explizite ACE weiterhin auf die Ersatzdatei
uebertragen wird.


# MITTEL

## `0_Vorab-Check.py` Zeile 2666  `[BEHOBEN]`

**Warnungen gehen weder in die Abschlussmeldung noch in den Exit-Code ein: ein Lauf mit fehlendem veraPDF meldet 'Alles bereits korrekt eingerichtet' und endet mit Exit-Code 0.**

```
nothing = not any(STATE[k] for k in (
        "created_dirs", "installed_packages", "installed_tools",
        "copied_files", "env_set", "path_added", "errors", "manual_actions",
        "admin_pending",
    ))
    if nothing:
        log_info("\n ✅ Alles bereits korrekt eingerichtet – keine Änderungen nötig.")
```

*Begruendung:* 'warnings' fehlt in der Liste, und main_stage1() (Zeile 2776-2780) wertet ebenfalls nur errors (->2) und manual_actions (->1) aus. Es gibt aber Pfade, die ein fehlendes Pflichtwerkzeug AUSSCHLIESSLICH als Warnung ablegen: _check_verapdf() Zeile 2454 log_warn('❌ veraPDF nicht im PATH/Standard-Pfaden gefunden.'), _install_verapdf_silent() Zeile 2382 log_warn('⚠️ verapdf-install.bat nicht in Deploy-Quelle gefunden.') mit return None, danach Zeile 2462 log_download_hint('verapdf') - ebenfalls nur log_warn. Weder errors noch manual_actions werden gesetzt. Der Kopfkommentar definiert 0 aber als 'alles OK'.

*Auswirkung:* Zweiter Lauf auf einem eingerichteten Rechner (Verzeichnisse existieren, Pakete installiert, ENV gesetzt), Deploy-Quelle erreichbar, enthaelt aber kein verapdf-install.bat: STATE hat nur warnings. Ausgabe: ' ✅ Alles bereits korrekt eingerichtet – keine Änderungen nötig.' und Rueckgabe 0. Ein Verteilwerkzeug/Batch, das auf ERRORLEVEL 0 prueft, meldet den Rechner als fertig - obwohl veraPDF fehlt und die spaetere PDF/A-Validierung im OCR-Skript nicht laufen kann. Die einzige Spur ist die Zeile ' Warnungen: 3'.

*Vorschlag:* 'warnings' in die nothing-Bedingung aufnehmen und in main_stage1()/main_stage2() einen eigenen Rueckgabewert fuer 'nur Warnungen' vergeben (oder die betroffenen Pfade in _check_verapdf/_install_verapdf_silent auf STATE['manual_actions'] umstellen, so wie try_install_tool() es in Zeile 1772 fuer die anderen Tools tut).


## `10_dateinamen_bereinigen.py` Zeile 864  `[BEHOBEN]`

**Phase B wandelt jedes '++' zwischen zwei ASCII-Buchstaben in 'ü' um, sobald der Name irgendwo einen Umlaut enthaelt - entgegen der Zusicherung im Docstring.**

```
r"(?<=[A-Za-z])\+\+(?=[A-Za-z])",
```

*Begruendung:* Der Gate in Zeile 836 ist _MOJIBAKE_HINT_RE, dessen Zeichenklasse laut Kommentar (Zeile 582-583) bewusst auch legitime deutsche Umlaute enthaelt. Damit ist der Gate bei praktisch jedem deutschen Dateinamen mit ae/oe/ue/ss offen, und die Regel greift auf jedem 'C++Wort'. Der Docstring in Zeile 819-820 versichert, dass legitime C++-Namen unangetastet bleiben; das Beispiel dort ('C++_Tutorial.pdf') traegt aber zufaellig einen Unterstrich hinter dem '++' und faellt deshalb als einziges nicht unter die Regel.

*Auswirkung:* Nachgestellt: 'C++Kurs für Anfänger.pdf' -> 'CüKurs für Anfänger.pdf'; 'C++Builder Handbuch Größe.doc' -> 'CüBuilder Handbuch Größe.doc'. Die Datei wird auf der Platte umbenannt, die Zeichenfolge 'C++' ist im Namen unwiederbringlich verloren.

*Vorschlag:* Phase B nur anwenden, wenn im Namen ein echter Mojibake-Beleg vorliegt, der ueber legitime Umlaute hinausgeht - also gegen _MOJIBAKE_RESULT_HINT_RE bzw. _CLUSTER_SUSP_CHARCLASS pruefen statt gegen das grosszuegige _MOJIBAKE_HINT_RE. Zusaetzlich 'C++', 'G++', 'A++' o.ae. als Ausnahmen fuehren.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `10_dateinamen_bereinigen.py` Zeile 1325  `[BEHOBEN]`

**Der Fortschrittsvermerk wird auch fuer Eintraege geschrieben, deren Umbenennung fehlgeschlagen ist (ERROR) oder die wegen Kollision uebersprungen wurden (SKIPPED) - nach einem Abbruch werden sie nie wieder angefasst.**

```
append_resume(resume_path, entry_path)
```

*Begruendung:* append_resume() wird ohne jede Statusabfrage direkt nach sanitize_entry_on_disk() aufgerufen. Ein Eintrag, dessen Umbenennung an WinError 32 (AV-Scanner/Indexer-Lock) gescheitert ist - genau der Fall, fuer den RENAME_RETRIES/RENAME_BASE_DELAY ueberhaupt existieren -, landet damit in der Resume-Datei. Bricht der Lauf danach ab, wird delete_resume_file() nicht erreicht (Zeile 1483 laeuft nur nach vollstaendigem Durchlauf), und der Wiederaufnahmelauf ueberspringt genau die Eintraege, die noch offen sind. Der Kommentar in Zeile 1240-1242 begruendet den Vermerk mit 'Eintraege, die schon geprueft wurden' - ein an einem Lock gescheiterter Eintrag ist aber gerade nicht erledigt.

*Auswirkung:* Nachgestellt: Datei 'Ma·e gesperrt.txt' mit offenem Handle gesperrt, process_directory() mit resume_path aufgerufen. Ergebnis STATS ERROR=1, RENAMED=0 - und die Resume-Datei enthaelt trotzdem die Zeile '\\?\...\Ma·e gesperrt.txt'. Ein Wiederaufnahmelauf ueberspringt die Datei, meldet 0 Fehler und der kaputte Name bleibt unbemerkt im Bestand.

*Vorschlag:* append_resume nur bei RenameStatus.RENAMED, CLEAN, EXCLUDED und WOULD_RENAME aufrufen; ERROR und SKIPPED bewusst nicht vermerken, damit ein Wiederaufnahmelauf sie erneut versucht.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `1_temp_dateien_entfernen.py` Zeile 736  `[BEHOBEN]`

**Die Kennung der Fortschrittsdatei enthaelt nur den Zielpfad, nicht die gewaehlten Loeschoptionen - ein Wiederaufnahmelauf mit schaerferen Optionen ueberspringt die bereits vermerkten Verzeichnisse.**

```
kennung = hashlib.sha1(
            os.path.normcase(aufgeloest).encode("utf-8", "surrogatepass")
        ).hexdigest()[:12]
```

*Begruendung:* Der Regelsatz wird in main() ueber die Abfragen 'delete_bak_log' (Zeile 1329) und 'delete_desktop_ini' (Zeile 1335) zusammengesetzt und variiert damit von Lauf zu Lauf. Der Vermerk unterscheidet diese Faelle nicht: Zeile 1407 bildet die Kennung ausschliesslich ueber target_dir. Ein abgebrochener Lauf mit engem Regelsatz blockiert deshalb einen spaeteren Lauf mit weiterem Regelsatz, ohne dass der Benutzer davon erfaehrt - die Meldung in Zeile 1178 nennt nur die Anzahl uebersprungener Verzeichnisse.

*Auswirkung:* Lauf 1 auf Q:\ ohne *.bak/*.log wird nach 2 Stunden abgebrochen; 30.000 Verzeichnisse sind vermerkt. Lauf 2 auf Q:\ mit aktiviertem *.bak/*.log ueberspringt genau diese 30.000 Verzeichnisse - dort wird keine einzige .bak/.log-Datei geprueft, obwohl der Benutzer die Option ausdruecklich eingeschaltet hat. Am Ende wird der Vermerk geloescht, der Lauf gilt als vollstaendig.

*Vorschlag:* Die Optionen in die Kennung aufnehmen (z. B. sha1 ueber Pfad + 'bak_log=1;desktop_ini=0;min_age=24') oder den Regelsatz als Kopfzeile in die Fortschrittsdatei schreiben und beim Laden vergleichen; bei Abweichung von vorne beginnen.


## `1_temp_dateien_entfernen.py` Zeile 1193  `[BEHOBEN]`

**Nach einem Abbruch melden die noch laufenden Worker ihre Datei als 'erledigt', obwohl process_file wegen abort_flag sofort zurueckkehrt und nichts geprueft hat.**

```
finally:
            # Erst hier gilt die Datei als verarbeitet - siehe Fortschritt.
            fortschritt.erledigt(verzeichnis)
```

*Begruendung:* process_file beginnt mit 'if abort_flag.is_set(): return' (Zeile 1044/1045) - die Datei wird also NICHT geprueft. Der finally-Block des Wrappers ruft trotzdem fortschritt.erledigt() auf und zaehlt sie damit als verarbeitet. Faellt der Zaehler dadurch auf 0, wird das Verzeichnis in die Fortschrittsdatei geschrieben und beim naechsten Lauf uebersprungen. Unter Python < 3.9 (Zeile 1225: shutdown(wait=False) ohne cancel_futures) trifft das nicht nur die bis zu 8 laufenden, sondern alle bis zu SUBMIT_BUFFER = 400 bereits eingereihten Dateien, die dann samt und sonders ungeprueft als 'erledigt' durchlaufen.

*Auswirkung:* Strg+C waehrend der Verarbeitung von Q:\Projekte: die 8 gerade laufenden Dateien kehren ungeprueft zurueck, werden aber als erledigt gezaehlt. Ist darunter die letzte offene Datei ihres Verzeichnisses, gilt das Verzeichnis als abgeschlossen und wird beim naechsten Start uebersprungen - seine Temp-Dateien bleiben dauerhaft liegen. Unter Python 3.8 betrifft das bis zu 400 Dateien pro Abbruch.

*Vorschlag:* Im Wrapper unterscheiden, ob wirklich verarbeitet wurde: 'if not abort_flag.is_set(): fortschritt.erledigt(verzeichnis)' - oder process_file einen Rueckgabewert liefern lassen (verarbeitet ja/nein) und nur dann fertigmelden. Der Semaphor muss weiterhin in jedem Fall freigegeben werden.


## `2b_entferne_schutz_excel.ps1` Zeile 1025  `[BEHOBEN]`

**Unter PowerShell 7 gibt es FileInfo.GetAccessControl nicht mehr: die ACL-/Owner-Uebernahme bei der Dateiersetzung faellt still aus und wird nur als DEBUG vermerkt.**

```
$acl = $fi.GetAccessControl(
```

*Begruendung:* Die Methoden FileInfo.GetAccessControl/SetAccessControl existieren nur im .NET Framework. In .NET Core / .NET 5+ - und damit in PowerShell 7 - wurden sie entfernt; der Zugriff laeuft dort ueber die Erweiterungsmethoden aus System.Security.AccessControl.FileSystemAclExtensions. Nachgestellt auf demselben Rechner: unter Windows PowerShell 5.1.26100.9168 'GetAccessControl: OK', unter PowerShell 7.6.4 'GetAccessControl FEHLER: Method invocation failed because [System.IO.FileInfo] does not contain a method named GetAccessControl'. Der Fehler laeuft in den catch in Zeile 1036, der ausschliesslich Write-DetailedLog mit Level DEBUG schreibt und $null zurueckgibt; Set-FileSecuritySnapshot steigt daraufhin in Zeile 1044 sofort wieder aus. Es gibt keine Warnung im Haupt-Log, keinen CSV-Eintrag und keine Zeile in der Abschlussstatistik. Der Kommentarblock in Zeile 952-969 begruendet diese Uebernahme ausdruecklich als Schutzmassnahme: 'Laeuft das Skript unter einem Admin-Konto ueber Ablagen mit expliziten Benutzer-ACEs oder NAS-Owner-Mapping, verliert der Nutzer dadurch den Zugriff.'

*Auswirkung:* Wird das Skript mit pwsh statt powershell.exe gestartet - auf diesem System ist PowerShell 7.6.4 installiert und die Standard-Shell -, so ersetzt es weiterhin Dateien (Zeile 2328 Copy work->final, Zeile 2340 Delete des Originals), uebernimmt aber weder DACL noch Owner. Jede konvertierte oder entsperrte Datei erbt danach nur die vererbbaren ACEs des Zielordners und gehoert dem ausfuehrenden Konto. Auf einer Archiv-Ablage mit expliziten Benutzer- oder Gruppen-ACEs verlieren die eigentlichen Nutzer den Zugriff auf genau die Dateien, die das Skript angefasst hat - stillschweigend, denn im Haupt-Log steht fuer jede Datei nur 'Erledigt'. Derselbe Ausfall trifft auch den Rollback-Pfad in Zeile 2387.

*Vorschlag:* Entweder die Ausfuehrung auf Windows PowerShell 5.1 erzwingen (#requires -Version 5.1 reicht nicht, da 7 die Bedingung ebenfalls erfuellt - stattdessen $PSVersionTable.PSEdition -ne 'Desktop' pruefen und mit klarer Meldung abbrechen), oder editionsneutral ueber die statischen Erweiterungsmethoden gehen: [System.Security.AccessControl.FileSystemAclExtensions]::GetAccessControl($fi, $sections) bzw. ::SetAccessControl($fi, $fs). In jedem Fall den Fehlschlag von Get-FileSecuritySnapshot mindestens als WARN ins Haupt-Log schreiben statt nur als DEBUG, damit ein flaechiger ACL-Verlust nicht unbemerkt bleibt.

*Behoben (14.08.2026):* Empirisch bestaetigt - `FileInfo.GetAccessControl` ist
unter Windows PowerShell 5.1.26100.9168 vorhanden, unter PowerShell 7.6.4
(Edition Core) nicht. Ein editionsneutraler Ersatz existiert nicht ohne
Weiteres: `System.Security.AccessControl.FileSystemAclExtensions` ist in
BEIDEN Editionen nicht geladen (nachgeprueft), und `Get-Acl`/`Set-Acl` haetten
das Langpfad-Problem, dessentwegen hier ueberhaupt .NET benutzt wird.
Deshalb der erste im Vorschlag genannte Weg: Beim Start wird
`$PSVersionTable.PSEdition -eq 'Core'` geprueft und mit klarer Meldung und
Exit-Code 2 abgebrochen - ein stiller Rechteverlust auf der Ablage ist
schlimmer als ein deutlicher Abbruch. Nachgestellt: `pwsh` bricht mit
Exit-Code 2 ab, `powershell.exe` (Desktop) laeuft unveraendert weiter.
Dieselbe Pruefung in 2a, 2c und 7 ergaenzt - dort steht dasselbe Konstrukt,
ohne dass es gemeldet war. Zusaetzlich meldet `Get-FileSecuritySnapshot` einen
Lesefehler jetzt als WARN statt als DEBUG: als DEBUG-Zeile ging der
Rechteverlust im Detail-Log unter.


*Vom Pruefer ausgefuehrt nachgestellt.*


## `2c_entferne_schutz_powerpoint.ps1` Zeile 474  `[BEHOBEN]`

**Im trap-Handler ist das Einsammeln der .pid-Dateien genau durch die Liste bedingt, die es fuellen soll - bei leerer Liste laeuft es nie.**

```
if ($script:TrackedPptPids -and (Test-Path -LiteralPath $script:TempPath)) {
```

*Begruendung:* $script:TrackedPptPids ist eine System.Collections.Generic.List[int] (Z. 193). PowerShell wertet eine leere Auflistung im booleschen Kontext als $false aus - nachgestellt unter powershell.exe 5.1: 'empty list truthy? False'. Der ganze Zweck des Blocks Z. 475-485 ist aber, PIDs aus den zurueckgebliebenen .pid-Dateien nachzutragen, also genau den Fall abzudecken, in dem noch nichts getrackt ist. Ist die Liste leer, wird nicht gelesen, anschliessend ueberspringt auch Z. 486 Clear-TrackedPowerPointInstances, und Z. 487 loescht $script:TempPath samt der .pid-Dateien - die Information ist danach unwiederbringlich weg. Die Liste ist regelmaessig leer, weil Clear-TrackedPowerPointInstances in Z. 867 .Clear() aufruft (u. a. nach jedem Konvertierungs-Timeout, Z. 1916).

*Auswirkung:* Ablauf: Datei A laeuft in den Job-Timeout -> Z. 1916 Clear-TrackedPowerPointInstances -> Liste leer. Bei Datei B startet Convert-PptToPptx eine POWERPNT-Instanz und schreibt deren PID nach $TempPath\PptDeepClean_*.pid. Tritt jetzt ein unerwarteter Fehler auf (Netzlaufwerk weg, Logger-Fehler), greift der trap: die .pid-Datei wird NICHT gelesen, die unsichtbare POWERPNT.EXE bleibt als Zombie mit gesperrter Temp-Datei zurueck, und der Temp-Ordner mit dem einzigen Beleg wird geloescht. Bei wiederholten Laeufen sammeln sich unsichtbare PowerPoint-Prozesse an.

*Vorschlag:* Die Liste aus der Bedingung nehmen: if (Test-Path -LiteralPath $script:TempPath) { ... }. Analog Z. 486 auf 'if ($script:TrackedPptPids.Count -gt 0)' umstellen (oder bedingungslos aufrufen), damit die .Count-Semantik statt der Auflistungs-Wahrheit greift.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `2c_entferne_schutz_powerpoint.ps1` Zeile 961  `[BEHOBEN]`

**Remove-OrphanedBackups stuft ein abgeschnittenes Original als intakt ein und loescht das Backup, das die einzige vollstaendige Kopie ist.**

```
try { $origOk = ([System.IO.FileInfo]::new($orig)).Length -gt 0 } catch { $origOk = $false }
```

*Begruendung:* Der Docstring (Z. 936-937) begruendet die Schonung damit, dass 'das Backup die einzige intakte Kopie sein' kann - geprueft wird aber nur 'existiert und > 0 Byte'. Der Zustand, gegen den das Backup ueberhaupt schuetzt, ist ein harter Abbruch (Prozess-Kill, Stromausfall, Netz-Drop) zwischen Z. 2032 ([System.IO.File]::Delete($finalDest), das Original ist da bereits geloescht) und dem Abschluss von Z. 2034. File.Copy laesst dabei typischerweise eine teilweise geschriebene Zieldatei zurueck - die ist groesser als 0 Byte und besteht die Pruefung. Der interne Rollback (Z. 2072-2094) greift nur bei einer abgefangenen Ausnahme, nicht bei einem harten Abbruch. Die Groesse des .bak wird nirgends mit der des Originals verglichen.

*Auswirkung:* Praesentation.pptx (40 MB) wird in-place bearbeitet; der Rechner faellt waehrend Z. 2034 aus. Zurueck bleiben Praesentation.pptx mit 6 MB (unbrauchbar) und Praesentation.pptx.bak mit den vollstaendigen 40 MB. Laeuft das Skript erst spaeter als $BackupCleanupMinAgeHours = 24 h wieder (Z. 162), liefert Z. 961 $origOk = $true, Z. 970 sieht das Backup als alt genug an und Z. 974 loescht es. Die vollstaendige Datei ist endgueltig weg, ohne Warnung - im Gegenteil, sie wird als 'Verwaistes Backup entfernt' protokolliert.

*Vorschlag:* Zusaetzlich die Groessen vergleichen und im Zweifel schonen, z. B.: $origOk = ($origLen -gt 0 -and $origLen -ge $fi.Length). Ist das Original kleiner als sein Backup, den Fund wie einen Waisen behandeln ($res.Orphans++, OrphanList, Write-Log WARN) statt zu loeschen.


## `3a_doc_docx_auf_neueste_Version_aktualisieren.py` Zeile 731  `[BEHOBEN]`

**Staging-Kopien (.tmp_new) und Sicherungen (.bak) bleiben nach hartem Abbruch im Zielverzeichnis liegen und werden von keinem Aufraeumpfad je erfasst**

```
stage_s = long_path(dst + ".tmp_new")
```

*Begruendung:* Der Temp-Ordner liegt auf C: (%USERPROFILE%\Documents, Zeile 156-160), die Ziele auf Q:/R:/G:/UNC. os.replace() ueber Volumegrenzen wirft immer OSError, d.h. der Staging-Pfad (Zeile 731-743) ist bei diesem Skript nicht der Ausnahme-, sondern der REGELFALL: fuer JEDE Datei entsteht kurzzeitig '<ziel>.docx.tmp_new' im Archiv-Ordner, und bei jedem gleichnamigen Update zusaetzlich '<ziel>.docx.bak' (Zeile 2925). Der Fehlerpfad in Zeile 744-751 faengt nur 'Exception'. Der Signal-Handler (Zeile 310-337) loest per sys.exit() ein SystemExit aus - das ist KEINE Exception-Unterklasse und wird dort nicht gefangen; Strg+C mitten in shutil.copy2 laesst die Staging-Datei also stehen. Aufgeraeumt wird ausschliesslich TEMP_PROCESS_PATH (Zeile 324-329, 3746-3751), cleanup_orphaned_temp_dirs() sieht nur TEMP_PROCESS_PARENT und cleanup_windows_temp() nur %TEMP%. file_generator() liefert '.tmp_new'/'.bak' wegen des Extension-Filters nie, ein Folgelauf bemerkt sie also auch nicht. Das Schwesterskript 5_OCR_PDF.py hat genau dafuer bereits einen Rest-Suchlauf ('Suche nach Resten (*.backup und *.tmp_new ...)', dort Zeile 4927, Kommentar 1636: 'Stromausfall, harter Abbruch -, bleibt die .tmp_new liegen'); in 3a fehlt dieser Mechanismus vollstaendig.

*Auswirkung:* Strg+C oder Netzwerkabriss waehrend der Kopie auf Q: hinterlaesst 'Q:\...\Bericht.docx.tmp_new' (vollstaendige Zweitkopie, Groesse des Dokuments) und/oder 'Bericht.docx.bak' dauerhaft in der Nutzerablage. Nach mehreren abgebrochenen Laeufen wachsen dort unsichtbare Dubletten in Gigabyte-Groesse, die kein Skript wieder entfernt; zusaetzlich verwirren sie die Anwender und die spaetere Google-Drive-Migration (Schritt 8).

*Vorschlag:* Analog zu 5_OCR_PDF.py einen Rest-Suchlauf ergaenzen: beim Start (nach cleanup_orphaned_temp_dirs) und/oder am Ende das Zielverzeichnis nach '*.tmp_new' und verwaisten '*.bak'/'*.bak_*' durchsuchen, die aelter als der Laufbeginn sind, und nach Rueckfrage bzw. im --auto-start-Modus automatisch entfernen. Zusaetzlich in robust_move() 'except BaseException' verwenden, damit auch SystemExit/KeyboardInterrupt die Staging-Datei vor dem Weiterreichen loeschen.


## `3a_doc_docx_auf_neueste_Version_aktualisieren.py` Zeile 2696  `[BEHOBEN]`

**Alle Info-Zaehler und die Modus-Verteilung werden bei jedem Wiederholungsversuch erneut hochgezaehlt, obwohl die Datei nur einmal verarbeitet wird**

```
mode_distribution[original_mode] += 1
```

*Begruendung:* process_file_with_retries() ruft end_compatibility_mode() bis zu MAX_RETRIES (=3) mal mit DENSELBEN Dicts info_counters/mode_distribution auf (Zeile 3104-3111). Saemtliche Info-Zaehler werden aber INNERHALB von end_compatibility_mode() gesetzt, und zwar VOR der haeufigsten Fehlerquelle (SaveAs2, Zeile 2871): mode_distribution (2696), BINARY_CONVERTED (2774), MODE_UPGRADED (2778), APPVERSION_UPDATED (2787), WORD_AUTO_REPAIRED (2793), MACROS_FOUND (2824), MACRO_SIGNATURE_INVALIDATED (2837). Schlaegt SaveAs2/Verschieben fehl, wird die Ausnahme in Zeile 3043 weitergereicht, der Wrapper zaehlt einen neuen Versuch und die Zaehler laufen erneut durch. Nur 'stats' (UPDATED/ERROR) wird korrekt einmal pro Datei gefuehrt, weil es im Aufrufer liegt.

*Auswirkung:* Eine .doc-Datei, die beim SaveAs zweimal scheitert und im dritten Anlauf gelingt, erzeugt BINARY_CONVERTED=3, MACROS_FOUND=3 und mode_distribution[11]=3 bei stats['UPDATED']=1. Auf einer Ablage mit vielen sperrigen Dateien meldet der Abschlussbericht (Konsole, JSON-Summary, E-Mail) mehr Binaerkonvertierungen und Makrodateien als ueberhaupt Dateien vorhanden sind; die 'Urspruenglichen Compat-Modi' sind systematisch zugunsten der problematischen Dateien verzerrt. Genau die Dateien, die Aufmerksamkeit brauchen, verfaelschen die Bilanz am staerksten.

*Vorschlag:* Die Info-Zaehler nicht in end_compatibility_mode() mutieren, sondern lokale Flags/Werte sammeln und im Rueckgabewert (z.B. als drittes Tupel-Element) an process_file_with_retries() geben, das sie nur beim endgueltig uebernommenen Versuch (break-Pfad) auf info_counters/mode_distribution addiert. Alternativ pro Datei ein frisches Zaehler-Dict uebergeben und nur nach dem erfolgreichen bzw. letzten Versuch in die Gesamtzaehler mergen.


## `3a_doc_docx_auf_neueste_Version_aktualisieren.py` Zeile 2903  `[BEHOBEN]`

**Der Zielname wird nur dann atomar reserviert, wenn er bereits belegt ist - im haeufigeren Fall 'noch frei' bleibt das vom Autor selbst beschriebene Ueberschreib-Zeitfenster offen**

```
if new_ext != ext and safe_exists(target_path):
```

*Begruendung:* reserve_unique_path() wurde laut eigenem Docstring (Zeile 433-447) genau deshalb eingefuehrt: "'Pruefen und danach benutzen' laesst ein Zeitfenster offen, in dem ein parallel laufender Durchgang - oder ein Nutzer, der gerade speichert - denselben Namen belegen kann; die Datei wuerde beim anschliessenden Schreiben ueberschrieben." Aufgerufen wird die Reservierung hier aber nur, WENN target_path schon existiert. Ist 'Bericht.docx' beim Pruefen noch nicht da, wird kein Platzhalter angelegt; zwischen dieser Pruefung (2903) und dem tatsaechlichen Schreiben in robust_move() (2955) liegen Backup-Kopie, Temp-Aufraeumen und die vollstaendige Netzkopie - je nach Dateigroesse und VPN mehrere Sekunden. robust_move() verwendet os.replace() bzw. os.replace(stage, dst), was ein in dieser Zeit neu entstandenes Ziel KOMMENTARLOS ueberschreibt, und ohne Backup, weil backup_path nur bei safe_exists(target_path) gesetzt wird (2924). Dasselbe Muster in sanitize_file_on_disk() (Zeile 1804 pruefen, 1825 os.replace).

*Auswirkung:* Das Skript konvertiert 'Bericht.doc' auf Q: (Ziel 'Bericht.docx' existiert noch nicht). Waehrend der Sekunden bis zum Verschieben speichert ein Anwender aus Word heraus 'Bericht.docx' in denselben Ordner. Das Skript ueberschreibt diese frische, fremde Datei ersatzlos - kein .bak, kein Ausweichname, kein Logeintrag. Die Arbeit des Anwenders ist verloren.

*Vorschlag:* reserve_unique_path()-Logik generalisieren: sobald new_ext != ext, den Zielnamen IMMER atomar belegen - zuerst per os.open(target, O_CREAT|O_EXCL) den Wunschnamen selbst reservieren, und nur wenn das mit FileExistsError scheitert, auf den Ausweichnamen ausweichen. target_was_reserved entsprechend setzen, damit die Freigabe im finally weiter greift.


## `3a_doc_docx_auf_neueste_Version_aktualisieren.py` Zeile 3262  `[BEHOBEN]`

**Voruebergehende Oeffnen-Fehler landen als 'SKIPPED' in der Resume-Liste und werden dadurch in allen Folgelaeufen dauerhaft uebersprungen**

```
"UPDATED", "ALREADY_CURRENT", "SKIPPED")
```

*Begruendung:* Der Kommentar direkt darueber (3257-3260) haelt ausdruecklich fest, dass nur DAUERHAFT erledigte Dateien in die Resume-Datei gehoeren, und nimmt SKIPPED_LOCKED korrekt aus. Der Sammelrueckgabewert 'SKIPPED' aus Zeile 2689 wird aber fuer JEDEN gescheiterten Open vergeben - unabhaengig von der Ursache. In diesen Zweig faellt nach drei Versuchen alles, was keine TimeoutError und kein HR_OPEN_ESCALATE-Fehler war: DFS-/VPN-Aussetzer, RPC-Fehler, kurzfristig nicht erreichbare Freigabe, Datei waehrend des Zugriffs verschoben. Die Ausgabe lautet dann 'ÜBERSPRUNGEN: Zugriff verweigert' und der Zaehler PROTECTED_SKIPPED wird erhoeht (2688), obwohl der Grund haeufig transient ist. Anschliessend wird der Pfad in already_done aufgenommen (3264) und in die Resume-Datei geschrieben (3269).

*Auswirkung:* Faellt die Q:-Freigabe waehrend eines Laufs fuer wenige Sekunden aus, werden die in diesem Fenster bearbeiteten Dateien als 'SKIPPED (Passwortgeschuetzt/kein Zugriff)' verbucht und in die Resume-Datei geschrieben. Jeder Folgelauf mit --resume bzw. mit der automatischen Resume-Datei ueberspringt sie ohne erneuten Versuch (Zeile 3231-3235, gezaehlt als RESUMED). Diese Dateien bleiben unbemerkt fuer immer im alten Kompatibilitaetsmodus bzw. als .doc liegen, ohne in der Fehlerstatistik aufzutauchen.

*Vorschlag:* Den Open-Fehlerpfad differenzieren: bei last_exception mit is_transient_error(...) oder wenn ueberhaupt kein Passwortproblem vorlag, einen eigenen Status (z.B. 'SKIPPED_TRANSIENT') zurueckgeben und diesen - wie SKIPPED_LOCKED - aus resume_eligible ausnehmen. Nur echte Passwort-/Rechtefehler duerfen dauerhaft als erledigt gelten.


## `3a_doc_docx_auf_neueste_Version_aktualisieren.py` Zeile 3878  `[BEHOBEN]`

**Abschliessende Enter-Abfrage laeuft auch mit --auto-start und blockiert einen geplanten Task unbegrenzt**

```
input("Drücken Sie Enter, um das Fenster zu schließen ...")
```

*Begruendung:* Alle anderen Rueckfragen des Skripts sind mit 'if not auto_mode' abgesichert (warn_running_word 944, Trust-Center 3434, Startbestaetigung 3680); nur diese letzte nicht. Die Schwesterskripte 4b (Zeile 3677 '_is_tty() and not auto_mode'), 4c (2959), 10 (1552) und 5_OCR_PDF (5237) kapseln dieselbe Zeile bereits mit einer auto/tty-Pruefung - 3a ist hier zurueckgeblieben. Der 'except EOFError' faengt nur den Fall eines geschlossenen Eingabekanals ab. Nachgestellt: mit stdin=DEVNULL kommt tatsaechlich EOFError (unkritisch), mit einer offenen stdin-Pipe - also sobald das Skript aus einem Wrapper-Skript, aus einem Dienst oder aus einer CI/Orchestrierung heraus gestartet wird - blockiert der Prozess dauerhaft (Test lief nach 8 s noch, ohne Fortschritt).

*Auswirkung:* '3a_... --dir Q:\ --auto-start', aus einem PowerShell-Wrapper mit umgeleiteter Standardeingabe oder aus SCCM/CI gestartet, arbeitet die gesamte Ablage ab und bleibt danach fuer immer an der Enter-Abfrage stehen. Der geplante Task laeuft nie 'fertig', der Einzelinstanz-Lock (LOCK_FILE) wird nie freigegeben - _release_lock() wurde zwar im finally bereits ausgefuehrt, der naechste Lauf startet also, kollidiert aber mit dem noch laufenden Prozess und dessen Word-Instanz. Bei Nachtlaeufen bleiben Word-Prozesse und Prozesskarteileichen bis zum manuellen Eingriff bestehen.

*Vorschlag:* Zeile 3877-3880 in 'if not auto_mode and sys.stdin is not None and sys.stdin.isatty():' kapseln (wie in 4b_ersetze_font_in_excel.py Zeile 3677) und zusaetzlich KeyboardInterrupt mit abfangen.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `3b_xls_xlsx_auf_neueste_Version_aktualisieren.py` Zeile 197  `[BEHOBEN]`

**detail_logger wird in _resolve_documents_dir benutzt, bevor es existiert - im Fehlerfall bricht das Skript beim Import mit NameError ab.**

```
except Exception as _e:
        detail_logger.debug(f"_resolve_documents_dir: Exception verworfen: {_e!r}")
```

*Begruendung:* _resolve_documents_dir() wird auf Modulebene in Zeile 200 aufgerufen (_docs_base = _resolve_documents_dir()). detail_logger entsteht aber erst in Zeile 297 (file_logger, detail_logger = _setup_logging()). Wirft win32api.GetLongPathName(base) - etwa bei einem per Ordnerumleitung auf eine Serverfreigabe gelegten Dokumente-Ordner mit eingeschraenkten Rechten -, laeuft der except-Zweig in einen NameError, der nicht mehr gefangen wird. Nachgestellt: das Muster liefert reproduzierbar 'NameError: name detail_logger is not defined'. Der Rueckfall 'return base' in Zeile 198, der genau diesen Fall abfangen soll, wird dadurch nie erreicht.

*Auswirkung:* Auf einem Rechner mit umgeleitetem Dokumente-Ordner startet das Skript ueberhaupt nicht: Abbruch beim Import mit 'NameError: name detail_logger is not defined' statt mit dem vorgesehenen stillen Rueckfall auf den Kurzpfad. Ohne Log, da das Logging zu diesem Zeitpunkt noch nicht existiert.

*Vorschlag:* Im except-Zweig von _resolve_documents_dir nicht loggen (nur 'pass' bzw. return base), oder _setup_logging() vor der Berechnung von _docs_base ausfuehren.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `3b_xls_xlsx_auf_neueste_Version_aktualisieren.py` Zeile 1204  `[BEHOBEN]`

**Staging-Datei '<Ziel>.tmp_new' bleibt nach Strg+C auf der Freigabe liegen und wird von keinem Aufraeumpfad erfasst.**

```
stage      = dst + ".tmp_new"
```

*Begruendung:* Im Cross-Volume-Fall (TEMP auf C:, Ziel auf Q:/UNC - der Regelfall dieses Skripts) kopiert robust_move zuerst nach '<Ziel>.tmp_new' AUF DAS ZIELVOLUME. Aufgeraeumt wird die Staging-Datei nur im 'except Exception'-Zweig (Zeile 1224-1230). Der Signal-Handler beendet den Prozess mit sys.exit(130) (Zeile 696); die dabei ausgeloeste SystemExit ist von BaseException abgeleitet und wird von 'except Exception' NICHT gefangen - der Zweig laeuft also nie. _safe_cleanup_temp() raeumt ausschliesslich TEMP_PROCESS_PATH unter Documents auf, die Staging-Datei liegt aber im Zielverzeichnis. Auch ein spaeterer Lauf entfernt sie nicht: file_generator filtert ueber die Endungsliste (Zeile 2172-2176, 2227), '.tmp_new' steht dort nicht, und der ~$/._-Aufraeumzweig greift nur bei diesen Praefixen.

*Auswirkung:* Strg+C (oder Netzabbruch/Prozess-Kill) waehrend shutil.copy2 einer grossen Mappe hinterlaesst dauerhaft z.B. 'Q:\Sammlung\Inventar.xlsx.tmp_new' - eine halbe Datei, die aussieht wie ein Dokument, von keinem Lauf mehr entfernt wird und Anwender zum Umbenennen/Oeffnen verleitet.

*Vorschlag:* Eindeutiges Staging-Praefix mit PID/UUID verwenden und beim Start bzw. im Signal-Handler die Reste des eigenen Laufs im Zielbaum entfernen; im Signal-Handler zusaetzlich BaseException-sicher aufraeumen (try/finally in robust_move statt 'except Exception').


## `3b_xls_xlsx_auf_neueste_Version_aktualisieren.py` Zeile 2803  `[BEHOBEN]`

**Der Excel-Neustart innerhalb von convert_excel_file bindet nur die lokale Variable neu; der Aufrufer arbeitet danach mit einem toten COM-Objekt.**

```
excel_app = _create_excel_app()
```

*Begruendung:* excel_app ist Parameter von convert_excel_file. Die Zuweisung in Zeile 2803 (und identisch in Zeile 2888) wirkt nur lokal. process_directory haelt weiterhin current_excel (Zeile 3717/3765) auf die per _quit_excel_app beendete Instanz. Zwei Folgen: (1) process_file_with_retries (Zeile 3357-3391) wiederholt convert_excel_file nach einer Ausnahme bis zu MAX_RETRIES-mal mit genau diesem toten Objekt - jeder COM-Aufruf wirft, wird in der pw-Schleife (Zeile 2754) still als last_exception geschluckt, und die Datei endet in Zeile 2943 mit 'SKIPPED / Uebersprungen (Passwort/kein Zugriff)', obwohl weder ein Passwort noch fehlende Rechte die Ursache waren. (2) Beim naechsten Schleifendurchlauf schlaegt der Lebenszeichen-Test (Zeile 3736) fehl und _quit_excel_app(current_excel) laeuft - dort ist pid_to_wait aber das globale excel_app_pid, also die PID der NEUEN Instanz (Zeile 3438). Quit() auf dem toten Objekt scheitert still, danach werden 5 s auf die falsche PID gewartet und die frisch erzeugte, gesunde Instanz per proc.kill() abgeschossen.

*Auswirkung:* Datei mit definedName-Konflikt loest Pre-Clean + internen Excel-Neustart aus und wirft danach beim SaveAs. Retry laeuft gegen die tote Instanz, die Datei wird als SKIPPED mit dem sachlich falschen Grund 'Passwort/kein Zugriff' protokolliert; anschliessend rund 7 s Leerlauf plus ein zusaetzlicher unnoetiger Excel-Neustart pro betroffener Datei.

*Vorschlag:* convert_excel_file die aktuelle Instanz zurueckgeben lassen (oder ueber excel_app_global arbeiten) und in process_file_with_retries/process_directory current_excel nach jedem Aufruf aus excel_app_global aktualisieren; in _quit_excel_app nicht das globale excel_app_pid, sondern die zum uebergebenen Objekt gehoerende PID verwenden.


## `3c_ppt_pptx_auf_neueste_Version_aktualisieren.py` Zeile 224  `[BEHOBEN]`

**detail_logger wird in _resolve_documents_dir benutzt, bevor er existiert - der Rueckfall im except-Zweig wirft dann NameError und das Skript startet nicht.**

```
detail_logger.debug(f"_resolve_documents_dir: Exception verworfen: {_e!r}")
```

*Begruendung:* _resolve_documents_dir wird beim Import in Zeile 235 ausgefuehrt ('_docs_base = _resolve_documents_dir()'), detail_logger entsteht erst in Zeile 314 ('file_logger, detail_logger = _setup_logging()'). Schlaegt win32api.GetLongPathName(base) fehl (z.B. umgeleiteter Dokumente-Ordner auf einer gerade nicht erreichbaren Freigabe), soll der except-Zweig den Fehler verwerfen und 'base' zurueckgeben - stattdessen wirft die Logzeile selbst NameError. Der Fehler entsteht IM Handler und wird von keinem weiteren except gefangen. Nachgestellt: das Muster bricht mit "NameError: name 'detail_logger' is not defined" ab, statt den Rueckfallwert zu liefern.

*Auswirkung:* Auf einem Rechner mit auf eine Netzfreigabe umgeleitetem Dokumente-Ordner, die im Startmoment nicht antwortet, endet der Aufruf des Skripts vor jeder Ausgabe mit einem NameError-Traceback statt mit dem vorgesehenen Rueckfall auf den kurzen Pfad. Kein Log, keine Fehlermeldung, der geplante Task laeuft gar nicht an.

*Vorschlag:* Im except-Zweig von _resolve_documents_dir (wie in _resolve_script_directory und _select_log_directory) keinen Logger benutzen - schlicht 'pass' bzw. Sammeln der Meldung in einer Variablen, die nach dem Logging-Setup nachgetragen wird.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `3c_ppt_pptx_auf_neueste_Version_aktualisieren.py` Zeile 2325  `[BEHOBEN]`

**Abgestuerztes PowerPoint wird bei skip_password als 'SKIPPED' verbucht; der Wiederanlauf der Engine haengt aber allein an 'ERROR' und laeuft deshalb nicht an.**

```
if not pres_opened or pres is None:
            if skip_password:
                pbar.write("  → ÜBERSPRUNGEN: Zugriff verweigert.")
                detail_logger.warning(
                    f"Übersprungen (kein Zugriff): {display_path(original_path)}"
                )
                return "SKIPPED"
```

*Begruendung:* Stuerzt die COM-Instanz zwischen zwei Dateien ab (der Code rechnet ausdruecklich damit, siehe Meldung Zeile 2640 'PowerPoint beschäftigt oder abgestürzt' und das Neustart-Intervall PPT_RESTART_INTERVAL), liefert Presentations.Open in allen MAX_RETRIES Versuchen einen com_error, der weder Passwortfehler noch Timeout ist. pres_opened bleibt False und die Funktion liefert 'SKIPPED'. Die Gesundheitspruefung samt Neustart der Engine in process_directory haengt jedoch an 'if result == "ERROR":' (Zeile 2725) und wird damit nie ausgeloest. skip_password ist im Automatikmodus per Vorgabe True (Zeile 2941-2942), also genau im unbeaufsichtigten Betrieb.

*Auswirkung:* Geplanter Task mit --auto-start: PowerPoint stuerzt bei Datei 130 ab. Die Dateien 131 bis 149 werden ohne einen einzigen COM-Kontakt als 'ÜBERSPRUNGEN: Zugriff verweigert' protokolliert und nicht konvertiert; erst der turnusmaessige Neustart bei Datei 150 (processed_count % 25) heilt den Zustand. Der Laufbericht weist die 19 Dateien als 'Übersprungen' statt als Fehler aus, ein Alarm bleibt aus.

*Vorschlag:* Vor dem SKIPPED-Rueckgabewert in Zeile 2331 die COM-Instanz pruefen (ppt_app.Version) und bei nicht erreichbarer Instanz 'ERROR' zurueckgeben, oder in process_directory die Neustartpruefung fuer 'SKIPPED' ebenso durchfuehren wie fuer 'ERROR'.


## `3c_ppt_pptx_auf_neueste_Version_aktualisieren.py` Zeile 2674  `[BEHOBEN]`

**Der Probelauf (--dry-run) loescht Dateien auf der Ablage, obwohl er zusichert, das Dateisystem nicht anzutasten.**

```
gen = file_generator(directory, clean_temp=True,
                         exclude_patterns=exclude_patterns)
```

*Begruendung:* clean_temp wird fest auf True gesetzt, dry_run wird an file_generator nicht durchgereicht. In file_generator loescht Zeile 2074 (safe_remove(full_t)) jede nicht gesperrte Datei, die mit '~$' oder '._' beginnt - im gesamten durchlaufenen Baum. Dass der Probelauf das Dateisystem unangetastet lassen soll, ist im Code ausdruecklich festgehalten: Zeile 2705-2707 ('Im Probelauf NICHT umbenennen - das Dateisystem bleibt vollstaendig unangetastet') und die Abschlussmeldung Zeile 3178 ('PROBELAUF ABGESCHLOSSEN (es wurde NICHTS geändert)'). Die Umbenennung ist mit 'if not dry_run' abgesichert, die Loeschung nicht - die Zusicherung gilt also nur zur Haelfte.

*Auswirkung:* 'python 3c_... --dir Q:\ --dry-run' zum Abschaetzen des Aufwands loescht auf Q:\ alle verwaisten '~$...'-Owner-Dateien und alle '._...'-Dateien (macOS-Ressourcezweige von Mac-Arbeitsplaetzen, die bei alten Dateien Nutzdaten enthalten koennen) - und meldet anschliessend, es sei nichts geaendert worden.

*Vorschlag:* clean_temp=not dry_run uebergeben (bzw. dry_run an file_generator durchreichen und die Loeschschleife 2065-2076 im Probelauf ueberspringen).


## `4a_ersetze_font_in_word.py` Zeile 1325  `[BEHOBEN]`

**Bei hartem Abbruch bleibt die Staging-Kopie '<name>.docx.tmp_new' in der Archivablage liegen; 'except Exception' faengt das SystemExit des Signal-Handlers nicht, und kein Aufraeumpfad des Skripts erfasst diese Reste je wieder.**

```
if last_exc is not None:
            raise last_exc
    except Exception:
```

*Begruendung:* _replace_file_atomic legt in Zeile 1292 'stage = dst + ".tmp_new"' NEBEN dem Original auf dem Zielvolume an (also auf Q:/R:/G:). Die Bereinigung dieser Stage-Datei haengt allein am 'except Exception' in Zeile 1325. _signal_handler (Zeile 1033-1060) beendet mit sys.exit(1); das erzeugt SystemExit, eine BaseException - sie laeuft an 'except Exception' vorbei. Danach greift keine weitere Bereinigung: _safe_cleanup_temp_process() raeumt nur TEMP_PROCESS_PATH, _cleanup_windows_temp() nur %TEMP%, und file_generator() filtert in Zeile 1768 nach Endungen ('.docx' usw.), sodass 'Vertrag.docx.tmp_new' auch von Folgelaeufen nie gesehen wird. 5_OCR_PDF.py kennt genau dieses Problem und raeumt es in Zeile 1645 ('if low.endswith(".pdf.tmp_new")') explizit auf - hier fehlt das Gegenstueck.

*Auswirkung:* Lauf ueber Q:\, 40-MB-Datei, Strg+C waehrend _robust_copy in die Stage-Datei: nachgestellt mit einem Testskript im Temp-Verzeichnis - nach dem SystemExit liegen 'Original.docx' UND 'Original.docx.tmp_new' im Zielordner, das Original ist korrekt unveraendert. Bei einem Abbruch nach 3000 Dateien sammeln sich diese Halbdateien in der Ablage, wandern in die Migration nach Google Drive und sind fuer den Anwender nicht von Nutzdaten zu unterscheiden.

*Vorschlag:* In Zeile 1325 'except BaseException:' verwenden (oder die Stage-Bereinigung in ein finally verlagern) und zusaetzlich einen Restesuchlauf fuer '*.tmp_new' analog 5_OCR_PDF.py Zeile 1645 vorsehen.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `4b_ersetze_font_in_excel.py` Zeile 1858  `[BEHOBEN]`

**Der Probelauf meldet gesperrte Dateien als 'wuerden uebersprungen', der Echtlauf prueft die Sperre an keiner Stelle.**

```
if is_locked_by_other(file_path):
            stats["GESPERRT"] += 1
            tag = "GESPERRT     "
```

*Begruendung:* is_locked_by_other wird ausschliesslich in dry_run_directory aufgerufen (Grep ueber die Datei: nur Definition 1828 und Aufruf 1858). Weder replace_fonts_in_workbook noch process_directory pruefen die Office-Owner-Datei ~$<name>. Der Probelauf verspricht in Zeile 1877 woertlich 'Gesperrt (wuerden uebersprungen)'; der Echtlauf oeffnet diese Mappen jedoch und faehrt den vollen Ersetzungs- und Ruecksschreibpfad. Damit weicht die Vorschau systematisch vom tatsaechlichen Verhalten ab - genau die Zahl, auf die sich der Anwender vor dem Echtlauf verlaesst.

*Auswirkung:* Kollege hat Q:\Ablage\Kasse.xlsx geoeffnet (~$Kasse.xlsx liegt daneben). Der Probelauf zaehlt die Datei unter GESPERRT und sagt 'wird uebersprungen'. Der Echtlauf oeffnet sie (Excel faellt wegen DisplayAlerts=False stillschweigend auf schreibgeschuetzt zurueck), schreibt Stage 2 und versucht Original -> .bak -> Move; im besten Fall endet das als ERROR mit .bak-Rest, im schlechteren wird an einer Datei gearbeitet, die parallel offen ist.

*Vorschlag:* is_locked_by_other am Anfang von replace_fonts_in_workbook aufrufen und mit einem eigenen Status (SKIPPED) zurueckkehren - dann stimmen Probelauf und Echtlauf ueberein.


## `4b_ersetze_font_in_excel.py` Zeile 2124  `[BEHOBEN]`

**Schriftart in bedingter Formatierung wird nie gesetzt - FormatCondition.Font.Name ist in Excel nicht beschreibbar, der Fehler wird still verschluckt.**

```
fc.Font.Name = font_name
```

*Begruendung:* Das Font-Objekt einer FormatCondition unterstuetzt nur Schriftschnitt, Unterstreichung, Farbe und Durchstreichung - dieselbe Einschraenkung, die im Excel-Dialog 'Zellen formatieren' fuer bedingte Formatierung sichtbar ist (Feld 'Schriftart' und 'Schriftgrad' sind ausgegraut); das zugrunde liegende dxf-Format kennt keinen Schriftnamen. Die Zuweisung laeuft daher in Laufzeitfehler 1004, der in Zeile 2125-2126 nur nach DEBUG geloggt wird. Die Konfigurationsausgabe in Zeile 3493 nennt 'Bedingte Formatierung' dennoch ausdruecklich als Verarbeitungsumfang. Sekundaer: sheet.UsedRange.FormatConditions liefert bei Blaettern mit uneinheitlichen Regeln ohnehin nicht die Gesamtheit der Regeln.

*Auswirkung:* Eine Mappe, in der Kennzahlen per bedingter Formatierung in Calibri hervorgehoben werden, wird als SUCCESS gemeldet; die bedingte Formatierung behaelt Calibri, obwohl das Protokoll und die Startausgabe vollstaendige Ersetzung zusichern.

*Vorschlag:* Entweder den Verarbeitungsumfang in Zeile 3493 und die Dokumentation um den Hinweis ergaenzen, dass bedingte Formatierung technisch keinen Schriftnamen traegt, oder _set_conditional_formatting_fonts entfernen; alternativ das Ergebnis pruefen und die Datei als PARTIAL melden statt still zu scheitern.


## `4b_ersetze_font_in_excel.py` Zeile 3231  `[BEHOBEN]`

**Wiederholungslauf nach abgebrochener Konvertierung erzeugt eine zweite Zieldatei (_1) und laesst die ungeprueftete erste Fassung dauerhaft im Bestand.**

```
result = replace_fonts_in_workbook(
                            file_path, excel, pbar, font_name,
                            metadata_selected)
```

*Begruendung:* FILE_PROCESSING_ATTEMPTS=2 wiederholt jede Datei, die ERROR liefert. Bei .xls/.xlt ist die Konvertierung aber nicht rein: Stage 1 (Zeile 2513-2517) hat new_path bereits auf die Freigabe geschrieben, bevor irgendeine Schriftersetzung stattfand. Faellt der Lauf danach aus (Save-Fehler Zeile 2654, Fehler in Block 7), bleibt diese Datei liegen; nichts loescht oder prueft sie. Der zweite Versuch findet new_path vor und weicht ueber die Kollisionsschleife (Zeile 2476-2482) auf '<name>_1.xlsx' aus und loescht bei Erfolg das .xls-Original (Block 7).

*Auswirkung:* Inventar.xls, erster Versuch bricht nach SaveAs beim Save ab: es entstehen Inventar.xlsx (alte Schriftarten oder trunkiert, nie mit verify_saved_file geprueft) und beim zweiten Versuch Inventar_1.xlsx; Inventar.xls wird geloescht. Die Konvertierungs-CSV verzeichnet nur Inventar.xls -> Inventar_1.xlsx, die verwaiste Inventar.xlsx ist in keinem Protokoll als fragwuerdig markiert.

*Vorschlag:* Bei Abbruch nach Stage 1 den bereits geschriebenen new_path im Fehlerpfad wieder entfernen (oder als RESCUE_/.bak kennzeichnen), oder .xls/.xlt-Dateien, deren Stage 1 bereits gelaufen ist, von der Wiederholung ausnehmen (Rueckgabe eines eigenen, nicht wiederholbaren Fehlerstatus).


## `4c_ersetze_font_in_powerpoint.py` Zeile 2255  `[BEHOBEN]`

**Nach fehlgeschlagener Re-Serialisierung wird temp_stage2_path nicht zurueckgesetzt (im Parallelzweig Zeile 2288 geschieht das) - der Schreibschutz des Originals wird entfernt und nie wiederhergestellt, Zeitstempel bleiben veraendert, und die Datei wird faelschlich als ERROR gezaehlt.**

```
except Exception as e:
                    detail_logger.warning(f"Re-Serialisierung fehlgeschlagen: {e}")
```

*Begruendung:* temp_stage2_path wird in Zeile 2229 GESETZT, bevor der SaveAs laeuft. Scheitert der SaveAs mit einem gewoehnlichen COM-Fehler (kein Timeout - z. B. Temp-Ordner nicht beschreibbar, Datentraeger voll, Trust-Center-Ablehnung), faengt Zeile 2254 die Ausnahme ab, protokolliert nur eine Warnung und laesst temp_stage2_path auf einen Pfad zeigen, unter dem KEINE Datei existiert. Der Parallelzweig 'kein OOXML-Mapping' macht genau dafuer in Zeile 2288 ein 'temp_stage2_path = None'; hier fehlt es. Folgen im weiteren Ablauf: (a) die Praesentation zeigt weiter auf das Original, presentation.Save() in Zeile 2352 schreibt also direkt ins Original - genau das, was der Kommentar Zeile 2209-2224 ausdruecklich verhindern soll; (b) Zeile 2437 'elif temp_stage2_path is not None' wird wahr, in Zeile 2445-2450 wird ein etwaiges Schreibschutz-Attribut des Originals per os.chmod entfernt und was_read_only gesetzt; (c) die Pruefung Zeile 2454 schlaegt fehl und wirft RuntimeError; (d) der Rettungspfad Zeile 2480 kopiert nichts (Datei existiert nicht) und Zeile 2501 'return "ERROR"' verlaesst die Funktion VOR der Schreibschutz-Wiederherstellung (Zeile 2504), vor _apply_security_descriptor (2518) und vor set_file_timestamps (2519).

*Auswirkung:* Eine schreibgeschuetzte Q:\...\Vortrag.pptx, bei der der Stage-2-SaveAs an einem vollen Documents-Laufwerk scheitert: Schriftart und Metadaten werden in die Originaldatei geschrieben, das Schreibschutz-Attribut ist danach dauerhaft entfernt, mtime/atime/ctime sind veraendert (obwohl die Konfigurationsanzeige 'Zeitstempel: BLEIBEN ERHALTEN' verspricht), die Datei wird als Fehler gezaehlt, nicht in die Done-Liste geschrieben und beim naechsten Lauf erneut angefasst.

*Vorschlag:* In Zeile 2254-2255 analog zu Zeile 2288 'temp_stage2_path = None' ergaenzen und die Datei anschliessend als ERROR/SKIPPED behandeln, ohne die Fonts ueberhaupt zu schreiben - oder die Wiederherstellung von Schreibschutz/ACL/Zeitstempel aus dem else-Zweig in den finally-Block verlagern, damit sie auf allen Fehlerpfaden greift.


## `4c_ersetze_font_in_powerpoint.py` Zeile 2558  `[BEHOBEN]`

**Bei Konvertierung ohne Temp-Kopie schreibt SaveAs die neue Datei direkt in die Freigabe; bricht der spaetere Save/Timeout ab, raeumt der finally-Block diese halbfertige Zieldatei nicht weg - Original und Teilkonvertat liegen danach nebeneinander.**

```
if temp_conv_path and os.path.exists(_long_path(temp_conv_path)):
            _safe_remove_with_retry(temp_conv_path)
```

*Begruendung:* Fuer .ppt/.pps/.pot ohne Long-Path-Temp-Kopie ist das SaveAs-Ziel in Zeile 2181 direkt new_path in der Quellfreigabe ('_saveas_target = new_path'). Danach folgen die Font-Durchlaeufe und presentation.Save() (Zeile 2352). Laeuft dieser Save in den Watchdog-Timeout (Zeile 2355-2363, PowerPoint wird hart gekillt) oder scheitert die Metadaten-Speicherung mit Timeout (Zeile 2379-2387), verlaesst die Funktion mit 'return "ERROR"'. Der finally-Block (Zeile 2551-2564) raeumt file_path, temp_conv_path und temp_stage2_path auf - new_path kommt darin nicht vor. Das Original wird korrekterweise nicht geloescht (Zeile 2414 wird nie erreicht), die bereits erzeugte konvertierte Datei bleibt aber liegen.

*Auswirkung:* Q:\Sammlung\Fuehrung.ppt laeuft beim Font-Save in den 240-s-Timeout: In der Freigabe liegen danach Fuehrung.ppt (unveraendert) UND Fuehrung.pptx (konvertiert, Schriftersetzung unvollstaendig, nicht in der Done-Liste). Der naechste Lauf verarbeitet beide Dateien, findet die Namenskollision (Zeile 2155) und legt zusaetzlich Fuehrung_1.pptx an - aus einem Dokument werden drei.

*Vorschlag:* Auf allen Fehlerpfaden nach erfolgreicher Konvertierung (was_converted True, is_temp_copy False) die angelegte new_path-Datei entfernen - z. B. new_path im finally-Block mitfuehren und loeschen, solange die Verarbeitung nicht mit SUCCESS endete; alternativ auch bei is_temp_copy=False zuerst nach TEMP_PROCESS_PATH konvertieren und erst nach dem letzten Save in die Freigabe verschieben.


## `5_OCR_PDF.py` Zeile 1091  `[BEHOBEN]`

**Die Whitelist des System-Temp-Aufraeumens enthaelt 'ocrmypdf.' - seit der Temp-Umlenkung koennen diese Ordner nur noch FREMDEN Prozessen gehoeren, deren aktive Zwischendaten am Skriptende geloescht werden.**

```
"ocrmypdf.", "test_ocr", "5_ocr_pdf",
```

*Begruendung:* _worker_init_tempdir (Zeile 1124-1142) setzt tempfile.tempdir sowie TMP/TEMP auf das eigene, PID-getrennte Arbeitsverzeichnis. Das gilt nicht nur fuer die Worker (Zeile 4090, 4537), sondern seit Zeile 4960/4961 auch fuer den Hauptprozess im sequenziellen Modus. Damit legt dieses Skript im echten System-Temp KEINE 'ocrmypdf.io.*'-Ordner mehr an - jeder dort gefundene Ordner mit diesem Praefix stammt von einem anderen Prozess. _cleanup_windows_temp (Zeile 1100-1121) laeuft am Skriptende (Zeile 5186) ueber _ORIG_WINDOWS_TEMP und loescht Treffer mit shutil.rmtree(..., ignore_errors=True) bzw. os.remove. ignore_errors ueberspringt zwar gesperrte Handles, nicht aber die vielen bereits geschlossenen Zwischendateien darin - die Annahme im Kommentar Zeile 1102-1104 ('Gesperrte/in-Nutzung-Dateien werden stillschweigend uebersprungen') traegt fuer Verzeichnisse nicht. Damit verletzt der Code genau die Regel, die er selbst in Zeile 1085-1086 aufstellt ('NIEMALS pauschal leeren: Fremdprozesse legen aktive Daten ohne Lock in %TEMP% ab; blindes Loeschen zerstoert sie'). Dasselbe gilt fuer das Praefix '5_ocr_pdf': faellt temp_dir auf den Ausweichpfad in Zeile 4946/4947 ('%TEMP%\5_OCR_PDF_Processing_<pid>'), loescht eine endende Instanz das Arbeitsverzeichnis einer parallel laufenden zweiten Instanz mit - eine PID-Pruefung findet nicht statt.

*Auswirkung:* Laeuft waehrend des OCR-Laufs eine zweite PDF-Verarbeitung mit ocrmypdf (CLI-Aufruf eines Kollegen, ein anderes Werkzeug, eine aeltere Skriptversion ohne Temp-Umlenkung), werden deren Zwischenordner '%TEMP%\ocrmypdf.io.<xyz>' beim Beenden dieses Skripts geloescht; der fremde Lauf bricht mit fehlenden Seitenbildern ab. Im Ausweichfall aus Zeile 4946 trifft es zusaetzlich die Temp-Kopien und OCR-Ergebnisse einer parallelen zweiten Instanz dieses Skripts.

*Vorschlag:* Beim Aufraeumen nur eigene Artefakte adressieren: die im Lauf tatsaechlich erzeugten Pfade mitschreiben und gezielt entfernen, oder das Praefix um die eigene PID ergaenzen (z.B. nur 'ocrmypdf.'-Ordner loeschen, deren mtime aelter als der Laufbeginn ist UND die im eigenen temp_dir liegen). Fuer '5_ocr_pdf' zusaetzlich pruefen, ob die im Ordnernamen enthaltene PID noch lebt (psutil.pid_exists), und laufende Instanzen ueberspringen.


## `5_OCR_PDF.py` Zeile 1970  `[BEHOBEN]`

**Die XMP-PDF/A-Kennung einer Datei wird als 'bereits OCR-verarbeitet' gewertet - reine Bild-Scans, die als PDF/A-2u/2b vorliegen, werden dadurch ohne Textebene uebersprungen.**

```
if xmp_code in ("PDFA_2U", "PDFA_2B"):
```

*Begruendung:* Die Marker in Zeile 1924-1931 stammen aus dem Subject und werden ausschliesslich von diesem Skript gesetzt (MARKER_VALUE_*, Zeile 502-506) - sie belegen also tatsaechlich eine OCR-Verarbeitung. Der Rueckfall in Zeile 1968-1971 liest dagegen mit _read_xmp_pdfa_code nur das pdfaid-Schema aus dem XMP (Zeile 1722-1744). Das sagt lediglich aus, dass die Datei PDF/A-konform ist, und nichts darueber, ob sie eine Textebene hat. Zeile 1973 setzt daraufhin has_ocr_marker=True. Weiter unten wird das ungeprueft als 'fertig' behandelt: Zeile 3282-3287 ueberspringt bei PDFA_2U bedingungslos ('Bereits im Zielformat PDF/A-2u (Metadaten-Marker)'), Zeile 3289-3297 bei PDFA_2B, sofern nicht auf pdfa-2u aufgeruestet wird, und Zeile 3959-3961 zaehlt die Datei als 'OCR-Marker vorhanden'. PDF/A-2u ist fuer eine reine Bild-PDF trivial erfuellt (kein Text -> keine ToUnicode-Anforderung verletzt), Scanner und DMS liefern solche Dateien routinemaessig. Das ohnehin ermittelte info['has_text'] (Zeile 1940-1943) wird an dieser Stelle nicht herangezogen.

*Auswirkung:* Ein als PDF/A-2u abgelegter Scan ohne Textebene (Scanner-/DMS-Ausgabe) wird in Zeile 3283 mit '✓ Bereits im Zielformat PDF/A-2u' als SKIPPED protokolliert und nie durch die OCR geschickt. Auf einem Bestand mit solchen Dateien laeuft der Kernzweck des Skripts fuer diese Dateien nie - und zwar ohne Warnung, die Statistik meldet sie als erledigt.

*Vorschlag:* Den XMP-Rueckfall nur greifen lassen, wenn die Datei tatsaechlich Text enthaelt, z.B. in Zeile 1968 'if info["existing_marker"] is None and info["has_text"]:'. Alternativ das Ergebnis in einem eigenen Feld (z.B. info['pdfa_level']) fuehren und has_ocr_marker nur aus den Subject-Markern ableiten, damit Formatkonformitaet und OCR-Zustand nicht mehr vermischt werden.


## `5_OCR_PDF.py` Zeile 2020  `[BEHOBEN]`

**Die Zwischendatei '<name>.pdf.tmp_marker' wird neben dem Original in der Ablage angelegt, von keinem Aufraeumpfad erfasst und wandert nach einem harten Abbruch mit in die Cloud.**

```
tmp = p + ".tmp_marker"
```

*Begruendung:* p ist _lp(file_path), also die Originaldatei in der Zielablage; tmp liegt damit als '<name>.pdf.tmp_marker' direkt daneben (identisch in _add_ocr_marker_legacy, Zeile 1985). Die Aufraeumpfade decken diesen Namen nicht ab: cleanup_orphaned_backups (Zeile 1612-1683) prueft ausschliesslich low.endswith('.pdf.tmp_new') (Zeile 1645) und low.endswith('.pdf.backup') (Zeile 1658); _purge_orphan_temp_files (Zeile 1149) und der rmtree am Skriptende betreffen nur das eigene Arbeitsverzeichnis; _cleanup_windows_temp (Zeile 1100) nur den System-Temp. Genau dieses Problem ist fuer .tmp_new im Kommentar 1630-1644 beschrieben und behoben worden ('Aufgeraeumt wurde sie bisher NIRGENDS ... Damit wanderten solche Reste am Ende mit in die Cloud') - fuer .tmp_marker aber nicht. Ein harter Abbruch in dem Fenster ist im Betrieb vorgesehen: Worker-Hardkill nach WORKER_TIMEOUT (Zeile 583) und zweimal Strg+C -> os._exit(130) (Zeile 907).

*Auswirkung:* set_ocr_marker laeuft auf der Originaldatei (Aufrufe Zeile 3278 und 3573). Wird der Prozess zwischen pdf.save(tmp) (Zeile 2035) und dem os.replace in safe_replace_with_retry hart beendet, bleibt z.B. 'Q:\...\Inventarbuch.pdf.tmp_marker' dauerhaft in der Archivablage liegen. Kein spaeterer Lauf entfernt sie; bei der anschliessenden Migration wird sie als zusaetzliche, unbrauchbare Datei mit in die Cloud kopiert.

*Vorschlag:* In cleanup_orphaned_backups neben dem .pdf.tmp_new-Zweig (Zeile 1645) einen gleichartigen Zweig fuer low.endswith('.pdf.tmp_marker') ergaenzen (dieselbe Alterspruefung ueber cutoff, kein Schutzbedarf, da das Original stets unveraendert danebenliegt). Besser noch: die Marker-Zwischendatei ueberhaupt nicht in der Ablage anlegen, sondern in temp_dir schreiben und von dort per safe_replace_with_retry ueberschieben.

*Nachpruefung 14.08.2026:* Bereits behoben, nur nicht markiert -
`cleanup_orphaned_backups` erfasst `.pdf.tmp_marker` inzwischen zusammen mit
`.pdf.tmp_new`, samt Kommentar, der genau diesen Fehler beschreibt. Keine
Aenderung noetig.



## `5_OCR_PDF.py` Zeile 3247  `[BEHOBEN]`

**Nach fitz.FileDataError laufen alle Schutzpruefungen auf leeren Werten – Formular-/Groessenschutz greift nicht, pdf_info wird nach der Reparatur nicht neu gelesen**

```
if pdf_info["needs_repair"]:
```

*Begruendung:* get_pdf_info() kehrt bei 'except fitz.FileDataError' sofort zurueck (Zeile 1897-1900) und liefert dabei needs_repair=True, is_valid=True, aber pages=0, has_forms=False, is_oversized=False, max_image_pixels=0, subject='' und existing_marker=None. Nachgestellt mit einer Datei aus '%PDF-1.4' + Fuellbytes: fitz wirft FileDataError. Alle Schutz-Returns davor (Zeile 3210 has_forms 'wird zum Schutz uebersprungen', 3217 is_oversized, 3226 max_image_pixels, 3237 check_memory_before_ocr mit pages=0 -> gibt wegen 'pages < 50' sofort True zurueck) passieren dann folgenlos. Anders als im Entschluesselungszweig (Zeile 3196 'pdf_info = get_pdf_info(file_path)') wird pdf_info nach der Reparatur NICHT neu erhoben. Hinzu kommt: repair_pdf ruft in Zeile 2057 dasselbe fitz.open() auf, das gerade FileDataError geworfen hat – in genau diesem Fall kann die Reparatur nie gelingen, der Ablauf landet immer bei 'Reparatur fehlgeschlagen – versuche OCR trotzdem' (Zeile 3253).

*Auswirkung:* Eine PDF, die PyMuPDF nicht oeffnen kann, die ocrmypdf/qpdf aber verarbeitet, wird ohne Formularfeld-, Uebergroessen- und RAM-Pruefung an Ort und Stelle ersetzt; enthaelt sie AcroForm-Felder, gehen die Formulardaten verloren, obwohl der Schutz dafuer existiert. Ausserdem werden Seitenzahl 0 in CSV/Statistik geschrieben und veraPDF mit dem Mindest-Timeout aufgerufen (_verapdf_timeout_for(0)).

*Vorschlag:* Nach erfolgreicher Reparatur pdf_info neu einlesen (wie in Zeile 3196) und die Schutzpruefungen (Formular/Uebergroesse/Bildpixel/RAM) danach erneut ausfuehren; scheitert die Reparatur bei needs_repair aus dem FileDataError-Zweig, die Datei als ERROR/SKIP behandeln statt blind zu OCRen.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `5_OCR_PDF.py` Zeile 3249  `[BEHOBEN]`

**repair_pdf ersetzt das Original vor der Zeitstempel-/ACL-Sicherung – restauriert wird der Reparaturzeitpunkt**

```
if repair_pdf(file_path, temp_dir):
```

*Begruendung:* repair_pdf() (Zeile 2053-2060) schreibt die reparierte Datei per safe_replace_with_retry ueber das Original und sichert dabei weder Zeitstempel noch Sicherheitsdeskriptor (anders als try_remove_empty_password, das beides in Zeile 2083/2084 vorher liest). Auch hier laeuft die Sicherung im Aufrufer erst danach: 'orig_times = _read_file_times(file_path)' in Zeile 3409. Am Ende (Zeile 3707-3708) wird deshalb genau der Zeitpunkt der Reparatur zurueckgeschrieben, nicht das urspruengliche Aenderungsdatum. Gleicher Mechanismus wie oben nachgestellt (os.replace setzt mtime neu).

*Auswirkung:* Bei jeder reparaturbeduerftigen PDF (Erkennung ueber doc.is_repaired, Zeile 1915) geht das Original-Aenderungsdatum unwiederbringlich verloren, ebenso explizit gesetzte Datei-ACEs – obwohl der Rest des Skripts beides bewusst erhaelt.

*Vorschlag:* Sicherung von orig_times/orig_sd vor den Reparaturblock (Zeile 3247) verschieben, oder in repair_pdf analog zu try_remove_empty_password Zeitstempel und Sicherheitsdeskriptor vor dem Ersetzen lesen und danach wieder anwenden.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `5_OCR_PDF.py` Zeile 3573  `[BEHOBEN]`

**Zwischendatei <name>.pdf.tmp_marker in der Ablage wird von keinem Aufraeumpfad erfasst**

```
if set_ocr_marker(file_path,
                                  _compose_subject(final_marker_value, original_subject)):
```

*Begruendung:* set_ocr_marker() legt seine Zwischenkopie direkt neben dem Original an: 'tmp = p + ".tmp_marker"' (Zeile 2020, ebenso Zeile 1985 in _add_ocr_marker_legacy) und schiebt sie erst danach per safe_replace_with_retry darueber. Wird der Worker in diesem Fenster hart getoetet – _terminate_pool_workers ruft proc.terminate()/proc.kill() (Zeile 3809-3813), der Signal-Handler beendet mit os._exit(130) (Zeile 907) – laeuft kein except/finally, und '<name>.pdf.tmp_marker' bleibt in der Archiv-Ablage liegen. cleanup_orphaned_backups kennt nur '.pdf.tmp_new' (Zeile 1645) und '.pdf.backup' (Zeile 1658); _purge_orphan_temp_files raeumt ausschliesslich das eigene Temp-Verzeichnis. Der Kommentar ab Zeile 1630 beschreibt dieselbe Luecke fuer .tmp_new als behobenen Fehler – fuer .tmp_marker ist sie offen geblieben. Betroffen sind die Aufrufe auf file_path (Zeilen 3278 und 3573), nicht der auf temp_output (Zeile 3519).

*Auswirkung:* Reste im Produktivbestand, die niemand aufraeumt und die bei der spaeteren Migration mit in die Cloud wandern; ausserdem verwirrende Dateileichen mit vollstaendigem PDF-Inhalt neben dem Original.

*Vorschlag:* In cleanup_orphaned_backups zusaetzlich '.pdf.tmp_marker' behandeln (analog zu .tmp_new: Original existiert und ist lesbar -> loeschen), oder die Marker-Zwischendatei im Temp-Verzeichnis statt neben dem Original erzeugen.


## `5_OCR_PDF.py` Zeile 3668  `[BEHOBEN]`

**Speicher-Fallback setzt oversample=200 – das ist eine DPI-Untergrenze, keine Reduktion**

```
current_options["oversample"] = 200
```

*Begruendung:* Die Meldung eine Zeile darueber lautet 'wechsle auf Ghostscript-Pfad + DPI-Reduktion'. ocrmypdf verwendet oversample jedoch als Minimum: in _pipeline.py (get_page_square_dpi/get_canvas_square_dpi) geht der Wert in ein max(image_dpi.x, image_dpi.y, _vector_page_dpi, options.oversample) ein; die CLI-Hilfe lautet 'Oversample images to at least the specified DPI'. Bei einem typischen 300-dpi-Scan bleibt der Wert wirkungslos, bei niedriger aufgeloesten Seiten erhoeht er die Rasteraufloesung und damit den Speicherbedarf. Die beabsichtigte Entlastung nach MemoryError/'failed to fill bitmap' tritt also nie ein; wirksam ist nur der Plugin-Reset in Zeile 3667.

*Auswirkung:* Der Wiederholungsversuch nach einem Speicherproblem laeuft mit unveraendertem oder groesserem Speicherbedarf und scheitert erneut; die Datei landet unnoetig als ERROR_OCR im Protokoll, statt mit reduzierter Aufloesung durchzulaufen.

*Vorschlag:* oversample nicht setzen (bzw. entfernen) und stattdessen die tatsaechlich begrenzenden Optionen verwenden – z.B. jobs=1 fuer diesen Versuch und/oder max_image_mpixels senken; die Textmeldung entsprechend anpassen.


## `5_OCR_PDF.py` Zeile 4220  `[BEHOBEN]`

**Nach einem Worker-Timeout werden bereits fertige, aber noch nicht ausgewertete Ergebnisse verworfen – Datei ist ersetzt, aber ohne CSV-Zeile, ohne Resume-Eintrag und ohne Statistik**

```
for f, info in list(future_to_path.items()):
                                    fp_inflight, _ = info
                                    with heartbeat_lock:
                                        active_workers.pop(fp_inflight, None)
                                    if f in timed_out_futures:
                                        continue
```

*Begruendung:* concurrent.futures.wait (Zeile 4189) liefert 'done' zurueck; die Timeout-Erkennung (Zeile 4196-4201) laeuft aber VOR der Auswertungsschleife 'for future in done' (Zeile 4254). Wird ein Timeout erkannt, bricht der Code bei Zeile 4252 ab – die Futures aus 'done' stehen zu diesem Zeitpunkt noch in future_to_path und laufen deshalb in die Requeue-Schleife: ihr fertiges Ergebnis wird nie ueber _handle_result verarbeitet, sondern der Pfad wird erneut eingereiht (Zeile 4248) und future_to_path.clear() (4249) wirft das Ergebnis weg. Mit workers=3 ist das Zusammentreffen realistisch: Worker A haengt 60 Minuten, waehrend B/C im Minutentakt fertig werden – bei genau dem wait()-Durchlauf, in dem A den WORKER_TIMEOUT ueberschreitet, ist mindestens ein fertiges Ergebnis in 'done'. Derselbe Fehler steht im BrokenProcessPool-Zweig (Zeile 4287-4294). Zusaetzlich zaehlt requeue_counts fuer diese Datei hoch; beim zweiten Vorfall wird eine tatsaechlich erfolgreich verarbeitete Datei als ERROR_REQUEUE_LIMIT ins Fehler-CSV geschrieben (Zeile 4239).

*Auswirkung:* Die betroffene PDF ist auf der Ablage bereits ersetzt, im Protokoll (CSV und _protokoll/migration.jsonl) fehlt sie jedoch vollstaendig – die Nachweiskette fuer den ersetzten Bestand hat eine Luecke. Ausserdem fehlt der Resume-Eintrag, die Datei wird sofort erneut eingelesen (und dann wegen OCR-Marker als SKIPPED gezaehlt), und im Extremfall erscheint sie als ERROR_REQUEUE_LIMIT, obwohl sie fehlerfrei verarbeitet wurde. Statistik und Fehlerliste sind damit falsch.

*Vorschlag:* Die fertigen Futures aus 'done' zuerst auswerten (_handle_result) und erst danach die Timeout-/Pool-Bruch-Behandlung ausfuehren; bzw. in beiden Requeue-Schleifen die Futures aus 'done' ueberspringen und ihr Ergebnis regulaer verarbeiten, statt den Pfad blind neu einzureihen.


## `5_OCR_PDF.py` Zeile 4622  `[BEHOBEN]`

**Watch-Modus: jeder Fehler wird als ERROR_TIMEOUT gezaehlt und protokolliert, auch wenn gar kein Timeout vorlag**

```
except Exception as e:
                    with thread_lock:
                        stats["ERROR"] += 1
                        stats["ERROR_TIMEOUT"] = stats.get("ERROR_TIMEOUT", 0) + 1
                        print(f"  [ERROR] {os.path.basename(file_path)}: {_fmt_exc(e)}")
                    log_error("watch_mode", f"Fehler bei {file_path}: {_fmt_exc(e)}", exc_info=True)
                    _write_error_csv(file_path, "ERROR_TIMEOUT", _fmt_exc(e))
```

*Begruendung:* Der zugehoerige try-Block beginnt in Zeile 4587 und umfasst nicht nur den Timeout-Pfad (4591-4601), sondern auch den BrokenProcessPool-Pfad (4602-4607, der bewusst ein generisches Exception-Objekt wirft), das Einsammeln des Ergebnisses, die Ausgabe der Meldungen sowie _flush_log_entries/_write_csv_row (4609-4617). Jede dieser Fehlerquellen landet im selben except-Zweig und wird unterschiedslos als ERROR_TIMEOUT gebucht – sowohl im Statistik-Dict als auch als 'detail'-Spalte im Ergebnis-CSV.

*Auswirkung:* Falsche Fehlerstatistik und falscher Fehlergrund im CSV: gestorbene Worker-Prozesse, Schreibfehler beim Protokoll und sonstige Ausnahmen erscheinen in der Auswertung (print_detailed_stats, Zeile 4764 'Worker-Timeout') als Zeitueberschreitungen. Die Ursachensuche nach einem Nachtlauf laeuft dadurch in die falsche Richtung; ein echtes Timeout-Problem wird vorgetaeuscht bzw. ein reales Pool-Problem verdeckt.

*Vorschlag:* Den Grund mitfuehren: die drei Faelle getrennt behandeln (concurrent.futures.TimeoutError -> ERROR_TIMEOUT, BrokenProcessPool -> ERROR_OCR bzw. eigener Schluessel, Rest -> ERROR_OCR) oder eine lokale Variable 'fehlergrund' im jeweiligen inneren except setzen und im aeusseren except statt der Konstante verwenden.


## `9_fehlerhafte_Dateien_finden.ps1` Zeile 314  `[BEHOBEN]`

**Preflight liefert immer HRESULT 0x80131501 (PowerShell-Wrapper) statt des echten Datei-HRESULT - der ganze HRESULT-Switch der Fehlerkategorisierung ist auf diesem Weg tot**

```
return @{ OK = $false; Error = $_.Exception.Message; HResult = $_.Exception.HResult }
```

*Begruendung:* Beim Aufruf einer .NET-Methode verpackt PowerShell die Ausnahme in eine System.Management.Automation.MethodInvocationException. Deren HResult ist konstant 0x80131501, nicht der HResult der eigentlichen IOException. Nachgestellt in Windows PowerShell 5.1 mit einer exklusiv gesperrten Datei: $_.Exception.GetType() = MethodInvocationException, $_.Exception.HResult = 0x80131501, waehrend die darunterliegende System.IO.IOException 0x80070020 (Sharing Violation) traegt. Damit trifft in Get-ErrorCategory (Zeile 395-403) kein einziger Switch-Zweig - weder 0x80070020 noch 0x80070005. Die Textregeln fangen den Fall nicht auf: die deutsche Meldung lautet '... zugreifen, da sie von einem anderen Prozess verwendet wird', das Muster in Zeile 411 heisst aber 'wird verwendet' (andere Wortstellung) und 'zugriff' in Zeile 414 ist kein Teilstring von 'zugreifen'. Nachgestellt: Get-ErrorCategory liefert fuer die gesperrte Datei 'Unbekannt'. (Der COM-Pfad ist NICHT betroffen - dort kommt die COMException ungekapselt an, gepruefte Gegenprobe: 0x800A004C bzw. 0x800A03EC kamen korrekt durch.)

*Auswirkung:* Eine Datei, die ein Kollege gerade geoeffnet hat - auf einer Archivfreigabe der Normalfall - erscheint im Bericht unter '[Preflight] Nicht klassifizierbarer Fehler ... (HRESULT 0x80131501)' in der Kategorie 'Unbekannt' statt 'Datei gesperrt'. Die Kategorienstatistik und die Filterung im XLSX-Bericht sind dadurch systematisch falsch; harmlose Sperren sehen aus wie ungeklaerte Defekte und werden nachtraeglich von Hand untersucht.

*Vorschlag:* Den inneren HRESULT durchreichen: $ex = $_.Exception; if ($ex.InnerException) { $ex = $ex.InnerException }; return @{ OK = $false; Error = $ex.Message; HResult = $ex.HResult }. Zusaetzlich in Get-ErrorCategory Zeile 411 die tatsaechlichen Wortstellungen aufnehmen ('verwendet wird', 'used by another process') und in Zeile 414 'zugreifen' bzw. 'verweigert' ergaenzen.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `9_fehlerhafte_Dateien_finden.ps1` Zeile 396  `[BEHOBEN]`

**Kennwortgeschuetzte Excel-Dateien werden als 'Office-Fehler' berichtet - die Kategorie 'Passwortschutz' ist fuer Excel unerreichbar**

```
"0x800A03EC" { return @{ Kategorie = "Office-Fehler";      Details = "Office meldet einen allgemeinen Anwendungsfehler (0x800A03EC). Ursprungstext: $Message" } }
```

*Begruendung:* Der Kommentar in Zeile 1609-1615 begruendet die Injektion von $Config.DummyPassword damit, dass passwortgeschuetzte Dateien sofort erkannt statt in den 120s-Timeout laufen sollen. Nachgestellt mit echtem Excel-COM: eine mit Oeffnungskennwort gespeicherte .xlsx, geoeffnet mit exakt dem Open-Aufruf aus Zeile 1617-1619, wirft COMException 0x800A03EC mit dem Text 'Die Open-Eigenschaft des Workbooks-Objektes kann nicht zugeordnet werden.' Der Switch in dieser Zeile greift vor allen Textregeln und liefert 'Office-Fehler'. Die Passwort-Regel in Zeile 405 wuerde aber auch dann nicht greifen, wenn sie vor dem Switch stuende - die Meldung enthaelt kein einziges der Woerter passwort/password/kennwort/verschluessel/encrypt (im Test geprueft: Treffer = False). Die Datei hebt Pfadlaenge (Zeile 376) und Namenskonflikt (Zeile 388) genau wegen dieses Problems ueber den Switch; der Passwortfall, fuer den eigens ein Dummy-Kennwort erzeugt wird (Zeile 87), ist uebersehen worden.

*Auswirkung:* Jede kennwortgeschuetzte Excel-Datei im Bestand landet im Bericht als 'Office-Fehler / Office meldet einen allgemeinen Anwendungsfehler (0x800A03EC)' und ist damit nicht von echten Defekten zu unterscheiden. Die Kategorie 'Passwortschutz' bleibt fuer Excel leer, obwohl passwortgeschuetzte Tabellen der haeufigste gutartige Treffer sind. Wer den Bericht abarbeitet, oeffnet diese Dateien als vermeintlich beschaedigt einzeln von Hand.

*Vorschlag:* Den Passwortschutz nicht aus der Office-Meldung ableiten, sondern vor dem COM-Open aus der Datei selbst: eine verschluesselte OOXML-Datei ist kein ZIP, sondern ein CFB-Container mit der Signatur D0 CF 11 E0 A1 B1 1A E1 (statt 'PK'). System.IO.Compression ist in Zeile 14-15 ohnehin schon geladen; die ersten acht Bytes zu lesen genuegt und kostet keinen COM-Aufruf. Alternativ - analog zu Pfadlaenge und Namenskonflikt - im Aufrufer einen eindeutigen Marker an $reason haengen und in Get-ErrorCategory vor dem Switch auswerten.

*Vom Pruefer ausgefuehrt nachgestellt.*



# GERING

## `0_Vorab-Check.py` Zeile 273  `[BEHOBEN]`

**root.setLevel(logging.INFO) verwirft alle 12 logging.debug()-Aufrufe, mit denen verschluckte Exceptions protokolliert werden sollen - die catch-Bloecke sind faktisch weiterhin stumm.**

```
root.setLevel(logging.INFO)
```

*Begruendung:* Das Skript enthaelt 12 Stellen nach dem Muster 'except Exception as _e: logging.debug(f"...: Exception verworfen: {_e!r}")' (z. B. Zeile 269, 305, 422, 491, 637, 1068, 1106, 2739). Ein auf INFO gesetzter Root-Logger filtert DEBUG-Records vor jedem Handler heraus. Nachgestellt: Root auf INFO, FileHandler angehaengt, logging.debug('DEBUG-ZEILE') und logging.info('INFO-ZEILE') -> die Datei enthaelt nur 'INFO-ZEILE'. Der offensichtliche Zweck dieser Zeilen (keine stillen except-Bloecke) wird durch die Logger-Konfiguration aufgehoben.

*Auswirkung:* Beispiel Zeile 2735-2739: das Loeschen einer alten stage_result.json schlaegt fehl (Datei durch Stage 2 eines Vorlaufs gesperrt). Der Grund wird nach logging.debug geschrieben und verworfen; im Logfile steht nichts. Anschliessend liest _post_stage2_recheck() ueber read_stage_result() das ALTE Ergebnis und meldet 'Stage 2 hat installiert: ...' fuer Tools, die dieser Lauf nie angefasst hat - ohne dass im Log eine Spur auf die Ursache fuehrt.

*Vorschlag:* root.setLevel(logging.DEBUG) setzen und stattdessen dem FileHandler das gewuenschte Level geben (fh.setLevel(logging.DEBUG) fuers Log, Konsole laeuft ohnehin ueber print()). Alternativ die betreffenden Aufrufe auf logging.info/warning heben.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `10_dateinamen_bereinigen.py` Zeile 760  `[BEHOBEN]`

**Die Whitelist-Tokens '_─' und ' ─' haben entgegen dem eigenen Kommentar keinen Buchstaben-Kontext und ersetzen U+2500 auch in legitimen Namen.**

```
" ─": " Ä",   # mit Leerzeichen davor
```

*Begruendung:* Der Kommentar in Zeile 755-758 nennt als Schutz ausdruecklich 'Token mit Buchstaben-Kontext, weil U+2500 theoretisch auch in ASCII-Art-Dateinamen legitim sein koennte'. Fuer die Tokens '─n' und '─N' stimmt das, fuer '_─' (Zeile 759) und ' ─' (Zeile 760) nicht: dort steht links nur Unterstrich bzw. Leerzeichen. Der beschriebene Schutz existiert fuer diese beiden Eintraege also nicht. Da U+2500 selbst im Hint-Zeichensatz liegt, oeffnet ein solcher Name den Heiler zusaetzlich immer selbst.

*Auswirkung:* Nachgestellt: 'ASCII-Art ─── Trennlinie Öl.txt' -> 'ASCII-Art Ä── Trennlinie Öl.txt'. Der Name wird auf der Platte geaendert und ist danach sinnlos.

*Vorschlag:* Die beiden Tokens auf einen Buchstaben-Kontext ziehen, z. B. '_─n': '_Än' und ' ─n': ' Än' (analog fuer die Grossschreibvariante), damit eine reine Trennlinie aus U+2500 nicht getroffen wird.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `2a_entferne_schutz_word.ps1` Zeile 2040  `[BEHOBEN]`

**COM-Timeouts werden in der Endstatistik als 'Verschluesselt' ausgewiesen.**

```
$stats.Encrypted++
```

*Begruendung:* Der Zweig Z. 2037 fasst die Kategorien 'Timeout' und 'Passwort' zusammen und erhoeht fuer beide $stats.Encrypted. Get-ConversionErrorCategory liefert 'Timeout' fuer die von Convert-DocToDocx bei Wait-Job-Ablauf geworfene Meldung 'TIMEOUT: Datei reagiert nicht ...' (Z. 1220) - das ist ein haengender COM-Aufruf bzw. ein offener Word-Dialog, nicht zwingend eine Verschluesselung. Konsole (Z. 2041) und CSV (Z. 2043) fuehren die Kategorie korrekt, nur der Zaehler nicht.

*Auswirkung:* Ein Lauf mit 300 Dateien, bei dem Word wegen eines nicht konfigurierten Trust Centers 120-mal in den 45-s-Timeout laeuft, meldet am Ende 'Verschluesselt: 120'. Der Betreiber sucht daraufhin nach Passwoertern, statt die eigentliche Ursache (Trust Center / Office-Profil) zu beheben.

*Vorschlag:* Eigenen Zaehler einfuehren (z. B. $stats.Timeouts) und in der Endstatistik getrennt ausweisen; $stats.Encrypted nur bei $category -eq 'Passwort' erhoehen.


## `2b_entferne_schutz_excel.ps1` Zeile 2055  `[BEHOBEN]`

**Sperrdatei-Loeschung: der Zaehler und das Protokoll melden 'Junk geloescht', auch wenn das Delete in den leeren catch gelaufen ist.**

```
$stats.Junk++
```

*Begruendung:* In Zeile 2054 steht 'try { [System.IO.File]::Delete((Add-LongPathPrefix $file.FullName)) } catch {}'. Der catch ist leer, es gibt keine Erfolgspruefung. Unmittelbar danach wird $stats.Junk unbedingt erhoeht, Write-DetailedLog schreibt 'Junk entfernt' und Write-CsvLog setzt Status JUNK mit Aktion 'Junk geloescht'. Zum Vergleich: an allen anderen Loeschstellen prueft das Skript sehr wohl nach - Remove-OrphanedBackups zaehlt in Zeile 622/625 getrennt Deleted und Failed, und Invoke-WindowsTempCleanup prueft in Zeile 735/738 explizit mit Test-Path nach, bevor es hochzaehlt. Hier fehlt diese Nachpruefung als einziger Loeschstelle.

*Auswirkung:* Eine '~$'-Sperrdatei ist genau dann noch gesperrt, wenn die zugehoerige Mappe tatsaechlich noch irgendwo geoeffnet ist. In diesem Fall wirft File.Delete eine IOException, die verschluckt wird - die Datei bleibt liegen, Statistik und CSV behaupten aber die Loeschung. Konkret: 30 alte Sperrdateien, davon 12 noch von offenen Sitzungen gehalten, ergeben die Meldung 'Junk geloescht: 30', obwohl 12 unveraendert auf der Freigabe stehen. Wer die CSV zur Abnahme der Migration heranzieht, haelt die Ablage faelschlich fuer bereinigt und uebersieht, dass genau die kritischen Faelle (aktive Sperren) uebrig sind.

*Vorschlag:* Erfolg pruefen und getrennt zaehlen, analog zu Remove-OrphanedBackups: das Delete in try/catch mit einem eigenen Fehlerzweig fuehren oder danach mit [System.IO.File]::Exists gegenpruefen; nur bei tatsaechlichem Verschwinden $stats.Junk++ und Status JUNK schreiben, sonst einen eigenen Zaehler (z. B. $stats.JunkFailed) hochzaehlen und den Fall als WARN mit der Ausnahmemeldung protokollieren.


## `2c_entferne_schutz_powerpoint.ps1` Zeile 1954  `[BEHOBEN]`

**Korrupte (nicht entpackbare) Dateien werden auf den Zaehler 'Verschluesselt' gebucht.**

```
if ($zipStatus -eq "Invalid") {
                $stats.Encrypted++
```

*Begruendung:* $zipStatus = 'Invalid' bedeutet laut Test-IsValidZip (Z. 640-642) genau das Gegenteil von verschluesselt: die Datei ist weder IOException-gesperrt noch zugriffsgeschuetzt, sondern kein gueltiges ZIP - Log und CSV benennen das in Z. 1957-1959 korrekt als 'korrupt'. Der Zaehler dagegen laeuft auf $stats.Encrypted, und die Abschlussstatistik gibt ihn in Z. 2247 als 'Verschluesselt:' aus. Bei der Timeout-Buchung in Z. 1918 ist die Gleichsetzung durch den Skriptkopf (Z. 30-31) ausdruecklich gedeckt, hier nicht - fuer korrupte Dateien existiert kein eigener Zaehler.

*Auswirkung:* Eine Ablage mit 30 beschaedigten .pptx meldet am Ende 'Verschluesselt: 30'. Der Verantwortliche schliesst daraus auf passwortgeschuetzte Dateien und sucht Passwoerter, statt die 30 defekten Dateien aus dem Backup zu restaurieren. Die richtige Information steht nur verstreut in CSV und Detail-Log.

*Vorschlag:* Eigenen Zaehler einfuehren: $stats.Corrupt initialisieren (bei Z. 1709-1721), in Z. 1954 $stats.Corrupt++ statt $stats.Encrypted++ setzen und in der Statistik als 'Korrupt/kein ZIP:' getrennt ausgeben.


## `3a_doc_docx_auf_neueste_Version_aktualisieren.py` Zeile 316  `[BEHOBEN]`

**Nach Strg+C werden die profil-persistenten Word-Optionen nie zurueckgesetzt - der Signal-Handler verhindert sogar das Zuruecksetzen im finalen finally**

```
word_app_global.Quit(SaveChanges=COM_FALSE)
```

*Begruendung:* Der Kommentar in Zeile 205-208 sichert ausdruecklich zu, dass UpdateLinksAtOpen und DoNotPromptForConvert 'im finalen finally wiederhergestellt' werden, 'damit die Word-Einstellungen des Benutzers nach dem Lauf unveraendert sind'. Der Signal-Handler beendet Word aber ohne vorherigen _restore_word_options()-Aufruf und setzt anschliessend word_app_global = None (Zeile 319). Das nachfolgende sys.exit(1) loest SystemExit aus; das finally in Zeile 3730 laeuft zwar, findet dort aber 'word_app_global is not None' als False vor und ueberspringt _restore_word_options() (3734). Dieselbe Luecke greift, wenn die letzte Word-Instanz vom Watchdog getoetet wurde - dann scheitert setattr in _restore_word_options() still (Zeile 1214-1215).

*Auswirkung:* Bricht der Anwender einen Lauf mit Strg+C ab (bei Laufzeiten von Stunden ueber Q:/R: der Normalfall), behaelt sein Word-Profil dauerhaft DoNotPromptForConvert=True und UpdateLinksAtOpen=False. Word fragt danach beim Oeffnen alter Formate nicht mehr nach und aktualisiert Feldverknuepfungen beim Oeffnen nicht mehr - eine stille, nicht dokumentierte Veraenderung der Arbeitsumgebung, die niemand mit dem Skript in Verbindung bringt.

*Vorschlag:* Im _signal_handler vor dem Quit '_restore_word_options(word_app_global)' aufrufen. Zusaetzlich die Originalwerte unabhaengig von einer lebenden COM-Instanz sichern/zuruecksetzen (z.B. beim Restore notfalls eine frische Word-Instanz starten), damit auch nach einem Watchdog-Kill der Ausgangszustand wiederhergestellt wird.


## `3a_doc_docx_auf_neueste_Version_aktualisieren.py` Zeile 1567  `[BEHOBEN]`

**check_required_modules() kann nie ausloesen - alle geprueften Module werden bereits beim Import auf Modulebene geladen**

```
missing = [pkg for mod, pkg in required.items()
```

*Begruendung:* Geprueft werden win32com.client, psutil, tqdm und msoffcrypto (Zeile 1561-1566). Genau diese Module werden aber schon in den Zeilen 84 (win32com.client), 85 (psutil), 87 (from tqdm import tqdm) und 92 (import msoffcrypto) unbedingt importiert - also lange bevor der Aufruf in Zeile 3317 stattfindet. Fehlt eines davon, scheitert bereits der Modulimport mit ModuleNotFoundError; die Funktion wird nie erreicht. Der gesamte Zweig ist toter Code (Fehlerklasse 'ein ganzer Programmzweig lief nie').

*Auswirkung:* Auf einem frisch aufgesetzten Arbeitsplatz ohne msoffcrypto-tool bricht das Skript mit einem nackten Python-Traceback ab, statt die vorgesehene Meldung 'FEHLENDE MODULE: pip install msoffcrypto-tool' auszugeben. Der Anwender - laut Kopfzeilen des Skripts ausdruecklich ein Einsteiger, dem sogar der pip-Befehl vorgegeben wird - steht ohne Handlungsanweisung da.

*Vorschlag:* Die vier Importe aus dem Modulkopf in die Funktionen verlagern, die sie brauchen (bzw. lazy importieren), oder den Modulkopf mit try/except ImportError umschliessen und im except-Zweig dieselbe Klartextmeldung ausgeben. Danach check_required_modules() vor dem ersten Zugriff aufrufen.


## `3b_xls_xlsx_auf_neueste_Version_aktualisieren.py` Zeile 1890  `[BEHOBEN]`

**Die definedName-Regex verschmilzt selbstschliessende Eintraege mit dem Folgeeintrag: der Pre-Clean entfernt den Konflikt dann nicht bzw. loescht einen fremden Namen mit.**

```
entries = re.findall(r"<definedName\b[^>]*?>.*?</definedName>", inner, re.DOTALL)
```

*Begruendung:* Das Muster verlangt ein schliessendes </definedName>. Ein selbstschliessender Eintrag <definedName name="X"/> hat keines, deshalb frisst '.*?</definedName>' den Eintrag zusammen mit dem NAECHSTEN Eintrag zu einem einzigen 'entry'. Die anschliessende Auswertung liest per re.search nur das ERSTE name-Attribut (Zeile 1899) und entscheidet fuer beide zusammen. Nachgestellt: bei '<definedName name="Wichtig"/><definedName name="_xlnm.Print_Area" ...>...</definedName>' wird ein einziger entry mit dem Namen 'Wichtig' erkannt -> nicht geloescht, der _xlnm.Print_Area-Eintrag bleibt trotz Pre-Clean erhalten. Umgekehrt reisst ein fuehrender _xlnm-Eintrag den nachfolgenden Benutzernamen mit in die Loeschung. Steht ein selbstschliessender Eintrag am Blockende, faellt er ganz aus 'entries' heraus und verschwindet beim Neuaufbau in Zeile 1928 ('<definedNames>' + nur die kept-Eintraege) ersatzlos.

*Auswirkung:* Bei einer von einem Drittanbieter-Werkzeug erzeugten Mappe mit selbstschliessenden definedName-Eintraegen: der Pre-Clean meldet Erfolg, der Namenskonflikt bleibt aber bestehen (Excel haengt beim Re-Open erneut im Modaldialog) - oder ein benannter Bereich des Anwenders wird beim Umschreiben der Datei still geloescht.

*Vorschlag:* Muster um die selbstschliessende Form erweitern, z.B. r'<definedName\\b[^>]*/>|<definedName\\b[^>]*?>.*?</definedName>', oder den Block mit xml.etree/lxml statt per Regex verarbeiten.

*Vom Pruefer ausgefuehrt nachgestellt.*


## `3c_ppt_pptx_auf_neueste_Version_aktualisieren.py` Zeile 1629  `[BEHOBEN]`

**Staging-Datei '<Ziel>.tmp_new' auf der Ablage wird von keinem Aufraeumpfad erfasst, wenn der Prozess hart abbricht.**

```
stage      = dst + ".tmp_new"
```

*Begruendung:* Die Staging-Datei entsteht im ZIELverzeichnis (Q:/R:/DFS), nicht im PID-eigenen Temp-Ordner. Sie wird nur im except-Zweig von robust_move (Zeile 1660-1664) entfernt. Der Signal-Handler (Zeile 661-695) raeumt ausschliesslich TEMP_PROCESS_PATH und beendet danach mit os._exit(1); der finally-Block im Hauptteil (Zeile 3143-3172) ebenfalls nur TEMP_PROCESS_PATH und %TEMP%. file_generator entfernt nur '~$'- und '._'-Reste (Zeile 2067), nicht '*.tmp_new'. Fuer zurueckgebliebene '.bak' gibt es dagegen eine ausdrueckliche Behandlung (stale_bak, Zeile 2465-2486) - fuer '.tmp_new' fehlt das Gegenstueck.

*Auswirkung:* Strg+C oder 'Task beenden' waehrend eines volumeuebergreifenden Verschiebens hinterlaesst 'Jahresbericht.pptx.tmp_new' dauerhaft neben der Originaldatei auf der Archivablage. Kein spaeterer Lauf entfernt oder erkennt sie; die Reste sammeln sich ueber die Laeufe an und sind fuer Anwender nicht von echten Dateien zu unterscheiden.

*Vorschlag:* Den aktuell benutzten stage-Pfad in einer globalen Variablen fuehren und im Signal-Handler sowie im finally-Block des Hauptteils mitloeschen; zusaetzlich in file_generator verwaiste '*.tmp_new' (nicht gesperrt, wie bei den '~$'-Resten geprueft) mit einraeumen.


## `4b_ersetze_font_in_excel.py` Zeile 3663  `[BEHOBEN]`

**Nach dem Log-Fallback zeigen die Abschlussmeldungen weiterhin die nicht beschreibbaren Ursprungspfade der Logdateien an.**

```
print(f"  Fehler-Log:  {os.path.abspath(LOG_FILE)}")
    print(f"  Detail-Log:  {os.path.abspath(DETAILED_LOG_FILE)}")
```

*Begruendung:* _setup_logging biegt im Fallback-Zweig (Zeile 333-347) RUN_SUMMARY_FILE, CONVERSIONS_CSV und _LOG_DIR global um, die Handler schreiben nach fh_path/dh_path - die Globals LOG_FILE und DETAILED_LOG_FILE bleiben aber unveraendert auf BASE_DIR stehen (global-Deklaration Zeile 327 nennt sie nicht). Die Abschlussausgabe und die Probelauf-Ausgabe in Zeile 3518 nennen daher Pfade, unter denen keine Datei existiert.

*Auswirkung:* Laeuft das Skript aus einem schreibgeschuetzten Ordner (z.B. Freigabe oder Program Files), verweist die Endausgabe auf ...\Claude Skripte\4b_..._20260813_101500.log; die Logs liegen tatsaechlich unter Dokumente\4b_ersetze_font_in_excel_logs.

*Vorschlag:* LOG_FILE und DETAILED_LOG_FILE in die global-Zeile von _setup_logging aufnehmen und im Fallback auf fh_path/dh_path setzen.


## `5_OCR_PDF.py` Zeile 4911  `[BEHOBEN]`

**Probelauf raeumt trotzdem auf: --dry-run schliesst das Loeschen von *.backup und *.tmp_new nicht aus**

```
do_cleanup = False
    if args.cleanup_backups:
        do_cleanup = True
    elif not args.auto and not args.watch:
```

*Begruendung:* Weder der --cleanup-backups-Zweig noch die interaktive Rueckfrage (Zeile 4924) beruecksichtigen args.dry_run. cleanup_orphaned_backups (Zeile 1612) loescht anschliessend tatsaechlich Dateien: verwaiste *.pdf.tmp_new (Zeile 1645-1654) und *.pdf.backup (1658-1676). Das widerspricht der Zusage des Probelaufs ('Probelauf: zeigt je Datei die Entscheidung, ohne etwas zu schreiben', Docstring Zeile 3898, und 'Es wurde nichts geschrieben.', Zeile 4000). Der Loeschpfad fuer *.backup ist zwar gegen Datenverlust abgesichert (Zeile 1667 prueft, ob das Original existiert und lesbar ist), *.tmp_new wird jedoch ungeprueft entfernt.

*Auswirkung:* Ein als ungefaehrlich angekuendigter Probelauf veraendert die Ablage. In Verbindung mit dem uebersprungenen Einzelinstanz-Schutz (Zeile 4861) kann der Probelauf ausserdem parallel zu einem echten Lauf Reste loeschen, die dort gerade entstehen.

*Vorschlag:* Die Aufraeum-Entscheidung um 'and not args.dry_run' erweitern (bzw. im Probelauf nur anzeigen, was geloescht wuerde).


## `5_OCR_PDF.py` Zeile 5202  `[BEHOBEN]`

**Probelauf ueberschreibt die Run-Summary des letzten echten Laufs und meldet dort 'Status: OK'**

```
write_run_summary(
            stats_result, target_dir, output_type,
            start_time, end_time,
            aborted=run_aborted, abort_reason=abort_reason,
        )
```

*Begruendung:* Der Abschlussblock unterscheidet nicht zwischen Probelauf und echtem Lauf; stats_result stammt bei --dry-run aus dry_run_directory. write_run_summary schreibt in die feste Datei RUN_SUMMARY_FILE (Zeile 420, '5_OCR_PDF_last_run.txt') im Modus 'w' (Zeile 3857) und vermerkt bei aborted=False 'Status: OK' (Zeile 3868). Ein Hinweis auf den Probelauf fehlt in der Datei vollstaendig.

*Auswirkung:* Der Nachweis ueber den letzten echten Verarbeitungslauf geht verloren und wird durch eine Datei ersetzt, die wie ein regulaerer Lauf mit 0 verarbeiteten und N uebersprungenen Dateien aussieht. Wer nach einem Nachtlauf die Zusammenfassung prueft, bekommt ein falsches Bild.

*Vorschlag:* Im Probelauf entweder keine Run-Summary schreiben oder in eine eigene Datei (z.B. '..._last_dryrun.txt') und die Kopfzeile mit 'PROBELAUF' kennzeichnen.


## `8_verschieben_auf_Google_Drive.ps1` Zeile 2098  `[BEHOBEN]`

**Die Zwischendatei '<log>.log.tmp' der Log-Substitution bleibt bei Fehler oder hartem Abbruch liegen und wird von keinem Aufraeumpfad erfasst.**

```
$tmpFile    = "$LogFile.tmp"
```

*Begruendung:* Die Log-Pfad-Substitution schreibt zuerst nach '$LogFile.tmp' und ersetzt das Original erst, wenn $logSuccess gesetzt ist (Zeile 2125-2131). Faengt der catch in Zeile 2119 eine Ausnahme ab (gesperrte Logdatei, volle Platte) oder bricht der Lauf zwischen Zeile 2104 und 2127 hart ab, bleibt die .tmp-Datei stehen. Clear-TempLeftovers (Zeile 217-245) raeumt ausschliesslich Verzeichnisse mit den Praefixen '8_verschieben_auf_Google_Drive_empty_' und 'RoboMirrorEmpty_' auf und kennt dieses Muster nicht. Dieselbe Stelle existiert ein zweites Mal im rclone-Zweig (Zeile 2644).

*Auswirkung:* Robocopy-Lauf, bei dem die Logdatei waehrend der Substitution durch einen Viewer gesperrt ist: es erscheint 'Log konnte nicht formatiert werden', und im Protokollverzeichnis bleibt dauerhaft '8_verschieben_auf_Google_Drive_<Zeitstempel>_Robocopy.log.tmp' neben dem eigentlichen Log liegen. Bei woechentlichen Laeufen sammeln sich unvollstaendige Teilkopien des Migrationsprotokolls an, die spaeter nicht mehr von der gueltigen Fassung zu unterscheiden sind.

*Vorschlag:* Die .tmp-Datei im finally-Block loeschen, wenn $logSuccess nicht gesetzt ist, und das Praefix-Muster '*.log.tmp' zusaetzlich in Clear-TempLeftovers aufnehmen (dort dann auch Dateien, nicht nur -Directory, beruecksichtigen).



# WIDERLEGT (keine Korrektur noetig)

- `3c_ppt_pptx_auf_neueste_Version_aktualisieren.py:1595` — robust_copy akzeptiert eine abgeschnittene Kopie als Erfolg - es wird nur 'groesser als 100 Byte' geprueft, ke
  - Widerlegt. Die Codebeschreibung stimmt zwar (verify_file, Zeile 1533-1540, prueft nur Existenz und min_size=100, waehrend robust_move in Zeile 1633/1638 die Quellgroesse vergleicht), der behauptete Fehler ist aber nicht erreichbar. Damit robust_copy eine abgeschnittene Kopie als 
- `4a_ersetze_font_in_word.py:2263` — Das 'doc.Saved = True' vor doc.Convert() loescht das Dirty-Flag von AcceptAllRevisions(); die Datei kann danac
  - WIDERLEGT durch Nachstellung mit Word 16 per COM (Testskript im Scratchpad, kein Zugriff auf Nutzdaten). Die Beweiskette des Pruefers haengt an dem Satz 'RemoveDocumentInformation(2) ist nach AcceptAllRevisions ein No-Op'. Genau das ist falsch: gemessen an einem frisch angelegten
- `4a_ersetze_font_in_word.py:182` — Die Tabelle METADATA_TYPES weicht von der Word-Aufzaehlung WdRemoveDocInfoType ab: 18/19/20 sind dort keine gu
  - WIDERLEGT durch Auslesen der Office-Typbibliothek, die der Pruefer selbst als Gegenprobe verlangt hat. Die Word-16-Typbibliothek liegt auf diesem Rechner vor (HKLM\SOFTWARE\Classes\TypeLib\{00020905-0000-0000-C000-000000000046}\8.7 -> C:\Program Files\Microsoft Office\Root\Office
- `4c_ersetze_font_in_powerpoint.py:2204` — Fehlgeschlagene Schema-Migration (.ppt/.pps/.pot) wird nur als Detail-Log-Warnung abgelegt; die Datei zaehlt d
  - Der beschriebene Ablauf stimmt technisch (was_converted bleibt False, Z. 2401 faellt in den else-Zweig, weder is_temp_copy noch temp_stage2_path greifen, Z. 2521-2525 meldet OK/append_done/SUCCESS), aber ein Kommentar erklaert das Verhalten als beabsichtigt: Z. 2172-2175 dokument
