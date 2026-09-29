# ==============================================================================
# Skript:  Office Macro Deep-Converter (Network Stable Edition)
# Stand:   11.06.2026
# Pfad:    Scannt rekursiv das gewaehlte Verzeichnis auf .docm, .xlsm, .pptm
# Logik:   Konvertiert NUR, wenn 0 Zeilen VBA-Code enthalten sind.
# Methode: Dateien werden lokal nach %LOCALAPPDATA%\Dateimigration-Arbeitskopien kopiert, dort
#          verarbeitet und anschliessend zurueck auf das Netzlaufwerk geschrieben.
#
# Excel 4.0 Makros (XLM):
#   HasVBProject / VBComponents erkennen nur Standard-VBA-Makros. Sehr alte
#   Excel 4.0-Makros (XLM) liegen in eigenen Arbeitsblaettern und werden vom
#   Office-Objektmodell NICHT als VBProject erfasst - solche Dateien galten
#   frueher faelschlich als "makrofrei" und verloren beim Umwandeln nach
#   .xlsx ihren Makroinhalt.
#   Test-HasExcel4Macro prueft daher zusaetzlich auf ZIP-Ebene, ob das
#   OOXML-Paket 'xl/macrosheets/' bzw. den Makroblatt-Content-Type enthaelt.
#   Treffer werden beibehalten und in der CSV zur manuellen Pruefung vermerkt.
#
# UMWANDLUNG IN .EXE (PS2EXE):
#   Install-Module ps2exe -Scope CurrentUser
#   Invoke-ps2exe -inputFile '.\7_Dateien_ohne_Makro_finden.ps1' -outputFile '.\7_Dateien_ohne_Makro_finden.exe' -iconFile '.\powershell_icon.ico' -title 'Office Macro Deep-Converter' -description 'Identifiziert .docm/.xlsm/.pptm ohne Makro-Inhalt und konvertiert sie zu .docx/.xlsx/.pptx' -product 'Office Macro Deep-Converter' -version '1.0.0.0' -STA -x64 -supportOS
#
#   Hinweis: -noConsole NICHT setzen (Read-Host / ReadKey / ESC-Abbruch
#            benoetigen ein interaktives Konsolenfenster).
#
#   Hinweis zu -x64 / -x86 bei gemischten Office-Bitness-Umgebungen:
#     Word, Excel und PowerPoint registrieren ihre COM-Server jeweils als
#     LocalServer32 (out-of-process: WINWORD.EXE / EXCEL.EXE / POWERPNT.EXE).
#     Das Windows-COM-Subsystem marshallt Aufrufe zwischen x64-Aufrufer und
#     x86-Server (und umgekehrt) automatisch ueber DCOM/LRPC. Eine mit -x64
#     kompilierte EXE funktioniert daher auch mit 32-Bit-Office (und eine
#     -x86-EXE mit 64-Bit-Office). Bitness-Match liefert minimal bessere
#     Performance, ist aber nicht erforderlich. -x64 ist eine sichere
#     Default-Wahl, da modernes Office (2019+, M365) standardmaessig x64 ist.
# ==============================================================================

[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$Automated,
    [string]$RootPath = "",

    # Nur pruefen, nichts umwandeln (identisch zu -WhatIf, aber sprechender).
    [switch]$ReadOnlyMode,

    # Ueberspringt das Aufraeumen zurueckgebliebener .bak_-Dateien.
    [switch]$SkipBackupCleanup
)

# ==================================================================
# Ausfuehrungsumgebung pruefen
# ==================================================================
# Dieses Skript setzt Windows PowerShell 5.1 voraus. Unter PowerShell 7
# (Edition 'Core') fehlen die Methoden FileInfo.GetAccessControl und
# .SetAccessControl - sie existieren nur im .NET Framework und wurden in
# .NET Core entfernt. Nachgestellt auf diesem Rechner: unter 5.1.26100.9168
# vorhanden, unter 7.6.4 nicht. Die Uebernahme von Rechten und Eigentuemer
# bei der Dateiersetzung faellt dort still aus (der Fehler landete nur als
# DEBUG im Detail-Log), und auf einer Ablage mit Owner-Mapping kann der
# Fachnutzer damit den Zugriff auf seine eigene Datei verlieren.
# Ein stiller Rechteverlust ist schlimmer als ein klarer Abbruch.
if ($PSVersionTable.PSEdition -eq 'Core') {
    Write-Host ""
    Write-Host ("=" * 70) -ForegroundColor Red
    Write-Host "  FALSCHE POWERSHELL-EDITION" -ForegroundColor Red
    Write-Host ("=" * 70) -ForegroundColor Red
    Write-Host "  Laeuft unter: PowerShell $($PSVersionTable.PSVersion) (Edition Core)"
    Write-Host "  Benoetigt   : Windows PowerShell 5.1 (Edition Desktop)"
    Write-Host ""
    Write-Host "  Grund: Unter PowerShell 7 lassen sich NTFS-Rechte und Eigentuemer"
    Write-Host "  der bearbeiteten Dateien nicht uebernehmen. Der Lauf wuerde die"
    Write-Host "  Berechtigungen Ihrer Ablage still veraendern."
    Write-Host ""
    Write-Host "  Bitte mit 'powershell.exe' starten, nicht mit 'pwsh'." -ForegroundColor Yellow
    Write-Host ("=" * 70) -ForegroundColor Red
    Write-Host ""
    exit 2
}


# $PSCmdlet ist nur im Skript-Scope verfuegbar, nicht in Funktionen.
$script:ScriptCmdlet = $PSCmdlet
$script:PreviewOnly  = $ReadOnlyMode.IsPresent

function Get-RunningOfficeSessions {
    <#
        Liefert Office-Prozesse mit sichtbarem Hauptfenster, also echte
        Sitzungen des Anwenders - im Unterschied zu unsichtbaren
        Automatisierungs-Instanzen.
    #>
    param([string[]]$ProcessNames = @('WINWORD', 'EXCEL', 'POWERPNT'))
    $found = @()
    foreach ($n in $ProcessNames) {
        try {
            $found += @(Get-Process -Name $n -ErrorAction SilentlyContinue |
                        Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero })
        } catch { }
    }
    return ,@($found)
}

function Show-OfficeRunningWarning {
    <#
        Weist vor dem ersten COM-Zugriff auf bereits laufende Sitzungen hin.

        Word, Excel und PowerPoint wird per COM automatisiert. Laeuft bereits eine Sitzung des
        Anwenders, hat das zwei Auswirkungen:
          1. Die Automatisierung kann sich an die vorhandene Sitzung haengen.
             Warnhinweise werden dann abgeschaltet und das Fenster
             ausgeblendet - fuer den Anwender sieht das aus wie ein Absturz.
          2. Die eigene Instanz laesst sich beim Aufraeumen nicht mehr
             zuverlaessig von der fremden unterscheiden.

        Die Sitzung wird NICHT beendet - dafuer sorgt die PID-Bindung an
        den Quit-Stellen. Ein sauberer Lauf setzt aber ein geschlossenes
        Office voraus.

        Rueckgabe: $true = fortfahren, $false = Anwender bricht ab.
    #>
    param([switch]$Silent)

    $sessions = Get-RunningOfficeSessions
    if ($sessions.Count -eq 0) { return $true }

    $namen = ($sessions | ForEach-Object { $_.ProcessName } | Select-Object -Unique) -join ', '
    $pids  = ($sessions | ForEach-Object { $_.Id }) -join ', '

    if ($Silent) {
        Write-Warning "Laufende Office-Sitzungen erkannt ($namen, PID: $pids) - sie werden geschuetzt, aber nicht geschlossen."
        return $true
    }

    Write-Host ""
    Write-Host ("=" * 66) -ForegroundColor Yellow
    Write-Host "  WARNUNG: Office laeuft bereits" -ForegroundColor Yellow
    Write-Host ("=" * 66) -ForegroundColor Yellow
    Write-Host "  Gefundene Sitzungen: $namen (PID: $pids)"
    Write-Host ""
    Write-Host "  Ihre Sitzung wird vom Skript NICHT beendet. Waehrend des Laufs"
    Write-Host "  kann sie aber ausgeblendet werden und Warnhinweise sind"
    Write-Host "  abgeschaltet - das wirkt wie ein Absturz."
    Write-Host "  Ausserdem laesst sich die eigene Automatisierungs-Instanz dann"
    Write-Host "  nicht mehr zuverlaessig von Ihrer Sitzung unterscheiden."
    Write-Host ""
    Write-Host "  EMPFEHLUNG: Office jetzt schliessen und das Skript neu starten." -ForegroundColor Yellow
    Write-Host ("=" * 66) -ForegroundColor Yellow
    Write-Host ""
    # Ohne echte Konsole NICHT fragen. Read-Host blockiert dort unbegrenzt
    # (Aufgabenplanung mit angehaengter Konsole) oder liefert sofort leer -
    # beides taugt nicht als Freigabe. Massgeblich ist die tatsaechliche
    # Eingabefaehigkeit, nicht der Hostname (ps2exe meldet 'PSRunspace-Host').
    $kannFragen = $false
    try { $kannFragen = -not [Console]::IsInputRedirected } catch { $kannFragen = $false }
    if (-not $kannFragen) {
        Write-Warning "Keine interaktive Konsole - Abbruch zum Schutz laufender Office-Sitzungen. Fuer den unbeaufsichtigten Betrieb -NoInteractive bzw. -Automated verwenden."
        return $false
    }

    $answer = Read-Host "Trotzdem fortfahren? [j/N]"
    if ($answer -notmatch '^[JjYy]') {
        Write-Host "Abgebrochen. Bitte Office schliessen und neu starten." -ForegroundColor Cyan
        return $false
    }
    return $true
}

function Confirm-Write {
    <#
        Zentrale Freigabe fuer jede schreibende Operation am Original.
        Vorher kannte das Skript keinen Probelauf - ein Testlauf auf einer
        fremden Ablage war nicht moeglich.
    #>
    param(
        [string]$Target,
        [string]$Action = 'Nach makrofreiem Format umwandeln'
    )
    if ($script:PreviewOnly) { return $false }
    if ($null -eq $script:ScriptCmdlet) { return $true }
    try {
        return $script:ScriptCmdlet.ShouldProcess($Target, $Action)
    } catch {
        return $true
    }
}

# Verzeichnisse, die bei der Suche nicht betreten werden.
$script:ExcludeDirNames = @(
    '$RECYCLE.BIN', 'System Volume Information', 'RECYCLER',
    '.git', '.svn', '__pycache__', 'node_modules'
)

# Verzeichnis-Presets fuer die Startauswahl.
$script:DirectoryPresets = @(
    'Q:\'
    'R:\'
    'G:\Geteilte Ablagen\'
    'G:\Meine Ablage\'
    '\\server\dfs\'
)

# Mindestalter, ab dem ein zurueckgebliebenes .bak_ als verwaist gilt.
$script:BackupCleanupMinAgeHours = 24

# Von Get-FilesRecursiveSafe gemeldete Pfade jenseits der Framework-Grenze.
# Vorher gab es den Statistikposten "Pfad zu lang", der nie hochgezaehlt wurde.
$script:SkippedPathTooLong = 0

# Wird nach der Log-Initialisierung gesetzt; leer = kein CSV-Protokoll.
$script:SkipLogPath = ''



# ==================================================================
# Gemeinsame Grundbibliothek (mit Rueckfall)
# ==================================================================
# Bindet _gemeinsam.psm1 ein, wenn vorhanden. Die eingebauten Kopien der
# Helfer bleiben bestehen und ueberschreiben das Modul absichtlich - so
# bleibt jedes Skript einzeln lauffaehig und die ps2exe-Uebersetzung
# funktioniert unveraendert. Genutzt wird das Modul fuer das gemeinsame
# Laufprotokoll (migration.jsonl) und die zentralen Verzeichnis-Presets
# aus pfade.json.
$script:GemeinsamGeladen = $false
try {
    $gemModul = Join-Path $PSScriptRoot '_gemeinsam.psm1'
    if (Test-Path -LiteralPath $gemModul) {
        Import-Module $gemModul -Force -DisableNameChecking -ErrorAction Stop
        $script:GemeinsamGeladen = $true
    }
} catch {
    # Ohne Modul laeuft das Skript mit seinen eingebauten Helfern weiter.
}

# ==============================================================================
# Grundkonfiguration
# ==============================================================================
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

# Fuer Test-HasExcel4Macro (ZIP-Zugriff auf das OOXML-Paket).
try { Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop } catch {
    try { [System.Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") | Out-Null } catch {}
}
# ZipArchive selbst liegt in System.IO.Compression; das laedt .NET erst mit
# dem ersten ZipFile-Aufruf nach. Test-HasExcel4Macro baut das Archiv aber
# direkt auf einem FileStream auf - ohne diese Zeile ist der Typ unbekannt
# und jede Pruefung endet als "nicht pruefbar" (gemessen unter 5.1).
try { Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop } catch {}

$rootPath = $RootPath
# Logs liegen direkt neben Skript/EXE. PSScriptRoot greift bei direktem
# Aufruf; bei ps2exe-EXEs ist PSScriptRoot leer - dann liefert die
# Entry-Assembly das EXE-Verzeichnis. $PWD nur als letzter Fallback.
function Resolve-ScriptDirectory {
    if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        return $PSScriptRoot
    }
    if ($MyInvocation.MyCommand.Path) {
        return Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    try {
        $entry = [System.Reflection.Assembly]::GetEntryAssembly()
        if ($entry -and $entry.Location) {
            return Split-Path -Parent $entry.Location
        }
    } catch {}
    return $PWD.Path
}
$scriptDir = Resolve-ScriptDirectory
$timestamp = (Get-Date).ToString("yyyy-MM-dd_HHmmss")
$logFile  = Join-Path -Path $scriptDir -ChildPath "7_Dateien_ohne_Makro_finden_$timestamp.log"
$skipLog  = Join-Path -Path $scriptDir -ChildPath "7_Dateien_ohne_Makro_finden_ManuellePruefung_$timestamp.csv"
# Gemeinsamer Arbeitsordner der Office-Skripte unter %LOCALAPPDATA%. Bis
# 29.09.2026 "Dokumente" - bei OneDrive-Ordnersicherung wanderte jede
# Arbeitskopie in die Cloud; %LOCALAPPDATA% wird nie umgeleitet.
$arbeitsBasis = Join-Path $(if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { $env:TEMP }) 'Dateimigration-Arbeitskopien'
$tempDir  = Join-Path -Path $arbeitsBasis -ChildPath "7_Dateien_ohne_Makro_finden_$PID"
$utf8Bom  = New-Object System.Text.UTF8Encoding $true
$restartThreshold = 500
# Frist je Datei fuer Oeffnen, VBA-Pruefung und Speichern (siehe
# Start-DateiWaechter). Danach wird die eigene Office-Instanz beendet.
$script:DateiTimeoutSec = 180
$bannerDate      = Get-Date -Format "dd.MM.yyyy"
$officePids       = @{ Word = $null; Excel = $null; PowerPoint = $null }
$officeStartTimes = @{ Word = $null; Excel = $null; PowerPoint = $null }
$cancelled        = $false

# Alte Temp-Ordner (>24h) frueherer/abgestuerzter Laeufe entfernen -
# Filter passt zum aktuellen Skriptpraefix.
$staleTempThreshold = (Get-Date).AddHours(-24)
@($arbeitsBasis, [Environment]::GetFolderPath("MyDocuments")) | Where-Object { $_ -and (Test-Path -LiteralPath $_) } |
    ForEach-Object { Get-ChildItem -LiteralPath $_ -Directory -Filter "7_Dateien_ohne_Makro_finden_*" -ErrorAction SilentlyContinue } |
    Where-Object { $_.LastWriteTime -lt $staleTempThreshold } |
    ForEach-Object {
        try { Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop -WhatIf:$false } catch {}
    }

if (-not (Test-Path -LiteralPath $tempDir)) { New-Item -ItemType Directory -Path $tempDir -Force -WhatIf:$false | Out-Null }

# ==============================================================================
# Hilfsfunktionen - Anzeige
# ==============================================================================
function Write-CenteredLine {
    param([string]$Text, [int]$Width = 78, [string]$Color = "Cyan")
    $padding = [Math]::Max(0, $Width - $Text.Length)
    $left  = [Math]::Floor($padding / 2)
    $right = $padding - $left
    Write-Host ("║" + (" " * $left) + $Text + (" " * $right) + "║") -ForegroundColor $Color
}

function Show-Banner {
    $boxWidth = 78
    Write-Host ""
    Write-Host ("╔" + ("═" * $boxWidth) + "╗") -ForegroundColor Cyan
    Write-CenteredLine ""
    Write-CenteredLine "IDENTIFIKATION FALSCHER DATEIENDUNGEN"
    Write-CenteredLine "(als .docm, .xlsm, .pptm gespeicherte Dateien ohne Makroinhalt)"
    Write-CenteredLine ""
    Write-CenteredLine ("Stand: $bannerDate")
    Write-CenteredLine ""
    Write-Host ("╚" + ("═" * $boxWidth) + "╝") -ForegroundColor Cyan
    Write-Host ""
}

function Write-Log {
    param ([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry  = "$timestamp [$Level] $Message"
    $color = switch($Level) {
        "SUCCESS" { "Green" }
        "WARN"    { "Yellow" }
        "ERROR"   { "Red" }
        default   { "Gray" }
    }
    Write-Host $logEntry -ForegroundColor $color
    try {
        [System.IO.File]::AppendAllText($logFile, $logEntry + "`r`n", $utf8Bom)
    } catch {}
}

function Write-SkipLogEntry {
    <#
        Schreibt eine Zeile in die CSV-Datei zur manuellen Nachpruefung.
        Zentral, weil die Zeilen vorher an zwei Stellen von Hand
        zusammengesetzt wurden. Eine davon nutzte eine falsch geschriebene
        Umbruch-Escapesequenz und schrieb daher ein literales Backslash-r
        Backslash-n in die Datei, statt die Zeile umzubrechen - alle
        Datensaetze landeten in einer einzigen CSV-Zeile. Zusaetzlich fehlte
        die CSV-Maskierung: Pfade duerfen Semikolon und Anfuehrungszeichen
        enthalten, was in Excel die Spalten verschoben hat.
    #>
    param(
        [string]$Source,
        [string]$Target = '',
        [string]$Reason = ''
    )

    # Gemeinsames Laufprotokoll (migration.jsonl) - ergaenzt das
    # skripteigene Protokoll, ersetzt es nicht. Erst damit laesst sich
    # der Fortschritt ueber alle elf Schritte hinweg auswerten.
    if ($script:GemeinsamGeladen) {
        try {
            Write-Laufprotokoll -Skript '7_Dateien_ohne_Makro_finden' `
                -Pfad $Source -Aktion 'Makropruefung' `
                -Status 'PRUEFEN' -Detail $Reason
        } catch { }
    }
    if ([string]::IsNullOrWhiteSpace($script:SkipLogPath)) { return }
    $q = {
        param($v)
        if ($null -eq $v) { $v = '' }
        # Zeilenumbrueche im Feld wuerden den Datensatz zerreissen.
        $v = ($v -replace '\r?\n', ' ')
        if ($v -match '[";]') { return '"' + ($v -replace '"', '""') + '"' }
        return $v
    }
    $line = ((& $q (Remove-LongPathPrefix $Source)),
             (& $q (Remove-LongPathPrefix $Target)),
             (& $q $Reason)) -join ';'
    try {
        [System.IO.File]::AppendAllText($script:SkipLogPath, $line + "`r`n", $utf8Bom)
    } catch { }
}

# ==============================================================================
# Hilfsfunktionen - Pfade
# ==============================================================================
function Remove-LongPathPrefix {
    param ([string]$Path)
    # StartsWith statt -like '\\?\*': bei -like ist '?' ein Platzhalter,
    # ein UNC-Pfad mit einbuchstabigem Server ('\\s\share\...') galt dann
    # als bereits praefixiert (gemessen unter 5.1). Ebenso unten und in 9.
    if ($Path.StartsWith('\\?\UNC\', [System.StringComparison]::OrdinalIgnoreCase)) { return '\\' + $Path.Substring(8) }
    if ($Path.StartsWith('\\?\'))     { return $Path.Substring(4) }
    return $Path
}

function Add-LongPathPrefix {
    param ([string]$Path)
    if ($Path.StartsWith('\\?\')) { return $Path }
    if ($Path -like '\\*')   { return '\\?\UNC\' + $Path.TrimStart('\') }
    if ($Path -match '^[A-Za-z]:\\') { return '\\?\' + $Path }
    return $Path
}

function Get-UniqueTargetPath {
    param ([string]$BasePath)
    # [System.IO.File]::Exists statt Test-Path: Test-Path liefert bei
    # '\\?\'-Praefixen unter PowerShell 5.1 nicht zuverlaessig $true.
    # Ein falsches $false haette hier die schlimmste Folge - die Funktion
    # meldete dann "Ziel frei" und der Aufrufer ueberschriebe eine
    # bestehende Datei.
    $baseLong = Add-LongPathPrefix $BasePath
    if (-not ([System.IO.File]::Exists($baseLong) -or [System.IO.Directory]::Exists($baseLong))) {
        return $BasePath
    }
    $dir  = Split-Path -Parent $BasePath
    $name = [System.IO.Path]::GetFileNameWithoutExtension($BasePath)
    $ext  = [System.IO.Path]::GetExtension($BasePath)
    $i = 2
    while ($true) {
        $candidate     = Join-Path $dir ($name + "_" + $i + $ext)
        $candidateLong = Add-LongPathPrefix $candidate
        if (-not ([System.IO.File]::Exists($candidateLong) -or [System.IO.Directory]::Exists($candidateLong))) {
            return $candidate
        }
        $i++
    }
}

function Get-UniqueBackupPath {
    param ([string]$OriginalPath)
    return $OriginalPath + ".bak_" + [Guid]::NewGuid().ToString("N").Substring(0, 8)
}

function Test-HasExcel4Macro {
    <#
        Prueft auf ZIP-Ebene, ob eine .xlsm Excel-4.0-Makros (XLM) enthaelt.

        Warum noetig: XLM-Makros liegen in eigenen Makroblaettern, nicht im
        VBA-Projekt. HasVBProject liefert dafuer $false und VBComponents ist
        leer - die Datei galt damit als "makrofrei" und wurde nach .xlsx
        umgewandelt. Dabei gehen die Makroblaetter ersatzlos verloren, ohne
        dass Excel oder das Skript warnen.

        Erkannt wird beides:
          - der Ordner 'xl/macrosheets/' im Paket
          - der Content-Type 'ms-excel.macrosheet' in [Content_Types].xml

        Rueckgabe: $true (XLM vorhanden), $false (keine), $null (nicht
        pruefbar - dann konservativ wie $true behandeln).
    #>
    param([string]$FilePath)

    $zip = $null
    $fs  = $null
    try {
        # FileShare.ReadWrite statt ZipFile.OpenRead: die Pruefung laeuft,
        # waehrend Excel die Mappe zum Schreiben offen haelt. OpenRead teilt
        # nur Lesezugriff und scheiterte daran bei JEDER .xlsm - alle landeten
        # als "XLM-Pruefung nicht moeglich" im Behalten-Zweig (nachgestellt).
        $fs  = New-Object System.IO.FileStream($FilePath, [System.IO.FileMode]::Open,
                   [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)

        foreach ($entry in $zip.Entries) {
            if ($entry.FullName -like 'xl/macrosheets/*') { return $true }
        }

        $ct = $zip.Entries | Where-Object { $_.FullName -eq '[Content_Types].xml' } | Select-Object -First 1
        if ($ct) {
            $sr = New-Object System.IO.StreamReader($ct.Open())
            try {
                $inhalt = $sr.ReadToEnd()
            } finally {
                $sr.Close()
            }
            if ($inhalt -match 'ms-excel\.macrosheet') { return $true }
        }

        # XLM-Funktionen in definierten Namen (GET.CELL, GET.WORKBOOK,
        # EVALUATE ...). Sie stehen nicht auf einem Makroblatt, sondern in
        # xl/workbook.xml, und gehen beim Speichern als .xlsx genauso
        # verloren (Zellen darauf zeigen #NAME?). 3b und 4b sichern solche
        # Mappen deshalb als .xlsm - ohne diese Pruefung machte 7 sie danach
        # wieder zu .xlsx. In der Datei steht die Formel immer in der
        # englischen Makrosprache; Namen mit function/vbProcedure/xlm="1"
        # sind selbst XLM-Makros. Musterauswahl wie in 3b/4b (_XLM_IN_NAMEN_RE).
        $wbEntry = $zip.Entries | Where-Object { $_.FullName -eq 'xl/workbook.xml' } | Select-Object -First 1
        if ($wbEntry) {
            $sr = New-Object System.IO.StreamReader($wbEntry.Open())
            try { $wbXml = $sr.ReadToEnd() } finally { $sr.Close() }
            $xlmMuster = '(?i)(?<![A-Za-z0-9_.])(?:GET\.[A-Za-z0-9_.]+|EVALUATE|FILES|DOCUMENTS|DIRECTORY|ACTIVE\.CELL|CALLER|SELECTION|NAMES|LINKS|WINDOWS|REFTEXT|TEXTREF|ABSREF|RELREF|DEREF|CALL|REGISTER\.ID)\('
            foreach ($m in [regex]::Matches($wbXml, '(?s)<definedName\b([^>]*)>(.*?)</definedName>')) {
                if ($m.Groups[1].Value -match '\b(function|vbProcedure|xlm)="(1|true)"') { return $true }
                if ($m.Groups[2].Value -match $xlmMuster) { return $true }
            }
        }
        return $false
    } catch {
        return $null
    } finally {
        if ($null -ne $zip) { try { $zip.Dispose() } catch {} }
        if ($null -ne $fs)  { try { $fs.Dispose() }  catch {} }
    }
}

function Get-OoxmlKennwortGrund {
    <#
        '' wenn Office die Datei ohne Kennwortabfrage oeffnet, sonst der
        Grund ('verschluesselt' / 'Schreibkennwort').

        Gemessen am 29.09.2026 (Office 2024): eine verschluesselte .docm und
        .xlsm liefen je bis zum Zeitwaechter (180 s) und kosteten einen
        Office-Neustart. Erkannt wird:
          - verschluesselt: OOXML-Datei im CFB-Container statt ZIP
          - Schreibkennwort: w:writeProtection mit Hash (Word),
            fileSharing mit Kennwort (Excel), p:modifyVerifier (PowerPoint)
    #>
    param([string]$Path, [string]$Ext)

    $fs = $null; $zip = $null
    try {
        $fs = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Open,
                  [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $kopf = New-Object byte[] 4
        if ($fs.Read($kopf, 0, 4) -lt 4) { return '' }
        if ($kopf[0] -eq 0xD0 -and $kopf[1] -eq 0xCF -and $kopf[2] -eq 0x11 -and $kopf[3] -eq 0xE0) {
            return 'verschluesselt'
        }
        if ($kopf[0] -ne 0x50 -or $kopf[1] -ne 0x4B) { return '' }
        [void]$fs.Seek(0, [System.IO.SeekOrigin]::Begin)
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
        $teil = switch ($Ext) {
            '.docm' { 'word/settings.xml' }
            '.xlsm' { 'xl/workbook.xml' }
            '.pptm' { 'ppt/presentation.xml' }
            default { $null }
        }
        if (-not $teil) { return '' }
        $eintrag = $zip.Entries | Where-Object { $_.FullName -eq $teil } | Select-Object -First 1
        if (-not $eintrag) { return '' }
        $sr = New-Object System.IO.StreamReader($eintrag.Open())
        try { $xml = $sr.ReadToEnd() } finally { $sr.Close() }
        switch ($Ext) {
            '.docm' {
                $m = [regex]::Match($xml, '<w:writeProtection\b[^>]*>')
                if ($m.Success -and $m.Value -match 'w:(hashValue|hash|cryptProviderType|salt)=') { return 'Schreibkennwort' }
            }
            '.xlsm' {
                $m = [regex]::Match($xml, '<(?:\w+:)?fileSharing\b[^>]*>')
                if ($m.Success -and $m.Value -match '\b(reservationPassword|hashValue|algorithmName)=') { return 'Schreibkennwort' }
            }
            '.pptm' {
                if ($xml -match '<(?:\w+:)?modifyVerifier\b') { return 'Schreibkennwort' }
            }
        }
        return ''
    } catch {
        return ''
    } finally {
        if ($null -ne $zip) { try { $zip.Dispose() } catch {} }
        if ($null -ne $fs)  { try { $fs.Dispose() }  catch {} }
    }
}

function Remove-OoxmlSchreibkennwort {
    <#
        Entfernt das Schreib-/Aenderungskennwort aus der LOKALEN Kopie
        (Entscheidung des Anwenders vom 30.09.2026: entfernen statt
        ueberspringen). Office selbst kann solche Dateien nur
        schreibgeschuetzt oeffnen, PowerPoint sie dann gar nicht speichern
        (gemessen) - deshalb auf Dateiebene, wie 2a-2c:
          .docm  w:writeProtection  (word/settings.xml)
          .xlsm  fileSharing        (xl/workbook.xml)
          .pptm  p:modifyVerifier   (ppt/presentation.xml)
        Wandelt 7 die Datei um, entsteht sie ohne Kennwort; behaelt 7 sie
        (Makros), bleibt das Original wie immer unangetastet.
    #>
    param([string]$Path, [string]$Ext)
    $teil = $null; $muster = $null
    switch ($Ext) {
        '.docm' { $teil = 'word/settings.xml';    $muster = '<w:writeProtection\b[^>]*/>|<w:writeProtection\b[^>]*>.*?</w:writeProtection>' }
        '.xlsm' { $teil = 'xl/workbook.xml';      $muster = '<(?:\w+:)?fileSharing\b[^>]*/>|<(?:\w+:)?fileSharing\b[^>]*>.*?</(?:\w+:)?fileSharing>' }
        '.pptm' { $teil = 'ppt/presentation.xml'; $muster = '<(?:\w+:)?modifyVerifier\b[^>]*/>|<(?:\w+:)?modifyVerifier\b[^>]*>.*?</(?:\w+:)?modifyVerifier>' }
    }
    if (-not $teil) { return }
    $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $eintrag = $zip.GetEntry($teil)
        if (-not $eintrag) { return }
        $sr = New-Object System.IO.StreamReader($eintrag.Open())
        try { $xml = $sr.ReadToEnd() } finally { $sr.Close() }
        $neu = [regex]::Replace($xml, $muster, '', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        if ($neu -eq $xml) { return }
        $eintrag.Delete()
        $ersatz = $zip.CreateEntry($teil)
        $sw = New-Object System.IO.StreamWriter($ersatz.Open(), (New-Object System.Text.UTF8Encoding $false))
        try { $sw.Write($neu) } finally { $sw.Close() }
    } finally {
        $zip.Dispose()
    }
}

function Get-UserShellFolder {
    param([string]$Name, [string]$Fallback)
    try {
        $key  = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
        $prop = Get-ItemProperty -Path $key -Name $Name -ErrorAction Stop
        $val  = $prop | Select-Object -ExpandProperty $Name -ErrorAction Stop
        if ($val) {
            $expanded = [Environment]::ExpandEnvironmentVariables($val)
            if (Test-Path -LiteralPath $expanded) { return $expanded }
        }
    } catch {}
    return $Fallback
}

# ==============================================================================
# Hilfsfunktionen - Datei-Locks und Retries
# ==============================================================================
function Wait-FileAvailable {
    param([string]$Path, [int]$MaxAttempts = 10, [int]$DelayMs = 300)
    for ($i = 1; $i -le $MaxAttempts; $i++) {
        try {
            $fs = [System.IO.File]::Open($Path, 'Open', 'Read', 'None')
            $fs.Close()
            $fs.Dispose()
            return $true
        } catch {
            Start-Sleep -Milliseconds $DelayMs
        }
    }
    return $false
}

function Invoke-WithRetry {
    param([scriptblock]$Action, [int]$MaxAttempts = 4, [int]$DelayMs = 500)
    for ($i = 1; $i -le $MaxAttempts; $i++) {
        try {
            & $Action
            return
        } catch {
            if ($i -eq $MaxAttempts) { throw }
            Start-Sleep -Milliseconds ($DelayMs * $i)
        }
    }
}

# Whitelist fuer den Windows-Temp-Cleanup: NUR Eintraege mit diesen
# Praefixen (Office/COM-Reste sowie eigene Skript-Artefakte) duerfen
# geloescht werden. NIEMALS pauschal leeren: Fremdprozesse legen
# aktive Daten ohne Lock in %TEMP% ab; blindes Loeschen zerstoert sie.
$script:WindowsTempWhitelistPrefixes = @(
    '~$', '~df', 'gen_py', 'vbe', 'excel8.0',
    '7_dateien_ohne_makro_finden'
)

function Test-WhitelistedTempEntry {
    param([string]$Name)
    $nl = $Name.ToLowerInvariant()
    foreach ($p in $script:WindowsTempWhitelistPrefixes) {
        if ($nl.StartsWith($p)) { return $true }
    }
    if ($nl.StartsWith('cvr') -and $nl.EndsWith('.tmp')) { return $true }
    return $false
}

function Invoke-WindowsTempCleanup {
    # Best-effort Bereinigung von %LOCALAPPDATA%\Temp am Skriptende -
    # ausschliesslich per Whitelist bekannter Praefixe.
    # Gesperrte/in-Nutzung-Dateien werden stillschweigend uebersprungen.
    $winTemp = $env:TEMP
    if (-not $winTemp) { $winTemp = $env:TMP }
    if (-not $winTemp -or -not (Test-Path -LiteralPath $winTemp)) { return }
    Get-ChildItem -LiteralPath $winTemp -Force -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            if (-not (Test-WhitelistedTempEntry $_.Name)) { return }
            if ($_.PSIsContainer) {
                Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false
            } else {
                Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue -WhatIf:$false
            }
        } catch {}
    }
}

# ==================================================================
# NTFS-SICHERHEITSUEBERNAHME (ACL/OWNER) BEI DATEI-ERSETZUNG
# ==================================================================
# Das Skript ersetzt Dateien bei der Konvertierung, statt sie in-place
# zu aendern. Neu erzeugte Dateien erben nur die vererbbaren ACEs des
# Zielordners und gehoeren dem AUSFUEHRENDEN Konto. Laeuft das Skript
# unter einem Admin-Konto ueber Ablagen mit expliziten Benutzer-ACEs
# oder NAS-Owner-Mapping, verliert der Nutzer dadurch den Zugriff.
# Daher: Owner/Group/DACL des Originals VOR der Ersetzung sichern und
# auf der neuen Datei wiederherstellen.
#  - Admin-Konto: SeRestore-/SeTakeOwnership-Privileg wird beim Start
#    aktiviert; Owner-Wiederherstellung funktioniert vollstaendig.
#  - Nutzer-Konto: Privileg-Aktivierung schlaegt still fehl; die neue
#    Datei gehoert ohnehin dem Nutzer. DACL-Restore funktioniert (der
#    Owner hat implizit WRITE_DAC); ein fremder Original-Owner wird
#    toleriert und nur im Detail-Log vermerkt.
# Alles best effort: ein fehlgeschlagener ACL-Restore bricht die
# Dateiverarbeitung nicht ab.

function Enable-RestorePrivileges {
    $script:RestorePrivilegesEnabled = $false
    try {
        if (-not ('DM.PrivilegeHelper' -as [type])) {
            Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace DM {
    public static class PrivilegeHelper {
        [StructLayout(LayoutKind.Sequential)]
        struct LUID { public uint LowPart; public int HighPart; }
        [StructLayout(LayoutKind.Sequential)]
        struct TOKEN_PRIVILEGES { public int PrivilegeCount; public LUID Luid; public int Attributes; }
        [DllImport("advapi32.dll", SetLastError = true)]
        static extern bool OpenProcessToken(IntPtr h, int acc, out IntPtr tok);
        [DllImport("advapi32.dll", SetLastError = true)]
        static extern bool LookupPrivilegeValue(string sys, string name, out LUID luid);
        [DllImport("advapi32.dll", SetLastError = true)]
        static extern bool AdjustTokenPrivileges(IntPtr tok, bool dis, ref TOKEN_PRIVILEGES newst, int len, IntPtr prev, IntPtr ret);
        [DllImport("kernel32.dll")]
        static extern IntPtr GetCurrentProcess();
        public static bool Enable(string priv) {
            IntPtr tok;
            if (!OpenProcessToken(GetCurrentProcess(), 0x28, out tok)) { return false; }
            LUID luid;
            if (!LookupPrivilegeValue(null, priv, out luid)) { return false; }
            TOKEN_PRIVILEGES tp;
            tp.PrivilegeCount = 1;
            tp.Luid = luid;
            tp.Attributes = 0x2;
            if (!AdjustTokenPrivileges(tok, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero)) { return false; }
            return Marshal.GetLastWin32Error() == 0;
        }
    }
}
'@
        }
        $okRestore = [DM.PrivilegeHelper]::Enable('SeRestorePrivilege')
        $okOwner   = [DM.PrivilegeHelper]::Enable('SeTakeOwnershipPrivilege')
        $script:RestorePrivilegesEnabled = ($okRestore -or $okOwner)
        if ($script:RestorePrivilegesEnabled) {
            Write-Log "Restore-Privilegien aktiviert (Admin-Kontext) - Owner-Wiederherstellung vollstaendig verfuegbar."
        } else {
            Write-Log "Restore-Privilegien nicht zugewiesen (Nutzer-Kontext) - Owner-Restore nur auf eigene Dateien moeglich."
        }
    } catch {
        Write-Log "Privileg-Aktivierung fehlgeschlagen: $_" "WARN"
    }
}

function Get-FileSecuritySnapshot {
    param([string]$Path)
    try {
        $fi  = [System.IO.FileInfo]::new($Path)
        $acl = $fi.GetAccessControl(
            [System.Security.AccessControl.AccessControlSections]::Access -bor
            [System.Security.AccessControl.AccessControlSections]::Owner  -bor
            [System.Security.AccessControl.AccessControlSections]::Group)
        return @{
            Dacl       = $acl.GetSecurityDescriptorSddlForm(
                [System.Security.AccessControl.AccessControlSections]::Access)
            OwnerGroup = $acl.GetSecurityDescriptorSddlForm(
                [System.Security.AccessControl.AccessControlSections]::Owner -bor
                [System.Security.AccessControl.AccessControlSections]::Group)
        }
    } catch {
        Write-Log "Sicherheitsinfo nicht lesbar ($Path): $_" "WARN"
        return $null
    }
}

function Set-FileSecuritySnapshot {
    param([string]$Path, $Snapshot)
    if ($null -eq $Snapshot) { return }
    try {
        $fsDacl = [System.Security.AccessControl.FileSecurity]::new()
        $fsDacl.SetSecurityDescriptorSddlForm(
            $Snapshot.Dacl,
            [System.Security.AccessControl.AccessControlSections]::Access)
        [System.IO.FileInfo]::new($Path).SetAccessControl($fsDacl)
    } catch {
        Write-Log "DACL-Wiederherstellung fehlgeschlagen ($Path): $_" "WARN"
    }
    try {
        $fsOwn = [System.Security.AccessControl.FileSecurity]::new()
        $fsOwn.SetSecurityDescriptorSddlForm(
            $Snapshot.OwnerGroup,
            [System.Security.AccessControl.AccessControlSections]::Owner -bor
            [System.Security.AccessControl.AccessControlSections]::Group)
        [System.IO.FileInfo]::new($Path).SetAccessControl($fsOwn)
    } catch {
        if ($script:RestorePrivilegesEnabled) {
            Write-Log "Owner-Wiederherstellung fehlgeschlagen ($Path): $_" "WARN"
        } else {
            # Nutzer-Kontext ohne Restore-Privileg: fremder Owner wird toleriert.
        }
    }
}

# ==============================================================================
# Hilfsfunktionen - COM / Office
# ==============================================================================
function Release-ComObject {
    param ($obj)
    if ($null -ne $obj) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null } catch {}
    }
}

function Stop-OfficeApp {
    <#
        Beendet eine Office-Instanz - aber nur die selbst gestartete.

        Word und PowerPoint sind Einzelinstanz-COM-Server: laeuft bereits
        eine Sitzung des Anwenders, liefert New-Object -ComObject GENAU
        DIESE Instanz statt einer neuen. Start-TrackedOfficeApp erkennt das
        daran, dass kein NEUER Prozess auftaucht, und liefert ProcessId
        $null. Wurde in diesem Fall trotzdem Quit() gerufen, schloss das
        Skript die Word-/PowerPoint-Sitzung des Anwenders - samt aller
        ungespeicherten Dokumente.

        Ohne eigene PID wird die Referenz daher nur freigegeben und das
        Fenster wieder sichtbar gemacht. Sollte es doch eine eigene Instanz
        gewesen sein, raeumt Stop-TrackedOfficeProcess sie spaeter ab.
    #>
    param ($app, [Nullable[int]]$OwnPid = $null)
    if ($null -ne $app) {
        if ($null -ne $OwnPid -and $OwnPid -gt 0) {
            try { $app.Quit() } catch {}
        } else {
            try { $app.Visible = $true } catch {}
        }
        Release-ComObject $app
    }
}

# Fehlertolerante rekursive Datei-Suche.
# .NET's EnumerateFiles(..., AllDirectories) ist ein lazy IEnumerable -
# UnauthorizedAccessException, PathTooLongException und Konsorten
# treten erst beim Iterieren (MoveNext) auf, nicht beim Aufruf. Auf
# Netzlaufwerken mit "System Volume Information", versteckten Papier-
# koerben oder per ACL geschuetzten Unterordnern reisst genau das die
# foreach-Schleife im Hauptloop auf - die Suche endet vorzeitig.
# Daher: Stack-basierte manuelle Rekursion, jedes Verzeichnis einzeln
# in try/catch, Zugriffsfehler werden still uebersprungen.
function Get-FilesRecursiveSafe {
    param(
        [string]$RootPath,
        [string]$Pattern
    )
    $result = New-Object System.Collections.Generic.List[string]
    $stack  = New-Object System.Collections.Generic.Stack[string]
    $stack.Push($RootPath)
    while ($stack.Count -gt 0) {
        $current = $stack.Pop()
        try {
            foreach ($f in [System.IO.Directory]::EnumerateFiles($current, $Pattern)) {
                $result.Add($f)
            }
        }
        catch [System.UnauthorizedAccessException]   { }
        catch [System.IO.PathTooLongException]       { $script:SkippedPathTooLong++ }
        catch [System.IO.DirectoryNotFoundException] { }
        catch [System.ArgumentException] {
            # Long-Path-Praefix wird vom Framework nicht unterstuetzt -
            # nur beim Root rethrowen, damit der Fallback ohne \\?\ greift.
            # Bei tiefer liegenden Verzeichnissen einfach ueberspringen.
            if ($current -eq $RootPath -and $current.StartsWith('\\?\')) { throw }
        }
        catch [System.NotSupportedException] {
            if ($current -eq $RootPath -and $current.StartsWith('\\?\')) { throw }
        }
        catch { }

        try {
            foreach ($d in [System.IO.Directory]::EnumerateDirectories($current)) {
                $leaf = [System.IO.Path]::GetFileName($d.TrimEnd('\'))
                if ($script:ExcludeDirNames -contains $leaf) { continue }
                # Junctions/Symlinks nicht folgen: sonst werden Dateien
                # doppelt verarbeitet oder die Suche laeuft im Kreis.
                try {
                    $attr = [System.IO.File]::GetAttributes($d)
                    if ($attr -band [System.IO.FileAttributes]::ReparsePoint) { continue }
                } catch { continue }
                $stack.Push($d)
            }
        }
        catch [System.UnauthorizedAccessException]   { }
        catch [System.IO.PathTooLongException]       { $script:SkippedPathTooLong++ }
        catch [System.IO.DirectoryNotFoundException] { }
        catch [System.ArgumentException] {
            if ($current -eq $RootPath -and $current.StartsWith('\\?\')) { throw }
        }
        catch [System.NotSupportedException] {
            if ($current -eq $RootPath -and $current.StartsWith('\\?\')) { throw }
        }
        catch { }
    }
    return $result
}

function Stop-TrackedOfficeProcess {
    param([Nullable[int]]$ProcessId, [Nullable[datetime]]$StartTime = $null)
    if ($null -eq $ProcessId -or $ProcessId -le 0) { return }
    try {
        $p = Get-Process -Id $ProcessId -ErrorAction Stop
        # PID-Recycling-Schutz: Windows kann die freigewordene PID
        # binnen Millisekunden an einen anderen Prozess (Virenscanner,
        # Updater, andere User-Session) vergeben. Ohne Namens-Check
        # wuerde dieser fremde Prozess hier hart gekillt. Nur weiter,
        # wenn der Prozess noch ein Office-Programm ist.
        if ($p.Name -notmatch '^(WINWORD|EXCEL|POWERPNT)$') { return }
        # Erstellungszeit-Check ergaenzt den Namens-Check: die PID kann an
        # eine NEUE Office-Sitzung des Benutzers vergeben worden sein, die
        # der Namens-Check allein nicht erkennt.
        if ($null -ne $StartTime) {
            try { if ($p.StartTime -ne $StartTime) { return } } catch {}
        }
        if (-not $p.HasExited) {
            $p | Stop-Process -Force -ErrorAction SilentlyContinue -WhatIf:$false
        }
    } catch {}
}

function Start-OfficeApp {
    param ([string]$AppType)
    switch ($AppType) {
        "Word" {
            $a = New-Object -ComObject Word.Application -ErrorAction Stop
            $a.DisplayAlerts      = 0
            $a.Visible            = $false
            $a.AutomationSecurity = 3
            try { $a.ScreenUpdating = $false } catch {}
            return $a
        }
        "Excel" {
            $a = New-Object -ComObject Excel.Application -ErrorAction Stop
            $a.DisplayAlerts      = $false
            $a.Visible            = $false
            $a.AutomationSecurity = 3
            $a.ScreenUpdating     = $false
            $a.EnableEvents       = $false
            return $a
        }
        "PowerPoint" {
            $a = New-Object -ComObject PowerPoint.Application -ErrorAction Stop
            $a.Visible            = -1
            try { $a.WindowState = 2 } catch {}
            $a.DisplayAlerts      = 1
            $a.AutomationSecurity = 3
            return $a
        }
    }
}

function Start-TrackedOfficeApp {
    param([string]$AppType, [string]$ProcessName)
    $before = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $app    = Start-OfficeApp $AppType
    Start-Sleep -Milliseconds 150
    $after  = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $newPid = $after | Where-Object { $before -notcontains $_ } | Select-Object -First 1
    $startTime = $null
    if ($newPid) {
        try { $startTime = (Get-Process -Id $newPid -ErrorAction Stop).StartTime } catch {}
    }
    return @{ App = $app; ProcessId = $newPid; StartTime = $startTime }
}

function Close-OfficeDocument {
    param ($obj, [string]$ext)
    if ($null -eq $obj) { return }
    try {
        switch ($ext) {
            ".docm" { $obj.Close(0) }
            ".xlsm" { $obj.Close($false) }
            ".pptm" { $obj.Close() }
        }
    } catch {}
}

# ------------------------------------------------------------------------------
# Zeitwaechter je Datei
# ------------------------------------------------------------------------------
# Open, VBA-Pruefung und SaveAs laufen synchron im Hauptthread. Haengt Office
# an einem unsichtbaren Dialog oder einer defekten Datei, kehrt der Aufruf nie
# zurueck und der Lauf steht - der Smoke-Test sichert nur den Start ab. Ein
# Hintergrund-Runspace beendet deshalb nach Fristablauf die EIGENE
# Office-Instanz (Name und Startzeit geprueft, wie in Skript 9). Der blockierte
# COM-Aufruf wirft daraufhin im Hauptthread, die Datei faellt in den
# Fehlerzweig und bleibt unveraendert (das Original wird erst NACH dem SaveAs
# angefasst), und die Instanz wird neu gestartet.
# Scharfschalten, Entschaerfen und das Beenden laufen unter demselben Schloss.
# Sonst kann der Waechter eine Instanz beenden, deren Aufruf gerade noch
# fertig wurde, ohne dass der Hauptthread es erfaehrt - die naechste Datei
# liefe dann gegen eine tote Instanz.
function Start-DateiWaechter {
    param([int]$PollMs = 250)
    $state = [hashtable]::Synchronized(@{
        Armed       = $false
        TargetPid   = 0
        TargetStart = $null
        Deadline    = [DateTime]::MaxValue
        TimedOut    = $false
        Running     = $true
    })
    $rs = [runspacefactory]::CreateRunspace()
    $rs.Open()
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    $null = $ps.AddScript({
        param($state, $pollMs)
        while ($state.Running) {
            [System.Threading.Monitor]::Enter($state.SyncRoot)
            try {
                if ($state.Armed -and [DateTime]::UtcNow -gt $state.Deadline) {
                    $state.Armed    = $false
                    $state.TimedOut = $true
                    try {
                        $p = Get-Process -Id $state.TargetPid -ErrorAction Stop
                        if ($p.Name -match '^(WINWORD|EXCEL|POWERPNT)$' -and
                            ($null -eq $state.TargetStart -or $p.StartTime -eq $state.TargetStart)) {
                            Stop-Process -Id $state.TargetPid -Force -ErrorAction Stop
                        }
                    } catch {}
                }
            } finally {
                [System.Threading.Monitor]::Exit($state.SyncRoot)
            }
            Start-Sleep -Milliseconds $pollMs
        }
    }).AddArgument($state).AddArgument($PollMs)
    $handle = $ps.BeginInvoke()
    return @{ State = $state; PS = $ps; Runspace = $rs; Handle = $handle }
}

# Ohne eigene PID (Einzelinstanz-Server hat die Sitzung des Anwenders
# geliefert) bleibt der Waechter aus - fremde Instanzen werden nie beendet.
function Set-DateiWaechter {
    param($Waechter, [Nullable[int]]$ProcessId, $StartTime, [int]$TimeoutSec)
    if ($null -eq $Waechter) { return }
    $s = $Waechter.State
    [System.Threading.Monitor]::Enter($s.SyncRoot)
    try {
        $s.TimedOut = $false
        if ($null -eq $ProcessId -or $ProcessId -le 0) { $s.Armed = $false; return }
        $s.TargetPid   = [int]$ProcessId
        $s.TargetStart = $StartTime
        $s.Deadline    = [DateTime]::UtcNow.AddSeconds($TimeoutSec)
        $s.Armed       = $true
    } finally {
        [System.Threading.Monitor]::Exit($s.SyncRoot)
    }
}

# Entschaerft und meldet, ob die Frist abgelaufen ist (die Instanz also
# beendet wurde). Ohne -Quittieren bleibt der Merker stehen, damit der
# Fehlerzweig und der finally-Zweig derselben Datei ihn noch sehen; der
# finally-Zweig quittiert, sonst loeste der Merker bei der naechsten Datei,
# die vor dem Oeffnen abbricht (Kopie gescheitert), einen zweiten Neustart aus.
function Clear-DateiWaechter {
    param($Waechter, [switch]$Quittieren)
    if ($null -eq $Waechter) { return $false }
    $s = $Waechter.State
    [System.Threading.Monitor]::Enter($s.SyncRoot)
    try {
        $s.Armed = $false
        $war = [bool]$s.TimedOut
        if ($Quittieren) { $s.TimedOut = $false }
        return $war
    } finally {
        [System.Threading.Monitor]::Exit($s.SyncRoot)
    }
}

function Stop-DateiWaechter {
    param($Waechter)
    if ($null -eq $Waechter) { return }
    $Waechter.State.Armed   = $false
    $Waechter.State.Running = $false
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while (-not $Waechter.Handle.IsCompleted -and $sw.ElapsedMilliseconds -lt 2000) {
        Start-Sleep -Milliseconds 50
    }
    if ($Waechter.Handle.IsCompleted) {
        try { $Waechter.PS.EndInvoke($Waechter.Handle) } catch {}
    } else {
        try { $Waechter.PS.Stop() } catch {}
    }
    try { $Waechter.PS.Dispose() } catch {}
    try { $Waechter.Runspace.Close() } catch {}
    try { $Waechter.Runspace.Dispose() } catch {}
}

# Beendet die eigene Instanz fuer einen Dateityp und startet sie neu -
# periodisch gegen Speicher-Drift und nach einem Waechter-Eingriff.
# Die Hauptschleife laeuft auf Skriptebene, daher $script:.
function Restart-OfficeFuerTyp {
    param([string]$Ext)
    switch ($Ext) {
        ".docm" {
            Stop-OfficeApp $script:word $officePids.Word
            [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
            Stop-TrackedOfficeProcess $officePids.Word $officeStartTimes.Word
            $r = Start-TrackedOfficeApp -AppType "Word" -ProcessName "WINWORD"
            $script:word = $r.App; $officePids.Word = $r.ProcessId; $officeStartTimes.Word = $r.StartTime
        }
        ".xlsm" {
            Stop-OfficeApp $script:excel $officePids.Excel
            [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
            Stop-TrackedOfficeProcess $officePids.Excel $officeStartTimes.Excel
            $r = Start-TrackedOfficeApp -AppType "Excel" -ProcessName "EXCEL"
            $script:excel = $r.App; $officePids.Excel = $r.ProcessId; $officeStartTimes.Excel = $r.StartTime
        }
        ".pptm" {
            Stop-OfficeApp $script:pptx $officePids.PowerPoint
            [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
            Stop-TrackedOfficeProcess $officePids.PowerPoint $officeStartTimes.PowerPoint
            $r = Start-TrackedOfficeApp -AppType "PowerPoint" -ProcessName "POWERPNT"
            $script:pptx = $r.App; $officePids.PowerPoint = $r.ProcessId; $officeStartTimes.PowerPoint = $r.StartTime
        }
    }
}

# ------------------------------------------------------------------------------
# Office-Smoke-Test (mit Watchdog-Timeout)
# ------------------------------------------------------------------------------
# Falsch konfiguriertes Trust-Center kann den COM-Open in einen unsichtbaren
# Modal-Dialog (Geschuetzte Ansicht) fuehren - der Aufruf wuerde sonst
# dauerhaft blockieren. Der Test laeuft daher in einem isolierten Background-
# Job (eigener PowerShell-Prozess) mit eigener Office-Instanz; bei Hänger
# beendet der Watchdog den Job hart und der haengende Office-Prozess kann
# zusaetzlich per gemerkter PID gekillt werden.
function Test-TrustCenterApp {
    param(
        [Parameter(Mandatory)] [string] $AppType,    # "Word", "Excel", "PowerPoint"
        [Parameter(Mandatory)] [string] $TestPath,
        [Parameter(Mandatory)] [int]    $MacroFormat,
        [Parameter(Mandatory)] [string] $Ext,        # ".docm", ".xlsm", ".pptm"
        [int]                           $TimeoutSec = 30
    )

    # Job in einer eigenen PowerShell-Instanz: erstellt eine frische
    # Office-Instanz, fuehrt Add+SaveAs+Open+VBProject-Check aus, gibt
    # @{ Success=$true/$false; ChildPid=<PID> } zurueck. Bei Timeout
    # killt der Watchdog den Job und ggf. die Office-PID.
    $job = Start-Job -ScriptBlock {
        param($AppType, $TestPath, $MacroFormat, $Ext)

        $progName = switch ($AppType) {
            "Word"       { "Word.Application" }
            "Excel"      { "Excel.Application" }
            "PowerPoint" { "PowerPoint.Application" }
        }
        $procName = switch ($AppType) {
            "Word"       { "WINWORD" }
            "Excel"      { "EXCEL" }
            "PowerPoint" { "POWERPNT" }
        }

        $existingPids = @()
        if ($AppType -eq "Word") {
            $existingPids = @(Get-Process -Name $procName -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
        } else {
            try {
                Add-Type -Name "User32MacroSmokeTest" -Namespace "Win32" -MemberDefinition @'
[DllImport("user32.dll")]
public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
'@ -ErrorAction SilentlyContinue
            } catch {}
        }

        $app     = $null
        $childPid = $null
        $doc     = $null
        $success = $false
        try {
            $app = New-Object -ComObject $progName
            try { $app.Visible       = $false } catch {}
            try { $app.DisplayAlerts = 0 }     catch {}
            try { $app.AutomationSecurity = 3 } catch {}
            if ($AppType -eq "PowerPoint") {
                # PowerPoint duldet kein Visible=$false; minimieren stattdessen.
                try { $app.Visible = $true } catch {}
                try { $app.WindowState = 2 } catch {}
            }

            if ($AppType -eq "Word") {
                Start-Sleep -Milliseconds 200
                foreach ($p in (Get-Process -Name $procName -ErrorAction SilentlyContinue)) {
                    if ($existingPids -notcontains $p.Id) { $childPid = $p.Id; break }
                }
            } else {
                try {
                    $procId = 0
                    [Win32.User32MacroSmokeTest]::GetWindowThreadProcessId([IntPtr]$app.Hwnd, [ref]$procId) | Out-Null
                    if ($procId -gt 0) { $childPid = [int]$procId }
                } catch {}
            }

            switch ($Ext) {
                ".docm" {
                    $doc = $app.Documents.Add()
                    # AddToRecentFiles = $false (5. Argument): siehe Hauptlauf.
                    $doc.SaveAs2($TestPath, $MacroFormat, $false, "", $false)
                    $doc.Close($false)
                    Start-Sleep -Milliseconds 200
                    $doc = $app.Documents.Open($TestPath, $false, $false, $false)
                }
                ".xlsm" {
                    $doc = $app.Workbooks.Add()
                    $doc.SaveAs($TestPath, $MacroFormat)
                    $doc.Close($false)
                    Start-Sleep -Milliseconds 200
                    $doc = $app.Workbooks.Open($TestPath, 0, $false)
                }
                ".pptm" {
                    $doc = $app.Presentations.Add(-1)
                    $doc.SaveAs($TestPath, $MacroFormat)
                    $doc.Close()
                    Start-Sleep -Milliseconds 200
                    $doc = $app.Presentations.Open($TestPath, 0, 0, -1)
                }
            }

            try {
                $vbp = $doc.VBProject
                if ($null -ne $vbp) {
                    $null = $vbp.VBComponents.Count
                    $success = $true
                }
            } catch {
                $success = $false
            }

            try {
                switch ($Ext) {
                    ".docm" { $doc.Close(0) }
                    ".xlsm" { $doc.Close($false) }
                    ".pptm" { $doc.Close() }
                }
            } catch {}
        } catch {
            $success = $false
        } finally {
            if ($app) {
                # Auch hier: nur die eigene Instanz beenden.
                if ($childPid) { try { $app.Quit() } catch {} }
                else { try { $app.Visible = $true } catch {} }
                try {
                    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null
                } catch {}
            }
            [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
            if (Test-Path -LiteralPath $TestPath) {
                Remove-Item -LiteralPath $TestPath -Force -ErrorAction SilentlyContinue
            }
        }
        return @{ Success = $success; ChildPid = $childPid }
    } -ArgumentList $AppType, $TestPath, $MacroFormat, $Ext

    $finished = Wait-Job -Job $job -Timeout $TimeoutSec

    if ($null -eq $finished) {
        # Timeout - Job killen + ggf. zugehoerigen Office-Prozess.
        try { Stop-Job -Job $job -ErrorAction SilentlyContinue -WhatIf:$false } catch {}
        # Bevor wir die Job-Output entsorgen, versuchen wir noch die ChildPid
        # zu lesen (Job hat sie evtl. geschrieben, bevor er hing).
        $childPid = $null
        try {
            $partial = Receive-Job -Job $job -Keep -ErrorAction SilentlyContinue
            if ($partial -and $partial.ChildPid) { $childPid = [int]$partial.ChildPid }
        } catch {}
        if ($childPid) {
            # Via Stop-TrackedOfficeProcess, damit der Office-Namens-
            # Check greift (PID-Recycling-Schutz - siehe dort).
            Stop-TrackedOfficeProcess $childPid
        }
        try { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue -WhatIf:$false } catch {}
        if (Test-Path -LiteralPath $TestPath) {
            Remove-Item -LiteralPath $TestPath -Force -ErrorAction SilentlyContinue -WhatIf:$false
        }
        return @{ Success = $false; TimedOut = $true; Message = "TIMEOUT nach $TimeoutSec s" }
    }

    $jobResult = $null
    try { $jobResult = Receive-Job -Job $job -ErrorAction SilentlyContinue } catch {}
    try { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue -WhatIf:$false } catch {}

    if ($jobResult -and $jobResult.Success) {
        return @{ Success = $true; TimedOut = $false; Message = "" }
    }
    return @{ Success = $false; TimedOut = $false; Message = "VBA-Zugriff/Open fehlgeschlagen" }
}

# ==============================================================================
# Banner
# ==============================================================================
if (-not $Automated) { Show-Banner }

# Vor dem ersten COM-Zugriff auf laufende Office-Sitzungen hinweisen.
if (-not (Show-OfficeRunningWarning -Silent:$Automated)) { exit 0 }


# ==============================================================================
# Hinweis Trust Center (vor Verzeichnisauswahl, damit der Anwender frueh weiss,
# was vorzubereiten ist)
# ==============================================================================
if (-not $Automated) {
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host "WICHTIGER HINWEIS - Vertrauenswuerdige Speicherorte"      -ForegroundColor Cyan
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Damit das Skript VBA-Projekte pruefen kann, muss der"     -ForegroundColor Cyan
    Write-Host "temporaere Arbeitsordner als vertrauenswuerdig"            -ForegroundColor Cyan
    Write-Host "hinterlegt sein:"                                          -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  $tempDir"                                                -ForegroundColor White
    Write-Host ""
    Write-Host "Bitte fuehren Sie VORAB folgende Schritte in jeder"        -ForegroundColor Cyan
    Write-Host "Office-Anwendung (Word, Excel, PowerPoint) durch:"         -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  1.  Datei > Optionen > Trust Center > Einstellungen"     -ForegroundColor Cyan
    Write-Host "      fuer das Trust Center"                               -ForegroundColor Cyan
    Write-Host "  2.  Vertrauenswuerdige Speicherorte > Neuen Speicherort" -ForegroundColor Cyan
    Write-Host "      hinzufuegen"                                         -ForegroundColor Cyan
    Write-Host "  3.  Pfad eintragen:  $tempDir"                           -ForegroundColor Cyan
    Write-Host "  4.  Haken setzen bei: 'Unterordner dieses Speicherorts"  -ForegroundColor Cyan
    Write-Host "      sind ebenfalls vertrauenswuerdig'"                   -ForegroundColor Cyan
    Write-Host "  5.  Haken setzen bei: 'Vertrauenswuerdige Speicherorte"  -ForegroundColor Cyan
    Write-Host "      im Netzwerk zulassen' (ganz unten im Fenster)"       -ForegroundColor Cyan
    Write-Host "  6.  Mit OK bestaetigen"                                  -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Zusaetzlich muss der VBA-Projektzugriff aktiviert sein:"   -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Trust Center > Makroeinstellungen >"                     -ForegroundColor Cyan
    Write-Host "  'Zugriff auf das VBA-Projektobjektmodell vertrauen'"     -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Hinweis: Mit ESC kann der Lauf jederzeit sauber"           -ForegroundColor Yellow
    Write-Host "        abgebrochen werden."                               -ForegroundColor Yellow
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host ""

    $answer = Read-Host "Haben Sie die Einstellungen vorgenommen? (J/N)"
    if ($answer -notmatch '^[jJyY]') {
        Write-Host "Skript abgebrochen. Bitte zuerst die Einstellungen vornehmen." -ForegroundColor Yellow
        exit
    }
}

# ==============================================================================
# Verzeichnisauswahl
# ==============================================================================
$desktopPath   = Get-UserShellFolder -Name "Desktop" `
                                     -Fallback (Join-Path $env:USERPROFILE "Desktop")
$downloadsPath = Get-UserShellFolder -Name "{374DE290-123F-4565-9164-39C4925E467B}" `
                                     -Fallback (Join-Path $env:USERPROFILE "Downloads")

if (-not $Automated) {
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host "Welches Verzeichnis soll durchsucht werden?"              -ForegroundColor Cyan
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host ""
    $presetCount = $script:DirectoryPresets.Count
    for ($i = 0; $i -lt $presetCount; $i++) {
        Write-Host ("  [" + ($i + 1) + "]  " + $script:DirectoryPresets[$i]) -ForegroundColor White
    }
    # Verkettung statt -f: Pfade koennen '{' oder '}' enthalten.
    Write-Host ("  [" + ($presetCount + 1) + "]  Desktop      ($desktopPath)")   -ForegroundColor White
    Write-Host ("  [" + ($presetCount + 2) + "]  Downloads    ($downloadsPath)") -ForegroundColor White
    Write-Host ("  [" + ($presetCount + 3) + "]  Eigenen Pfad eingeben")         -ForegroundColor White
    Write-Host ""

    $driveChoice = Read-Host ("Ihre Wahl (1-" + ($presetCount + 3) + ")")
    $choiceNum = 0
    [void][int]::TryParse($driveChoice, [ref]$choiceNum)

    switch ($true) {
        ($choiceNum -ge 1 -and $choiceNum -le $presetCount) {
            $rootPath = $script:DirectoryPresets[$choiceNum - 1]
            break
        }
        ($choiceNum -eq $presetCount + 1) { $rootPath = $desktopPath;   break }
        ($choiceNum -eq $presetCount + 2) { $rootPath = $downloadsPath; break }
        ($choiceNum -eq $presetCount + 3) {
            $rootPath = Read-Host "Bitte vollstaendigen Pfad eingeben (z.B. S:\Abteilung)"
            $rootPath = $rootPath.Trim().Trim('"').Trim("'")
            if ($rootPath -match '^[A-Za-z]:$') { $rootPath = $rootPath + '\' }
            break
        }
        default {
            Write-Host "Ungueltige Auswahl. Skript wird abgebrochen." -ForegroundColor Red
            exit
        }
    }
    if ($rootPath -notmatch '\\$') { $rootPath = $rootPath + '\' }

    if (-not (Test-Path -LiteralPath $rootPath)) {
        Write-Host "FEHLER: Der Pfad '$rootPath' ist nicht erreichbar." -ForegroundColor Red
        Write-Host "Bitte pruefen Sie die Netzwerkverbindung und starten Sie das Skript erneut." -ForegroundColor Yellow
        exit
    }

    Write-Host ""
    Write-Host "Gewaehlter Pfad: $rootPath" -ForegroundColor Green
    Write-Host ""

    # --- Vorschau-Modus ---
    # -WhatIf und -ReadOnlyMode haben Vorrang; sonst interaktiv nachfragen.
    if (-not $script:PreviewOnly -and -not $PSBoundParameters.ContainsKey('WhatIf')) {
        $pv = Read-Host "Nur pruefen, nichts umwandeln (Probelauf)? [j/N]"
        if ($pv -and $pv.Trim().ToLowerInvariant() -in @('j','ja','y','yes')) {
            $script:PreviewOnly = $true
        }
    }
    if ($script:PreviewOnly) {
        Write-Host "PROBELAUF - es wird NICHTS umgewandelt." -ForegroundColor Yellow
        Write-Host ""
    }
} else {
    if ($rootPath -ne "") {
        $rootPath = $rootPath.Trim().Trim('"').Trim("'")
        if ($rootPath -match '^[A-Za-z]:$') { $rootPath = $rootPath + '\' }
        if ($rootPath -notmatch '\\$')      { $rootPath = $rootPath + '\' }
    } else {
        Write-Host "FEHLER: Im Modus -Automated ist -RootPath erforderlich." -ForegroundColor Red
        throw "RootPath fehlt (Automated-Modus)"
    }
    if (-not (Test-Path -LiteralPath $rootPath)) {
        Write-Host "FEHLER: Der Pfad '$rootPath' ist nicht erreichbar." -ForegroundColor Red
        throw "Pfad nicht erreichbar (Automated-Modus): $rootPath"
    }
}

# Relativen Pfad absolut machen. Test-Path loest gegen $PWD auf, die
# [System.IO]-Aufrufe der Suche aber gegen das Arbeitsverzeichnis des
# PROZESSES - bei '-RootPath .' oder 'Ordner' wurde so ein anderer Baum
# durchsucht bzw. aus 'Ordner' ein ungueltiges '\\?\Ordner' (nachgestellt in
# 9, dasselbe Muster). ProviderPath liefert den echten Dateisystempfad, auch
# fuer gemappte Laufwerke und UNC.
try {
    $rootPath = (Resolve-Path -LiteralPath $rootPath -ErrorAction Stop).ProviderPath
    if ($rootPath -notmatch '\\$') { $rootPath = $rootPath + '\' }
} catch {}

# ==============================================================================
# Fortschrittsanzeige (3-Optionen-Menu, analog zu 2a/2b/2c)
# ==============================================================================
$useProgress = $false
$skipPreScan = $true
$script:LastProgressUpdate    = [DateTime]::MinValue
$script:ProgressMinIntervalMs = 500
$script:ShouldStop            = $false

# Strg+C als Eingabe behandeln und zusammen mit ESC in der Hauptschleife
# pollen: ein add_CancelKeyPress-ScriptBlock wuerde auf dem Ctrl+C-Thread
# ohne PowerShell-Runspace laufen und dort fehlschlagen - der Abbruch
# kaeme dann doch hart mitten in der Datei-Ersetzung. Nur interaktiv
# (im Automated-Modus pollt niemand die Tasten).
$script:CtrlCAsInput = $false
try {
    if (-not $Automated -and $Host.Name -eq 'ConsoleHost') {
        [Console]::TreatControlCAsInput = $true
        $script:CtrlCAsInput = $true
    }
} catch {}
if (-not $Automated) {
    Write-Host "Fortschrittsanzeige:" -ForegroundColor Yellow
    Write-Host "  [1] Keine Fortschrittsleiste (laufende Zählung inline) [Standard]" -ForegroundColor White
    Write-Host "  [2] Fortschrittsleiste mit ETA (Vorab-Scan nötig, Start verzögert!)" -ForegroundColor White
    Write-Host "  [3] Fortschrittsleiste ohne ETA (sofortiger Start, kein Total)" -ForegroundColor White
    Write-Host ""
    Write-Host "  Hinweis zu [2]: Bei UNC-Pfaden oder sehr großen Freigaben kann" -ForegroundColor DarkGray
    Write-Host "  die Indizierung 10–30 Minuten dauern, bevor die erste Datei" -ForegroundColor DarkGray
    Write-Host "  verarbeitet wird. Bei > 100.000 Dateien besser [1] oder [3] wählen." -ForegroundColor DarkGray
    Write-Host ""
    $progressChoice = (Read-Host "Auswahl [1-3]").Trim()
    switch ($progressChoice) {
        "2"     { $useProgress = $true;  $skipPreScan = $false }
        "3"     { $useProgress = $true;  $skipPreScan = $true  }
        default { $useProgress = $false; $skipPreScan = $true  }
    }
    Write-Host ""
}

# ==============================================================================
# Logfile-Initialisierung
# ==============================================================================
# Log-Verzeichnis: bevorzugt neben dem Skript, sonst LOCALAPPDATA, sonst TEMP.
# Vorher wurde ungeprueft in $scriptDir geschrieben - lag das Skript auf einer
# schreibgeschuetzten Freigabe, brach der Lauf hier mit unbehandelter
# Ausnahme ab, bevor eine einzige Datei geprueft war.
$logDirCandidates = @($scriptDir,
                      (Join-Path $env:LOCALAPPDATA 'OfficeMacroDeepConverter'),
                      $env:TEMP)
$logInitOk = $false
foreach ($cand in $logDirCandidates) {
    if ([string]::IsNullOrWhiteSpace($cand)) { continue }
    try {
        if (-not [System.IO.Directory]::Exists($cand)) {
            New-Item -ItemType Directory -Path $cand -Force -ErrorAction Stop -WhatIf:$false | Out-Null
        }
        $tryLog  = Join-Path $cand (Split-Path $logFile -Leaf)
        $trySkip = Join-Path $cand (Split-Path $skipLog -Leaf)
        [System.IO.File]::WriteAllText($tryLog,  "--- START DER TIEFENPRUEFUNG: $(Get-Date) ---`r`n", $utf8Bom)
        [System.IO.File]::WriteAllText($trySkip, "Quelle;Zieldatei;Grund`r`n",            $utf8Bom)
        $logFile            = $tryLog
        $skipLog            = $trySkip
        $script:SkipLogPath = $trySkip
        $logInitOk          = $true
        break
    } catch { continue }
}
if (-not $logInitOk) {
    Write-Host "WARNUNG: Keine Logdatei schreibbar - Lauf wird nicht protokolliert." -ForegroundColor Red
}

Write-Log "Temporaerer Arbeitsordner: $tempDir"

# Restore-Privilegien (Admin-Kontext) fuer ACL/Owner-Erhalt aktivieren.
Enable-RestorePrivileges

Write-Log "Initialisiere Office-Komponenten..."

# Zuletzt verwendet: Ausgangszustand merken; am Ende verschwinden nur die
# Verknuepfungen, die Office fuer eigene Arbeitskopien bzw. das bearbeitete
# Verzeichnis angelegt hat (_gemeinsam.psm1; gemessen 29.09.2026: 13 je Lauf).
if ($script:GemeinsamGeladen) {
    try { Start-RecentMomentaufnahme; Add-RecentWurzel $RootPath } catch { }
}

# ==============================================================================
# Office-Initialisierung
# ==============================================================================
$word = $null; $excel = $null; $pptx = $null
$dateiWaechter = $null
$abortedByException = $false

try {

try {
    $wRes  = Start-TrackedOfficeApp -AppType "Word"       -ProcessName "WINWORD"
    $word  = $wRes.App;  $officePids.Word       = $wRes.ProcessId;  $officeStartTimes.Word       = $wRes.StartTime

    $eRes  = Start-TrackedOfficeApp -AppType "Excel"      -ProcessName "EXCEL"
    $excel = $eRes.App;  $officePids.Excel      = $eRes.ProcessId;  $officeStartTimes.Excel      = $eRes.StartTime

    $pRes  = Start-TrackedOfficeApp -AppType "PowerPoint" -ProcessName "POWERPNT"
    $pptx  = $pRes.App;  $officePids.PowerPoint = $pRes.ProcessId;  $officeStartTimes.PowerPoint = $pRes.StartTime

    Write-Log ("Office-Schnittstellen erfolgreich geladen (PIDs: Word={0}, Excel={1}, PPT={2})." -f `
        $officePids.Word, $officePids.Excel, $officePids.PowerPoint)
} catch {
    Write-Log "Kritischer Fehler beim Office-Start: $($_.Exception.Message)" -Level "ERROR"
    Stop-OfficeApp $word $officePids.Word
    Stop-OfficeApp $excel $officePids.Excel
    Stop-OfficeApp $pptx $officePids.PowerPoint
    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    Stop-TrackedOfficeProcess $officePids.Word $officeStartTimes.Word
    Stop-TrackedOfficeProcess $officePids.Excel $officeStartTimes.Excel
    Stop-TrackedOfficeProcess $officePids.PowerPoint $officeStartTimes.PowerPoint
    throw "Office-Start fehlgeschlagen: $($_.Exception.Message)"
}

# ==============================================================================
# Office-Smoke-Test (mit Watchdog-Timeout)
# ==============================================================================
Write-Log "Fuehre Office-Smoke-Test durch (Timeout 30 s pro App)..."

$smokeWord  = Test-TrustCenterApp -AppType "Word"       -TestPath (Join-Path $tempDir "_smoketest.docm") -MacroFormat 13 -Ext ".docm" -TimeoutSec 30
$smokeExcel = Test-TrustCenterApp -AppType "Excel"      -TestPath (Join-Path $tempDir "_smoketest.xlsm") -MacroFormat 52 -Ext ".xlsm" -TimeoutSec 30
$smokePpt   = Test-TrustCenterApp -AppType "PowerPoint" -TestPath (Join-Path $tempDir "_smoketest.pptm") -MacroFormat 25 -Ext ".pptm" -TimeoutSec 30

$allOk = $smokeWord.Success -and $smokeExcel.Success -and $smokePpt.Success
if (-not $allOk) {
    function Format-SmokeStatus {
        param($Result)
        if ($Result.Success)  { return "OK" }
        if ($Result.TimedOut) { return "TIMEOUT" }
        return "FAIL"
    }
    $statusWord  = Format-SmokeStatus $smokeWord
    $statusExcel = Format-SmokeStatus $smokeExcel
    $statusPpt   = Format-SmokeStatus $smokePpt
    Write-Log "Office-Smoke-Test fehlgeschlagen: Word=$statusWord, Excel=$statusExcel, PowerPoint=$statusPpt" -Level "ERROR"
    if ($smokeWord.TimedOut)  { Write-Log "  Word:       $($smokeWord.Message)"  -Level "ERROR" }
    if ($smokeExcel.TimedOut) { Write-Log "  Excel:      $($smokeExcel.Message)" -Level "ERROR" }
    if ($smokePpt.TimedOut)   { Write-Log "  PowerPoint: $($smokePpt.Message)"   -Level "ERROR" }
    Write-Log "Bitte Trust-Center-Einstellungen fuer alle drei Office-Anwendungen wie beschrieben konfigurieren und Skript neu starten." -Level "ERROR"
    Stop-OfficeApp $word $officePids.Word
    Stop-OfficeApp $excel $officePids.Excel
    Stop-OfficeApp $pptx $officePids.PowerPoint
    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    Stop-TrackedOfficeProcess $officePids.Word $officeStartTimes.Word
    Stop-TrackedOfficeProcess $officePids.Excel $officeStartTimes.Excel
    Stop-TrackedOfficeProcess $officePids.PowerPoint $officeStartTimes.PowerPoint
    if ($Automated) {
        throw "Office-Smoke-Test fehlgeschlagen"
    }
    Write-Host ""
    Write-Host "Druecken Sie eine beliebige Taste zum Beenden..." -ForegroundColor Cyan
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}
Write-Log "Office-Smoke-Test erfolgreich (Word, Excel, PowerPoint)." -Level "SUCCESS"

# ==============================================================================
# Konfiguration Dateitypen
# ==============================================================================
$fileTypes = @(
    @{ Ext = ".docm"; Target = ".docx"; Format = 12; App = "Word" }
    @{ Ext = ".xlsm"; Target = ".xlsx"; Format = 51; App = "Excel" }
    @{ Ext = ".pptm"; Target = ".pptx"; Format = 24; App = "PowerPoint" }
)

# ==============================================================================
# Zeitwaechter je Datei starten (siehe Start-DateiWaechter)
# ==============================================================================
$dateiWaechter = Start-DateiWaechter
foreach ($t in $fileTypes) {
    if (-not ($officePids[$t.App] -gt 0)) {
        Write-Log (("Zeitwaechter fuer {0} wirkungslos: keine eigene Instanz erkannt (laeuft {0} bereits?). " +
                    "Eine haengende {1}-Datei haelt den Lauf dann an.") -f $t.App, $t.Ext) -Level "WARN"
    }
}

# ==============================================================================
# Long-Path Vorbereitung
# ==============================================================================
$longPath = Add-LongPathPrefix $rootPath

# ==============================================================================
# Statistik
# ==============================================================================
$stats = @{
    Scanned        = 0
    Converted      = 0
    Renamed        = 0
    KeptMacro      = 0
    KeptXlm        = 0
    KeptXlmUnknown = 0
    KeptNoAccess   = 0
    SkippedLong    = 0
    SkippedPassword = 0
    Errors         = 0
    WouldConvert   = 0
}

# ==============================================================================
# Vorab-Zaehlung (nur bei Option [2] mit ETA)
# ==============================================================================
$totalFiles = 0
if ($useProgress -and -not $skipPreScan) {
    Write-Log "Vorab-Zaehlung der Dateien fuer ETA-Anzeige (kann bei grossen Netzlaufwerken dauern)..."
    foreach ($type in $fileTypes) {
        try {
            $totalFiles += (Get-FilesRecursiveSafe -RootPath $longPath -Pattern "*$($type.Ext)").Count
        } catch {
            try {
                $totalFiles += (Get-FilesRecursiveSafe -RootPath $rootPath -Pattern "*$($type.Ext)").Count
            } catch {}
        }
    }
    Write-Log "Vorab-Zaehlung: $totalFiles Datei(en) gefunden." -Level "INFO"
}

# ==============================================================================
# Hauptschleife
# ==============================================================================
$startTime = Get-Date
foreach ($type in $fileTypes) {
    if ($cancelled) { break }
    Write-Log "Suche nach $($type.Ext) Dateien in: $rootPath"

    $filePaths = $null
    try {
        $filePaths = Get-FilesRecursiveSafe -RootPath $longPath -Pattern "*$($type.Ext)"
    } catch {
        Write-Log "Long-Path-Suche fehlgeschlagen, verwende Fallback: $($_.Exception.Message)" -Level "WARN"
        try {
            $filePaths = Get-FilesRecursiveSafe -RootPath $rootPath -Pattern "*$($type.Ext)"
        } catch {
            Write-Log "Fallback-Suche ebenfalls fehlgeschlagen: $($_.Exception.Message)" -Level "ERROR"
            continue
        }
    }

    $processedInBatch = 0

    foreach ($rawPath in $filePaths) {

        # ==============================================================================
        # ESC-Abbruch-Pruefung
        # ==============================================================================
        if (-not $Automated) {
            try {
                while ([Console]::KeyAvailable) {
                    $key = [Console]::ReadKey($true)
                    if ($key.Key -eq 'Escape') {
                        Write-Log "Abbruch durch Anwender (ESC)." -Level "WARN"
                        $cancelled = $true
                        break
                    }
                    if ($key.Key -eq [ConsoleKey]::C -and
                        (($key.Modifiers -band [ConsoleModifiers]::Control) -ne 0)) {
                        Write-Log "Abbruch durch Anwender (Strg+C)." -Level "WARN"
                        $cancelled = $true
                        break
                    }
                }
            } catch {}
            if ($cancelled) { break }
        }

        if ($script:ShouldStop) {
            Write-Log "Abbruch durch Anwender (Strg+C)." -Level "WARN"
            $cancelled = $true
            break
        }

        # ==============================================================================
        # Periodischer App-Restart gegen COM-Speicher-Drift
        # ==============================================================================
        if ($processedInBatch -ge $restartThreshold) {
            Write-Log "App-Restart nach $processedInBatch Dateien (Memory-Hygiene)..." -Level "INFO"
            Restart-OfficeFuerTyp $type.Ext
            $processedInBatch = 0
        }

        $stats.Scanned++
        $processedInBatch++

        $filePath         = Remove-LongPathPrefix $rawPath
        $filePathLong     = Add-LongPathPrefix $filePath
        $fileName         = Split-Path $filePath -Leaf

        # ==============================================================================
        # Fortschrittsanzeige
        # ==============================================================================
        if ($useProgress) {
            $now = Get-Date
            $isFinal = ($totalFiles -gt 0 -and $stats.Scanned -ge $totalFiles)
            $skipUpdate = (-not $isFinal) -and `
                (($now - $script:LastProgressUpdate).TotalMilliseconds -lt $script:ProgressMinIntervalMs)
            if (-not $skipUpdate) {
                $script:LastProgressUpdate = $now
                $elapsed = $now - $startTime
                if ($skipPreScan -or $totalFiles -le 0) {
                    # Option [3]: ohne ETA, kein Total
                    $activity = "Office Macro Deep-Converter  |  $($stats.Scanned) Datei(en)"
                    Write-Progress -Activity $activity -Status $fileName -PercentComplete 0
                } else {
                    # Option [2]: mit ETA und Total
                    $pct = [int](($stats.Scanned / $totalFiles) * 100)
                    if ($pct -gt 100) { $pct = 100 }
                    $etaStr = "Berechne..."
                    if ($stats.Scanned -gt 0) {
                        $secPerFile = $elapsed.TotalSeconds / $stats.Scanned
                        $remaining  = ($totalFiles - $stats.Scanned) * $secPerFile
                        $etaSpan    = [TimeSpan]::FromSeconds([Math]::Max(0, $remaining))
                        $etaStr     = "ETA: " + $etaSpan.ToString('hh\:mm\:ss')
                    }
                    $activity = "Office Macro Deep-Converter  |  $etaStr  |  $($stats.Scanned) / $totalFiles"
                    Write-Progress -Activity $activity -Status $fileName -PercentComplete $pct
                }
            }
        }

        $hasCode          = $false
        $noAccessOccurred = $false
        $xlmHandled       = $false
        $obj              = $null
        $origTimestamp    = $null

        $guidDir     = Join-Path $tempDir ([Guid]::NewGuid().ToString("N"))
        $localCopy   = $null
        $localTarget = $null

        try {
            New-Item -ItemType Directory -Path $guidDir -Force -WhatIf:$false | Out-Null

            # Lokale Kopie unter generischem Kurznamen, damit Original-Filename
            # die 259-Zeichen-Grenze des lokalen Temp-Pfads nicht reisst.
            #
            # [string] ist Pflicht, nicht Zierde: Join-Path liefert einen in
            # PSObject verpackten String, und den reicht PowerShell an COM
            # nicht als Text weiter. Word.SaveAs2 blieb damit bei JEDER .docm
            # haengen - der Hauptthread drehte mit voller CPU-Last und kam
            # auch nicht zurueck, als Word beendet wurde (nachgestellt unter
            # 5.1 mit Word 2024: mit Join-Path haengt es, mit [string] ist
            # die Datei nach 0,1 s geschrieben). Documents.Open vertrug den
            # verpackten Wert zufaellig.
            $localCopy = [string](Join-Path $guidDir "work$($type.Ext)")

            try {
                $origTimestamp = [System.IO.File]::GetLastWriteTime($filePathLong)
            } catch {}

            # ==============================================================================
            # Lokale Kopie mit Retry und AV-Wartelogik
            # ==============================================================================
            # [System.IO] statt Copy-Item: die Provider-Cmdlets von
            # PowerShell 5.1 kommen mit '\\?\'-Praefixen nicht zuverlaessig
            # zurecht (siehe die gleichlautende Begruendung in Skript 6).
            try {
                Invoke-WithRetry -Action {
                    [System.IO.File]::Copy($filePathLong, $localCopy, $true)
                } -MaxAttempts 4 -DelayMs 500
            } catch {
                Write-Log "Kopie nach lokal fehlgeschlagen fuer $fileName : $($_.Exception.Message)" -Level "WARN"
                $stats.Errors++
                continue
            }

            if (-not (Wait-FileAvailable -Path $localCopy -MaxAttempts 12 -DelayMs 300)) {
                Write-Log "Lokale Kopie nicht freigegeben (AV-Scanner?): $fileName" -Level "WARN"
                $stats.Errors++
                continue
            }

            $kennwort = Get-OoxmlKennwortGrund -Path $localCopy -Ext $type.Ext
            if ($kennwort -eq 'verschluesselt') {
                Write-Log "UEBERSPRUNGEN ($kennwort): $fileName" -Level "WARN"
                Write-SkipLogEntry -Source $filePath -Reason "Kennwort ($kennwort) - nicht geprueft"
                $stats.SkippedPassword++
                continue
            }
            if ($kennwort) {
                try {
                    Remove-OoxmlSchreibkennwort -Path $localCopy -Ext $type.Ext
                    Write-Log "Schreibkennwort in der Arbeitskopie entfernt: $fileName" -Level "INFO"
                } catch {
                    Write-Log "Schreibkennwort nicht entfernbar ($fileName): $($_.Exception.Message)" -Level "WARN"
                    Write-SkipLogEntry -Source $filePath -Reason "Schreibkennwort nicht entfernbar"
                    $stats.SkippedPassword++
                    continue
                }
            }

            Set-DateiWaechter $dateiWaechter -ProcessId $officePids[$type.App] `
                -StartTime $officeStartTimes[$type.App] -TimeoutSec $script:DateiTimeoutSec
            switch ($type.Ext) {
                # 4. Argument AddToRecentFiles = $false. Ohne es trug Word jede
                # Arbeitskopie in seine Liste "Zuletzt verwendet" ein - am
                # 29.09.2026 standen 21 solche Eintraege aus 7 in der Registry.
                ".docm" { $obj = $word.Documents.Open($localCopy, $false, $false, $false) }
                ".xlsm" { $obj = $excel.Workbooks.Open($localCopy, 0, $false) }
                ".pptm" { $obj = $pptx.Presentations.Open($localCopy, 0, 0, -1) }
            }

            # Excel-4.0-Makros erkennen, BEVOR ueber "makrofrei" entschieden
            # wird - HasVBProject sieht sie nicht (siehe Test-HasExcel4Macro).
            if ($type.Ext -eq ".xlsm") {
                $xlm = Test-HasExcel4Macro $localCopy
                if ($xlm -ne $false) {
                    $hasCode = $true
                    # Merker fuer den Abschlusszweig weiter unten. Ohne ihn
                    # wurde die Datei dort ein ZWEITES Mal gezaehlt
                    # ($stats.KeptMacro) und bekam eine zweite CSV-Zeile mit
                    # dem Grund "VBA-Projekt vorhanden" - sachlich falsch,
                    # denn ein XLM-Makroblatt ist gerade KEIN VBA-Projekt.
                    $xlmHandled = $true
                    if ($null -eq $xlm) {
                        Write-Log "XLM-Pruefung nicht moeglich fuer $fileName - Datei wird zur Sicherheit behalten." -Level "WARN"
                        $stats.KeptXlmUnknown++
                        $csvGrund = "XLM-Pruefung nicht moeglich - manuell pruefen"
                    } else {
                        Write-Log "EXCEL-4.0-MAKRO (XLM) gefunden: $fileName (Wird beibehalten)" -Level "WARN"
                        $stats.KeptXlm++
                        $csvGrund = "Excel-4.0-Makro (XLM) in Makroblatt - kein VBA, aber Makroinhalt"
                    }
                    Write-SkipLogEntry -Source $filePath -Reason $csvGrund
                }
            }

            if ($null -ne $obj) {
                $canCheckVBA = $false
                try {
                    $canCheckVBA = $obj.HasVBProject
                } catch {
                    try {
                        $testVB = $obj.VBProject
                        $canCheckVBA = ($null -ne $testVB)
                        Release-ComObject $testVB
                    } catch {
                        Write-Log "Kein VBA-Zugriff moeglich fuer $fileName. Behalte Datei zur Sicherheit." -Level "WARN"
                        $hasCode = $true
                        $noAccessOccurred = $true
                    }
                }

                if ($canCheckVBA -and -not $hasCode) {
                    $vbProject    = $null
                    $vbComponents = $null
                    $module       = $null
                    $codeModule   = $null
                    try {
                        $lineCount    = 0
                        $vbProject    = $obj.VBProject
                        $vbComponents = $vbProject.VBComponents
                        foreach ($module in $vbComponents) {
                            $codeModule = $module.CodeModule
                            $lineCount += $codeModule.CountOfLines
                            Release-ComObject $codeModule
                            $codeModule = $null
                        }
                        if ($lineCount -gt 0) { $hasCode = $true }
                    } catch {
                        Write-Log "Zugriff verweigert auf VBA in $fileName. Behalte Datei zur Sicherheit." -Level "WARN"
                        $hasCode = $true
                        $noAccessOccurred = $true
                    } finally {
                        Release-ComObject $codeModule
                        Release-ComObject $vbComponents
                        Release-ComObject $vbProject
                    }
                }

                # Lief die Frist waehrend der VBA-Pruefung ab, sind die
                # Fehler der inneren catch-Zweige Folge des Eingriffs und
                # kein Trust-Center-Problem - sonst stuende die Datei als
                # "Kein VBA-Zugriff (Trust Center)" in der CSV statt als
                # TIMEOUT. Die Ausnahme fuehrt in den Fehlerzweig, der die
                # Datei unveraendert laesst und TIMEOUT protokolliert.
                if ($null -ne $dateiWaechter -and $dateiWaechter.State.TimedOut) {
                    throw "Zeitwaechter hat $($type.App) waehrend der VBA-Pruefung beendet"
                }

                if (-not $hasCode) {

                    # ==============================================================================
                    # Zielpfad eindeutig machen, lokal speichern, anschliessend zurueckkopieren
                    # ==============================================================================
                    # Klammern um Muster und Ersetzung sind zwingend: der
                    # Komma-Operator bindet in PowerShell STAERKER als '+'.
                    # Ohne sie liest der Parser
                    #   $filePath -replace ($muster + @('$', '.docx'))
                    # also ein EINARMIGES -replace mit dem Muster
                    # '\.docm$ .docx', das auf nichts passt. Der Zielname
                    # behielt dadurch die Makro-Endung, Get-UniqueTargetPath
                    # machte daraus '<name>_2.docm', und dort landete
                    # anschliessend docx-Inhalt - waehrend das Original
                    # geloescht wurde. Das Skript tat damit das Gegenteil
                    # seiner Aufgabe und meldete es als "KONVERTIERT".
                    $networkTargetBase = $filePath -replace ([regex]::Escape($type.Ext) + '$'), $type.Target
                    $networkTarget     = Get-UniqueTargetPath $networkTargetBase
                    $wasRenamed        = ($networkTarget -ne $networkTargetBase)

                    # Zentrale Freigabe VOR dem SaveAs. Vorher stand die
                    # Pruefung erst nach der vollstaendigen COM-Konvertierung:
                    # der Probelauf liess Office also jede Datei komplett
                    # umwandeln, nur um das Ergebnis zu verwerfen. Ausserdem
                    # war $obj an der alten Stelle bereits freigegeben und auf
                    # $null gesetzt, sodass der Aufraeumcode im Probelauf-Zweig
                    # wirkungslos war. Hier ist das Dokument noch offen und
                    # wird regulaer geschlossen.
                    if (-not (Confirm-Write -Target $filePath `
                              -Action "Nach $($type.Target) umwandeln")) {
                        Write-Log "[PROBELAUF] Wuerde umgewandelt: $fileName -> $(Split-Path $networkTarget -Leaf)" -Level "INFO"
                        $stats.WouldConvert++
                        Write-SkipLogEntry -Source $filePath -Target $networkTarget `
                            -Reason 'PROBELAUF - wuerde umgewandelt, nichts geaendert'
                        Close-OfficeDocument $obj $type.Ext
                        Release-ComObject $obj
                        $obj = $null
                        continue
                    }

                    # Lokales Save-Target unter Kurznamen, damit langer Original-Filename
                    # die 259-Zeichen-Grenze des lokalen Pfades nicht reisst.
                    # [string]: siehe $localCopy - ohne haengt SaveAs2.
                    $localTarget = [string](Join-Path $guidDir "work$($type.Target)")

                    switch ($type.Ext) {
                        ".docm" { $obj.SaveAs2($localTarget, $type.Format, $false, "", $false) }
                        ".xlsm" { $obj.SaveAs($localTarget, $type.Format) }
                        ".pptm" { $obj.SaveAs($localTarget, $type.Format) }
                    }

                    Close-OfficeDocument $obj $type.Ext
                    Release-ComObject $obj
                    $obj = $null
                    # Ab hier nur noch Dateioperationen - ein langsames
                    # Netzlaufwerk darf die Office-Instanz nicht kosten.
                    # Lief die Frist trotzdem gerade ab, ist das Ergebnis
                    # vollstaendig (SaveAs kehrte zurueck); den Neustart
                    # erledigt der finally-Zweig.
                    $null = Clear-DateiWaechter $dateiWaechter

                    if ($null -ne $origTimestamp) {
                        try { [System.IO.File]::SetLastWriteTime($localTarget, $origTimestamp) } catch {}
                    }

                    # NTFS-Sicherheitsinfo (Owner/Group/DACL) des Originals
                    # sichern - wird auf die konvertierte Datei uebertragen.
                    # Langpfad-Fassung, sonst scheitert der Snapshot genau bei
                    # den Dateien, fuer die das Skript den Praefix ueberhaupt
                    # eingefuehrt hat.
                    $secSnapshot = Get-FileSecuritySnapshot $filePathLong

                    $backupPath     = Get-UniqueBackupPath $filePath
                    $backupPathLong = Add-LongPathPrefix $backupPath
                    $networkTargetLong = Add-LongPathPrefix $networkTarget
                    # Muss vor dem try stehen: der Fehlerzweig raeumt darueber
                    # eine liegengebliebene Zwischenkopie ab, und ein Wert aus
                    # dem vorigen Schleifendurchlauf duerfte dort nicht stehen.
                    $stagingPathLong = $null

                    Start-Sleep -Milliseconds 200

                    # Die gesamte Ersetzungssequenz laeuft ueber [System.IO]
                    # statt ueber Rename-Item/Copy-Item/Remove-Item. Das ist
                    # hier besonders wichtig: scheitert ein Schritt in der
                    # Mitte, ist das Original bereits umbenannt - und der
                    # Wiederherstellungspfad benutzte bisher dasselbe
                    # unzuverlaessige Cmdlet, konnte also aus demselben Grund
                    # scheitern.
                    try {
                        Invoke-WithRetry -Action {
                            [System.IO.File]::Move($filePathLong, $backupPathLong)
                        } -MaxAttempts 4 -DelayMs 500

                        try {
                            # Zwischen Get-UniqueTargetPath (oben) und diesem
                            # Schreibvorgang liegt die komplette
                            # COM-Konvertierung, das Wartefenster und das
                            # Umbenennen des Originals - in dieser Zeit kann
                            # ein anderer Lauf oder Anwender genau diesen
                            # Zielnamen belegt haben. Ein Copy mit
                            # overwrite=$true hat die fremde Datei dann
                            # ersatzlos ueberschrieben (nachgestellt:
                            # Fremdinhalt weg).
                            #
                            # Deshalb zweistufig: erst unter einem eigenen,
                            # garantiert freien Namen im Zielverzeichnis
                            # ablegen (dort darf overwrite stehen - die Datei
                            # gehoert uns, und die Wiederholung braucht es
                            # nach einem Teilabbruch), dann per File.Move an
                            # den endgueltigen Platz. Move legt NICHT ueber
                            # eine bestehende Datei, sondern wirft
                            # ERROR_ALREADY_EXISTS (0x800700B7, gemessen) -
                            # das ist die unteilbare Pruefung, die dem
                            # getrennten Exists() fehlt.
                            $stagingPath     = $networkTargetBase + ".tmp_" + [Guid]::NewGuid().ToString("N").Substring(0, 8)
                            $stagingPathLong = Add-LongPathPrefix $stagingPath

                            Invoke-WithRetry -Action {
                                [System.IO.File]::Copy($localTarget, $stagingPathLong, $true)
                            } -MaxAttempts 4 -DelayMs 500

                            # Platz belegen. Bei Kollision einen neuen Namen
                            # ziehen statt zu ueberschreiben. Begrenzt, damit
                            # ein dauerhaft blockiertes Ziel nicht endlos
                            # dreht; der Fehler faellt dann in den
                            # Wiederherstellungszweig unten.
                            $claimed = $false
                            for ($claimTry = 0; $claimTry -lt 20 -and -not $claimed; $claimTry++) {
                                try {
                                    [System.IO.File]::Move($stagingPathLong, $networkTargetLong)
                                    $claimed = $true
                                } catch [System.IO.IOException] {
                                    if (-not ([System.IO.File]::Exists($networkTargetLong) -or
                                              [System.IO.Directory]::Exists($networkTargetLong))) {
                                        throw   # andere Ursache (Sperre, Netz) - nicht als Kollision behandeln
                                    }
                                    $networkTarget     = Get-UniqueTargetPath $networkTargetBase
                                    $networkTargetLong = Add-LongPathPrefix $networkTarget
                                    $wasRenamed        = ($networkTarget -ne $networkTargetBase)
                                    Write-Log ("Zielname war beim Schreiben belegt, weiche aus auf: " +
                                               (Split-Path $networkTarget -Leaf)) -Level "WARN"
                                }
                            }
                            if (-not $claimed) {
                                try { [System.IO.File]::Delete($stagingPathLong) } catch {}
                                throw "Zielname konnte nicht belegt werden: $networkTarget"
                            }

                            if ($null -ne $origTimestamp) {
                                try { [System.IO.File]::SetLastWriteTime($networkTargetLong, $origTimestamp) } catch {}
                            }

                            # ACL/Owner des Originals auf die konvertierte
                            # Datei uebertragen (Admin: vollstaendig; Nutzer: DACL).
                            Set-FileSecuritySnapshot $networkTargetLong $secSnapshot

                            # Ab hier steht das Ziel. Ein Fehler beim Entfernen
                            # des Backups darf deshalb NICHT in die
                            # Wiederherstellung unten laufen: die holte das
                            # Original zurueck und liess die fertige Datei
                            # daneben stehen - jeder weitere Lauf erzeugte dann
                            # '_2', '_3' ... Nachgestellt mit einem
                            # schreibgeschuetzten Original: File.Delete auf das
                            # Backup wirft "Zugriff verweigert". Das Attribut
                            # wird deshalb vorher entfernt und auf das Ziel
                            # uebertragen; bleibt das Backup trotzdem stehen,
                            # wird nur gewarnt (das Aufraeumen spaeterer Laeufe
                            # erfasst es).
                            $bakAttr = [System.IO.FileAttributes]::Normal
                            try { $bakAttr = [System.IO.File]::GetAttributes($backupPathLong) } catch {}
                            if ($bakAttr -band [System.IO.FileAttributes]::ReadOnly) {
                                try {
                                    [System.IO.File]::SetAttributes($backupPathLong, ($bakAttr -band (-bnot [System.IO.FileAttributes]::ReadOnly)))
                                    $zielAttr = [System.IO.File]::GetAttributes($networkTargetLong)
                                    [System.IO.File]::SetAttributes($networkTargetLong, ($zielAttr -bor [System.IO.FileAttributes]::ReadOnly))
                                } catch {}
                            }
                            try {
                                Invoke-WithRetry -Action {
                                    [System.IO.File]::Delete($backupPathLong)
                                } -MaxAttempts 4 -DelayMs 500
                            } catch {
                                Write-Log "Umgewandelt, aber Backup nicht loeschbar - bleibt liegen: $backupPath ($($_.Exception.Message))" -Level "WARN"
                            }

                            if ($wasRenamed) {
                                Write-Log "KONVERTIERT (umbenannt): $fileName -> $(Split-Path $networkTarget -Leaf)" -Level "SUCCESS"
                                Write-SkipLogEntry -Source $filePath -Target $networkTarget `
                                    -Reason 'Zieldatei existierte bereits, neuer Name vergeben'
                                $stats.Renamed++
                            } else {
                                Write-Log "KONVERTIERT: $fileName (Kein Code gefunden)" -Level "SUCCESS"
                            }
                            $stats.Converted++
                        } catch {
                            Write-Log "Rueckkopie fehlgeschlagen fuer $fileName - stelle Original wieder her: $($_.Exception.Message)" -Level "ERROR"
                            $stats.Errors++
                            # Zwischenkopie abraeumen, sonst bleibt eine
                            # '<name>.docx.tmp_<8 Hex>' im Zielverzeichnis liegen.
                            if ($null -ne $stagingPathLong) {
                                try {
                                    if ([System.IO.File]::Exists($stagingPathLong)) {
                                        [System.IO.File]::Delete($stagingPathLong)
                                    }
                                } catch {
                                    Write-Log "Zwischenkopie nicht loeschbar: $stagingPathLong" -Level "WARN"
                                }
                            }
                            try {
                                [System.IO.File]::Move($backupPathLong, $filePathLong)
                            } catch {
                                Write-Log "KRITISCH: Weder Rueckkopie noch Wiederherstellung moeglich fuer $fileName - Backup liegt als $backupPath" -Level "ERROR"
                            }
                        }
                    } catch {
                        Write-Log "Original nicht umbenennbar (Dateisperre?): $fileName : $($_.Exception.Message)" -Level "WARN"
                        $stats.Errors++
                    }
                } else {
                    Close-OfficeDocument $obj $type.Ext
                    Release-ComObject $obj
                    $obj = $null
                    if ($xlmHandled) {
                        # Bereits oben gezaehlt und protokolliert (Excel-4.0-Makro).
                        # Hier nichts weiter tun - sonst doppelte Zaehlung und
                        # eine zweite, inhaltlich falsche CSV-Zeile.
                    } elseif ($noAccessOccurred) {
                        $stats.KeptNoAccess++
                        Write-SkipLogEntry -Source $filePath `
                            -Reason 'Kein VBA-Zugriff (Trust Center) - Makrostatus unbekannt, manuell pruefen'
                    } else {
                        Write-Log "MAKRO GEFUNDEN: $fileName (Wird beibehalten)" -Level "INFO"
                        $stats.KeptMacro++
                        Write-SkipLogEntry -Source $filePath `
                            -Reason 'VBA-Projekt vorhanden - bewusst beibehalten'
                    }
                }
            } else {
                Write-Log "Konnte Datei nicht oeffnen: $fileName" -Level "WARN"
                $stats.Errors++
                Write-SkipLogEntry -Source $filePath -Reason 'Datei liess sich nicht oeffnen'
            }
        } catch {
            $fehlerText = $_.Exception.Message
            if ($dateiWaechter.State.TimedOut) {
                $fehlerText = "TIMEOUT nach $($script:DateiTimeoutSec) s - $($type.App) blockierte (Dialog, Passwort oder defekte Datei); Instanz beendet, Datei unveraendert"
            }
            Write-Log "Fehler bei $fileName : $fehlerText" -Level "ERROR"
            $stats.Errors++
            Write-SkipLogEntry -Source $filePath -Reason ("Fehler: " + $fehlerText)
            if ($null -ne $obj) {
                Close-OfficeDocument $obj $type.Ext
            }
        } finally {
            $abgelaufen = Clear-DateiWaechter $dateiWaechter -Quittieren
            Release-ComObject $obj
            $obj = $null
            if ($abgelaufen) {
                Write-Log "Zeitwaechter hat $($type.App) beendet - starte neu..." -Level "WARN"
                Restart-OfficeFuerTyp $type.Ext
                $processedInBatch = 0
            }

            if (Test-Path -LiteralPath $guidDir) {
                Remove-Item -LiteralPath $guidDir -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false
            }
        }
    }
}

} catch {
    $abortedByException = $true
    Write-Log "ABBRUCH durch Exception: $($_.Exception.Message)" -Level "ERROR"
    try { Write-Log $_.ScriptStackTrace -Level "ERROR" } catch {}
} finally {
    # ==============================================================================
    # Cleanup (laeuft auch bei Exception/Abbruch garantiert)
    # ==============================================================================
    Write-Log "Beende Prozesse..."
    try { Stop-DateiWaechter $dateiWaechter } catch {}
    try { Stop-OfficeApp $word $officePids.Word } catch {}
    try { Stop-OfficeApp $excel $officePids.Excel } catch {}
    try { Stop-OfficeApp $pptx $officePids.PowerPoint } catch {}

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 500

    try { Stop-TrackedOfficeProcess $officePids.Word $officeStartTimes.Word       } catch {}
    try { Stop-TrackedOfficeProcess $officePids.Excel $officeStartTimes.Excel      } catch {}
    try { Stop-TrackedOfficeProcess $officePids.PowerPoint $officeStartTimes.PowerPoint } catch {}

    try {
        if ((Get-ChildItem $tempDir -ErrorAction SilentlyContinue | Measure-Object).Count -eq 0) {
            Remove-Item -LiteralPath $tempDir -Force -ErrorAction SilentlyContinue -WhatIf:$false
        }
    } catch {}

    # Fortschrittsbalken am Ende ausblenden, falls aktiv.
    if ($useProgress) {
        try { Write-Progress -Activity "Office Macro Deep-Converter" -Completed } catch {}
    }
}

# ==============================================================================
# Verwaiste Backups frueherer Laeufe aufraeumen
# ==============================================================================
# Vor dem Ersetzen benennt das Skript das Original in '<datei>.bak_<8 Hex>' um
# und loescht es nach erfolgreicher Umwandlung. Nach einem harten Abbruch
# (Stromausfall, Task-Manager) bleiben diese Reste liegen. Geloescht wird nur,
# wenn alle drei Bedingungen erfuellt sind:
#   1. Name entspricht exakt dem eigenen Muster '<office-datei>.bak_<8 Hex>'
#   2. eine gleichnamige Zieldatei existiert und ist nicht leer
#   3. das Backup ist aelter als $script:BackupCleanupMinAgeHours
# Backups ohne intakte Zieldatei werden gemeldet, aber nie geloescht.
if (-not $SkipBackupCleanup -and -not $cancelled -and -not $abortedByException) {
    Write-Log "Suche zurueckgebliebene Backups frueherer Laeufe ..."
    $bakDeleted = 0
    $bakKept    = 0
    $bakOrphans = 0
    $bakCutoff  = (Get-Date).AddHours(-$script:BackupCleanupMinAgeHours)
    try {
        foreach ($bak in (Get-FilesRecursiveSafe -RootPath $longPath -Pattern "*.bak_*")) {
            try {
                $bakPlain = Remove-LongPathPrefix $bak
                if ($bakPlain -notmatch '\.bak_[0-9a-f]{8}$') { continue }
                $orig = $bakPlain -replace '\.bak_[0-9a-f]{8}$', ''
                # Nur eigene Office-Backups anfassen - fremde .bak_-Dateien bleiben unberuehrt.
                if ($orig -notmatch '\.(docm|xlsm|pptm|doc|xls|ppt)$') { continue }

                $fi = New-Object System.IO.FileInfo (Add-LongPathPrefix $bakPlain)
                if (-not $fi.Exists) { continue }

                # Zieldatei: makrofreie Variante oder das unveraenderte Original.
                $kandidaten = @($orig,
                                ($orig -replace '\.docm$', '.docx'),
                                ($orig -replace '\.xlsm$', '.xlsx'),
                                ($orig -replace '\.pptm$', '.pptx')) | Select-Object -Unique
                $zielOk = $false
                foreach ($k in $kandidaten) {
                    try {
                        $kfi = New-Object System.IO.FileInfo (Add-LongPathPrefix $k)
                        if ($kfi.Exists -and $kfi.Length -gt 0) { $zielOk = $true; break }
                    } catch { }
                }
                if (-not $zielOk) {
                    $bakOrphans++
                    Write-Log "Backup ohne intakte Zieldatei - NICHT geloescht: $bakPlain" -Level "WARN"
                    continue
                }
                if ($fi.LastWriteTime -gt $bakCutoff) { $bakKept++; continue }

                if (Confirm-Write -Target $bakPlain -Action 'Verwaistes Backup loeschen') {
                    try {
                        [System.IO.File]::Delete($fi.FullName)
                        $bakDeleted++
                    } catch {
                        $bakKept++
                        Write-Log "Backup nicht loeschbar: $bakPlain - $($_.Exception.Message)" -Level "WARN"
                    }
                } else {
                    $bakKept++
                }
            } catch { }
        }
        Write-Log ("Backups: {0} entfernt, {1} behalten, {2} ohne Zieldatei" -f $bakDeleted, $bakKept, $bakOrphans)
    } catch {
        Write-Log "Backup-Aufraeumen uebersprungen: $($_.Exception.Message)" -Level "WARN"
    }
}

# ==============================================================================
# Statistik-Ausgabe
# ==============================================================================
if ($abortedByException) {
    Write-Log "--- ZUSAMMENFASSUNG (DURCH FEHLER ABGEBROCHEN) ---" -Level "ERROR"
} elseif ($cancelled) {
    Write-Log "--- ZUSAMMENFASSUNG (DURCH ANWENDER ABGEBROCHEN) ---" -Level "WARN"
} else {
    Write-Log "--- ZUSAMMENFASSUNG ---"
}
Write-Log ("Gescannt:                          {0}" -f $stats.Scanned)
Write-Log ("Konvertiert:                       {0}" -f $stats.Converted)
Write-Log ("  davon umbenannt (_2, _3, ...):   {0}" -f $stats.Renamed)
Write-Log ("Beibehalten (VBA-Projekt):          {0}" -f $stats.KeptMacro)
# Eigene Zeilen statt "davon": XLM-Dateien werden seit der Korrektur
# NICHT mehr zusaetzlich als VBA-Makro gezaehlt, sind also keine
# Teilmenge von KeptMacro mehr.
if ($stats.KeptXlm -gt 0) {
    Write-Log ("Beibehalten (Excel-4.0-Makro/XLM): {0}" -f $stats.KeptXlm) -Level "WARN"
}
if ($stats.KeptXlmUnknown -gt 0) {
    Write-Log ("Beibehalten (XLM nicht pruefbar):  {0}" -f $stats.KeptXlmUnknown) -Level "WARN"
}
if ($stats.WouldConvert -gt 0) {
    Write-Log ("Wuerde umwandeln (Probelauf):       {0}" -f $stats.WouldConvert) -Level "WARN"
}
Write-Log ("Beibehalten (kein VBA-Zugriff):    {0}" -f $stats.KeptNoAccess)
$stats.SkippedLong = $script:SkippedPathTooLong
Write-Log ("Uebersprungen (Pfad zu lang):      {0}" -f $stats.SkippedLong)
Write-Log ("Uebersprungen (Kennwort):          {0}" -f $stats.SkippedPassword)
Write-Log ("Fehler:                            {0}" -f $stats.Errors)
Write-Log "Logdateien:"
Write-Log "  $logFile"
Write-Log "  $skipLog"
Write-Log "--- SKRIPT BEENDET ---"

if ($script:GemeinsamGeladen) {
    try { [void](Remove-EigeneRecentEintraege) } catch { }
}
Invoke-WindowsTempCleanup

try { if ($script:CtrlCAsInput) { [Console]::TreatControlCAsInput = $false } } catch {}

if (-not $Automated) {
    Write-Host ""
    Write-Host "Druecken Sie eine beliebige Taste zum Beenden..." -ForegroundColor Cyan
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

if ($abortedByException -and $Automated) {
    exit 1
}
