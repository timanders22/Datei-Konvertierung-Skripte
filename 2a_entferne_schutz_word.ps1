<#
.SYNOPSIS
    Word Deep Unprotect
.DESCRIPTION
    Entfernt Dokumentschutz, Schreibschutz und Abschnittsschutz aus Word-Dateien.
    Konvertiert alte .doc/.dot-Formate nach .docx/.dotx.
    Unterstützt Netzlaufwerke und Pfade > 260 Zeichen.

    Behandelte Schutzarten (XML-Ebene, ohne Passwort):
      word/settings.xml : documentProtection, writeProtection
      word/document.xml : formProt (Abschnitte), permStart/permEnd, fieldLock,
                          lock (Content-Control-Sperren in sdtPr)

    Zone.Identifier (Alternate Data Stream):
      Wird vor der COM-Verarbeitung und nach dem Rückschreiben entfernt.
      Verhindert, dass Word in der "Geschützten Ansicht" (Protected View)
      startet, was die COM-Automatisierung zum Einfrieren bringt.

    Temporärer Ordner:
      Das Skript nutzt den Dokumente-Ordner des Benutzers als temporären
      Arbeitsbereich (nicht %TEMP%, da Word diesen Pfad nicht als
      "Vertrauenswürdigen Speicherort" akzeptiert). Der Ordner muss in
      Word unter Datei > Optionen > Trust Center > Vertrauenswürdige
      Speicherorte eingetragen sein (inkl. Unterordner). Bei Netzwerk-
      umgeleiteten Dokumente-Ordnern zusätzlich die Option
      "Vertrauenswürdige Speicherorte im Netzwerk zulassen" aktivieren.
      Das Skript führt vor dem Hauptlauf einen Trust-Center-Smoke-Test durch
      und bricht bei Fehlschlag ab (bzw. fragt interaktiv nach), um stunden-
      lange Timeouts pro Datei zu vermeiden.

    MRU (Most Recently Used):
      Documents.Open und SaveAs2 setzen AddToRecentFiles=$false,
      damit verarbeitete Dateien nicht in der MRU-Liste erscheinen.

    AV-Resilienz (Retry + Wait):
      I/O-Operationen sind in einen Retry-Wrapper mit exponentiellem
      Backoff gekapselt. Nach dem initialen Copy in den Arbeitsordner
      wird zusätzlich auf exklusiven Dateizugriff gewartet, um AV-Scanner-
      Sperren zu überbrücken. Erfolgreiche Zweit-/Drittversuche werden
      in der Endstatistik ausgewiesen.

    Nicht entfernbar ohne Passwort:
      - Verschlüsselte Dokumente (Open-Passwort) → werden übersprungen
      - VBA-Projekt-Passwort in .docm/.dotm (binäre vbaProject.bin)

    Voraussetzung für Aufgabenplanung / Headless-Betrieb (Dienstkonto):
      Word benötigt einen existierenden Desktop-Ordner im Profilpfad des
      ausführenden Kontos. Fehlen diese Ordner, schlägt die COM-Initialisierung
      fehl. Vor dem ersten Einsatz als Administrator anlegen:
        C:\Windows\SysWOW64\config\systemprofile\Desktop
        C:\Windows\System32\config\systemprofile\Desktop

    Bekanntes Infrastruktur-Verhalten (kein Fehler im Skript):
      Bei DFS-Umleitungen auf offline/abgestürzte Server kann die Dateisuche
      an einzelnen Ordnern einfrieren, bis Windows den Netzwerk-Timeout auslöst.
      Das Skript setzt danach automatisch fort.

    Compilieren in eine .exe (Windows PowerShell 5.1, STA-Konsolenanwendung):
      # Einmalig: ps2exe installieren
      Install-Module -Name ps2exe -Scope CurrentUser -Force

      # Hilfscheck: Office-Bitness ermitteln (optional zur Optimierung)
      #   Rueckgabe 'x64' -> -x64 verwenden ; 'x86' -> -x86 verwenden
      Get-ItemProperty 'HKLM:\Software\Microsoft\Office\ClickToRun\Configuration' -Name Platform -ErrorAction SilentlyContinue

      # Hinweis zu -x64 / -x86 bei gemischten Office-Bitness-Umgebungen:
      # Word registriert seinen COM-Server als LocalServer32 (out-of-process,
      # WINWORD.EXE). Das Windows-COM-Subsystem marshallt Aufrufe zwischen
      # x64-Aufrufer und x86-Server (und umgekehrt) automatisch ueber DCOM/LRPC.
      # Daher funktioniert eine mit -x64 kompilierte EXE auch mit 32-Bit-Word
      # (und eine -x86-EXE mit 64-Bit-Word). Bitness-Match liefert minimal
      # bessere Performance, ist aber nicht erforderlich. -x64 ist eine sichere
      # Default-Wahl, da modernes Office (2019+, M365) standardmaessig x64 ist.

      Invoke-ps2exe -inputFile   '.\2a_entferne_schutz_word.ps1' -outputFile  '.\2a_entferne_schutz_word.exe' -iconFile    '.\powershell_icon.ico' -title       'Word Deep Unprotect' -description 'Entfernt Schutz aus Word-Dateien' -product     'Word Deep Unprotect' -version     '1.0.1.0' -STA -x64 -supportOS

    Simulation:
      -WhatIf zeigt an, welche Dateien geaendert wuerden, ohne sie
      anzufassen. Empfohlen fuer den ersten Lauf auf einer neuen Ablage.

    Stand: 10.06.2026
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Position=0, Mandatory=$false)]
    [string]$TargetPath,

    [Parameter(Mandatory=$false)]
    [switch]$ShowProgress,

    [Parameter(Mandatory=$false)]
    [switch]$NoInteractive,

    # Ueberspringt den Backup-Aufraeumlauf. Auf sehr grossen Ablagen spart das
    # einen kompletten zusaetzlichen Verzeichnisdurchlauf.
    [Parameter(Mandatory=$false)]
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



# ==================================================================
# Gemeinsame Grundbibliothek (mit Rueckfall)
# ==================================================================
# Bindet _gemeinsam.psm1 ein, wenn vorhanden. Die eingebauten Kopien der
# Helfer bleiben bestehen - so bleibt jedes Skript einzeln lauffaehig und
# die ps2exe-Uebersetzung funktioniert unveraendert. Genutzt wird das
# Modul fuer das gemeinsame Laufprotokoll (migration.jsonl).
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


# ==================================================================
# KONFIGURATION
# ==================================================================
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

$scriptDir              = Resolve-ScriptDirectory

# Log-Verzeichnis: bevorzugt neben dem Skript, sonst LOCALAPPDATA.
# Vorher wurde ungeprueft in $scriptDir geschrieben. Liegt das Skript auf einer
# schreibgeschuetzten Freigabe oder unter Program Files, schlugen ALLE
# Log-Schreibvorgaenge still fehl (Add-Content mit -ErrorAction
# SilentlyContinue) - der Lauf lief ohne jede Protokollierung durch.
function Resolve-LogDirectory {
    param([string]$Preferred)
    foreach ($cand in @($Preferred,
                        (Join-Path $env:LOCALAPPDATA 'WordDeepUnprotect'),
                        $env:TEMP)) {
        if ([string]::IsNullOrWhiteSpace($cand)) { continue }
        try {
            if (-not [System.IO.Directory]::Exists($cand)) {
                New-Item -ItemType Directory -Path $cand -Force -ErrorAction Stop -WhatIf:$false | Out-Null
            }
            $probe = Join-Path $cand (".writetest_{0}" -f ([Guid]::NewGuid().ToString('N')))
            [System.IO.File]::WriteAllText($probe, 'x')
            [System.IO.File]::Delete($probe)
            return $cand
        } catch { continue }
    }
    return $Preferred
}

$LogDir                 = Resolve-LogDirectory -Preferred $scriptDir
$TempPath               = Join-Path ([Environment]::GetFolderPath('MyDocuments')) "2a_entferne_schutz_word_$PID"
$RunTimestamp           = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$LogFilePath            = Join-Path $LogDir "2a_entferne_schutz_word_$RunTimestamp.log"
$DetailedLogPath        = Join-Path $LogDir "2a_entferne_schutz_word_detailed_$RunTimestamp.log"
$CsvLogPath             = Join-Path $LogDir "2a_entferne_schutz_word_$RunTimestamp.csv"

# Mindestalter fuer das Loeschen von '~$'-Sperrdateien. Word legt sie an,
# solange ein Dokument GEOEFFNET ist - sofortiges Loeschen bricht die Sperre
# fuer andere Nutzer auf der Freigabe. Einheitlich zum Temp-Cleaner.
$JunkMinAgeHours        = 24

# Maximale Anzahl Ausweichnamen (Datei_2.docx ... Datei_100.docx), bevor
# zum Schutz fremder Dateien abgebrochen wird.
$MaxNameClashRetries    = 100

# Mindestalter, ab dem ein zurueckgebliebenes '.bak' als verwaist gilt.
# Grosszuegig bemessen, damit ein parallel laufender Durchgang eines anderen
# Nutzers auf derselben Ablage nicht seine frischen Backups verliert.
$BackupCleanupMinAgeHours = 24

# Verzeichnisse, die bei der Suche nicht betreten werden. Ohne diese Liste
# landeten auch Dokumente im Papierkorb und in Schattenkopie-Verzeichnissen
# in der Verarbeitung.
$script:ExcludeDirNames = @(
    '$RECYCLE.BIN', 'System Volume Information', 'RECYCLER',
    '.git', '.svn', '__pycache__', 'node_modules'
)

# Verzeichnis-Presets fuer die Startauswahl. Hier die im eigenen Netz
# gebraeuchlichen Laufwerke/Shares eintragen.
$DirectoryPresets = @(
    'Q:\'
    'R:\'
    'G:\Geteilte Ablagen'
    'G:\Meine Ablage'
    '\\server\dfs'
)
$FileOpenTimeoutSeconds = 45
$TrustCenterTimeoutSec  = 25
# Retry-Parameter sind so kalibriert, dass die Gesamtwartezeit ueber alle
# Versuche hinweg ca. 14 Sekunden umfasst (200+400+800+1600+3200+4000+4000ms).
# Das deckt typische AV-Scanner-Lockups nach Copy/Unblock-File ab, die in der
# Praxis 5-10 Sekunden dauern koennen (Defender mit Cloud-Scan).
$RetryMaxAttempts       = 7
$RetryInitialDelayMs    = 200
$RetryMaxDelayMs        = 4000
$PostComStabilizeMs     = 150
$PidCleanupInterval     = 50
# Vor-Open-Wait nach Copy+ReadOnly-Reset+Unblock-File. Hilft, dass ZipFile.Open
# nicht direkt in einen AV-Scan laeuft.
$WaitFileAvailableMs    = 8000
$script:Passwords       = [System.Collections.Generic.List[string]]::new()
$ScriptStartTime        = Get-Date

# Word SaveAs2 Format-Konstanten (WdSaveFormat)
$WD_FORMAT_DOCX = 12
$WD_FORMAT_DOCM = 13
$WD_FORMAT_DOTX = 14
$WD_FORMAT_DOTM = 15

# ==================================================================
# INIT & LIBRARIES
# ==================================================================
try {
    Add-Type -AssemblyName "System.IO.Compression.FileSystem" -ErrorAction Stop
} catch {
    try { Add-Type -AssemblyName "System.IO.Compression" -ErrorAction SilentlyContinue } catch {}
}

$script:ShouldStop      = $false
$script:UseProgress     = $false
$script:SkipPreScan     = $false
$script:TotalFiles      = 0
$script:ProcessedCount  = 0
$script:TrackedWordPids = [System.Collections.Generic.List[int]]::new()
# Startzeit je verfolgter PID (PID -> StartTime), siehe Add-TrackedWordPid.
$script:TrackedWordStart = @{}
$script:RetryStats      = @{ Recovered = 0; Failed = 0 }
$script:RestorePrivilegesEnabled = $false
$script:LastProgressUpdate    = [DateTime]::MinValue
$script:ProgressMinIntervalMs = 500

# $PSCmdlet ist nur im Skript-Scope verfuegbar, nicht in einfachen Funktionen.
# Referenz merken, damit Confirm-Write ShouldProcess aufrufen kann.
$script:ScriptCmdlet = $PSCmdlet
$script:DryRun       = $false

function Get-RunningOfficeSessions {
    <#
        Liefert Office-Prozesse mit sichtbarem Hauptfenster, also echte
        Sitzungen des Anwenders - im Unterschied zu unsichtbaren
        Automatisierungs-Instanzen.
    #>
    param([string[]]$ProcessNames = @('WINWORD'))
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

        Word wird per COM automatisiert. Laeuft bereits eine Sitzung des
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
        Zentrale Freigabe fuer jede schreibende Operation.

        Das Skript deklarierte bisher SupportsShouldProcess, rief ShouldProcess
        aber an KEINER Stelle auf. -WhatIf wurde damit stillschweigend
        angenommen und die Dateien trotzdem geaendert - wer einen Testlauf
        erwartete, bekam einen echten Lauf.
    #>
    param(
        [string]$Target,
        [string]$Action = 'Ändern'
    )
    if ($script:DryRun) { return $false }
    if ($null -eq $script:ScriptCmdlet) { return $true }
    try {
        return $script:ScriptCmdlet.ShouldProcess($Target, $Action)
    } catch {
        return $true
    }
}

function Test-IsOwnWordProcess {
    <#
        Prueft, ob eine PID wirklich zu einer von UNS gestarteten Word-Instanz
        gehoert. Windows vergibt PIDs wieder; ohne den Startzeit-Vergleich
        koennte Cleanup-AllWord eine zwischenzeitlich vom Nutzer geoeffnete
        Word-Sitzung abschiessen, die zufaellig dieselbe PID bekommen hat.

        Verglichen wird EXAKT mit der Startzeit, die Add-TrackedWordPid beim
        Aufnehmen der PID gemerkt hat (wie Stop-TrackedOfficeProcess
        -StartTime in _gemeinsam.psm1). Vorher genuegte 'StartTime nach
        Skriptstart': eine Word-Sitzung, die der Anwender waehrend des Laufs
        oeffnet und die eine frei gewordene, noch in der Liste stehende PID
        erbt, bestand diese Pruefung ebenfalls und wurde von Cleanup-AllWord
        per Stop-Process -Force beendet (samt ungespeicherter Dokumente).
        Nicht gemerkte PIDs gelten nie als eigene.
    #>
    param([int]$ProcessId)
    if (-not $script:TrackedWordStart.ContainsKey($ProcessId)) { return $false }
    try {
        $proc = Get-Process -Id $ProcessId -ErrorAction Stop
        if ($proc.Name -ne 'WINWORD') { return $false }
        try {
            if ($proc.StartTime -ne $script:TrackedWordStart[$ProcessId]) { return $false }
        } catch {
            # StartTime nicht lesbar (Rechte) - im Zweifel nicht anfassen
            return $false
        }
        return $true
    } catch {
        return $false
    }
}

function Add-TrackedWordPid {
    <#
        Nimmt eine PID in die Aufraeumliste auf und merkt sich ihre Startzeit.
        Aufgenommen wird nur ein laufender WINWORD-Prozess, der nach dem
        Skriptstart gestartet wurde; eine bereits beendete PID kommt gar nicht
        erst auf die Liste (sonst koennte sie spaeter wiedervergeben werden).
    #>
    param([int]$ProcessId)
    if ($ProcessId -le 0) { return }
    try {
        $proc = Get-Process -Id $ProcessId -ErrorAction Stop
        if ($proc.Name -ne 'WINWORD') { return }
        $start = $proc.StartTime
        if ($start -lt $ScriptStartTime) { return }
    } catch {
        return
    }
    $script:TrackedWordStart[$ProcessId] = $start
    if (-not $script:TrackedWordPids.Contains($ProcessId)) { $script:TrackedWordPids.Add($ProcessId) }
}

# ==================================================================
# LOGGING & PROZESS-CLEANUP (Trap-relevant)
# ==================================================================

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    # -WhatIf:$false: Protokollieren ist keine simulierbare Fachaktion. Ohne
    # das schweigt unter -WhatIf das gesamte Log, und der Probelauf laesst
    # sich hinterher nicht nachlesen.
    Add-Content -LiteralPath $LogFilePath -Value $entry -Encoding utf8 -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false
}

function Write-DetailedLog {
    param([string]$Message, [string]$Level = "DEBUG")
    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff')] [$Level] $Message"
    Add-Content -LiteralPath $DetailedLogPath -Value $entry -Encoding utf8 -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false
}

function Write-CsvLog {
    <#
        Eine Zeile je verarbeiteter Datei. Semikolon als Trennzeichen und
        UTF-8 mit BOM, damit Excel die Datei im deutschen Gebietsschema
        direkt per Doppelklick korrekt oeffnet.

        Bewusst Add-Content wie bei den uebrigen Logs statt eines
        persistenten StreamWriters: 2a hat acht exit-Pfade, ein nicht
        geschlossener Writer wuerde dort die letzten Zeilen verlieren.
    #>
    param(
        [string]$Status,
        [string]$Actions,
        [string]$Path,
        [string]$Details = ''
    )

    # Gemeinsames Laufprotokoll (migration.jsonl) - ergaenzt das
    # skripteigene Protokoll, ersetzt es nicht. Erst damit laesst sich
    # der Fortschritt ueber alle elf Schritte hinweg auswerten.
    if ($script:GemeinsamGeladen) {
        try {
            Write-Laufprotokoll -Skript '2a_entferne_schutz_word' `
                -Pfad $Path -Aktion 'Schutz entfernen' `
                -Status $Status -Detail $Details
        } catch { }
    }
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $q  = {
        param($v)
        '"' + ((("$v") -replace '"', '""') -replace "`r`n|`r|`n", ' ') + '"'
    }
    $line = "$ts;$Status;$(& $q $Actions);$(& $q $Path);$(& $q $Details)"
    Add-Content -LiteralPath $CsvLogPath -Value $line -Encoding utf8 -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false
}

function Cleanup-AllWord {
    $trackedIds = @($script:TrackedWordPids | Select-Object -Unique)
    foreach ($id in $trackedIds) {
        if (Test-IsOwnWordProcess -ProcessId $id) {
            try { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue -WhatIf:$false } catch {}
        }
    }
    [System.GC]::Collect()
    $script:TrackedWordPids.Clear()
    $script:TrackedWordStart.Clear()
}

function Sweep-DeadPids {
    if ($script:TrackedWordPids.Count -eq 0) { return }
    $alive = [System.Collections.Generic.List[int]]::new()
    foreach ($id in @($script:TrackedWordPids)) {
        if (Test-IsOwnWordProcess -ProcessId $id) { $alive.Add($id) }
    }
    $script:TrackedWordPids.Clear()
    foreach ($id in $alive) { $script:TrackedWordPids.Add($id) }
    # Gemerkte Startzeiten toter PIDs mit entfernen.
    foreach ($id in @($script:TrackedWordStart.Keys)) {
        if (-not $alive.Contains([int]$id)) { $script:TrackedWordStart.Remove($id) }
    }
}

# ==================================================================
# ABBRUCH-HANDLER & TRAP
# ==================================================================
# Hier stand ein [Console]::add_CancelKeyPress-Handler als ScriptBlock. Der
# Handler laeuft auf einem Threadpool-Thread ohne PowerShell-Runspace: Strg+C
# setzte nicht $script:ShouldStop, sondern warf dort eine
# PSInvalidOperationException (ScriptBlock.GetContextFromTLS) und riss
# powershell.exe hart herunter - mitten in einer Datei-Ersetzung
# (nachgestellt per GenerateConsoleCtrlEvent, WER-Bericht). Jetzt wie in 7/9:
# Strg+C als Eingabe behandeln und zusammen mit ESC in der Hauptschleife
# abfragen. Eingeschaltet wird das erst unmittelbar vor der Hauptschleife,
# damit Strg+C in den Eingabeaufforderungen davor wie gewohnt abbricht.
$script:CtrlCAsInput = $false

trap {
    try { if ($script:CtrlCAsInput) { [Console]::TreatControlCAsInput = $false } } catch {}
    Write-Warning "Unerwarteter Fehler: $_"
    Start-Sleep -Milliseconds 200
    if (Test-Path -LiteralPath $TempPath) {
        Get-ChildItem -LiteralPath $TempPath -Filter "*.pid" -ErrorAction SilentlyContinue |
            ForEach-Object {
                try {
                    $p = [int][System.IO.File]::ReadAllText($_.FullName).Trim()
                    if ($p -gt 0) { Add-TrackedWordPid -ProcessId $p }
                } catch {}
            }
    }
    Cleanup-AllWord
    if (Test-Path -LiteralPath $TempPath) { Remove-Item -LiteralPath $TempPath -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false }
    exit 1
}

# ==================================================================
# PFAD-HELPER
# ==================================================================

function Add-LongPathPrefix {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
    if ($Path -match "^\\\\\?\\")            { return $Path }
    if ($Path -match "^\\\\")                { return "\\?\UNC" + $Path.Substring(1) }
    if ($Path -match "^[a-zA-Z]:")           { return "\\?\$Path" }
    return $Path
}

function Remove-LongPathPrefix {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
    if ($Path.StartsWith('\\?\UNC\')) { return '\\' + $Path.Substring(8) }
    if ($Path.StartsWith('\\?\'))      { return $Path.Substring(4) }
    return $Path
}

function Test-PathLong {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    $lp = Add-LongPathPrefix $Path
    return ([System.IO.Directory]::Exists($lp) -or [System.IO.File]::Exists($lp))
}

function Get-UserShellFolder {
    param(
        [Parameter(Mandatory)]
        [string]$Name,
        [string]$Fallback
    )
    try {
        $key = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
        $val = (Get-ItemProperty -LiteralPath $key -Name $Name -ErrorAction Stop).$Name
        if (-not [string]::IsNullOrWhiteSpace($val)) {
            $expanded = [Environment]::ExpandEnvironmentVariables($val)
            if ([System.IO.Directory]::Exists($expanded)) {
                return $expanded
            }
        }
    } catch {}
    return $Fallback
}

# ==================================================================
# I/O-RESILIENZ
# ==================================================================

function Invoke-WithRetry {
    param(
        [scriptblock]$Action,
        [int]$MaxAttempts   = $RetryMaxAttempts,
        [int]$InitialDelay  = $RetryInitialDelayMs,
        [int]$MaxDelay      = $RetryMaxDelayMs,
        [string]$Context    = ""
    )
    $attempt = 0
    $delay   = $InitialDelay
    while ($true) {
        $attempt++
        try {
            $result = & $Action
            if ($attempt -gt 1) {
                $script:RetryStats.Recovered++
                Write-DetailedLog "Retry-Erfolg bei Versuch $attempt [$Context]" "INFO"
            }
            return $result
        } catch {
            if ($attempt -ge $MaxAttempts) {
                $script:RetryStats.Failed++
                Write-DetailedLog "Retry aufgegeben nach $attempt Versuchen [$Context]: $_" "WARN"
                throw
            }
            Write-DetailedLog "Retry $attempt/$MaxAttempts nach ${delay}ms [$Context]: $($_.Exception.Message)" "DEBUG"
            Start-Sleep -Milliseconds $delay
            $delay = [Math]::Min($delay * 2, $MaxDelay)
        }
    }
}

function Wait-FileAvailable {
    param(
        [string]$Path,
        [int]$TimeoutMs = $WaitFileAvailableMs
    )
    $deadline = (Get-Date).AddMilliseconds($TimeoutMs)
    $delay    = $RetryInitialDelayMs
    while ((Get-Date) -lt $deadline) {
        try {
            $fi = New-Object System.IO.FileInfo $Path
            if (-not $fi.Exists) { return $false }
            $stream = $fi.Open([System.IO.FileMode]::Open,
                               [System.IO.FileAccess]::ReadWrite,
                               [System.IO.FileShare]::None)
            $stream.Close()
            return $true
        } catch {
            Start-Sleep -Milliseconds $delay
            $delay = [Math]::Min($delay * 2, $RetryMaxDelayMs)
        }
    }
    return $false
}

function Test-IsValidZip {
    param([string]$FilePath)
    $attempt = 0
    $delay   = $RetryInitialDelayMs
    while ($true) {
        $attempt++
        try {
            $z = [System.IO.Compression.ZipFile]::OpenRead($FilePath)
            $z.Dispose()
            if ($attempt -gt 1) { $script:RetryStats.Recovered++ }
            return $true
        } catch [System.IO.IOException] {
            if ($attempt -ge $RetryMaxAttempts) { return $false }
            Start-Sleep -Milliseconds $delay
            $delay = [Math]::Min($delay * 2, $RetryMaxDelayMs)
        } catch {
            return $false
        }
    }
}

function Test-IsOleEncrypted {
    param([string]$FilePath)
    # Nur die ersten max. 8192 Bytes per Stream lesen statt die komplette
    # Datei via ReadAllBytes in den RAM zu laden. Bei alten .doc-Dateien mit
    # eingebetteten Grafiken (100+ MB sind nicht selten) sparen wir damit
    # massiv RAM und vermeiden OutOfMemoryException bei parallelen Workern.
    $stream = $null
    try {
        $stream = [System.IO.File]::OpenRead($FilePath)
        $buffer = New-Object byte[] 8192
        $bytesRead = $stream.Read($buffer, 0, 8192)
        if ($bytesRead -lt 8) { return $false }
        $content = [System.Text.Encoding]::ASCII.GetString($buffer, 0, $bytesRead)
        return ($content -match "EncryptedPackage" -or $content -match "EncryptionInfo")
    } catch {
        Write-DetailedLog "OLE-Header-Check fehlgeschlagen: $FilePath - $_" "WARN"
        return $false
    } finally {
        if ($null -ne $stream) {
            try { $stream.Dispose() } catch {}
        }
    }
}

function Test-FileIsLocked {
    param([string]$FilePath)
    $lp = Add-LongPathPrefix $FilePath
    try {
        $fi = New-Object System.IO.FileInfo $lp
        if (-not $fi.Exists) { return $false }
        # Das ReadOnly-Attribut ist keine Sperre. Ein Oeffnen mit ReadWrite
        # scheitert daran mit UnauthorizedAccessException, und die Datei
        # wurde als 'Zugriff verweigert' uebersprungen - im Probelauf immer
        # (dort wird das Attribut nicht entfernt), statt 'Wuerde aendern' zu
        # melden (nachgestellt). Fuer schreibgeschuetzte Dateien daher nur
        # lesend oeffnen; FileShare.None erkennt fremde Handles weiterhin.
        $zugriff = if ($fi.IsReadOnly) { [System.IO.FileAccess]::Read } else { [System.IO.FileAccess]::ReadWrite }
        $stream = $fi.Open([System.IO.FileMode]::Open,
                           $zugriff,
                           [System.IO.FileShare]::None)
        $stream.Close()
        return $false
    } catch [System.UnauthorizedAccessException] {
        return "AccessDenied"
    } catch {
        return $true
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
        if (-not ('WordUnprotect.PrivilegeHelper' -as [type])) {
            Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace WordUnprotect {
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
        $okRestore = [WordUnprotect.PrivilegeHelper]::Enable('SeRestorePrivilege')
        $okOwner   = [WordUnprotect.PrivilegeHelper]::Enable('SeTakeOwnershipPrivilege')
        $script:RestorePrivilegesEnabled = ($okRestore -or $okOwner)
        if ($script:RestorePrivilegesEnabled) {
            Write-DetailedLog "Restore-Privilegien aktiviert (Admin-Kontext) - Owner-Wiederherstellung vollstaendig verfuegbar." "DEBUG"
        } else {
            Write-DetailedLog "Restore-Privilegien nicht zugewiesen (Nutzer-Kontext) - Owner-Restore nur auf eigene Dateien moeglich." "DEBUG"
        }
    } catch {
        Write-DetailedLog "Privileg-Aktivierung fehlgeschlagen: $_" "DEBUG"
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
        # Bewusst WARN statt DEBUG: Schlaegt das Lesen fehl, werden Rechte und
        # Eigentuemer der Datei nach der Ersetzung NICHT wiederhergestellt.
        # Als DEBUG-Zeile ging dieser Rechteverlust im Detail-Log unter.
        # Gleichlautend in 2a, 2b und 2c.
        Write-Log "Sicherheitsinfo nicht lesbar - Rechte gehen bei der Ersetzung verloren ($Path): $_" "WARN"
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
        Write-DetailedLog "DACL wiederhergestellt: $Path" "DEBUG"
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
        Write-DetailedLog "Owner wiederhergestellt: $Path" "DEBUG"
    } catch {
        if ($script:RestorePrivilegesEnabled) {
            Write-Log "Owner-Wiederherstellung fehlgeschlagen ($Path): $_" "WARN"
        } else {
            Write-DetailedLog "Owner nicht gesetzt (kein Admin-Privileg - im Nutzer-Kontext unkritisch): $Path" "DEBUG"
        }
    }
}

# ==================================================================
# DATEI-ENUMERATION & FORTSCHRITT
# ==================================================================

function Get-WordFilesRobust {
    param([string]$RootPath, [string]$ExtPattern, [switch]$StringsOnly)
    $queue = [System.Collections.Queue]::new()
    $queue.Enqueue($RootPath)
    while ($queue.Count -gt 0) {
        $dir = $queue.Dequeue()
        try {
            $dirInfo = [System.IO.DirectoryInfo]::new($dir)
            if ($dirInfo.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { continue }
            foreach ($file in [System.IO.Directory]::EnumerateFiles($dir, "*.*")) {
                if ([System.IO.Path]::GetExtension($file) -match $ExtPattern) {
                    if ($StringsOnly) { $file } else { [System.IO.FileInfo]::new($file) }
                }
            }
        } catch {}
        try {
            foreach ($sub in [System.IO.Directory]::EnumerateDirectories($dir)) {
                $leaf = [System.IO.Path]::GetFileName($sub.TrimEnd('\'))
                if ($script:ExcludeDirNames -contains $leaf) {
                    Write-DetailedLog "Verzeichnis übersprungen (Ausschlussliste): $sub" "DEBUG"
                    continue
                }
                $queue.Enqueue($sub)
            }
        } catch {}
    }
}

function Reserve-UniqueDestination {
    <#
        Reserviert einen freien Ausweichnamen ATOMAR.

        Die vorherige Fassung prüfte erst mit File::Exists und benutzte den
        Namen danach. Zwischen Prüfung und Verwendung liegt ein Zeitfenster,
        in dem ein parallel laufender Durchgang - oder ein Nutzer, der gerade
        speichert - denselben Namen belegen kann; die Datei würde dann beim
        Copy(..., overwrite=$true) überschrieben.

        FileMode::CreateNew schlägt fehl, wenn die Datei bereits existiert.
        Das Anlegen der 0-Byte-Platzhalterdatei ist damit die Reservierung.
        Der spätere Copy überschreibt den Platzhalter; scheitert er, räumt
        der catch-Zweig die Datei wieder weg.

        Rückgabe: reservierter Pfad, oder $null wenn alle Varianten belegt.
    #>
    param(
        [string]$BasePath,
        [string]$Extension,
        [int]$MaxRetries = $MaxNameClashRetries
    )
    $counter = 2
    while ($counter -le $MaxRetries) {
        $candidate = "${BasePath}_${counter}${Extension}"
        try {
            $fs = [System.IO.File]::Open($candidate,
                  [System.IO.FileMode]::CreateNew,
                  [System.IO.FileAccess]::Write,
                  [System.IO.FileShare]::None)
            $fs.Close()
            return $candidate
        } catch {
            $counter++
        }
    }
    return $null
}

function Move-PasswordToFront {
    <#
        Sortiert ein erfolgreiches Passwort an den Anfang der Liste.

        Ohne das probiert jede Datei die Passwoerter in unveraenderter
        Reihenfolge durch. Liegt das richtige an dritter Stelle, kostet das
        pro Datei zwei zusaetzliche COM-Oeffnungsversuche - bei einem
        Bestand mit einheitlichem Passwort summiert sich das erheblich.
    #>
    param([string]$Password)
    if ([string]::IsNullOrEmpty($Password))         { return }
    if (-not $script:Passwords.Contains($Password)) { return }
    $idx = $script:Passwords.IndexOf($Password)
    if ($idx -le 0) { return }
    $script:Passwords.RemoveAt($idx)
    $script:Passwords.Insert(0, $Password)
    Write-DetailedLog "Passwort-Reihenfolge angepasst (Treffer nach vorn sortiert)." "DEBUG"
}

function Get-ConversionErrorCategory {
    <#
        Klassifiziert eine COM-Fehlermeldung. Ersetzt die bisherige
        Zwei-Klassen-Logik, die per '-like "*Passwort*"' auf einen vom
        Job selbst gesetzten Prefix geprueft hat - dieser Prefix ist
        entfallen, klassifiziert wird jetzt am Originaltext.
    #>
    param([string]$Message)
    if     ($Message -like  "*TIMEOUT*")                              { return "Timeout" }
    elseif ($Message -match "Passwort|password|Kennwort")             { return "Passwort" }
    elseif ($Message -match "Datentr.ger.*voll|nicht gen.gend Speicherplatz|not enough space|disk (is )?full") {
                                                                        return "Datentraeger voll" }
    elseif ($Message -match "Arbeitsspeicher|out of memory|insufficient memory|not enough memory") {
                                                                        return "Arbeitsspeicher" }
    elseif ($Message -match "nicht gefunden|not found|be found|cannot find|kann .* nicht finden") { return "Datei nicht gefunden" }
    else                                                              { return "Konvert" }
}

function Update-Progress {
    param([string]$Status, [string]$CurrentFile = "")
    if (-not $script:UseProgress) { return }

    $now = Get-Date
    $isFinal = ($script:TotalFiles -gt 0 -and $script:ProcessedCount -ge $script:TotalFiles)
    if (-not $isFinal) {
        if (($now - $script:LastProgressUpdate).TotalMilliseconds -lt $script:ProgressMinIntervalMs) {
            return
        }
    }
    $script:LastProgressUpdate = $now

    $pct     = 0
    $etaStr  = if ($script:SkipPreScan) { "ETA: --:--:--" } else { "Berechne..." }
    $elapsed = $now - $ScriptStartTime

    if ($script:TotalFiles -gt 0) {
        $pct = [int](($script:ProcessedCount / $script:TotalFiles) * 100)
        if ($script:ProcessedCount -gt 0) {
            $secPerFile = $elapsed.TotalSeconds / $script:ProcessedCount
            $remaining  = ($script:TotalFiles - $script:ProcessedCount) * $secPerFile
            $etaSpan    = [TimeSpan]::FromSeconds([Math]::Max(0, $remaining))
            $etaStr     = "ETA: " + $etaSpan.ToString('hh\:mm\:ss')
        }
    }

    $totalStr = if ($script:SkipPreScan) { "?" } else { "$($script:TotalFiles)" }
    $activity = "Word Deep Unprotect  |  $etaStr  |  $($script:ProcessedCount) / $totalStr"
    Write-Progress -Activity $activity -Status $Status -CurrentOperation $CurrentFile -PercentComplete ([Math]::Min($pct, 100))
}

# ==================================================================
# TEMP-CLEANUP (Reste alter Läufe)
# ==================================================================

function Remove-StaleTempFolders {
    $parent = [Environment]::GetFolderPath('MyDocuments')
    if ([string]::IsNullOrWhiteSpace($parent)) { return }
    if (-not [System.IO.Directory]::Exists($parent)) { return }
    # Get-ChildItem mit -ErrorAction SilentlyContinue ist robuster als
    # [System.IO.Directory]::EnumerateDirectories: letzteres kann durch eine
    # einzelne UnauthorizedAccessException auf einem Geschwister-Ordner
    # (DLP-gesperrt, EFS, OneDrive-Konflikt) die gesamte Iteration abbrechen.
    $candidates = @()
    try {
        $candidates = @(Get-ChildItem -LiteralPath $parent -Directory `
                                      -Filter "2a_entferne_schutz_word_*" `
                                      -Force -ErrorAction SilentlyContinue)
    } catch {}
    foreach ($d in $candidates) {
        try {
            if ($d.FullName -eq $TempPath) { continue }
            if ($d.Name -match '_(\d+)$') {
                $oldPid = [int]$Matches[1]
                $stillRunning = $false
                try {
                    $null = Get-Process -Id $oldPid -ErrorAction Stop
                    $stillRunning = $true
                } catch {}
                if (-not $stillRunning) {
                    try { Remove-Item -LiteralPath $d.FullName -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false } catch {}
                }
            }
        } catch {}
    }
}

function Get-OriginalFromBackupPath {
    <#
        Leitet aus einem Backup-Pfad den Pfad des Originals ab.
          Vertrag.docx.bak         -> Vertrag.docx
          Vertrag.docx.bak_a1b2c3  -> Vertrag.docx
        Der Suffix wird nur EINMAL abgeschnitten (letztes Vorkommen).
    #>
    param([string]$BackupPath)
    if ($BackupPath -match '^(?<orig>.+?)\.bak(?:_[0-9a-fA-F]+)?$') {
        return $Matches['orig']
    }
    return $null
}

function Remove-OrphanedBackups {
    <#
        Raeumt '.bak'/'.bak_xxxxxx'-Dateien auf, die nach einem harten Abbruch
        eines frueheren Laufs liegen geblieben sind.

        Sicherheitsregeln - ein Backup wird NUR geloescht, wenn ALLE zutreffen:
          1. Der abgeleitete Originalname endet auf eine Word-Endung.
             Fremde '.bak'-Dateien (Editor-Backups o. ae.) bleiben unangetastet.
          2. Das Original existiert und ist groesser als 0 Byte.
          3. Das Backup ist aelter als $BackupCleanupMinAgeHours.

        Fehlt das Original oder ist es leer, ist das Backup moeglicherweise die
        EINZIGE intakte Kopie - genau der Fall, fuer den es angelegt wurde.
        Solche Funde werden gemeldet, niemals geloescht.
    #>
    param([string]$RootPath)

    $res = @{
        Deleted    = 0
        Kept       = 0
        Orphans    = 0
        Failed     = 0
        OrphanList = [System.Collections.Generic.List[string]]::new()
    }
    $cutoff = (Get-Date).AddHours(-$BackupCleanupMinAgeHours)

    foreach ($bak in (Get-WordFilesRobust $RootPath '^\.bak(_[0-9a-fA-F]+)?$' -StringsOnly)) {
        try {
            $orig = Get-OriginalFromBackupPath -BackupPath $bak
            if ([string]::IsNullOrWhiteSpace($orig)) { continue }

            # Regel 1: nur unsere eigenen Word-Backups anfassen
            if ($orig -notmatch '\.do[ct][xm]?$') {
                Write-DetailedLog "Fremde .bak-Datei ignoriert: $(Remove-LongPathPrefix $bak)" "DEBUG"
                continue
            }

            $fi = [System.IO.FileInfo]::new($bak)
            if (-not $fi.Exists) { continue }

            # Regel 2: Original muss existieren und Inhalt haben
            $origOk  = $false
            $origLen = -1
            if ([System.IO.File]::Exists($orig)) {
                try {
                    $origLen = ([System.IO.FileInfo]::new($orig)).Length
                    $origOk  = $origLen -gt 0
                } catch { $origOk = $false; $origLen = -1 }
            }
            if (-not $origOk) {
                $res.Orphans++
                $res.OrphanList.Add((Remove-LongPathPrefix $bak))
                Write-Log "Backup ohne intaktes Original - NICHT gelöscht, bitte prüfen: $(Remove-LongPathPrefix $bak)" "WARN"
                continue
            }

            # Regel 2b: Das Original muss mindestens so gross sein wie das Backup.
            # 'Length -gt 0' allein genuegt nicht. [System.IO.File]::Copy
            # trunkiert die Zieldatei beim Start und schreibt fortlaufend - ein
            # harter Abbruch (Stromausfall, Kill, BSOD) hinterlaesst deshalb
            # typischerweise ein TEILWEISE geschriebenes Original mit Groesse
            # > 0, nicht 0 Byte. Genau dieser Zustand bestand Regel 2, und beim
            # naechsten Lauf wurde das Backup geloescht, das die einzige intakte
            # Kopie war (nachgestellt: Original 380 KB, Backup 2,4 MB - 'wird
            # geloescht'). Diese .bak-Dateien stammen laut Funktionskopf
            # ohnehin nur aus abgebrochenen Laeufen; ein faelschlich behaltenes
            # Backup kostet eine Meldung, ein faelschlich geloeschtes die Datei.
            if ($origLen -lt $fi.Length) {
                $res.Orphans++
                $res.OrphanList.Add((Remove-LongPathPrefix $bak))
                Write-Log ("Backup groesser als das Original ({0:N0} statt {1:N0} Bytes) - " +
                           "Original moeglicherweise abgeschnitten. NICHT gelöscht, bitte prüfen: {2}" -f `
                           $fi.Length, $origLen, (Remove-LongPathPrefix $bak)) "WARN"
                continue
            }

            # Regel 3: Mindestalter - an der CreationTime gemessen (wie 2b),
            # NICHT an der LastWriteTime. Das Backup entsteht per
            # [System.IO.File]::Copy, und Copy uebernimmt die LastWriteTime der
            # QUELLE. Auf Archiv-Ablagen ist die Jahre alt; ein Sekunden altes
            # .bak (abgebrochener oder parallel laufender Durchgang) galt damit
            # sofort als verwaist und wurde geloescht (nachgestellt: Quelle mit
            # LastWriteTime 2019 kopiert -> Kopie LastWriteTime 2019,
            # CreationTime jetzt). Die Frist war wirkungslos.
            if ($fi.CreationTime -gt $cutoff) {
                $res.Kept++
                Write-DetailedLog "Backup zu jung, behalten: $(Remove-LongPathPrefix $bak)" "DEBUG"
                continue
            }

            if (Confirm-Write (Remove-LongPathPrefix $bak) 'Verwaistes Backup löschen') {
                try {
                    [System.IO.File]::Delete((Add-LongPathPrefix $bak))
                    $res.Deleted++
                    Write-DetailedLog "Verwaistes Backup entfernt: $(Remove-LongPathPrefix $bak)" "DEBUG"
                } catch {
                    $res.Failed++
                    Write-DetailedLog "Backup nicht löschbar: $(Remove-LongPathPrefix $bak) - $_" "WARN"
                }
            } else {
                $res.Kept++
            }
        } catch {
            $res.Failed++
            Write-DetailedLog "Backup-Prüfung fehlgeschlagen ($bak): $_" "WARN"
        }
    }
    return $res
}

function Clear-WindowsTempWhitelist {
    # Whitelist-Cleanup von %TEMP% am Skriptende: Es werden AUSSCHLIESSLICH
    # Eintraege mit bekannten Praefixen entfernt (Office-/COM-Reste sowie
    # eigene Skript-Artefakte). NIEMALS pauschal leeren - Fremdprozesse
    # legen aktive Daten ohne Lock in %TEMP% ab; blindes Loeschen
    # zerstoert sie.
    $winTemp = $env:TEMP
    if ([string]::IsNullOrWhiteSpace($winTemp)) { $winTemp = $env:TMP }
    if ([string]::IsNullOrWhiteSpace($winTemp) -or -not (Test-Path -LiteralPath $winTemp)) { return }

    $prefixes = @('~$', '~df', 'gen_py', 'vbe', 'excel8.0', '2a_entferne_schutz_word')
    $entries  = @()
    try {
        $entries = @(Get-ChildItem -LiteralPath $winTemp -Force -ErrorAction SilentlyContinue)
    } catch { return }

    foreach ($e in $entries) {
        $low = $e.Name.ToLowerInvariant()
        $hit = $false
        foreach ($p in $prefixes) {
            if ($low.StartsWith($p)) { $hit = $true; break }
        }
        if (-not $hit -and $low.StartsWith('cvr') -and $low.EndsWith('.tmp')) { $hit = $true }
        if (-not $hit) { continue }
        try {
            if ($e.PSIsContainer) {
                Remove-Item -LiteralPath $e.FullName -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false
            } else {
                Remove-Item -LiteralPath $e.FullName -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false
            }
        } catch {}
    }
}

# ==================================================================
# KONVERTIERUNG (alt .doc/.dot -> .docx/.dotx)
# ==================================================================
function Convert-DocToDocx {
    param([string]$SourcePath, [string]$DestPathBase, [string[]]$Passwords)

    $pidFile = Join-Path $TempPath "WordDeepClean_$([Guid]::NewGuid().ToString()).pid"
    $job = Start-Job -ScriptBlock {
        param($src, $destBase, [object[]]$pws, $pidFile,
              $WD_FORMAT_DOCX, $WD_FORMAT_DOCM, $WD_FORMAT_DOTX, $WD_FORMAT_DOTM)

        $existingPids = @(Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })

        $word = New-Object -ComObject Word.Application
        $word.Visible       = $false
        $word.DisplayAlerts = 0
        try { $word.Interactive = $false } catch {}
        try { $word.AutomationSecurity = 3 } catch {}
        try { $word.Options.UpdateLinksAtOpen = $false } catch {}
        try { $word.Options.DoNotPromptForConvert = $true } catch {}

        Start-Sleep -Milliseconds 50
        $myWordPid = $null
        foreach ($p in (Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue)) {
            if ($existingPids -notcontains $p.Id) { $myWordPid = $p.Id; break }
        }
        if ($myWordPid) { try { [System.IO.File]::WriteAllText($pidFile, "$myWordPid") } catch {} }

        $isOleFormat = ($src -match "\.do[ct]$")
        # tryList defensiv aufbauen: Über die Start-Job-Boundary werden Argumente
        # via PSRP serialisiert. Ein [string[]] wird im Job-Runspace zu einem
        # Deserialized.System.String[]-Wrapper, bei dem der @()-Operator NICHT
        # zuverlässig idempotent ist – er kann den ganzen Wrapper als 1-Element-
        # Array umfassen statt ihn zu entrollen. Wenn dann später foreach($pw in
        # $tryList) läuft, ist $pw plötzlich eine ArrayList statt ein String,
        # und der COM-Aufruf $word.Documents.Open(..., $pw, $pw, ...) scheitert
        # mit "Ausnahme beim Festlegen von 'Open': ArrayList kann nicht in
        # Object konvertiert werden". List[string] + expliziter [string]-Cast
        # eliminiert das.
        $tryList = [System.Collections.Generic.List[string]]::new()
        if ($null -ne $pws) {
            foreach ($p in $pws) {
                if ($null -ne $p) { $tryList.Add([string]$p) }
            }
        }
        if ($isOleFormat) {
            $tryList.Insert(0, "")
        } else {
            $tryList.Add("")
        }

        $doc          = $null
        $lastErr      = "Kein Versuch erfolgreich"
        $successfulPw = $null
        $usedRepair   = $false

        # Eskalation analog zum CorruptLoad-Pfad in 2b: erst normal oeffnen,
        # dann mit Word-Reparatur. Word kennt kein xlExtractData-Aequivalent,
        # OpenAndRepair ist die letzte Stufe.
        #   Documents.Open(FileName, ConfirmConversions, ReadOnly,
        #                  AddToRecentFiles, PasswordDocument, PasswordTemplate,
        #                  Revert, WritePasswordDocument, WritePasswordTemplate,
        #                  Format, Encoding, Visible, OpenAndRepair)
        foreach ($repair in @($false, $true)) {
            foreach ($pwRaw in $tryList) {
                # Paranoia-Cast: garantiert, dass das 5./6. Argument von Open
                # ein primitiver String ist (kein Wrapper über die Job-Boundary).
                $pw = [string]$pwRaw
                try {
                    if ($repair) {
                        $doc = $word.Documents.Open(
                            $src, $false, $true, $false, $pw, $pw, $true, "", "",
                            0, "", $false, $true
                        )
                    } else {
                        $doc = $word.Documents.Open(
                            $src, $false, $true, $false, $pw, $pw, $true, "", ""
                        )
                    }
                    $successfulPw = $pw
                    $usedRepair   = $repair
                    break
                } catch {
                    $lastErr = $_.Exception.Message
                    $doc     = $null
                }
            }
            if ($doc) { break }
        }

        if (-not $doc) {
            try { $word.Interactive = $true } catch {}
            # Quit() nur fuer die selbst gestartete Instanz.
            # Word ist ein Einzelinstanz-COM-Server: laeuft bereits eine
            # Word-Sitzung des Anwenders, liefert New-Object -ComObject
            # GENAU DIESE Instanz statt einer neuen. $myWordPid bleibt dann
            # leer - und ein Quit() darauf schliesst die Sitzung des
            # Anwenders samt ungespeicherter Dokumente. In dem Fall wird die
            # Referenz nur freigegeben; eine evtl. doch eigene Instanz raeumt
            # die Zombie-Erkennung des Aufrufers spaeter ab.
            if ($myWordPid) { try { $word.Quit() } catch {} }
            else { try { $word.Visible = $true } catch {} }
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
            # Originale COM-Fehlermeldung weiterreichen (kein "Passwort:"-Prefix
            # mehr): der Categorizer im Aufrufer entscheidet am echten Text.
            return @{ Status = "ERROR"; Msg = $lastErr; WordPid = $myWordPid }
        }

        try {
            $hasMacros  = $true
            $isTemplate = ($src -match "\.dot$")
            try { $hasMacros = [bool]$doc.HasVBProject } catch { $hasMacros = $true }

            if ($isTemplate) {
                $format = if ($hasMacros) { $WD_FORMAT_DOTM } else { $WD_FORMAT_DOTX }
                $ext    = if ($hasMacros) { ".dotm" } else { ".dotx" }
            } else {
                $format = if ($hasMacros) { $WD_FORMAT_DOCM } else { $WD_FORMAT_DOCX }
                $ext    = if ($hasMacros) { ".docm" } else { ".docx" }
            }

            $final = "$destBase$ext"
            $doc.SaveAs2($final, $format, $false, "", $false)
            $doc.Close($false)

            return @{
                Status       = "OK"
                Path         = $final
                WordPid      = $myWordPid
                UsedPassword = $successfulPw
                UsedRepair   = $usedRepair
            }
        } catch {
            return @{ Status = "ERROR"; Msg = $_.Exception.Message; WordPid = $myWordPid }
        } finally {
            try { $word.Interactive = $true } catch {}
            # Quit nur fuer die eigene Instanz (siehe Hinweis oben).
            if ($myWordPid) { try { $word.Quit() } catch {} }
            else { try { $word.Visible = $true } catch {} }
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
        }
    # KEIN fuehrendes Komma vor $Passwords: (,$Passwords) erzeugt ein
    # 1-Element-Array, das das string[] umschliesst. Als Element einer
    # aeusseren Argumentliste wird es NICHT entrollt - der Job-Parameter
    # empfaengt dann ein Array mit genau einem Element (dem inneren Array),
    # die foreach-Schleife laeuft einmal, und [string] auf ein string[] ergibt
    # die leerzeichen-verbundene Darstellung. Bei drei Passwoertern wurde also
    # nur das Phantom-Passwort 'Ablage1 Archiv2 Depot3' probiert (nachgestellt).
    # Bei genau EINEM Passwort faellt das nicht auf - deshalb hat der Fehler
    # jeden Ein-Passwort-Test ueberlebt.
    # Die Sorge, die das Komma motiviert hat, trifft nicht zu: auch bei LEERER
    # Liste verschieben sich die Folgeparameter nicht ($pidFile & Co. kommen
    # korrekt an, ebenfalls nachgestellt).
    } -ArgumentList $SourcePath, $DestPathBase, $Passwords, $pidFile,
                    $WD_FORMAT_DOCX, $WD_FORMAT_DOCM, $WD_FORMAT_DOTX, $WD_FORMAT_DOTM

    if (-not (Wait-Job $job -Timeout $FileOpenTimeoutSeconds)) {
        try { Stop-Job  $job -ErrorAction SilentlyContinue -WhatIf:$false } catch {}
        try { Remove-Job $job -Force -ErrorAction SilentlyContinue -WhatIf:$false } catch {}
        if ([System.IO.File]::Exists($pidFile)) {
            try {
                $savedPid = [int][System.IO.File]::ReadAllText($pidFile).Trim()
                if ($savedPid -gt 0) { Add-TrackedWordPid -ProcessId $savedPid }
            } catch {}
            try { [System.IO.File]::Delete($pidFile) } catch {}
        }
        throw "TIMEOUT: Datei reagiert nicht (evtl. verschlüsselt oder Dialog hängt)."
    }

    $result = Receive-Job $job
    try { Remove-Job $job -Force -ErrorAction SilentlyContinue -WhatIf:$false } catch {}
    if ([System.IO.File]::Exists($pidFile)) { try { [System.IO.File]::Delete($pidFile) } catch {} }

    # Add-TrackedWordPid prueft Name und Startzeit selbst.
    if ($result.WordPid) { Add-TrackedWordPid -ProcessId ([int]$result.WordPid) }

    if ($result.Status -eq "OK") {
        Start-Sleep -Milliseconds $PostComStabilizeMs
        # Hashtable statt nacktem Pfad: der Aufrufer braucht zusaetzlich das
        # erfolgreiche Passwort (Reihenfolge-Lernen) und die Reparatur-Info.
        return @{
            Path         = $result.Path
            UsedPassword = $result.UsedPassword
            UsedRepair   = [bool]$result.UsedRepair
        }
    } else {
        throw $result.Msg
    }
}

# ==================================================================
# ÖFFNEN-PASSWORT PER COM ENTFERNEN
# ==================================================================
function Remove-OpenPassword {
    param([string]$FilePath, [string[]]$Passwords, [string]$OriginalExtension = "")

    $pidFile = Join-Path $TempPath "WordPwRemove_$([Guid]::NewGuid().ToString()).pid"
    $job = Start-Job -ScriptBlock {
        param($fp, [object[]]$pws, $origExt, $pidFile)

        $existingPids = @(Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })

        $word = New-Object -ComObject Word.Application
        $word.Visible       = $false
        $word.DisplayAlerts = 0
        try { $word.Interactive = $false } catch {}
        try { $word.AutomationSecurity = 3 } catch {}
        try { $word.Options.UpdateLinksAtOpen = $false } catch {}
        try { $word.Options.DoNotPromptForConvert = $true } catch {}

        Start-Sleep -Milliseconds 50
        $myWordPid = $null
        foreach ($p in (Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue)) {
            if ($existingPids -notcontains $p.Id) { $myWordPid = $p.Id; break }
        }
        if ($myWordPid) { try { [System.IO.File]::WriteAllText($pidFile, "$myWordPid") } catch {} }

        $isOleFormat = ($origExt -match "^\.do[ct]$")
        # tryList defensiv aufbauen – siehe Begründung in Convert-DocToDocx:
        # Über die Start-Job-Boundary serialisierte [string[]] werden zu
        # Deserialized-Wrappern, die @() nicht zuverlässig entrollt.
        $tryList = [System.Collections.Generic.List[string]]::new()
        if ($null -ne $pws) {
            foreach ($p in $pws) {
                if ($null -ne $p) { $tryList.Add([string]$p) }
            }
        }
        if ($isOleFormat) {
            $tryList.Insert(0, "")
        } else {
            $tryList.Add("")
        }

        $doc          = $null
        $lastErr      = "Kein Passwort hat funktioniert"
        $successfulPw = $null

        foreach ($pwRaw in $tryList) {
            $pw = [string]$pwRaw
            try {
                $doc = $word.Documents.Open(
                    $fp, $false, $false, $false, $pw, $pw, $false, $pw, $pw
                )
                $successfulPw = $pw
                break
            } catch {
                $lastErr = $_.Exception.Message
                $doc     = $null
            }
        }

        if (-not $doc) {
            try { $word.Interactive = $true } catch {}
            # Quit nur fuer die eigene Instanz (siehe Hinweis oben).
            if ($myWordPid) { try { $word.Quit() } catch {} }
            else { try { $word.Visible = $true } catch {} }
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
            return @{ Status = "ERROR"; Msg = $lastErr; WordPid = $myWordPid }
        }

        try {
            $doc.Password      = ""
            $doc.WritePassword = ""
            $doc.Save()
            $doc.Close($false)
            return @{ Status = "OK"; WordPid = $myWordPid; UsedPassword = $successfulPw }
        } catch {
            return @{ Status = "ERROR"; Msg = $_.Exception.Message; WordPid = $myWordPid }
        } finally {
            try { $word.Interactive = $true } catch {}
            # Quit nur fuer die eigene Instanz (siehe Hinweis oben).
            if ($myWordPid) { try { $word.Quit() } catch {} }
            else { try { $word.Visible = $true } catch {} }
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
        }
    # Kein fuehrendes Komma - Begruendung siehe Convert-DocToDocx oben.
    } -ArgumentList $FilePath, $Passwords, $OriginalExtension, $pidFile

    if (-not (Wait-Job $job -Timeout $FileOpenTimeoutSeconds)) {
        try { Stop-Job  $job -ErrorAction SilentlyContinue -WhatIf:$false } catch {}
        try { Remove-Job $job -Force -ErrorAction SilentlyContinue -WhatIf:$false } catch {}
        if ([System.IO.File]::Exists($pidFile)) {
            try {
                $savedPid = [int][System.IO.File]::ReadAllText($pidFile).Trim()
                if ($savedPid -gt 0) { Add-TrackedWordPid -ProcessId $savedPid }
            } catch {}
            try { [System.IO.File]::Delete($pidFile) } catch {}
        }
        # Haengende Instanz sofort abraeumen. Stop-Job beendet nur den
        # Job-Prozess, nicht das per DCOM gestartete WINWORD. Der Aufrufer
        # bucht den Timeout als 'SKIP (Passwort)' und macht weiter - anders als
        # der Konvertierungspfad ruft er Cleanup-AllWord NICHT auf. Das Word
        # (ggf. mit sichtbarem Kennwortdialog) lief so bis zum Laufende weiter,
        # hielt die Temp-Kopie offen, und jeder weitere Timeout kam eine
        # Instanz dazu. Es laeuft immer nur ein Job zugleich; die Liste
        # enthaelt hier also keine noch gebrauchte Instanz.
        Cleanup-AllWord
        # Einheitliches Rueckgabeformat auch im Timeout-Fall: der Aufrufer
        # greift auf .Success zu, ein nacktes $false haette dort still $null
        # ergeben.
        return @{ Success = $false; UsedPassword = $null }
    }

    $result = Receive-Job $job
    try { Remove-Job $job -Force -ErrorAction SilentlyContinue -WhatIf:$false } catch {}
    if ([System.IO.File]::Exists($pidFile)) { try { [System.IO.File]::Delete($pidFile) } catch {} }

    # Add-TrackedWordPid prueft Name und Startzeit selbst.
    if ($result.WordPid) { Add-TrackedWordPid -ProcessId ([int]$result.WordPid) }

    if ($result.Status -eq "OK") {
        Start-Sleep -Milliseconds $PostComStabilizeMs
        return @{ Success = $true; UsedPassword = $result.UsedPassword }
    }
    return @{ Success = $false; UsedPassword = $null }
}

# ==================================================================
# SCHUTZ ENTFERNEN (XML-Ebene)
# ==================================================================
function Remove-WordProtection {
    param([string]$FilePath)

    $actionsTaken = [System.Collections.ArrayList]::new()

    $zip = $null
    try {
        $zip = Invoke-WithRetry -Context "ZipFile.Open Update" -Action {
            [System.IO.Compression.ZipFile]::Open($FilePath, "Update")
        }

        $targetNames = @("word/settings.xml", "word/document.xml")
        foreach ($e in $zip.Entries) {
            if ($e.FullName -match '^word/(header|footer|footnotes|endnotes)\d*\.xml$') {
                $targetNames += $e.FullName
            }
        }
        $targetNames = @($targetNames | Select-Object -Unique)

        foreach ($targetName in $targetNames) {
            $entry = $zip.GetEntry($targetName)
            if (-not $entry) {
                Write-DetailedLog "Eintrag nicht gefunden: $targetName" "DEBUG"
                continue
            }

            $stream  = $entry.Open()
            $reader  = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::UTF8)
            $content = $reader.ReadToEnd()
            $reader.Close()
            $stream.Close()

            $xmlDoc = New-Object System.Xml.XmlDocument
            try {
                $xmlDoc.LoadXml($content)
            } catch {
                Write-DetailedLog "XML-Parsefehler in ${targetName}: $_" "WARN"
                continue
            }

            $modified = $false

            if ($targetName -eq "word/settings.xml") {
                $xpaths = @(
                    "//*[local-name()='documentProtection']",
                    "//*[local-name()='writeProtection']"
                )
            } else {
                $xpaths = @(
                    "//*[local-name()='formProt']",
                    "//*[local-name()='permStart']",
                    "//*[local-name()='permEnd']",
                    "//*[local-name()='fieldLock']",
                    "//*[local-name()='lock' and parent::*[local-name()='sdtPr']]"
                )
            }

            foreach ($xpath in $xpaths) {
                $nodes = @($xmlDoc.SelectNodes($xpath))
                foreach ($node in $nodes) {
                    $nodeName = $node.LocalName
                    Write-DetailedLog "Entferne <$nodeName> aus $targetName" "DEBUG"
                    $node.ParentNode.RemoveChild($node) | Out-Null
                    $modified = $true

                    switch ($nodeName) {
                        "documentProtection" { $actionsTaken.Add("Dokumentschutz")   | Out-Null }
                        "writeProtection"    { $actionsTaken.Add("Schreibschutz")    | Out-Null }
                        "formProt"           { $actionsTaken.Add("Abschnittsschutz") | Out-Null }
                        "permStart"          { $actionsTaken.Add("Inhaltsbereich")   | Out-Null }
                        "permEnd"            { }
                        "fieldLock"          { $actionsTaken.Add("Feldsperre")       | Out-Null }
                        "lock"               { $actionsTaken.Add("ContentControl-Sperre") | Out-Null }
                        default              { $actionsTaken.Add($nodeName)          | Out-Null }
                    }
                }
            }

            if ($modified) {
                $entry.Delete()
                $newEntry    = $zip.CreateEntry($targetName)
                $writeStream = $newEntry.Open()

                $utf8NoBom         = [System.Text.UTF8Encoding]::new($false)
                $writer            = [System.Xml.XmlTextWriter]::new($writeStream, $utf8NoBom)
                $writer.Formatting = [System.Xml.Formatting]::None
                $xmlDoc.Save($writer)
                $writer.Close()
                $writeStream.Close()

                Write-DetailedLog "Eintrag $targetName neu geschrieben" "DEBUG"
            }
        }

        $zip.Dispose()
        $zip = $null

    } catch {
        if ($zip) { try { $zip.Dispose() } catch {} }
        throw "XML-Fehler: $_"
    }

    # Fuehrendes Komma ist ZWINGEND: PowerShell entrollt Sammlungen beim
    # Schreiben in die Ausgabe. 'return @()' liefert dadurch GAR NICHTS, die
    # aufrufende Variable wird $null - und '@($null).Count' ist 1, nicht 0.
    # Ohne das Komma galt eine Datei ohne jeden Schutz als "eine Aktion":
    # sie wurde gesichert, zurueckgeschrieben und als entsperrt gezaehlt,
    # der Zweig "Kein Schutz" war toter Code.
    return ,@($actionsTaken | Select-Object -Unique)
}

# ==================================================================
# TRUST-CENTER SMOKE-TEST
# ==================================================================

function New-MinimalDocx {
    param([string]$Path)

    if ([System.IO.File]::Exists($Path)) {
        [System.IO.File]::Delete($Path)
    }

    $contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    $rels         = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    $document     = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>'

    $entries = @(
        @{ Name = "[Content_Types].xml"; Content = $contentTypes }
        @{ Name = "_rels/.rels";         Content = $rels         }
        @{ Name = "word/document.xml";   Content = $document     }
    )

    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    $zip       = [System.IO.Compression.ZipFile]::Open($Path, 'Create')
    try {
        foreach ($e in $entries) {
            $entry = $zip.CreateEntry($e.Name)
            $s     = $entry.Open()
            $bytes = $utf8NoBom.GetBytes($e.Content)
            $s.Write($bytes, 0, $bytes.Length)
            $s.Close()
        }
    } finally {
        $zip.Dispose()
    }
}

function Test-WordTrustCenter {
    param(
        [string]$TempDir,
        [int]$TimeoutSec = 25
    )

    $testDocx = Join-Path $TempDir ("trustcheck_{0}.docx" -f ([Guid]::NewGuid().ToString("N")))
    $pidFile  = Join-Path $TempDir ("trustcheck_{0}.pid"  -f ([Guid]::NewGuid().ToString("N")))

    try {
        New-MinimalDocx -Path $testDocx
    } catch {
        return @{ Ok = $false; Msg = "Test-DOCX konnte nicht erstellt werden: $_" }
    }

    $job = Start-Job -ScriptBlock {
        param($fp, $pidFile)

        $existingPids = @(Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })

        $word = New-Object -ComObject Word.Application
        $word.Visible       = $false
        $word.DisplayAlerts = 0
        try { $word.Interactive = $false } catch {}
        try { $word.AutomationSecurity = 3 } catch {}
        try { $word.Options.UpdateLinksAtOpen = $false } catch {}
        try { $word.Options.DoNotPromptForConvert = $true } catch {}

        Start-Sleep -Milliseconds 50
        $myWordPid = $null
        foreach ($p in (Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue)) {
            if ($existingPids -notcontains $p.Id) { $myWordPid = $p.Id; break }
        }
        if ($myWordPid) { try { [System.IO.File]::WriteAllText($pidFile, "$myWordPid") } catch {} }

        try {
            # Documents.Open: FileName, ConfirmConversions=$false, ReadOnly=$true, AddToRecentFiles=$false
            $doc = $word.Documents.Open($fp, $false, $true, $false)
            if (-not $doc) {
                return @{ Ok = $false; Msg = "Open lieferte kein Document"; WordPid = $myWordPid }
            }
            try { $doc.Close($false) } catch {}
            return @{ Ok = $true; WordPid = $myWordPid }
        } catch {
            return @{ Ok = $false; Msg = $_.Exception.Message; WordPid = $myWordPid }
        } finally {
            try { $word.Interactive = $true } catch {}
            # Quit nur fuer die eigene Instanz (siehe Hinweis oben).
            if ($myWordPid) { try { $word.Quit() } catch {} }
            else { try { $word.Visible = $true } catch {} }
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    } -ArgumentList $testDocx, $pidFile

    $completed = Wait-Job $job -Timeout $TimeoutSec

    if (-not $completed) {
        try { Stop-Job   $job -ErrorAction SilentlyContinue -WhatIf:$false } catch {}
        try { Remove-Job $job -Force -ErrorAction SilentlyContinue -WhatIf:$false } catch {}
        if ([System.IO.File]::Exists($pidFile)) {
            try {
                $savedPid = [int][System.IO.File]::ReadAllText($pidFile).Trim()
                if ($savedPid -gt 0) {
                    Add-TrackedWordPid -ProcessId $savedPid
                    # Nur beenden, wenn es noch derselbe Prozess ist (Name und
                    # gemerkte Startzeit) - hier stand ein ungeprueftes
                    # Stop-Process auf die PID aus der Datei.
                    if (Test-IsOwnWordProcess -ProcessId $savedPid) {
                        try { Stop-Process -Id $savedPid -Force -ErrorAction SilentlyContinue -WhatIf:$false } catch {}
                    }
                }
            } catch {}
            try { [System.IO.File]::Delete($pidFile) } catch {}
        }
        try { [System.IO.File]::Delete($testDocx) } catch {}
        return @{ Ok = $false; Msg = "TIMEOUT nach ${TimeoutSec}s - Trust Center vermutlich nicht konfiguriert oder Office-Profil beschaedigt" }
    }

    $result = Receive-Job $job
    try { Remove-Job $job -Force -ErrorAction SilentlyContinue -WhatIf:$false } catch {}
    if ([System.IO.File]::Exists($pidFile)) { try { [System.IO.File]::Delete($pidFile) } catch {} }
    try { [System.IO.File]::Delete($testDocx) } catch {}
    if ($result -and $result.WordPid) { Add-TrackedWordPid -ProcessId ([int]$result.WordPid) }

    if ($result -and $result.Ok) {
        return @{ Ok = $true }
    } else {
        $m = if ($result) { $result.Msg } else { "Kein Ergebnis vom COM-Job" }
        return @{ Ok = $false; Msg = $m }
    }
}

# ==================================================================
# INTERAKTIVER START
# ==================================================================
if (-not $NoInteractive) {

    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║                WORD DEEP UNPROTECT                 ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════╝" -ForegroundColor Cyan

    # Vor dem ersten COM-Zugriff auf laufende Office-Sitzungen hinweisen.
    if (-not (Show-OfficeRunningWarning -Silent:$NoInteractive)) { exit 0 }
    Write-Host ""
    Write-Host "HINWEIS: Das Skript verwendet den Ordner 'Dokumente' als temporären" -ForegroundColor DarkYellow
    Write-Host "Arbeitsordner für die COM-Verarbeitung (Word öffnet Dateien daraus)." -ForegroundColor DarkYellow
    Write-Host "" -ForegroundColor DarkYellow
    Write-Host "Bitte in Word sicherstellen:" -ForegroundColor DarkYellow
    Write-Host "  Datei > Optionen > Trust Center > Einstellungen für das Trust Center" -ForegroundColor Gray
    Write-Host "  > Vertrauenswürdige Speicherorte:" -ForegroundColor Gray
    Write-Host "    1. Den Dokumente-Ordner hinzufügen (falls nicht vorhanden):" -ForegroundColor Gray
    Write-Host "       $([Environment]::GetFolderPath('MyDocuments'))" -ForegroundColor White
    Write-Host "       [x] Unterordner dieser Speicherorte sind ebenfalls vertrauenswürdig" -ForegroundColor Gray
    Write-Host "    2. [x] Vertrauenswürdige Speicherorte im Netzwerk zulassen" -ForegroundColor Gray
    Write-Host "       (erforderlich, wenn Dokumente-Ordner auf Netzlaufwerk umgeleitet)" -ForegroundColor Gray
    Write-Host ""

    if ([string]::IsNullOrWhiteSpace($TargetPath)) {
        $desktopPath   = Get-UserShellFolder -Name 'Desktop' `
                                             -Fallback (Join-Path $env:USERPROFILE 'Desktop')
        $downloadsPath = Get-UserShellFolder -Name '{374DE290-123F-4565-9164-39C4925E467B}' `
                                             -Fallback (Join-Path $env:USERPROFILE 'Downloads')

        $presetCount = $DirectoryPresets.Count
        Write-Host "Zielpfad auswählen:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $presetCount; $i++) {
            Write-Host ("  [{0}] {1}" -f ($i + 1), $DirectoryPresets[$i]) -ForegroundColor White
        }
        # Bewusst Verkettung statt -f: Pfade koennen '{' oder '}' enthalten,
        # was den Formatoperator zur Laufzeit sprengen wuerde.
        Write-Host ("  [" + ($presetCount + 1) + "] Desktop      ($desktopPath)")   -ForegroundColor White
        Write-Host ("  [" + ($presetCount + 2) + "] Downloads    ($downloadsPath)") -ForegroundColor White
        Write-Host ("  [" + ($presetCount + 3) + "] Eigenen Pfad eingeben")         -ForegroundColor White
        Write-Host ""
        $pathChoice = (Read-Host ("Auswahl [1-{0}]" -f ($presetCount + 3))).Trim()

        $choiceNum = 0
        [void][int]::TryParse($pathChoice, [ref]$choiceNum)

        switch ($true) {
            ($choiceNum -ge 1 -and $choiceNum -le $presetCount) {
                $TargetPath = $DirectoryPresets[$choiceNum - 1]
                if (-not (Test-PathLong $TargetPath)) {
                    Write-Error "Nicht erreichbar: $TargetPath  (Laufwerk verbunden?)"
                    exit 1
                }
                break
            }
            ($choiceNum -eq $presetCount + 1) { $TargetPath = $desktopPath;   break }
            ($choiceNum -eq $presetCount + 2) { $TargetPath = $downloadsPath; break }
            default {
                $TargetPath = (Read-Host "Pfad eingeben").Trim().Trim('"').Trim("'").Trim()
                if ([string]::IsNullOrWhiteSpace($TargetPath)) {
                    Write-Error "Kein Pfad angegeben. Abbruch."
                    exit 1
                }
                if (-not (Test-PathLong $TargetPath)) {
                    Write-Error "Pfad existiert nicht: $TargetPath"
                    exit 1
                }
            }
        }
    }

    Write-Host ""
    Write-Host "Passworte zum Öffnen geschützter Word-Dateien (bis zu 3, leer = fertig):" -ForegroundColor Yellow
    Write-Host "HINWEIS: Eingabe wird NICHT getrimmt – Leerzeichen am Anfang/Ende bleiben erhalten." -ForegroundColor DarkGray
    for ($pwIdx = 1; $pwIdx -le 3; $pwIdx++) {
        $pwInput = Read-Host "  Passwort $pwIdx (leer = fertig)"
        if ([string]::IsNullOrEmpty($pwInput)) { break }
        $script:Passwords.Add($pwInput)
    }

    Write-Host ""
    Write-Host "Fortschrittsanzeige:" -ForegroundColor Yellow
    Write-Host "  [1] Keine Fortschrittsleiste (laufende Zählung inline) [Standard]" -ForegroundColor White
    Write-Host "  [2] Fortschrittsleiste mit ETA (Vorab-Scan nötig, Start verzögert!)" -ForegroundColor White
    Write-Host "  [3] Fortschrittsleiste ohne ETA (sofortiger Start, kein Total)" -ForegroundColor White
    Write-Host ""
    Write-Host "  Hinweis zu [2]: Bei UNC-Pfaden oder sehr großen Freigaben kann" -ForegroundColor DarkGray
    Write-Host "  die Indizierung 10–30 Minuten dauern, bevor die erste Datei" -ForegroundColor DarkGray
    Write-Host "  verarbeitet wird. Bei > 100.000 Dateien besser [1] oder [3]" -ForegroundColor DarkGray
    Write-Host "  wählen (Generator-Variante ohne Vorab-Zählung)." -ForegroundColor DarkGray
    Write-Host ""
    $progressChoice = (Read-Host "Auswahl [1-3]").Trim()
    switch ($progressChoice) {
        "2" { $script:UseProgress = $true;  $script:SkipPreScan = $false }
        "3" { $script:UseProgress = $true;  $script:SkipPreScan = $true  }
        default { $script:UseProgress = $false; $script:SkipPreScan = $true }
    }

    Write-Host ""
} else {
    if ([string]::IsNullOrWhiteSpace($TargetPath)) {
        Write-Error "Im Modus -NoInteractive ist -TargetPath erforderlich."
        exit 1
    }
    $script:UseProgress = $ShowProgress.IsPresent
    $script:SkipPreScan = $true

    # Auch ohne Dialog protokollieren, dass fremde Office-Sitzungen
    # laufen - im Fehlerfall ist das die haeufigste Ursache.
    [void](Show-OfficeRunningWarning -Silent)
}

# ==================================================================
# VORBEREITUNG
# ==================================================================
if (-not (Test-PathLong $TargetPath)) {
    Write-Error "Pfad nicht gefunden: $TargetPath"
    exit 1
}

Remove-StaleTempFolders

if (Test-Path -LiteralPath $TempPath) {
    Remove-Item -LiteralPath $TempPath -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false
}
# Provider-frei anlegen statt per New-Item: Das Skript deklariert
# SupportsShouldProcess, und unter -WhatIf setzt PowerShell
# $WhatIfPreference fuer den gesamten Skript-Scope. JEDES nachgelagerte
# ShouldProcess-faehige Cmdlet gehorcht dem - auch New-Item und die
# Set-Content/Add-Content der Protokolle. Der Arbeitsordner entstand dann
# nicht, der Smoke-Test scheiterte, und mit -NoInteractive brach der Lauf
# ab; interaktiv scheiterte danach jede einzelne Datei an der Arbeitskopie.
# Genau der im Kopf empfohlene erste Probelauf lieferte damit 'Wuerde
# aendern: 0' und 'Fehler: <alle Dateien>' - ohne Log zum Nachlesen.
# [System.IO.Directory]::CreateDirectory kennt kein ShouldProcess.
[System.IO.Directory]::CreateDirectory($TempPath) | Out-Null

# Restore-Privilegien (Admin-Kontext) fuer ACL/Owner-Erhalt aktivieren.
Enable-RestorePrivileges

# Protokollkoepfe ebenfalls provider-frei schreiben (siehe oben): unter
# -WhatIf haette Set-Content sie unterdrueckt, und der Probelauf haette
# weder Log noch CSV zum Nachlesen hinterlassen. UTF-8 MIT BOM, damit
# Excel die CSV im deutschen Gebietsschema per Doppelklick korrekt oeffnet
# - das entspricht dem bisherigen '-Encoding utf8' der Windows-PowerShell.
$script:LogEncoding = [System.Text.UTF8Encoding]::new($true)
[System.IO.File]::WriteAllText(
    $LogFilePath,
    "Log Start: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  |  Ziel: $TargetPath`r`n",
    $script:LogEncoding)
[System.IO.File]::WriteAllText(
    $DetailedLogPath,
    "Debug-Log Start: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`r`n",
    $script:LogEncoding)
[System.IO.File]::WriteAllText(
    $CsvLogPath,
    "Zeitstempel;Status;Aktionen;Pfad;Details`r`n",
    $script:LogEncoding)

Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host " WORD DEEP UNPROTECT"
Write-Host " Ziel: $TargetPath"
Write-Host "=====================================================" -ForegroundColor Cyan

# ------------------------------------------------------------------
# Trust-Center Smoke-Test
# ------------------------------------------------------------------
Write-Host "Prüfe COM-Subsystem und Trust-Center..." -ForegroundColor DarkGray
$smokeTest = Test-WordTrustCenter -TempDir $TempPath -TimeoutSec $TrustCenterTimeoutSec

if (-not $smokeTest.Ok) {
    Write-Host ""
    Write-Host "WORD-SMOKE-TEST FEHLGESCHLAGEN" -ForegroundColor Red
    Write-Host "   Grund: $($smokeTest.Msg)"      -ForegroundColor Yellow
    Write-Host ""
    Write-Host "   Mögliche Ursachen:" -ForegroundColor Gray
    Write-Host "     - Dokumentenordner nicht als vertrauenswürdiger Speicherort eingetragen" -ForegroundColor Gray
    Write-Host "     - Word nicht installiert oder Office-Profil beschädigt" -ForegroundColor Gray
    Write-Host "     - Fehlender Desktop-Ordner für Dienstkonto (siehe Skript-Header)" -ForegroundColor Gray
    Write-Host ""
    Write-Log "Trust-Center-Test fehlgeschlagen: $($smokeTest.Msg)" "ERROR"

    if ($NoInteractive) {
        Write-Log "Abbruch (NoInteractive) nach fehlgeschlagenem Trust-Center-Test." "ERROR"
        Cleanup-AllWord
        if (Test-Path -LiteralPath $TempPath) { Remove-Item -LiteralPath $TempPath -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false }
        exit 2
    }

    $choice = (Read-Host "Trotzdem fortfahren? (Timeout-Risiko pro Datei!) [j/N]").Trim().ToUpperInvariant()
    if ($choice -ne "J" -and $choice -ne "Y") {
        Write-Host "Abbruch durch Benutzer." -ForegroundColor Yellow
        Write-Log "Abbruch durch Benutzer nach Trust-Center-Warnung." "INFO"
        Cleanup-AllWord
        if (Test-Path -LiteralPath $TempPath) { Remove-Item -LiteralPath $TempPath -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false }
        exit 0
    }
    Write-Log "Trust-Center-Warnung vom Benutzer ignoriert - Fortsetzung." "WARN"
} else {
    Write-Host "Smoke-Test OK." -ForegroundColor DarkGray
    Write-DetailedLog "Trust-Center-Smoke-Test erfolgreich." "DEBUG"
}

# ------------------------------------------------------------------
# Verwaiste Backups früherer Läufe aufräumen
# ------------------------------------------------------------------
$backupCleanup = $null
if (-not $SkipBackupCleanup) {
    Write-Host "Suche zurückgebliebene Backups früherer Läufe..." -ForegroundColor DarkGray
    $backupCleanup = Remove-OrphanedBackups -RootPath (Add-LongPathPrefix $TargetPath)

    if ($backupCleanup.Deleted -gt 0 -or $backupCleanup.Orphans -gt 0 -or $backupCleanup.Kept -gt 0) {
        Write-Host ("  Backups: {0} entfernt, {1} behalten, {2} ohne Original" -f `
                    $backupCleanup.Deleted, $backupCleanup.Kept, $backupCleanup.Orphans) -ForegroundColor DarkGray
    } else {
        Write-Host "  Keine Backup-Reste gefunden." -ForegroundColor DarkGray
    }

    if ($backupCleanup.Orphans -gt 0) {
        Write-Host ""
        Write-Host "ACHTUNG: $($backupCleanup.Orphans) Backup(s) ohne intaktes Original gefunden." -ForegroundColor Yellow
        Write-Host "Diese wurden NICHT gelöscht - sie könnten die einzige Kopie sein:" -ForegroundColor Yellow
        $shown = 0
        foreach ($o in $backupCleanup.OrphanList) {
            Write-Host "  $o" -ForegroundColor Gray
            $shown++
            if ($shown -ge 10) {
                Write-Host "  ... und $($backupCleanup.OrphanList.Count - 10) weitere (siehe Log)" -ForegroundColor Gray
                break
            }
        }
        Write-Host "Zum Wiederherstellen die Endung '.bak' bzw. '.bak_xxxxxx' entfernen." -ForegroundColor Gray
        Write-Host ""
    }
    Write-Log ("Backup-Aufräumen: {0} entfernt, {1} behalten, {2} ohne Original, {3} Fehler" -f `
               $backupCleanup.Deleted, $backupCleanup.Kept, $backupCleanup.Orphans, $backupCleanup.Failed) "INFO"
} else {
    Write-DetailedLog "Backup-Aufräumen übersprungen (-SkipBackupCleanup)." "DEBUG"
}

if ($script:UseProgress -and -not $script:SkipPreScan) {
    Write-Host "Zähle Dateien für ETA (bitte warten)..." -ForegroundColor DarkGray
    $script:TotalFiles = (Get-WordFilesRobust (Add-LongPathPrefix $TargetPath) "^\.do[ct][xm]?$" -StringsOnly |
                          Measure-Object).Count
    Write-Host "Gefunden: $($script:TotalFiles) Word-Dateien`n" -ForegroundColor DarkGray
}

$stats = @{
    Processed   = 0
    Unlocked    = 0
    Skipped     = 0
    Errors      = 0
    Junk        = 0
    JunkKept    = 0
    Converted   = 0
    Locked      = 0
    Encrypted   = 0
    # COM-Timeouts getrennt fuehren: sie wurden bisher auf
    # 'Encrypted' gebucht und erschienen in der Endstatistik als
    # verschluesselte Dateien - eine Fehldiagnose, die den
    # Betreiber nach Passwoertern suchen laesst, obwohl Word
    # schlicht nicht geantwortet hat.
    Timeouts    = 0
    WouldChange = 0
    Repaired    = 0
}

# ==================================================================
# HAUPTSCHLEIFE
# ==================================================================
Write-Host "Starte Verarbeitung..." -ForegroundColor Cyan

# Strg+C als Eingabe behandeln (Begruendung beim Trap oben); nur interaktiv
# an der Konsole - im Modus -NoInteractive fragt niemand die Tasten ab.
try {
    if (-not $NoInteractive -and $Host.Name -eq 'ConsoleHost') {
        [Console]::TreatControlCAsInput = $true
        $script:CtrlCAsInput = $true
        Write-Host "Abbruch mit ESC oder Strg+C (nach der laufenden Datei)." -ForegroundColor DarkGray
    }
} catch {}

try {
Get-WordFilesRobust (Add-LongPathPrefix $TargetPath) "^\.do[ct][xm]?$" |
    ForEach-Object {

    # ESC / Strg+C abfragen (wie 7): wirkt vor der naechsten Datei, die
    # laufende wird nie mittendrin verlassen.
    if ($script:CtrlCAsInput) {
        try {
            while ([Console]::KeyAvailable) {
                $key = [Console]::ReadKey($true)
                if ($key.Key -eq 'Escape' -or
                    ($key.Key -eq [ConsoleKey]::C -and
                     (($key.Modifiers -band [ConsoleModifiers]::Control) -ne 0))) {
                    $script:ShouldStop = $true
                    break
                }
            }
        } catch {}
    }

    if ($script:ShouldStop) { throw [System.OperationCanceledException]::new() }

    $file = $_

    $workFile = $null

    $stats.Processed++
    $script:ProcessedCount = $stats.Processed
    Update-Progress -Status "Verarbeite..." -CurrentFile $file.FullName

    if (($stats.Processed % $PidCleanupInterval) -eq 0) { Sweep-DeadPids }

    if (-not $script:UseProgress) {
        Write-Host -NoNewline "[#$($stats.Processed)] $($file.Name) "
    }

    # --- Junk-Dateien löschen ---
    # '~$'-Dateien sind Word-SPERRDATEIEN: sie existieren, solange ein Dokument
    # geoeffnet ist. Sofortiges Loeschen bricht die Sperre und andere Nutzer
    # bekommen keinen Hinweis mehr, dass die Datei in Bearbeitung ist.
    # Deshalb erst ab $JunkMinAgeHours entfernen.
    if ($file.Name.StartsWith("~`$") -or $file.Name.StartsWith("._")) {
        $ageHours = ((Get-Date) - $file.LastWriteTime).TotalHours
        if ($ageHours -lt $JunkMinAgeHours) {
            $stats.JunkKept++
            $stats.Skipped++
            if (-not $script:UseProgress) {
                Write-Host ("-> SKIP (Sperrdatei, erst {0:N1} h alt)" -f $ageHours) -ForegroundColor DarkGray
            }
            Write-DetailedLog ("Sperrdatei geschont ({0:N1}h < {1}h): {2}" -f $ageHours, $JunkMinAgeHours, $file.FullName) "DEBUG"
            Write-CsvLog -Status "SKIP" -Actions "Sperrdatei geschont" -Path $file.FullName -Details ("{0:N1} h alt" -f $ageHours)
            return
        }
        if (Confirm-Write (Remove-LongPathPrefix $file.FullName) 'Verwaiste Sperrdatei löschen') {
            try { [System.IO.File]::Delete((Add-LongPathPrefix $file.FullName)) } catch {}
            $stats.Junk++
            if (-not $script:UseProgress) { Write-Host "-> JUNK" -ForegroundColor DarkGray }
            Write-DetailedLog "Junk entfernt: $($file.FullName)" "DEBUG"
            Write-CsvLog -Status "JUNK" -Actions "Sperrdatei geloescht" -Path $file.FullName
        } else {
            $stats.WouldChange++
            if (-not $script:UseProgress) { Write-Host "-> WHATIF (Junk)" -ForegroundColor DarkCyan }
            Write-CsvLog -Status "WHATIF" -Actions "Wuerde Sperrdatei loeschen" -Path $file.FullName
        }
        return
    }

    if ($file.Length -eq 0) {
        $stats.Skipped++
        if (-not $script:UseProgress) { Write-Host "-> SKIP (0 Byte / korrupt)" -ForegroundColor DarkGray }
        Write-Log         "0-Byte-Datei übersprungen (korrupt): $($file.FullName)" "WARN"
        Write-DetailedLog "0-Byte-Datei übersprungen: $($file.FullName)" "WARN"
        Write-CsvLog -Status "SKIP" -Actions "0 Byte / korrupt" -Path $file.FullName
        return
    }

    $srcLong = Add-LongPathPrefix $file.FullName

    # --- Schreibschutz auf Datei-Ebene ---
    # Nur das ReadOnly-Bit loeschen. Vorher wurde pauschal auf 'Normal' gesetzt
    # und damit auch Hidden, System, Archive und NotContentIndexed entfernt -
    # ein stiller Nebeneffekt, der z. B. Backup-Werkzeuge und Suchindizes
    # durcheinanderbringt.
    # $roEntfernt/$zurueckgeschrieben: Das Attribut wird hier entfernt, bevor
    # feststeht, ob die Datei ueberhaupt geaendert wird. Endete die Datei
    # danach als 'Kein Schutz', SKIP oder Fehler, blieb sie ohne
    # Schreibschutz zurueck, obwohl das Protokoll 'keine Aenderung' meldete.
    # Deshalb wird das Attribut wiederhergestellt, wenn nicht erfolgreich
    # zurueckgeschrieben wurde (Sperr-Check unten und finally).
    $roEntfernt         = $false
    $zurueckgeschrieben = $false
    if ($file.IsReadOnly) {
        if (Confirm-Write (Remove-LongPathPrefix $srcLong) 'Schreibschutz-Attribut entfernen') {
            try {
                $srcAttrs = [System.IO.File]::GetAttributes($srcLong)
                [System.IO.File]::SetAttributes(
                    $srcLong, $srcAttrs -band (-bnot [System.IO.FileAttributes]::ReadOnly))
                $roEntfernt = $true
                Write-DetailedLog "Schreibschutz (Attribut) entfernt: $($file.FullName)" "DEBUG"
            } catch {
                Write-DetailedLog "Attribut-Reset fehlgeschlagen: $($file.FullName) - $_" "WARN"
            }
        }
    }

    # --- Sperr-Check ---
    $lockResult = Test-FileIsLocked -FilePath $file.FullName
    if ($lockResult) {
        if ($roEntfernt) {
            try { [System.IO.File]::SetAttributes($srcLong, [System.IO.File]::GetAttributes($srcLong) -bor [System.IO.FileAttributes]::ReadOnly) } catch {}
        }
        $stats.Locked++
        $stats.Skipped++
        if ($lockResult -eq "AccessDenied") {
            if (-not $script:UseProgress) { Write-Host "-> SKIP (Zugriff verweigert)" -ForegroundColor Red }
            Write-Log "Zugriff verweigert: $($file.FullName)" "WARN"
            Write-DetailedLog "Keine NTFS-Berechtigung oder Datei gesperrt: $($file.FullName)" "WARN"
            Write-CsvLog -Status "SKIP" -Actions "Zugriff verweigert" -Path $file.FullName
        } else {
            if (-not $script:UseProgress) { Write-Host "-> SKIP (gesperrt)" -ForegroundColor Yellow }
            Write-Log "Gesperrt/geöffnet: $($file.FullName)" "WARN"
            Write-DetailedLog "Datei gesperrt: $($file.FullName)" "WARN"
            Write-CsvLog -Status "SKIP" -Actions "Gesperrt/geoeffnet" -Path $file.FullName
        }
        return
    }

    # --- Zeitstempel sichern ---
    $origCreationTime   = $file.CreationTime
    $origLastWriteTime  = $file.LastWriteTime
    $origLastAccessTime = $file.LastAccessTime

    # --- NTFS-Sicherheitsinfo (Owner/Group/DACL) sichern ---
    $origSecurity = Get-FileSecuritySnapshot -Path $srcLong

    $guid         = [Guid]::NewGuid().ToString()
    $tempFile     = Join-Path $TempPath "$guid$($file.Extension)"
    $tempBase     = Join-Path $TempPath $guid
    $workFile     = $tempFile
    $backupPath   = $null
    $wasConverted = $false
    $convRepaired = $false
    # Von Reserve-UniqueDestination angelegte 0-Byte-Platzhalterdatei.
    # Wird im finally entfernt, falls das Zurueckschreiben nicht bis zum
    # Ende kam - sonst bliebe sie z. B. nach einem Fehler beim Backup liegen.
    $reservedPlaceholder = $null

    try {
        # --- In Temp kopieren (Retry + Verfügbarkeits-Wait gegen AV-Sperren) ---
        Invoke-WithRetry -Context "Copy src->temp" -Action {
            [System.IO.File]::Copy($srcLong, $tempFile, $true)
        } | Out-Null

        try {
            $tempLong = Add-LongPathPrefix $tempFile
            $tempAttrs = [System.IO.File]::GetAttributes($tempLong)
            if ($tempAttrs -band [System.IO.FileAttributes]::ReadOnly) {
                [System.IO.File]::SetAttributes($tempLong, $tempAttrs -band (-bnot [System.IO.FileAttributes]::ReadOnly))
                Write-DetailedLog "ReadOnly-Attribut auf Temp-Kopie entfernt: $tempFile" "DEBUG"
            }
        } catch {
            Write-DetailedLog "ReadOnly-Attribut-Reset auf Temp-Kopie fehlgeschlagen: $_" "WARN"
        }

        # --- Zone.Identifier (ADS) auf Temp entfernen ---
        # Reihenfolge wichtig: Unblock-File triggert ggf. AV-Scan. Wir wollen
        # diesen Trigger VOR Wait-FileAvailable haben, damit Wait-FileAvailable
        # das letzte Gate vor ZipFile.Open ist.
        try { Unblock-File -LiteralPath (Remove-LongPathPrefix $tempFile) -ErrorAction SilentlyContinue -WhatIf:$false } catch {}

        if (-not (Wait-FileAvailable -Path (Add-LongPathPrefix $tempFile))) {
            $stats.Skipped++
            $stats.Locked++
            if (-not $script:UseProgress) { Write-Host "-> SKIP (Temp-Datei blockiert)" -ForegroundColor Yellow }
            Write-Log         "Temp-Datei blockiert nach Copy (AV-Lock?): $($file.FullName)" "WARN"
            Write-DetailedLog "Wait-FileAvailable Timeout: $tempFile" "WARN"
            Write-CsvLog -Status "SKIP" -Actions "Temp-Datei blockiert (AV-Lock?)" -Path $file.FullName
            return
        }

        # --- Konvertierung alter Formate (.doc / .dot) ---
        if ($file.Extension -match "^\.do[ct]$") {
            $isOleEncrypted = Test-IsOleEncrypted -FilePath $tempFile

            if ($isOleEncrypted -and $script:Passwords.Count -eq 0) {
                $stats.Encrypted++
                $stats.Skipped++
                if (-not $script:UseProgress) { Write-Host "-> SKIP (verschlüsselt, kein PW)" -ForegroundColor Magenta }
                Write-Log "Übersprungen (verschlüsselt, kein Passwort): $($file.FullName)" "WARN"
                Write-DetailedLog "OLE-Datei ist verschlüsselt, keine Passwörter verfügbar: $($file.FullName)" "WARN"
                Write-CsvLog -Status "SKIP" -Actions "Verschluesselt, kein Passwort" -Path $file.FullName
                return
            }

            if ($isOleEncrypted) {
                Write-DetailedLog "OLE-Datei erkannt als verschlüsselt: $($file.FullName)" "DEBUG"
            }

            try {
                if (-not $script:UseProgress) { Write-Host -NoNewline "[Conv] " -ForegroundColor Cyan }
                $convResult = Convert-DocToDocx -SourcePath $tempFile -DestPathBase $tempBase -Passwords $script:Passwords
                $workFile   = $convResult.Path
                if ($convResult.UsedPassword) { Move-PasswordToFront $convResult.UsedPassword }
                if ($convResult.UsedRepair) {
                    $convRepaired = $true
                    $stats.Repaired++
                    if (-not $script:UseProgress) { Write-Host -NoNewline "[Repariert] " -ForegroundColor Yellow }
                    Write-Log         "Word-Reparatur genutzt (OpenAndRepair): $($file.FullName)" "WARN"
                    Write-DetailedLog "Datei liess sich nur mit OpenAndRepair oeffnen - Inhalt bitte stichprobenartig pruefen: $($file.FullName)" "WARN"
                }
                $wasConverted = $true
                $stats.Converted++
                Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false
                Write-DetailedLog "Konvertiert: $($file.FullName) -> $workFile" "DEBUG"
            } catch {
                $msg      = $_.Exception.Message
                $category = Get-ConversionErrorCategory -Message $msg
                if ($category -in @("Timeout", "Passwort")) {
                    Cleanup-AllWord
                    $stats.Skipped++
                    if ($category -eq 'Passwort') { $stats.Encrypted++ } else { $stats.Timeouts++ }
                    if (-not $script:UseProgress) { Write-Host "-> SKIP ($category)" -ForegroundColor Magenta }
                    Write-Log "Übersprungen ($category): $($file.FullName) - $_" "WARN"
                    Write-CsvLog -Status "SKIP" -Actions $category -Path $file.FullName -Details $msg
                } else {
                    $stats.Errors++
                    if (-not $script:UseProgress) { Write-Host "-> ERR ($category)" -ForegroundColor Red }
                    Write-Log         "Konvert-Fehler [$category]: $($file.FullName) - $_" "ERROR"
                    Write-DetailedLog "Konvert-Fehler Detail: $_" "ERROR"
                    Write-CsvLog -Status "ERR" -Actions "Konvert: $category" -Path $file.FullName -Details $msg
                }
                return
            }
        } else {
            $dest     = "$tempBase$($file.Extension)"
            $workFile = $dest
        }

        # --- OOXML-Verschlüsselungs-Check + Passwort-Entfernung ---
        # Muss VOR dem if stehen: der Wert geht unten in die
        # Rueckschreib-Bedingung ein und darf dort nicht undefiniert sein.
        $pwRemoved  = $false
        $isOoxmlExt = ($file.Extension -match "^\.do[ct][xm]$") -and (-not $wasConverted)
        if ($isOoxmlExt -and -not (Test-IsValidZip -FilePath $workFile)) {
            if ($script:Passwords.Count -eq 0) {
                $stats.Encrypted++
                $stats.Skipped++
                if (-not $script:UseProgress) { Write-Host "-> SKIP (verschlüsselt, kein PW)" -ForegroundColor Magenta }
                Write-Log         "Übersprungen (verschlüsselt, kein Passwort): $($file.FullName)" "WARN"
                Write-DetailedLog "OOXML-Datei verschlüsselt, keine Passwörter verfügbar: $($file.FullName)" "WARN"
                Write-CsvLog -Status "SKIP" -Actions "Verschluesselt, kein Passwort" -Path $file.FullName
                return
            }
            if (-not $script:UseProgress) { Write-Host -NoNewline "[PW] " -ForegroundColor Cyan }
            $pwResult  = Remove-OpenPassword -FilePath $workFile -Passwords $script:Passwords -OriginalExtension $file.Extension
            $pwRemoved = $pwResult.Success
            if ($pwRemoved -and $pwResult.UsedPassword) { Move-PasswordToFront $pwResult.UsedPassword }
            if (-not $pwRemoved) {
                $stats.Skipped++
                $stats.Encrypted++
                if (-not $script:UseProgress) { Write-Host "-> SKIP (Passwort)" -ForegroundColor Magenta }
                Write-Log         "Übersprungen (Passwort): $($file.FullName)" "WARN"
                Write-DetailedLog "Passwort-Entfernung fehlgeschlagen: $($file.FullName)" "WARN"
                Write-CsvLog -Status "SKIP" -Actions "Passwort nicht erkannt" -Path $file.FullName
                return
            }
            Write-DetailedLog "Öffnen-Passwort entfernt: $($file.FullName)" "DEBUG"
        }

        # --- Schutz per XML entfernen ---
        $actions = Remove-WordProtection -FilePath $workFile
        Write-DetailedLog "Aktionen ($($file.Name)): $($actions -join ', ')" "DEBUG"

        # --- Rückschreiben ---
        # $pwRemoved MUSS mit in die Bedingung: Remove-OpenPassword hat die
        # Entschluesselung nur auf der Temp-Kopie ($workFile) vorgenommen.
        # Findet Remove-WordProtection danach keinen XML-Schutz - der Normalfall,
        # denn die Datei war ja nur mit einem Oeffnen-Passwort verschluesselt und
        # nicht zusaetzlich formulargeschuetzt -, ist $actions leer und
        # $wasConverted fuer .docx/.docm ohnehin $false. Ohne $pwRemoved lief
        # dann der else-Zweig: die entschluesselte Temp-Datei wurde im finally
        # geloescht, die Datei auf der Ablage blieb verschluesselt, und
        # protokolliert wurde 'Kein Schutz gefunden'. Der Betreiber hielt die
        # Ablage danach fuer passwortfrei, obwohl sie es nicht war.
        if (@($actions).Count -gt 0 -or $wasConverted -or $pwRemoved) {

            # Vorspann zusammensetzen statt sequentiell ueberschreiben, damit
            # keine Aktion verlorengeht (Konvertierung und Passwort schliessen
            # sich derzeit aus, weil $isOoxmlExt '-not $wasConverted' verlangt -
            # darauf soll sich die Meldung aber nicht verlassen muessen).
            $vorspann = @()
            if ($wasConverted) { $vorspann += 'Konvertierung' }
            if ($pwRemoved)    { $vorspann += 'Öffnen-Passwort' }
            $actionStrPre = (@($vorspann) + @($actions)) -join ', '

            # Zentrale Freigabe: ab hier wird das Original angefasst. Bei
            # -WhatIf bzw. im Simulationsmodus endet die Verarbeitung hier,
            # die Temp-Kopie raeumt der finally-Block auf.
            if (-not (Confirm-Write (Remove-LongPathPrefix $srcLong) "Schutz entfernen ($actionStrPre)")) {
                if ($wasConverted) { $stats.Converted-- }
                $stats.WouldChange++
                if (-not $script:UseProgress) {
                    Write-Host "-> WHATIF ($actionStrPre)" -ForegroundColor DarkCyan
                }
                Write-Log "Simulation: würde ändern: $($file.FullName)  [$actionStrPre]" "INFO"
                Write-CsvLog -Status "WHATIF" -Actions $actionStrPre -Path $file.FullName
                return
            }

            $newExt     = [System.IO.Path]::GetExtension($workFile)
            $finalDest  = [System.IO.Path]::ChangeExtension($srcLong, $newExt)

            # Namenskollisions-Schutz: existiert beim Formatwechsel bereits
            # eine fremde Datei mit dem Zielnamen, wird ein nummerierter
            # Ausweichname gewaehlt, statt sie via Copy(..., $true)
            # stillschweigend zu ueberschreiben (analog 3a/3b/3c).
            if ($wasConverted -and ($finalDest -ne $srcLong) -and [System.IO.File]::Exists($finalDest)) {
                $collisionDir  = [System.IO.Path]::GetDirectoryName($finalDest)
                $collisionStem = [System.IO.Path]::GetFileNameWithoutExtension($finalDest)
                $candidate = Reserve-UniqueDestination `
                                -BasePath  ([System.IO.Path]::Combine($collisionDir, $collisionStem)) `
                                -Extension $newExt
                if (-not $candidate) {
                    throw "Mehr als $MaxNameClashRetries Namenskollisionen - Abbruch zum Schutz fremder Dateien."
                }
                $finalDest = $candidate
                $reservedPlaceholder = $candidate
                if (-not $script:UseProgress) {
                    Write-Host -NoNewline "[Ziel: $([System.IO.Path]::GetFileName($finalDest))] " -ForegroundColor Yellow
                }
                Write-Log "Namenskollision gelöst: $($file.FullName) -> $(Remove-LongPathPrefix $finalDest)" "WARN"
            }

            $backupPath = "$srcLong.bak"

            if ([System.IO.File]::Exists($backupPath)) {
                # Ein zurueckgebliebenes .bak stammt aus einem harten Abbruch
                # eines frueheren Laufs und kann die einzige intakte Kopie des
                # Originals sein (Crash mitten im Copy work->final trunkiert
                # das Original). Daher beiseite legen statt ueberschreiben.
                $staleBak = "$srcLong.bak_$([Guid]::NewGuid().ToString('N').Substring(0,6))"
                try {
                    [System.IO.File]::Move($backupPath, $staleBak)
                    Write-Log "Veraltetes Backup beiseite gelegt (nicht gelöscht): $(Remove-LongPathPrefix $staleBak)" "WARN"
                } catch {
                    Write-DetailedLog "Veraltetes Backup nicht verschiebbar (wird überschrieben): $_" "WARN"
                }
            }

            Invoke-WithRetry -Context "Backup src" -Action {
                [System.IO.File]::Copy($srcLong, $backupPath, $true)
            } | Out-Null
            # Erzeugungszeit ausdruecklich stempeln (wie 2b): Remove-OrphanedBackups
            # misst das Alter daran. Musste ein vorhandenes .bak oben
            # ueberschrieben werden (Move fehlgeschlagen), behielte die Datei
            # sonst die ALTE CreationTime.
            try { [System.IO.File]::SetCreationTime($backupPath, (Get-Date)) } catch {}

            try {
                Invoke-WithRetry -Context "Copy work->final" -Action {
                    [System.IO.File]::Copy($workFile, $finalDest, $true)
                } | Out-Null

                if (-not [System.IO.File]::Exists($finalDest)) { throw "Verifizierung fehlgeschlagen" }
                $reservedPlaceholder = $null   # Ziel ist real geschrieben

                # --- Zone.Identifier entfernen ---
                try {
                    $unblockPath = Remove-LongPathPrefix $finalDest
                    Remove-Item -LiteralPath $unblockPath -Stream Zone.Identifier -ErrorAction Stop
                } catch {
                    try { Unblock-File -LiteralPath $unblockPath -ErrorAction SilentlyContinue } catch {}
                }

                # --- ACL/Owner des Originals auf die neue Datei uebertragen ---
                # (Admin-Kontext: vollstaendig; Nutzer-Kontext: DACL)
                Set-FileSecuritySnapshot -Path $finalDest -Snapshot $origSecurity

                # --- Backup + ggf. alte .doc entfernen ---
                try { [System.IO.File]::Delete($backupPath) } catch {}
                if ($wasConverted -and $srcLong -ne $finalDest -and [System.IO.File]::Exists($srcLong)) {
                    # Mit Retry und Protokoll wie 2b. Vorher lief der Delete in
                    # ein leeres catch: hielt z. B. der Virenscanner die alte .doc
                    # kurz offen, blieb sie ohne jeden Hinweis neben der neuen
                    # .docx liegen - und der naechste Lauf konvertierte sie
                    # erneut, in eine '_2'-Variante (Doppel auf der Ablage).
                    try {
                        Invoke-WithRetry -Context "Delete original" -Action {
                            [System.IO.File]::Delete($srcLong)
                        } | Out-Null
                    } catch {
                        Write-DetailedLog "Alte Quelldatei konnte nicht geloescht werden (wird beibehalten): $srcLong - $_" "WARN"
                        Write-Log "Quelldatei beibehalten: $($file.FullName) - $_" "WARN"
                    }
                }

                # --- Zeitstempel wiederherstellen ---
                # Auf Netzlaufwerken kann der serverseitige AV-Scanner die frisch
                # geschriebene Datei kurz sperren. Wir warten kurz, bevor wir die
                # Zeitstempel zurücksetzen. Bei Timeout trotzdem versuchen – das
                # try/catch unten fängt Reststörungen ab, dann gehen lediglich die
                # Zeitstempel verloren (Datei selbst ist sicher).
                $null = Wait-FileAvailable -Path $finalDest
                try {
                    [System.IO.File]::SetCreationTimeUtc($finalDest,   $origCreationTime.ToUniversalTime())
                    [System.IO.File]::SetLastWriteTimeUtc($finalDest,  $origLastWriteTime.ToUniversalTime())
                    [System.IO.File]::SetLastAccessTimeUtc($finalDest, $origLastAccessTime.ToUniversalTime())
                } catch {}

                $zurueckgeschrieben = $true
                $stats.Unlocked++
                $actionStr = $actions -join ', '
                if (-not $script:UseProgress) {
                    Write-Host "-> OK ($actionStr)" -ForegroundColor Green
                }
                Write-Log "Erledigt: $($file.FullName)  [$actionStr]" "INFO"
                $csvDetails = Remove-LongPathPrefix $finalDest
                if ($convRepaired) { $csvDetails += "  [mit OpenAndRepair geoeffnet]" }
                Write-CsvLog -Status "OK" -Actions $actionStr -Path $file.FullName -Details $csvDetails

            } catch {
                $stats.Errors++
                # Wenn die Datei zuvor erfolgreich von .doc nach .docx konvertiert
                # wurde ($wasConverted=true, $stats.Converted++), aber das
                # Zurueckschreiben hier scheitert, wird der Rollback unten
                # durchgefuehrt - die Datei landet wieder als Original auf der
                # Platte. Der Converted-Counter muss deshalb hier korrigiert
                # werden. Der globale catch ganz aussen (mit identischer
                # Korrekturlogik) wird NICHT erreicht, weil dieser innere
                # catch den Fehler absorbiert.
                if ($wasConverted) { $stats.Converted-- }
                if ($finalDest -and [System.IO.File]::Exists($finalDest)) {
                    try { [System.IO.File]::Delete($finalDest) } catch {}
                    Write-DetailedLog "Unvollständige Zieldatei entfernt: $finalDest" "WARN"
                }
                if ([System.IO.File]::Exists($backupPath)) {
                    $backupSafeToDelete = $true
                    if ($finalDest -eq $srcLong) {
                        $backupSafeToDelete = $false
                        try {
                            Invoke-WithRetry -Context "Rollback from Backup" -Action {
                                [System.IO.File]::Copy($backupPath, $srcLong, $true)
                            } | Out-Null
                            if ([System.IO.File]::Exists($srcLong)) {
                                $backupSafeToDelete = $true
                                Set-FileSecuritySnapshot -Path $srcLong -Snapshot $origSecurity
                                Write-DetailedLog "Rollback erfolgreich: Original aus Backup wiederhergestellt: $srcLong" "WARN"
                            }
                        } catch {
                            Write-Log "KRITISCH: Rollback fehlgeschlagen - Backup bleibt erhalten: $backupPath" "ERROR"
                            Write-DetailedLog "Rollback-Kopie fehlgeschlagen: $_" "ERROR"
                        }
                    }
                    if ($backupSafeToDelete) {
                        try { [System.IO.File]::Delete($backupPath) } catch {}
                    }
                }
                if (-not $script:UseProgress) { Write-Host "-> ERR (Schreiben)" -ForegroundColor Red }
                Write-Log         "Schreib-Fehler: $($file.FullName) - $_" "ERROR"
                Write-DetailedLog "Schreib-Fehler Detail: $_" "ERROR"
                Write-CsvLog -Status "ERR" -Actions "Schreib-Fehler" -Path $file.FullName -Details $_.Exception.Message
            }

        } else {
            if (-not $script:UseProgress) { Write-Host "-> Kein Schutz" -ForegroundColor Gray }
            Write-CsvLog -Status "NO_CHANGE" -Actions "Kein Schutz gefunden" -Path $file.FullName
        }

    } catch {
        $stats.Errors++
        if ($wasConverted) { $stats.Converted-- }
        # KEIN blindes Rollback hier! Wenn der aeussere catch greift, ist
        # $srcLong byte-identisch unveraendert (alle bisherigen Operationen
        # lesen entweder nur oder schreiben in $tempFile/$workFile bzw.
        # $backupPath). Die einzige Operation, die das Original tatsaechlich
        # ueberschreibt, ist Copy work->final im INNEREN try-Block - der
        # innere catch (oben) rollback-t das bereits sauber.
        #
        # Der frueher hier stehende Rollback war gefaehrlich: Wenn die
        # Backup-Erstellung selbst (Invoke-WithRetry "Backup src", Z. 1344)
        # mitten im Copy bricht - z.B. Netzwerk-Drop, Quota voll, AV-Lock -
        # liegt $backupPath als partial/0-Byte-Datei auf der Platte. [Exists]
        # liefert true. Der alte Code hat dieses korrupte Backup blind ueber
        # das INTAKTE Original kopiert -> unwiderruflicher Datenverlust.
        #
        # Backup wird hier nicht geloescht, damit der Nutzer es bei Bedarf
        # manuell pruefen kann (im Normalfall ist das Original ja intakt,
        # das Backup also ueberfluessig; bei Hyper-Edge-Cases koennte es
        # aber wertvoll sein).
        if ($backupPath -and [System.IO.File]::Exists($backupPath)) {
            Write-DetailedLog "Hinweis: Backup zur manuellen Pruefung erhalten: $backupPath" "WARN"
        }
        if (-not $script:UseProgress) { Write-Host "-> ERROR: $($_.Exception.Message)" -ForegroundColor Red }
        Write-Log         "Allg. Fehler: $($file.FullName) - $_" "ERROR"
        Write-DetailedLog "Allg. Fehler Detail: $_" "ERROR"
        Write-CsvLog -Status "ERR" -Actions "Allg. Fehler" -Path $file.FullName -Details $_.Exception.Message

    } finally {
        if (-not [string]::IsNullOrWhiteSpace($workFile)) {
            Remove-Item -LiteralPath $workFile -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false
        }
        # Oben entferntes ReadOnly-Attribut zuruecksetzen, wenn die Datei
        # nicht ersetzt wurde (Kein Schutz, SKIP, WHATIF, Fehler/Rollback).
        if ($roEntfernt -and -not $zurueckgeschrieben -and [System.IO.File]::Exists($srcLong)) {
            try {
                [System.IO.File]::SetAttributes($srcLong, [System.IO.File]::GetAttributes($srcLong) -bor [System.IO.FileAttributes]::ReadOnly)
                Write-DetailedLog "Schreibschutz (Attribut) wiederhergestellt, Datei unveraendert: $($file.FullName)" "DEBUG"
            } catch {
                Write-DetailedLog "Schreibschutz-Attribut nicht wiederherstellbar: $($file.FullName) - $_" "WARN"
            }
        }
        # Reservierten Platzhalter nur entfernen, wenn er leer geblieben ist -
        # eine Datei mit Inhalt wird niemals angefasst.
        if ($reservedPlaceholder) {
            try {
                $ph = [System.IO.FileInfo]::new($reservedPlaceholder)
                if ($ph.Exists -and $ph.Length -eq 0) {
                    [System.IO.File]::Delete($reservedPlaceholder)
                    Write-DetailedLog "Reservierten Platzhalter entfernt: $reservedPlaceholder" "DEBUG"
                }
            } catch {}
        }
    }
}
} catch [System.OperationCanceledException] {
    Write-Host "`nVerarbeitung durch Nutzer abgebrochen." -ForegroundColor Yellow
    Write-Log "Verarbeitung durch Nutzer abgebrochen." "WARN"
}

try { if ($script:CtrlCAsInput) { [Console]::TreatControlCAsInput = $false; $script:CtrlCAsInput = $false } } catch {}

# ==================================================================
# AUFRÄUMEN & STATISTIK
# ==================================================================
if ($script:UseProgress) { Write-Progress -Activity "Fertig" -Completed }

Cleanup-AllWord
if (Test-Path -LiteralPath $TempPath) { Remove-Item -LiteralPath $TempPath -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false }
Clear-WindowsTempWhitelist

$duration = (Get-Date) - $ScriptStartTime
$durStr   = $duration.ToString('hh\:mm\:ss')

Write-Host "`n=====================================================" -ForegroundColor Cyan
Write-Host " FERTIG in $durStr" -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host "Geprüft:         $($stats.Processed)"   -ForegroundColor White
Write-Host "Entsperrt/Konv:  $($stats.Unlocked)"    -ForegroundColor Green
Write-Host "Konvertiert:     $($stats.Converted)"   -ForegroundColor Green
if ($stats.Repaired -gt 0) {
    Write-Host "  davon repariert: $($stats.Repaired)  (nur mit OpenAndRepair lesbar – prüfen!)" -ForegroundColor Yellow
}
if ($stats.WouldChange -gt 0) {
    Write-Host "Würde ändern:    $($stats.WouldChange)  (Simulation – nichts geschrieben)" -ForegroundColor DarkCyan
}
if ($backupCleanup) {
    Write-Host "Backups entf.:   $($backupCleanup.Deleted)" -ForegroundColor DarkGray
    if ($backupCleanup.Orphans -gt 0) {
        Write-Host "Backups o. Orig: $($backupCleanup.Orphans)  (NICHT gelöscht – prüfen!)" -ForegroundColor Yellow
    }
}
Write-Host "Junk gelöscht:   $($stats.Junk)"        -ForegroundColor DarkGray
if ($stats.JunkKept -gt 0) {
    Write-Host "Sperrdateien:    $($stats.JunkKept)  (jünger als $JunkMinAgeHours h – geschont)" -ForegroundColor DarkGray
}
Write-Host "Übersprungen:    $($stats.Skipped)"     -ForegroundColor Magenta
Write-Host "Verschlüsselt:   $($stats.Encrypted)"   -ForegroundColor Magenta
Write-Host "COM-Timeouts:    $($stats.Timeouts)"    -ForegroundColor Magenta
Write-Host "Gesperrt:        $($stats.Locked)"      -ForegroundColor Yellow
Write-Host "Fehler:          $($stats.Errors)"      -ForegroundColor Red
if ($script:RetryStats.Recovered -gt 0 -or $script:RetryStats.Failed -gt 0) {
    Write-Host "-----------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "Retry-Treffer:   $($script:RetryStats.Recovered)" -ForegroundColor DarkYellow
    Write-Host "Retry-Aufgaben:  $($script:RetryStats.Failed)"    -ForegroundColor DarkYellow
    if ($script:RetryStats.Recovered -gt 0) {
        Write-Host "                 (deutet auf AV-Scanner-Verzögerungen hin)" -ForegroundColor DarkGray
    }
}
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host "Logs:" -ForegroundColor Cyan
Write-Host "  $LogFilePath"     -ForegroundColor Gray
Write-Host "  $DetailedLogPath" -ForegroundColor Gray
Write-Host "  $CsvLogPath"      -ForegroundColor Gray
Write-Host "=====================================================" -ForegroundColor Cyan

@(
    "=================================================================="
    "Log Ende:  $(Get-Date)"
    "Dauer:     $durStr"
    "Geprüft:   $($stats.Processed)  |  Entsperrt: $($stats.Unlocked)  |  Fehler: $($stats.Errors)"
    "Konv:      $($stats.Converted)  |  Skip: $($stats.Skipped)  |  Verschlüsselt: $($stats.Encrypted)  |  Timeouts: $($stats.Timeouts)  |  Gesperrt: $($stats.Locked)"
    "Simuliert: $($stats.WouldChange)  |  Sperrdateien geschont: $($stats.JunkKept)  |  Repariert: $($stats.Repaired)"
    $(if ($backupCleanup) {
        "Backups:   entfernt=$($backupCleanup.Deleted)  behalten=$($backupCleanup.Kept)  ohne Original=$($backupCleanup.Orphans)  Fehler=$($backupCleanup.Failed)"
      } else {
        "Backups:   Aufräumen übersprungen (-SkipBackupCleanup)"
      })
    "Retry:     Recovered=$($script:RetryStats.Recovered)  Failed=$($script:RetryStats.Failed)"
    "=================================================================="
) | Add-Content -LiteralPath $LogFilePath -Encoding utf8 -WhatIf:$false -Confirm:$false

if (-not $NoInteractive) {
    Write-Host ""
    Write-Host "Beliebige Taste drücken zum Beenden..." -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
