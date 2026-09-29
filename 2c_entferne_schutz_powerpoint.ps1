<#
.SYNOPSIS
    PowerPoint Deep Unprotect
.DESCRIPTION
    Entfernt Schutz aus PowerPoint-Dateien (XML-Ebene) und konvertiert
    alte .ppt/.pps/.pot -> .pptx/.ppsx/.potx. Netzlaufwerk- und
    Langpfad-faehig. Robust gegen AV-Locks durch Retry, Priming-Read
    und Stabilitaets-Polling der Temp-Kopie.

    Voraussetzung Vertrauenswuerdige Speicherorte:
      Das Skript verwendet den Ordner
      "%LOCALAPPDATA%\Dateimigration-Arbeitskopien\2c_entferne_schutz_powerpoint_<PID>"
      als temporaeren Arbeitsordner fuer die COM-Automatisierung (bis
      29.09.2026 "Dokumente" - der wird per OneDrive-Ordnersicherung oft
      synchronisiert, jede Arbeitskopie ging dann in die Cloud). Damit
      PowerPoint Dateien dort ohne Geschuetzte Ansicht oeffnet, diesen
      Ordner als vertrauenswuerdigen Speicherort eintragen:
        PowerPoint -> Datei -> Optionen -> Trust Center -> Einstellungen...
          -> Vertrauenswuerdige Speicherorte
          -> "Neuen Speicherort hinzufuegen..."
          -> Pfad: C:\Users\<Benutzername>\AppData\Local\Dateimigration-Arbeitskopien
          -> Haken: "Unterordner dieses Speicherorts sind ebenfalls vertrauenswuerdig"
      Bei Netzlaufwerken als Zielpfad zusaetzlich:
          -> "Vertrauenswuerdige Speicherorte im Netzwerk zulassen"
      Das Skript fuehrt vor dem Hauptlauf einen PowerPoint-Smoke-Test durch
      und bricht bei Fehlschlag ab (bzw. fragt interaktiv nach), um stunden-
      lange Timeouts pro Datei zu vermeiden.

    PowerPoint muss geschlossen sein:
      PowerPoint ist ein Einzelinstanz-COM-Server. Laeuft beim Start bereits
      ein POWERPNT-Prozess (auch ohne Fenster), arbeitet das Skript nicht per
      COM: interaktiv wird nach dem Schliessen erneut geprueft, mit
      -NoInteractive endet der Lauf mit Exitcode 3. Waehrend des Laufs
      PowerPoint nicht starten - erkennt ein Job eine fremde Instanz, bricht
      der Lauf ebenfalls mit Exitcode 3 ab.

    Hinweis zu Passwoertern in PowerPoint:
      Anders als Word.Documents.Open und Excel.Workbooks.Open akzeptiert
      PowerPoint.Presentations.Open KEINEN Password-Parameter. Verschluesselte
      Praesentationen (OLE/CFBF mit EncryptedPackage) blockieren in COM auf
      einem Passwort-Dialog und werden vom Skript per Job-Timeout uebersprungen.
      Modify-Passwoerter (WritePassword) werden durch das ReadOnly-Open umgangen
      und beim SaveAs durch hardcoded Reset auf "" entfernt. Eine interaktive
      Passwort-Abfrage entfaellt deshalb bewusst.

    Compilieren in eine .exe (Windows PowerShell 5.1, STA-Konsolenanwendung):
      # Einmalig: ps2exe installieren
      Install-Module -Name ps2exe -Scope CurrentUser -Force

      # Hilfscheck: Office-Bitness ermitteln (optional zur Optimierung)
      #   Rueckgabe 'x64' -> -x64 verwenden ; 'x86' -> -x86 verwenden
      Get-ItemProperty 'HKLM:\Software\Microsoft\Office\ClickToRun\Configuration' -Name Platform -ErrorAction SilentlyContinue

      # Hinweis zu -x64 / -x86 bei gemischten Office-Bitness-Umgebungen:
      # PowerPoint registriert seinen COM-Server als LocalServer32 (out-of-process,
      # POWERPNT.EXE). Das Windows-COM-Subsystem marshallt Aufrufe zwischen
      # x64-Aufrufer und x86-Server (und umgekehrt) automatisch ueber DCOM/LRPC.
      # Daher funktioniert eine mit -x64 kompilierte EXE auch mit 32-Bit-PowerPoint
      # (und eine -x86-EXE mit 64-Bit-PowerPoint). Bitness-Match liefert minimal
      # bessere Performance, ist aber nicht erforderlich. -x64 ist eine sichere
      # Default-Wahl, da modernes Office (2019+, M365) standardmaessig x64 ist.

      Invoke-ps2exe -inputFile '.\2c_entferne_schutz_powerpoint.ps1' -outputFile '.\2c_entferne_schutz_powerpoint.exe' -iconFile '.\powershell_icon.ico' -title 'Powerpoint Deep Unprotect' -description 'Entfernt Schutz aus Powerpoint-Dateien' -product 'Powerpoint Deep Unprotect' -version '1.0.1.0' -STA -x64 -supportOS

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
# Logs liegen direkt neben Skript/EXE. PSScriptRoot greift bei direktem
# Aufruf; bei ps2exe-EXEs ist PSScriptRoot leer - dann liefert die
# Entry-Assembly das EXE-Verzeichnis. $PWD nur als letzter Fallback.
#
# Der frueher hier stehende Einzeiler fiel bei einer ps2exe-EXE sofort auf
# $PWD zurueck, also auf das aktuelle Arbeitsverzeichnis: bei Doppelklick
# aus dem Explorer zufaellig, bei einer Verknuepfung frei einstellbar.
# Haupt-Log, Detail-Log und Ergebnis-CSV landeten damit an wechselnden
# Orten - und das, obwohl der Skriptkopf die ps2exe-Uebersetzung
# ausdruecklich dokumentiert. Gleiche Fassung wie in 2a, 6, 7, 8 und 9.
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

$script:scriptDir              = Resolve-ScriptDirectory

# Log-Verzeichnis: bevorzugt neben dem Skript, sonst LOCALAPPDATA, sonst TEMP.
# Vorher wurde ungeprueft in $scriptDir geschrieben; auf einer schreibgeschuetzten
# Freigabe scheiterte Initialize-Loggers still und der Lauf lief ohne Protokoll.
function Resolve-LogDirectory {
    param([string]$Preferred)
    foreach ($cand in @($Preferred,
                        (Join-Path $env:LOCALAPPDATA 'PptDeepUnprotect'),
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
$script:LogDir                 = Resolve-LogDirectory -Preferred $script:scriptDir

# Gemeinsamer Arbeitsordner der Office-Skripte, NICHT "Dokumente" (siehe
# Kopf: OneDrive-Ordnersicherung laedt dort jede Arbeitskopie hoch).
$script:ArbeitsBasis           = Join-Path $(if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { $env:TEMP }) 'Dateimigration-Arbeitskopien'
$script:TempPath               = Join-Path $script:ArbeitsBasis "2c_entferne_schutz_powerpoint_$PID"
$script:RunTimestamp           = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$script:LogFilePath            = Join-Path $script:LogDir "2c_entferne_schutz_powerpoint_$($script:RunTimestamp).log"
$script:DetailedLogPath        = Join-Path $script:LogDir "2c_entferne_schutz_powerpoint_detailed_$($script:RunTimestamp).log"
$script:CsvLogPath             = Join-Path $script:LogDir "2c_entferne_schutz_powerpoint_$($script:RunTimestamp).csv"

# Mindestalter fuer das Loeschen von '~$'-Sperrdateien und fuer die
# Einstufung eines '.bak' als verwaist.
$script:JunkMinAgeHours          = 24
$script:BackupCleanupMinAgeHours = 24

# Verzeichnisse, die bei der Suche nicht betreten werden.
$script:ExcludeDirNames = @(
    '$RECYCLE.BIN', 'System Volume Information', 'RECYCLER',
    '.git', '.svn', '__pycache__', 'node_modules'
)

# Verzeichnis-Presets fuer die Startauswahl.
$script:DirectoryPresets = @(
    'Q:\'
    'R:\'
    'G:\Geteilte Ablagen'
    'G:\Meine Ablage'
    '\\server\dfs'
)
$script:FileOpenTimeoutSeconds = 45
$script:TrustCenterTimeoutSec  = 25
$script:ScriptStartTime        = Get-Date
$script:MaxDirectoryDepth      = 200
$script:MaxNameClashRetries    = 100
$script:AvStableMaxAttempts    = 12
$script:AvStableDelayMs        = 350

$script:Utf8Bom = New-Object System.Text.UTF8Encoding($true)

$script:ShouldStop      = $false
# Gesetzt, sobald ein COM-Job eine PowerPoint-Sitzung des Anwenders erkennt.
$script:FremdePptAbbruch = $false
$script:UseProgress     = $false
$script:SkipPreScan     = $false
$script:TotalFiles      = 0
$script:ProcessedCount  = 0
$script:TrackedPptPids  = [System.Collections.Generic.List[int]]::new()
$script:LogWriter       = $null
$script:DebugLogWriter  = $null
$script:CsvLogWriter    = $null
# $PSCmdlet ist nur im Skript-Scope verfuegbar, nicht in einfachen Funktionen.
$script:ScriptCmdlet    = $PSCmdlet
$script:LastProgressUpdate    = [DateTime]::MinValue
$script:ProgressMinIntervalMs = 500
$script:RestorePrivilegesEnabled = $false

# ==================================================================
# POWERPOINT-FORMAT-KONSTANTEN
# ==================================================================
$script:PptConstants = @{
    ppSaveAsOpenXMLPresentation             = 24
    ppSaveAsOpenXMLPresentationMacroEnabled = 25
    ppSaveAsOpenXMLTemplate                 = 26
    ppSaveAsOpenXMLTemplateMacroEnabled     = 27
    ppSaveAsOpenXMLShow                     = 28
    ppSaveAsOpenXMLShowMacroEnabled         = 29
    msoTrue                                 = -1
    ppWindowMinimized                       =  2
    # 1, nicht 2: laut Typbibliothek (MSPPT.OLB, PpAlertLevel) ist
    # ppAlertsNone = 1 und ppAlertsAll = 2. Mit 2 wurden Warnungen also
    # gerade NICHT unterdrueckt (3c, 4c, 7 und 9 setzen 1).
    ppAlertsNone                            =  1
    msoFeatureInstallNone                   =  0
    msoAutomationSecurityForceDisable       =  3
}

# ==================================================================
# INIT & LIBRARIES
# ==================================================================
try {
    Add-Type -AssemblyName "System.IO.Compression.FileSystem" -ErrorAction Stop
} catch {
    [System.Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") | Out-Null
}

# ==================================================================
# LOGGING (persistente Writer mit AutoFlush)
# ==================================================================
function Initialize-Loggers {
    try {
        $script:LogWriter = [System.IO.StreamWriter]::new($script:LogFilePath, $false, $script:Utf8Bom)
        $script:LogWriter.AutoFlush = $true
        $script:LogWriter.WriteLine("Log Start: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  |  Ziel: $TargetPath")
    } catch { $script:LogWriter = $null }
    try {
        $script:DebugLogWriter = [System.IO.StreamWriter]::new($script:DetailedLogPath, $false, $script:Utf8Bom)
        $script:DebugLogWriter.AutoFlush = $true
        $script:DebugLogWriter.WriteLine("Debug-Log Start: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    } catch { $script:DebugLogWriter = $null }
    try {
        $script:CsvLogWriter = [System.IO.StreamWriter]::new($script:CsvLogPath, $false, $script:Utf8Bom)
        $script:CsvLogWriter.AutoFlush = $true
        $script:CsvLogWriter.WriteLine("Zeitstempel;Status;Aktionen;Pfad;Details")
    } catch { $script:CsvLogWriter = $null }
}

function Close-Loggers {
    if ($script:LogWriter)      { try { $script:LogWriter.Close() }      catch {}; $script:LogWriter = $null }
    if ($script:DebugLogWriter) { try { $script:DebugLogWriter.Close() } catch {}; $script:DebugLogWriter = $null }
    if ($script:CsvLogWriter)   { try { $script:CsvLogWriter.Close() }   catch {}; $script:CsvLogWriter = $null }
}

function Invoke-WindowsTempCleanup {
    # Whitelist-Cleanup von %LOCALAPPDATA%\Temp am Skriptende: Es werden
    # AUSSCHLIESSLICH Eintraege mit bekannten Praefixen entfernt (Office-/
    # COM-Reste sowie eigene Skript-Artefakte). NIEMALS pauschal leeren -
    # Fremdprozesse legen aktive Daten ohne Lock in %TEMP% ab; blindes
    # Loeschen zerstoert sie.
    $winTemp = $env:TEMP
    if (-not $winTemp) { $winTemp = $env:TMP }
    if (-not $winTemp -or -not (Test-Path -LiteralPath $winTemp)) { return }

    $prefixes = @('~$', '~df', 'gen_py', 'vbe', 'excel8.0', '2c_entferne_schutz_powerpoint')
    $removedFiles = 0
    $removedDirs  = 0
    $skipped      = 0
    Get-ChildItem -LiteralPath $winTemp -Force -ErrorAction SilentlyContinue | ForEach-Object {
        $low = $_.Name.ToLowerInvariant()
        $hit = $false
        foreach ($p in $prefixes) {
            if ($low.StartsWith($p)) { $hit = $true; break }
        }
        if (-not $hit -and $low.StartsWith('cvr') -and $low.EndsWith('.tmp')) { $hit = $true }
        if (-not $hit -and $low.StartsWith('ppt') -and $low.EndsWith('.tmp')) { $hit = $true }
        if (-not $hit) { $skipped++; return }
        try {
            if ($_.PSIsContainer) {
                Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false
                if (-not (Test-Path -LiteralPath $_.FullName)) { $removedDirs++ }
            } else {
                Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false
                if (-not (Test-Path -LiteralPath $_.FullName)) { $removedFiles++ }
            }
        } catch {}
    }
    Write-DetailedLog "Windows-Temp-Cleanup (Whitelist): $removedFiles Dateien, $removedDirs Ordner aus $winTemp entfernt, $skipped fremde Eintraege unangetastet." "DEBUG"
}

function Remove-StaleTempFolders {
    # Heutiger Arbeitsordner UND "Dokumente" (Reste von Laeufen vor dem
    # Umzug des Arbeitsordners am 29.09.2026).
    foreach ($p in @($script:ArbeitsBasis, [Environment]::GetFolderPath('MyDocuments'))) {
        Remove-StaleTempFoldersIn -Parent $p
    }
}

function Remove-StaleTempFoldersIn {
    # Verwaiste Arbeitsordner abgebrochener Laeufe entfernen - nur eigene
    # 2c_entferne_schutz_powerpoint_<PID>-Ordner, deren PID nicht mehr lebt.
    param([string]$Parent)
    $parent = $Parent
    if ([string]::IsNullOrWhiteSpace($parent)) { return }
    if (-not (Test-Path -LiteralPath $parent)) { return }
    $candidates = @()
    try {
        $candidates = @(Get-ChildItem -LiteralPath $parent -Directory `
                                      -Filter "2c_entferne_schutz_powerpoint_*" `
                                      -Force -ErrorAction SilentlyContinue)
    } catch {}
    foreach ($d in $candidates) {
        try {
            if ($d.FullName -eq $script:TempPath) { continue }
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

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    if (-not $script:LogWriter) { return }
    try { $script:LogWriter.WriteLine("[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message") } catch {}
}

function Write-DetailedLog {
    param([string]$Message, [string]$Level = "DEBUG")
    if (-not $script:DebugLogWriter) { return }
    try { $script:DebugLogWriter.WriteLine("[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff')] [$Level] $Message") } catch {}
}

function Write-CsvLog {
    <#
        Eine Zeile je verarbeiteter Datei. Semikolon und UTF-8 mit BOM,
        damit Excel die Datei im deutschen Gebietsschema direkt oeffnet.
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
            Write-Laufprotokoll -Skript '2c_entferne_schutz_powerpoint' `
                -Pfad $Path -Aktion 'Schutz entfernen' `
                -Status $Status -Detail $Details
        } catch { }
    }
    if (-not $script:CsvLogWriter) { return }
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $q  = {
        param($v)
        '"' + ((("$v") -replace '"', '""') -replace "`r`n|`r|`n", ' ') + '"'
    }
    try {
        $script:CsvLogWriter.WriteLine("$ts;$Status;$(& $q $Actions);$(& $q $Path);$(& $q $Details)")
    } catch {}
}

function Get-PowerPointProzesseSitzung {
    <#
        Alle POWERPNT-Prozesse der eigenen Windows-Sitzung - MIT und OHNE
        sichtbares Fenster.

        Frueher zaehlten nur Prozesse mit Hauptfenster. Eine Instanz ohne
        Fenster (Rest eines abgebrochenen Laufs, Vorschau, Add-in-Host)
        uebernimmt 'New-Object -ComObject PowerPoint.Application' aber genauso.
        Prozesse anderer Sitzungen (Terminalserver) bleiben aussen vor: COM
        verbindet nur mit Klassenobjekten der eigenen Anmeldesitzung (nicht
        gemessen, hier nur Schutz vor einem Dauer-Abbruch auf Mehrbenutzer-
        Rechnern).
    #>
    $sid = -1
    try { $sid = (Get-Process -Id $PID -ErrorAction Stop).SessionId } catch {}
    return ,@(Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue |
              Where-Object { $sid -lt 0 -or $_.SessionId -eq $sid })
}

function Confirm-PowerPointGeschlossen {
    <#
        Vor dem ersten COM-Zugriff: laeuft PowerPoint, arbeitet das Skript
        NICHT per COM.

        PowerPoint ist ein Einzelinstanz-COM-Server. Vom Auftraggeber
        gemessen (29.09.2026): laeuft eine Sitzung des Anwenders, liefert
        'New-Object -ComObject PowerPoint.Application' genau diese Sitzung,
        und ihr HWND fuehrt auf ihre PID. Die fruehere Fassung dieser Funktion
        (Show-OfficeRunningWarning) versprach "Ihre Sitzung wird NICHT
        beendet" und liess mit -NoInteractive einfach weiterlaufen ("sie
        werden geschuetzt") - tatsaechlich hielt das Skript die fremde PID fuer
        die eigene, schloss die Sitzung per Quit() bzw. per Stop-Process
        -Force im Timeout-Zweig und verstellte DisplayAlerts und WindowState.

        Interaktiv: Hinweis, nach Enter erneut pruefen, 'A' bricht ab.
        -NoInteractive oder umgeleitete Eingabe: sofort $false (Abbruch).
        Rueckgabe: $true = kein PowerPoint aktiv, $false = abbrechen.
    #>
    param([switch]$NoInteractive)

    while ($true) {
        $procs = Get-PowerPointProzesseSitzung
        if ($procs.Count -eq 0) { return $true }
        $pids = ($procs | ForEach-Object { $_.Id }) -join ', '

        Write-Host ""
        Write-Host ("=" * 66) -ForegroundColor Yellow
        Write-Host "  POWERPOINT LAEUFT (PID: $pids)" -ForegroundColor Yellow
        Write-Host ("=" * 66) -ForegroundColor Yellow
        Write-Host "  PowerPoint laesst sich nur einmal starten. Das Skript wuerde"
        Write-Host "  sich an diese Sitzung haengen und sie beim Aufraeumen"
        Write-Host "  schliessen - ungespeicherte Praesentationen gingen verloren."
        Write-Host "  Deshalb arbeitet es nicht, solange PowerPoint laeuft."
        Write-Host ""
        Write-Host "  Bitte PowerPoint vollstaendig beenden. Eine Instanz ohne Fenster"
        Write-Host "  (POWERPNT.EXE) im Task-Manager beenden."
        Write-Host "  Waehrend des Laufs PowerPoint NICHT starten."
        Write-Host ("=" * 66) -ForegroundColor Yellow
        Write-Host ""

        # Ohne echte Konsole NICHT fragen. Read-Host blockiert dort unbegrenzt
        # (Aufgabenplanung mit angehaengter Konsole) oder liefert sofort leer.
        # Massgeblich ist die tatsaechliche Eingabefaehigkeit, nicht der
        # Hostname (ps2exe meldet 'PSRunspace-Host').
        $kannFragen = $false
        try { $kannFragen = -not [Console]::IsInputRedirected } catch { $kannFragen = $false }
        if ($NoInteractive -or -not $kannFragen) {
            return $false
        }
        $antwort = Read-Host "PowerPoint schliessen, dann Enter = erneut pruefen, A = Abbruch"
        if ($antwort -match '^\s*[AaNn]') { return $false }
    }
}

function Confirm-Write {
    <#
        Zentrale Freigabe fuer jede schreibende Operation am Original.
        Das Skript kannte bisher gar kein -WhatIf; ein Probelauf auf einer
        fremden Ablage war damit nicht moeglich.

        Wirft ShouldProcess, wird NICHT geschrieben. Vorher lieferte der
        catch-Zweig $true - gemessen unter 5.1: mit -Confirm im
        NonInteractive-Modus wirft ShouldProcess ("Lese- und
        Eingabeaufforderungsfunktionen sind nicht verfuegbar"), und das
        Skript schrieb dann jede Datei OHNE die verlangte Bestaetigung.
    #>
    param(
        [string]$Target,
        [string]$Action = 'Aendern'
    )
    if ($null -eq $script:ScriptCmdlet) { return $true }
    try {
        return $script:ScriptCmdlet.ShouldProcess($Target, $Action)
    } catch {
        if (-not $script:ConfirmFehlerGemeldet) {
            $script:ConfirmFehlerGemeldet = $true
            Write-Warning "Bestaetigung nicht moeglich ($($_.Exception.Message)) - es wird NICHTS geschrieben."
        }
        Write-Log "Bestaetigung nicht moeglich, nicht geschrieben: $Target [$Action] - $($_.Exception.Message)" "WARN"
        return $false
    }
}

# ==================================================================
# ABBRUCH PER TASTE (Strg+C / ESC)
# ==================================================================
# Kein [Console]::add_CancelKeyPress mehr: der dort registrierte Skriptblock
# laeuft auf einem Threadpool-Thread ohne PowerShell-Runspace. Gemessen unter
# 5.1 (Kindprozess, GenerateConsoleCtrlEvent, wortgleicher Handler, 4 von 4
# Laeufen): ShouldStop wurde nie gesetzt, powershell.exe endete hart mit
# Exitcode 2, und weder finally noch trap liefen - mitten im Rueckschreiben
# haette das ein halbes Original hinterlassen. Stattdessen wie in 6 und 7:
# Strg+C wird vor der Hauptschleife zur Eingabe (TreatControlCAsInput) und
# dort zusammen mit ESC abgefragt (Test-AbortRequested).
$script:CtrlCAsInput = $false
$script:ConfirmFehlerGemeldet = $false

function Test-AbortRequested {
    if (-not $script:CtrlCAsInput) { return }
    try {
        while ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            $istStrgC = ($key.Key -eq [ConsoleKey]::C -and
                         (($key.Modifiers -band [ConsoleModifiers]::Control) -ne 0))
            if ($istStrgC -or $key.Key -eq [ConsoleKey]::Escape) {
                if (-not $script:ShouldStop) {
                    Write-Host "`n!!  ABBRUCH angefordert - laufende Datei wird noch fertiggestellt..." -ForegroundColor Yellow
                }
                $script:ShouldStop = $true
            }
        }
    } catch {}
}

trap {
    Write-Warning "Unerwarteter Fehler: $_"
    try { if ($script:CtrlCAsInput) { [Console]::TreatControlCAsInput = $false } } catch {}
    Start-Sleep -Milliseconds 200
    # Die Liste gehoert NICHT in die Bedingung: Sinn dieses Blocks ist es
    # gerade, die PIDs aus den .pid-Dateien nachzutragen, wenn die Liste noch
    # LEER ist (Absturz vor dem ersten Add). PowerShell wertet eine leere
    # System.Collections.Generic.List im booleschen Kontext als $false aus
    # (nachgestellt unter 5.1) - der Block lief also nie, wenn er gebraucht
    # wurde, und die verwaisten PowerPoint-Prozesse blieben stehen.
    if (Test-Path -LiteralPath $script:TempPath) {
        Get-ChildItem -LiteralPath $script:TempPath -Filter "*.pid" -ErrorAction SilentlyContinue |
            ForEach-Object {
                try {
                    $pidContent = [System.IO.File]::ReadAllText($_.FullName).Trim()
                    if ($pidContent -match '^\d+$') {
                        $p = [int]$pidContent
                        if ($p -gt 0) { $script:TrackedPptPids.Add($p) }
                    }
                } catch {}
            }
    }
    # .Count statt Auflistungs-Wahrheitswert (siehe oben). Der trap gilt fuer
    # das ganze Skript, die Funktion wird aber erst weiter unten definiert -
    # ein Fehler davor liefe sonst in CommandNotFound (gemessen unter 5.1).
    if ($script:TrackedPptPids.Count -gt 0 -and (Get-Command Clear-TrackedPowerPointInstances -ErrorAction SilentlyContinue)) { Clear-TrackedPowerPointInstances }
    if (Test-Path -LiteralPath $script:TempPath) { Remove-Item -LiteralPath $script:TempPath -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false }
    Close-Loggers
    break
}

# ==================================================================
# HELPER: PFADE & UMGEBUNG
# ==================================================================
function Add-LongPathPrefix {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
    if ($Path -match "^\\\\\?\\")            { return $Path }
    if ($Path -match "^\\\\")                { return "\\?\UNC" + $Path.Substring(1) }
    if ($Path -match "^[a-zA-Z]:")           { return "\\?\$Path" }
    return $Path
}

function Get-UserDesktopPath {
    try {
        $reg = Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'Desktop' -ErrorAction Stop
        $raw = $reg.Desktop
        if (-not [string]::IsNullOrWhiteSpace($raw)) {
            return [Environment]::ExpandEnvironmentVariables($raw)
        }
    } catch {}
    return (Join-Path $env:USERPROFILE 'Desktop')
}

function Get-UserDownloadsPath {
    $guidName = '{374DE290-123F-4565-9164-39C4925E467B}'
    try {
        $reg = Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name $guidName -ErrorAction Stop
        $raw = $reg.$guidName
        if (-not [string]::IsNullOrWhiteSpace($raw)) {
            return [Environment]::ExpandEnvironmentVariables($raw)
        }
    } catch {}
    return (Join-Path $env:USERPROFILE 'Downloads')
}

function Test-PowerPointInstalled {
    foreach ($p in @(
        'Registry::HKEY_CLASSES_ROOT\PowerPoint.Application\CLSID',
        'HKLM:\SOFTWARE\Classes\PowerPoint.Application\CLSID',
        'HKLM:\SOFTWARE\WOW6432Node\Classes\PowerPoint.Application\CLSID'
    )) {
        try {
            $null = Get-ItemProperty -Path $p -ErrorAction Stop
            return $true
        } catch {}
    }
    return $false
}

# ==================================================================
# HELPER: I/O-ROBUSTHEIT
# ==================================================================
function Invoke-WithRetry {
    param(
        [Parameter(Mandatory=$true)][scriptblock]$Action,
        [int]$MaxRetries     = 10,
        [int]$InitialDelayMs = 200,
        [int]$MaxDelayMs     = 8000,
        [string]$Context     = ""
    )
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            $result = & $Action
            if ($attempt -gt 1) {
                Write-DetailedLog "Retry-Erfolg bei Versuch $attempt/$MaxRetries [$Context]" "DEBUG"
            }
            return $result
        } catch [System.IO.IOException], [System.UnauthorizedAccessException] {
            if ($attempt -eq $MaxRetries) {
                Write-DetailedLog "Retry aufgegeben nach $attempt/$MaxRetries Versuchen [$Context]: $($_.Exception.Message)" "WARN"
                throw
            }
            $delay = [Math]::Min([int]($InitialDelayMs * [Math]::Pow(2, $attempt - 1)), $MaxDelayMs)
            Write-DetailedLog "Retry $attempt/$MaxRetries nach ${delay}ms [$Context]: $($_.Exception.Message)" "DEBUG"
            Start-Sleep -Milliseconds $delay
        }
    }
}

function Invoke-FilePrimingRead {
    param([string]$Path)
    $fs = $null
    try {
        $fs = [System.IO.File]::Open($Path,
              [System.IO.FileMode]::Open,
              [System.IO.FileAccess]::Read,
              [System.IO.FileShare]::ReadWrite)
        $buf = New-Object byte[] 4096
        [void]$fs.Read($buf, 0, $buf.Length)
    } catch {} finally {
        if ($fs) { try { $fs.Close() } catch {} }
    }
}

function Wait-FileStable {
    param(
        [string]$Path,
        [int]$MaxAttempts = $script:AvStableMaxAttempts,
        [int]$DelayMs     = $script:AvStableDelayMs
    )
    $lastSize    = [long]-1
    $stableHits  = 0
    for ($i = 1; $i -le $MaxAttempts; $i++) {
        $size      = [long]-1
        $exclusive = $false
        try {
            $fi   = New-Object System.IO.FileInfo($Path)
            $size = $fi.Length
        } catch {
            Start-Sleep -Milliseconds $DelayMs
            continue
        }
        $fs = $null
        try {
            $fs = [System.IO.File]::Open($Path,
                  [System.IO.FileMode]::Open,
                  [System.IO.FileAccess]::Read,
                  [System.IO.FileShare]::None)
            $exclusive = $true
        } catch {} finally {
            if ($fs) { try { $fs.Close() } catch {} }
        }
        if ($exclusive -and $size -ge 0 -and $size -eq $lastSize) {
            $stableHits++
            if ($stableHits -ge 1) { return $true }
        } else {
            $stableHits = 0
            $lastSize   = $size
        }
        Start-Sleep -Milliseconds $DelayMs
    }
    return $false
}

function Test-IsValidZip {
    param([string]$FilePath)
    try {
        Invoke-WithRetry -Action {
            $z = [System.IO.Compression.ZipFile]::OpenRead($FilePath)
            $z.Dispose()
        }
        return "Valid"
    } catch [System.IO.IOException] {
        Write-DetailedLog "ZIP-Check: gesperrt nach Retry: $FilePath" "WARN"
        return "Locked"
    } catch [System.UnauthorizedAccessException] {
        Write-DetailedLog "ZIP-Check: Zugriff verweigert: $FilePath" "WARN"
        return "Locked"
    } catch {
        return "Invalid"
    }
}

function Test-IsOleEncrypted {
    param([string]$FilePath)
    $bufferSize = 8192
    $fs = $null
    try {
        $fs = Invoke-WithRetry -Action { [System.IO.File]::OpenRead($FilePath) }
    } catch {
        Write-DetailedLog "OLE-Header-Check fehlgeschlagen (Lock): $FilePath - $_" "WARN"
        return $false
    }
    $bytes = $null
    try {
        if ($fs.Length -lt 8) { return $false }
        $readLen = [int][Math]::Min([long]$bufferSize, [long]$fs.Length)
        $bytes   = New-Object byte[] $readLen
        $fs.Read($bytes, 0, $readLen) | Out-Null
    } finally { $fs.Close() }

    if ($bytes[0] -ne 0xD0 -or $bytes[1] -ne 0xCF -or
        $bytes[2] -ne 0x11 -or $bytes[3] -ne 0xE0) {
        return $false
    }

    $content = [System.Text.Encoding]::ASCII.GetString($bytes) -replace '\x00',''
    return ($content -match "EncryptedPackage" -or $content -match "EncryptionInfo")
}

function Test-FileIsLocked {
    param([string]$FilePath)
    $lp = Add-LongPathPrefix $FilePath
    $stream = $null
    try {
        if (-not [System.IO.File]::Exists($lp)) { return $false }
        # Schreibgeschuetzte Dateien nur lesend oeffnen. ReadWrite wirft bei
        # gesetztem ReadOnly-Attribut UnauthorizedAccessException (gemessen
        # unter 5.1) - der Probelauf meldete jede solche Datei als "Zugriff
        # verweigert", weil das Attribut dort nicht entfernt wird. Fuer die
        # Sperrpruefung genuegt FileShare.None: jeder fremde Handle laesst
        # auch den lesenden Open scheitern.
        $zugriff = [System.IO.FileAccess]::ReadWrite
        try {
            if ([System.IO.File]::GetAttributes($lp) -band [System.IO.FileAttributes]::ReadOnly) {
                $zugriff = [System.IO.FileAccess]::Read
            }
        } catch {}
        $stream = [System.IO.File]::Open($lp,
                       [System.IO.FileMode]::Open,
                       $zugriff,
                       [System.IO.FileShare]::None)
        return $false
    } catch [System.UnauthorizedAccessException] {
        return "AccessDenied"
    } catch {
        return $true
    } finally {
        if ($stream) { try { $stream.Close() } catch {} }
    }
}

function Clear-ReadOnlyAttribut {
    # Entfernt nur das ReadOnly-Bit. $true = war gesetzt und ist entfernt.
    param([string]$LongPath)
    try {
        $a = [System.IO.File]::GetAttributes($LongPath)
        if (-not ($a -band [System.IO.FileAttributes]::ReadOnly)) { return $false }
        [System.IO.File]::SetAttributes($LongPath, $a -band (-bnot [System.IO.FileAttributes]::ReadOnly))
        return $true
    } catch { return $false }
}

function Set-ReadOnlyAttribut {
    # Setzt das ReadOnly-Bit wieder (Rollback: Original unveraendert zurueck).
    param([string]$LongPath)
    try {
        $a = [System.IO.File]::GetAttributes($LongPath)
        [System.IO.File]::SetAttributes($LongPath, $a -bor [System.IO.FileAttributes]::ReadOnly)
    } catch {
        Write-DetailedLog "ReadOnly-Attribut nicht wiederherstellbar: $LongPath - $_" "WARN"
    }
}

function Reserve-UniqueDestination {
    param(
        [string]$BasePath,
        [string]$Extension,
        [int]$MaxRetries = $script:MaxNameClashRetries
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
        if (-not ('PptUnprotect.PrivilegeHelper' -as [type])) {
            Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace PptUnprotect {
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
        $okRestore = [PptUnprotect.PrivilegeHelper]::Enable('SeRestorePrivilege')
        $okOwner   = [PptUnprotect.PrivilegeHelper]::Enable('SeTakeOwnershipPrivilege')
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
# HELPER: PROZESSE
# ==================================================================
function Test-IsOwnPptProcess {
    <#
        Prueft, ob eine PID wirklich zu einer von UNS gestarteten
        PowerPoint-Instanz gehoert. Windows vergibt PIDs wieder; ohne den
        Startzeit-Vergleich koennte hier eine zwischenzeitlich vom Nutzer
        geoeffnete Praesentation abgeschossen werden.
    #>
    param([int]$ProcessId)
    try {
        $proc = Get-Process -Id $ProcessId -ErrorAction Stop
        if ($proc.ProcessName -ne 'POWERPNT') { return $false }
        try {
            if ($proc.StartTime -lt $script:ScriptStartTime) { return $false }
        } catch {
            return $false
        }
        return $true
    } catch {
        return $false
    }
}

function Clear-TrackedPowerPointInstances {
    foreach ($id in @($script:TrackedPptPids | Select-Object -Unique)) {
        if (Test-IsOwnPptProcess -ProcessId $id) {
            try { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue -WhatIf:$false } catch {}
        }
    }
    [System.GC]::Collect()
    $script:TrackedPptPids.Clear()
}

function Wait-TrackedPowerPointExit {
    <#
        Vor jedem COM-Job: eigene, bereits per Quit() beendete Instanzen
        laufen oft noch einen Moment weiter. Der naechste Job erfasst vor
        New-Object alle vorhandenen POWERPNT-Prozesse als fremd - ein noch
        auslaufender eigener Prozess wuerde dort als Sitzung des Anwenders
        gelten und den Lauf grundlos abbrechen. Daher bis zu $TimeoutSec
        warten und Reste danach beenden - nur eigene (Name + Startzeit).
    #>
    param([int]$TimeoutSec = 10)
    if ($script:TrackedPptPids.Count -eq 0) { return }
    $bis = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $bis) {
        $lebend = @($script:TrackedPptPids | Select-Object -Unique |
                    Where-Object { Test-IsOwnPptProcess -ProcessId $_ })
        if ($lebend.Count -eq 0) { break }
        Start-Sleep -Milliseconds 250
    }
    Clear-TrackedPowerPointInstances
}

# ==================================================================
# EIGENE POWERPOINT-INSTANZ ERKENNEN (Code fuer die COM-Jobs)
# ==================================================================
# PowerPoint ist ein Einzelinstanz-COM-Server: laeuft bereits eine Sitzung,
# liefert 'New-Object -ComObject PowerPoint.Application' GENAU DIESE, und ihr
# HWND fuehrt auf ihre PID (vom Auftraggeber gemessen, 29.09.2026). Die
# frueheren Kommentare nahmen an, die PID bleibe dann leer - tatsaechlich
# galt die Sitzung des Anwenders als eigene Instanz.
# Deshalb erfasst jeder Job VOR New-Object die POWERPNT-Prozesse der eigenen
# Sitzung. Gehoert das gelieferte Objekt zu einem davon (oder laesst es sich
# nicht zuordnen), ist es NICHT die eigene Instanz: keine Eigenschaft wird
# verstellt, kein Quit, keine PID-Datei, kein Kill - der Job meldet 'FREMD',
# das Hauptskript bricht den Lauf ab.
# Die Jobs laufen in eigenen Prozessen und sehen die Skriptfunktionen nicht;
# der Code geht deshalb als Text an beide Jobs (Konvertierung, Smoke-Test) -
# eine Quelle statt zwei Kopien.
$script:PptInstanzCode = @'
function Get-PptSitzungsPids {
    $sid = -1
    try { $sid = (Get-Process -Id $PID -ErrorAction Stop).SessionId } catch {}
    return ,@(Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue |
              Where-Object { $sid -lt 0 -or $_.SessionId -eq $sid } |
              ForEach-Object { [int]$_.Id })
}

function Get-PptInstanz {
    # Eigen = $true nur, wenn die Instanz nachweislich NEU gestartet wurde.
    param($App, [int[]]$VorherPids = @())
    $hwndPid = 0
    try {
        if (-not ('Win32.User32PptInst' -as [type])) {
            Add-Type -Name 'User32PptInst' -Namespace 'Win32' -ErrorAction Stop -MemberDefinition '[DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);'
        }
        $procId = 0
        [void][Win32.User32PptInst]::GetWindowThreadProcessId([IntPtr]$App.HWND, [ref]$procId)
        if ($procId -gt 0) { $hwndPid = $procId }
    } catch {}
    if ($hwndPid -gt 0) {
        return @{ Eigen = ($VorherPids -notcontains $hwndPid); Pid = $hwndPid }
    }
    # HWND nicht lesbar: nur eindeutig, wenn genau EIN Prozess neu hinzukam.
    $nachher = Get-PptSitzungsPids
    $neu = @($nachher | Where-Object { $VorherPids -notcontains $_ })
    if ($neu.Count -eq 1) { return @{ Eigen = $true; Pid = [int]$neu[0] } }
    return @{ Eigen = $false; Pid = $null }
}

function Get-PptFremdePraesentationen {
    # Offene Praesentationen AUSSERHALB des eigenen Arbeitsordners; -1 = nicht lesbar.
    param($App, [string]$EigenerOrdner)
    try {
        $n = 0
        $liste = $App.Presentations
        if ($null -eq $liste) { return -1 }
        foreach ($p in $liste) {
            $fn = ''
            try { $fn = [string]$p.FullName } catch {}
            if (-not $fn.StartsWith($EigenerOrdner, [System.StringComparison]::OrdinalIgnoreCase)) { $n++ }
        }
        return $n
    } catch { return -1 }
}
'@

# ==================================================================
# HELPER: DATEISYSTEM-TRAVERSIERUNG
# ==================================================================
function Get-PptFilesRobust {
    param([string]$RootPath, [string]$ExtPattern)
    $queue   = [System.Collections.Queue]::new()
    $visited = [System.Collections.Generic.HashSet[string]]::new(
                   [System.StringComparer]::OrdinalIgnoreCase)
    $queue.Enqueue(@($RootPath, 0))
    while ($queue.Count -gt 0) {
        $item  = $queue.Dequeue()
        $dir   = $item[0]
        $depth = $item[1]
        if ($depth -gt $script:MaxDirectoryDepth) {
            Write-DetailedLog "Tiefenlimit ($($script:MaxDirectoryDepth)) erreicht, ueberspringe: $dir" "WARN"
            continue
        }
        if (-not $visited.Add($dir)) { continue }
        try {
            foreach ($f in [System.IO.Directory]::EnumerateFiles($dir, "*.*")) {
                if ([System.IO.Path]::GetExtension($f) -match $ExtPattern) {
                    $f
                }
            }
        } catch {}
        try {
            foreach ($sub in [System.IO.Directory]::EnumerateDirectories($dir)) {
                $isReparse = $false
                try {
                    $attr = [System.IO.File]::GetAttributes($sub)
                    if ($attr -band [System.IO.FileAttributes]::ReparsePoint) { $isReparse = $true }
                } catch { continue }
                if ($isReparse) {
                    Write-DetailedLog "Reparse-Point/Junction uebersprungen: $sub" "DEBUG"
                    continue
                }
                $leaf = [System.IO.Path]::GetFileName($sub.TrimEnd('\'))
                if ($script:ExcludeDirNames -contains $leaf) {
                    Write-DetailedLog "Verzeichnis uebersprungen (Ausschlussliste): $sub" "DEBUG"
                    continue
                }
                $queue.Enqueue(@($sub, $depth + 1))
            }
        } catch {}
    }
}

# Eigene Sicherungen heissen '<Original>.bak_<8 Hex, klein, mindestens ein
# Buchstabe>' (New-BackupPfad). Nur GENAU dieses Muster raeumt das Skript
# auf. Vorher hiess die Sicherung schlicht '<Original>.bak', und das
# Aufraeummuster '^(.+?)\.bak(_[0-9a-fA-F]+)?$' traf damit auch eine vom
# Anwender selbst angelegte 'Vortrag.pptx.bak' (oder 'Vortrag.pptx.bak_2024')
# - geloescht, sobald das Original gleich gross oder groesser war. Eine
# vorhandene 'Vortrag.pptx.bak' wurde zudem vor dem Schreiben nach
# '.bak_<6 Hex>' umbenannt und fiel damit im naechsten Lauf ebenfalls unter
# das Muster. Der Buchstabe ist Pflicht, damit eine Datumsendung wie
# '.bak_20240101' (acht Ziffern) nie als eigene Sicherung gilt.
# Alte '.bak'- und '.bak_<6 Hex>'-Reste frueherer Fassungen werden nicht mehr
# angefasst - lieber ein liegengebliebener Rest als eine fremde Sicherung weg.
$script:EigenesBackupMuster = '^(?<orig>.+)\.bak_(?=[0-9a-f]*[a-f])[0-9a-f]{8}$'

function New-BackupPfad {
    param([string]$Original)
    for ($i = 0; $i -lt 100; $i++) {
        $kennung = [Guid]::NewGuid().ToString('N').Substring(0, 8)
        if ($kennung -notmatch '[a-f]') { continue }
        $kandidat = "$Original.bak_$kennung"
        if (-not [System.IO.File]::Exists($kandidat)) { return $kandidat }
    }
    throw "Kein freier Backup-Name fuer $Original"
}

function Get-OriginalFromBackupPath {
    # Vortrag.pptx.bak_1a2b3c4d  ->  Vortrag.pptx ; alles andere -> $null
    param([string]$BackupPath)
    if ($BackupPath -cmatch $script:EigenesBackupMuster) {
        return $Matches['orig']
    }
    return $null
}

function Remove-OrphanedBackups {
    <#
        Raeumt '.bak_xxxxxxxx'-Reste frueherer Laeufe auf (nur das eigene
        Muster, siehe $script:EigenesBackupMuster).

        Geloescht wird nur, wenn ALLE Bedingungen zutreffen:
          1. Der abgeleitete Originalname endet auf eine PowerPoint-Endung
             (fremde .bak-Dateien bleiben unangetastet),
          2. das Original existiert und ist mindestens so gross wie das Backup,
          3. das Backup ist aelter als $script:BackupCleanupMinAgeHours.

        Fehlt das Original oder ist es leer, kann das Backup die einzige
        intakte Kopie sein - solche Funde werden gemeldet, nie geloescht.
    #>
    param([string]$RootPath)

    $res = @{
        Deleted = 0; Kept = 0; Orphans = 0; Failed = 0
        OrphanList = [System.Collections.Generic.List[string]]::new()
    }
    $cutoff = (Get-Date).AddHours(-$script:BackupCleanupMinAgeHours)

    foreach ($bak in (Get-PptFilesRobust $RootPath '^\.bak_[0-9a-f]{8}$')) {
        try {
            $orig = Get-OriginalFromBackupPath -BackupPath $bak
            if ([string]::IsNullOrWhiteSpace($orig)) { continue }
            if ($orig -notmatch '\.(ppt|pot|pps|pptx|pptm|potx|potm|ppsx|ppsm)$') {
                Write-DetailedLog "Fremde .bak-Datei ignoriert: $bak" "DEBUG"
                continue
            }

            $fi = [System.IO.FileInfo]::new($bak)
            if (-not $fi.Exists) { continue }

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
                $res.OrphanList.Add($bak)
                Write-Log "Backup ohne intaktes Original - NICHT geloescht, bitte pruefen: $bak" "WARN"
                continue
            }

            # Das Original muss mindestens so gross sein wie das Backup.
            # 'Length -gt 0' allein genuegt nicht: [System.IO.File]::Copy
            # trunkiert die Zieldatei beim Start und schreibt fortlaufend -
            # ein harter Abbruch hinterlaesst deshalb typischerweise ein
            # TEILWEISE geschriebenes Original mit Groesse > 0. Genau dieser
            # Zustand bestand die Pruefung, und das Backup mit der einzigen
            # vollstaendigen Fassung wurde geloescht. Gleiche Ursache und
            # gleiche Korrektur wie in 2a und 2b.
            if ($origLen -lt $fi.Length) {
                $res.Orphans++
                $res.OrphanList.Add($bak)
                Write-Log ("Backup groesser als das Original ({0:N0} statt {1:N0} Bytes) - Original " +
                           "moeglicherweise abgeschnitten. NICHT geloescht, bitte pruefen: {2}" -f `
                           $fi.Length, $origLen, $bak) "WARN"
                continue
            }

            # Altersfrist an der CreationTime messen, NICHT an der
            # LastWriteTime: das Backup entsteht per [System.IO.File]::Copy,
            # und Copy uebertraegt die LastWriteTime der Quelle auf das Ziel.
            # Auf Ablagen sind Dokumente typischerweise Jahre alt, das Backup
            # erbt das - ein Sekunden altes .bak galt damit sofort als
            # verwaist und der Kept-Zweig war praktisch toter Code.
            if ($fi.CreationTime -gt $cutoff) { $res.Kept++; continue }

            if (Confirm-Write $bak 'Verwaistes Backup loeschen') {
                try {
                    [System.IO.File]::Delete((Add-LongPathPrefix $bak))
                    $res.Deleted++
                    Write-DetailedLog "Verwaistes Backup entfernt: $bak" "DEBUG"
                } catch {
                    $res.Failed++
                    Write-DetailedLog "Backup nicht loeschbar: $bak - $_" "WARN"
                }
            } else {
                $res.Kept++
            }
        } catch {
            $res.Failed++
        }
    }
    return $res
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
    $elapsed = $now - $script:ScriptStartTime

    if ($script:TotalFiles -gt 0) {
        $pct = [Math]::Max(0, [Math]::Min(100,
                [int](($script:ProcessedCount / $script:TotalFiles) * 100)))
        if ($script:ProcessedCount -gt 0) {
            $secPerFile = $elapsed.TotalSeconds / $script:ProcessedCount
            $remaining  = ($script:TotalFiles - $script:ProcessedCount) * $secPerFile
            $etaSpan    = [TimeSpan]::FromSeconds([Math]::Max(0, $remaining))
            $etaStr     = "ETA: " + $etaSpan.ToString('hh\:mm\:ss')
        }
    }

    $totalStr = if ($script:SkipPreScan) { "?" } else { "$($script:TotalFiles)" }
    $activity = "PPT Deep Unprotect  |  $etaStr  |  $($script:ProcessedCount) / $totalStr"
    Write-Progress -Activity $activity -Status $Status -CurrentOperation $CurrentFile -PercentComplete $pct
}

# ==================================================================
# KONVERTIERUNG (alt .ppt/.pps/.pot -> .pptx/.ppsx/.potx)
# ==================================================================
function Convert-PptToPptx {
    param([string]$SourcePath, [string]$DestPathBase, [string]$OriginalExt)

    # Auslaufende eigene Instanzen des vorigen Jobs abwarten, sonst gelten
    # sie im naechsten Job als fremd (siehe Wait-TrackedPowerPointExit).
    Wait-TrackedPowerPointExit

    $pidFile = Join-Path $script:TempPath "PptDeepClean_$([Guid]::NewGuid().ToString()).pid"
    $job = Start-Job -ScriptBlock {
        param($src, $destBase, $origExt, $consts, $pidFile, $instanzCode, $eigenerOrdner)

        . ([scriptblock]::Create($instanzCode))

        # Vor New-Object erfassen, welche PowerPoint-Prozesse schon laufen
        # (Begruendung bei $script:PptInstanzCode).
        $vorher = Get-PptSitzungsPids
        $ppt    = New-Object -ComObject PowerPoint.Application
        $inst   = Get-PptInstanz -App $ppt -VorherPids $vorher
        if (-not $inst.Eigen) {
            # Fremde oder nicht zuordenbare Instanz: NICHTS daran verstellen,
            # kein Quit, keine PID-Datei - nur die eigene Referenz freigeben.
            try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ppt) | Out-Null } catch {}
            $wer = if ($inst.Pid) { "PID $($inst.Pid)" } else { "PID nicht zuzuordnen" }
            return @{ Status = "FREMD"; PptPid = $null
                      Msg = "PowerPoint des Anwenders laeuft ($wer) - Lauf abgebrochen" }
        }
        $myPptPid = $inst.Pid
        try { [System.IO.File]::WriteAllText($pidFile, "$myPptPid") } catch {}

        $ppt.Visible            = $consts.msoTrue
        $ppt.DisplayAlerts      = $consts.ppAlertsNone
        $ppt.FeatureInstall     = $consts.msoFeatureInstallNone
        try { $ppt.AutomationSecurity = $consts.msoAutomationSecurityForceDisable } catch {}
        try { $ppt.WindowState = $consts.ppWindowMinimized } catch {}

        $pres = $null
        $res  = @{ Status = "ERROR"; Msg = "Oeffnen fehlgeschlagen" }
        try {
            try {
                $pres = $ppt.Presentations.Open($src, -1, -1, 0)
            } catch {
                $res.Msg = $_.Exception.Message
            }
            if (-not $pres) { throw [System.InvalidOperationException]::new($res.Msg) }

            $hasMacros = $false
            try { $hasMacros = [bool]$pres.HasVBProject } catch { $hasMacros = $true }
            $origExt = $origExt.ToLower()

            if ($origExt -match "^\.pps") {
                $format = if ($hasMacros) { $consts.ppSaveAsOpenXMLShowMacroEnabled         } else { $consts.ppSaveAsOpenXMLShow         }
                $ext    = if ($hasMacros) { ".ppsm" } else { ".ppsx" }
            } elseif ($origExt -match "^\.pot") {
                $format = if ($hasMacros) { $consts.ppSaveAsOpenXMLTemplateMacroEnabled     } else { $consts.ppSaveAsOpenXMLTemplate     }
                $ext    = if ($hasMacros) { ".potm" } else { ".potx" }
            } else {
                $format = if ($hasMacros) { $consts.ppSaveAsOpenXMLPresentationMacroEnabled } else { $consts.ppSaveAsOpenXMLPresentation }
                $ext    = if ($hasMacros) { ".pptm" } else { ".pptx" }
            }

            $final = "$destBase$ext"
            # Beim SaveAs etwaige Open-/Modify-Passwoerter mitabraeumen, damit
            # sie nicht in die Zieldatei uebernommen werden. PowerPoint COM
            # akzeptiert beim Open keinen Passwort-Parameter; das ReadOnly-Open
            # umgeht ein gesetztes WritePassword und der Reset auf "" entfernt
            # es endgueltig beim SaveAs.
            try { $pres.Password      = "" } catch {}
            try { $pres.WritePassword = "" } catch {}
            $pres.SaveAs($final, $format)
            # Saved=true verhindert unsichtbare "Speichern?"-Dialoge in Close()
            # bei alten OLE-Embeds oder gebrochenen Verknuepfungen.
            try { $pres.Saved = $consts.msoTrue } catch {}
            $pres.Close()

            $res = @{ Status = "OK"; Path = $final }
        } catch {
            try {
                if ($pres) {
                    try { $pres.Saved = $consts.msoTrue } catch {}
                    $pres.Close()
                }
            } catch {}
            $res = @{ Status = "ERROR"; Msg = $_.Exception.Message }
        } finally {
            if ($pres) {
                try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($pres) | Out-Null } catch {}
                $pres = $null
            }
        }

        # Quit nur, wenn ausser den eigenen Arbeitsdateien NICHTS offen ist.
        # PowerPoint startet nur einmal: oeffnet der Anwender waehrend des
        # Laufs eine Praesentation, landet sie in DIESER Instanz - ein Quit()
        # schloesse sie. Dann Instanz stehen lassen, Fenster wiederherstellen
        # (ppWindowNormal=1, ppAlertsAll=2, Werte aus der Interop-Assembly
        # gelesen), PID NICHT zur Nachverfolgung melden (sonst beendet sie
        # Clear-TrackedPowerPointInstances) und warnen.
        $fremdOffen = Get-PptFremdePraesentationen -App $ppt -EigenerOrdner $eigenerOrdner
        if ($fremdOffen -eq 0) {
            try { $ppt.Quit() } catch {}
            $res.PptPid = $myPptPid
        } else {
            try { $ppt.WindowState   = 1 } catch {}
            try { $ppt.DisplayAlerts = 2 } catch {}
            $res.PptPid  = $null
            $res.Warnung = if ($fremdOffen -lt 0) {
                "Offene Praesentationen der PowerPoint-Instanz (PID $myPptPid) nicht lesbar - PowerPoint NICHT beendet"
            } else {
                "In der PowerPoint-Instanz (PID $myPptPid) hat der Anwender $fremdOffen Praesentation(en) geoeffnet - PowerPoint NICHT beendet"
            }
        }
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ppt) | Out-Null } catch {}
        return $res
    } -ArgumentList $SourcePath, $DestPathBase, $OriginalExt, $script:PptConstants, $pidFile, $script:PptInstanzCode, ($script:TempPath.TrimEnd('\') + '\')

    if (-not (Wait-Job $job -Timeout $script:FileOpenTimeoutSeconds)) {
        Stop-Job  $job -WhatIf:$false
        Remove-Job $job -WhatIf:$false
        if ([System.IO.File]::Exists($pidFile)) {
            try {
                $pidContent = [System.IO.File]::ReadAllText($pidFile).Trim()
                if ($pidContent -match '^\d+$') {
                    $savedPid = [int]$pidContent
                    # Nur beenden, was nachweislich die eigene Instanz ist
                    # (Name + Startzeit nach Skriptbeginn). Vorher genuegte der
                    # Name - bei einer Sitzung des Anwenders, deren PID in der
                    # Datei stand, beendete dieser Zweig sie per -Force.
                    if ($savedPid -gt 0 -and (Test-IsOwnPptProcess -ProcessId $savedPid)) {
                        try { Stop-Process -Id $savedPid -Force -ErrorAction SilentlyContinue -WhatIf:$false } catch {}
                        Write-DetailedLog "Timeout: POWERPNT.EXE PID $savedPid sofort beendet" "WARN"
                    }
                }
            } catch {}
            [System.IO.File]::Delete($pidFile)
        }
        throw [System.TimeoutException]::new("Datei reagiert nicht (evtl. Oeffnungspasswort oder haengender Dialog).")
    }

    $result = Receive-Job $job
    Remove-Job $job -WhatIf:$false
    if ([System.IO.File]::Exists($pidFile)) { try { [System.IO.File]::Delete($pidFile) } catch {} }
    if ($result.PptPid) { $script:TrackedPptPids.Add([int]$result.PptPid) }
    if ($result.Warnung) {
        # Der Anwender arbeitet in PowerPoint - weiter per COM zu arbeiten,
        # hiesse, seine Sitzung zu uebernehmen. Die laufende Datei wird noch
        # fertiggestellt, dann bricht die Hauptschleife ab.
        $script:FremdePptAbbruch = $true
        Write-Host ""
        Write-Host "WARNUNG: $($result.Warnung). Lauf wird nach dieser Datei beendet." -ForegroundColor Yellow
        Write-Log "$($result.Warnung) - Lauf wird nach dieser Datei beendet." "WARN"
    }

    if ($result.Status -eq "OK") { return $result.Path }
    if ($result.Status -eq "FREMD") {
        $script:FremdePptAbbruch = $true
        Write-Log "$($result.Msg)" "ERROR"
    }
    throw $result.Msg
}

# ==================================================================
# SCHUTZ ENTFERNEN (XML-EBENE)
#   [A] ppt/presentation.xml
#         - modifyVerifier         (Aenderungs-Passwort)
#         - writeProtection        (Immer schreibgeschuetzt oeffnen)
#         - readOnlyRecommended    (Attribut auf p:presentation)
#   [B] ppt/presProps.xml          (Fallback: readOnlyRecommended)
#   [C-H] Folien/Master/Layouts/Notizen/Handouts:
#         cNvSpPr / cNvPicPr / cNvCxnSpPr / cNvGraphicFramePr / cNvGrpSpPr locked="1"
#         spLocks / picLocks / cxnSpLocks / graphicFrameLocks / grpSpLocks
# ==================================================================
function Remove-PptxProtection {
    param([string]$FilePath)

    $actionsTaken = [System.Collections.ArrayList]::new()
    $zip          = $null

    # Nur Sperren, die ein Anwender setzt. 'noGrp' und 'noChangeAspect'
    # schreibt PowerPoint von sich aus: noGrp an jeden Platzhalter und an
    # Tabellenrahmen (gemessen: 72x in einer frisch erzeugten Datei mit
    # Standardvorlage), noChangeAspect an eingefuegte Bilder. Standen sie in
    # der Liste, schrieb das Skript JEDE .pptx um, auch voellig ungeschuetzte
    # (Befundbericht Abschnitt 9 a: 14 XML-Teile geaendert, "Entsperrt: 6"
    # bei einer einzigen wirklich geschuetzten Datei) - mit neuem Zeitstempel
    # und vollstaendigem Neu-Upload auf Google Drive. Sie verhindern zudem
    # nur Gruppieren bzw. freies Verzerren und sind kein Bearbeitungsschutz.
    #
    # Nachgezaehlt am 29.09.2026 in den Office-2024-Vorlagen dieses Rechners
    # (root\Templates: 6 .potx, root\Document Themes 16: 11 .thmx):
    #  - 'noChangeArrowheads' steht in 3 der 6 .potx (Pitchbook, Training,
    #    WidescreenPresentation) - 25x in Notizseiten, 22x auf Folien (auch
    #    picLocks an Bildern). Wie noChangeAspect nur eine Formatsperre
    #    (Pfeilspitzen), kein Bearbeitungsschutz -> ganz aus der Liste.
    #  - 'noRot' steht in ALLEN 6 .potx, 60x - ausnahmslos am Folienbild-
    #    Platzhalter (ph type="sldImg") in notesMaster/notesSlide:
    #    <a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/>. Jede .pptx mit
    #    Notizen haette damit als geschuetzt gegolten und waere neu
    #    geschrieben worden. Auf Folien, Layouts und Mastern kam noRot kein
    #    einziges Mal vor - dort bleibt es eine Anwendersperre.
    #  - 'noTextEdit' steht 5x am selben sldImg-Platzhalter (Training.potx);
    #    der Platzhalter traegt keinen Text.
    #    -> noRot und noTextEdit werden nur am sldImg-Platzhalter der
    #       Notizteile uebergangen ($sldImgIgnorieren).
    #  - 'noEditPoints' steht 10x im Design "Ion Boardroom" (.thmx) an
    #    Schmuckformen in Layouts und Master, nie auf Folien.
    #    -> in slideLayout/slideMaster uebergangen, auf Folien entfernt.
    # Probe: die 6 .potx als .pptx durch Remove-PptxProtection - vorher bei
    # allen 6 "Shape-Lock" (7 bis 25 Teile je Datei), nachher bei keiner.
    $lockAttrs = @("noSelect","noMove","noResize","noRot",
                   "noTextEdit",
                   "noAdjustHandles","noEditPoints","noUngrp")
    $sldImgIgnorieren = @("noRot","noTextEdit")

    try {
        $zip = Invoke-WithRetry -Context "ZipFile.Open Update" -Action {
            [System.IO.Compression.ZipFile]::Open($FilePath, "Update")
        }

        $allEntryNames  = @($zip.Entries | Select-Object -ExpandProperty FullName)
        $staticTargets  = @("ppt/presentation.xml", "ppt/presProps.xml")
        $dynamicTargets = $allEntryNames | Where-Object {
            $_ -match "^ppt/(slides/slide|slideMasters/slideMaster|slideLayouts/slideLayout|notesSlides/notesSlide|handoutMasters/handoutMaster|notesMasters/notesMaster)\d+\.xml$"
        }
        $allTargets = $staticTargets + @($dynamicTargets)

        foreach ($targetName in $allTargets) {
            $entry = $zip.GetEntry($targetName)
            if (-not $entry) { continue }

            $stream  = $entry.Open()
            $reader  = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::UTF8, $true)
            $content = $reader.ReadToEnd()
            $reader.Close()
            $stream.Close()

            if ([string]::IsNullOrWhiteSpace($content)) { continue }

            $xmlDoc = New-Object System.Xml.XmlDocument
            $xmlDoc.PreserveWhitespace = $true
            try { $xmlDoc.LoadXml($content) }
            catch {
                Write-DetailedLog "XML-Parsefehler in ${targetName}: $_" "WARN"
                continue
            }

            $modified = $false

            # --------------------------------------------------------
            # [A] ppt/presentation.xml
            # --------------------------------------------------------
            if ($targetName -eq "ppt/presentation.xml") {

                $nodes = @($xmlDoc.SelectNodes("//*[local-name()='modifyVerifier']"))
                foreach ($node in $nodes) {
                    $node.ParentNode.RemoveChild($node) | Out-Null
                    $modified = $true
                    $actionsTaken.Add("Aenderungs-Passwort") | Out-Null
                    Write-DetailedLog "Entfernt: modifyVerifier" "DEBUG"
                }

                $wp = @($xmlDoc.SelectNodes("//*[local-name()='writeProtection']"))
                foreach ($node in $wp) {
                    $node.ParentNode.RemoveChild($node) | Out-Null
                    $modified = $true
                    $actionsTaken.Add("Schreibschutz (writeProtection)") | Out-Null
                    Write-DetailedLog "Entfernt: writeProtection" "DEBUG"
                }

                $presRoot = $xmlDoc.SelectSingleNode("//*[local-name()='presentation']")
                if ($presRoot -and $presRoot.Attributes) {
                    $toRemove = @()
                    foreach ($a in $presRoot.Attributes) {
                        if ($a.LocalName -eq "readOnlyRecommended") { $toRemove += $a }
                    }
                    foreach ($a in $toRemove) {
                        $presRoot.Attributes.Remove($a) | Out-Null
                        $modified = $true
                        $actionsTaken.Add("Schreibschutz-Empfehlung") | Out-Null
                        Write-DetailedLog "Entfernt: readOnlyRecommended (Attribut auf p:presentation)" "DEBUG"
                    }
                }
            }

            # --------------------------------------------------------
            # [B] ppt/presProps.xml (Fallback)
            # --------------------------------------------------------
            if ($targetName -eq "ppt/presProps.xml") {

                $nodes = @($xmlDoc.SelectNodes("//*[local-name()='readOnlyRecommended']"))
                foreach ($node in $nodes) {
                    $node.ParentNode.RemoveChild($node) | Out-Null
                    $modified = $true
                    $actionsTaken.Add("Schreibschutz-Empfehlung") | Out-Null
                    Write-DetailedLog "Entfernt: readOnlyRecommended (Knoten in presProps)" "DEBUG"
                }

                $presNode = $xmlDoc.SelectSingleNode("//*[local-name()='presentationPr']")
                if ($presNode -and $presNode.Attributes) {
                    $toRemove = @()
                    foreach ($a in $presNode.Attributes) {
                        if ($a.LocalName -eq "readOnlyRecommended") { $toRemove += $a }
                    }
                    foreach ($a in $toRemove) {
                        $presNode.Attributes.Remove($a) | Out-Null
                        $modified = $true
                        $actionsTaken.Add("Schreibschutz-Empfehlung") | Out-Null
                        Write-DetailedLog "Entfernt: readOnlyRecommended (Attribut in presProps)" "DEBUG"
                    }
                }
            }

            # --------------------------------------------------------
            # [C-H] Folien, Master, Layouts, Notizen, Handouts
            # --------------------------------------------------------
            if ($targetName -ne "ppt/presentation.xml" -and $targetName -ne "ppt/presProps.xml") {

                $lockContainers = @($xmlDoc.SelectNodes(
                    "//*[local-name()='cNvSpPr'          or local-name()='cNvPicPr' or
                         local-name()='cNvCxnSpPr'       or local-name()='cNvGraphicFramePr' or
                         local-name()='cNvGrpSpPr']"))
                foreach ($node in $lockContainers) {
                    $lockedAttr = $node.Attributes["locked"]
                    if ($lockedAttr -and $lockedAttr.Value -eq "1") {
                        $lockedAttr.Value = "0"
                        $modified = $true
                        Write-DetailedLog "locked=0 gesetzt in $targetName ($($node.LocalName))" "DEBUG"
                    }
                }

                $lockNodes = @($xmlDoc.SelectNodes(
                    "//*[local-name()='spLocks'          or local-name()='picLocks' or
                         local-name()='cxnSpLocks'       or local-name()='graphicFrameLocks' or
                         local-name()='grpSpLocks']"))
                $lockTypeCounts = @{}
                $istNotizTeil  = $targetName -match '^ppt/(notesSlides/notesSlide|notesMasters/notesMaster)\d+\.xml$'
                $istLayoutTeil = $targetName -match '^ppt/(slideLayouts/slideLayout|slideMasters/slideMaster)\d+\.xml$'
                foreach ($node in $lockNodes) {
                    $hadLock = $false
                    # Von Office selbst gesetzte Attribute je Fundort uebergehen
                    # (Messung siehe $lockAttrs).
                    $uebergehen = @()
                    if ($istLayoutTeil) { $uebergehen += 'noEditPoints' }
                    if ($istNotizTeil -and $node.LocalName -eq 'spLocks') {
                        # spLocks -> cNvSpPr -> nvSpPr; dort nvPr/ph
                        $ph = $null
                        try {
                            $ph = $node.ParentNode.ParentNode.SelectSingleNode("*[local-name()='nvPr']/*[local-name()='ph']")
                        } catch {}
                        if ($ph -and $ph.GetAttribute('type') -eq 'sldImg') { $uebergehen += $sldImgIgnorieren }
                    }
                    foreach ($attr in $lockAttrs) {
                        if ($uebergehen -contains $attr) { continue }
                        $a = $node.Attributes[$attr]
                        if ($a -and $a.Value -eq "1") {
                            $a.Value  = "0"
                            $hadLock  = $true
                            $modified = $true
                        }
                    }
                    if ($hadLock) {
                        $key = $node.LocalName
                        if ($lockTypeCounts.ContainsKey($key)) { $lockTypeCounts[$key]++ }
                        else                                    { $lockTypeCounts[$key] = 1 }
                    }
                }
                if ($lockTypeCounts.Count -gt 0) {
                    $summary = ($lockTypeCounts.GetEnumerator() |
                        ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ", "
                    Write-DetailedLog "Sperrknoten bereinigt in ${targetName}: $summary" "DEBUG"
                }

                if ($modified) {
                    $shortName = $targetName -replace "^ppt/", ""
                    $actionsTaken.Add("Shape-Lock ($shortName)") | Out-Null
                }
            }

            # --------------------------------------------------------
            # Rueckschreiben (UTF-8 ohne BOM - OOXML-Anforderung)
            # --------------------------------------------------------
            if ($modified) {
                $entry.Delete()
                $newEntry    = $zip.CreateEntry($targetName)
                $writeStream = $newEntry.Open()
                $writer      = $null
                try {
                    $settings = New-Object System.Xml.XmlWriterSettings
                    $settings.Encoding           = [System.Text.UTF8Encoding]::new($false)
                    $settings.Indent             = $false
                    $settings.OmitXmlDeclaration = $false
                    $settings.CloseOutput        = $false
                    $writer = [System.Xml.XmlWriter]::Create($writeStream, $settings)
                    $xmlDoc.Save($writer)
                    $writer.Flush()
                } finally {
                    if ($writer)      { try { $writer.Close()      } catch {} }
                    if ($writeStream) { try { $writeStream.Close() } catch {} }
                }

                Write-DetailedLog "Eintrag neu geschrieben: $targetName" "DEBUG"
            }
        }

        $zip.Dispose()
        $zip = $null

    } catch {
        if ($zip) { try { $zip.Dispose() } catch {} }
        throw "XML-Fehler: $_"
    }

    # Fuehrendes Komma: 'return @()' wuerde beim Entrollen GAR NICHTS
    # liefern und die aufrufende Variable auf $null setzen.
    return ,@($actionsTaken | Select-Object -Unique)
}

# ==================================================================
# TRUST-CENTER SMOKE-TEST
# ==================================================================
# PowerPoint hat keine simple "Datei via XML zusammenbauen"-Loesung wie
# Word/Excel, da der OOXML-Standard fuer .pptx mindestens slideMaster,
# slideLayout und Theme verlangt. Wir nutzen daher COM zur Erstellung der
# Test-Praesentation: PowerPoint erzeugt einmal eine leere .pptx im Temp-
# ordner, schliesst sie, und oeffnet sie dann wieder. Letzterer Schritt ist
# die eigentliche Trust-Center-Pruefung (Open ohne Geschuetzte Ansicht).
# ==================================================================
function Test-PowerPointTrustCenter {
    param(
        [string]$TempDir,
        [int]$TimeoutSec = 25
    )

    $testPath = Join-Path $TempDir ("trustcheck_{0}.pptx" -f ([Guid]::NewGuid().ToString("N")))
    $pidFile  = Join-Path $TempDir ("trustcheck_{0}.pid"  -f ([Guid]::NewGuid().ToString("N")))

    $job = Start-Job -ScriptBlock {
        param($fp, $pidFile, $consts, $instanzCode, $eigenerOrdner)

        . ([scriptblock]::Create($instanzCode))

        # Vor New-Object erfassen, welche PowerPoint-Prozesse schon laufen
        # (Begruendung bei $script:PptInstanzCode).
        $vorher = Get-PptSitzungsPids
        $ppt    = New-Object -ComObject PowerPoint.Application
        $inst   = Get-PptInstanz -App $ppt -VorherPids $vorher
        if (-not $inst.Eigen) {
            try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ppt) | Out-Null } catch {}
            $wer = if ($inst.Pid) { "PID $($inst.Pid)" } else { "PID nicht zuzuordnen" }
            return @{ Ok = $false; Fremd = $true; PptPid = $null
                      Msg = "PowerPoint des Anwenders laeuft ($wer)" }
        }
        $myPptPid = $inst.Pid
        try { [System.IO.File]::WriteAllText($pidFile, "$myPptPid") } catch {}

        $ppt.Visible       = $consts.msoTrue
        $ppt.DisplayAlerts = $consts.ppAlertsNone
        $ppt.FeatureInstall = $consts.msoFeatureInstallNone
        try { $ppt.AutomationSecurity = $consts.msoAutomationSecurityForceDisable } catch {}
        try { $ppt.WindowState = $consts.ppWindowMinimized } catch {}

        $pres1 = $null
        $pres2 = $null
        $res   = $null
        try {
            # 1) Leere Praesentation erzeugen und in den Temp-Ordner speichern
            $pres1 = $ppt.Presentations.Add($consts.msoTrue)
            if (-not $pres1) {
                throw [System.InvalidOperationException]::new("Presentations.Add lieferte keine Praesentation")
            }
            $pres1.SaveAs($fp, $consts.ppSaveAsOpenXMLPresentation)
            $pres1.Close()
            $pres1 = $null

            # 2) Re-Open aus dem Temp-Ordner -> eigentlicher Trust-Center-Test
            $pres2 = $ppt.Presentations.Open($fp, -1, -1, 0)
            if (-not $pres2) {
                throw [System.InvalidOperationException]::new("Presentations.Open lieferte keine Praesentation")
            }
            $pres2.Close()
            $pres2 = $null

            $res = @{ Ok = $true }
        } catch {
            $res = @{ Ok = $false; Msg = $_.Exception.Message }
        } finally {
            if ($pres1) { try { $pres1.Close() } catch {} }
            if ($pres2) { try { $pres2.Close() } catch {} }
        }

        # Quit nur, wenn ausser den eigenen Arbeitsdateien nichts offen ist
        # (Begruendung in Convert-PptToPptx).
        $fremdOffen = Get-PptFremdePraesentationen -App $ppt -EigenerOrdner $eigenerOrdner
        if ($fremdOffen -eq 0) {
            try { $ppt.Quit() } catch {}
            $res.PptPid = $myPptPid
        } else {
            try { $ppt.WindowState   = 1 } catch {}
            try { $ppt.DisplayAlerts = 2 } catch {}
            $res.PptPid = $null
            $res.Ok     = $false
            $res.Fremd  = $true
            $res.Msg    = "In der PowerPoint-Instanz (PID $myPptPid) ist eine fremde Praesentation geoeffnet - PowerPoint NICHT beendet"
        }
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ppt) | Out-Null } catch {}
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        return $res
    } -ArgumentList $testPath, $pidFile, $script:PptConstants, $script:PptInstanzCode, ($TempDir.TrimEnd('\') + '\')

    $completed = Wait-Job $job -Timeout $TimeoutSec

    if (-not $completed) {
        try { Stop-Job   $job -ErrorAction SilentlyContinue -WhatIf:$false } catch {}
        try { Remove-Job $job -Force -ErrorAction SilentlyContinue -WhatIf:$false } catch {}
        if ([System.IO.File]::Exists($pidFile)) {
            try {
                $pidContent = [System.IO.File]::ReadAllText($pidFile).Trim()
                if ($pidContent -match '^\d+$') {
                    $savedPid = [int]$pidContent
                    if ($savedPid -gt 0) {
                        $script:TrackedPptPids.Add($savedPid)
                        # Nur die nachweislich eigene Instanz beenden: Name UND
                        # Startzeit nach Skriptbeginn (Test-IsOwnPptProcess).
                        # Die reine Namenspruefung schuetzte zwar vor PID-
                        # Recycling an Systemprozesse, nicht aber vor einer
                        # PowerPoint-Sitzung des Anwenders - deren PID stand
                        # bei uebernommener Sitzung selbst in der PID-Datei.
                        # Gleich wie Convert-PptToPptx und
                        # Clear-TrackedPowerPointInstances.
                        if (Test-IsOwnPptProcess -ProcessId $savedPid) {
                            try { Stop-Process -Id $savedPid -Force -ErrorAction SilentlyContinue -WhatIf:$false } catch {}
                            Write-DetailedLog "Trust-Center-Timeout: POWERPNT.EXE PID $savedPid sofort beendet" "WARN"
                        }
                    }
                }
            } catch {}
            try { [System.IO.File]::Delete($pidFile) } catch {}
        }
        try { [System.IO.File]::Delete($testPath) } catch {}
        return @{ Ok = $false; Msg = "TIMEOUT nach ${TimeoutSec}s - Trust Center vermutlich nicht konfiguriert oder Office-Profil beschaedigt" }
    }

    $result = Receive-Job $job
    try { Remove-Job $job -Force -ErrorAction SilentlyContinue -WhatIf:$false } catch {}
    if ([System.IO.File]::Exists($pidFile)) { try { [System.IO.File]::Delete($pidFile) } catch {} }
    try { [System.IO.File]::Delete($testPath) } catch {}
    if ($result -and $result.PptPid) { $script:TrackedPptPids.Add([int]$result.PptPid) }

    if ($result -and $result.Ok) {
        return @{ Ok = $true }
    } else {
        $m = if ($result) { $result.Msg } else { "Kein Ergebnis vom COM-Job" }
        return @{ Ok = $false; Msg = $m; Fremd = [bool]($result -and $result.Fremd) }
    }
}

# ==================================================================
# INTERAKTIVER START
# ==================================================================
if (-not $NoInteractive.IsPresent) {

    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║             POWERPOINT DEEP UNPROTECT              ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════╝" -ForegroundColor Cyan

    # Laufendes PowerPoint wird unmittelbar vor dem ersten COM-Zugriff
    # geprueft (Confirm-PowerPointGeschlossen, vor dem Smoke-Test).
    Write-Host ""
    Write-Host "HINWEIS: Das Skript verwendet diesen Arbeitsordner für die COM-Verarbeitung" -ForegroundColor DarkYellow
    Write-Host "(PowerPoint öffnet Dateien daraus; nicht synchronisiert, anders als 'Dokumente'):" -ForegroundColor DarkYellow
    Write-Host "       $($script:ArbeitsBasis)" -ForegroundColor White
    Write-Host "" -ForegroundColor DarkYellow
    Write-Host "Bitte in PowerPoint sicherstellen:" -ForegroundColor DarkYellow
    Write-Host "  Datei > Optionen > Trust Center > Einstellungen für das Trust Center" -ForegroundColor Gray
    Write-Host "  > Vertrauenswürdige Speicherorte: diesen Ordner hinzufügen" -ForegroundColor Gray
    Write-Host "       [x] Unterordner dieser Speicherorte sind ebenfalls vertrauenswürdig" -ForegroundColor Gray
    Write-Host ""

    if ([string]::IsNullOrWhiteSpace($TargetPath)) {
        $desktopPath   = Get-UserDesktopPath
        $downloadsPath = Get-UserDownloadsPath

        $presetCount = $script:DirectoryPresets.Count
        Write-Host "Zielpfad auswählen:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $presetCount; $i++) {
            Write-Host ("  [" + ($i + 1) + "] " + $script:DirectoryPresets[$i]) -ForegroundColor White
        }
        # Verkettung statt -f: Pfade koennen '{' oder '}' enthalten.
        Write-Host ("  [" + ($presetCount + 1) + "] Desktop      ($desktopPath)")   -ForegroundColor White
        Write-Host ("  [" + ($presetCount + 2) + "] Downloads    ($downloadsPath)") -ForegroundColor White
        Write-Host ("  [" + ($presetCount + 3) + "] Eigenen Pfad eingeben")         -ForegroundColor White
        Write-Host ""
        $pathChoice = (Read-Host ("Auswahl [1-" + ($presetCount + 3) + "]")).Trim()

        $choiceNum = 0
        [void][int]::TryParse($pathChoice, [ref]$choiceNum)

        switch ($true) {
            ($choiceNum -ge 1 -and $choiceNum -le $presetCount) {
                $TargetPath = $script:DirectoryPresets[$choiceNum - 1]
                break
            }
            ($choiceNum -eq $presetCount + 1) { $TargetPath = $desktopPath;   break }
            ($choiceNum -eq $presetCount + 2) { $TargetPath = $downloadsPath; break }
            default {
                $TargetPath = (Read-Host "Pfad eingeben").Trim().Trim('"',"'")
                if ($TargetPath -match '^[a-zA-Z]:$') { $TargetPath += '\' }
                if ([string]::IsNullOrWhiteSpace($TargetPath)) {
                    Write-Error "Kein Pfad angegeben. Abbruch."
                    exit 1
                }
            }
        }
    }

    Write-Host ""
    Write-Host "Fortschrittsanzeige:" -ForegroundColor Yellow
    Write-Host "  [1] Keine Fortschrittsleiste (laufende Zählung inline) [Standard]" -ForegroundColor White
    Write-Host "  [2] Fortschrittsleiste mit ETA (Vorab-Scan nötig, Start verzögert!)" -ForegroundColor White
    Write-Host "  [3] Fortschrittsleiste ohne ETA (sofortiger Start, kein Total)" -ForegroundColor White
    Write-Host ""
    Write-Host "  Hinweis zu [2]: Bei UNC-Pfaden oder sehr grossen Freigaben kann" -ForegroundColor DarkGray
    Write-Host "  die Indizierung 10-30 Minuten dauern, bevor die erste Datei" -ForegroundColor DarkGray
    Write-Host "  verarbeitet wird. Bei > 100.000 Dateien besser [1] oder [3]" -ForegroundColor DarkGray
    Write-Host "  waehlen (Generator-Variante ohne Vorab-Zaehlung)." -ForegroundColor DarkGray
    Write-Host ""
    $progressChoice = (Read-Host "Auswahl [1-3]").Trim()
    switch ($progressChoice) {
        "2"     { $script:UseProgress = $true;  $script:SkipPreScan = $false }
        "3"     { $script:UseProgress = $true;  $script:SkipPreScan = $true  }
        default { $script:UseProgress = $false; $script:SkipPreScan = $true  }
    }

    Write-Host ""
} else {
    if ([string]::IsNullOrWhiteSpace($TargetPath)) {
        Write-Error "NoInteractive-Modus erfordert -TargetPath."
        exit 1
    }
    $script:UseProgress = $ShowProgress.IsPresent
    $script:SkipPreScan = $true
    # Laufendes PowerPoint: siehe Confirm-PowerPointGeschlossen vor dem
    # Smoke-Test. Die fruehere stille Warnung ("sie werden geschuetzt")
    # liess den Lauf trotzdem an der Sitzung des Anwenders arbeiten.
}

# ==================================================================
# VORBEREITUNG
# ==================================================================
if (-not (Test-PowerPointInstalled)) {
    Write-Error "PowerPoint (COM-Komponente PowerPoint.Application) ist auf diesem System nicht registriert. Abbruch."
    exit 2
}

if (-not (Test-Path -LiteralPath $TargetPath)) {
    Write-Error "Pfad nicht gefunden: $TargetPath"
    exit 1
}
$tgtItem = Get-Item -LiteralPath $TargetPath -ErrorAction SilentlyContinue
if ($tgtItem -and -not $tgtItem.PSIsContainer) {
    Write-Error "Zielpfad ist eine Datei, kein Verzeichnis: $TargetPath"
    exit 1
}

# Absoluten Pfad festschreiben. Test-Path loest einen relativen Pfad gegen
# den PowerShell-Ort ($PWD) auf, [System.IO] dagegen gegen das Arbeits-
# verzeichnis des Prozesses - die beiden laufen nach einem Set-Location
# auseinander. Gemessen unter 5.1 (PWD = ...\A, Prozess-CWD = ...\B):
# Test-Path 'ziel' pruefte A\ziel, EnumerateFiles('ziel') lieferte B\ziel -
# das Skript haette einen fremden Baum bearbeitet. Gilt auch fuer die
# manuelle Pfadeingabe im Menue.
try {
    $TargetPath = (Resolve-Path -LiteralPath $TargetPath -ErrorAction Stop).ProviderPath
} catch {
    Write-Error "Pfad nicht aufloesbar: $TargetPath - $_"
    exit 1
}

Remove-StaleTempFolders

if (Test-Path -LiteralPath $script:TempPath) {
    Remove-Item -LiteralPath $script:TempPath -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false
}
# Provider-frei anlegen statt per New-Item: Das Skript deklariert
# SupportsShouldProcess, und unter -WhatIf gehorcht jedes ShouldProcess-faehige
# Cmdlet dem $WhatIfPreference des Skript-Scopes - auch New-Item. Der
# Arbeitsordner entstuende dann nicht, der PowerPoint-Smoke-Test darunter
# schlaegt fehl, und der Lauf endet vor der ersten Datei. Die Protokolle sind
# hier nicht betroffen, weil Initialize-Loggers StreamWriter verwendet.
[System.IO.Directory]::CreateDirectory($script:TempPath) | Out-Null

# Restore-Privilegien (Admin-Kontext) fuer ACL/Owner-Erhalt aktivieren.
Enable-RestorePrivileges

Initialize-Loggers

Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host " POWERPOINT DEEP UNPROTECT"
Write-Host " Ziel: $TargetPath"
Write-Host "=====================================================" -ForegroundColor Cyan

# ------------------------------------------------------------------
# Laufendes PowerPoint? Dann KEIN COM (Begruendung bei
# Confirm-PowerPointGeschlossen). Exitcode 3 = PowerPoint laeuft.
# ------------------------------------------------------------------
if (-not (Confirm-PowerPointGeschlossen -NoInteractive:$NoInteractive)) {
    Write-Host "Abbruch: PowerPoint laeuft. Bitte PowerPoint beenden und das Skript neu starten." -ForegroundColor Red
    Write-Log "Abbruch: PowerPoint laeuft bereits - COM-Automatisierung wuerde die Sitzung des Anwenders uebernehmen." "ERROR"
    Close-Loggers
    if (Test-Path -LiteralPath $script:TempPath) { Remove-Item -LiteralPath $script:TempPath -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false }
    exit 3
}

# ------------------------------------------------------------------
# PowerPoint Smoke-Test
# ------------------------------------------------------------------
Write-Host "Prüfe COM-Subsystem und Trust-Center..." -ForegroundColor DarkGray
$smokeTest = Test-PowerPointTrustCenter -TempDir $script:TempPath -TimeoutSec $script:TrustCenterTimeoutSec

if ($smokeTest.Fremd) {
    # Zwischen Pruefung und Smoke-Test wurde PowerPoint gestartet: kein
    # "Trotzdem fortfahren" - jeder weitere Job haette dieselbe Sitzung.
    Write-Host ""
    Write-Host "ABBRUCH: $($smokeTest.Msg)" -ForegroundColor Red
    Write-Host "Bitte PowerPoint beenden und das Skript neu starten." -ForegroundColor Red
    Write-Log "Abbruch im Smoke-Test: $($smokeTest.Msg)" "ERROR"
    Clear-TrackedPowerPointInstances
    Close-Loggers
    if (Test-Path -LiteralPath $script:TempPath) { Remove-Item -LiteralPath $script:TempPath -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false }
    exit 3
}

if (-not $smokeTest.Ok) {
    Write-Host ""
    Write-Host "POWERPOINT-SMOKE-TEST FEHLGESCHLAGEN" -ForegroundColor Red
    Write-Host "   Grund: $($smokeTest.Msg)"          -ForegroundColor Yellow
    Write-Host ""
    Write-Host "   Mögliche Ursachen:" -ForegroundColor Gray
    Write-Host "     - Arbeitsordner (siehe Hinweis beim Start) nicht als vertrauenswürdiger Speicherort eingetragen" -ForegroundColor Gray
    Write-Host "     - PowerPoint nicht installiert oder Office-Profil beschädigt" -ForegroundColor Gray
    Write-Host "     - Fehlender Desktop-Ordner für Dienstkonto" -ForegroundColor Gray
    Write-Host ""
    Write-Log "PowerPoint-Smoke-Test fehlgeschlagen: $($smokeTest.Msg)" "ERROR"

    if ($NoInteractive.IsPresent) {
        Write-Log "Abbruch (NoInteractive) nach fehlgeschlagenem Smoke-Test." "ERROR"
        Clear-TrackedPowerPointInstances
        Close-Loggers
        if (Test-Path $script:TempPath) { Remove-Item $script:TempPath -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false }
        exit 2
    }

    $choice = (Read-Host "Trotzdem fortfahren? (Timeout-Risiko pro Datei!) [j/N]").Trim().ToUpperInvariant()
    if ($choice -ne "J" -and $choice -ne "Y") {
        Write-Host "Abbruch durch Benutzer." -ForegroundColor Yellow
        Write-Log "Abbruch durch Benutzer nach Smoke-Test-Warnung." "INFO"
        Clear-TrackedPowerPointInstances
        Close-Loggers
        if (Test-Path $script:TempPath) { Remove-Item $script:TempPath -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false }
        exit 0
    }
    Write-Log "Smoke-Test-Warnung vom Benutzer ignoriert - Fortsetzung." "WARN"
} else {
    Write-Host "Smoke-Test OK." -ForegroundColor DarkGray
    Write-DetailedLog "PowerPoint-Smoke-Test erfolgreich." "DEBUG"
}

$allPptExt = "^\.p(pt|ot|ps)[xm]?$"

# ------------------------------------------------------------------
# Verwaiste Backups frueherer Laeufe aufraeumen
# ------------------------------------------------------------------
$backupCleanup = $null
if (-not $SkipBackupCleanup) {
    Write-Host "Suche zurueckgebliebene Backups frueherer Laeufe..." -ForegroundColor DarkGray
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
        Write-Host "Diese wurden NICHT geloescht - sie koennten die einzige Kopie sein:" -ForegroundColor Yellow
        $shown = 0
        foreach ($o in $backupCleanup.OrphanList) {
            Write-Host "  $o" -ForegroundColor Gray
            $shown++
            if ($shown -ge 10) {
                Write-Host "  ... und $($backupCleanup.OrphanList.Count - 10) weitere (siehe Log)" -ForegroundColor Gray
                break
            }
        }
        Write-Host "Zum Wiederherstellen die Endung '.bak_xxxxxxxx' entfernen." -ForegroundColor Gray
        Write-Host ""
    }
    Write-Log ("Backup-Aufraeumen: {0} entfernt, {1} behalten, {2} ohne Original, {3} Fehler" -f `
               $backupCleanup.Deleted, $backupCleanup.Kept, $backupCleanup.Orphans, $backupCleanup.Failed) "INFO"
} else {
    Write-DetailedLog "Backup-Aufraeumen uebersprungen (-SkipBackupCleanup)." "DEBUG"
}

if ($script:UseProgress -and -not $script:SkipPreScan) {
    Write-Host "Zaehle Dateien fuer ETA (bitte warten)..." -ForegroundColor DarkGray
    $etaCount = 0
    Get-PptFilesRobust (Add-LongPathPrefix $TargetPath) $allPptExt | ForEach-Object { $etaCount++ }
    $script:TotalFiles = $etaCount
    Write-Host "Gefunden: $($script:TotalFiles) PowerPoint-Dateien`n" -ForegroundColor DarkGray
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
    # Korrupte Dateien und COM-Timeouts getrennt fuehren - beide
    # wurden bisher auf 'Encrypted' gebucht und in der
    # Endstatistik als verschluesselt ausgewiesen.
    Corrupt     = 0
    Timeouts    = 0
    AvBlocked   = 0
    WouldChange = 0
}

# ==================================================================
# HAUPTSCHLEIFE
# ==================================================================
Write-Host "Starte Verarbeitung..." -ForegroundColor Cyan

# Strg+C ab hier als Eingabe behandeln und in der Schleife abfragen (siehe
# "ABBRUCH PER TASTE"). Erst jetzt, damit die Read-Host-Abfragen davor sich
# normal abbrechen lassen. Nur mit echter, nicht umgeleiteter Konsole; mit
# -NoInteractive fragt niemand Tasten ab.
if (-not $NoInteractive.IsPresent) {
    try {
        if (-not [Console]::IsInputRedirected) {
            [Console]::TreatControlCAsInput = $true
            $script:CtrlCAsInput = $true
            Write-Host "Abbruch mit Strg+C oder ESC (die laufende Datei wird fertiggestellt)." -ForegroundColor DarkGray
        }
    } catch {}
}

try {
Get-PptFilesRobust (Add-LongPathPrefix $TargetPath) $allPptExt |
    ForEach-Object {

    Test-AbortRequested
    if ($script:ShouldStop) { throw [System.OperationCanceledException]::new() }
    # Fremde PowerPoint-Sitzung erkannt (Job meldete FREMD oder der Anwender
    # hat in der Automatisierungs-Instanz etwas geoeffnet): nicht weiter per
    # COM arbeiten.
    if ($script:FremdePptAbbruch) { throw [System.OperationCanceledException]::new() }

    $filePath = $_
    $fileName = [System.IO.Path]::GetFileName($filePath)
    $fileExt  = [System.IO.Path]::GetExtension($filePath)

    $workFile     = $null
    $backupPath   = $null
    $wasConverted = $false

    $stats.Processed++
    $script:ProcessedCount = $stats.Processed
    Update-Progress -Status "Verarbeite..." -CurrentFile $filePath

    if (-not $script:UseProgress) {
        Write-Host -NoNewline "[#$($stats.Processed)] $fileName "
    }

    # 1. Junk-Dateien (~$temp, ._macOS)
    # '~$'-Dateien sind PowerPoint-SPERRDATEIEN: sie existieren, solange eine
    # Praesentation geoeffnet ist. Sofortiges Loeschen bricht die Sperre.
    if ($fileName.StartsWith("~`$") -or $fileName.StartsWith("._")) {
        # $srcLong wurde hier frueher BENUTZT, bevor es weiter unten
        # zugewiesen wurde. In PowerShell ist eine nicht zugewiesene
        # Variable $null, GetLastWriteTime($null) wirft - und der leere
        # catch-Block liess $ageHours auf 0.0 stehen. Damit galt JEDE
        # Sperrdatei als "zu jung" und wurde geschont: das Aufraeumen
        # verwaister ~$-Dateien hat in diesem Skript nie stattgefunden,
        # und der Zaehler "Sperrdateien geschont" war entsprechend
        # aufgeblaeht. 2a und 2b holen das Alter korrekt.
        $ageHours = 0.0
        try {
            $ageHours = ((Get-Date) - ([System.IO.File]::GetLastWriteTime((Add-LongPathPrefix $filePath)))).TotalHours
        } catch {
            # Alter nicht ermittelbar: konservativ schonen (wie bisher),
            # aber sichtbar machen statt still zu verschlucken.
            Write-DetailedLog "Alter der Sperrdatei nicht ermittelbar, wird geschont: $filePath - $_" "WARN"
        }
        if ($ageHours -lt $script:JunkMinAgeHours) {
            $stats.JunkKept++
            $stats.Skipped++
            if (-not $script:UseProgress) {
                Write-Host ("-> SKIP (Sperrdatei, erst {0:N1} h alt)" -f $ageHours) -ForegroundColor DarkGray
            }
            Write-CsvLog -Status "SKIP" -Actions "Sperrdatei geschont" -Path $filePath
            return
        }
        if (-not (Confirm-Write $filePath 'Verwaiste Sperrdatei loeschen')) {
            $stats.WouldChange++
            if (-not $script:UseProgress) { Write-Host "-> WHATIF (Junk)" -ForegroundColor DarkCyan }
            Write-CsvLog -Status "WHATIF" -Actions "Wuerde Sperrdatei loeschen" -Path $filePath
            return
        }
        # Nur zaehlen, was wirklich geloescht ist. Vorher lief $stats.Junk++
        # auch nach einem gescheiterten Delete (leerer catch). Scheitert das
        # Loeschen, ist die Sperrdatei meist doch in Benutzung - dann schonen.
        $junkWeg = $false
        try {
            [System.IO.File]::Delete((Add-LongPathPrefix $filePath))
            $junkWeg = -not [System.IO.File]::Exists((Add-LongPathPrefix $filePath))
        } catch {
            Write-DetailedLog "Sperrdatei nicht loeschbar: $filePath - $_" "WARN"
        }
        if (-not $junkWeg) {
            $stats.Skipped++
            if (-not $script:UseProgress) { Write-Host "-> SKIP (Sperrdatei nicht loeschbar)" -ForegroundColor Yellow }
            Write-Log "Sperrdatei nicht loeschbar (in Benutzung?): $filePath" "WARN"
            Write-CsvLog -Status "SKIP" -Actions "Sperrdatei nicht loeschbar" -Path $filePath
            return
        }
        $stats.Junk++
        if (-not $script:UseProgress) { Write-Host "-> JUNK" -ForegroundColor DarkGray }
        Write-DetailedLog "Junk entfernt: $filePath" "DEBUG"
        Write-CsvLog -Status "JUNK" -Actions "Sperrdatei geloescht" -Path $filePath
        return
    }

    $srcLong = Add-LongPathPrefix $filePath

    # 2. Schreibschutz-Attribut: hier nur FESTSTELLEN. Entfernt wird es erst,
    #    wenn das Ergebnis feststeht - beim Rueckschreiben bzw. im Zweig
    #    "Kein Schutz" - und jeweils nach Confirm-Write, mit Ausweis im
    #    Protokoll. Vorher wurde es hier fuer JEDE Datei entfernt, auch fuer
    #    danach gesperrte, verschluesselte oder defekte, ohne Eintrag in der
    #    CSV (dort stand "Kein Schutz gefunden"). Im Probelauf blieb es
    #    stehen, und die Sperrpruefung meldete die Datei als "Zugriff
    #    verweigert" (siehe Test-FileIsLocked). Das Entfernen selbst bleibt
    #    gewollt: es ist Teil von "Schutz entfernen", wie in 2a und 2b.
    $istReadOnly = $false
    try {
        $istReadOnly = [bool]([System.IO.File]::GetAttributes($srcLong) -band [System.IO.FileAttributes]::ReadOnly)
    } catch {}
    $roEntfernt = $false

    # 3. Sperr-Check
    $lockResult = Test-FileIsLocked -FilePath $filePath
    if ($lockResult) {
        $stats.Locked++
        $stats.Skipped++
        if ($lockResult -eq "AccessDenied") {
            if (-not $script:UseProgress) { Write-Host "-> SKIP (Zugriff verweigert)" -ForegroundColor Red }
            Write-Log "Zugriff verweigert: $filePath" "WARN"
            Write-DetailedLog "Keine NTFS-Berechtigung: $filePath" "WARN"
            Write-CsvLog -Status "SKIP" -Actions "Zugriff verweigert" -Path $filePath
        } else {
            if (-not $script:UseProgress) { Write-Host "-> SKIP (gesperrt)" -ForegroundColor Yellow }
            Write-Log "Gesperrt/geoeffnet: $filePath" "WARN"
            Write-DetailedLog "Datei gesperrt: $filePath" "WARN"
            Write-CsvLog -Status "SKIP" -Actions "Gesperrt/geoeffnet" -Path $filePath
        }
        return
    }

    try {
        $origCreationTimeUtc   = [System.IO.File]::GetCreationTimeUtc($srcLong)
        $origLastWriteTimeUtc  = [System.IO.File]::GetLastWriteTimeUtc($srcLong)
        $origLastAccessTimeUtc = [System.IO.File]::GetLastAccessTimeUtc($srcLong)
    } catch {
        # Ohne diesen Schutz wuerde eine Ausnahme hier (Netzwerk-Drop,
        # AV-Lock) ausserhalb des Datei-try landen und via trap den
        # gesamten Lauf abbrechen.
        $now = [DateTime]::UtcNow
        $origCreationTimeUtc   = $now
        $origLastWriteTimeUtc  = $now
        $origLastAccessTimeUtc = $now
        Write-DetailedLog "Zeitstempel-Auslesen fehlgeschlagen, Fallback auf Now: $filePath - $_" "WARN"
    }

    # --- NTFS-Sicherheitsinfo (Owner/Group/DACL) sichern ---
    $origSecurity = Get-FileSecuritySnapshot -Path $srcLong

    # 0-Byte-Platzhalter aus Reserve-UniqueDestination (siehe finally).
    $reservedPlaceholder = $null

    $guid     = [Guid]::NewGuid().ToString()
    $tempFile = Join-Path $script:TempPath "$guid$fileExt"
    $tempBase = Join-Path $script:TempPath $guid
    $workFile = $tempFile

    try {
        # 4. In Temp kopieren (mit Retry gegen AV-Locks)
        Invoke-WithRetry -Context "Copy src->temp" -Action { [System.IO.File]::Copy($srcLong, $tempFile, $true) }
        # File.Copy uebernimmt das ReadOnly-Attribut auf die Temp-Kopie
        # (gemessen unter 5.1), und ZipFile.Open(Update) scheitert daran mit
        # UnauthorizedAccessException - nach ~36 s Retry. Seit das Attribut am
        # Original erst beim Rueckschreiben entfernt wird, traefe das jede
        # schreibgeschuetzte Datei; die Temp-Kopie gehoert dem Skript.
        [void](Clear-ReadOnlyAttribut $tempFile)

        # Priming-Read: AV synchron abschliessen lassen
        Invoke-FilePrimingRead -Path $tempFile

        # AV-Stabilitaets-Polling (exklusiver Open + Groessen-Stabilitaet)
        if (-not (Wait-FileStable -Path $tempFile)) {
            $stats.AvBlocked++
            $stats.Skipped++
            if (-not $script:UseProgress) { Write-Host "-> SKIP (AV-Lock anhaltend)" -ForegroundColor Magenta }
            Write-Log "Uebersprungen (AV-Lock anhaltend nach Polling): $filePath" "WARN"
            Write-DetailedLog "Wait-FileStable: Datei nicht stabil/exklusiv oeffenbar: $tempFile" "WARN"
            Write-CsvLog -Status "SKIP" -Actions "AV-Lock anhaltend" -Path $filePath
            return
        }

        $hadZoneIdentifier = $false
        try {
            $streams = Get-Item -LiteralPath $tempFile -Stream "Zone.Identifier" -ErrorAction SilentlyContinue
            if ($streams) { $hadZoneIdentifier = $true }
        } catch {}
        if ($hadZoneIdentifier) {
            Unblock-File -LiteralPath $tempFile -ErrorAction SilentlyContinue -WhatIf:$false
            Write-DetailedLog "Zone.Identifier entfernt auf Temp-Kopie: $fileName" "DEBUG"
            # Unblock-File kann einen erneuten AV-Scan triggern (sehr klassischer
            # Defender-Trigger nach ADS-Aenderung). Daher Wait-FileStable nochmal
            # laufen lassen, damit ZipFile.Open nicht direkt in den Scan laeuft.
            if (-not (Wait-FileStable -Path $tempFile)) {
                $stats.AvBlocked++
                $stats.Skipped++
                if (-not $script:UseProgress) { Write-Host "-> SKIP (AV-Lock nach Unblock-File)" -ForegroundColor Magenta }
                Write-Log "Uebersprungen (AV-Lock nach Unblock-File): $filePath" "WARN"
                Write-DetailedLog "Wait-FileStable nach Unblock-File fehlgeschlagen: $tempFile" "WARN"
                Write-CsvLog -Status "SKIP" -Actions "AV-Lock nach Unblock-File" -Path $filePath
                return
            }
        }

        # 5. Konvertierung alter Formate
        $ext = $fileExt.ToLower()

        if ($ext -match "^\.p(pt|ot|ps)$") {
            $isOleEncrypted = Test-IsOleEncrypted -FilePath $tempFile

            if ($isOleEncrypted) {
                $stats.Encrypted++
                $stats.Skipped++
                if (-not $script:UseProgress) { Write-Host "-> SKIP (verschluesselt)" -ForegroundColor Magenta }
                Write-Log "Uebersprungen (OLE verschluesselt): $filePath" "WARN"
                Write-DetailedLog "OLE-Datei ist verschluesselt: $filePath" "WARN"
                Write-CsvLog -Status "SKIP" -Actions "Verschluesselt (kein PW-Parameter in PowerPoint-COM)" -Path $filePath
                return
            }

            try {
                if (-not $script:UseProgress) { Write-Host -NoNewline "[Conv] " -ForegroundColor Cyan }
                $workFile     = Convert-PptToPptx -SourcePath $tempFile -DestPathBase $tempBase -OriginalExt $ext
                $wasConverted = $true
                # $stats.Converted wird erst nach erfolgreichem Rueckschreiben
                # gezaehlt (siehe dort) - vorher hier, und ein gescheitertes
                # Rueckschreiben liess den Zaehler zu hoch stehen.
                Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false
                Write-DetailedLog "Konvertiert: $filePath -> $workFile" "DEBUG"
            } catch [System.TimeoutException] {
                Clear-TrackedPowerPointInstances
                $stats.Skipped++
                $stats.Timeouts++
                if (-not $script:UseProgress) { Write-Host "-> SKIP (Timeout)" -ForegroundColor Magenta }
                Write-Log "Uebersprungen (Timeout): $filePath - $_" "WARN"
                Write-CsvLog -Status "SKIP" -Actions "Timeout" -Path $filePath -Details $_.Exception.Message
                return
            } catch {
                $stats.Errors++
                if ($script:FremdePptAbbruch) {
                    # Job hat eine fremde PowerPoint-Sitzung erkannt; die
                    # Hauptschleife bricht vor der naechsten Datei ab.
                    if (-not $script:UseProgress) { Write-Host "-> ABBRUCH (PowerPoint des Anwenders laeuft)" -ForegroundColor Red }
                    Write-CsvLog -Status "ERR" -Actions "PowerPoint des Anwenders laeuft - Lauf abgebrochen" -Path $filePath -Details $_.Exception.Message
                    return
                }
                if (-not $script:UseProgress) { Write-Host "-> ERR (Konvert)" -ForegroundColor Red }
                Write-Log "Konvert-Fehler: $filePath - $_" "ERROR"
                Write-DetailedLog "Konvert-Fehler Detail: $_" "ERROR"
                Write-CsvLog -Status "ERR" -Actions "Konvert-Fehler" -Path $filePath -Details $_.Exception.Message
                return
            }
        } else {
            # OOXML mit Oeffnungspasswort = OLE2/CFBF-Container mit EncryptedPackage
            if (Test-IsOleEncrypted -FilePath $tempFile) {
                $stats.Encrypted++
                $stats.Skipped++
                if (-not $script:UseProgress) { Write-Host "-> SKIP (verschluesselt)" -ForegroundColor Magenta }
                Write-Log "Uebersprungen (OOXML verschluesselt): $filePath" "WARN"
                Write-DetailedLog "OOXML-Datei ist verschluesselt (OLE-Container): $filePath" "WARN"
                Write-CsvLog -Status "SKIP" -Actions "Verschluesselt (OLE-Container)" -Path $filePath
                return
            }

            $zipStatus = Test-IsValidZip -FilePath $tempFile
            if ($zipStatus -eq "Locked") {
                $stats.AvBlocked++
                $stats.Skipped++
                if (-not $script:UseProgress) { Write-Host "-> SKIP (AV-Lock)" -ForegroundColor Magenta }
                Write-Log "Uebersprungen (AV-Lock nach Retry): $filePath" "WARN"
                Write-DetailedLog "Datei nach Retry weiterhin gesperrt: $filePath" "WARN"
                Write-CsvLog -Status "SKIP" -Actions "AV-Lock nach Retry" -Path $filePath
                return
            }
            if ($zipStatus -eq "Invalid") {
                # Korrupt, nicht verschluesselt: die Meldung sagt 'korrupt',
                # gezaehlt wurde bisher aber auf 'Encrypted'. In der
                # Endstatistik erschienen defekte Dateien damit als
                # kennwortgeschuetzt - der Betreiber suchte nach Passwoertern
                # statt die Datei zu ersetzen.
                $stats.Corrupt++
                $stats.Skipped++
                if (-not $script:UseProgress) { Write-Host "-> SKIP (korrupt)" -ForegroundColor Magenta }
                Write-Log "Uebersprungen (kein gueltiges ZIP-Archiv): $filePath" "WARN"
                Write-DetailedLog "OOXML-Datei ist kein gueltiges ZIP: $filePath" "WARN"
                Write-CsvLog -Status "SKIP" -Actions "Kein gueltiges ZIP (korrupt)" -Path $filePath
                return
            }
            $workFile = $tempFile
        }

        # 6. Schutz per XML entfernen
        $actions    = Remove-PptxProtection -FilePath $workFile
        $actionsLog = $actions -join ', '
        Write-DetailedLog "Aktionen ($fileName): $actionsLog" "DEBUG"

        # 7. Rueckschreiben
        if (-not [string]::IsNullOrWhiteSpace($actionsLog) -or $wasConverted) {

            $actionStrPre = $actionsLog
            if ($wasConverted) {
                $actionStrPre = (@('Konvertierung') + @($actions)) -join ', '
            }
            if ($istReadOnly) { $actionStrPre = (@($actionStrPre, 'Schreibschutz-Attribut') | Where-Object { $_ }) -join ', ' }

            # Zentrale Freigabe: ab hier wird das Original angefasst.
            if (-not (Confirm-Write $filePath "Schutz entfernen ($actionStrPre)")) {
                $stats.WouldChange++
                if (-not $script:UseProgress) {
                    Write-Host "-> WHATIF ($actionStrPre)" -ForegroundColor DarkCyan
                }
                Write-Log    "Simulation: wuerde aendern: $filePath  [$actionStrPre]" "INFO"
                Write-CsvLog -Status "WHATIF" -Actions $actionStrPre -Path $filePath
                return
            }

            $newExt    = [System.IO.Path]::GetExtension($workFile)
            $finalDest = [System.IO.Path]::ChangeExtension($srcLong, $newExt)

            if ($finalDest -ne $srcLong -and [System.IO.File]::Exists($finalDest)) {
                $base = [System.IO.Path]::Combine(
                            [System.IO.Path]::GetDirectoryName($finalDest),
                            [System.IO.Path]::GetFileNameWithoutExtension($finalDest))

                $reserved = Reserve-UniqueDestination -BasePath $base -Extension $newExt
                if (-not $reserved) {
                    $stats.Errors++
                    if (-not $script:UseProgress) { Write-Host "-> ERR (>$($script:MaxNameClashRetries) Namenskonflikte)" -ForegroundColor Red }
                    Write-Log "Namenskonflikt-Limit ($($script:MaxNameClashRetries)) erreicht: $filePath" "ERROR"
                    Write-DetailedLog "Alle Umbenennungen _2 bis _$($script:MaxNameClashRetries) existieren: $filePath" "ERROR"
                    # Vorher ohne CSV-Zeile: die Datei fehlte im Ergebnisprotokoll.
                    Write-CsvLog -Status "ERR" -Actions "Namenskonflikt-Limit erreicht" -Path $filePath -Details "_2 bis _$($script:MaxNameClashRetries) belegt"
                    return
                }

                $finalDest           = $reserved
                $reservedPlaceholder = $reserved
                Write-DetailedLog "Namenskonflikt: Ziel reserviert als $finalDest" "WARN"
                if (-not $script:UseProgress) { Write-Host -NoNewline "[Umben.] " -ForegroundColor Yellow }
            }

            # Eigener, eindeutiger Backup-Name (siehe $script:EigenesBackupMuster).
            # Ein vorhandenes '.bak' des Anwenders wird nicht mehr umbenannt
            # oder ueberschrieben.
            $backupPath = New-BackupPfad -Original $srcLong

            # Schreibschutz erst jetzt entfernen (nach der Freigabe) - noetig
            # fuer Delete/Copy am Original und damit das Backup ihn nicht erbt.
            if ($istReadOnly) {
                $roEntfernt = Clear-ReadOnlyAttribut $srcLong
                if ($roEntfernt) { Write-DetailedLog "Datei-Schreibschutz (Attribut) entfernt: $filePath" "DEBUG" }
            }

            Invoke-WithRetry -Action { [System.IO.File]::Copy($srcLong, $backupPath, $true) }
            # Erzeugungszeit ausdruecklich stempeln: Remove-OrphanedBackups
            # misst das Alter daran (File.Copy uebernimmt die LastWriteTime der
            # Quelle; die CreationTime koennte per NTFS-Tunneling alt sein).
            try { [System.IO.File]::SetCreationTime($backupPath, (Get-Date)) } catch {}

            try {
                if ($finalDest -eq $srcLong -and [System.IO.File]::Exists($finalDest)) {
                    [System.IO.File]::Delete($finalDest)
                }
                Invoke-WithRetry -Action { [System.IO.File]::Copy($workFile, $finalDest, $true) }
                if (-not [System.IO.File]::Exists($finalDest)) { throw "Verifizierung fehlgeschlagen" }
                $reservedPlaceholder = $null   # Ziel ist real geschrieben

                try { [System.IO.File]::Delete($backupPath) } catch {}
                if ($wasConverted -and [System.IO.File]::Exists($srcLong) -and ($finalDest -ne $srcLong)) {
                    try { [System.IO.File]::Delete($srcLong) }
                    catch {
                        Write-DetailedLog "Alte Quelldatei konnte nicht geloescht werden (wird beibehalten): $srcLong - $_" "WARN"
                        Write-Log "Quelldatei beibehalten: $filePath - $_" "WARN"
                    }
                }

                try {
                    [System.IO.File]::SetCreationTimeUtc($finalDest,   $origCreationTimeUtc)
                    [System.IO.File]::SetLastWriteTimeUtc($finalDest,  $origLastWriteTimeUtc)
                    [System.IO.File]::SetLastAccessTimeUtc($finalDest, $origLastAccessTimeUtc)
                } catch {
                    Write-DetailedLog "Zeitstempel-Wiederherstellung fehlgeschlagen: $finalDest - $_" "WARN"
                }

                # --- ACL/Owner des Originals auf die neue Datei uebertragen ---
                # (Admin-Kontext: vollstaendig; Nutzer-Kontext: DACL). Hier
                # besonders relevant: der In-Place-Pfad loescht das Original
                # vor dem Copy - die neue Datei erbte sonst nur Ordner-ACEs.
                Set-FileSecuritySnapshot -Path $finalDest -Snapshot $origSecurity

                $stats.Unlocked++
                if ($wasConverted) { $stats.Converted++ }
                $actionStr = $actions -join ', '
                if ([string]::IsNullOrWhiteSpace($actionStr)) { $actionStr = "Nur Format konvertiert" }
                if ($roEntfernt) { $actionStr += ", Schreibschutz-Attribut" }
                if (-not $script:UseProgress) {
                    Write-Host "-> OK ($actionStr)" -ForegroundColor Green
                }
                Write-Log "Erledigt: $filePath  [$actionStr]" "INFO"
                Write-CsvLog -Status "OK" -Actions $actionStr -Path $filePath -Details $finalDest

            } catch {
                $stats.Errors++
                if ([System.IO.File]::Exists($finalDest)) {
                    try { [System.IO.File]::Delete($finalDest) } catch {}
                    Write-DetailedLog "Unvollstaendige Zieldatei entfernt: $finalDest" "WARN"
                }
                # Original unveraendert? (Konvertierung: .ppt nie angefasst;
                # In-Place: nur nach erfolgreichem Rollback.)
                $originalIntakt = ($finalDest -ne $srcLong)
                if ([System.IO.File]::Exists($backupPath)) {
                    $backupSafeToDelete = $true
                    if ($finalDest -eq $srcLong) {
                        $backupSafeToDelete = $false
                        try {
                            [System.IO.File]::Copy($backupPath, $srcLong, $true)
                            if ([System.IO.File]::Exists($srcLong)) {
                                $backupSafeToDelete = $true
                                $originalIntakt     = $true
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
                # Nichts geschrieben -> auch das Schreibschutz-Attribut zurueck.
                if ($roEntfernt -and $originalIntakt -and [System.IO.File]::Exists($srcLong)) {
                    Set-ReadOnlyAttribut $srcLong
                }
                if (-not $script:UseProgress) { Write-Host "-> ERR (Schreiben)" -ForegroundColor Red }
                Write-Log "Schreib-Fehler: $filePath - $_" "ERROR"
                Write-DetailedLog "Schreib-Fehler Detail: $_" "ERROR"
                Write-CsvLog -Status "ERR" -Actions "Schreib-Fehler" -Path $filePath -Details $_.Exception.Message
            }

        } else {
            if ($hadZoneIdentifier) {
                # Zentrale Freigabe - MUSS vor der .bak-Behandlung stehen.
                # Dieser Zweig griff bisher ohne jede Freigabe: er ersetzt das
                # Original per Delete + Copy ueber [System.IO.File]-Methoden,
                # die kein ShouldProcess kennen. Damit war die im Skriptkopf
                # zugesicherte Eigenschaft ('-WhatIf zeigt an, welche Dateien
                # geaendert wuerden, ohne sie anzufassen') fuer ihn nicht
                # erfuellt: bei einem Probelauf wurde jede ungeschuetzte .pptx
                # mit Zone.Identifier - auf Ablagen sehr haeufig, weil einmal
                # aus Mail oder Internet geladen - tatsaechlich geloescht und
                # durch die Temp-Kopie ersetzt. Auch das Beiseitelegen eines
                # vorhandenen .bak gehoert hinter die Freigabe, sonst benennt
                # schon die Simulation Dateien um.
                $zoneAktion = if ($istReadOnly) { 'Zone.Identifier, Schreibschutz-Attribut' } else { 'Zone.Identifier' }
                if (-not (Confirm-Write $filePath "$zoneAktion entfernen")) {
                    $stats.WouldChange++
                    if (-not $script:UseProgress) {
                        Write-Host "-> WHATIF ($zoneAktion)" -ForegroundColor DarkCyan
                    }
                    Write-Log    "Simulation: wuerde $zoneAktion entfernen: $filePath" "INFO"
                    Write-CsvLog -Status "WHATIF" -Actions "Wuerde $zoneAktion entfernen" -Path $filePath
                    return
                }

                # Eigener, eindeutiger Backup-Name (siehe Rueckschreib-Pfad).
                $zoneBak = New-BackupPfad -Original $srcLong

                # Schreibschutz erst nach der Freigabe entfernen - Delete am
                # Original scheitert sonst, und das Backup erbte ihn.
                if ($istReadOnly) {
                    $roEntfernt = Clear-ReadOnlyAttribut $srcLong
                    if ($roEntfernt) { Write-DetailedLog "Datei-Schreibschutz (Attribut) entfernt: $filePath" "DEBUG" }
                }

                try {
                    Invoke-WithRetry -Action { [System.IO.File]::Copy($srcLong, $zoneBak, $true) }
                    try { [System.IO.File]::SetCreationTime($zoneBak, (Get-Date)) } catch {}
                    [System.IO.File]::Delete($srcLong)
                    Invoke-WithRetry -Action { [System.IO.File]::Copy($workFile, $srcLong, $true) }
                    if (-not [System.IO.File]::Exists($srcLong)) { throw "Verifizierung fehlgeschlagen" }
                    try {
                        [System.IO.File]::SetCreationTimeUtc($srcLong,   $origCreationTimeUtc)
                        [System.IO.File]::SetLastWriteTimeUtc($srcLong,  $origLastWriteTimeUtc)
                        [System.IO.File]::SetLastAccessTimeUtc($srcLong, $origLastAccessTimeUtc)
                    } catch {
                        Write-DetailedLog "Zeitstempel-Wiederherstellung fehlgeschlagen (Zone.Identifier-Pfad): $filePath - $_" "WARN"
                    }
                    # ACL/Owner zurueckschreiben - die Datei wurde via
                    # Delete+Copy neu erzeugt.
                    Set-FileSecuritySnapshot -Path $srcLong -Snapshot $origSecurity
                    try { [System.IO.File]::Delete($zoneBak) } catch {}
                    $stats.Unlocked++
                    $zoneErg = if ($roEntfernt) { 'Zone.Identifier, Schreibschutz-Attribut' } else { 'Zone.Identifier' }
                    if (-not $script:UseProgress) { Write-Host "-> OK ($zoneErg)" -ForegroundColor Green }
                    Write-Log "Erledigt: $filePath  [$zoneErg]" "INFO"
                    Write-CsvLog -Status "OK" -Actions "$zoneErg entfernt" -Path $filePath
                } catch {
                    $stats.Errors++
                    # Potenziell korrupte Zieldatei zuerst bedingungslos
                    # entfernen: Bei einem AV-Scanner-Eingriff, Netzwerkhaenger
                    # oder Disk-Full mitten im Copy existiert $srcLong typisch
                    # als unvollstaendige/0-Byte-Datei. Ohne dieses Loeschen
                    # wuerde der nachfolgende Rollback ueberspringt und das
                    # intakte Backup wuerde am Ende verworfen werden -
                    # permanenter Datenverlust (Original weg, Backup geloescht).
                    if ([System.IO.File]::Exists($srcLong)) {
                        try { [System.IO.File]::Delete($srcLong) } catch {}
                        Write-DetailedLog "Unvollstaendige Zieldatei entfernt (Zone.Identifier-Pfad): $srcLong" "WARN"
                    }
                    if ([System.IO.File]::Exists($zoneBak)) {
                        $backupSafeToDelete = $false
                        try {
                            Invoke-WithRetry -Action { [System.IO.File]::Copy($zoneBak, $srcLong, $true) }
                            if ([System.IO.File]::Exists($srcLong)) {
                                $backupSafeToDelete = $true
                                Set-FileSecuritySnapshot -Path $srcLong -Snapshot $origSecurity
                                # Original unveraendert zurueck -> Attribut auch.
                                if ($roEntfernt) { Set-ReadOnlyAttribut $srcLong }
                                Write-DetailedLog "Rollback erfolgreich (Zone.Identifier-Pfad): $srcLong" "WARN"
                            }
                        } catch {
                            Write-Log "KRITISCH: Rollback fehlgeschlagen - Backup bleibt erhalten: $zoneBak" "ERROR"
                            Write-DetailedLog "Rollback-Kopie fehlgeschlagen (Zone.Identifier-Pfad): $_" "ERROR"
                        }
                        if ($backupSafeToDelete) {
                            try { [System.IO.File]::Delete($zoneBak) } catch {}
                        }
                    }
                    Write-DetailedLog "Rueckkopie (Zone.Identifier-Strip) fehlgeschlagen: $filePath - $_" "WARN"
                    Write-Log "Schreib-Fehler (Zone.Identifier): $filePath - $_" "ERROR"
                    Write-CsvLog -Status "ERR" -Actions "Schreib-Fehler (Zone.Identifier)" -Path $filePath -Details $_.Exception.Message
                    if (-not $script:UseProgress) { Write-Host "-> WARN (ADS)" -ForegroundColor Yellow }
                }
            } elseif ($istReadOnly) {
                # Kein Schutz im Inhalt, aber Schreibschutz-Attribut: das ist
                # die einzige Aenderung - mit Freigabe und Ausweis.
                if (-not (Confirm-Write $filePath 'Schreibschutz-Attribut entfernen')) {
                    $stats.WouldChange++
                    if (-not $script:UseProgress) { Write-Host "-> WHATIF (Schreibschutz-Attribut)" -ForegroundColor DarkCyan }
                    Write-Log    "Simulation: wuerde Schreibschutz-Attribut entfernen: $filePath" "INFO"
                    Write-CsvLog -Status "WHATIF" -Actions "Wuerde Schreibschutz-Attribut entfernen" -Path $filePath
                } elseif (Clear-ReadOnlyAttribut $srcLong) {
                    $roEntfernt = $true
                    $stats.Unlocked++
                    if (-not $script:UseProgress) { Write-Host "-> OK (Schreibschutz-Attribut)" -ForegroundColor Green }
                    Write-Log    "Erledigt: $filePath  [Schreibschutz-Attribut]" "INFO"
                    Write-CsvLog -Status "OK" -Actions "Schreibschutz-Attribut entfernt" -Path $filePath
                } else {
                    $stats.Errors++
                    if (-not $script:UseProgress) { Write-Host "-> ERR (Schreibschutz-Attribut)" -ForegroundColor Red }
                    Write-Log    "Schreibschutz-Attribut nicht entfernbar: $filePath" "ERROR"
                    Write-CsvLog -Status "ERR" -Actions "Schreibschutz-Attribut nicht entfernbar" -Path $filePath
                }
            } else {
                if (-not $script:UseProgress) { Write-Host "-> Kein Schutz" -ForegroundColor Gray }
                Write-CsvLog -Status "NO_CHANGE" -Actions "Kein Schutz gefunden" -Path $filePath
            }
        }

    } catch {
        $stats.Errors++
        if ($backupPath -and [System.IO.File]::Exists($backupPath)) {
            try { [System.IO.File]::Delete($backupPath) } catch {}
        }
        # Hier landet nur, was VOR dem inneren Schreib-try scheitert (z. B.
        # die Backup-Kopie) - das Original ist dann unveraendert.
        if ($roEntfernt -and [System.IO.File]::Exists($srcLong)) { Set-ReadOnlyAttribut $srcLong }
        if (-not $script:UseProgress) { Write-Host "-> ERROR: $($_.Exception.Message)" -ForegroundColor Red }
        Write-Log "Allg. Fehler: $filePath - $_" "ERROR"
        Write-DetailedLog "Allg. Fehler Detail: $_" "ERROR"
        Write-CsvLog -Status "ERR" -Actions "Allg. Fehler" -Path $filePath -Details $_.Exception.Message

    } finally {
        if (-not [string]::IsNullOrWhiteSpace($workFile)) {
            Remove-Item -LiteralPath $workFile -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false
        }
        if (-not [string]::IsNullOrWhiteSpace($tempFile) -and ($tempFile -ne $workFile)) {
            Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false
        }
        # Reservierten Platzhalter nur entfernen, wenn er leer geblieben ist.
        # Bricht z. B. die Backup-Kopie ab, laeuft der aeussere catch an, der
        # $finalDest nicht aufraeumt - die 0-Byte-Datei bliebe sonst liegen.
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
    if ($script:FremdePptAbbruch -and -not $script:ShouldStop) {
        Write-Host "`nVerarbeitung abgebrochen: PowerPoint des Anwenders laeuft." -ForegroundColor Red
        Write-Log "Verarbeitung abgebrochen: PowerPoint des Anwenders laeuft - COM-Arbeit eingestellt." "ERROR"
    } else {
        Write-Host "`nVerarbeitung durch Nutzer abgebrochen." -ForegroundColor Yellow
        Write-Log "Verarbeitung durch Nutzer abgebrochen." "WARN"
    }
}
try { if ($script:CtrlCAsInput) { [Console]::TreatControlCAsInput = $false; $script:CtrlCAsInput = $false } } catch {}

# ==================================================================
# AUFRAEUMEN & STATISTIK
# ==================================================================
if ($script:UseProgress) { Write-Progress -Activity "Fertig" -Completed }

Clear-TrackedPowerPointInstances
if (Test-Path $script:TempPath) { Remove-Item $script:TempPath -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false }

$duration = (Get-Date) - $script:ScriptStartTime
$durStr   = $duration.ToString('hh\:mm\:ss')

Write-Host "`n=====================================================" -ForegroundColor Cyan
Write-Host " FERTIG in $durStr" -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host "Geprueft:        $($stats.Processed)"  -ForegroundColor White
Write-Host "Entsperrt/Konv:  $($stats.Unlocked)"   -ForegroundColor Green
Write-Host "Konvertiert:     $($stats.Converted)"  -ForegroundColor Green
if ($stats.WouldChange -gt 0) {
    Write-Host "Wuerde aendern:  $($stats.WouldChange)  (Simulation - nichts geschrieben)" -ForegroundColor DarkCyan
}
if ($backupCleanup) {
    Write-Host "Backups entf.:   $($backupCleanup.Deleted)" -ForegroundColor DarkGray
    if ($backupCleanup.Orphans -gt 0) {
        Write-Host "Backups o. Orig: $($backupCleanup.Orphans)  (NICHT geloescht - pruefen!)" -ForegroundColor Yellow
    }
}
Write-Host "Junk geloescht:  $($stats.Junk)"       -ForegroundColor DarkGray
if ($stats.JunkKept -gt 0) {
    Write-Host "Sperrdateien:    $($stats.JunkKept)  (juenger als $($script:JunkMinAgeHours) h - geschont)" -ForegroundColor DarkGray
}
Write-Host "Uebersprungen:   $($stats.Skipped)"    -ForegroundColor Magenta
Write-Host "Verschluesselt:  $($stats.Encrypted)"  -ForegroundColor Magenta
Write-Host "Korrupt/kein ZIP:$($stats.Corrupt)"    -ForegroundColor Magenta
Write-Host "COM-Timeouts:    $($stats.Timeouts)"   -ForegroundColor Magenta
Write-Host "AV-blockiert:    $($stats.AvBlocked)"  -ForegroundColor Magenta
Write-Host "Gesperrt:        $($stats.Locked)"     -ForegroundColor Yellow
Write-Host "Fehler:          $($stats.Errors)"     -ForegroundColor Red
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host "Logs:" -ForegroundColor Cyan
Write-Host "  $([System.IO.Path]::GetFullPath($script:LogFilePath))"     -ForegroundColor Gray
Write-Host "  $([System.IO.Path]::GetFullPath($script:DetailedLogPath))" -ForegroundColor Gray
Write-Host "  $([System.IO.Path]::GetFullPath($script:CsvLogPath))"      -ForegroundColor Gray
Write-Host "=====================================================" -ForegroundColor Cyan

$footerLines = @(
    "=================================================================="
    "Log Ende:  $(Get-Date)"
    "Dauer:     $durStr"
    "Geprueft:  $($stats.Processed)  |  Entsperrt: $($stats.Unlocked)  |  Fehler: $($stats.Errors)"
    "Konv:      $($stats.Converted)  |  Skip: $($stats.Skipped)  |  Verschluesselt: $($stats.Encrypted)  |  Korrupt: $($stats.Corrupt)  |  Timeouts: $($stats.Timeouts)  |  AV-blockiert: $($stats.AvBlocked)  |  Gesperrt: $($stats.Locked)"
    "=================================================================="
)
if ($script:LogWriter) {
    foreach ($line in $footerLines) { try { $script:LogWriter.WriteLine($line) } catch {} }
}

Invoke-WindowsTempCleanup
Close-Loggers

if ($script:FremdePptAbbruch) {
    Write-Host ""
    Write-Host "ACHTUNG: Lauf wegen laufender PowerPoint-Sitzung abgebrochen (Exitcode 3)." -ForegroundColor Red
    Write-Host "PowerPoint beenden und das Skript erneut starten." -ForegroundColor Red
}

if (-not $NoInteractive.IsPresent) {
    Write-Host ""
    Write-Host "Beliebige Taste druecken zum Beenden..." -ForegroundColor DarkGray
    try { $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") } catch {}
}

# Exitcode 3 = PowerPoint des Anwenders lief (Abbruch zu seinem Schutz).
if ($script:FremdePptAbbruch) { exit 3 }
