<#
.SYNOPSIS
    Excel Deep Unprotect
.DESCRIPTION
    Entfernt Arbeitsmappen-, Blatt-, Chart- und Bereichsschutz aus Excel-Dateien.
    Konvertiert alte .xls/.xlt/.xla- und binäre .xlsb-Formate nach .xlsx/.xltx/.xlam.
    Unterstützt Netzlaufwerke und Pfade > 260 Zeichen.

    Behandelte Schutzarten (XML-Ebene, ohne Passwort):
      workbook.xml          : workbookProtection, fileSharing
      worksheets/*.xml      : sheetProtection, protectedRanges
      chartsheets/*.xml     : sheetProtection
      macrosheets/*.xml     : sheetProtection
      dialogsheets/*.xml    : sheetProtection
      tables/*.xml          : tableProtection
      pivotTables/*.xml     : pivotTableProtection

    Nicht entfernbar ohne Passwort:
      - Verschlüsselte Arbeitsmappen (Open-Passwort) -> werden übersprungen
      - VBA-Projekt-Passwort in .xlsm/.xlam (binäre vbaProject.bin)

    Voraussetzung Vertrauenswürdige Speicherorte:
      Das Skript verwendet den Ordner "Dokumente\2b_entferne_schutz_excel_<PID>" als
      temporären Arbeitsordner für die COM-Automatisierung. Damit Excel Dateien
      dort ohne Geschützte Ansicht öffnet, muss der Dokumentenordner des
      ausführenden Benutzers als vertrauenswürdiger Speicherort eingetragen sein:
        Excel -> Datei -> Optionen -> Trust Center -> Einstellungen...
          -> Vertrauenswürdige Speicherorte
          -> "Neuen Speicherort hinzufügen..."
          -> Pfad: C:\Users\<Benutzername>\Documents
          -> Haken: "Unterordner dieses Speicherorts sind ebenfalls vertrauenswürdig"
      Bei Netzlaufwerken als Zielpfad zusätzlich den Haken setzen bei:
          -> "Vertrauenswürdige Speicherorte im Netzwerk zulassen"
      Das Skript führt vor dem Hauptlauf einen Trust-Center-Smoke-Test durch und
      bricht bei Fehlschlag ab (bzw. fragt interaktiv nach), um stundenlange
      Timeouts pro Datei zu vermeiden.

    Voraussetzung für Aufgabenplanung / Headless-Betrieb (Dienstkonto):
      Excel benötigt einen existierenden Desktop-Ordner im Profilpfad des
      ausführenden Kontos. Fehlen diese Ordner, schlägt die COM-Initialisierung
      mit Fehler 0x800A03EC oder "Out of Memory" fehl. Vor dem ersten Einsatz
      als Administrator auf dem Zielserver anlegen:
        C:\Windows\SysWOW64\config\systemprofile\Desktop
        C:\Windows\System32\config\systemprofile\Desktop

    Bekanntes Infrastruktur-Verhalten (kein Fehler im Skript):
      Bei DFS-Umleitungen auf offline/abgestürzte Server kann die Dateisuche
      an einzelnen Ordnern bis zu 60 Sekunden einfrieren, bis Windows den
      Netzwerk-Timeout auslöst und die IOException abgefangen wird. Das Skript
      setzt danach automatisch fort. Bei Hunderten toter Links kann die
      Gesamtlaufzeit erheblich steigen und die ETA-Anzeige stagnieren.

    Log-Dateien im Skriptverzeichnis (mit Zeitstempel pro Lauf):
      2b_entferne_schutz_excel_<TS>.log           : Haupt-Log (UTF-8 mit BOM)
      2b_entferne_schutz_excel_detailed_<TS>.log  : Detail-Log mit DEBUG-Eintraegen (UTF-8 mit BOM)
      2b_entferne_schutz_excel_<TS>.csv           : Ergebnis-CSV pro Datei (UTF-8 mit BOM, Semikolon)

    Retry-Verhalten bei transienten Dateisperren (Virenscanner):
      Frisch in den Temp-Ordner geschriebene Dateien werden von Defender oder
      anderen AV-Engines oft noch exklusiv gehalten, wenn ZipFile.Open(Update)
      oder File.Copy darauf zugreifen will. Das Skript wartet nach jeder
      Schreiboperation aktiv mit Wait-FileAvailable (FileShare.None-Probe) auf
      Freigabe und wrapped zusaetzlich kritische I/O-Aufrufe mit einem
      Retry-Wrapper (bis zu 5 Versuche, exponentielles Backoff
      200 ms -> 2 s). Nur Sharing-Violations und IOException werden retried;
      echte ZIP-/Format-Fehler (verschluesselt, korrupt) führen direkt zum
      Standard-Skip-Pfad. Die Retry-Statistik erscheint am Ende des Laufs,
      sofern mindestens ein Retry stattgefunden hat.

    Stat-Hinweis:
      "Gesperrt" ist eine Teilmenge von "Uebersprungen". Eine gesperrte Datei
      erhoeht beide Zaehler.

    Compilieren in eine .exe (Windows PowerShell 5.1, STA-Konsolenanwendung):
      # Einmalig: ps2exe installieren
      Install-Module -Name ps2exe -Scope CurrentUser -Force

      # Hilfscheck: Office-Bitness ermitteln (optional zur Optimierung)
      #   Rueckgabe 'x64' -> -x64 verwenden ; 'x86' -> -x86 verwenden
      Get-ItemProperty 'HKLM:\Software\Microsoft\Office\ClickToRun\Configuration' -Name Platform -ErrorAction SilentlyContinue

      # Hinweis zu -x64 / -x86 bei gemischten Office-Bitness-Umgebungen:
      # Excel registriert seinen COM-Server als LocalServer32 (out-of-process,
      # EXCEL.EXE). Das Windows-COM-Subsystem marshallt Aufrufe zwischen
      # x64-Aufrufer und x86-Server (und umgekehrt) automatisch ueber DCOM/LRPC.
      # Daher funktioniert eine mit -x64 kompilierte EXE auch mit 32-Bit-Excel
      # (und eine -x86-EXE mit 64-Bit-Excel). Bitness-Match liefert minimal
      # bessere Performance, ist aber nicht erforderlich. -x64 ist eine sichere
      # Default-Wahl, da modernes Office (2019+, M365) standardmaessig x64 ist.

      Invoke-ps2exe -inputFile   '.\2b_entferne_schutz_excel.ps1' -outputFile  '.\2b_entferne_schutz_excel.exe' -iconFile    '.\powershell_icon.ico' -title       'Excel Deep Unprotect' -description 'Entfernt Arbeitsmappen-, Blatt-, Bereichs- und Pivot-Schutz aus Excel-Dateien' -product     'Excel Deep Unprotect' -version     '1.0.1.0' -STA -x64 -supportOS

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

$scriptDir = Resolve-ScriptDirectory

# Log-Verzeichnis: bevorzugt neben dem Skript, sonst LOCALAPPDATA, sonst TEMP.
# Vorher wurde ungeprueft in $scriptDir geschrieben. Liegt das Skript auf einer
# schreibgeschuetzten Freigabe, scheiterte das Oeffnen der drei Log-Writer und
# der gesamte Lauf lief ohne Protokoll und ohne CSV durch.
function Resolve-LogDirectory {
    param([string]$Preferred)
    foreach ($cand in @($Preferred,
                        (Join-Path $env:LOCALAPPDATA 'ExcelDeepUnprotect'),
                        $env:TEMP)) {
        if ([string]::IsNullOrWhiteSpace($cand)) { continue }
        try {
            if (-not [System.IO.Directory]::Exists($cand)) {
                New-Item -ItemType Directory -Path $cand -Force -ErrorAction Stop | Out-Null
            }
            $probe = Join-Path $cand (".writetest_{0}" -f ([Guid]::NewGuid().ToString('N')))
            [System.IO.File]::WriteAllText($probe, 'x')
            [System.IO.File]::Delete($probe)
            return $cand
        } catch { continue }
    }
    return $Preferred
}
$LogDir = Resolve-LogDirectory -Preferred $scriptDir

$docRoot = [Environment]::GetFolderPath("MyDocuments")
if ([string]::IsNullOrWhiteSpace($docRoot) -or -not (Test-Path -LiteralPath $docRoot)) {
    $docRoot = $env:TEMP
    if ([string]::IsNullOrWhiteSpace($docRoot)) { $docRoot = "C:\Windows\Temp" }
}
$TempPath               = Join-Path $docRoot "2b_entferne_schutz_excel_$PID"

$RunTimestamp           = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$LogFilePath            = Join-Path $LogDir "2b_entferne_schutz_excel_$RunTimestamp.log"
$DetailedLogPath        = Join-Path $LogDir "2b_entferne_schutz_excel_detailed_$RunTimestamp.log"
$CsvLogPath             = Join-Path $LogDir "2b_entferne_schutz_excel_$RunTimestamp.csv"
$FileOpenTimeoutSeconds = 45
$TrustCenterTimeoutSec  = 25

# Mindestalter fuer das Loeschen von '~$'-Sperrdateien. Excel legt sie an,
# solange eine Mappe GEOEFFNET ist - sofortiges Loeschen bricht die Sperre
# fuer andere Nutzer auf der Freigabe.
$JunkMinAgeHours          = 24

# Maximale Anzahl Ausweichnamen (Datei_2.xlsx ... Datei_100.xlsx), bevor
# zum Schutz fremder Dateien abgebrochen wird.
$MaxNameClashRetries      = 100

# Mindestalter, ab dem ein zurueckgebliebenes '.bak' als verwaist gilt.
$BackupCleanupMinAgeHours = 24

# Verzeichnisse, die bei der Suche nicht betreten werden. Ohne diese Liste
# landeten auch Mappen im Papierkorb und in Schattenkopien in der Verarbeitung.
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

# ------------------------------------------------------------------
# Retry-Konfiguration (Virenscanner-Locks auf frisch geschriebenen Temp-Dateien)
# ------------------------------------------------------------------
$RetryMaxAttempts       = 5
$RetryInitialDelayMs    = 200
$RetryMaxDelayMs        = 2000
$WaitFileReadyTimeoutSec = 10
$PidCleanupInterval     = 50

$script:Passwords       = [System.Collections.Generic.List[string]]::new()
$script:RetryRecovered  = 0
$script:RetryFailed     = 0
$script:RestorePrivilegesEnabled = $false

$Utf8Bom   = New-Object System.Text.UTF8Encoding($true)
$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)

$ScriptStartTime        = Get-Date

# ==================================================================
# INIT & LIBRARIES
# ==================================================================
try {
    Add-Type -AssemblyName "System.IO.Compression.FileSystem" -ErrorAction Stop
} catch {
    [System.Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") | Out-Null
}

$script:ShouldStop          = $false
$script:UseProgress         = $false
$script:SkipPreScan         = $false
$script:TotalFiles          = 0
$script:ProcessedCount      = 0
$script:TrackedExcelPids    = [System.Collections.Generic.List[int]]::new()
$script:LogWriter           = $null
$script:DetailedLogWriter   = $null
$script:CsvLogWriter        = $null
$script:LastProgressUpdate    = [DateTime]::MinValue
$script:ProgressMinIntervalMs = 500

# $PSCmdlet ist nur im Skript-Scope verfuegbar, nicht in einfachen Funktionen.
$script:ScriptCmdlet = $PSCmdlet

function Confirm-Write {
    <#
        Zentrale Freigabe fuer jede schreibende Operation am Original.

        Das Skript deklarierte SupportsShouldProcess, rief ShouldProcess aber an
        keiner Stelle auf: -WhatIf wurde stillschweigend angenommen und die
        Dateien trotzdem geaendert.
    #>
    param(
        [string]$Target,
        [string]$Action = 'Aendern'
    )
    if ($null -eq $script:ScriptCmdlet) { return $true }
    try {
        return $script:ScriptCmdlet.ShouldProcess($Target, $Action)
    } catch {
        return $true
    }
}

function Test-IsOwnExcelProcess {
    <#
        Prueft, ob eine PID wirklich zu einer von UNS gestarteten Excel-Instanz
        gehoert. Windows vergibt PIDs wieder; ohne Startzeit-Vergleich koennte
        Stop-AllTrackedExcel eine zwischenzeitlich vom Nutzer geoeffnete
        Excel-Sitzung abschiessen, die dieselbe PID bekommen hat.
    #>
    param([int]$ProcessId)
    try {
        $proc = Get-Process -Id $ProcessId -ErrorAction Stop
        if ($proc.Name -ne 'EXCEL') { return $false }
        try {
            if ($proc.StartTime -lt $ScriptStartTime) { return $false }
        } catch {
            return $false
        }
        return $true
    } catch {
        return $false
    }
}

# ==================================================================
# HELPER-FUNKTIONEN
# ==================================================================

function Add-LongPathPrefix {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
    if ($Path -match "^\\\\\?\\")            { return $Path }
    if ($Path -match "^\\\\")                { return "\\?\UNC" + $Path.Substring(1) }
    if ($Path -match "^[a-zA-Z]:")           { return "\\?\$Path" }
    return $Path
}

# ------------------------------------------------------------------
# User Shell Folder Aufloesung (Registry mit Fallback)
# ------------------------------------------------------------------
function Get-UserShellFolder {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Desktop','Downloads')]
        [string]$Name
    )

    $regKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    $regValueName = if ($Name -eq 'Downloads') {
        '{374DE290-123F-4565-9164-39C4925E467B}'
    } else {
        'Desktop'
    }

    try {
        $rawProp = Get-ItemProperty -Path $regKey -Name $regValueName -ErrorAction Stop
        $raw = $rawProp.$regValueName
        if (-not [string]::IsNullOrWhiteSpace($raw)) {
            $expanded = [Environment]::ExpandEnvironmentVariables($raw)
            if (Test-Path -LiteralPath $expanded) {
                return $expanded
            }
        }
    } catch {}

    if ($Name -eq 'Desktop') {
        $fb = [Environment]::GetFolderPath('Desktop')
        if (-not [string]::IsNullOrWhiteSpace($fb) -and (Test-Path -LiteralPath $fb)) {
            return $fb
        }
        return (Join-Path $env:USERPROFILE 'Desktop')
    }

    return (Join-Path $env:USERPROFILE 'Downloads')
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
            if ($attempt -gt 1) { $script:RetryRecovered++ }
            return $true
        } catch [System.IO.IOException] {
            if ($attempt -ge $RetryMaxAttempts) {
                $script:RetryFailed++
                Write-DetailedLog "Test-IsValidZip: $RetryMaxAttempts Versuche erfolglos (IOException): $FilePath - $($_.Exception.Message)" "WARN"
                return $false
            }
            Write-DetailedLog "Test-IsValidZip IOException (Versuch $attempt/$RetryMaxAttempts, ${delay}ms): $FilePath" "DEBUG"
            Start-Sleep -Milliseconds $delay
            $delay = [Math]::Min($delay * 2, $RetryMaxDelayMs)
        } catch {
            return $false
        }
    }
}

# ------------------------------------------------------------------
# Retry-Wrapper für I/O gegen Virenscanner- und transiente Dateisperren
# ------------------------------------------------------------------
function Invoke-WithRetry {
    param(
        [Parameter(Mandatory)] [scriptblock]$ScriptBlock,
        [string]$OperationName = "I/O",
        [int]$MaxAttempts      = $RetryMaxAttempts,
        [int]$InitialDelayMs   = $RetryInitialDelayMs,
        [int]$MaxDelayMs       = $RetryMaxDelayMs
    )

    $attempt = 0
    $delay   = $InitialDelayMs
    $lastErr = $null

    while ($attempt -lt $MaxAttempts) {
        $attempt++
        try {
            $result = & $ScriptBlock
            if ($attempt -gt 1) {
                $script:RetryRecovered++
                Write-DetailedLog "Retry erfolgreich nach Versuch ${attempt} für ${OperationName}" "DEBUG"
            }
            return $result
        } catch {
            $lastErr = $_
            $msg     = $_.Exception.Message
            $isTransient = (
                $msg -match "verwendet wird"                      -or
                $msg -match "being used by another process"       -or
                $msg -match "wird von einem anderen Prozess"      -or
                $msg -match "sharing violation"                    -or
                $msg -match "cannot access the file"               -or
                $msg -match "process cannot access"                -or
                $_.Exception -is [System.IO.IOException]
            )
            if (-not $isTransient -or $attempt -ge $MaxAttempts) {
                if ($attempt -gt 1) { $script:RetryFailed++ }
                throw
            }
            Write-DetailedLog "Retry ${OperationName} (Versuch ${attempt}/${MaxAttempts}, ${delay}ms): $msg" "DEBUG"
            Start-Sleep -Milliseconds $delay
            $delay = [Math]::Min($delay * 2, $MaxDelayMs)
        }
    }
    if ($lastErr) { throw $lastErr }
}

# ------------------------------------------------------------------
# Aktive Wartephase auf Datei-Freigabe (AV-Scanner / OS-Cache)
# ------------------------------------------------------------------
function Wait-FileAvailable {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [int]$TimeoutSec     = $WaitFileReadyTimeoutSec,
        [int]$InitialDelayMs = 80,
        [int]$MaxDelayMs     = 800
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    $delay    = $InitialDelayMs
    $attempts = 0

    while ((Get-Date) -lt $deadline) {
        $attempts++
        if (-not [System.IO.File]::Exists($Path)) {
            Start-Sleep -Milliseconds $delay
            $delay = [Math]::Min($delay * 2, $MaxDelayMs)
            continue
        }
        try {
            $fs = [System.IO.File]::Open(
                $Path,
                [System.IO.FileMode]::Open,
                [System.IO.FileAccess]::ReadWrite,
                [System.IO.FileShare]::None
            )
            $fs.Close()
            $fs.Dispose()
            if ($attempts -gt 1) {
                $script:RetryRecovered++
                Write-DetailedLog "Wait-FileAvailable erfolgreich nach $attempts Versuchen: $Path" "DEBUG"
            }
            return $true
        } catch [System.UnauthorizedAccessException] {
            return $false
        } catch {
            Start-Sleep -Milliseconds $delay
            $delay = [Math]::Min($delay * 2, $MaxDelayMs)
        }
    }

    $script:RetryFailed++
    Write-DetailedLog "Wait-FileAvailable Timeout nach ${TimeoutSec}s ($attempts Versuche): $Path" "WARN"
    return $false
}

# ------------------------------------------------------------------
# OLE-Verschlüsselung erkennen (.xls, .xlt, .xla)
# ------------------------------------------------------------------
function Test-OleFileIsEncrypted {
    param([string]$FilePath)

    # Nur die ersten 32 KB per Stream lesen statt die komplette Datei via
    # ReadAllBytes in den RAM zu laden: alte .xls mit eingebetteten Grafiken
    # erreichen nicht selten 100+ MB - das sparte unnoetig RAM und riskierte
    # OutOfMemoryException, obwohl ohnehin nur die ersten 32 KB durchsucht
    # werden.
    $stream = $null
    try {
        $stream = [System.IO.File]::OpenRead($FilePath)
        $buffer = New-Object byte[] 32768
        $bytesRead = $stream.Read($buffer, 0, 32768)

        if ($bytesRead -lt 512) { return $false }
        if ($buffer[0] -ne 0xD0 -or $buffer[1] -ne 0xCF -or $buffer[2] -ne 0x11 -or $buffer[3] -ne 0xE0) {
            return $false
        }

        $contentU16 = [System.Text.Encoding]::Unicode.GetString($buffer, 0, $bytesRead)
        $contentAsc = [System.Text.Encoding]::ASCII.GetString($buffer, 0, $bytesRead)

        $pattern = "EncryptedPackage|EncryptionInfo|EncryptedSummary|StrongEncryptionDataSpace|DataSpaceMap"
        if ($contentU16 -match $pattern) { return $true }
        if ($contentAsc -match $pattern) { return $true }

        return $false
    } catch {
        return $false
    } finally {
        if ($null -ne $stream) {
            try { $stream.Dispose() } catch {}
        }
    }
}

function Get-ExcelFilesRobust {
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
                if ($script:ExcludeDirNames -contains $leaf) { continue }
                $queue.Enqueue($sub)
            }
        } catch {}
    }
}

function Get-OriginalFromBackupPath {
    # Vertrag.xlsx.bak / Vertrag.xlsx.bak_a1b2c3  ->  Vertrag.xlsx
    param([string]$BackupPath)
    if ($BackupPath -match '^(?<orig>.+?)\.bak(?:_[0-9a-fA-F]+)?$') {
        return $Matches['orig']
    }
    return $null
}

function Remove-OrphanedBackups {
    <#
        Raeumt '.bak'/'.bak_xxxxxx'-Reste frueherer Laeufe auf.

        Geloescht wird nur, wenn ALLE Bedingungen zutreffen:
          1. Der abgeleitete Originalname endet auf eine Excel-Endung
             (fremde .bak-Dateien bleiben unangetastet),
          2. das Original existiert und ist groesser als 0 Byte,
          3. das Backup ist aelter als $BackupCleanupMinAgeHours.

        Fehlt das Original oder ist es leer, kann das Backup die einzige
        intakte Kopie sein - solche Funde werden gemeldet, nie geloescht.
    #>
    param([string]$RootPath)

    $res = @{
        Deleted = 0; Kept = 0; Orphans = 0; Failed = 0
        OrphanList = [System.Collections.Generic.List[string]]::new()
    }
    $cutoff = (Get-Date).AddHours(-$BackupCleanupMinAgeHours)

    foreach ($bak in (Get-ExcelFilesRobust $RootPath '^\.bak(_[0-9a-fA-F]+)?$' -StringsOnly)) {
        try {
            $orig = Get-OriginalFromBackupPath -BackupPath $bak
            if ([string]::IsNullOrWhiteSpace($orig)) { continue }
            if ($orig -notmatch '\.(xls|xlt|xla|xlsx|xlsm|xltx|xltm|xlsb|xlam)$') {
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
            # 'Length -gt 0' allein genuegt nicht: ein Abbruch mitten im
            # 'Copy work->final' hinterlaesst ein TEILWEISE geschriebenes
            # Original mit Groesse > 0 - genau der Fall, den der Kommentar an
            # der Backup-Erzeugung als Grund nennt, ein altes .bak keinesfalls
            # zu ueberschreiben. Ohne diese Pruefung wurde beim naechsten Lauf
            # die einzige unversehrte Kopie geloescht.
            if ($origLen -lt $fi.Length) {
                $res.Orphans++
                $res.OrphanList.Add($bak)
                Write-Log ("Backup groesser als das Original ({0:N0} statt {1:N0} Bytes) - Original " +
                           "moeglicherweise abgeschnitten. NICHT geloescht, bitte pruefen: {2}" -f `
                           $fi.Length, $origLen, $bak) "WARN"
                continue
            }

            # Altersfrist an der CreationTime messen, NICHT an der
            # LastWriteTime. Das Backup entsteht per [System.IO.File]::Copy,
            # und Copy uebertraegt die LastWriteTime der QUELLE auf das Ziel -
            # nur die CreationTime wird neu gesetzt. Auf Archiv-Ablagen ist
            # die LastWriteTime der Dokumente typischerweise Jahre alt, das
            # Backup erbt sie. Damit galt ein Sekunden altes .bak sofort als
            # verwaist (nachgestellt: Original von 2020, Backup gerade erzeugt
            # -> 'wird geloescht'), die Frist war wirkungslos und der
            # Kept-Zweig praktisch toter Code. Mit der CreationTime bleibt
            # auch das .bak eines parallel laufenden Durchgangs verschont.
            if ($fi.CreationTime -gt $cutoff) {
                $res.Kept++
                continue
            }

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

# ------------------------------------------------------------------
# Log-Writer (persistent; Fallback auf offen-schliessen pro Aufruf)
# ------------------------------------------------------------------
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    if ($script:LogWriter) {
        try { $script:LogWriter.WriteLine($entry) } catch {}
    } else {
        $sw = $null
        try {
            $sw = [System.IO.StreamWriter]::new($LogFilePath, $true, $Utf8Bom)
            $sw.WriteLine($entry)
        } catch {} finally { if ($sw) { $sw.Close() } }
    }
}

function Write-DetailedLog {
    param([string]$Message, [string]$Level = "DEBUG")
    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff')] [$Level] $Message"
    if ($script:DetailedLogWriter) {
        try { $script:DetailedLogWriter.WriteLine($entry) } catch {}
    } else {
        $sw = $null
        try {
            $sw = [System.IO.StreamWriter]::new($DetailedLogPath, $true, $Utf8Bom)
            $sw.WriteLine($entry)
        } catch {} finally { if ($sw) { $sw.Close() } }
    }
}

function Write-CsvLog {
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
            Write-Laufprotokoll -Skript '2b_entferne_schutz_excel' `
                -Pfad $Path -Aktion 'Schutz entfernen' `
                -Status $Status -Detail $Details
        } catch { }
    }
    if (-not $script:CsvLogWriter) { return }
    $ts      = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $escAct  = '"' + (($Actions  -replace '"', '""') -replace "`r`n|`r|`n", ' ') + '"'
    $escPath = '"' + (($Path     -replace '"', '""') -replace "`r`n|`r|`n", ' ') + '"'
    $escDet  = '"' + (($Details  -replace '"', '""') -replace "`r`n|`r|`n", ' ') + '"'
    try { $script:CsvLogWriter.WriteLine("$ts;$Status;$escAct;$escPath;$escDet") } catch {}
}

function Sync-LogWriters {
    try { if ($script:LogWriter)         { $script:LogWriter.Flush() } }         catch {}
    try { if ($script:DetailedLogWriter) { $script:DetailedLogWriter.Flush() } } catch {}
    try { if ($script:CsvLogWriter)      { $script:CsvLogWriter.Flush() } }      catch {}
}

function Close-LogWriters {
    try { if ($script:LogWriter)         { $script:LogWriter.Close() } }         catch {}
    try { if ($script:DetailedLogWriter) { $script:DetailedLogWriter.Close() } } catch {}
    try { if ($script:CsvLogWriter)      { $script:CsvLogWriter.Close() } }      catch {}
    $script:LogWriter         = $null
    $script:DetailedLogWriter = $null
    $script:CsvLogWriter      = $null
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

    $prefixes = @('~$', '~df', 'gen_py', 'vbe', 'excel8.0', '2b_entferne_schutz_excel')
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
    # Verwaiste Arbeitsordner abgebrochener Laeufe entfernen - nur eigene
    # 2b_entferne_schutz_excel_<PID>-Ordner, deren PID nicht mehr lebt.
    $parent = [Environment]::GetFolderPath('MyDocuments')
    if ([string]::IsNullOrWhiteSpace($parent)) { return }
    if (-not (Test-Path -LiteralPath $parent)) { return }
    $candidates = @()
    try {
        $candidates = @(Get-ChildItem -LiteralPath $parent -Directory `
                                      -Filter "2b_entferne_schutz_excel_*" `
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

# ------------------------------------------------------------------
# Lock-Status einer Datei (einheitlicher String-Rueckgabewert)
# ------------------------------------------------------------------
function Test-FileIsLocked {
    param([string]$FilePath)
    $lp = Add-LongPathPrefix $FilePath
    try {
        if (-not [System.IO.File]::Exists($lp)) { return 'Free' }
        $stream = [System.IO.File]::Open(
            $lp,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::ReadWrite,
            [System.IO.FileShare]::None
        )
        $stream.Close()
        return 'Free'
    } catch [System.UnauthorizedAccessException] {
        return 'AccessDenied'
    } catch {
        return 'Locked'
    }
}

function Stop-AllTrackedExcel {
    foreach ($id in @($script:TrackedExcelPids | Select-Object -Unique)) {
        if (Test-IsOwnExcelProcess -ProcessId $id) {
            try { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue } catch {}
        }
    }
    [System.GC]::Collect()
    $script:TrackedExcelPids.Clear()
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

    $pct    = 0
    $etaStr = if ($script:SkipPreScan) { "ETA: --:--:--" } else { "Berechne..." }
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
    $activity = "Excel Deep Unprotect  |  $etaStr  |  $($script:ProcessedCount) / $totalStr"
    Write-Progress -Activity $activity -Status $Status -CurrentOperation $CurrentFile -PercentComplete ([Math]::Min($pct, 100))
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

function Get-ConversionErrorCategory {
    <#
        Klassifiziert eine COM-Fehlermeldung.

        Zwei Korrekturen gegenueber der frueheren Inline-Fassung:
          - "Datentraeger voll" (mit ae) konnte nie greifen; Windows meldet
            auf Deutsch "Datentraeger" mit Umlaut, auf Englisch "not enough
            space on the disk". Der Zweig war damit toter Code.
          - Die Datentraeger-Pruefung steht jetzt VOR der Speicherpruefung.
            "Auf dem Datentraeger ist nicht genuegend Speicherplatz" wurde
            sonst vom Muster 'nicht gen.gend' als Arbeitsspeicher-Problem
            fehlklassifiziert.
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

function Move-PasswordToFront {
    param([string]$Password)
    if ([string]::IsNullOrEmpty($Password))         { return }
    if (-not $script:Passwords.Contains($Password)) { return }
    $idx = $script:Passwords.IndexOf($Password)
    if ($idx -le 0) { return }
    $script:Passwords.RemoveAt($idx)
    $script:Passwords.Insert(0, $Password)
}

# ------------------------------------------------------------------
# Minimale gueltige XLSX fuer Trust-Center-Smoke-Test
# ------------------------------------------------------------------
function New-MinimalXlsx {
    param([string]$Path)

    if ([System.IO.File]::Exists($Path)) {
        [System.IO.File]::Delete($Path)
    }

    $contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
    $rels         = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    $wbRels       = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
    $workbook     = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'
    $sheet        = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>'

    $entries = @(
        @{ Name = "[Content_Types].xml";          Content = $contentTypes }
        @{ Name = "_rels/.rels";                  Content = $rels         }
        @{ Name = "xl/_rels/workbook.xml.rels";   Content = $wbRels       }
        @{ Name = "xl/workbook.xml";              Content = $workbook     }
        @{ Name = "xl/worksheets/sheet1.xml";     Content = $sheet        }
    )

    $zip = [System.IO.Compression.ZipFile]::Open($Path, 'Create')
    try {
        foreach ($e in $entries) {
            $entry = $zip.CreateEntry($e.Name)
            $s     = $entry.Open()
            $bytes = $Utf8NoBom.GetBytes($e.Content)
            $s.Write($bytes, 0, $bytes.Length)
            $s.Close()
        }
    } finally {
        $zip.Dispose()
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
        if (-not ('ExcelUnprotect.PrivilegeHelper' -as [type])) {
            Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace ExcelUnprotect {
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
        $okRestore = [ExcelUnprotect.PrivilegeHelper]::Enable('SeRestorePrivilege')
        $okOwner   = [ExcelUnprotect.PrivilegeHelper]::Enable('SeTakeOwnershipPrivilege')
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
# LAUFENDE OFFICE-SITZUNGEN
# ==================================================================
# Diese beiden Funktionen fehlten hier als einzigem der COM-Skripte
# (2a, 2c, 7 und 9 haben sie). Fuer Excel gilt die Begruendung
# unveraendert: 'New-Object -ComObject Excel.Application' startet keine
# neue Instanz, wenn Excel bereits laeuft, sondern haengt sich an die
# vorhandene Sitzung. Visible=$false laesst sie fuer den Anwender wie
# abgestuerzt aussehen, DisplayAlerts=$false schaltet seine Warnhinweise
# ab, und die eigene Instanz ist danach nicht mehr zuverlaessig von der
# fremden zu unterscheiden - womit auch die PID-Bindung beim Aufraeumen
# ins Leere greift.

function Get-RunningOfficeSessions {
    <#
        Liefert Office-Prozesse mit sichtbarem Hauptfenster, also echte
        Sitzungen des Anwenders - im Unterschied zu unsichtbaren
        Automatisierungs-Instanzen.
    #>
    param([string[]]$ProcessNames = @('EXCEL'))
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
        Die Sitzung wird NICHT beendet - dafuer sorgt die PID-Bindung an
        den Quit-Stellen. Ein sauberer Lauf setzt aber ein geschlossenes
        Excel voraus.

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
    Write-Host "  WARNUNG: Excel laeuft bereits" -ForegroundColor Yellow
    Write-Host ("=" * 66) -ForegroundColor Yellow
    Write-Host "  Gefundene Sitzungen: $namen (PID: $pids)"
    Write-Host ""
    Write-Host "  Ihre Sitzung wird vom Skript NICHT beendet. Waehrend des Laufs"
    Write-Host "  kann sie aber ausgeblendet werden und Warnhinweise sind"
    Write-Host "  abgeschaltet - das wirkt wie ein Absturz."
    Write-Host "  Ausserdem laesst sich die eigene Automatisierungs-Instanz dann"
    Write-Host "  nicht mehr zuverlaessig von Ihrer Sitzung unterscheiden."
    Write-Host ""
    Write-Host "  EMPFEHLUNG: Excel jetzt schliessen und das Skript neu starten." -ForegroundColor Yellow
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
        Write-Host "Abgebrochen. Bitte Excel schliessen und neu starten." -ForegroundColor Cyan
        return $false
    }
    return $true
}

# ==================================================================
# GEMEINSAME COM-INITIALISIERUNG (wird in Start-Job-Scriptbloecke injiziert)
# ==================================================================
$ComInitCode = @'
Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);' -Name "User32" -Namespace "Win32" -ErrorAction SilentlyContinue
$ex = New-Object -ComObject Excel.Application
$ex.Visible             = $false
$ex.DisplayAlerts       = $false
$ex.AskToUpdateLinks    = $false
$ex.AutomationSecurity  = 3
$ex.EnableEvents        = $false
$ex.ScreenUpdating      = $false
try { $ex.Interactive = $false } catch {}
try { $ex.WindowState = -4140 } catch {}
$myExcelPid = $null
try {
    $hwnd = [IntPtr]$ex.Hwnd
    $procId = [uint32]0
    [Win32.User32]::GetWindowThreadProcessId($hwnd, [ref]$procId) | Out-Null
    if ($procId -gt 0) { $myExcelPid = [int]$procId }
} catch {}
if ($myExcelPid) { try { [System.IO.File]::WriteAllText($pidFile, "$myExcelPid") } catch {} }
'@

$ComTeardownCode = @'
if ($wb) {
    try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null } catch {}
    $wb = $null
}
try { $ex.Interactive = $true } catch {}
try { $ex.Quit() } catch {}
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ex) | Out-Null
$ex = $null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
[System.GC]::Collect()
'@

# ==================================================================
# TRUST-CENTER SMOKE-TEST
# ==================================================================
function Test-ExcelTrustCenter {
    param(
        [string]$TempDir,
        [int]$TimeoutSec = 25
    )

    $testXlsx = Join-Path $TempDir ("trustcheck_{0}.xlsx" -f ([Guid]::NewGuid().ToString("N")))
    $pidFile  = Join-Path $TempDir ("trustcheck_{0}.pid"  -f ([Guid]::NewGuid().ToString("N")))

    try {
        New-MinimalXlsx -Path $testXlsx
    } catch {
        return @{ Ok = $false; Msg = "Test-XLSX konnte nicht erstellt werden: $_" }
    }

    $job = Start-Job -ScriptBlock {
        param($fp, $pidFile, $comInit, $comDown)
        Invoke-Expression $comInit
        try {
            $wb = $ex.Workbooks.Open($fp, 0, $true)
            if (-not $wb) { return @{ Ok = $false; Msg = "Open lieferte kein Workbook"; ExcelPid = $myExcelPid } }
            try { $wb.Close($false) } catch {}
            return @{ Ok = $true; ExcelPid = $myExcelPid }
        } catch {
            return @{ Ok = $false; Msg = $_.Exception.Message; ExcelPid = $myExcelPid }
        } finally {
            Invoke-Expression $comDown
        }
    } -ArgumentList $testXlsx, $pidFile, $ComInitCode, $ComTeardownCode

    $completed = Wait-Job $job -Timeout $TimeoutSec

    if (-not $completed) {
        Stop-Job   $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        if ([System.IO.File]::Exists($pidFile)) {
            try {
                $savedPid = [int][System.IO.File]::ReadAllText($pidFile).Trim()
                if ($savedPid -gt 0) {
                    # PID nur tracken, NICHT direkt killen. Bei einem Timeout
                    # kann Excel parallel abgestuerzt sein (AV-Kill, OOM,
                    # interner Crash) und Windows die PID an einen voellig
                    # unbeteiligten Prozess vergeben haben. Ein blindes
                    # Stop-Process wuerde diesen Fremdprozess abschiessen.
                    # Der Cleanup laeuft sicher ueber Stop-AllTrackedExcel
                    # (siehe Z. 1169 im Trust-Center-Fail-Pfad und Z. 1744
                    # im Endlauf), wo Get-Process -Name "EXCEL" + PID-Filter
                    # den Process-Name vorbildlich validiert.
                    $script:TrackedExcelPids.Add($savedPid)
                }
            } catch {}
            try { [System.IO.File]::Delete($pidFile) } catch {}
        }
        try { [System.IO.File]::Delete($testXlsx) } catch {}
        return @{ Ok = $false; Msg = "TIMEOUT nach ${TimeoutSec}s - Trust Center vermutlich nicht konfiguriert oder Office-Profil beschaedigt" }
    }

    $result = Receive-Job $job
    Remove-Job $job -Force -ErrorAction SilentlyContinue
    if ([System.IO.File]::Exists($pidFile)) { try { [System.IO.File]::Delete($pidFile) } catch {} }
    try { [System.IO.File]::Delete($testXlsx) } catch {}
    if ($result -and $result.ExcelPid) { $script:TrackedExcelPids.Add([int]$result.ExcelPid) }

    if ($result -and $result.Ok) {
        return @{ Ok = $true }
    } else {
        $m = if ($result) { $result.Msg } else { "Kein Ergebnis vom COM-Job" }
        return @{ Ok = $false; Msg = $m }
    }
}

# ==================================================================
# KONVERTIERUNG via COM (xls / xlt / xla / xlsb -> xlsx / xltx / xlam)
# ==================================================================
function Convert-ExcelViaCom {
    param(
        [string]$SourcePath,
        [string]$DestPathBase,
        [string]$OriginalExtension,
        [string[]]$Passwords,
        [switch]$PreferPasswordFirst
    )

    $pidFile = Join-Path $TempPath ("ExcelDeepClean_{0}.pid" -f ([Guid]::NewGuid().ToString("N")))

    $job = Start-Job -ScriptBlock {
        param($src, $destBase, $origExt, [object[]]$pws, $pidFile, $comInit, $comDown, $preferPwFirst)

        Invoke-Expression $comInit

        $isOleFormat = ($origExt -match "^\.(xls|xlt|xla)$")
        if ($preferPwFirst -and $pws.Count -gt 0) {
            $tryList = @($pws) + @("")
        } elseif ($isOleFormat) {
            $tryList = @("") + @($pws)
        } else {
            $tryList = @($pws) + @("")
        }

        $wb            = $null
        $lastErr       = "Kein Versuch erfolgreich"
        $successfulPw  = $null
        $usedRepair    = $false
        $usedExtract   = $false

        Unblock-File -LiteralPath $src -ErrorAction SilentlyContinue

        # 0 = xlNormalLoad, 1 = xlRepairFile, 2 = xlExtractData
        # Reihenfolge: erst normal, dann reparieren, zuletzt nur Daten extrahieren.
        # xlExtractData ist Excels letzter Notnagel bei alten/defekten .xls -
        # wird nur erreicht, wenn xlNormalLoad und xlRepairFile beide gescheitert sind.
        foreach ($corruptLoad in @(0, 1, 2)) {
            foreach ($pw in $tryList) {
                try {
                    # Bewusst KEIN [System.Reflection.Missing]::Value mehr fuer die
                    # Skip-Slots: PowerShell marshallt Missing in Late-Bound-COM-
                    # Aufrufen mit vielen Args unzuverlaessig - Excel reagiert dann
                    # mit "Die Open-Eigenschaft des Workbooks-Objektes kann nicht
                    # zugeordnet werden" (vor allem bei alten .xls/.xlsb). Stattdessen
                    # die expliziten API-Defaults aus der Excel-Object-Model-Referenz.
                    $wb = $ex.Workbooks.Open(
                        $src,         # Filename
                        0,            # UpdateLinks: xlUpdateLinksNever
                        $true,        # ReadOnly
                        5,            # Format: xlAutomatic (Default)
                        $pw,          # Password
                        "",           # WriteResPassword: IMMER leer; sonst validiert Excel das Open-PW als Schreibschutz-PW (faelschlicher "Passwort"-Fehler bei .xls mit ReadOnlyRecommended/WriteRes-Hash)
                        $true,        # IgnoreReadOnlyRecommended
                        2,            # Origin: xlWindows (Default)
                        ",",          # Delimiter (irrelevant ohne Format=6/xlCSV)
                        $false,       # Editable: $false bei alten OLE-.xls korrekt; 3b laesst weg, was demselben Defaultverhalten entspricht
                        $false,       # Notify
                        0,            # Converter: kein zusaetzlicher Konverter
                        $false,       # AddToMru
                        $false,       # Local
                        $corruptLoad  # CorruptLoad
                    )
                    $successfulPw = $pw
                    $usedRepair   = ($corruptLoad -ge 1)
                    $usedExtract  = ($corruptLoad -eq 2)
                    break
                } catch {
                    $lastErr = $_.Exception.Message
                    $wb      = $null
                }
            }
            if ($wb) { break }
        }

        if (-not $wb) {
            Invoke-Expression $comDown
            # Originale COM-Fehlermeldung weiterreichen (kein "Passwort:"-Prefix mehr);
            # Categorizer im Aufrufer entscheidet anhand des echten Texts.
            return @{ Status = "ERROR"; Msg = $lastErr; ExcelPid = $myExcelPid }
        }

        try {
            $hasMacros  = $false
            $isTemplate = ($origExt -match "^\.xlt")
            $isAddin    = ($origExt -match "^\.xla$")

            try { if ($wb.HasVBProject) { $hasMacros = $true } } catch { $hasMacros = $true }

            if ($isAddin) {
                $format = 55
                $ext    = ".xlam"
            } elseif ($isTemplate) {
                $format = if ($hasMacros) { 53 } else { 54 }
                $ext    = if ($hasMacros) { ".xltm" } else { ".xltx" }
            } else {
                $format = if ($hasMacros) { 52 } else { 51 }
                $ext    = if ($hasMacros) { ".xlsm" } else { ".xlsx" }
            }

            $final = "$destBase$ext"
            # Defaults statt Missing::Value (siehe Open-Aufruf).
            try {
                $wb.SaveAs($final, $format, "", "", $false, $false, 1, 1, $false)
            } catch {
                $fallbackFormat = $null
                $fallbackExt    = $null
                if     ($format -eq 51) { $fallbackFormat = 52; $fallbackExt = ".xlsm" }
                elseif ($format -eq 54) { $fallbackFormat = 53; $fallbackExt = ".xltm" }

                if ($fallbackFormat) {
                    $format = $fallbackFormat
                    $ext    = $fallbackExt
                    $final  = "$destBase$ext"
                    $wb.SaveAs($final, $format, "", "", $false, $false, 1, 1, $false)
                } else {
                    throw
                }
            }
            $wb.Close($false)

            return @{ Status = "OK"; Path = $final; ExcelPid = $myExcelPid; UsedPassword = $successfulPw; UsedRepair = $usedRepair; UsedExtract = $usedExtract }
        } catch {
            return @{ Status = "ERROR"; Msg = $_.Exception.Message; ExcelPid = $myExcelPid }
        } finally {
            Invoke-Expression $comDown
        }
    # KEIN fuehrendes Komma vor $Passwords: (,$Passwords) umschliesst das
    # string[] mit einem 1-Element-Array, das ueber die Job-Grenze nicht
    # entrollt wird. $tryList enthaelt dann als erstes Element kein Passwort,
    # sondern eine ArrayList - genau der COM-Fehler 'ArrayList kann nicht in
    # Object konvertiert werden'. Die einzelnen Passwoerter werden nie
    # probiert (nachgestellt: Typen=[ArrayList, String] statt vier Strings).
    # Bei leerer Liste verschieben sich die Folgeparameter nicht.
    } -ArgumentList $SourcePath, $DestPathBase, $OriginalExtension, $Passwords, $pidFile, $ComInitCode, $ComTeardownCode, $PreferPasswordFirst.IsPresent

    if (-not (Wait-Job $job -Timeout $FileOpenTimeoutSeconds)) {
        Stop-Job  $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        if ([System.IO.File]::Exists($pidFile)) {
            try {
                $savedPid = [int][System.IO.File]::ReadAllText($pidFile).Trim()
                if ($savedPid -gt 0) {
                    # PID nur tracken, NICHT direkt killen - PID-Recycling-
                    # Schutz. Cleanup via Stop-AllTrackedExcel im catch
                    # der File-Schleife (Timeout-Pfad triggert das zwingend).
                    $script:TrackedExcelPids.Add($savedPid)
                }
            } catch {}
            try { [System.IO.File]::Delete($pidFile) } catch {}
        }
        throw "TIMEOUT: Datei reagiert nicht (evtl. verschluesselt oder Dialog haengt)."
    }

    $result = Receive-Job $job
    Remove-Job $job -Force -ErrorAction SilentlyContinue

    if ([System.IO.File]::Exists($pidFile)) { try { [System.IO.File]::Delete($pidFile) } catch {} }
    if ($result -and $result.ExcelPid) { $script:TrackedExcelPids.Add([int]$result.ExcelPid) }

    if ($result -and $result.Status -eq "OK") {
        return @{ Path = $result.Path; UsedPassword = $result.UsedPassword; UsedRepair = $result.UsedRepair; UsedExtract = $result.UsedExtract }
    } else {
        throw $(if ($result) { $result.Msg } else { "COM-Job lieferte kein Ergebnis" })
    }
}

# ==================================================================
# OEFFNEN-PASSWORT PER COM ENTFERNEN (EXCEL)
# ==================================================================
function Remove-OpenPassword {
    param([string]$FilePath, [string[]]$Passwords)

    $pidFile = Join-Path $TempPath ("ExcelPwRemove_{0}.pid" -f ([Guid]::NewGuid().ToString("N")))

    $job = Start-Job -ScriptBlock {
        param($fp, [object[]]$pws, $pidFile, $comInit, $comDown)

        Invoke-Expression $comInit

        $tryList      = @($pws) + @("")
        $wb           = $null
        $lastErr      = "Kein Passwort hat funktioniert"
        $successfulPw = $null

        Unblock-File -LiteralPath $fp -ErrorAction SilentlyContinue

        foreach ($pw in $tryList) {
            try {
                # Defaults statt Missing::Value (siehe Convert-ExcelViaCom).
                $wb = $ex.Workbooks.Open(
                    $fp,    # Filename
                    0,      # UpdateLinks
                    $false, # ReadOnly
                    5,      # Format: xlAutomatic
                    $pw,    # Password
                    "",     # WriteResPassword
                    $true,  # IgnoreReadOnlyRecommended
                    2,      # Origin: xlWindows
                    ",",    # Delimiter
                    $false, # Editable
                    $false, # Notify
                    0,      # Converter
                    $false, # AddToMru
                    $false, # Local
                    0       # CorruptLoad
                )
                $successfulPw = $pw
                break
            } catch {
                $lastErr = $_.Exception.Message
                $wb      = $null
            }
        }

        if (-not $wb) {
            Invoke-Expression $comDown
            return @{ Status = "ERROR"; Msg = $lastErr; ExcelPid = $myExcelPid }
        }

        try {
            $wb.Password         = ""
            $wb.WriteResPassword = ""
            $wb.Save()
            $wb.Close($false)
            return @{ Status = "OK"; ExcelPid = $myExcelPid; UsedPassword = $successfulPw }
        } catch {
            return @{ Status = "ERROR"; Msg = $_.Exception.Message; ExcelPid = $myExcelPid }
        } finally {
            Invoke-Expression $comDown
        }
    # Kein fuehrendes Komma - Begruendung siehe oben.
    } -ArgumentList $FilePath, $Passwords, $pidFile, $ComInitCode, $ComTeardownCode

    if (-not (Wait-Job $job -Timeout $FileOpenTimeoutSeconds)) {
        Stop-Job  $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        if ([System.IO.File]::Exists($pidFile)) {
            try {
                $savedPid = [int][System.IO.File]::ReadAllText($pidFile).Trim()
                if ($savedPid -gt 0) {
                    # PID nur tracken, NICHT direkt killen - PID-Recycling-
                    # Schutz. Cleanup via Stop-AllTrackedExcel im catch
                    # der File-Schleife (Timeout-Pfad triggert das zwingend).
                    $script:TrackedExcelPids.Add($savedPid)
                }
            } catch {}
            try { [System.IO.File]::Delete($pidFile) } catch {}
        }
        return @{ Success = $false; UsedPassword = $null }
    }

    $result = Receive-Job $job
    Remove-Job $job -Force -ErrorAction SilentlyContinue
    if ([System.IO.File]::Exists($pidFile)) { try { [System.IO.File]::Delete($pidFile) } catch {} }
    if ($result -and $result.ExcelPid) { $script:TrackedExcelPids.Add([int]$result.ExcelPid) }

    if ($result -and $result.Status -eq "OK") {
        return @{ Success = $true; UsedPassword = $result.UsedPassword }
    } else {
        return @{ Success = $false; UsedPassword = $null }
    }
}

# ==================================================================
# SCHUTZ ENTFERNEN (XML-Ebene im ZIP)
# ==================================================================
function Remove-ExcelProtection {
    param([string]$FilePath)

    $actionsTaken = [System.Collections.ArrayList]::new()
    $zip = $null

    try {
        $zip = Invoke-WithRetry -OperationName "ZipFile.Open(Update)" -ScriptBlock {
            [System.IO.Compression.ZipFile]::Open($FilePath, "Update")
        }

        $targetNames = [System.Collections.Generic.List[string]]::new()
        foreach ($entry in $zip.Entries) {
            $fn = $entry.FullName
            if     ($fn -eq "xl/workbook.xml")                                               { $targetNames.Add($fn) }
            elseif ($fn -match "^xl/worksheets/sheet.*\.xml$")                               { $targetNames.Add($fn) }
            elseif ($fn -match "^xl/(chartsheets|macrosheets|dialogsheets)/sheet.*\.xml$")   { $targetNames.Add($fn) }
            elseif ($fn -match "^xl/tables/table.*\.xml$")                                   { $targetNames.Add($fn) }
            elseif ($fn -match "^xl/pivotTables/pivotTable.*\.xml$")                         { $targetNames.Add($fn) }
        }

        if ($targetNames.Count -eq 0) {
            $zip.Dispose()
            $zip = $null
            return ,@()
        }

        # ------------------------------------------------------------------
        # Sheet-Namen-Mapping: technischer Dateiname (z.B. "sheet1") -> angezeigter
        # Tab-Name (z.B. "Übersicht"). Die sheetN.xml-Nummerierung entspricht NICHT
        # zwangsläufig der Anzeige-Reihenfolge im Workbook. Mapping aus:
        #   xl/workbook.xml          : <sheet name="..." r:id="rId..."/>
        #   xl/_rels/workbook.xml.rels: <Relationship Id="rId..." Target="worksheets/sheetN.xml"/>
        # ------------------------------------------------------------------
        $sheetDisplayNames = @{}
        try {
            $wbEntry   = $zip.GetEntry("xl/workbook.xml")
            $relsEntry = $zip.GetEntry("xl/_rels/workbook.xml.rels")
            if ($wbEntry -and $relsEntry) {
                # rId -> Target (z.B. "worksheets/sheet1.xml")
                $relsMap = @{}
                $rs = $relsEntry.Open()
                try {
                    $rr  = [System.IO.StreamReader]::new($rs, [System.Text.Encoding]::UTF8)
                    $rxd = New-Object System.Xml.XmlDocument
                    $rxd.LoadXml($rr.ReadToEnd())
                    $rr.Close()
                    foreach ($rel in @($rxd.SelectNodes("//*[local-name()='Relationship']"))) {
                        $rid    = $rel.GetAttribute("Id")
                        $target = $rel.GetAttribute("Target")
                        if ($rid -and $target) { $relsMap[$rid] = $target }
                    }
                } finally { $rs.Close() }

                # sheet name + r:id aus workbook.xml
                $ws = $wbEntry.Open()
                try {
                    $wr  = [System.IO.StreamReader]::new($ws, [System.Text.Encoding]::UTF8)
                    $wxd = New-Object System.Xml.XmlDocument
                    $wxd.LoadXml($wr.ReadToEnd())
                    $wr.Close()
                    foreach ($sh in @($wxd.SelectNodes("//*[local-name()='sheet']"))) {
                        $name = $sh.GetAttribute("name")
                        # r:id ist namespaced – über LocalName=id durchforsten
                        $rid = $null
                        foreach ($a in $sh.Attributes) {
                            if ($a.LocalName -eq "id") { $rid = $a.Value; break }
                        }
                        if (-not $name -or -not $rid) { continue }
                        $target = $relsMap[$rid]
                        if (-not $target) { continue }
                        # Target ist relativ zu "xl/" – Pfad ggf. normalisieren
                        $targetPath = if ($target -like "/*") { $target.TrimStart("/") }
                                      else                    { "xl/$target" }
                        $short = [System.IO.Path]::GetFileNameWithoutExtension($targetPath)
                        if ($short) { $sheetDisplayNames[$short] = $name }
                    }
                } finally { $ws.Close() }
            }
        } catch {
            # Mapping ist Komfort – Fehler nicht propagieren, einfach mit shortName weiterarbeiten
            Write-DetailedLog "Sheet-Namen-Mapping konnte nicht ermittelt werden: $_" "DEBUG"
        }

        foreach ($targetName in $targetNames) {
            $entry = $zip.GetEntry($targetName)
            if (-not $entry) { continue }

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

            $modified     = $false
            $isWorkbook   = ($targetName -eq "xl/workbook.xml")
            $isTable      = ($targetName -match "^xl/tables/")
            $isPivotTable = ($targetName -match "^xl/pivotTables/")
            $shortName    = [System.IO.Path]::GetFileNameWithoutExtension($targetName)
            # Anzeige-Name (z.B. "Übersicht"); Fallback auf technischen Namen wenn unbekannt
            $displayName  = if ($sheetDisplayNames.ContainsKey($shortName)) {
                                $sheetDisplayNames[$shortName]
                            } else {
                                $shortName
                            }

            if ($isWorkbook) {
                $xpaths = @(
                    "//*[local-name()='workbookProtection']",
                    "//*[local-name()='fileSharing']"
                )
            } elseif ($isTable) {
                $xpaths = @(
                    "//*[local-name()='tableProtection']"
                )
            } elseif ($isPivotTable) {
                $xpaths = @(
                    "//*[local-name()='pivotTableProtection']"
                )
            } else {
                $xpaths = @(
                    "//*[local-name()='sheetProtection']",
                    "//*[local-name()='protectedRanges']"
                )
            }

            foreach ($xpath in $xpaths) {
                $nodes = @($xmlDoc.SelectNodes($xpath))
                foreach ($node in $nodes) {
                    $nodeName = $node.LocalName
                    Write-DetailedLog "Entferne <$nodeName> aus $targetName" "DEBUG"
                    $node.ParentNode.RemoveChild($node) | Out-Null
                    $modified = $true

                    $label = switch ($nodeName) {
                        "workbookProtection"   { "Arbeitsmappenschutz" }
                        "fileSharing"          { "Freigabeschutz" }
                        "tableProtection"      { "Tabellenschutz ($shortName)" }
                        "pivotTableProtection" { "PivotTableSchutz ($shortName)" }
                        "sheetProtection"      {
                            if     ($targetName -match "macrosheets")  { "Makroblattschutz ($displayName)" }
                            elseif ($targetName -match "dialogsheets") { "Dialogblattschutz ($displayName)" }
                            elseif ($targetName -match "chartsheets")  { "Chartschutz ($displayName)" }
                            else                                        { "Blattschutz ($displayName)" }
                        }
                        "protectedRanges"      { "Bereichsschutz ($displayName)" }
                        default                { "$nodeName ($displayName)" }
                    }
                    $actionsTaken.Add($label) | Out-Null
                }
            }

            if ($modified) {
                $entry.Delete()
                $newEntry    = $zip.CreateEntry($targetName)
                $writeStream = $newEntry.Open()

                $writer    = [System.Xml.XmlTextWriter]::new($writeStream, $Utf8NoBom)
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
        throw "XML/ZIP-Fehler: $_"
    } finally {
        if ($zip) { try { $zip.Dispose() } catch {} }
    }

    return ,@($actionsTaken | Select-Object -Unique)
}

# ==================================================================
# ABBRUCH-HANDLER UND TRAP (nach Funktionsdefinitionen)
# ==================================================================
try {
    if ($Host.Name -eq 'ConsoleHost') {
        [Console]::TreatControlCAsInput = $false
        [Console]::add_CancelKeyPress([System.ConsoleCancelEventHandler]{
            param($sender, $e)
            $e.Cancel = $true
            Write-Host "`nABBRUCH angefordert - laufende Datei wird noch fertiggestellt..." -ForegroundColor Yellow
            $script:ShouldStop = $true
        })
    }
} catch {}

trap {
    Write-Warning "Unerwarteter Fehler: $_"
    Start-Sleep -Milliseconds 200
    if (Test-Path -LiteralPath $TempPath) {
        Get-ChildItem -LiteralPath $TempPath -Filter "*.pid" -ErrorAction SilentlyContinue |
            ForEach-Object {
                try {
                    $p = [int][System.IO.File]::ReadAllText($_.FullName).Trim()
                    if ($p -gt 0) { $script:TrackedExcelPids.Add($p) }
                } catch {}
            }
    }
    try { Stop-AllTrackedExcel } catch {}
    try { Sync-LogWriters }      catch {}
    try { Close-LogWriters }     catch {}
    if (Test-Path -LiteralPath $TempPath) { Remove-Item -LiteralPath $TempPath -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false }
    break
}

# ==================================================================
# INTERAKTIVER START
# ==================================================================
if (-not $NoInteractive) {

    Write-Host ""
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "            EXCEL DEEP UNPROTECT                  " -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "VORAUSSETZUNG: Vertrauenswuerdiger Speicherort" -ForegroundColor Yellow
    Write-Host "  Damit die COM-Automatisierung nicht in der Geschuetzten Ansicht" -ForegroundColor Gray
    Write-Host "  einfriert, muss der Dokumentenordner als vertrauenswuerdig" -ForegroundColor Gray
    Write-Host "  eingetragen sein:" -ForegroundColor Gray
    Write-Host "    Excel -> Datei -> Optionen -> Trust Center -> Einstellungen..." -ForegroundColor White
    Write-Host "      -> Vertrauenswuerdige Speicherorte -> Neuen Speicherort hinzufuegen..." -ForegroundColor White
    Write-Host "      -> Pfad: $docRoot" -ForegroundColor White
    Write-Host "      -> Haken: 'Unterordner ... sind ebenfalls vertrauenswuerdig'" -ForegroundColor White
    Write-Host "  Bei Netzlaufwerken als Zielpfad zusaetzlich:" -ForegroundColor Gray
    Write-Host "      -> Haken: 'Vertrauenswuerdige Speicherorte im Netzwerk zulassen'" -ForegroundColor White
    Write-Host ""

    if ([string]::IsNullOrWhiteSpace($TargetPath)) {

        # --------------------------------------------------------------
        # User Shell Folder Aufloesung fuer Desktop und Downloads
        # --------------------------------------------------------------
        $desktopDir   = Get-UserShellFolder -Name 'Desktop'
        $downloadsDir = Get-UserShellFolder -Name 'Downloads'

        $presetCount = $DirectoryPresets.Count
        Write-Host "Zielpfad auswaehlen:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $presetCount; $i++) {
            Write-Host ("  [" + ($i + 1) + "] " + $DirectoryPresets[$i]) -ForegroundColor White
        }
        # Verkettung statt -f: Pfade koennen '{' oder '}' enthalten.
        Write-Host ("  [" + ($presetCount + 1) + "] Desktop ($desktopDir)")     -ForegroundColor White
        Write-Host ("  [" + ($presetCount + 2) + "] Downloads ($downloadsDir)") -ForegroundColor White
        Write-Host ("  [" + ($presetCount + 3) + "] Eigenen Pfad eingeben")     -ForegroundColor White
        Write-Host ""
        $pathChoice = (Read-Host ("Auswahl [1-" + ($presetCount + 3) + "]")).Trim()

        $choiceNum = 0
        [void][int]::TryParse($pathChoice, [ref]$choiceNum)

        switch ($true) {
            ($choiceNum -ge 1 -and $choiceNum -le $presetCount) {
                $TargetPath = $DirectoryPresets[$choiceNum - 1]
                break
            }
            ($choiceNum -eq $presetCount + 1) { $TargetPath = $desktopDir;   break }
            ($choiceNum -eq $presetCount + 2) { $TargetPath = $downloadsDir; break }
            ($choiceNum -eq $presetCount + 3) {
                $rawInput   = Read-Host "Pfad eingeben"
                $TargetPath = if ($null -ne $rawInput) {
                    $rawInput.Trim().Trim([char[]]@('"',"'"))
                } else { '' }
                if ([string]::IsNullOrWhiteSpace($TargetPath)) {
                    Write-Error "Kein Pfad angegeben. Abbruch."
                    exit 1
                }
                break
            }
            default {
                Write-Error ("Ungueltige Auswahl: '$pathChoice'. Erlaubt: 1-" + ($presetCount + 3) + ".")
                exit 1
            }
        }
    }

    Write-Host ""
    Write-Host "Passworte zum Oeffnen geschuetzter Excel-Dateien (bis zu 3, leer = fertig):" -ForegroundColor Yellow
    Write-Host "HINWEIS: Eingabe wird NICHT getrimmt - Leerzeichen am Anfang/Ende bleiben erhalten." -ForegroundColor DarkGray
    for ($pwIdx = 1; $pwIdx -le 3; $pwIdx++) {
        $pwInput = Read-Host "  Passwort $pwIdx (leer = fertig)"
        if ([string]::IsNullOrEmpty($pwInput)) { break }
        $script:Passwords.Add($pwInput)
    }

    Write-Host ""
    Write-Host "Fortschrittsanzeige:" -ForegroundColor Yellow
    Write-Host "  [1] Keine Fortschrittsleiste (laufende Zaehlung inline) [Standard]" -ForegroundColor White
    Write-Host "  [2] Fortschrittsleiste mit ETA (Vorab-Scan noetig, Start verzoegert!)" -ForegroundColor White
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
        Write-Error "Im nicht-interaktiven Modus muss -TargetPath angegeben werden."
        exit 1
    }
    $script:UseProgress = $ShowProgress.IsPresent
    $script:SkipPreScan = $true
}

# ==================================================================
# VORBEREITUNG
# ==================================================================
if (-not (Test-Path -LiteralPath $TargetPath)) {
    Write-Error "Pfad nicht gefunden: $TargetPath"
    exit 1
}

# Vor dem ersten COM-Zugriff (Smoke-Test weiter unten) auf laufende
# Excel-Sitzungen hinweisen. Im Modus -NoInteractive nur warnen, nicht
# fragen - dort sieht niemand die Rueckfrage.
if (-not (Show-OfficeRunningWarning -Silent:$NoInteractive)) { exit 0 }

Remove-StaleTempFolders

if (Test-Path -LiteralPath $TempPath) {
    Remove-Item -LiteralPath $TempPath -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false
}
# [System.IO.Directory]::CreateDirectory statt New-Item.
#
# Hier stand 'New-Item -LiteralPath ... -ItemType Directory'. New-Item
# hat aber gar keinen Parameter -LiteralPath (nur -Path, -Name, -ItemType,
# -Value, -Force, -Credential) - der Aufruf scheiterte mit einer
# ParameterBindingException. Da er nicht abbrechend ist, lief das Skript
# weiter, nur ohne Arbeitsordner: die Zeile darueber hatte einen evtl.
# vorhandenen Vorgaenger gerade geloescht, es gab also keinen Rueckfall.
#
# Folge: der Trust-Center-Smoke-Test konnte seine Testmappe nicht anlegen
# und meldete faelschlich "Dokumentenordner nicht vertrauenswuerdig";
# danach scheiterte fuer JEDE Datei die Arbeitskopie. Das Skript hat in
# diesem Zustand keine einzige Datei entschuetzt.
#
# CreateDirectory ist ausserdem provider-frei und kommt - anders als die
# PowerShell-Cmdlets - zuverlaessig mit '\\?\'-Praefixen zurecht. Es wirft
# nicht, wenn das Verzeichnis bereits existiert.
try {
    [void][System.IO.Directory]::CreateDirectory($TempPath)
} catch {
    Write-Host "FEHLER: Arbeitsordner konnte nicht angelegt werden: $TempPath" -ForegroundColor Red
    Write-Host "        $($_.Exception.Message)" -ForegroundColor Red
    exit 2
}
if (-not [System.IO.Directory]::Exists($TempPath)) {
    Write-Host "FEHLER: Arbeitsordner fehlt nach dem Anlegen: $TempPath" -ForegroundColor Red
    exit 2
}

# Restore-Privilegien (Admin-Kontext) fuer ACL/Owner-Erhalt aktivieren.
Enable-RestorePrivileges

# ------------------------------------------------------------------
# Persistente Log-Writer oeffnen (mit Null-Schutz im Anschluss)
# ------------------------------------------------------------------
try {
    $script:LogWriter         = [System.IO.StreamWriter]::new($LogFilePath,     $false, $Utf8Bom)
    $script:DetailedLogWriter = [System.IO.StreamWriter]::new($DetailedLogPath, $false, $Utf8Bom)
    $script:CsvLogWriter      = [System.IO.StreamWriter]::new($CsvLogPath,      $false, $Utf8Bom)
} catch {
    Write-Warning "Log-Dateien konnten nicht geoeffnet werden: $_"
}

if ($script:LogWriter) {
    try { $script:LogWriter.WriteLine("Log Start: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  |  Ziel: $TargetPath") } catch {}
}
if ($script:DetailedLogWriter) {
    try { $script:DetailedLogWriter.WriteLine("Debug-Log Start: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')") } catch {}
}
if ($script:CsvLogWriter) {
    try { $script:CsvLogWriter.WriteLine("Zeitstempel;Status;Aktionen;Pfad;Details") } catch {}
}
Sync-LogWriters

Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host " EXCEL DEEP UNPROTECT"
Write-Host " Ziel: $TargetPath"
Write-Host "=====================================================" -ForegroundColor Cyan

# ------------------------------------------------------------------
# Excel Smoke-Test
# ------------------------------------------------------------------
Write-Host "Pruefe COM-Subsystem und Trust-Center..." -ForegroundColor DarkGray
$smokeTest = Test-ExcelTrustCenter -TempDir $TempPath -TimeoutSec $TrustCenterTimeoutSec

if (-not $smokeTest.Ok) {
    Write-Host ""
    Write-Host "EXCEL-SMOKE-TEST FEHLGESCHLAGEN" -ForegroundColor Red
    Write-Host "   Grund: $($smokeTest.Msg)"     -ForegroundColor Yellow
    Write-Host ""
    Write-Host "   Moegliche Ursachen:" -ForegroundColor Gray
    Write-Host "     - Dokumentenordner nicht als vertrauenswuerdiger Speicherort eingetragen" -ForegroundColor Gray
    Write-Host "     - Excel nicht installiert oder Office-Profil beschaedigt" -ForegroundColor Gray
    Write-Host "     - Fehlender Desktop-Ordner fuer Dienstkonto (siehe Skript-Header)" -ForegroundColor Gray
    Write-Host ""
    Write-Log "Excel-Smoke-Test fehlgeschlagen: $($smokeTest.Msg)" "ERROR"

    if ($NoInteractive) {
        Write-Log "Abbruch (NoInteractive) nach fehlgeschlagenem Smoke-Test." "ERROR"
        try { Stop-AllTrackedExcel } catch {}
        Sync-LogWriters
        Close-LogWriters
        if (Test-Path -LiteralPath $TempPath) { Remove-Item -LiteralPath $TempPath -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false }
        exit 2
    }

    $choice = (Read-Host "Trotzdem fortfahren? (Timeout-Risiko pro Datei!) [j/N]").Trim().ToUpperInvariant()
    if ($choice -ne "J" -and $choice -ne "Y") {
        Write-Host "Abbruch durch Benutzer." -ForegroundColor Yellow
        Write-Log "Abbruch durch Benutzer nach Smoke-Test-Warnung." "INFO"
        try { Stop-AllTrackedExcel } catch {}
        Sync-LogWriters
        Close-LogWriters
        if (Test-Path -LiteralPath $TempPath) { Remove-Item -LiteralPath $TempPath -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false }
        exit 0
    }
    Write-Log "Smoke-Test-Warnung vom Benutzer ignoriert - Fortsetzung." "WARN"
} else {
    Write-Host "Smoke-Test OK." -ForegroundColor DarkGray
    Write-DetailedLog "Excel-Smoke-Test erfolgreich." "DEBUG"
}

# ------------------------------------------------------------------
# Datei-Extensions
# ------------------------------------------------------------------
$comConvertExt  = "^\.(xls|xlt|xla|xlsb)$"
$xmlDirectExt   = "^\.(xlsx|xlsm|xltx|xltm|xlam)$"
$allExcelExt    = "^\.(xls|xlt|xla|xlsx|xlsm|xltx|xltm|xlsb|xlam)$"

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
        Write-Host "Zum Wiederherstellen die Endung '.bak' bzw. '.bak_xxxxxx' entfernen." -ForegroundColor Gray
        Write-Host ""
    }
    Write-Log ("Backup-Aufraeumen: {0} entfernt, {1} behalten, {2} ohne Original, {3} Fehler" -f `
               $backupCleanup.Deleted, $backupCleanup.Kept, $backupCleanup.Orphans, $backupCleanup.Failed) "INFO"
    Sync-LogWriters
} else {
    Write-DetailedLog "Backup-Aufraeumen uebersprungen (-SkipBackupCleanup)." "DEBUG"
}

if ($script:UseProgress -and -not $script:SkipPreScan) {
    Write-Host "Zaehle Dateien fuer ETA (bitte warten)..." -ForegroundColor DarkGray
    $etaCount = 0
    Get-ExcelFilesRobust (Add-LongPathPrefix $TargetPath) $allExcelExt -StringsOnly | ForEach-Object { $etaCount++ }
    $script:TotalFiles = $etaCount
    Write-Host "Gefunden: $($script:TotalFiles) Excel-Dateien`n" -ForegroundColor DarkGray
}

# Hinweis: 'Locked' ist eine Teilmenge von 'Skipped'.
$stats = @{
    Processed   = 0
    Unlocked    = 0
    Skipped     = 0
    Errors      = 0
    Junk        = 0
    JunkKept    = 0
    Converted   = 0
    Locked      = 0
    WouldChange = 0
}

# ==================================================================
# HAUPTSCHLEIFE
# ==================================================================
Write-Host "Starte Verarbeitung..." -ForegroundColor Cyan

try {
Get-ExcelFilesRobust (Add-LongPathPrefix $TargetPath) $allExcelExt |
    ForEach-Object {

    if ($script:ShouldStop) { throw [System.OperationCanceledException]::new() }

    $file = $_

    $workFile     = $null
    $wasConverted = $false
    $backupPath   = $null
    # 0-Byte-Platzhalter aus Reserve-UniqueDestination (siehe finally).
    $reservedPlaceholder = $null

    $stats.Processed++
    $script:ProcessedCount = $stats.Processed
    Update-Progress -Status "Verarbeite..." -CurrentFile $file.FullName

    if (-not $script:UseProgress) {
        Write-Host -NoNewline "[#$($stats.Processed)] $($file.Name) "
    }

    # ------------------------------------------------------------------
    # Junk-Dateien (Office-Sperrzeichen, macOS-Metadaten)
    # ------------------------------------------------------------------
    # '~$'-Dateien sind Excel-SPERRDATEIEN: sie existieren, solange eine Mappe
    # geoeffnet ist. Sofortiges Loeschen bricht die Sperre und andere Nutzer
    # bekommen keinen Hinweis mehr, dass die Datei in Bearbeitung ist.
    if ($file.Name.StartsWith("~`$") -or $file.Name.StartsWith("._")) {
        $ageHours = ((Get-Date) - $file.LastWriteTime).TotalHours
        if ($ageHours -lt $JunkMinAgeHours) {
            $stats.JunkKept++
            $stats.Skipped++
            if (-not $script:UseProgress) {
                Write-Host ("-> SKIP (Sperrdatei, erst {0:N1} h alt)" -f $ageHours) -ForegroundColor DarkGray
            }
            Write-CsvLog -Status "SKIP" -Actions "Sperrdatei geschont" -Path $file.FullName
            Sync-LogWriters
            return
        }
        if (Confirm-Write $file.FullName 'Verwaiste Sperrdatei loeschen') {
            # Erfolg pruefen, statt ihn anzunehmen. Der Delete lief zuvor in
            # ein leeres catch: schlug er fehl (Datei noch gehalten, keine
            # Rechte), meldete das Skript trotzdem 'Junk geloescht' und zaehlte
            # ihn mit - die Sperrdatei blieb aber liegen und blockierte die
            # zugehoerige Mappe weiter.
            $junkWeg = $false
            try {
                [System.IO.File]::Delete((Add-LongPathPrefix $file.FullName))
                $junkWeg = -not [System.IO.File]::Exists((Add-LongPathPrefix $file.FullName))
            } catch {
                Write-DetailedLog "Sperrdatei nicht loeschbar: $($file.FullName) - $_" "WARN"
            }
            if ($junkWeg) {
                $stats.Junk++
                if (-not $script:UseProgress) { Write-Host "-> JUNK" -ForegroundColor DarkGray }
                Write-DetailedLog "Junk entfernt: $($file.FullName)" "DEBUG"
                Write-CsvLog -Status "JUNK" -Actions "Junk geloescht" -Path $file.FullName
            } else {
                $stats.Errors++
                if (-not $script:UseProgress) { Write-Host "-> ERR (Sperrdatei)" -ForegroundColor Red }
                Write-Log "Verwaiste Sperrdatei nicht loeschbar: $($file.FullName)" "WARN"
                Write-CsvLog -Status "ERR" -Actions "Sperrdatei nicht loeschbar" -Path $file.FullName
            }
        } else {
            $stats.WouldChange++
            if (-not $script:UseProgress) { Write-Host "-> WHATIF (Junk)" -ForegroundColor DarkCyan }
            Write-CsvLog -Status "WHATIF" -Actions "Wuerde Junk loeschen" -Path $file.FullName
        }
        Sync-LogWriters
        return
    }

    $srcLong = Add-LongPathPrefix $file.FullName

    # Nur das ReadOnly-Bit loeschen. Vorher wurde pauschal auf 'Normal' gesetzt
    # und damit auch Hidden, System, Archive und NotContentIndexed entfernt.
    if ($file.IsReadOnly) {
        if (Confirm-Write $file.FullName 'Schreibschutz-Attribut entfernen') {
            try {
                $srcAttrs = [System.IO.File]::GetAttributes($srcLong)
                [System.IO.File]::SetAttributes(
                    $srcLong, $srcAttrs -band (-bnot [System.IO.FileAttributes]::ReadOnly))
                Write-DetailedLog "Dateiattribut ReadOnly entfernt: $($file.FullName)" "DEBUG"
            } catch {
                Write-DetailedLog "Attribut-Reset fehlgeschlagen: $($file.FullName) - $_" "WARN"
            }
        }
    }

    # ------------------------------------------------------------------
    # Lock-Pruefung (einheitlicher String-Rueckgabewert)
    # ------------------------------------------------------------------
    $lockResult = Test-FileIsLocked -FilePath $file.FullName
    if ($lockResult -ne 'Free') {
        $stats.Locked++
        $stats.Skipped++
        if ($lockResult -eq 'AccessDenied') {
            if (-not $script:UseProgress) { Write-Host "-> SKIP (Zugriff verweigert)" -ForegroundColor Red }
            Write-Log         "Zugriff verweigert: $($file.FullName)" "WARN"
            Write-DetailedLog "Keine NTFS-Berechtigung: $($file.FullName)" "WARN"
            Write-CsvLog -Status "SKIP" -Actions "Zugriff verweigert" -Path $file.FullName
        } else {
            if (-not $script:UseProgress) { Write-Host "-> SKIP (gesperrt)" -ForegroundColor Yellow }
            Write-Log         "Gesperrt/geoeffnet: $($file.FullName)" "WARN"
            Write-DetailedLog "Datei gesperrt: $($file.FullName)" "WARN"
            Write-CsvLog -Status "SKIP" -Actions "Gesperrt/geoeffnet" -Path $file.FullName
        }
        Sync-LogWriters
        return
    }

    try {
        $origCreationTimeUtc   = [System.IO.File]::GetCreationTimeUtc($srcLong)
        $origLastWriteTimeUtc  = [System.IO.File]::GetLastWriteTimeUtc($srcLong)
        $origLastAccessTimeUtc = [System.IO.File]::GetLastAccessTimeUtc($srcLong)
    } catch {
        $now = [DateTime]::UtcNow
        $origCreationTimeUtc   = $now
        $origLastWriteTimeUtc  = $now
        $origLastAccessTimeUtc = $now
        Write-DetailedLog "Zeitstempel-Auslesen fehlgeschlagen, Fallback auf Now: $($file.FullName) - $_" "WARN"
    }

    # --- NTFS-Sicherheitsinfo (Owner/Group/DACL) sichern ---
    $origSecurity = Get-FileSecuritySnapshot -Path $srcLong

    $guid     = [Guid]::NewGuid().ToString("N")
    $tempFile = Join-Path $TempPath "$guid$($file.Extension)"
    $tempBase = Join-Path $TempPath $guid
    $workFile = $tempFile

    try {
        # ------------------------------------------------------------------
        # Kopie ins lokale Temp-Verzeichnis + AV-Wartephase
        # ------------------------------------------------------------------
        Invoke-WithRetry -OperationName "Copy src->temp ($($file.Name))" -ScriptBlock {
            [System.IO.File]::Copy($srcLong, $tempFile, $true)
        }
        Unblock-File -LiteralPath $tempFile -ErrorAction SilentlyContinue

        if (-not (Wait-FileAvailable -Path $tempFile -TimeoutSec $WaitFileReadyTimeoutSec)) {
            Write-DetailedLog "Wait-FileAvailable nach Copy ausgelaufen, fahre fort: $tempFile" "WARN"
        }

        $wasConverted        = $false
        $isOleEncrypted      = $false

        # ------------------------------------------------------------------
        # Konvertierungspfad (COM) fuer Legacy-Formate
        # ------------------------------------------------------------------
        if ($file.Extension -match $comConvertExt) {

            if ($file.Extension -match "^\.(xls|xlt|xla)$") {
                $isOleEncrypted = Test-OleFileIsEncrypted -FilePath $tempFile
                if ($isOleEncrypted -and $script:Passwords.Count -eq 0) {
                    $stats.Skipped++
                    if (-not $script:UseProgress) { Write-Host "-> SKIP (OLE-verschluesselt, kein PW)" -ForegroundColor Magenta }
                    Write-Log "Uebersprungen (OLE-verschluesselt, kein PW): $($file.FullName)" "WARN"
                    Write-CsvLog -Status "SKIP" -Actions "OLE-verschluesselt, kein Passwort" -Path $file.FullName
                    Sync-LogWriters
                    return
                }
            }

            try {
                if (-not $script:UseProgress) { Write-Host -NoNewline "[Conv] " -ForegroundColor Cyan }

                $convResult = Convert-ExcelViaCom `
                    -SourcePath $tempFile `
                    -DestPathBase $tempBase `
                    -OriginalExtension $file.Extension `
                    -Passwords $script:Passwords `
                    -PreferPasswordFirst:$isOleEncrypted

                $workFile = $convResult.Path
                if ($convResult.UsedPassword) { Move-PasswordToFront $convResult.UsedPassword }
                if ($convResult.UsedExtract) {
                    Write-Log         "EXTRAKT-Fallback genutzt (xlExtractData=2) - Datenrettungsmodus: $($file.FullName)" "WARN"
                    Write-DetailedLog "EXTRAKT-Fallback genutzt (xlExtractData=2). Datei war schwer beschaedigt; nur Rohdaten gerettet, Formatierung/Formeln/Makros verloren: $($file.FullName)" "WARN"
                } elseif ($convResult.UsedRepair) {
                    Write-DetailedLog "REPAIR-Fallback genutzt (xlRepairFile=1): $($file.FullName)" "WARN"
                }

                if (-not (Wait-FileAvailable -Path $workFile -TimeoutSec $WaitFileReadyTimeoutSec)) {
                    Write-DetailedLog "Wait-FileAvailable nach Konvertierung ausgelaufen: $workFile" "WARN"
                }

                $wasConverted = $true
                $stats.Converted++
                Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false
                Write-DetailedLog "Konvertiert: $($file.FullName) -> $workFile" "DEBUG"

            } catch {
                $msg      = $_.Exception.Message
                $category = Get-ConversionErrorCategory -Message $msg

                if ($category -in @("Timeout", "Passwort")) {
                    Stop-AllTrackedExcel
                    $stats.Skipped++
                    if (-not $script:UseProgress) { Write-Host "-> SKIP ($category)" -ForegroundColor Magenta }
                    Write-Log    "Uebersprungen ($category): $($file.FullName) - $_" "WARN"
                    Write-CsvLog -Status "SKIP" -Actions $category -Path $file.FullName -Details $msg
                } else {
                    $stats.Errors++
                    if (-not $script:UseProgress) { Write-Host "-> ERR ($category)" -ForegroundColor Red }
                    Write-Log         "Konvert-Fehler [$category]: $($file.FullName) - $_" "ERROR"
                    Write-DetailedLog "Konvert-Fehler Detail: $_" "ERROR"
                    Write-CsvLog -Status "ERR" -Actions "Konvert: $category" -Path $file.FullName -Details $msg
                }
                Sync-LogWriters
                return
            }

        } elseif ($file.Extension -match $xmlDirectExt) {
            $workFile = $tempFile

        } else {
            $stats.Skipped++
            if (-not $script:UseProgress) { Write-Host "-> SKIP (Format unbekannt)" -ForegroundColor Gray }
            Write-CsvLog -Status "SKIP" -Actions "Format unbekannt" -Path $file.FullName
            Sync-LogWriters
            return
        }

        # ------------------------------------------------------------------
        # OOXML-Validitaet pruefen, ggf. Oeffnen-Passwort entfernen
        # ------------------------------------------------------------------
        if (-not (Test-IsValidZip -FilePath $workFile)) {
            if ($script:Passwords.Count -gt 0) {
                if (-not $script:UseProgress) { Write-Host -NoNewline "[PW] " -ForegroundColor Cyan }
                $pwResult = Remove-OpenPassword -FilePath $workFile -Passwords $script:Passwords
                if (-not $pwResult.Success) {
                    $stats.Skipped++
                    if (-not $script:UseProgress) { Write-Host "-> SKIP (Passwort)" -ForegroundColor Magenta }
                    Write-Log         "Uebersprungen (Passwort): $($file.FullName)" "WARN"
                    Write-DetailedLog "Passwort-Entfernung fehlgeschlagen: $($file.FullName)" "WARN"
                    Write-CsvLog -Status "SKIP" -Actions "Passwort nicht erkannt" -Path $file.FullName
                    Sync-LogWriters
                    return
                }
                if ($pwResult.UsedPassword) { Move-PasswordToFront $pwResult.UsedPassword }

                if (-not (Wait-FileAvailable -Path $workFile -TimeoutSec $WaitFileReadyTimeoutSec)) {
                    Write-DetailedLog "Wait-FileAvailable nach Passwort-Entfernung ausgelaufen: $workFile" "WARN"
                }
                Write-DetailedLog "Oeffnen-Passwort entfernt: $($file.FullName)" "DEBUG"

                if (-not (Test-IsValidZip -FilePath $workFile)) {
                    $stats.Skipped++
                    if (-not $script:UseProgress) { Write-Host "-> SKIP (nach PW kein ZIP)" -ForegroundColor Magenta }
                    Write-Log "Nach Passwortentfernung kein gueltiges ZIP: $($file.FullName)" "WARN"
                    Write-CsvLog -Status "SKIP" -Actions "Nach PW kein gueltiges ZIP" -Path $file.FullName
                    Sync-LogWriters
                    return
                }
            } else {
                $stats.Skipped++
                if (-not $script:UseProgress) { Write-Host "-> SKIP (verschluesselt, kein PW)" -ForegroundColor Magenta }
                Write-Log "Uebersprungen (verschluesselt, kein PW): $($file.FullName)" "WARN"
                Write-CsvLog -Status "SKIP" -Actions "Verschluesselt, kein PW" -Path $file.FullName
                Sync-LogWriters
                return
            }
        }

        # ------------------------------------------------------------------
        # Schutz-XML entfernen
        # ------------------------------------------------------------------
        $actions    = Remove-ExcelProtection -FilePath $workFile
        $actionsLog = $actions -join ' | '
        Write-DetailedLog "Aktionen ($($file.FullName)): $actionsLog" "DEBUG"

        if (@($actions).Count -gt 0 -or $wasConverted) {

            $actionStrPre = @($actions) -join ' | '
            if ($wasConverted) {
                $actionStrPre = (@('Konvertierung') + @($actions)) -join ' | '
            }

            # Zentrale Freigabe: ab hier wird das Original angefasst. Bei -WhatIf
            # endet die Verarbeitung hier; die Temp-Kopie raeumt der finally-Block.
            if (-not (Confirm-Write $file.FullName "Schutz entfernen ($actionStrPre)")) {
                if ($wasConverted) { $stats.Converted-- }
                $stats.WouldChange++
                if (-not $script:UseProgress) {
                    Write-Host "-> WHATIF ($actionStrPre)" -ForegroundColor DarkCyan
                }
                Write-Log    "Simulation: wuerde aendern: $($file.FullName)  [$actionStrPre]" "INFO"
                Write-CsvLog -Status "WHATIF" -Actions $actionStrPre -Path $file.FullName
                Sync-LogWriters
                return
            }

            $newExt    = [System.IO.Path]::GetExtension($workFile)
            $finalDest = [System.IO.Path]::ChangeExtension($srcLong, $newExt)

            # Namenskollision beim Formatwechsel atomar aufloesen.
            # '$finalDest -ne $srcLong' ergaenzt (Parität zu 2a): ohne diese
            # Bedingung wuerde bei gleichbleibender Endung die eigene Quelldatei
            # als Kollision gewertet und eine ueberfluessige '_2'-Variante
            # angelegt, statt das Original zu ersetzen.
            if ($wasConverted -and ($finalDest -ne $srcLong) -and [System.IO.File]::Exists($finalDest)) {
                $basePath  = $srcLong.Substring(0, $srcLong.LastIndexOf("."))
                $candidate = Reserve-UniqueDestination -BasePath $basePath -Extension $newExt
                if (-not $candidate) {
                    throw "Kollisionsschutz: mehr als $MaxNameClashRetries Varianten der Zieldatei vorhanden"
                }
                $finalDest = $candidate
                $reservedPlaceholder = $candidate
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
                    Write-Log "Veraltetes Backup beiseite gelegt (nicht geloescht): $staleBak" "WARN"
                } catch {
                    Write-DetailedLog "Veraltetes Backup nicht verschiebbar (wird ueberschrieben): $_" "WARN"
                }
            }

            Invoke-WithRetry -OperationName "Copy Backup ($($file.Name))" -ScriptBlock {
                [System.IO.File]::Copy($srcLong, $backupPath, $true)
            }
            # Erzeugungszeit ausdruecklich stempeln: Remove-OrphanedBackups
            # misst das Alter daran. Musste das vorhandene .bak oben
            # ueberschrieben werden (Move fehlgeschlagen), behielte die Datei
            # sonst die ALTE CreationTime und ein parallel laufender Durchgang
            # koennte dieses frische Backup als verwaist loeschen.
            try { [System.IO.File]::SetCreationTime($backupPath, (Get-Date)) } catch {}

            try {
                Invoke-WithRetry -OperationName "Copy work->final ($($file.Name))" -ScriptBlock {
                    [System.IO.File]::Copy($workFile, $finalDest, $true)
                }
                if (-not [System.IO.File]::Exists($finalDest)) {
                    Start-Sleep -Milliseconds 500
                    if (-not [System.IO.File]::Exists($finalDest)) { throw "Verifizierung fehlgeschlagen" }
                }
                $reservedPlaceholder = $null   # Ziel ist real geschrieben

                try { [System.IO.File]::Delete($backupPath) } catch {}
                if ($wasConverted -and $srcLong -ne $finalDest -and [System.IO.File]::Exists($srcLong)) {
                    try {
                        Invoke-WithRetry -OperationName "Delete original ($($file.Name))" -ScriptBlock {
                            [System.IO.File]::Delete($srcLong)
                        }
                    } catch {
                        Write-DetailedLog "Alte Quelldatei konnte nicht geloescht werden (wird beibehalten): $srcLong - $_" "WARN"
                        Write-Log "Quelldatei beibehalten: $($file.FullName) - $_" "WARN"
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
                # (Admin-Kontext: vollstaendig; Nutzer-Kontext: DACL)
                Set-FileSecuritySnapshot -Path $finalDest -Snapshot $origSecurity

                $stats.Unlocked++
                $actionStr = $actions -join ' | '
                if ([string]::IsNullOrWhiteSpace($actionStr)) { $actionStr = "Nur Format konvertiert" }
                if (-not $script:UseProgress) {
                    Write-Host "-> OK ($actionStr)" -ForegroundColor Green
                }
                Write-Log    "Erledigt: $($file.FullName)  [$actionStr]" "INFO"
                Write-CsvLog -Status "OK" -Actions $actionStr -Path $file.FullName -Details $finalDest

            } catch {
                $stats.Errors++
                if ($finalDest -ne $srcLong -and [System.IO.File]::Exists($finalDest)) {
                    try {
                        [System.IO.File]::Delete($finalDest)
                        Write-DetailedLog "Unvollstaendige Zieldatei entfernt: $finalDest" "WARN"
                    } catch {}
                }
                if ([System.IO.File]::Exists($backupPath)) {
                    $backupSafeToDelete = $true
                    if ($finalDest -eq $srcLong) {
                        $backupSafeToDelete = $false
                        try {
                            Invoke-WithRetry -OperationName "Rollback from Backup" -ScriptBlock {
                                [System.IO.File]::Copy($backupPath, $srcLong, $true)
                            }
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
                if ($wasConverted) { $stats.Converted-- }
                if (-not $script:UseProgress) { Write-Host "-> ERR (Schreiben)" -ForegroundColor Red }
                Write-Log         "Schreib-Fehler: $($file.FullName) - $_" "ERROR"
                Write-DetailedLog "Schreib-Fehler Detail: $_" "ERROR"
                Write-CsvLog -Status "ERR" -Actions "Schreib-Fehler" -Path $file.FullName -Details $_.Exception.Message
            }

        } else {
            if (-not $script:UseProgress) { Write-Host "-> Kein Schutz" -ForegroundColor Gray }
            Write-CsvLog -Status "NO_CHANGE" -Actions "Kein Schutz" -Path $file.FullName
        }

    } catch {
        $stats.Errors++
        if ($wasConverted) { $stats.Converted-- }
        if (-not $script:UseProgress) { Write-Host "-> ERROR: $($_.Exception.Message)" -ForegroundColor Red }
        Write-Log         "Allg. Fehler: $($file.FullName) - $_" "ERROR"
        Write-DetailedLog "Allg. Fehler Detail: $_" "ERROR"
        Write-CsvLog -Status "ERR" -Actions "Allg. Fehler" -Path $file.FullName -Details $_.Exception.Message

    } finally {
        if (-not [string]::IsNullOrWhiteSpace($workFile)) {
            Remove-Item -LiteralPath $workFile -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false
        }
        # Reservierten Platzhalter nur entfernen, wenn er leer geblieben ist.
        if ($reservedPlaceholder) {
            try {
                $ph = [System.IO.FileInfo]::new($reservedPlaceholder)
                if ($ph.Exists -and $ph.Length -eq 0) {
                    [System.IO.File]::Delete($reservedPlaceholder)
                    Write-DetailedLog "Reservierten Platzhalter entfernt: $reservedPlaceholder" "DEBUG"
                }
            } catch {}
        }
        Sync-LogWriters
    }
}
} catch [System.OperationCanceledException] {
    Write-Host "`nVerarbeitung durch Nutzer abgebrochen." -ForegroundColor Yellow
    Write-Log "Verarbeitung durch Nutzer abgebrochen." "WARN"
}

# ==================================================================
# AUFRAEUMEN & STATISTIK
# ==================================================================
if ($script:UseProgress) { Write-Progress -Activity "Fertig" -Completed }

try {
    Stop-AllTrackedExcel
} catch {
    Write-DetailedLog "Stop-AllTrackedExcel Fehler beim Endlauf: $_" "WARN"
}
if (Test-Path -LiteralPath $TempPath) { Remove-Item -LiteralPath $TempPath -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$false }

$duration = (Get-Date) - $ScriptStartTime
$durStr   = $duration.ToString('hh\:mm\:ss')

Write-Host "`n=====================================================" -ForegroundColor Cyan
Write-Host " FERTIG in $durStr" -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host "Geprueft:        $($stats.Processed)"   -ForegroundColor White
Write-Host "Entsperrt/Konv:  $($stats.Unlocked)"    -ForegroundColor Green
Write-Host "Konvertiert:     $($stats.Converted)"   -ForegroundColor Green
if ($stats.WouldChange -gt 0) {
    Write-Host "Wuerde aendern:  $($stats.WouldChange)  (Simulation - nichts geschrieben)" -ForegroundColor DarkCyan
}
if ($backupCleanup) {
    Write-Host "Backups entf.:   $($backupCleanup.Deleted)" -ForegroundColor DarkGray
    if ($backupCleanup.Orphans -gt 0) {
        Write-Host "Backups o. Orig: $($backupCleanup.Orphans)  (NICHT geloescht - pruefen!)" -ForegroundColor Yellow
    }
}
Write-Host "Junk geloescht:  $($stats.Junk)"        -ForegroundColor DarkGray
if ($stats.JunkKept -gt 0) {
    Write-Host "Sperrdateien:    $($stats.JunkKept)  (juenger als $JunkMinAgeHours h - geschont)" -ForegroundColor DarkGray
}
Write-Host "Uebersprungen:   $($stats.Skipped)"     -ForegroundColor Magenta
Write-Host "  davon Locked:  $($stats.Locked)"      -ForegroundColor Yellow
Write-Host "Fehler:          $($stats.Errors)"      -ForegroundColor Red
if ($script:RetryRecovered -gt 0 -or $script:RetryFailed -gt 0) {
    Write-Host "Retry erfolgr.:  $($script:RetryRecovered)"  -ForegroundColor DarkCyan
    Write-Host "Retry gescheit.: $($script:RetryFailed)"     -ForegroundColor DarkCyan
}
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host "Logs:" -ForegroundColor Cyan
Write-Host "  $LogFilePath"     -ForegroundColor Gray
Write-Host "  $DetailedLogPath" -ForegroundColor Gray
Write-Host "  $CsvLogPath"      -ForegroundColor Gray
Write-Host "=====================================================" -ForegroundColor Cyan

$footerLines = @(
    "=================================================================="
    "Log Ende:  $(Get-Date)"
    "Dauer:     $durStr"
    "Geprueft:  $($stats.Processed)  |  Entsperrt: $($stats.Unlocked)  |  Fehler: $($stats.Errors)"
    "Konv:      $($stats.Converted)  |  Skip: $($stats.Skipped)  (davon Locked: $($stats.Locked))  |  Junk: $($stats.Junk)"
)
if ($script:RetryRecovered -gt 0 -or $script:RetryFailed -gt 0) {
    $footerLines += "Retry:     erfolgreich $($script:RetryRecovered)  |  gescheitert $($script:RetryFailed)"
}
$footerLines += "=================================================================="
if ($script:LogWriter) {
    foreach ($line in $footerLines) {
        try { $script:LogWriter.WriteLine($line) } catch {}
    }
}
Sync-LogWriters
Invoke-WindowsTempCleanup
Close-LogWriters

Write-Host ""
if (-not $NoInteractive) {
    Write-Host "Beliebige Taste druecken zum Beenden..." -ForegroundColor DarkGray
    $useReadKey = $false
    try {
        if ($Host.Name -eq 'ConsoleHost' -and -not [Console]::IsInputRedirected) {
            $useReadKey = $true
        }
    } catch {}
    if ($useReadKey) {
        try {
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        } catch {
            try { Read-Host | Out-Null } catch {}
        }
    } else {
        try { Read-Host | Out-Null } catch {}
    }
}
