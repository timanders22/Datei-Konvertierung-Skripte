<#
.SYNOPSIS
    PowerPoint Deep Unprotect
.DESCRIPTION
    Entfernt Schutz aus PowerPoint-Dateien (XML-Ebene) und konvertiert
    alte .ppt/.pps/.pot -> .pptx/.ppsx/.potx. Netzlaufwerk- und
    Langpfad-faehig. Robust gegen AV-Locks durch Retry, Priming-Read
    und Stabilitaets-Polling der Temp-Kopie.

    Voraussetzung Vertrauenswuerdige Speicherorte:
      Das Skript verwendet den Ordner "Dokumente\2c_entferne_schutz_powerpoint_<PID>" als
      temporaeren Arbeitsordner fuer die COM-Automatisierung. Damit PowerPoint
      Dateien dort ohne Geschuetzte Ansicht oeffnet, muss der Dokumentenordner
      des ausfuehrenden Benutzers als vertrauenswuerdiger Speicherort
      eingetragen sein:
        PowerPoint -> Datei -> Optionen -> Trust Center -> Einstellungen...
          -> Vertrauenswuerdige Speicherorte
          -> "Neuen Speicherort hinzufuegen..."
          -> Pfad: C:\Users\<Benutzername>\Documents
          -> Haken: "Unterordner dieses Speicherorts sind ebenfalls vertrauenswuerdig"
      Bei Netzlaufwerken als Zielpfad zusaetzlich:
          -> "Vertrauenswuerdige Speicherorte im Netzwerk zulassen"
      Das Skript fuehrt vor dem Hauptlauf einen PowerPoint-Smoke-Test durch
      und bricht bei Fehlschlag ab (bzw. fragt interaktiv nach), um stunden-
      lange Timeouts pro Datei zu vermeiden.

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
# KONFIGURATION
# ==================================================================
$script:scriptDir              = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) { $PWD.Path } else { $PSScriptRoot }

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
$script:LogDir                 = Resolve-LogDirectory -Preferred $script:scriptDir

$script:TempPath               = Join-Path ([Environment]::GetFolderPath("MyDocuments")) "2c_entferne_schutz_powerpoint_$PID"
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
    ppAlertsNone                            =  2
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
                Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
                if (-not (Test-Path -LiteralPath $_.FullName)) { $removedDirs++ }
            } else {
                Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
                if (-not (Test-Path -LiteralPath $_.FullName)) { $removedFiles++ }
            }
        } catch {}
    }
    Write-DetailedLog "Windows-Temp-Cleanup (Whitelist): $removedFiles Dateien, $removedDirs Ordner aus $winTemp entfernt, $skipped fremde Eintraege unangetastet." "DEBUG"
}

function Remove-StaleTempFolders {
    # Verwaiste Arbeitsordner abgebrochener Laeufe entfernen - nur eigene
    # 2c_entferne_schutz_powerpoint_<PID>-Ordner, deren PID nicht mehr lebt.
    $parent = [Environment]::GetFolderPath('MyDocuments')
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
                    try { Remove-Item -LiteralPath $d.FullName -Recurse -Force -ErrorAction SilentlyContinue } catch {}
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

function Get-RunningOfficeSessions {
    <#
        Liefert Office-Prozesse mit sichtbarem Hauptfenster, also echte
        Sitzungen des Anwenders - im Unterschied zu unsichtbaren
        Automatisierungs-Instanzen.
    #>
    param([string[]]$ProcessNames = @('POWERPNT'))
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

        PowerPoint wird per COM automatisiert. Laeuft bereits eine Sitzung des
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
        Das Skript kannte bisher gar kein -WhatIf; ein Probelauf auf einer
        fremden Ablage war damit nicht moeglich.
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

# ==================================================================
# ABBRUCH-HANDLER
# ==================================================================
try {
    if ($Host.Name -eq 'ConsoleHost') {
        [Console]::TreatControlCAsInput = $false
        [Console]::add_CancelKeyPress([System.ConsoleCancelEventHandler]{
            param($sender, $e)
            $e.Cancel = $true
            Write-Host "`n!!  ABBRUCH angefordert - laufende Datei wird noch fertiggestellt..." -ForegroundColor Yellow
            $script:ShouldStop = $true
        })
    }
} catch {}

trap {
    Write-Warning "Unerwarteter Fehler: $_"
    Start-Sleep -Milliseconds 200
    if ($script:TrackedPptPids -and (Test-Path -LiteralPath $script:TempPath)) {
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
    if ($script:TrackedPptPids) { Clear-TrackedPowerPointInstances }
    if (Test-Path -LiteralPath $script:TempPath) { Remove-Item -LiteralPath $script:TempPath -Recurse -Force -ErrorAction SilentlyContinue }
    Close-Loggers
    break
}

# ==================================================================
# HELPER: PFADE & UMGEBUNG
# ==================================================================
function Get-LongPath {
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
function Invoke-WithFileRetry {
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
        Invoke-WithFileRetry -Action {
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
        $fs = Invoke-WithFileRetry -Action { [System.IO.File]::OpenRead($FilePath) }
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
    $lp = Get-LongPath $FilePath
    $stream = $null
    try {
        if (-not [System.IO.File]::Exists($lp)) { return $false }
        $stream = [System.IO.File]::Open($lp,
                       [System.IO.FileMode]::Open,
                       [System.IO.FileAccess]::ReadWrite,
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
        Write-DetailedLog "Sicherheitsinfo nicht lesbar ($Path): $_" "DEBUG"
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
            try { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue } catch {}
        }
    }
    [System.GC]::Collect()
    $script:TrackedPptPids.Clear()
}

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

function Get-OriginalFromBackupPath {
    # Vortrag.pptx.bak / Vortrag.pptx.bak_a1b2c3  ->  Vortrag.pptx
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
          1. Der abgeleitete Originalname endet auf eine PowerPoint-Endung
             (fremde .bak-Dateien bleiben unangetastet),
          2. das Original existiert und ist groesser als 0 Byte,
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

    foreach ($bak in (Get-PptFilesRobust $RootPath '^\.bak(_[0-9a-fA-F]+)?$')) {
        try {
            $orig = Get-OriginalFromBackupPath -BackupPath $bak
            if ([string]::IsNullOrWhiteSpace($orig)) { continue }
            if ($orig -notmatch '\.(ppt|pot|pps|pptx|pptm|potx|potm|ppsx|ppsm)$') {
                Write-DetailedLog "Fremde .bak-Datei ignoriert: $bak" "DEBUG"
                continue
            }

            $fi = [System.IO.FileInfo]::new($bak)
            if (-not $fi.Exists) { continue }

            $origOk = $false
            if ([System.IO.File]::Exists($orig)) {
                try { $origOk = ([System.IO.FileInfo]::new($orig)).Length -gt 0 } catch { $origOk = $false }
            }
            if (-not $origOk) {
                $res.Orphans++
                $res.OrphanList.Add($bak)
                Write-Log "Backup ohne intaktes Original - NICHT geloescht, bitte pruefen: $bak" "WARN"
                continue
            }

            if ($fi.LastWriteTime -gt $cutoff) { $res.Kept++; continue }

            if (Confirm-Write $bak 'Verwaistes Backup loeschen') {
                try {
                    [System.IO.File]::Delete((Get-LongPath $bak))
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

    $pidFile = Join-Path $script:TempPath "PptDeepClean_$([Guid]::NewGuid().ToString()).pid"
    $job = Start-Job -ScriptBlock {
        param($src, $destBase, $origExt, $consts, $pidFile)

        $ppt = New-Object -ComObject PowerPoint.Application

        $ppt.Visible            = $consts.msoTrue
        $ppt.DisplayAlerts      = $consts.ppAlertsNone
        $ppt.FeatureInstall     = $consts.msoFeatureInstallNone
        try { $ppt.AutomationSecurity = $consts.msoAutomationSecurityForceDisable } catch {}
        try { $ppt.WindowState = $consts.ppWindowMinimized } catch {}

        $myPptPid = $null
        try {
            Add-Type -Name "User32Ppt" -Namespace "Win32" -MemberDefinition @'
[DllImport("user32.dll")]
public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
'@ -ErrorAction SilentlyContinue
            $procId = 0
            [Win32.User32Ppt]::GetWindowThreadProcessId([IntPtr]$ppt.Hwnd, [ref]$procId) | Out-Null
            if ($procId -gt 0) { $myPptPid = $procId }
        } catch {}
        if ($myPptPid) { try { [System.IO.File]::WriteAllText($pidFile, "$myPptPid") } catch {} }

        $pres    = $null
        $lastErr = "Oeffnen fehlgeschlagen"
        try {
            $pres = $ppt.Presentations.Open($src, -1, -1, 0)
        } catch {
            $lastErr = $_.Exception.Message
        }

        if (-not $pres) {
            # Quit() nur fuer die selbst gestartete Instanz.
            # PowerPoint ist wie Word ein Einzelinstanz-COM-Server: laeuft
            # bereits eine Sitzung des Anwenders, liefert New-Object
            # -ComObject GENAU DIESE Instanz. $myPptPid bleibt dann leer -
            # und ein Quit() darauf schliesst die Sitzung des Anwenders
            # samt ungespeicherter Praesentationen. Dann nur die Referenz
            # freigeben; eine evtl. doch eigene Instanz raeumt die
            # Zombie-Erkennung des Aufrufers spaeter ab.
            if ($myPptPid) { try { $ppt.Quit() } catch {} }
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ppt) | Out-Null
            return @{ Status = "ERROR"; Msg = $lastErr; PptPid = $myPptPid }
        }

        try {
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

            return @{ Status = "OK"; Path = $final; PptPid = $myPptPid }
        } catch {
            try {
                if ($pres) {
                    try { $pres.Saved = $consts.msoTrue } catch {}
                    $pres.Close()
                }
            } catch {}
            return @{ Status = "ERROR"; Msg = $_.Exception.Message; PptPid = $myPptPid }
        } finally {
            if ($pres) {
                try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($pres) | Out-Null } catch {}
                $pres = $null
            }
            # Quit nur fuer die eigene Instanz (siehe Hinweis oben).
            if ($myPptPid) { try { $ppt.Quit() } catch {} }
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ppt) | Out-Null
        }
    } -ArgumentList $SourcePath, $DestPathBase, $OriginalExt, $script:PptConstants, $pidFile

    if (-not (Wait-Job $job -Timeout $script:FileOpenTimeoutSeconds)) {
        Stop-Job  $job
        Remove-Job $job
        if ([System.IO.File]::Exists($pidFile)) {
            try {
                $pidContent = [System.IO.File]::ReadAllText($pidFile).Trim()
                if ($pidContent -match '^\d+$') {
                    $savedPid = [int]$pidContent
                    if ($savedPid -gt 0) {
                        $p = Get-Process -Id $savedPid -ErrorAction SilentlyContinue
                        if ($p -and $p.ProcessName -eq 'POWERPNT') {
                            try { Stop-Process -Id $savedPid -Force -ErrorAction SilentlyContinue } catch {}
                            Write-DetailedLog "Timeout: POWERPNT.EXE PID $savedPid sofort beendet" "WARN"
                        }
                    }
                }
            } catch {}
            [System.IO.File]::Delete($pidFile)
        }
        throw [System.TimeoutException]::new("Datei reagiert nicht (evtl. Oeffnungspasswort oder haengender Dialog).")
    }

    $result = Receive-Job $job
    Remove-Job $job
    if ([System.IO.File]::Exists($pidFile)) { try { [System.IO.File]::Delete($pidFile) } catch {} }
    if ($result.PptPid) { $script:TrackedPptPids.Add([int]$result.PptPid) }

    if ($result.Status -eq "OK") { return $result.Path }
    else                         { throw $result.Msg }
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

    $lockAttrs = @("noSelect","noMove","noResize","noRot","noGrp",
                   "noChangeAspect","noChangeArrowheads","noTextEdit",
                   "noAdjustHandles","noEditPoints","noUngrp")

    try {
        $zip = Invoke-WithFileRetry -Context "ZipFile.Open Update" -Action {
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
                foreach ($node in $lockNodes) {
                    $hadLock = $false
                    foreach ($attr in $lockAttrs) {
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
        param($fp, $pidFile, $consts)

        $ppt = New-Object -ComObject PowerPoint.Application
        $ppt.Visible       = $consts.msoTrue
        $ppt.DisplayAlerts = $consts.ppAlertsNone
        $ppt.FeatureInstall = $consts.msoFeatureInstallNone
        try { $ppt.AutomationSecurity = $consts.msoAutomationSecurityForceDisable } catch {}
        try { $ppt.WindowState = $consts.ppWindowMinimized } catch {}

        $myPptPid = $null
        try {
            Add-Type -Name "User32PptTc" -Namespace "Win32" -MemberDefinition @'
[DllImport("user32.dll")]
public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
'@ -ErrorAction SilentlyContinue
            $procId = 0
            [Win32.User32PptTc]::GetWindowThreadProcessId([IntPtr]$ppt.Hwnd, [ref]$procId) | Out-Null
            if ($procId -gt 0) { $myPptPid = $procId }
        } catch {}
        if ($myPptPid) { try { [System.IO.File]::WriteAllText($pidFile, "$myPptPid") } catch {} }

        $pres1 = $null
        $pres2 = $null
        try {
            # 1) Leere Praesentation erzeugen und in den Temp-Ordner speichern
            $pres1 = $ppt.Presentations.Add($consts.msoTrue)
            if (-not $pres1) {
                return @{ Ok = $false; Msg = "Presentations.Add lieferte keine Praesentation"; PptPid = $myPptPid }
            }
            $pres1.SaveAs($fp, $consts.ppSaveAsOpenXMLPresentation)
            $pres1.Close()
            $pres1 = $null

            # 2) Re-Open aus dem Temp-Ordner -> eigentlicher Trust-Center-Test
            $pres2 = $ppt.Presentations.Open($fp, -1, -1, 0)
            if (-not $pres2) {
                return @{ Ok = $false; Msg = "Presentations.Open lieferte keine Praesentation"; PptPid = $myPptPid }
            }
            $pres2.Close()
            $pres2 = $null

            return @{ Ok = $true; PptPid = $myPptPid }
        } catch {
            return @{ Ok = $false; Msg = $_.Exception.Message; PptPid = $myPptPid }
        } finally {
            if ($pres1) { try { $pres1.Close() } catch {} }
            if ($pres2) { try { $pres2.Close() } catch {} }
            # Quit nur fuer die eigene Instanz (siehe Hinweis oben).
            if ($myPptPid) { try { $ppt.Quit() } catch {} }
            try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ppt) | Out-Null } catch {}
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    } -ArgumentList $testPath, $pidFile, $script:PptConstants

    $completed = Wait-Job $job -Timeout $TimeoutSec

    if (-not $completed) {
        try { Stop-Job   $job -ErrorAction SilentlyContinue } catch {}
        try { Remove-Job $job -Force -ErrorAction SilentlyContinue } catch {}
        if ([System.IO.File]::Exists($pidFile)) {
            try {
                $pidContent = [System.IO.File]::ReadAllText($pidFile).Trim()
                if ($pidContent -match '^\d+$') {
                    $savedPid = [int]$pidContent
                    if ($savedPid -gt 0) {
                        $script:TrackedPptPids.Add($savedPid)
                        # Process-Name-Check vor Stop-Process schuetzt vor
                        # PID-Recycling: Wenn PowerPoint waehrend des Smoke-
                        # Tests abstuerzt (AV-Kill, Office-Profilschaden),
                        # gibt Windows die PID sofort frei und kann sie an
                        # einen unbeteiligten Systemprozess (svchost, Update,
                        # Virenscanner) vergeben. Ein blindes Stop-Process
                        # wuerde diesen Fremdprozess abschiessen.
                        # Konsistent mit der Implementierung in
                        # Convert-PptToPptx (Z. 633-637) und mit
                        # Clear-TrackedPowerPointInstances (Z. 442-449).
                        $p = Get-Process -Id $savedPid -ErrorAction SilentlyContinue
                        if ($p -and $p.ProcessName -eq 'POWERPNT') {
                            try { Stop-Process -Id $savedPid -Force -ErrorAction SilentlyContinue } catch {}
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
    try { Remove-Job $job -Force -ErrorAction SilentlyContinue } catch {}
    if ([System.IO.File]::Exists($pidFile)) { try { [System.IO.File]::Delete($pidFile) } catch {} }
    try { [System.IO.File]::Delete($testPath) } catch {}
    if ($result -and $result.PptPid) { $script:TrackedPptPids.Add([int]$result.PptPid) }

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
if (-not $NoInteractive.IsPresent) {

    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║             POWERPOINT DEEP UNPROTECT              ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════╝" -ForegroundColor Cyan

    # Vor dem ersten COM-Zugriff auf laufende Office-Sitzungen hinweisen.
    if (-not (Show-OfficeRunningWarning)) { exit 0 }
    Write-Host ""
    Write-Host "HINWEIS: Das Skript verwendet den Ordner 'Dokumente' als temporären" -ForegroundColor DarkYellow
    Write-Host "Arbeitsordner für die COM-Verarbeitung (PowerPoint öffnet Dateien daraus)." -ForegroundColor DarkYellow
    Write-Host "" -ForegroundColor DarkYellow
    Write-Host "Bitte in PowerPoint sicherstellen:" -ForegroundColor DarkYellow
    Write-Host "  Datei > Optionen > Trust Center > Einstellungen für das Trust Center" -ForegroundColor Gray
    Write-Host "  > Vertrauenswürdige Speicherorte:" -ForegroundColor Gray
    Write-Host "    1. Den Dokumente-Ordner hinzufügen (falls nicht vorhanden):" -ForegroundColor Gray
    Write-Host "       $([Environment]::GetFolderPath('MyDocuments'))" -ForegroundColor White
    Write-Host "       [x] Unterordner dieser Speicherorte sind ebenfalls vertrauenswürdig" -ForegroundColor Gray
    Write-Host "    2. [x] Vertrauenswürdige Speicherorte im Netzwerk zulassen" -ForegroundColor Gray
    Write-Host "       (erforderlich, wenn Dokumente-Ordner auf Netzlaufwerk umgeleitet)" -ForegroundColor Gray
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

    # Auch ohne Dialog protokollieren, dass fremde Office-Sitzungen
    # laufen - im Fehlerfall ist das die haeufigste Ursache.
    [void](Show-OfficeRunningWarning -Silent)
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

Remove-StaleTempFolders

if (Test-Path -LiteralPath $script:TempPath) {
    Remove-Item -LiteralPath $script:TempPath -Recurse -Force -ErrorAction SilentlyContinue
}
New-Item -Path $script:TempPath -ItemType Directory -Force | Out-Null

# Restore-Privilegien (Admin-Kontext) fuer ACL/Owner-Erhalt aktivieren.
Enable-RestorePrivileges

Initialize-Loggers

Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host " POWERPOINT DEEP UNPROTECT"
Write-Host " Ziel: $TargetPath"
Write-Host "=====================================================" -ForegroundColor Cyan

# ------------------------------------------------------------------
# PowerPoint Smoke-Test
# ------------------------------------------------------------------
Write-Host "Prüfe COM-Subsystem und Trust-Center..." -ForegroundColor DarkGray
$smokeTest = Test-PowerPointTrustCenter -TempDir $script:TempPath -TimeoutSec $script:TrustCenterTimeoutSec

if (-not $smokeTest.Ok) {
    Write-Host ""
    Write-Host "POWERPOINT-SMOKE-TEST FEHLGESCHLAGEN" -ForegroundColor Red
    Write-Host "   Grund: $($smokeTest.Msg)"          -ForegroundColor Yellow
    Write-Host ""
    Write-Host "   Mögliche Ursachen:" -ForegroundColor Gray
    Write-Host "     - Dokumentenordner nicht als vertrauenswürdiger Speicherort eingetragen" -ForegroundColor Gray
    Write-Host "     - PowerPoint nicht installiert oder Office-Profil beschädigt" -ForegroundColor Gray
    Write-Host "     - Fehlender Desktop-Ordner für Dienstkonto" -ForegroundColor Gray
    Write-Host ""
    Write-Log "PowerPoint-Smoke-Test fehlgeschlagen: $($smokeTest.Msg)" "ERROR"

    if ($NoInteractive.IsPresent) {
        Write-Log "Abbruch (NoInteractive) nach fehlgeschlagenem Smoke-Test." "ERROR"
        Clear-TrackedPowerPointInstances
        Close-Loggers
        if (Test-Path $script:TempPath) { Remove-Item $script:TempPath -Recurse -Force -ErrorAction SilentlyContinue }
        exit 2
    }

    $choice = (Read-Host "Trotzdem fortfahren? (Timeout-Risiko pro Datei!) [j/N]").Trim().ToUpperInvariant()
    if ($choice -ne "J" -and $choice -ne "Y") {
        Write-Host "Abbruch durch Benutzer." -ForegroundColor Yellow
        Write-Log "Abbruch durch Benutzer nach Smoke-Test-Warnung." "INFO"
        Clear-TrackedPowerPointInstances
        Close-Loggers
        if (Test-Path $script:TempPath) { Remove-Item $script:TempPath -Recurse -Force -ErrorAction SilentlyContinue }
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
    $backupCleanup = Remove-OrphanedBackups -RootPath (Get-LongPath $TargetPath)

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
} else {
    Write-DetailedLog "Backup-Aufraeumen uebersprungen (-SkipBackupCleanup)." "DEBUG"
}

if ($script:UseProgress -and -not $script:SkipPreScan) {
    Write-Host "Zaehle Dateien fuer ETA (bitte warten)..." -ForegroundColor DarkGray
    $etaCount = 0
    Get-PptFilesRobust (Get-LongPath $TargetPath) $allPptExt | ForEach-Object { $etaCount++ }
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
    AvBlocked   = 0
    WouldChange = 0
}

# ==================================================================
# HAUPTSCHLEIFE
# ==================================================================
Write-Host "Starte Verarbeitung..." -ForegroundColor Cyan

try {
Get-PptFilesRobust (Get-LongPath $TargetPath) $allPptExt |
    ForEach-Object {

    if ($script:ShouldStop) { throw [System.OperationCanceledException]::new() }

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
        $ageHours = 0.0
        try { $ageHours = ((Get-Date) - ([System.IO.File]::GetLastWriteTime($srcLong))).TotalHours } catch {}
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
        try { [System.IO.File]::Delete((Get-LongPath $filePath)) } catch {}
        $stats.Junk++
        if (-not $script:UseProgress) { Write-Host "-> JUNK" -ForegroundColor DarkGray }
        Write-DetailedLog "Junk entfernt: $filePath" "DEBUG"
        Write-CsvLog -Status "JUNK" -Actions "Sperrdatei geloescht" -Path $filePath
        return
    }

    $srcLong = Get-LongPath $filePath

    # 2. Schreibschutz-Attribut entfernen
    $fileAttrs = $null
    try { $fileAttrs = [System.IO.File]::GetAttributes($srcLong) } catch {}
    if ($fileAttrs -and ($fileAttrs -band [System.IO.FileAttributes]::ReadOnly)) {
        try {
            if (Confirm-Write $filePath 'Schreibschutz-Attribut entfernen') {
                [System.IO.File]::SetAttributes($srcLong, $fileAttrs -band (-bnot [System.IO.FileAttributes]::ReadOnly))
                Write-DetailedLog "Datei-Schreibschutz (Attribut) entfernt: $filePath" "DEBUG"
            }
        } catch {}
    }

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
        Invoke-WithFileRetry -Context "Copy src->temp" -Action { [System.IO.File]::Copy($srcLong, $tempFile, $true) }

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
            Unblock-File -LiteralPath $tempFile -ErrorAction SilentlyContinue
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
                $stats.Converted++
                Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue
                Write-DetailedLog "Konvertiert: $filePath -> $workFile" "DEBUG"
            } catch [System.TimeoutException] {
                Clear-TrackedPowerPointInstances
                $stats.Skipped++
                $stats.Encrypted++
                if (-not $script:UseProgress) { Write-Host "-> SKIP (Timeout/verschluesselt)" -ForegroundColor Magenta }
                Write-Log "Uebersprungen (Timeout): $filePath - $_" "WARN"
                Write-CsvLog -Status "SKIP" -Actions "Timeout" -Path $filePath -Details $_.Exception.Message
                return
            } catch {
                $stats.Errors++
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
                $stats.Encrypted++
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

            # Zentrale Freigabe: ab hier wird das Original angefasst.
            if (-not (Confirm-Write $filePath "Schutz entfernen ($actionStrPre)")) {
                if ($wasConverted) { $stats.Converted-- }
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
                    return
                }

                $finalDest           = $reserved
                $reservedPlaceholder = $reserved
                Write-DetailedLog "Namenskonflikt: Ziel reserviert als $finalDest" "WARN"
                if (-not $script:UseProgress) { Write-Host -NoNewline "[Umben.] " -ForegroundColor Yellow }
            }

            $backupPath = "$srcLong.bak"

            if ([System.IO.File]::Exists($backupPath)) {
                # Ein zurueckgebliebenes .bak stammt aus einem harten Abbruch
                # eines frueheren Laufs und kann die einzige intakte Kopie des
                # Originals sein. Daher beiseite legen statt ueberschreiben.
                $staleBak = "$srcLong.bak_$([Guid]::NewGuid().ToString('N').Substring(0,6))"
                try {
                    [System.IO.File]::Move($backupPath, $staleBak)
                    Write-Log "Veraltetes Backup beiseite gelegt (nicht geloescht): $staleBak" "WARN"
                } catch {
                    Write-DetailedLog "Veraltetes Backup nicht verschiebbar (wird ueberschrieben): $_" "WARN"
                }
            }

            Invoke-WithFileRetry -Action { [System.IO.File]::Copy($srcLong, $backupPath, $true) }

            try {
                if ($finalDest -eq $srcLong -and [System.IO.File]::Exists($finalDest)) {
                    [System.IO.File]::Delete($finalDest)
                }
                Invoke-WithFileRetry -Action { [System.IO.File]::Copy($workFile, $finalDest, $true) }
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
                $actionStr = $actions -join ', '
                if ([string]::IsNullOrWhiteSpace($actionStr)) { $actionStr = "Nur Format konvertiert" }
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
                if ([System.IO.File]::Exists($backupPath)) {
                    $backupSafeToDelete = $true
                    if ($finalDest -eq $srcLong) {
                        $backupSafeToDelete = $false
                        try {
                            [System.IO.File]::Copy($backupPath, $srcLong, $true)
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
                Write-Log "Schreib-Fehler: $filePath - $_" "ERROR"
                Write-DetailedLog "Schreib-Fehler Detail: $_" "ERROR"
                Write-CsvLog -Status "ERR" -Actions "Schreib-Fehler" -Path $filePath -Details $_.Exception.Message
            }

        } else {
            if ($hadZoneIdentifier) {
                $zoneBak = "$srcLong.bak"

                if ([System.IO.File]::Exists($zoneBak)) {
                    # Stale-Backup beiseite legen statt ueberschreiben
                    # (siehe Begruendung im Rueckschreib-Pfad).
                    $staleBak = "$srcLong.bak_$([Guid]::NewGuid().ToString('N').Substring(0,6))"
                    try {
                        [System.IO.File]::Move($zoneBak, $staleBak)
                        Write-Log "Veraltetes Backup beiseite gelegt (nicht geloescht): $staleBak" "WARN"
                    } catch {
                        Write-DetailedLog "Veraltetes Backup nicht verschiebbar (wird ueberschrieben): $_" "WARN"
                    }
                }

                try {
                    Invoke-WithFileRetry -Action { [System.IO.File]::Copy($srcLong, $zoneBak, $true) }
                    [System.IO.File]::Delete($srcLong)
                    Invoke-WithFileRetry -Action { [System.IO.File]::Copy($workFile, $srcLong, $true) }
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
                    if (-not $script:UseProgress) { Write-Host "-> OK (Zone.Identifier)" -ForegroundColor Green }
                    Write-Log "Erledigt: $filePath  [Zone.Identifier]" "INFO"
                    Write-CsvLog -Status "OK" -Actions "Zone.Identifier entfernt" -Path $filePath
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
                            Invoke-WithFileRetry -Action { [System.IO.File]::Copy($zoneBak, $srcLong, $true) }
                            if ([System.IO.File]::Exists($srcLong)) {
                                $backupSafeToDelete = $true
                                Set-FileSecuritySnapshot -Path $srcLong -Snapshot $origSecurity
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
            } else {
                if (-not $script:UseProgress) { Write-Host "-> Kein Schutz" -ForegroundColor Gray }
                Write-CsvLog -Status "NO_CHANGE" -Actions "Kein Schutz gefunden" -Path $filePath
            }
        }

    } catch {
        $stats.Errors++
        if ($wasConverted) { $stats.Converted-- }
        if ($backupPath -and [System.IO.File]::Exists($backupPath)) {
            try { [System.IO.File]::Delete($backupPath) } catch {}
        }
        if (-not $script:UseProgress) { Write-Host "-> ERROR: $($_.Exception.Message)" -ForegroundColor Red }
        Write-Log "Allg. Fehler: $filePath - $_" "ERROR"
        Write-DetailedLog "Allg. Fehler Detail: $_" "ERROR"
        Write-CsvLog -Status "ERR" -Actions "Allg. Fehler" -Path $filePath -Details $_.Exception.Message

    } finally {
        if (-not [string]::IsNullOrWhiteSpace($workFile)) {
            Remove-Item -LiteralPath $workFile -Force -ErrorAction SilentlyContinue
        }
        if (-not [string]::IsNullOrWhiteSpace($tempFile) -and ($tempFile -ne $workFile)) {
            Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue
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
    Write-Host "`nVerarbeitung durch Nutzer abgebrochen." -ForegroundColor Yellow
    Write-Log "Verarbeitung durch Nutzer abgebrochen." "WARN"
}

# ==================================================================
# AUFRAEUMEN & STATISTIK
# ==================================================================
if ($script:UseProgress) { Write-Progress -Activity "Fertig" -Completed }

Clear-TrackedPowerPointInstances
if (Test-Path $script:TempPath) { Remove-Item $script:TempPath -Recurse -Force -ErrorAction SilentlyContinue }

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
    "Konv:      $($stats.Converted)  |  Skip: $($stats.Skipped)  |  Verschluesselt: $($stats.Encrypted)  |  AV-blockiert: $($stats.AvBlocked)  |  Gesperrt: $($stats.Locked)"
    "=================================================================="
)
if ($script:LogWriter) {
    foreach ($line in $footerLines) { try { $script:LogWriter.WriteLine($line) } catch {} }
}

Invoke-WindowsTempCleanup
Close-Loggers

if (-not $NoInteractive.IsPresent) {
    Write-Host ""
    Write-Host "Beliebige Taste druecken zum Beenden..." -ForegroundColor DarkGray
    try { $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") } catch {}
}
