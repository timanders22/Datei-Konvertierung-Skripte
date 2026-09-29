<#
.SYNOPSIS
    Excel Calculation Mode Fixer - LIVE MODUS
    Speichern als: UTF-8 mit BOM

    Korrigiert in xlsx/xlsm/xltx/xltm-Dateien einen versehentlich oder
    absichtlich gesetzten manuellen Berechnungsmodus zurueck auf 'auto',
    inklusive 'calcOnSave=0' und 'calcCompleted=0'. Arbeitet rein auf
    OOXML-Ebene (ZIP/XML), kein Excel/COM erforderlich.

    Compilieren in eine .exe (Windows PowerShell 5.1, STA-Konsolenanwendung):
      # Einmalig: ps2exe installieren
      Install-Module -Name ps2exe -Scope CurrentUser -Force

      Invoke-ps2exe -inputFile   '.\6_Excel_automatische_Berechnung.ps1' -outputFile  '.\6_Excel_automatische_Berechnung.exe' -iconFile    '.\powershell_icon.ico' -title       'Excel Calculation Mode Fixer' -description 'Setzt manuellen Excel-Berechnungsmodus auf automatisch zurueck' -product     'Excel Calculation Mode Fixer' -version     '1.0.1.0' -STA -x64 -supportOS

    Simulation:
      -WhatIf zeigt an, welche Dateien geaendert wuerden, ohne sie
      anzufassen. Ohne Parameter wird der Modus interaktiv abgefragt.

    Stand: 11.06.2026
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Position = 0, Mandatory = $false)]
    [string]$TargetPath,

    # Nur pruefen, nichts schreiben (identisch zu -WhatIf, aber sprechender).
    [Parameter(Mandatory = $false)]
    [switch]$ReadOnlyMode,

    [Parameter(Mandatory = $false)]
    [switch]$NoInteractive,

    # Ueberspringt das Aufraeumen zurueckgebliebener .bak_-Dateien.
    [Parameter(Mandatory = $false)]
    [switch]$SkipBackupCleanup
)

# $PSCmdlet ist nur im Skript-Scope verfuegbar, nicht in Funktionen.
$script:ScriptCmdlet = $PSCmdlet

# Steuert Wait-AnyKey. Im Modus -NoInteractive darf am Ende nichts auf
# einen Tastendruck warten, sonst haengt ein geplanter Task dauerhaft.
$script:NoWaitOnExit = $NoInteractive.IsPresent

function Confirm-Write {
    <#
        Zentrale Freigabe fuer jede schreibende Operation am Original.
        Vorher steuerte eine hartkodierte Variable ($ReadOnly = $false in
        Zeile 25) den Vorschau-Modus - fuer einen Probelauf musste man das
        Skript editieren.
    #>
    param(
        [string]$Target,
        [string]$Action = 'Berechnungsmodus korrigieren'
    )
    if ($script:PreviewOnly) { return $false }
    # Rueckfall ueber $WhatIfPreference statt hart $true: ist $PSCmdlet nicht
    # verfuegbar (dot-sourced, ps2exe-Host) oder wirft ShouldProcess, wurde
    # vorher trotz -WhatIf tatsaechlich geschrieben. $WhatIfPreference wird vom
    # gemeinsamen Parameter im Skript-Scope gesetzt und ist in Funktionen
    # sichtbar (gemessen).
    if ($null -eq $script:ScriptCmdlet) { return (-not $WhatIfPreference) }
    try {
        return $script:ScriptCmdlet.ShouldProcess($Target, $Action)
    } catch {
        return (-not $WhatIfPreference)
    }
}

# Verzeichnisse, die bei der Suche nicht betreten werden. Ohne diese Liste
# wurden auch Mappen im Papierkorb und in Schattenkopien korrigiert.
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

# Mindestalter, ab dem ein zurueckgebliebenes .bak_ als verwaist gilt.
$script:BackupCleanupMinAgeHours = 24

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Vorschau-Modus: entweder per -ReadOnlyMode/-WhatIf oder interaktiv.
# Wird weiter unten endgueltig gesetzt.
$script:PreviewOnly = $ReadOnlyMode.IsPresent
# Nur fuer die Anzeige (Kopfzeile, Abschlussbericht). Der Schreibschutz selbst
# haengt an Confirm-Write, nicht an dieser Variablen.
$ReadOnly           = $script:PreviewOnly -or $WhatIfPreference


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

# --- HILFSFUNKTIONEN ---

function Add-LongPathPrefix {
    param([string]$Path)
    if ([string]::IsNullOrEmpty($Path)) { return $Path }
    if ($Path -like "\\?\*") { return $Path }
    if ($Path -like "\\*")   { return "\\?\UNC\" + $Path.TrimStart('\') }
    return "\\?\" + $Path
}

function Remove-LongPathPrefix {
    param([string]$Path)
    if ([string]::IsNullOrEmpty($Path)) { return $Path }
    if ($Path -like "\\?\UNC\*") { return "\\" + $Path.Substring(8) }
    if ($Path -like "\\?\*")     { return $Path.Substring(4) }
    return $Path
}

function Format-CsvField {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return "" }
    $Text = $Text -replace "[`r`n`t]", ' '
    return ($Text -replace '"', '""')
}

function Test-IsZipFile {
    # Returns: $true (ZIP), $false (definitiv keine ZIP), $null (nicht pruefbar - z.B. gesperrt)
    param([string]$FilePath)
    try {
        $fs = [System.IO.File]::Open(
            $FilePath,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::ReadWrite
        )
        try {
            if ($fs.Length -lt 2) { return $false }
            $bytes = New-Object byte[] 2
            $read = $fs.Read($bytes, 0, 2)
            if ($read -lt 2) { return $false }
            return ($bytes[0] -eq 0x50 -and $bytes[1] -eq 0x4B)
        } finally {
            $fs.Close()
            $fs.Dispose()
        }
    } catch {
        return $null
    }
}

function Remove-RestrictiveAttributes {
    <#
        Entfernt ReadOnly, Hidden und System, damit File.Copy das Ziel
        ueberschreiben kann, und liefert die Liste der entfernten Attribute
        zurueck.

        Hidden und System werden nach dem Zurueckkopieren von
        Restore-FileAttributes wieder gesetzt - vorher blieben sie dauerhaft
        entfernt, was den Dateizustand still veraenderte. ReadOnly bleibt
        bewusst entfernt: sonst blockiert es jeden kuenftigen Lauf erneut.
    #>
    param([string]$FilePath)
    $removed = New-Object System.Collections.Generic.List[string]
    try {
        $attrs    = [System.IO.File]::GetAttributes($FilePath)
        $newAttrs = $attrs
        if ($attrs -band [System.IO.FileAttributes]::ReadOnly) {
            $newAttrs = $newAttrs -band (-bnot [System.IO.FileAttributes]::ReadOnly)
            $removed.Add("ReadOnly")
        }
        if ($attrs -band [System.IO.FileAttributes]::Hidden) {
            $newAttrs = $newAttrs -band (-bnot [System.IO.FileAttributes]::Hidden)
            $removed.Add("Hidden")
        }
        if ($attrs -band [System.IO.FileAttributes]::System) {
            $newAttrs = $newAttrs -band (-bnot [System.IO.FileAttributes]::System)
            $removed.Add("System")
        }
        if ($newAttrs -ne $attrs) {
            [System.IO.File]::SetAttributes($FilePath, $newAttrs)
        }
    } catch { }
    return ,$removed
}

function Restore-FileAttributes {
    <#
        Setzt Hidden/System nach dem Zurueckkopieren wieder. ReadOnly wird
        bewusst NICHT wiederhergestellt (siehe Remove-RestrictiveAttributes).
    #>
    param(
        [string]$FilePath,
        [string[]]$Attributes
    )
    if (-not $Attributes -or $Attributes.Count -eq 0) { return @() }
    $restored = New-Object System.Collections.Generic.List[string]
    try {
        $attrs = [System.IO.File]::GetAttributes($FilePath)
        $neu   = $attrs
        if ($Attributes -contains 'Hidden') {
            $neu = $neu -bor [System.IO.FileAttributes]::Hidden
            $restored.Add('Hidden')
        }
        if ($Attributes -contains 'System') {
            $neu = $neu -bor [System.IO.FileAttributes]::System
            $restored.Add('System')
        }
        if ($neu -ne $attrs) { [System.IO.File]::SetAttributes($FilePath, $neu) }
    } catch { }
    return ,$restored
}

function Get-FileTimestamps {
    param([string]$FilePath)
    try {
        return @{
            Creation   = [System.IO.File]::GetCreationTimeUtc($FilePath)
            LastWrite  = [System.IO.File]::GetLastWriteTimeUtc($FilePath)
            LastAccess = [System.IO.File]::GetLastAccessTimeUtc($FilePath)
        }
    } catch {
        return $null
    }
}

function Restore-FileTimestamps {
    <#
        Schreibt die vor der Aenderung gesicherten Zeitstempel zurueck.

        Ohne das traegt jede korrigierte Datei nach dem Lauf das heutige
        Aenderungsdatum: File.Copy uebernimmt die Zeitstempel der Temp-Kopie,
        und die wurde beim Schreiben ins ZIP auf "jetzt" gesetzt.
    #>
    param(
        [string]$FilePath,
        $Timestamps
    )
    if ($null -eq $Timestamps) { return $false }
    try {
        [System.IO.File]::SetCreationTimeUtc($FilePath,   $Timestamps.Creation)
        [System.IO.File]::SetLastWriteTimeUtc($FilePath,  $Timestamps.LastWrite)
        [System.IO.File]::SetLastAccessTimeUtc($FilePath, $Timestamps.LastAccess)
        return $true
    } catch {
        return $false
    }
}

function Wait-AnyKey {
    <#
        Wartet nur, wenn wirklich jemand zusehen kann.

        Vorher wartete das Skript bedingungslos - auch im Modus
        -NoInteractive, der ausdruecklich fuer geplante Tasks gedacht ist.
        Ein solcher Task beendete sich nie und belegte bei jedem Lauf
        einen Prozess. Zusaetzlich zur Parameterpruefung wird der Host
        abgefragt: unter ps2exe -noConsole oder bei umgeleiteter Eingabe
        wirft ReadKey, statt zu warten.
    #>
    if ($script:NoWaitOnExit) { return }
    if ($Host.Name -ne 'ConsoleHost') { return }
    try { if ([Console]::IsInputRedirected) { return } } catch {}
    Write-Host ""
    Write-Host "Beliebige Taste druecken zum Beenden..." -ForegroundColor DarkGray
    try {
        [void]$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    } catch {}
}

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

function Get-FilesRecursiveSafe {
    param(
        [Parameter(Mandatory)][string]$RootLongPath,
        [Parameter(Mandatory)][string[]]$Extensions
    )
    $extSet = New-Object System.Collections.Generic.HashSet[string](
        [string[]]$Extensions, [System.StringComparer]::OrdinalIgnoreCase)

    $stack = New-Object System.Collections.Generic.Stack[string]
    $stack.Push($RootLongPath)

    while ($stack.Count -gt 0) {
        $dir = $stack.Pop()

        try {
            foreach ($f in [System.IO.Directory]::EnumerateFiles($dir)) {
                $ext = [System.IO.Path]::GetExtension($f)
                if (-not [string]::IsNullOrEmpty($ext) -and $extSet.Contains($ext)) {
                    $f
                }
            }
        } catch [System.UnauthorizedAccessException] {
        } catch [System.IO.IOException] {
        } catch {
        }

        try {
            foreach ($sub in [System.IO.Directory]::EnumerateDirectories($dir)) {
                $leaf = [System.IO.Path]::GetFileName($sub.TrimEnd('\'))
                if ($script:ExcludeDirNames -contains $leaf) { continue }
                # Junctions/Symlinks nicht folgen: sonst werden Dateien
                # doppelt verarbeitet oder die Suche laeuft im Kreis.
                try {
                    $attr = [System.IO.File]::GetAttributes($sub)
                    if ($attr -band [System.IO.FileAttributes]::ReparsePoint) { continue }
                } catch { continue }
                $stack.Push($sub)
            }
        } catch [System.UnauthorizedAccessException] {
        } catch [System.IO.IOException] {
        } catch {
        }
    }
}

# --- VERZEICHNISAUSWAHL ---
if (-not $NoInteractive) { Clear-Host }
Write-Host "============================================" -ForegroundColor Cyan
Write-Host " Excel Calculation Mode Fixer" -ForegroundColor Cyan
Write-Host " Berechnungsmodus von 'Manuell' auf" -ForegroundColor Cyan
Write-Host " 'Automatisch' korrigieren" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host " Welches Verzeichnis soll durchsucht werden?" -ForegroundColor White
Write-Host ""

$desktopDir   = Get-UserShellFolder -Name 'Desktop'
$downloadsDir = Get-UserShellFolder -Name 'Downloads'

# Mit -TargetPath entfaellt das Menue vollstaendig (fuer geplante Tasks).
if (-not [string]::IsNullOrWhiteSpace($TargetPath)) {
    $rootPath = $TargetPath.Trim().Trim('"').Trim("'")
    if ($rootPath -match '^[A-Za-z]:$') { $rootPath += '\' }
} else {
    if ($NoInteractive) {
        Write-Host "Im Modus -NoInteractive ist -TargetPath erforderlich." -ForegroundColor Red
        exit 1
    }

$presetCount = $script:DirectoryPresets.Count
for ($i = 0; $i -lt $presetCount; $i++) {
    Write-Host ("  [" + ($i + 1) + "] " + $script:DirectoryPresets[$i]) -ForegroundColor Yellow
}
# Verkettung statt -f: Pfade koennen '{' oder '}' enthalten.
Write-Host ("  [" + ($presetCount + 1) + "] Desktop ($desktopDir)")     -ForegroundColor Yellow
Write-Host ("  [" + ($presetCount + 2) + "] Downloads ($downloadsDir)") -ForegroundColor Yellow
Write-Host ("  [" + ($presetCount + 3) + "] Eigenen Pfad eingeben")     -ForegroundColor Yellow
Write-Host ""
Write-Host "  [0] Abbrechen" -ForegroundColor DarkGray
Write-Host ""
Write-Host "--------------------------------------------"

$selection = Read-Host ("Auswahl (0-" + ($presetCount + 3) + ")")

$choiceNum = 0
[void][int]::TryParse($selection, [ref]$choiceNum)

switch ($true) {
    ($choiceNum -ge 1 -and $choiceNum -le $presetCount) {
        $rootPath = $script:DirectoryPresets[$choiceNum - 1]
        break
    }
    ($choiceNum -eq $presetCount + 1) { $rootPath = $desktopDir;   break }
    ($choiceNum -eq $presetCount + 2) { $rootPath = $downloadsDir; break }
    ($choiceNum -eq $presetCount + 3) {
        Write-Host ""
        $customPath = Read-Host "Vollstaendigen Pfad eingeben"
        $customPath = $customPath.Trim('"').Trim("'")
        if ($customPath -match '^[A-Za-z]:$') { $customPath = $customPath + '\' }
        if (-not [string]::IsNullOrWhiteSpace($customPath) -and (Test-Path -LiteralPath $customPath)) {
            $rootPath = $customPath
        } else {
            Write-Host "Pfad ungueltig oder nicht erreichbar: $customPath" -ForegroundColor Red
            Wait-AnyKey
            exit 1
        }
        break
    }
    ($selection -eq "0") {
        Write-Host "Abgebrochen." -ForegroundColor Yellow
        exit 0
    }
    default {
        Write-Host "Ungueltige Auswahl. Skript wird beendet." -ForegroundColor Red
        Wait-AnyKey
        exit 1
    }
}
}

# --- VORSCHAU-MODUS ---
# -WhatIf und -ReadOnlyMode haben Vorrang; sonst interaktiv nachfragen.
if (-not $script:PreviewOnly -and -not $NoInteractive -and
    -not $PSBoundParameters.ContainsKey('WhatIf')) {
    Write-Host ""
    $pv = Read-Host "Nur pruefen, nichts aendern (Probelauf)? [j/N]"
    if ($pv -and $pv.Trim().ToLowerInvariant() -in @('j','ja','y','yes')) {
        $script:PreviewOnly = $true
    }
}
$ReadOnly = $script:PreviewOnly -or $WhatIfPreference

if (-not (Test-Path -LiteralPath $rootPath)) {
    Write-Host ""
    Write-Host "FEHLER: Das gewaehlte Verzeichnis ist nicht erreichbar:" -ForegroundColor Red
    Write-Host "  $rootPath" -ForegroundColor Red
    Write-Host ""
    Write-Host "Moegliche Ursachen:" -ForegroundColor Yellow
    Write-Host "  - Laufwerk nicht verbunden / nicht gemountet" -ForegroundColor Yellow
    Write-Host "  - Netzwerkverbindung unterbrochen" -ForegroundColor Yellow
    Write-Host "  - Keine Berechtigung" -ForegroundColor Yellow
    Wait-AnyKey
    exit 1
}

Write-Host ""
Write-Host "Gewaehltes Verzeichnis: $rootPath" -ForegroundColor Green
Write-Host ""

# --- TEMPORAERER ORDNER ---
$runId      = [guid]::NewGuid().ToString("N").Substring(0, 8)
# GetFolderPath beruecksichtigt umgeleitete Documents-Ordner (z.B. OneDrive)
$docsFolder = [Environment]::GetFolderPath('MyDocuments')
$tempFolder = Join-Path $docsFolder "6_Excel_automatische_Berechnung_$runId"

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

# --- CLEANUP: alte Temp-Ordner (>24h) und alte Logs (>30 Tage) ---
$staleTempThreshold = (Get-Date).AddHours(-24)
Get-ChildItem -LiteralPath $docsFolder -Directory -Filter "6_Excel_automatische_Berechnung_*" -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -lt $staleTempThreshold } |
    ForEach-Object {
        try {
            Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop -WhatIf:$false
            Write-Host "[CLEANUP]   Alter Temp-Ordner entfernt: $($_.Name)" -ForegroundColor DarkGray
        } catch {
            Write-Host "[WARNUNG]   Alter Temp-Ordner nicht loeschbar: $($_.Name)" -ForegroundColor DarkYellow
        }
    }

# Provider-frei: New-Item gehorcht -WhatIf und legte den Ordner dann NICHT an.
# Die Arbeitskopie in Schritt 2 lief danach bei JEDER Datei auf
# DirectoryNotFoundException - der Probelauf meldete lauter Fehler statt
# Befunde. Der Tempordner ist Infrastruktur und muss auch im Probelauf stehen.
if (-not [System.IO.Directory]::Exists($tempFolder)) {
    [void][System.IO.Directory]::CreateDirectory($tempFolder)
}

# --- LOG-DATEI (CSV, UTF-8 ohne BOM) ---
$logEncoding = New-Object System.Text.UTF8Encoding($false)
$logName     = "6_Excel_automatische_Berechnung_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# Log-Verzeichnis: bevorzugt neben dem Skript, sonst LOCALAPPDATA, sonst TEMP.
# Vorher wurde ungeprueft in $scriptDir geschrieben - lag das Skript auf einer
# schreibgeschuetzten Freigabe, brach der Lauf hier mit unbehandelter
# Ausnahme ab, bevor ueberhaupt eine Datei geprueft wurde.
$logFile = $null
foreach ($cand in @($scriptDir,
                    (Join-Path $env:LOCALAPPDATA 'ExcelCalcModeFixer'),
                    $env:TEMP)) {
    if ([string]::IsNullOrWhiteSpace($cand)) { continue }
    try {
        if (-not [System.IO.Directory]::Exists($cand)) {
            # Ebenfalls provider-frei - sonst faellt das Log unter -WhatIf
            # stillschweigend auf den naechsten Kandidaten zurueck.
            [void][System.IO.Directory]::CreateDirectory($cand)
        }
        $try = Join-Path $cand $logName
        [System.IO.File]::WriteAllText(
            $try, '"Zeitstempel";"Dateiname";"Pfad";"Status";"Details"' + [Environment]::NewLine, $logEncoding)
        $logFile = $try
        break
    } catch { continue }
}
if (-not $logFile) {
    Write-Host "WARNUNG: Keine Log-Datei schreibbar - Lauf wird nicht protokolliert." -ForegroundColor Red
}

# --- CLEANUP: alte Logs (>30 Tage) ---
# Muss NACH der Wahl des Log-Verzeichnisses laufen und in genau diesem
# Verzeichnis. Vorher wurde fest in $scriptDir aufgeraeumt - also
# ausgerechnet dort nicht, wo das Skript bei schreibgeschuetztem
# Skriptordner tatsaechlich protokolliert (LOCALAPPDATA oder TEMP).
# Die Protokolle wuchsen dort unbegrenzt.
if ($logFile) {
    $staleLogThreshold = (Get-Date).AddDays(-30)
    $logDirActual      = Split-Path -Parent $logFile
    Get-ChildItem -LiteralPath $logDirActual -File -Filter "6_Excel_automatische_Berechnung_*.csv" -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $staleLogThreshold -and $_.FullName -ne $logFile } |
        ForEach-Object {
            try {
                Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop -WhatIf:$false
                Write-Host "[CLEANUP]   Alte Log-Datei entfernt: $($_.Name)" -ForegroundColor DarkGray
            } catch { }
        }
}

function Write-Log {
    param(
        [string]$FileName,
        [string]$FilePath,
        [string]$Status,
        [string]$Details = ""
    )

    # Gemeinsames Laufprotokoll (migration.jsonl) - ergaenzt die CSV,
    # ersetzt sie nicht. Erst damit laesst sich der Fortschritt ueber
    # alle elf Schritte hinweg auswerten.
    if ($script:GemeinsamGeladen) {
        try {
            Write-Laufprotokoll -Skript '6_Excel_automatische_Berechnung' `
                -Pfad $FilePath -Aktion 'Berechnungsmodus' `
                -Status $Status -Detail $Details
        } catch { }
    }

    if (-not $logFile) { return }
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $cleanPath = Remove-LongPathPrefix $FilePath
    $line = '"{0}";"{1}";"{2}";"{3}";"{4}"' -f `
        $timestamp, `
        (Format-CsvField $FileName), `
        (Format-CsvField $cleanPath), `
        $Status, `
        (Format-CsvField $Details)
    [System.IO.File]::AppendAllText($logFile, $line + [Environment]::NewLine, $logEncoding)
}

# --- STATISTIK ---
$countChecked       = 0
$countFound         = 0
$countFixed         = 0
$countLocked        = 0
$countErrors        = 0
$countSkippedTemp   = 0
$countSkippedBinary = 0
$countSkippedNoZip  = 0

$script:ShouldStop   = $false
$script:CtrlCAsInput = $false
try {
    if ($Host.Name -eq 'ConsoleHost') {
        [Console]::TreatControlCAsInput = $true
        $script:CtrlCAsInput = $true
    }
} catch {}

function Test-AbortRequested {
    if (-not $script:CtrlCAsInput) { return }
    try {
        while ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            if ($key.Key -eq [ConsoleKey]::C -and
                (($key.Modifiers -band [ConsoleModifiers]::Control) -ne 0)) {
                if (-not $script:ShouldStop) {
                    Write-Host "`n[ABBRUCH]   Strg+C erkannt - laufende Datei wird noch fertiggestellt..." -ForegroundColor Yellow
                }
                $script:ShouldStop = $true
            }
        }
    } catch {}
}

# --- PFAD-VORBEREITUNG ---
$longPath   = Add-LongPathPrefix $rootPath
$extensions = @('.xlsx', '.xlsm', '.xltx', '.xltm')

if ($ReadOnly) {
    Write-Host "--- Excel Calculation Mode Fixer (VORSCHAU) ---" -ForegroundColor Cyan -BackgroundColor Black
} else {
    Write-Host "--- Excel Calculation Mode Fixer (LIVE) ---" -ForegroundColor Red -BackgroundColor Black
}
Write-Host "Zielpfad  : $rootPath"
Write-Host "Tempordner: $tempFolder"
Write-Host "Log-Datei : $logFile"
if (-not $ReadOnly) {
    Write-Host "WARNUNG: Aenderungen werden direkt geschrieben." -ForegroundColor Yellow
    Write-Host "         Attribute ReadOnly/Hidden/System werden bei Korrektur entfernt." -ForegroundColor Yellow
}
Write-Host "--------------------------------------------"

# --- VERWAISTE BACKUPS FRUEHERER LAEUFE ---
# Das Skript legt vor dem Zurueckkopieren '<datei>.bak_<id>' an und loescht
# es bei Erfolg. Nach einem harten Abbruch bleiben diese Reste liegen.
# Geloescht wird nur, wenn das Original existiert, nicht leer und das
# Backup aelter als $script:BackupCleanupMinAgeHours ist.
if (-not $SkipBackupCleanup) {
    Write-Host ""
    Write-Host "Suche zurueckgebliebene Backups frueherer Laeufe ..." -ForegroundColor Cyan
    $bakDeleted = 0; $bakKept = 0; $bakOrphans = 0
    $bakCutoff  = (Get-Date).AddHours(-$script:BackupCleanupMinAgeHours)

    # Eigener Durchlauf: Get-FilesRecursiveSafe filtert auf Endungen und
    # taugt fuer das '*.bak_<id>'-Muster nicht.
    $bakStack = New-Object System.Collections.Generic.Stack[string]
    $bakStack.Push($longPath)
    while ($bakStack.Count -gt 0) {
        $d = $bakStack.Pop()
        try {
            foreach ($f in [System.IO.Directory]::EnumerateFiles($d, "*.bak_*")) {
                try {
                    $orig = $f -replace '\.bak_[0-9a-fA-F]+$', ''
                    if ($orig -eq $f) { continue }
                    if ($orig -notmatch '\.(xlsx|xlsm|xltx|xltm)$') { continue }
                    $fi = New-Object System.IO.FileInfo $f
                    if (-not $fi.Exists) { continue }
                    $origOk = $false
                    if ([System.IO.File]::Exists($orig)) {
                        try { $origOk = (New-Object System.IO.FileInfo $orig).Length -gt 0 } catch { }
                    }
                    if (-not $origOk) {
                        $bakOrphans++
                        Write-Host "[BACKUP]    Ohne intaktes Original - NICHT geloescht: $(Remove-LongPathPrefix $f)" -ForegroundColor Yellow
                        Write-Log -FileName (Split-Path $f -Leaf) -FilePath $f -Status "WARNUNG" -Details "Backup ohne intaktes Original - bitte pruefen"
                        continue
                    }
                    if ($fi.LastWriteTime -gt $bakCutoff) { $bakKept++; continue }
                    if (Confirm-Write (Remove-LongPathPrefix $f) 'Verwaistes Backup loeschen') {
                        try { [System.IO.File]::Delete($f); $bakDeleted++ } catch { }
                    } else { $bakKept++ }
                } catch { }
            }
        } catch { }
        try {
            foreach ($sub in [System.IO.Directory]::EnumerateDirectories($d)) {
                $leaf = [System.IO.Path]::GetFileName($sub.TrimEnd('\'))
                if ($script:ExcludeDirNames -contains $leaf) { continue }
                try {
                    $a = [System.IO.File]::GetAttributes($sub)
                    if ($a -band [System.IO.FileAttributes]::ReparsePoint) { continue }
                } catch { continue }
                $bakStack.Push($sub)
            }
        } catch { }
    }
    if ($bakDeleted -or $bakKept -or $bakOrphans) {
        Write-Host "  Backups: $bakDeleted entfernt, $bakKept behalten, $bakOrphans ohne Original" -ForegroundColor DarkGray
    } else {
        Write-Host "  Keine Backup-Reste gefunden." -ForegroundColor DarkGray
    }
}

# --- DATEISUCHE ---
Write-Host ""
Write-Host "Suche Excel-Dateien (rekursiv) ..." -ForegroundColor Cyan
Write-Host "Hinweis: Bei Netzlaufwerken (UNC) und sehr vielen Dateien" -ForegroundColor DarkGray
Write-Host "(>100.000) kann die Suche 10-30 Minuten dauern." -ForegroundColor DarkGray
try {
    $files = @(Get-FilesRecursiveSafe -RootLongPath $longPath -Extensions $extensions)
} catch {
    Write-Host "Schnelle Enumeration fehlgeschlagen, versuche Standardpfad..." -ForegroundColor Gray
    $files = Get-ChildItem -LiteralPath $rootPath -File -Recurse -ErrorAction SilentlyContinue |
             Where-Object { $extensions -contains $_.Extension.ToLowerInvariant() } |
             Select-Object -ExpandProperty FullName
}

Write-Host "$(@($files).Count) Excel-Dateien gefunden." -ForegroundColor Cyan

# --- HAUPTSCHLEIFE ---
foreach ($file in $files) {
    Test-AbortRequested
    if ($script:ShouldStop) { break }

    try {
        $fileName = Split-Path $file -Leaf
    } catch {
        $countErrors++
        $msg = $_.Exception.Message
        Write-Host "[FEHLER]    Ungueltiger Pfad: $msg" -ForegroundColor Red
        Write-Log -FileName "" -FilePath $file -Status "FEHLER" -Details "Ungueltiger Pfad: $msg"
        continue
    }

    # Excel-Lock-Dateien (~$datei.xlsx)
    if ($fileName -like '~$*') {
        $countSkippedTemp++
        Write-Host "[ SKIP ]    $fileName (temporaere Excel-Datei)" -ForegroundColor DarkGray
        Write-Log -FileName $fileName -FilePath $file -Status "SKIP" -Details "Temporaere Excel-Lock-Datei uebersprungen"
        continue
    }

    # .xlsb Binary-Format wird hier nicht unterstuetzt (kein ZIP-Container).
    # Defensiv: koennte indirekt durch Endungs-Filter auftauchen, falls jemand die Liste erweitert.
    if ([System.IO.Path]::GetExtension($fileName).ToLowerInvariant() -eq '.xlsb') {
        $countSkippedBinary++
        Write-Host "[ SKIP ]    $fileName (Binary-Format .xlsb)" -ForegroundColor DarkGray
        Write-Log -FileName $fileName -FilePath $file -Status "SKIP" -Details "Binary-Format .xlsb wird nicht unterstuetzt"
        continue
    }

    $countChecked++

    if ($countChecked % 100 -eq 0) {
        Write-Host "[FORTSCHRITT] $countChecked Dateien verarbeitet (Gefunden: $countFound, Korrigiert: $countFixed)..." -ForegroundColor Cyan
    }

    $stream  = $null
    $archive = $null
    $reader  = $null
    $writer  = $null

    try {
        $baseName  = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
        $extension = [System.IO.Path]::GetExtension($fileName)
        $uniqueId  = [guid]::NewGuid().ToString("N").Substring(0, 8)
        # Basisname kuerzen, damit Tempfile-Komponente nicht NTFS-Limit (255) reisst
        if ($baseName.Length -gt 180) { $baseName = $baseName.Substring(0, 180) }
        $tempFile     = Join-Path $tempFolder "${baseName}_${uniqueId}${extension}"
        $tempFileLong = Add-LongPathPrefix $tempFile
    } catch {
        $countErrors++
        $msg = $_.Exception.Message
        Write-Host "[FEHLER]    $fileName : Pfad-Operation fehlgeschlagen: $msg" -ForegroundColor Red
        Write-Log -FileName $fileName -FilePath $file -Status "FEHLER" -Details "Pfad-Operation: $msg"
        continue
    }

    try {
        # Schritt 1: Magic-Byte-Check (PK = ZIP)
        $zipCheck = Test-IsZipFile $file
        if ($zipCheck -eq $false) {
            $countSkippedNoZip++
            Write-Host "[ SKIP ]    $fileName (keine gueltige ZIP-Signatur)" -ForegroundColor DarkYellow
            Write-Log -FileName $fileName -FilePath $file -Status "SKIP" -Details "Keine gueltige ZIP/OOXML-Signatur"
            continue
        }
        # Bei $null (z.B. gesperrt) regulaer fortfahren - Copy meldet ggf. BLOCKIERT.

        # Schritt 2: In Tempordner kopieren und auf Kopie alle blockierenden Attribute entfernen
        [System.IO.File]::Copy($file, $tempFileLong, $true)
        [void](Remove-RestrictiveAttributes $tempFileLong)

        $stream         = $null
        $openRetry      = 0
        $maxOpenRetries = 4
        $openDelays     = @(100, 200, 400, 800)
        while ($null -eq $stream -and $openRetry -lt $maxOpenRetries) {
            try {
                $stream = New-Object System.IO.FileStream(
                    $tempFileLong,
                    [System.IO.FileMode]::Open,
                    [System.IO.FileAccess]::ReadWrite,
                    [System.IO.FileShare]::None
                )
            } catch [System.IO.IOException] {
                $openRetry++
                if ($openRetry -ge $maxOpenRetries) { throw }
                Start-Sleep -Milliseconds $openDelays[$openRetry - 1]
            }
        }

        $archive = New-Object System.IO.Compression.ZipArchive(
            $stream,
            [System.IO.Compression.ZipArchiveMode]::Update,
            $true
        )
        $entry = $archive.GetEntry("xl/workbook.xml")

        if ($null -ne $entry) {
            $utf8NoBom = New-Object System.Text.UTF8Encoding($false)

            $reader     = New-Object System.IO.StreamReader($entry.Open(), $utf8NoBom)
            $xmlContent = $reader.ReadToEnd()
            $reader.Close()
            $reader     = $null

            # Auswertung ueber den XML-Baum statt per regulaerem Ausdruck.
            # Der frueher verwendete -match lief gegen die GESAMTE
            # workbook.xml: ein gleichnamiges Attribut an einer anderen
            # Stelle (etwa in einer definedName-Formel oder einer
            # Erweiterungsliste) loeste dadurch eine Korrektur aus und
            # wurde vom -replace anschliessend mitentfernt.
            $xmlDoc   = $null
            $calcPr   = $null
            $xmlUsable = $false
            try {
                $xmlDoc = New-Object System.Xml.XmlDocument
                $xmlDoc.PreserveWhitespace = $true
                $xmlDoc.LoadXml($xmlContent)
                # local-name(): workbook.xml traegt einen Standard-Namensraum,
                # ein einfaches SelectSingleNode('calcPr') findet nichts.
                $calcPr = $xmlDoc.DocumentElement.SelectSingleNode("*[local-name()='calcPr']")
                $xmlUsable = $true
            } catch {
                Write-Host "[ SKIP ]    $fileName (workbook.xml nicht lesbar: $($_.Exception.Message))" -ForegroundColor DarkYellow
                Write-Log -FileName $fileName -FilePath $file -Status "SKIP" -Details "workbook.xml nicht als XML lesbar: $($_.Exception.Message)"
            }

            $hasManualCalc        = $false
            $hasCalcOnSaveZero    = $false
            $hasCalcCompletedZero = $false
            if ($xmlUsable -and $null -ne $calcPr) {
                $hasManualCalc        = ($calcPr.GetAttribute('calcMode')      -ieq 'manual')
                $hasCalcOnSaveZero    = ($calcPr.GetAttribute('calcOnSave')    -eq  '0')
                $hasCalcCompletedZero = ($calcPr.GetAttribute('calcCompleted') -eq  '0')
            }
            $needsFix = $hasManualCalc -or $hasCalcOnSaveZero -or $hasCalcCompletedZero

            if (-not $xmlUsable) {
                # Bereits oben als SKIP protokolliert - hier nichts weiter tun,
                # sonst meldete das Skript die unlesbare Datei zusaetzlich als
                # "Bereits automatisch".
                $countSkippedNoZip++
            } elseif ($needsFix) {
                $countFound++

                # Zentrale Freigabe: ab hier wird das Original angefasst.
                $darfSchreiben = Confirm-Write `
                    -Target (Remove-LongPathPrefix $file) `
                    -Action "Berechnungsmodus auf automatisch zuruecksetzen"

                if ($darfSchreiben) {
                    # Zeitstempel VOR jeder Aenderung sichern.
                    $origTimes = Get-FileTimestamps (Add-LongPathPrefix $file)

                    # Geaendert wird ausschliesslich der calcPr-Tag im
                    # Originaltext. Bewusst NICHT ueber $xmlDoc.OuterXml:
                    # das serialisiert die komplette workbook.xml neu und
                    # schreibt dabei Entities in unbeteiligten Textknoten um
                    # (&quot; wird zu "), was den Eingriff unnoetig gross
                    # macht. Die Erkennung oben laeuft weiterhin ueber den
                    # XML-Baum, der Schnitt hier ist auf den einen Tag
                    # begrenzt - Attributwerte koennen weder '<' noch '>'
                    # roh enthalten, das Muster ist also eindeutig.
                    # Excel setzt fuer alle drei Attribute den gewuenschten
                    # Wert als Vorgabe, wenn sie fehlen: calcMode=auto,
                    # calcOnSave=1, calcCompleted=1. 'autoNoTable' bleibt
                    # erhalten, weil oben nur exakt 'manual' geprueft wird.
                    $tagMatch = [regex]::Match($xmlContent, '<calcPr\b[^>]*?/?>')
                    if (-not $tagMatch.Success) {
                        throw "calcPr-Element im Originaltext nicht auffindbar"
                    }
                    $newTag = $tagMatch.Value
                    $attrsToDrop = New-Object System.Collections.Generic.List[string]
                    if ($hasManualCalc)        { $attrsToDrop.Add('calcMode') }
                    if ($hasCalcOnSaveZero)    { $attrsToDrop.Add('calcOnSave') }
                    if ($hasCalcCompletedZero) { $attrsToDrop.Add('calcCompleted') }
                    foreach ($attrName in $attrsToDrop) {
                        $newTag = $newTag -replace ("(?i)\s+" + $attrName + "\s*=\s*(""[^""]*""|'[^']*')"), ''
                    }
                    $newContent = $xmlContent.Substring(0, $tagMatch.Index) +
                                  $newTag +
                                  $xmlContent.Substring($tagMatch.Index + $tagMatch.Length)

                    $entry.Delete()
                    $newEntry = $archive.CreateEntry("xl/workbook.xml")

                    $writer = New-Object System.IO.StreamWriter($newEntry.Open(), $utf8NoBom)
                    $writer.Write($newContent)
                    $writer.Flush()
                    $writer.Close()
                    $writer = $null

                    # Schritt 3: Archiv/Stream schliessen vor dem Zurueckkopieren
                    $archive.Dispose(); $archive = $null
                    $stream.Close(); $stream.Dispose(); $stream = $null

                    # Schritt 4: Quellattribute (ReadOnly/Hidden/System) entfernen - bleiben dauerhaft entfernt
                    $longFile     = Add-LongPathPrefix $file
                    $removedAttrs = Remove-RestrictiveAttributes $longFile

                    # Schritt 5: Zurueckkopieren mit Backup-Schutz und Retry.
                    # File.Copy(...,$true) trunkiert das Ziel sofort auf
                    # 0 Bytes - bricht der Kopiervorgang ab (Netzwerk-Drop,
                    # AV-Scanner), waere das Original ohne Backup zerstoert
                    # (die Temp-Kopie wird im finally-Block geloescht).
                    $backupLong = $null
                    if ([System.IO.File]::Exists($longFile)) {
                        $backupLong = Add-LongPathPrefix ((Remove-LongPathPrefix $file) + ".bak_" + $uniqueId)
                        [System.IO.File]::Copy($longFile, $backupLong, $true)
                    }
                    $copySuccess = $false
                    $copyRetry   = 0
                    $maxRetries  = 3
                    while (-not $copySuccess -and $copyRetry -lt $maxRetries) {
                        try {
                            [System.IO.File]::Copy($tempFileLong, $longFile, $true)
                            $copySuccess = $true
                        } catch {
                            $copyRetry++
                            if ($copyRetry -lt $maxRetries) {
                                Write-Host "[RETRY]     $fileName - Zurueckkopieren fehlgeschlagen, Versuch $($copyRetry + 1)/$maxRetries..." -ForegroundColor DarkYellow
                                Start-Sleep -Seconds $copyRetry
                            } else {
                                if ($null -ne $backupLong -and [System.IO.File]::Exists($backupLong)) {
                                    try {
                                        [System.IO.File]::Copy($backupLong, $longFile, $true)
                                        [System.IO.File]::Delete($backupLong)
                                        Write-Host "[RESTORE]   $fileName - Original aus Backup wiederhergestellt" -ForegroundColor DarkYellow
                                        Write-Log -FileName $fileName -FilePath $file -Status "RESTORE" -Details "Zurueckkopieren fehlgeschlagen - Original aus Backup wiederhergestellt"
                                    } catch {
                                        $bakDisplay = Remove-LongPathPrefix $backupLong
                                        Write-Host "[WARNUNG]   $fileName - Backup-Restore fehlgeschlagen, Backup bleibt: $bakDisplay" -ForegroundColor Red
                                        Write-Log -FileName $fileName -FilePath $file -Status "WARNUNG" -Details "Backup-Restore fehlgeschlagen - Backup bleibt erhalten: $bakDisplay"
                                    }
                                }
                                throw $_
                            }
                        }
                    }
                    if ($null -ne $backupLong -and [System.IO.File]::Exists($backupLong)) {
                        try {
                            [System.IO.File]::Delete($backupLong)
                        } catch {
                            $bakDisplay = Remove-LongPathPrefix $backupLong
                            Write-Host "[WARNUNG]   Backup nicht loeschbar: $bakDisplay" -ForegroundColor DarkYellow
                            Write-Log -FileName $fileName -FilePath $file -Status "WARNUNG" -Details "Backup nicht loeschbar: $bakDisplay"
                        }
                    }

                    # Hidden/System zurueckgeben - ReadOnly bleibt bewusst
                    # entfernt (sonst blockiert es jeden kuenftigen Lauf).
                    $restoredAttrs = Restore-FileAttributes -FilePath $longFile -Attributes $removedAttrs

                    # Zeitstempel des Originals zurueckschreiben.
                    $timesOk = Restore-FileTimestamps -FilePath $longFile -Timestamps $origTimes

                    $countFixed++
                    $detailParts = New-Object System.Collections.Generic.List[string]
                    if ($hasManualCalc)        { $detailParts.Add("calcMode=manual entfernt") }
                    if ($hasCalcOnSaveZero)    { $detailParts.Add("calcOnSave=0 entfernt") }
                    if ($hasCalcCompletedZero) { $detailParts.Add("calcCompleted=0 entfernt") }
                    if ($removedAttrs.Count -gt 0) {
                        $detailParts.Add("Attribute temporaer entfernt: " + ($removedAttrs -join ','))
                    }
                    if ($restoredAttrs.Count -gt 0) {
                        $detailParts.Add("wiederhergestellt: " + ($restoredAttrs -join ','))
                    }
                    if ($null -ne $origTimes) {
                        $detailParts.Add($(if ($timesOk) { "Zeitstempel erhalten" }
                                           else          { "Zeitstempel NICHT wiederherstellbar" }))
                    }
                    $detail = $detailParts -join '; '
                    Write-Host "[FIXED]     $fileName" -ForegroundColor Green
                    Write-Log -FileName $fileName -FilePath $file -Status "FIXED" -Details $detail
                } else {
                    # Vorschau/-WhatIf: nichts anfassen, nur berichten.
                    $detailParts = New-Object System.Collections.Generic.List[string]
                    if ($hasManualCalc)        { $detailParts.Add("calcMode=manual") }
                    if ($hasCalcOnSaveZero)    { $detailParts.Add("calcOnSave=0") }
                    if ($hasCalcCompletedZero) { $detailParts.Add("calcCompleted=0") }
                    $detail = "VORSCHAU - vorhanden: " + ($detailParts -join ', ')
                    Write-Host "[GEFUNDEN]  $fileName (VORSCHAU - keine Aenderung)" -ForegroundColor Yellow
                    Write-Log -FileName $fileName -FilePath $file -Status "GEFUNDEN" -Details $detail
                }
            } else {
                Write-Host "[ OK ]      $fileName" -ForegroundColor DarkGray
                Write-Log -FileName $fileName -FilePath $file -Status "OK" -Details "Bereits automatisch"
            }
        } else {
            Write-Host "[ SKIP ]    $fileName (kein workbook.xml gefunden)" -ForegroundColor DarkGray
            Write-Log -FileName $fileName -FilePath $file -Status "SKIP" -Details "xl/workbook.xml nicht vorhanden"
        }

    } catch {
        $msg = $_.Exception.Message
        if ($msg -like "*verwendet wird*" -or
            $msg -like "*because it is being used*" -or
            $msg -like "*being used by another process*") {
            Write-Host "[BLOCKIERT] $fileName (Datei ist geoeffnet)" -ForegroundColor DarkYellow
            Write-Log -FileName $fileName -FilePath $file -Status "BLOCKIERT" -Details "Datei durch anderen Prozess gesperrt"
            $countLocked++
        } elseif ($msg -like "*Access*denied*" -or
                  $msg -like "*Zugriff*verweigert*" -or
                  $msg -like "*UnauthorizedAccess*") {
            Write-Host "[ZUGRIFF]   $fileName (Zugriff verweigert)" -ForegroundColor Red
            Write-Log -FileName $fileName -FilePath $file -Status "ZUGRIFF" -Details "Zugriff verweigert: $msg"
            $countErrors++
        } else {
            Write-Host "[FEHLER]    $fileName : $msg" -ForegroundColor Red
            Write-Log -FileName $fileName -FilePath $file -Status "FEHLER" -Details $msg
            $countErrors++
        }
    } finally {
        if ($null -ne $reader)  { try { $reader.Close() }  catch { } }
        if ($null -ne $writer)  { try { $writer.Close() }  catch { } }
        if ($null -ne $archive) { try { $archive.Dispose() } catch { } }
        if ($null -ne $stream)  { try { $stream.Close(); $stream.Dispose() } catch { } }

        # [System.IO.File] statt Test-Path/Remove-Item: die Provider-Cmdlets
        # von PowerShell 5.1 kommen mit '\\?\'-Praefixen nicht zuverlaessig
        # zurecht - die Temp-Datei blieb dann unbemerkt liegen.
        if ([System.IO.File]::Exists($tempFileLong)) {
            $rmRetry = 0
            while ([System.IO.File]::Exists($tempFileLong) -and ($rmRetry -lt 3)) {
                try {
                    [System.IO.File]::Delete($tempFileLong)
                } catch {
                    $rmRetry++
                    Start-Sleep -Milliseconds 500
                }
            }
            if ([System.IO.File]::Exists($tempFileLong)) {
                Write-Host "[WARNUNG]   Temp-Datei konnte nicht geloescht werden: $tempFile" -ForegroundColor DarkYellow
                Write-Log -FileName $fileName -FilePath $tempFile -Status "WARNUNG" -Details "Temp-Datei nicht loeschbar"
            }
        }
    }
}

# --- TEMP-ORDNER AUFRAEUMEN ---
if ((Test-Path -LiteralPath $tempFolder) -and (@(Get-ChildItem -LiteralPath $tempFolder -Force).Count -eq 0)) {
    # -WhatIf:$false, weil der Ordner oben bewusst auch im Probelauf angelegt
    # wird - ohne das bliebe nach jedem -WhatIf-Lauf ein leerer Ordner liegen.
    Remove-Item -LiteralPath $tempFolder -Force -ErrorAction SilentlyContinue -WhatIf:$false
} elseif (Test-Path -LiteralPath $tempFolder) {
    Write-Host "[WARNUNG]   Temp-Ordner nicht leer, bitte manuell pruefen: $tempFolder" -ForegroundColor DarkYellow
    Write-Log -FileName "" -FilePath $tempFolder -Status "WARNUNG" -Details "Temp-Ordner nicht leer - manuelle Pruefung noetig"
}

# --- ABSCHLUSSBERICHT ---
$totalSkipped = $countSkippedTemp + $countSkippedBinary + $countSkippedNoZip
Write-Host "--------------------------------------------" -ForegroundColor Cyan
Write-Host "Gepruefte Dateien gesamt : $countChecked"
Write-Host "Manuelle Modi gefunden   : $countFound"
if ($ReadOnly) {
    Write-Host "Wuerden korrigiert werden: $countFound" -ForegroundColor Yellow
} else {
    Write-Host "Erfolgreich korrigiert   : $countFixed" -ForegroundColor Green
}
Write-Host "Gesperrt (in Benutzung)  : $countLocked" -ForegroundColor Yellow
Write-Host "Uebersprungen gesamt     : $totalSkipped" -ForegroundColor DarkGray
Write-Host "  - Excel-Lock-Dateien   : $countSkippedTemp" -ForegroundColor DarkGray
Write-Host "  - .xlsb (Binary)       : $countSkippedBinary" -ForegroundColor DarkGray
Write-Host "  - keine ZIP-Signatur   : $countSkippedNoZip" -ForegroundColor DarkGray
Write-Host "Fehler                   : $countErrors" -ForegroundColor $(if ($countErrors -gt 0) { "Red" } else { "DarkGray" })
Write-Host "Log-Datei                : $logFile" -ForegroundColor Cyan
Write-Host "Vorgang abgeschlossen." -ForegroundColor Cyan
try { if ($script:CtrlCAsInput) { [Console]::TreatControlCAsInput = $false } } catch {}
Wait-AnyKey
