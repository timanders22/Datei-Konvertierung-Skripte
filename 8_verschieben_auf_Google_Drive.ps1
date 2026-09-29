
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

# ==================================================================
# VERSCHIEBEN VON VERZEICHNISSEN UND DATEIEN AUF GOOGLE DRIVE (G:)
# ------------------------------------------------------------------
# Stand   : 12.06.2026
#
# DATEI-ENCODING:
#   Diese .ps1 muss als "UTF-8 mit BOM" gespeichert sein, damit
#   PowerShell 5.1 deutsche Umlaute korrekt interpretiert.
#
# AUSFUEHRUNGSUMGEBUNG:
#   - Mindestens PowerShell 5.0 (Win10/11 Standard: 5.1)
#   - Long-Path-Support: HKLM\SYSTEM\CurrentControlSet\Control\
#     FileSystem\LongPathsEnabled = 1 (empfohlen)
#   - Methode rclone: rclone.exe >= 1.62 im Skript-Ordner oder PATH
#
# EXE-KOMPILIERUNG (PS2EXE):
#   Im PowerShell-Fenster:
#     Install-Module -Name ps2exe -Scope CurrentUser -Force
#     Import-Module ps2exe
#     Invoke-PS2EXE `
#         -inputFile  .\8_verschieben_auf_Google_Drive.ps1 `
#         -outputFile .\8_verschieben_auf_Google_Drive.exe `
#         -iconFile   .\powershell_icon.ico `
#         -title      "GDrive Migration" `
#         -product    "GDrive Migration" `
#         -requireAdmin
#
#   Hinweise:
#   - KEIN -noConsole verwenden, da Read-Host benoetigt wird.
#   - -requireAdmin laedt Q:/R: zuverlaessig (UAC-Mapping).
#     Alternativ ohne -requireAdmin und stattdessen UNC-Pfade nutzen.
# ==================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8

$Utf8Bom = New-Object System.Text.UTF8Encoding($true)

# ==================================================================
# KONFIGURATION
# ==================================================================
# Praesets fuer die Quellauswahl. Leere Werte werden im Menue automatisch
# ausgeblendet - so laesst sich das Skript ohne UNC-Basen betreiben.
$UncBaseQ            = "\\server\dfs"
$UncBaseR            = "\\server\home"
$RcloneRemoteQ       = "ProjektDrive:"
$RcloneRemoteMyDrive = "GDriveEnterprise:"

$RetryMaxAttempts        = 5
$RetryInitialDelaySec    = 1
$RetryMaxDelaySec        = 30

# Verzeichnisnamen, die UEBERALL ausgeschlossen werden: beim Kopieren
# (Robocopy /XD bzw. rclone --exclude), bei Pruefung/Verifikation und
# beim Loeschen. Muss fuer alle Pfade einheitlich sein, sonst meldet
# die Verifikation Dateien als "fehlend", die absichtlich nie kopiert
# wurden. '~snapshot'/'.snapshot' sind NAS-Schattenkopien (NetApp & Co.)
# auf R:\ -- read-only-Duplikate, die nicht migriert werden duerfen.
$script:ExcludedDirNames = @('$RECYCLE.BIN', 'System Volume Information', '~snapshot', '.snapshot')
$ExcludedDirNames        = $script:ExcludedDirNames

# Strg+C-Behandlung: Ein PowerShell-Scriptblock als CancelKeyPress-Handler
# laeuft beim Event auf einem fremden Thread ohne Runspace und schlaegt
# dort zur Laufzeit fehl ("There is no Runspace available ...") -- das
# Flag wuerde nie gesetzt. Daher setzt ein kleiner C#-Handler (Add-Type)
# ein statisches Flag, das die Schleifen ueber Test-ShouldStop pollen.
# Erstes Strg+C  = sanfter Abbruch (laufende Operation wird beendet),
# zweites Strg+C = sofortiges Beenden.
$script:ShouldStop = $false
$script:AbortHandlerActive = $false
try {
    [Console]::TreatControlCAsInput = $false
    if (-not ('GDriveMigrationAbort' -as [type])) {
        Add-Type -ErrorAction Stop -TypeDefinition @'
using System;
public static class GDriveMigrationAbort
{
    public static volatile bool Requested = false;
    private static bool _registered = false;
    public static void Register()
    {
        if (_registered) { return; }
        _registered = true;
        Console.CancelKeyPress += delegate(object sender, ConsoleCancelEventArgs e)
        {
            if (Requested) { e.Cancel = false; }  // zweites Strg+C: hart beenden
            else           { e.Cancel = true; Requested = true; }
        };
    }
}
'@
    }
    [GDriveMigrationAbort]::Register()
    $script:AbortHandlerActive = $true
} catch {
    Write-Warning "Strg+C-Abfangen nicht verfuegbar -- Strg+C beendet das Skript sofort. ($_)"
}

function Test-ShouldStop {
    if (-not $script:ShouldStop -and $script:AbortHandlerActive -and [GDriveMigrationAbort]::Requested) {
        $script:ShouldStop = $true
        Write-Host "`n⚠️  ABBRUCH angefordert – laufende Operation wird noch fertiggestellt... (erneutes Strg+C beendet sofort)" -ForegroundColor Yellow
    }
    return $script:ShouldStop
}

# ==================================================================
# HILFSFUNKTIONEN -- PFAD & SYSTEM
# ==================================================================

function Get-ScriptDirectory {
    if ($PSScriptRoot) { return $PSScriptRoot }
    if ($MyInvocation.MyCommand.Path) { return Split-Path -Parent $MyInvocation.MyCommand.Path }
    if ($PSCommandPath) { return Split-Path -Parent $PSCommandPath }
    try {
        $entry = [System.Reflection.Assembly]::GetEntryAssembly()
        if ($entry -and $entry.Location) { return Split-Path -Parent $entry.Location }
    } catch { }
    try {
        $proc = [System.Diagnostics.Process]::GetCurrentProcess()
        if ($proc.MainModule -and $proc.MainModule.FileName) {
            return Split-Path -Parent $proc.MainModule.FileName
        }
    } catch { }
    return (Get-Location).Path
}

function Add-LongPathPrefix {
    param([string]$Path)
    if ([string]::IsNullOrEmpty($Path))  { return $Path }
    if ($Path.StartsWith('\\?\'))         { return $Path }
    if ($Path -match '^[A-Za-z]:$')       { return '\\?\' + $Path + '\' }
    if ($Path.StartsWith('\\'))           { return '\\?\UNC\' + $Path.Substring(2) }
    return '\\?\' + $Path
}

function Remove-LongPathPrefix {
    param([string]$Path)
    if ([string]::IsNullOrEmpty($Path))   { return $Path }
    if ($Path.StartsWith('\\?\UNC\'))      { return '\\' + $Path.Substring(8) }
    if ($Path.StartsWith('\\?\'))          { return $Path.Substring(4) }
    return $Path
}

function Get-UserShellFolder {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Desktop','Downloads','Documents')]
        [string]$Name
    )
    $regKey    = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
    $valueName = switch ($Name) {
        'Desktop'   { 'Desktop' }
        'Documents' { 'Personal' }
        default     { '{374DE290-123F-4565-9164-39C4925E467B}' }
    }

    try {
        $raw = (Get-ItemProperty -Path $regKey -Name $valueName -ErrorAction Stop).$valueName
        if (-not [string]::IsNullOrWhiteSpace($raw)) {
            $expanded = [System.Environment]::ExpandEnvironmentStrings($raw)
            if ($expanded -and (Test-Path -LiteralPath $expanded)) {
                return $expanded
            }
        }
    } catch { }

    $fallback = Join-Path $env:USERPROFILE $Name
    return $fallback
}

function Get-OptimalRcloneSettings {
    $totalGB = 8
    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        if ($cs.TotalPhysicalMemory) {
            $totalGB = [Math]::Round($cs.TotalPhysicalMemory / 1GB, 0)
        }
    } catch {
        try {
            $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
            if ($os.TotalVisibleMemorySize) {
                $totalGB = [Math]::Round($os.TotalVisibleMemorySize / 1MB, 0)
            }
        } catch { }
    }
    if ($totalGB -le 8) {
        return @{ Transfers = 3; ChunkSize = '16M'; RamGB = $totalGB }
    } elseif ($totalGB -le 16) {
        return @{ Transfers = 4; ChunkSize = '32M'; RamGB = $totalGB }
    } else {
        return @{ Transfers = 8; ChunkSize = '64M'; RamGB = $totalGB }
    }
}

function Clear-TempLeftovers {
    # Whitelist-Cleanup am Skriptende: Es werden AUSSCHLIESSLICH eigene
    # Skript-Artefakte entfernt (verwaiste Leer-Ordner aus abgebrochenen
    # Laeufen). NIEMALS pauschal leeren - Fremdprozesse legen aktive
    # Daten ohne Lock in %TEMP% ab; blindes Loeschen zerstoert sie.
    # "RoboMirrorEmpty_" deckt Altlasten frueherer Skript-Versionen ab.
    $prefixes = @('8_verschieben_auf_Google_Drive_empty_', 'RoboMirrorEmpty_')
    # Alle Orte, an denen der Leerordner fuer den /MIR-Purge angelegt
    # werden kann - muss mit der Kandidatenliste dort deckungsgleich
    # bleiben, sonst bleiben Reste liegen.
    $roots = @()
    if ($env:TEMP          -and (Test-Path -LiteralPath $env:TEMP))          { $roots += $env:TEMP }
    if ($env:LOCALAPPDATA  -and (Test-Path -LiteralPath $env:LOCALAPPDATA))  { $roots += $env:LOCALAPPDATA }
    try {
        $docs = Get-UserShellFolder -Name 'Documents'
        if ($docs -and (Test-Path -LiteralPath $docs)) { $roots += $docs }
    } catch { }
    $roots = @($roots | Select-Object -Unique)
    foreach ($root in $roots) {
        foreach ($prefix in $prefixes) {
            try {
                Get-ChildItem -LiteralPath $root -Directory -Filter "$prefix*" -ErrorAction SilentlyContinue |
                    ForEach-Object {
                        Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
                    }
            } catch { }
        }
    }
}

function Wait-ForKey {
    param([string]$Message = "Druecken Sie eine Taste, um fortzufahren ...")
    Write-Host ""
    Write-Host $Message
    try {
        if ($Host.Name -eq 'ConsoleHost' -and $Host.UI -and $Host.UI.RawUI) {
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            return
        }
    } catch { }
    $null = Read-Host
}


function Test-SourcePathSafe {
    <#
        Sicherheitsnetz vor dem Kopieren UND vor dem Loeschen.

        Vorher wurde jeder existierende Pfad akzeptiert. Wer bei Option [7]
        versehentlich "C:\" eingab, bekam am Ende den Robocopy-/MIR-Purge
        gegen das Systemlaufwerk - der loescht ein komplettes Windows,
        weil er einen leeren Ordner ueber die Quelle spiegelt. Ebenso
        gefaehrlich: das Nutzerprofil, Programmverzeichnisse, das
        Skriptverzeichnis selbst und das Google-Drive-Laufwerk (dort
        wuerde die Quelle das eigene Ziel sein).

        Rueckgabe: $null wenn unbedenklich, sonst der Ablehnungsgrund.
    #>
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return "Leerer Pfad." }
    $p = $Path.TrimEnd('\')

    # --- Laufwerkswurzeln ---
    if ($p -match '^[A-Za-z]:$') {
        $sysDrive = ($env:SystemDrive).TrimEnd('\')
        if ($p -ieq $sysDrive) {
            return "'$Path' ist das Systemlaufwerk. Eine Migration des gesamten Windows-Laufwerks ist ausgeschlossen."
        }
        try {
            $di = New-Object System.IO.DriveInfo($p + '\')
            if ($di.DriveType -eq [System.IO.DriveType]::Fixed -and
                [System.IO.Directory]::Exists((Join-Path ($p + '\') 'Windows\System32'))) {
                return "'$Path' enthaelt eine Windows-Installation."
            }
        } catch { }
    }

    # --- UNC-Server-Wurzel (\\server ohne Freigabe) ---
    if ($p -match '^\\\\[^\\]+$') {
        return "'$Path' ist nur der Servername ohne Freigabe."
    }

    # --- Geschuetzte Systemverzeichnisse ---
    $protected = @(
        $env:SystemRoot,
        $env:ProgramFiles,
        ${env:ProgramFiles(x86)},
        $env:ProgramData,
        $env:USERPROFILE,
        $env:APPDATA,
        $env:LOCALAPPDATA,
        $env:TEMP
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    # Vorher nur auf Gleichheit geprueft: 'C:\Users' (Vorfahr aller
    # Profile), 'C:\Users\<Name>\AppData' (Vorfahr von APPDATA und
    # LOCALAPPDATA) oder 'C:\Program Files\Common Files' (Nachfahr von
    # ProgramFiles) gingen durch (nachgestellt per herausgeloester
    # Funktion). Jetzt gilt:
    #   - gleich oder VORFAHR eines geschuetzten Verzeichnisses: immer
    #     gesperrt (der Loeschschritt traefe das geschuetzte Verzeichnis
    #     mit).
    #   - NACHFAHR: nur fuer die systemeigenen Verzeichnisse (Windows,
    #     Programme, ProgramData) gesperrt. Unterhalb des Nutzerprofils
    #     bleibt es erlaubt - Desktop und Downloads sind feste Praesets
    #     dieses Skripts, Dokumente-Unterordner ein ueblicher Fall.
    foreach ($prot in $protected) {
        $pr = $prot.TrimEnd('\')
        if ($p -ieq $pr) {
            return "'$Path' ist ein geschuetztes Systemverzeichnis."
        }
        if ($pr.StartsWith($p + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
            return "'$Path' enthaelt das geschuetzte Systemverzeichnis '$pr'."
        }
    }
    # ProgramW6432 ergaenzt: in einem 32-Bit-Prozess zeigt ProgramFiles auf
    # 'Program Files (x86)', das 64-Bit-Verzeichnis fehlte dann ganz.
    $systemEigen = @(
        $env:SystemRoot,
        $env:ProgramFiles,
        ${env:ProgramFiles(x86)},
        $env:ProgramW6432,
        $env:ProgramData
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    foreach ($prot in $systemEigen) {
        $pr = $prot.TrimEnd('\')
        if (($p + '\').StartsWith($pr + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
            return "'$Path' liegt im geschuetzten Systemverzeichnis '$pr'."
        }
    }

    # --- Google-Drive-Laufwerk als Quelle ---
    if ($p -match '^[Gg]:') {
        return "'$Path' liegt auf dem Google-Drive-Laufwerk G: - Quelle und Ziel wuerden zusammenfallen."
    }

    # --- Skriptverzeichnis (das Skript wuerde sich selbst loeschen) ---
    try {
        $sd = (Get-ScriptDirectory).TrimEnd('\')
        if ($sd -and (($p -ieq $sd) -or $sd.StartsWith($p + '\', [System.StringComparison]::OrdinalIgnoreCase))) {
            return "'$Path' enthaelt dieses Skript selbst."
        }
    } catch { }

    return $null
}

function Test-PathInside {
    <# Liegt $Child in $Parent (oder ist identisch)? #>
    param([string]$Child, [string]$Parent)
    if ([string]::IsNullOrWhiteSpace($Child) -or [string]::IsNullOrWhiteSpace($Parent)) { return $false }
    $c = $Child.TrimEnd('\')  + '\'
    $pa = $Parent.TrimEnd('\') + '\'
    return $c.StartsWith($pa, [System.StringComparison]::OrdinalIgnoreCase)
}

function Stop-Script {
    <#
        Zentraler Ausstieg. Vorher wurde an 27 Stellen 'Wait-ForKey; exit'
        geschrieben - dabei blieb Clear-TempLeftovers ungenutzt, sodass
        verwaiste Leerordner frueherer Abbrueche liegenblieben.
    #>
    param([int]$Code = 0)
    try { Clear-TempLeftovers } catch { }
    Wait-ForKey
    exit $Code
}

function Stop-WennAbgebrochen {
    <#
        Strg+C-Pruefpunkt auf oberster Ebene. Die Schleifen der Phasen vor
        dem Kopieren (Namensbereinigung, Kompatibilitaetspruefung,
        Zeitstempel-Export) brechen bei Strg+C nur ab und liefern ihr
        Teilergebnis zurueck - bei einem Abbruch gleich zu Beginn also 0
        Probleme bzw. 0 Eintraege. Der Ablauf meldete dann gruen und startete
        danach den vollstaendigen Robocopy-Lauf, weil das Flag erst in der
        Verifikation wieder abgefragt wurde. Hier ist nach jeder Phase
        Schluss; die Quelle bleibt erhalten.
    #>
    param(
        [string]$Phase,
        [System.Text.Encoding]$Encoding = $null
    )
    if (-not (Test-ShouldStop)) { return }
    Write-Host ""
    Write-Warning "Abbruch durch Anwender (Strg+C) waehrend: $Phase"
    Write-Warning "Das Ergebnis dieser Phase ist unvollstaendig und wird nicht verwendet."
    Write-Warning "Es wird nichts weiter kopiert und nichts geloescht - die Quelle bleibt erhalten."
    if ($LogFile) {
        Add-LogLine -Path $LogFile -Message "[$(Get-Date)] ABBRUCH (Strg+C) waehrend: $Phase. Nichts geloescht, Quelle bleibt erhalten." `
                    -Encoding $Encoding
    }
    Stop-Script
}

# ==================================================================
# HILFSFUNKTIONEN -- I/O & RETRY (AV-SCANNER-RESILIENZ)
# ==================================================================

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory)] [scriptblock]$ScriptBlock,
        [int]$MaxAttempts        = $script:RetryMaxAttempts,
        [int]$InitialDelaySeconds = $script:RetryInitialDelaySec,
        [int]$MaxDelaySeconds    = $script:RetryMaxDelaySec,
        [string]$OperationName   = "I/O-Operation"
    )
    $attempt = 0
    $delay   = $InitialDelaySeconds
    while ($true) {
        $attempt++
        try {
            return & $ScriptBlock
        } catch [System.IO.IOException] {
            if ($attempt -ge $MaxAttempts) { throw }
            Start-Sleep -Seconds $delay
            $delay = [Math]::Min($delay * 2, $MaxDelaySeconds)
        } catch [System.UnauthorizedAccessException] {
            if ($attempt -ge $MaxAttempts) { throw }
            Start-Sleep -Seconds $delay
            $delay = [Math]::Min($delay * 2, $MaxDelaySeconds)
        } catch {
            throw
        }
    }
}

function Test-FileIsLocked {
    # Parameter heisst -Path wie in 2a, 2b und 2c. Vorher -LongPath:
    # derselbe Zweck unter anderem Namen, was den Vergleich zwischen den
    # Skripten unnoetig erschwerte.
    param([string]$Path)
    try {
        $fs = [System.IO.File]::Open($Path,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::None)
        $fs.Close()
        return $false
    } catch {
        return $true
    }
}

function Wait-FileAvailable {
    # Ebenfalls -Path statt -LongPath (siehe Test-FileIsLocked).
    param(
        [string]$Path,
        [int]$MaxWaitSeconds = 30
    )
    $waited = 0
    $delay  = 1
    while (Test-FileIsLocked -Path $Path) {
        if ($waited -ge $MaxWaitSeconds) { return $false }
        Start-Sleep -Seconds $delay
        $waited += $delay
        $delay  = [Math]::Min($delay * 2, 8)
    }
    return $true
}

function Test-IsReparsePoint {
    param([System.IO.FileSystemInfo]$Item)
    try {
        return (($Item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -eq [System.IO.FileAttributes]::ReparsePoint)
    } catch { return $false }
}

function Add-LogLine {
    param(
        [string]$Path,
        [string]$Message,
        [System.Text.Encoding]$Encoding
    )
    if ($null -eq $Encoding) { $Encoding = $Utf8Bom }
    $sw = $null
    try {
        $sw = [System.IO.StreamWriter]::new($Path, $true, $Encoding)
        $sw.WriteLine($Message)
    } catch {
        Write-Warning "Log-Eintrag fehlgeschlagen: $_"
    } finally {
        if ($sw) { $sw.Close() }
    }
}

# ==================================================================
# HILFSFUNKTIONEN -- STREAMING-ENUMERATION (RAM-SCHONEND)
# ==================================================================

function Get-FilesStreaming {
    param(
        [Parameter(Mandatory)] [string]$RootPath,
        [bool]$SkipReparsePoints = $true,
        [switch]$IncludeFiles,
        [switch]$IncludeDirectories,
        [System.Collections.Generic.List[string]]$ReparsePointLog = $null
    )
    if (-not $IncludeFiles -and -not $IncludeDirectories) {
        $IncludeFiles       = $true
        $IncludeDirectories = $true
    }

    $stack = New-Object System.Collections.Generic.Stack[string]
    $longRoot = Add-LongPathPrefix $RootPath.TrimEnd('\')
    $stack.Push($longRoot)

    while ($stack.Count -gt 0) {
        if (Test-ShouldStop) { break }
        $current = $stack.Pop()
        $dirInfo = $null
        try {
            $dirInfo = New-Object System.IO.DirectoryInfo($current)
            # Die Reparse-Pruefung gilt nicht fuer die Wurzel selbst: Der
            # Nutzer hat sie explizit gewaehlt (DFS-Links/Junctions sind als
            # Quelle legitim), und auch Robocopy /XJ kopiert den Inhalt
            # einer als Quelle angegebenen Junction.
            if ($SkipReparsePoints -and $current -ne $longRoot -and (Test-IsReparsePoint $dirInfo)) { continue }
        } catch { continue }

        $entries = $null
        try {
            $entries = $dirInfo.EnumerateFileSystemInfos()
        } catch { continue }

        foreach ($entry in $entries) {
            if (Test-ShouldStop) { break }
            try {
                if ($entry -is [System.IO.DirectoryInfo]) {
                    if ($script:ExcludedDirNames -contains $entry.Name) { continue }
                    if ($SkipReparsePoints -and (Test-IsReparsePoint $entry)) {
                        if ($null -ne $ReparsePointLog) { $ReparsePointLog.Add((Remove-LongPathPrefix $entry.FullName)) }
                        continue
                    }
                    if ($IncludeDirectories) { Write-Output $entry }
                    $stack.Push($entry.FullName)
                } else {
                    if ($SkipReparsePoints -and (Test-IsReparsePoint $entry)) {
                        if ($null -ne $ReparsePointLog) { $ReparsePointLog.Add((Remove-LongPathPrefix $entry.FullName)) }
                        continue
                    }
                    if ($IncludeFiles) { Write-Output $entry }
                }
            } catch { continue }
        }
    }
}

function Get-RemainingExcludedDir {
    <#
    .SYNOPSIS
        Noch vorhandene ausgeschlossene Verzeichnisse unterhalb von RootPath.
    .DESCRIPTION
        Wird nach dem /MIR-Purge gebraucht. Was dort stehen geblieben ist,
        wurde per /XD bewusst NICHT kopiert (NAS-Schattenkopien, Papierkorb,
        Systemordner) und darf deshalb auch nicht geloescht werden -- es ist
        aber der Grund, warum der Quellordner nicht leer ist, und gehoert
        darum ins Protokoll statt in ein rekursives Force-Delete.
    #>
    param([string]$RootPath)

    $found = New-Object System.Collections.Generic.List[string]
    $stack = New-Object System.Collections.Generic.Stack[string]
    $stack.Push((Add-LongPathPrefix $RootPath))
    while ($stack.Count -gt 0) {
        $current = $stack.Pop()
        $subs = $null
        try { $subs = [System.IO.Directory]::EnumerateDirectories($current) } catch { continue }
        foreach ($sub in $subs) {
            try {
                if ($script:ExcludedDirNames -contains ([System.IO.Path]::GetFileName($sub))) {
                    $found.Add((Remove-LongPathPrefix $sub))
                    continue   # nicht hineinsteigen: der Inhalt bleibt ohnehin
                }
                # Reparse-Punkte nicht betreten (Schleifengefahr).
                if (Test-IsReparsePoint (New-Object System.IO.DirectoryInfo($sub))) { continue }
                $stack.Push($sub)
            } catch { continue }
        }
    }
    return $found
}

function ConvertTo-CsvField {
    param([object]$Value)
    if ($null -eq $Value) { return '""' }
    $s = "$Value"
    if ($s -match '[",\r\n]') {
        return '"' + ($s -replace '"', '""') + '"'
    }
    return '"' + $s + '"'
}

# ==================================================================
# HILFSFUNKTIONEN -- ZEITSTEMPEL (NDJSON-STREAMING)
# ==================================================================

function Export-Timestamps {
    param(
        [string]$RootPath,
        [string]$OutputFile,
        [bool]$SkipReparsePoints = $true,
        [string]$ReparseReportFile = ''
    )
    $root  = $RootPath.TrimEnd('\')
    $count = 0
    $maxDepth   = 0
    $hugeFiles  = 0
    [long]$totalBytes = 0
    $reparseList = $null
    if ($SkipReparsePoints -and $ReparseReportFile) {
        $reparseList = New-Object System.Collections.Generic.List[string]
    }
    $sw    = $null
    try {
        $sw = [System.IO.StreamWriter]::new($OutputFile, $false, $Utf8Bom)
        Get-FilesStreaming -RootPath $root -SkipReparsePoints $SkipReparsePoints -IncludeFiles -IncludeDirectories -ReparsePointLog $reparseList |
            ForEach-Object {
                if (Test-ShouldStop) { return }
                $full = Remove-LongPathPrefix $_.FullName
                $rel  = $full
                if ($rel.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $rel = $rel.Substring($root.Length).TrimStart('\')
                }
                if ($_ -is [System.IO.DirectoryInfo]) {
                    $depth = $rel.Split('\').Length
                    if ($depth -gt $maxDepth) { $maxDepth = $depth }
                } else {
                    $totalBytes += $_.Length
                    if ($_.Length -gt 5TB) { $hugeFiles++ }
                }
                try {
                    $rec = [ordered]@{
                        RelativePath      = $rel
                        IsDirectory       = ($_ -is [System.IO.DirectoryInfo])
                        CreationTimeUtc   = $_.CreationTimeUtc.ToString('o')
                        LastWriteTimeUtc  = $_.LastWriteTimeUtc.ToString('o')
                        LastAccessTimeUtc = $_.LastAccessTimeUtc.ToString('o')
                    }
                    $sw.WriteLine(($rec | ConvertTo-Json -Compress -Depth 2))
                    $count++
                } catch { }
            }
    } catch {
        Write-Warning "Zeitstempel-Export teilweise fehlgeschlagen: $_"
    } finally {
        if ($sw) { $sw.Close() }
    }

    if ($reparseList -and $reparseList.Count -gt 0) {
        $rw = $null
        try {
            $rw = [System.IO.StreamWriter]::new($ReparseReportFile, $false, $Utf8Bom)
            $rw.WriteLine('"Pfad"')
            foreach ($p in $reparseList) { $rw.WriteLine((ConvertTo-CsvField $p)) }
        } catch {
            Write-Warning "Symlink-Liste konnte nicht geschrieben werden: $_"
        } finally {
            if ($rw) { $rw.Close() }
        }
        Write-Host "  $($reparseList.Count) Symlink(s)/Junction(s) uebersprungen (werden NICHT migriert)." -ForegroundColor Yellow
        Write-Host "  Liste: $ReparseReportFile" -ForegroundColor Yellow
    }

    if (Test-ShouldStop) {
        # Teilexport - nicht als abgeschlossen melden (der Aufrufer bricht ab).
        Write-Warning "  Zeitstempel-Export abgebrochen - nur $count Eintraege erfasst."
    } else {
        Write-Host "  Zeitstempel exportiert: $count Eintraege -> $OutputFile"
    }
    return @{
        Count      = $count
        MaxDepth   = $maxDepth
        HugeFiles  = $hugeFiles
        TotalBytes = $totalBytes
    }
}

function Import-Timestamps {
    param(
        [string]$DestPath,
        [string]$TimestampFile,
        [bool]$RestoreAccessTime = $false
    )
    if (-not (Test-Path -LiteralPath $TimestampFile)) {
        Write-Warning "Zeitstempel-Datei nicht gefunden: $TimestampFile"
        return
    }
    $destRoot  = $DestPath.TrimEnd('\')
    $ok = 0; $skip = 0; $miss = 0; $fail = 0
    $invariant = [System.Globalization.CultureInfo]::InvariantCulture
    $style     = [System.Globalization.DateTimeStyles]::RoundtripKind

    # Verarbeitungsreihenfolge: erst alle DATEIEN, danach die ORDNER
    # (tiefste zuerst). Jede Datei-Zeitstempel-Aenderung aktualisiert die
    # LastWriteTime des Elternordners - wuerden Ordner zuerst oder
    # gemischt gesetzt, waeren ihre Zeitstempel anschliessend wieder
    # verfaelscht.
    $processRecord = {
        param($rec)
        $target     = if ($rec.RelativePath) { Join-Path $destRoot $rec.RelativePath } else { $destRoot }
        $longTarget = Add-LongPathPrefix $target
        $isDir      = [bool]$rec.IsDirectory
        $attributes = $null
        $wasReadOnly = $false

        try {
            $exists = if ($isDir) { [System.IO.Directory]::Exists($longTarget) } else { [System.IO.File]::Exists($longTarget) }
            if (-not $exists) { return 'missing' }

            $cTime = [datetime]::Parse($rec.CreationTimeUtc,   $invariant, $style)
            $mTime = [datetime]::Parse($rec.LastWriteTimeUtc,  $invariant, $style)
            $aTime = [datetime]::Parse($rec.LastAccessTimeUtc, $invariant, $style)

            $currentMtime = if ($isDir) { [System.IO.Directory]::GetLastWriteTimeUtc($longTarget) } else { [System.IO.File]::GetLastWriteTimeUtc($longTarget) }
            $currentCtime = if ($isDir) { [System.IO.Directory]::GetCreationTimeUtc($longTarget)  } else { [System.IO.File]::GetCreationTimeUtc($longTarget)  }

            $isTimeMatch = ([Math]::Abs(($currentMtime - $mTime).TotalSeconds) -lt 2 -and
                            [Math]::Abs(($currentCtime - $cTime).TotalSeconds) -lt 2)

            if ($isTimeMatch -and $RestoreAccessTime) {
                $currentAtime = if ($isDir) { [System.IO.Directory]::GetLastAccessTimeUtc($longTarget) } else { [System.IO.File]::GetLastAccessTimeUtc($longTarget) }
                $isTimeMatch  = ([Math]::Abs(($currentAtime - $aTime).TotalSeconds) -lt 2)
            }

            if ($isTimeMatch) { return 'skip' }

            if (-not $isDir) {
                $attributes  = [System.IO.File]::GetAttributes($longTarget)
                $wasReadOnly = ($attributes -band [System.IO.FileAttributes]::ReadOnly) -eq [System.IO.FileAttributes]::ReadOnly
                if ($wasReadOnly) {
                    Invoke-WithRetry -ScriptBlock {
                        [System.IO.File]::SetAttributes($longTarget, $attributes -band (-bnot [System.IO.FileAttributes]::ReadOnly))
                    } | Out-Null
                }
            }

            Invoke-WithRetry -ScriptBlock {
                if ($isDir) {
                    [System.IO.Directory]::SetCreationTimeUtc($longTarget, $cTime)
                    [System.IO.Directory]::SetLastWriteTimeUtc($longTarget, $mTime)
                    if ($RestoreAccessTime) { [System.IO.Directory]::SetLastAccessTimeUtc($longTarget, $aTime) }
                } else {
                    [System.IO.File]::SetCreationTimeUtc($longTarget, $cTime)
                    [System.IO.File]::SetLastWriteTimeUtc($longTarget, $mTime)
                    if ($RestoreAccessTime) { [System.IO.File]::SetLastAccessTimeUtc($longTarget, $aTime) }
                }
            } | Out-Null

            if ($wasReadOnly -and $null -ne $attributes) {
                try { [System.IO.File]::SetAttributes($longTarget, $attributes) } catch { }
            }
            return 'ok'
        } catch {
            return 'fail'
        }
    }

    $dirRecords = New-Object System.Collections.Generic.List[object]
    $reader = $null
    try {
        $reader = [System.IO.StreamReader]::new($TimestampFile, $Utf8Bom)
        while ($null -ne ($line = $reader.ReadLine())) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            $rec = $null
            try { $rec = $line | ConvertFrom-Json } catch { $fail++; continue }

            if ([bool]$rec.IsDirectory) {
                $dirRecords.Add($rec)
                continue
            }
            switch (& $processRecord $rec) {
                'ok'      { $ok++ }
                'skip'    { $skip++ }
                'missing' { $miss++ }
                default   { $fail++ }
            }
        }
    } catch {
        Write-Warning "Zeitstempel-Datei konnte nicht gelesen werden: $_"
    } finally {
        if ($reader) { $reader.Close() }
    }

    $dirRecords.Sort([System.Collections.Generic.Comparer[object]]::Create({
        param($a, $b)
        ("$($b.RelativePath)").Length.CompareTo(("$($a.RelativePath)").Length)
    }))
    foreach ($rec in $dirRecords) {
        switch (& $processRecord $rec) {
            'ok'      { $ok++ }
            'skip'    { $skip++ }
            'missing' { $miss++ }
            default   { $fail++ }
        }
    }

    $color = if ($fail -gt 0) { 'Red' } else { 'Green' }
    Write-Host "  Zeitstempel wiederhergestellt: $ok gesetzt, $skip bereits aktuell, $miss im Ziel nicht vorhanden (uebersprungen), $fail fehlgeschlagen." -ForegroundColor $color
}

# ==================================================================
# HILFSFUNKTIONEN -- GOOGLE-DRIVE-NAMENSPRUEFUNG (CSV-STREAMING)
# ==================================================================

function Test-GDriveCompatibility {
    param(
        [string]$RootPath,
        [string]$LogFile,
        [bool]$SkipReparsePoints = $true,
        [string]$DestRoot = ''
    )
    $forbidden = '[\\/:*?"<>|]'
    $reserved  = '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\.|$)'
    $count     = 0
    $csvFile   = [System.IO.Path]::ChangeExtension($LogFile, '.gdrive-probleme.csv')
    $writer    = $null
    $srcRoot       = $RootPath.TrimEnd('\')
    $destRootClean = if ([string]::IsNullOrWhiteSpace($DestRoot)) { $null } else { $DestRoot.TrimEnd('\') }

    Write-Host "Pruefe Google Drive-Kompatibilitaet aller Dateinamen ..."

    try {
        Get-FilesStreaming -RootPath $RootPath -SkipReparsePoints $SkipReparsePoints -IncludeFiles -IncludeDirectories |
            ForEach-Object {
                $name     = $_.Name
                $realPath = Remove-LongPathPrefix $_.FullName
                $issues   = New-Object System.Collections.Generic.List[string]

                if ($name -match $forbidden)     { $issues.Add("Verbotenes Zeichen ($($Matches[0]))") }
                if ($name -imatch $reserved)     { $issues.Add("Reservierter Name") }
                if ($name.Length -gt 255)        { $issues.Add("Name > 255 Zeichen ($($name.Length))") }
                if ($name -match '^\s|\s$')      { $issues.Add("Fuehrendes/abschliessendes Leerzeichen") }
                if ($name -match '\.$')          { $issues.Add("Abschliessender Punkt") }
                if ($destRootClean) {
                    # Massgeblich ist die Pfadlaenge im ZIEL (z.B. G:\Geteilte
                    # Ablagen\...), nicht in der Quelle -- der Zielbasispfad
                    # ist oft deutlich laenger als der Quellbasispfad.
                    $rel = $realPath
                    if ($rel.StartsWith($srcRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
                        $rel = $rel.Substring($srcRoot.Length).TrimStart('\')
                    }
                    $destLen = $destRootClean.Length + 1 + $rel.Length
                    if ($destLen -gt 260) { $issues.Add("Zielpfad > 260 Zeichen ($destLen)") }
                } elseif ($realPath.Length -gt 260) {
                    $issues.Add("Pfad > 260 Zeichen ($($realPath.Length))")
                }

                if ($issues.Count -gt 0) {
                    if ($null -eq $writer) {
                        $writer = [System.IO.StreamWriter]::new($csvFile, $false, $Utf8Bom)
                        $writer.WriteLine('"Typ","Pfad","Probleme"')
                    }
                    $typ = if ($_ -is [System.IO.DirectoryInfo]) { 'Ordner' } else { 'Datei' }
                    # Doppelte Klammer ist Pflicht (hier und an den zwoelf
                    # weiteren WriteLine-Stellen mit -f): in einer
                    # Methoden-Argumentliste trennt das Komma die ARGUMENTE,
                    # -f bekam nur den ersten Wert und warf FormatException
                    # (gemessen unter 5.1). Die Pruefsummen- und
                    # Kompatibilitaetspruefung brachen dadurch bei der ersten
                    # Datei ab und meldeten fuer 0 Dateien "alles in Ordnung".
                    $writer.WriteLine(("{0},{1},{2}" -f `
                        (ConvertTo-CsvField $typ), `
                        (ConvertTo-CsvField $realPath), `
                        (ConvertTo-CsvField ($issues -join '; '))))
                    $count++
                }
            }
    } catch {
        Write-Warning "Kompatibilitaetspruefung fehlgeschlagen: $_"
    } finally {
        if ($writer) { $writer.Close() }
    }

    # Bei Strg+C bricht Get-FilesStreaming ab und $count ist nur ein
    # Teilergebnis - vorher folgte dann "Alle Dateinamen sind Google
    # Drive-kompatibel" in Gruen. Der Aufrufer bricht danach ab.
    if (Test-ShouldStop) {
        Write-Warning "  Kompatibilitaetspruefung abgebrochen - Ergebnis unvollstaendig ($count Problem(e) bis zum Abbruch)."
        return $count
    }
    if ($count -gt 0) {
        Write-Warning "  $count Datei(en)/Ordner mit GDrive-Kompatibilitaetsproblemen gefunden."
        Write-Warning "  Details: $csvFile"
        Write-Warning "  Bitte problematische Dateien vor der Migration umbenennen."
    } else {
        Write-Host "  Alle Dateinamen sind Google Drive-kompatibel." -ForegroundColor Green
    }
    return $count
}

function Repair-GDrivePaths {
    param(
        [string]$RootPath,
        [string]$LogFile,
        [bool]$SkipReparsePoints = $true
    )

    $forbidden = '[\\/:*?"<>|]'
    $reserved  = '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\.|$)'
    $csvFile   = [System.IO.Path]::ChangeExtension($LogFile, '.umbenennungen.csv')
    $writer    = $null
    $ok = 0; $fail = 0; $written = 0

    Write-Host "Bereinige Datei- und Ordnernamen ..."

    $items = New-Object System.Collections.Generic.List[System.IO.FileSystemInfo]
    Get-FilesStreaming -RootPath $RootPath -SkipReparsePoints $SkipReparsePoints -IncludeFiles -IncludeDirectories |
        ForEach-Object { $items.Add($_) }

    if ($items.Count -eq 0) {
        Write-Host "  Keine Dateien gefunden."
        return 0
    }

    $items.Sort([System.Collections.Generic.Comparer[System.IO.FileSystemInfo]]::Create({
        param($a, $b)
        $b.FullName.Length.CompareTo($a.FullName.Length)
    }))

    $colWidth = try { [Math]::Max(40, [Console]::WindowWidth - 6) } catch { 80 }

    foreach ($item in $items) {
        if (Test-ShouldStop) { break }

        $oldName = $item.Name
        $newName = $oldName

        $newName = $newName -replace '^\s+', ''
        $newName = $newName -replace '[\s.]+$', ''
        $beforeForbidden = $newName
        $newName = $newName -replace $forbidden, '-'
        if ($newName -cne $beforeForbidden) {
            # Nur kollabieren, wenn zuvor verbotene Zeichen ersetzt wurden --
            # legitime Mehrfach-Bindestriche in gueltigen Namen bleiben erhalten.
            $newName = $newName -replace '-{2,}', '-'
        }

        if ($newName -imatch $reserved) {
            $newName = '_' + $newName
        }

        if ([string]::IsNullOrWhiteSpace($newName)) {
            $ext = [System.IO.Path]::GetExtension($oldName)
            $newName = if ($ext -and $ext -ne '.') { "_bereinigt$ext" } else { '_bereinigt' }
        }

        if ($oldName -eq $newName) { continue }

        $parentReal = Remove-LongPathPrefix (Split-Path $item.FullName -Parent)
        $targetPath = Join-Path $parentReal $newName
        $longTarget = Add-LongPathPrefix $targetPath

        # [System.IO] statt Test-Path: Provider-Cmdlets sind mit '\\?\'-Pfaden
        # unter 5.1 nicht verlaesslich (siehe 7; fuer UNC/G: nicht messbar),
        # und ein falsches "frei" liesse unten eine bestehende Datei
        # ueberschreiben. Gemeldet von pruefe_alles (Muster 5.2).
        if ([System.IO.File]::Exists($longTarget) -or [System.IO.Directory]::Exists($longTarget)) {
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($newName)
            $ext      = [System.IO.Path]::GetExtension($newName)
            $counter  = 2
            do {
                $candidate  = "${baseName}_${counter}${ext}"
                $targetPath = Join-Path $parentReal $candidate
                $longTarget = Add-LongPathPrefix $targetPath
                $counter++
            } while ([System.IO.File]::Exists($longTarget) -or [System.IO.Directory]::Exists($longTarget))
            $newName = $candidate
        }

        $typ = if ($item -is [System.IO.DirectoryInfo]) { 'Ordner' } else { 'Datei' }

        try {
            Invoke-WithRetry -ScriptBlock {
                Move-Item -LiteralPath (Add-LongPathPrefix $item.FullName) `
                          -Destination $longTarget -Force -ErrorAction Stop
            } | Out-Null

            if ($null -eq $writer) {
                $writer = [System.IO.StreamWriter]::new($csvFile, $false, $Utf8Bom)
                $writer.WriteLine('"Typ","AlterName","NeuerName","Pfad","Status"')
            }
            $writer.WriteLine(("{0},{1},{2},{3},{4}" -f `
                (ConvertTo-CsvField $typ), `
                (ConvertTo-CsvField $oldName), `
                (ConvertTo-CsvField $newName), `
                (ConvertTo-CsvField $parentReal), `
                (ConvertTo-CsvField 'OK')))
            $written++
            $ok++

            $disp = "'$oldName' -> '$newName'"
            if ($disp.Length -gt ($colWidth - 4)) {
                $disp = $disp.Substring(0, $colWidth - 7) + '...'
            }
            Write-Host "  $disp"
        } catch {
            if ($null -eq $writer) {
                $writer = [System.IO.StreamWriter]::new($csvFile, $false, $Utf8Bom)
                $writer.WriteLine('"Typ","AlterName","NeuerName","Pfad","Status"')
            }
            $writer.WriteLine(("{0},{1},{2},{3},{4}" -f `
                (ConvertTo-CsvField $typ), `
                (ConvertTo-CsvField $oldName), `
                (ConvertTo-CsvField $newName), `
                (ConvertTo-CsvField $parentReal), `
                (ConvertTo-CsvField "FEHLER: $_")))
            $written++
            $fail++
            Write-Warning "  Fehler beim Umbenennen von '$oldName': $_"
        }
    }

    # Vorher stand Close() ausserhalb jedes finally: brach die Schleife mit
    # einer Ausnahme ab, blieb das Handle auf der CSV bis zur GC offen und
    # der Report war unvollstaendig/gesperrt.
    try { if ($writer) { $writer.Close() } } catch { }

    if ($written -gt 0) {
        $color = if ($fail -gt 0) { 'Red' } else { 'Green' }
        Write-Host "  $ok Umbenennung(en) erfolgreich, $fail fehlgeschlagen." -ForegroundColor $color
        Write-Host "  Umbenennungs-Report: $csvFile"

        Add-LogLine -Path $LogFile `
            -Message "[$(Get-Date)] BEREINIGUNG: $ok umbenannt, $fail fehlgeschlagen. Report: $csvFile" `
            -Encoding ([System.Text.Encoding]::Unicode)
    } else {
        Write-Host "  Keine Umbenennungen noetig." -ForegroundColor Green
    }

    return $ok
}

# ==================================================================
# HILFSFUNKTIONEN -- VERIFIKATION (CRC / LIGHT, CSV-STREAMING)
# ==================================================================

function Compare-FileChecksums {
    param(
        [string]$SourceRoot,
        [string]$DestRoot,
        [string]$ReportFile,
        [string]$LogFile,
        [bool]$SkipReparsePoints = $true
    )

    $srcRoot  = $SourceRoot.TrimEnd('\')
    $dstRoot  = $DestRoot.TrimEnd('\')
    $okPaths  = New-Object System.Collections.Generic.List[string]
    $counters = @{ Match = 0; Mismatch = 0; Missing = 0; Error = 0; Total = 0; Bytes = [long]0 }

    Write-Host "  Sammle Quelldateien (gestreamt) ..."

    $colWidth = try { [Math]::Max(40, [Console]::WindowWidth - 6) } catch { 80 }
    $sha      = [System.Security.Cryptography.SHA256]::Create()
    $sw       = [System.Diagnostics.Stopwatch]::StartNew()
    $writer   = $null
    $processed = 0

    try {
        $writer = [System.IO.StreamWriter]::new($ReportFile, $false, $Utf8Bom)
        $writer.WriteLine('"RelativerPfad","Status","SHA256_Quelle","SHA256_Ziel","Groesse_Bytes"')

        Get-FilesStreaming -RootPath $srcRoot -SkipReparsePoints $SkipReparsePoints -IncludeFiles |
            ForEach-Object {
                if (Test-ShouldStop) { return }
                $srcItem = $_
                $processed++
                $counters.Total++

                $rel = Remove-LongPathPrefix $srcItem.FullName
                if ($rel.StartsWith($srcRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $rel = $rel.Substring($srcRoot.Length).TrimStart('\')
                }

                if ($processed % 50 -eq 0 -or $sw.Elapsed.TotalSeconds -ge 3) {
                    $disp = if ($rel.Length -gt ($colWidth - 12)) {
                        '...' + $rel.Substring($rel.Length - $colWidth + 15)
                    } else { $rel }
                    Write-Host ("`r  [{0,6}] {1}" -f $processed, $disp.PadRight($colWidth - 10)) -NoNewline
                    $sw.Restart()
                }

                $srcLong = Add-LongPathPrefix (Remove-LongPathPrefix $srcItem.FullName)
                $srcHash = $null
                try {
                    # Initialize() vor JEDEM Versuch: ComputeHash(Stream) setzt
                    # das gemeinsame SHA256-Objekt nur nach einem vollstaendigen
                    # Durchlauf zurueck. Warf das Lesen mittendrin, blieb der
                    # Teilzustand stehen, und der naechste Versuch bzw. die
                    # naechste Datei bekam einen falschen Hash (nachgestellt
                    # mit einem Stream, der nach 64 KB wirft).
                    $srcHash = Invoke-WithRetry -ScriptBlock {
                        $fs = [System.IO.File]::OpenRead($srcLong)
                        try   { $sha.Initialize(); [BitConverter]::ToString($sha.ComputeHash($fs)).Replace('-', '') }
                        finally { $fs.Close() }
                    }
                } catch {
                    $writer.WriteLine(("{0},{1},{2},{3},{4}" -f `
                        (ConvertTo-CsvField $rel), `
                        (ConvertTo-CsvField 'FEHLER_QUELLE'), `
                        (ConvertTo-CsvField "LESEFEHLER: $_"), `
                        (ConvertTo-CsvField ''), `
                        (ConvertTo-CsvField $srcItem.Length)))
                    $counters.Error++
                    return
                }

                $dstPath = Join-Path $dstRoot $rel
                $dstLong = Add-LongPathPrefix $dstPath

                if (-not [System.IO.File]::Exists($dstLong)) {
                    $writer.WriteLine(("{0},{1},{2},{3},{4}" -f `
                        (ConvertTo-CsvField $rel), `
                        (ConvertTo-CsvField 'FEHLT_IM_ZIEL'), `
                        (ConvertTo-CsvField $srcHash), `
                        (ConvertTo-CsvField ''), `
                        (ConvertTo-CsvField $srcItem.Length)))
                    $counters.Missing++
                    return
                }

                $dstHash = $null
                try {
                    $dstHash = Invoke-WithRetry -ScriptBlock {
                        $fs = [System.IO.File]::OpenRead($dstLong)
                        try   { $sha.Initialize(); [BitConverter]::ToString($sha.ComputeHash($fs)).Replace('-', '') }
                        finally { $fs.Close() }
                    }
                } catch {
                    $writer.WriteLine(("{0},{1},{2},{3},{4}" -f `
                        (ConvertTo-CsvField $rel), `
                        (ConvertTo-CsvField 'FEHLER_ZIEL'), `
                        (ConvertTo-CsvField $srcHash), `
                        (ConvertTo-CsvField "LESEFEHLER: $_"), `
                        (ConvertTo-CsvField $srcItem.Length)))
                    $counters.Error++
                    return
                }

                if ($srcHash -eq $dstHash) {
                    $writer.WriteLine(("{0},{1},{2},{3},{4}" -f `
                        (ConvertTo-CsvField $rel), `
                        (ConvertTo-CsvField 'OK'), `
                        (ConvertTo-CsvField $srcHash), `
                        (ConvertTo-CsvField $dstHash), `
                        (ConvertTo-CsvField $srcItem.Length)))
                    $counters.Match++
                    $counters.Bytes += $srcItem.Length
                    $okPaths.Add($rel)
                } else {
                    $writer.WriteLine(("{0},{1},{2},{3},{4}" -f `
                        (ConvertTo-CsvField $rel), `
                        (ConvertTo-CsvField 'ABWEICHUNG'), `
                        (ConvertTo-CsvField $srcHash), `
                        (ConvertTo-CsvField $dstHash), `
                        (ConvertTo-CsvField $srcItem.Length)))
                    $counters.Mismatch++
                }
            }
    } catch {
        Write-Warning "Pruefsummen-Vergleich teilweise fehlgeschlagen: $_"
    } finally {
        if ($writer) { $writer.Close() }
        $sha.Dispose()
    }

    Write-Host ("`r" + (' ' * ($colWidth + 4)) + "`r") -NoNewline

    if ($counters.Total -eq 0) {
        Write-Host "  Keine Dateien zum Pruefen gefunden."
        return @{ Counters = $counters; OkPaths = $okPaths }
    }

    $allOk = ($counters.Mismatch -eq 0 -and $counters.Missing -eq 0 -and $counters.Error -eq 0)
    if ($allOk) {
        Write-Host "  Pruefsummen-Ergebnis: Alle $($counters.Match) Dateien sind bit-identisch." -ForegroundColor Green
    } else {
        Write-Host "  Pruefsummen-Ergebnis:"
        Write-Host "    OK (identisch):    $($counters.Match)" -ForegroundColor Green
        if ($counters.Mismatch -gt 0) { Write-Host "    ABWEICHUNG:        $($counters.Mismatch)" -ForegroundColor Red }
        if ($counters.Missing  -gt 0) { Write-Host "    FEHLT IM ZIEL:     $($counters.Missing)"  -ForegroundColor Red }
        if ($counters.Error    -gt 0) { Write-Host "    LESEFEHLER:        $($counters.Error)"    -ForegroundColor Yellow }
        Write-Host "    Details: $ReportFile"
    }

    $logMsg = "[$(Get-Date)] PRUEFSUMMEN: $($counters.Total) Dateien, $($counters.Match) OK, " +
              "$($counters.Mismatch) Abweichungen, $($counters.Missing) fehlend, $($counters.Error) Fehler. " +
              "Report: $ReportFile"
    Add-LogLine -Path $LogFile -Message $logMsg -Encoding ([System.Text.Encoding]::Unicode)

    return @{ Counters = $counters; OkPaths = $okPaths }
}

function Compare-FileExistence {
    param(
        [string]$SourceRoot,
        [string]$DestRoot,
        [string]$ReportFile,
        [string]$LogFile,
        [bool]$SkipReparsePoints = $true
    )

    $srcRoot  = $SourceRoot.TrimEnd('\')
    $dstRoot  = $DestRoot.TrimEnd('\')
    $okPaths  = New-Object System.Collections.Generic.List[string]
    $counters = @{ Match = 0; Mismatch = 0; Missing = 0; Error = 0; Total = 0; Bytes = [long]0; Placeholder = 0 }

    Write-Host "  Sammle Quelldateien (gestreamt) ..."

    $colWidth  = try { [Math]::Max(40, [Console]::WindowWidth - 6) } catch { 80 }
    $sw        = [System.Diagnostics.Stopwatch]::StartNew()
    $writer    = $null
    $processed = 0

    try {
        $writer = [System.IO.StreamWriter]::new($ReportFile, $false, $Utf8Bom)
        $writer.WriteLine('"RelativerPfad","Status","Groesse_Quelle","Groesse_Ziel"')

        Get-FilesStreaming -RootPath $srcRoot -SkipReparsePoints $SkipReparsePoints -IncludeFiles |
            ForEach-Object {
                if (Test-ShouldStop) { return }
                $srcItem = $_
                $processed++
                $counters.Total++

                $rel = Remove-LongPathPrefix $srcItem.FullName
                if ($rel.StartsWith($srcRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $rel = $rel.Substring($srcRoot.Length).TrimStart('\')
                }

                if ($processed % 200 -eq 0 -or $sw.Elapsed.TotalSeconds -ge 2) {
                    $disp = if ($rel.Length -gt ($colWidth - 12)) {
                        '...' + $rel.Substring($rel.Length - $colWidth + 15)
                    } else { $rel }
                    Write-Host ("`r  [{0,6}] {1}" -f $processed, $disp.PadRight($colWidth - 10)) -NoNewline
                    $sw.Restart()
                }

                $dstPath = Join-Path $dstRoot $rel
                $dstLong = Add-LongPathPrefix $dstPath

                if (-not [System.IO.File]::Exists($dstLong)) {
                    $writer.WriteLine(("{0},{1},{2},{3}" -f `
                        (ConvertTo-CsvField $rel), `
                        (ConvertTo-CsvField 'FEHLT_IM_ZIEL'), `
                        (ConvertTo-CsvField $srcItem.Length), `
                        (ConvertTo-CsvField '')))
                    $counters.Missing++
                    return
                }

                try {
                    $dstInfo = New-Object System.IO.FileInfo($dstLong)
                    $dstSize = $dstInfo.Length
                } catch {
                    $writer.WriteLine(("{0},{1},{2},{3}" -f `
                        (ConvertTo-CsvField $rel), `
                        (ConvertTo-CsvField 'FEHLER_ZIEL'), `
                        (ConvertTo-CsvField $srcItem.Length), `
                        (ConvertTo-CsvField "LESEFEHLER: $_")))
                    $counters.Error++
                    return
                }

                # Cloud-Platzhalter erkennen, BEVOR die Datei als geprueft
                # gilt. Ein Platzhalter (Attribute unten) ist eine Datei,
                # deren Inhalt NICHT lokal liegt, sondern erst beim Zugriff
                # aus der Cloud geholt wird; ihre Groesse ist nur die
                # gemeldete Groesse, kein Beleg fuer den Inhalt. Deshalb
                # zaehlt sie hier bewusst NICHT als OK und bleibt in der
                # Quelle. Das CRC-Verfahren liest den Inhalt (und holt ihn
                # dabei herunter) und vergleicht ihn damit tatsaechlich.
                # Frueher hiess es hier, der Inhalt sei "noch nicht oben" -
                # das ist umgekehrt: Drive for Desktop macht eine Datei erst
                # NACH dem Hochladen zum Platzhalter. Umgekehrt beweist eine
                # lokal vorhandene Datei (kein Platzhalter) NICHT, dass sie
                # schon hochgeladen ist - sie liegt womoeglich nur im Cache.
                # Der Light-Check prueft also Existenz und Groesse im
                # lokalen Ziel (Drive-Cache), nicht den Upload. Die
                # Zaehlung bleibt absichtlich so: sie erleichtert das
                # Loeschen der Quelle in keinem Fall.
                #   0x00400000 FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS
                #   0x00040000 FILE_ATTRIBUTE_RECALL_ON_OPEN
                #   0x00001000 FILE_ATTRIBUTE_OFFLINE
                $placeholderMask = 0x00400000 -bor 0x00040000 -bor 0x00001000
                $isPlaceholder   = $false
                try {
                    $isPlaceholder = ((([int]$dstInfo.Attributes) -band $placeholderMask) -ne 0)
                } catch { }

                if ($isPlaceholder) {
                    $writer.WriteLine(("{0},{1},{2},{3}" -f `
                        (ConvertTo-CsvField $rel), `
                        (ConvertTo-CsvField 'PLATZHALTER'), `
                        (ConvertTo-CsvField $srcItem.Length), `
                        (ConvertTo-CsvField "$dstSize (nur Platzhalter - Inhalt nicht lokal verfuegbar)")))
                    $counters.Placeholder++
                    return
                }

                if ($srcItem.Length -eq $dstSize) {
                    $writer.WriteLine(("{0},{1},{2},{3}" -f `
                        (ConvertTo-CsvField $rel), `
                        (ConvertTo-CsvField 'OK'), `
                        (ConvertTo-CsvField $srcItem.Length), `
                        (ConvertTo-CsvField $dstSize)))
                    $counters.Match++
                    $counters.Bytes += $srcItem.Length
                    $okPaths.Add($rel)
                } else {
                    $writer.WriteLine(("{0},{1},{2},{3}" -f `
                        (ConvertTo-CsvField $rel), `
                        (ConvertTo-CsvField 'ABWEICHUNG'), `
                        (ConvertTo-CsvField $srcItem.Length), `
                        (ConvertTo-CsvField $dstSize)))
                    $counters.Mismatch++
                }
            }
    } catch {
        Write-Warning "Light-Check teilweise fehlgeschlagen: $_"
    } finally {
        if ($writer) { $writer.Close() }
    }

    Write-Host ("`r" + (' ' * ($colWidth + 4)) + "`r") -NoNewline

    if ($counters.Total -eq 0) {
        Write-Host "  Keine Dateien zum Pruefen gefunden."
        return @{ Counters = $counters; OkPaths = $okPaths }
    }

    $allOk = ($counters.Mismatch -eq 0 -and $counters.Missing -eq 0 -and
              $counters.Error -eq 0 -and $counters.Placeholder -eq 0)
    if ($allOk) {
        Write-Host "  Light-Check: Alle $($counters.Match) Dateien im Ziel vorhanden (Groesse identisch)." -ForegroundColor Green
    } else {
        Write-Host "  Light-Check-Ergebnis:"
        Write-Host "    OK (Existenz+Groesse): $($counters.Match)" -ForegroundColor Green
        if ($counters.Mismatch -gt 0) { Write-Host "    GROESSE ABWEICHEND:    $($counters.Mismatch)" -ForegroundColor Red }
        if ($counters.Missing  -gt 0) { Write-Host "    FEHLT IM ZIEL:         $($counters.Missing)"  -ForegroundColor Red }
        if ($counters.Error    -gt 0) { Write-Host "    LESEFEHLER:            $($counters.Error)"    -ForegroundColor Yellow }
        if ($counters.Placeholder -gt 0) {
            Write-Host "    NUR PLATZHALTER:       $($counters.Placeholder)" -ForegroundColor Yellow
            Write-Host "      Der Inhalt dieser Dateien liegt nicht lokal (nur in der Cloud); die" -ForegroundColor DarkYellow
            Write-Host "      gemeldete Groesse belegt den Inhalt nicht. Sie gelten als NICHT" -ForegroundColor DarkYellow
            Write-Host "      verifiziert und werden in der Quelle behalten. Fuer einen" -ForegroundColor DarkYellow
            Write-Host "      Inhaltsvergleich das CRC-Verfahren waehlen - es liest die Dateien" -ForegroundColor DarkYellow
            Write-Host "      (und laedt sie dafuer herunter)." -ForegroundColor DarkYellow
        }
        Write-Host "    Details: $ReportFile"
    }

    $logMsg = "[$(Get-Date)] LIGHT-CHECK: $($counters.Total) Dateien, $($counters.Match) OK, " +
              "$($counters.Mismatch) Groessenabweichung, $($counters.Missing) fehlend, $($counters.Error) Fehler. " +
              "Report: $ReportFile"
    Add-LogLine -Path $LogFile -Message $logMsg -Encoding ([System.Text.Encoding]::Unicode)

    return @{ Counters = $counters; OkPaths = $okPaths }
}

# ==================================================================
# HILFSFUNKTIONEN -- RCLONE-EINRICHTUNG
# ==================================================================

function Show-RcloneSetupGuide {
    Write-Host ""
    Write-Host "=============================================================="
    Write-Host "  RCLONE-EINRICHTUNG -- SCHRITT FUER SCHRITT"
    Write-Host "=============================================================="
    Write-Host ""
    Write-Host "SCHRITT 1 -- rclone.exe herunterladen und bereitstellen"
    Write-Host "--------------------------------------------------------------"
    Write-Host "  (a) Oeffnen Sie https://rclone.org/downloads/"
    Write-Host "  (b) Laden Sie das ZIP 'Intel/AMD - 64 Bit' fuer Windows herunter."
    Write-Host "  (c) Entpacken Sie das Archiv."
    Write-Host "  (d) Kopieren Sie 'rclone.exe' in EINES der folgenden Verzeichnisse:"
    Write-Host "        -- denselben Ordner wie dieses Skript (empfohlen)"
    Write-Host "        -- C:\Windows\System32 (erfordert Admin-Rechte)"
    Write-Host "        -- ein Verzeichnis, das in der PATH-Umgebungsvariable steht"
    Write-Host "  (e) Mindestversion: 1.62 (wegen Parameter '--metadata')."
    Write-Host ""
    Write-Host "SCHRITT 2 -- Google-Workspace-Projekt vorbereiten (einmalig durch IT)"
    Write-Host "--------------------------------------------------------------"
    Write-Host "  Fuer Unternehmens-Migrationen wird ein Dienstkonto (Service Account)"
    Write-Host "  empfohlen. Dadurch entfaellt die persoenliche Browser-Authentifizierung,"
    Write-Host "  und der Zugriff laeuft auch ohne eingeloggten Nutzer."
    Write-Host ""
    Write-Host "  (a) Oeffnen Sie die Google Cloud Console: https://console.cloud.google.com"
    Write-Host "  (b) Projekt anlegen (z.B. 'rclone-migration') oder vorhandenes waehlen."
    Write-Host "  (c) Menue -> 'APIs und Dienste' -> 'Bibliothek':"
    Write-Host "        Suchen und AKTIVIEREN Sie die 'Google Drive API'."
    Write-Host "  (d) Menue -> 'IAM und Verwaltung' -> 'Dienstkonten':"
    Write-Host "        'Dienstkonto erstellen' -> Name vergeben (z.B. 'rclone-sa')."
    Write-Host "  (e) Fuer das neue Dienstkonto:"
    Write-Host "        Tab 'Schluessel' -> 'Schluessel hinzufuegen' -> 'Neuen Schluessel erstellen'"
    Write-Host "        Typ: JSON -> Download starten -> Datei sicher ablegen"
    Write-Host "        (z.B. C:\ProgramData\rclone\sa.json)."
    Write-Host "  (f) Details des Dienstkontos notieren:"
    Write-Host "        -- Client-ID (eindeutige numerische ID)"
    Write-Host "        -- Dienstkonto-E-Mail (endet auf .iam.gserviceaccount.com)"
    Write-Host ""
    Write-Host "SCHRITT 3 -- Domainweite Delegation (nur fuer 'Meine Ablage')"
    Write-Host "--------------------------------------------------------------"
    Write-Host "  Nur noetig, wenn das Dienstkonto im Namen eines konkreten Nutzers"
    Write-Host "  agieren soll (typisch fuer R:\, Desktop, Downloads -> Meine Ablage)."
    Write-Host "  Fuer Geteilte Ablagen (Q:\) ist SCHRITT 3 nicht erforderlich --"
    Write-Host "  dort genuegt SCHRITT 4."
    Write-Host ""
    Write-Host "  (a) Google Workspace Admin Console: https://admin.google.com"
    Write-Host "  (b) Sicherheit -> Zugriffs- und Datenkontrolle -> API-Steuerung"
    Write-Host "  (c) 'Domainweite Delegation verwalten' -> 'Neu hinzufuegen'"
    Write-Host "  (d) Client-ID des Dienstkontos einfuegen."
    Write-Host "  (e) OAuth-Bereiche (scopes): https://www.googleapis.com/auth/drive"
    Write-Host "  (f) Speichern."
    Write-Host ""
    Write-Host "SCHRITT 4 -- Geteilte Ablage freigeben (fuer Q:\-Szenarien)"
    Write-Host "--------------------------------------------------------------"
    Write-Host "  Fuer Geteilte Ablagen muss das Dienstkonto als Mitglied eingetragen werden."
    Write-Host ""
    Write-Host "  (a) In Google Drive: gewuenschte Geteilte Ablage oeffnen."
    Write-Host "  (b) 'Mitglieder verwalten' -> Dienstkonto-E-Mail eintragen."
    Write-Host "  (c) Rolle: 'Inhalte verwalten' oder 'Administrator'."
    Write-Host ""
    Write-Host "SCHRITT 5 -- rclone-Remote konfigurieren"
    Write-Host "--------------------------------------------------------------"
    Write-Host "  Interaktive Konfiguration auf der Kommandozeile:"
    Write-Host ""
    Write-Host "    > rclone config"
    Write-Host ""
    Write-Host "  Dialogfolge (Antworten jeweils ohne Anfuehrungszeichen eingeben):"
    Write-Host ""
    Write-Host "    n                                # New remote"
    Write-Host "    name> ProjektDrive               # Remote-Name -- MUSS dem Wert in"
    Write-Host "                                     #   `$RcloneRemoteQ entsprechen"
    Write-Host "                                     #   (oder GDriveEnterprise fuer R:/Desktop/Downloads)"
    Write-Host "    Storage> drive                   # oder Nummer fuer 'Google Drive'"
    Write-Host "    client_id>                       # leer lassen (Default) oder eigene ID"
    Write-Host "    client_secret>                   # leer lassen (Default)"
    Write-Host "    scope> 1                         # '1' = vollstaendiger Drive-Zugriff"
    Write-Host "    service_account_file> C:\ProgramData\rclone\sa.json"
    Write-Host "    Edit advanced config? y/n> n"
    Write-Host "    Use auto config? y/n> n          # auf Servern oder fuer Service Accounts"
    Write-Host "    Configure this as a team drive? y/n>"
    Write-Host "      -- 'y' fuer Geteilte Ablagen (Q:\)"
    Write-Host "      -- 'n' fuer Meine Ablage (R:\, Desktop, Downloads)"
    Write-Host "    Bei 'y': team_drive-ID auswaehlen (wird automatisch aufgelistet)."
    Write-Host "    Keep this remote? y/e/d> y"
    Write-Host "    q                                # Beenden"
    Write-Host ""
    Write-Host "  Falls zwei Remotes benoetigt werden (Q:\ und Meine Ablage), SCHRITT 5"
    Write-Host "  fuer jedes Remote einmal durchlaufen."
    Write-Host ""
    Write-Host "SCHRITT 6 -- Smoke-Test"
    Write-Host "--------------------------------------------------------------"
    Write-Host "  Pruefen Sie die Konfiguration auf der Kommandozeile:"
    Write-Host ""
    Write-Host "    > rclone listremotes"
    Write-Host "    > rclone lsd ProjektDrive:"
    Write-Host "    > rclone lsd GDriveEnterprise:"
    Write-Host ""
    Write-Host "  Wenn beide Befehle Ordner auflisten (ohne Fehler), ist die"
    Write-Host "  Einrichtung abgeschlossen."
    Write-Host ""
    Write-Host "SCHRITT 7 -- Hinweise zu den Transfer-Parametern"
    Write-Host "--------------------------------------------------------------"
    Write-Host "  --bwlimit '08:00,15M 18:00,off'   Bandbreite tagsueber 15 MB/s, nachts unlimitiert"
    Write-Host "  --transfers=N                     Parallele Uploads (dynamisch nach RAM)"
    Write-Host "  --tpslimit=8                      API-Aufrufe/Sek. (Rate-Limit-Schutz)"
    Write-Host "  --drive-chunk-size=NM             Upload-Chunkgroesse (dynamisch nach RAM)"
    Write-Host "  --drive-use-trash=false           Sofortiges Loeschen (kein Papierkorb)"
    Write-Host "  --retries=5 / --low-level-retries=15   Toleranz gegen Netzwerkfehler"
    Write-Host "  --metadata                        Uebertraegt Zeitstempel in die Cloud"
    Write-Host "  --drive-stop-on-upload-limit      Sauberer Stopp am Google-Upload-Tageslimit"
    Write-Host ""
    Write-Host "  RAM-Bedarf wird automatisch an verfuegbaren Arbeitsspeicher angepasst:"
    Write-Host "    <= 8 GB:  3 transfers x 16M = ca. 48 MB"
    Write-Host "    <= 16 GB: 4 transfers x 32M = ca. 128 MB"
    Write-Host "    > 16 GB:  8 transfers x 64M = ca. 512 MB"
    Write-Host ""
    Write-Host "  WICHTIG -- GOOGLE-UPLOAD-TAGESLIMIT:"
    Write-Host "  Google Drive erlaubt pro Konto (auch Dienstkonto) maximal 750 GB"
    Write-Host "  Upload pro Tag. Groessere Migrationen verteilen sich daher ueber"
    Write-Host "  mehrere Tage: Skript am Folgetag einfach erneut starten -- bereits"
    Write-Host "  verschobene Dateien sind aus der Quelle entfernt, der Lauf setzt"
    Write-Host "  automatisch dort fort, wo er gestoppt wurde."
    Write-Host ""
    Write-Host "=============================================================="
    Write-Host ""
}

# ==================================================================
# POWERSHELL-VERSIONS-CHECK
# ==================================================================
if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Host ""
    Write-Host "Dieses Skript benoetigt PowerShell 5.0 oder neuer." -ForegroundColor Red
    Write-Host "Aktuelle Version: $($PSVersionTable.PSVersion)" -ForegroundColor Red
    Write-Host "Bitte aktualisieren Sie PowerShell (Win10/11: bereits 5.1)."
    Stop-Script 1
}

# ==================================================================
# METHODEN-AUSWAHL
# ==================================================================
Write-Host "`n╔═══════════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║ Verschieben von Verzeichnissen und Dateien auf Google Drive - Laufwerk G: ║" -ForegroundColor Cyan
Write-Host "╚═══════════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
Write-Host "================================================"
Write-Host "  MIGRATIONSMETHODE WAEHLEN"
Write-Host "================================================"
Write-Host "  [1] Robocopy  (lokal, sofort einsatzbereit)"
Write-Host "  [2] rclone    (direkt in die Cloud, benoetigt Einrichtung)"
Write-Host "================================================`n"
$MethodChoice = Read-Host "Auswahl (1 oder 2)"

if ($MethodChoice -ne "1" -and $MethodChoice -ne "2") {
    Write-Warning "Ungueltige Auswahl. Bitte 1 oder 2 eingeben."
    Stop-Script
}

# ==================================================================
# RCLONE-VORAUSSETZUNGEN
# ==================================================================
$RcloneSettings = $null
if ($MethodChoice -eq "2") {
    Show-RcloneSetupGuide

    Write-Host "Sind alle oben beschriebenen Voraussetzungen erfuellt? [J/N]"
    $RcloneReady = Read-Host
    if ($RcloneReady -notmatch '^[JjYy]$') {
        Write-Host "Bitte zuerst alle Voraussetzungen erfuellen und das Skript erneut starten."
        Stop-Script
    }

    $ScriptDir = Get-ScriptDirectory
    $RcloneExe = $null
    if ($ScriptDir) {
        $localRclone = Join-Path $ScriptDir "rclone.exe"
        if (Test-Path -LiteralPath $localRclone) {
            $RcloneExe = $localRclone
        }
    }
    if (-not $RcloneExe -and (Get-Command rclone -ErrorAction SilentlyContinue)) {
        $RcloneExe = "rclone"
    }
    if (-not $RcloneExe) {
        Write-Warning "rclone.exe wurde weder im Skriptverzeichnis ('$ScriptDir') noch im PATH gefunden."
        Write-Warning "Bitte rclone.exe in denselben Ordner wie dieses Skript legen oder zum PATH hinzufuegen."
        Stop-Script
    }

    try {
        $rcloneVerOutput = & $RcloneExe version 2>&1 | Select-Object -First 1
        if ($rcloneVerOutput -match 'v(\d+\.\d+)') {
            $rcloneVer = [version]$Matches[1]
            Write-Host "  rclone-Version: $($Matches[0])"
            if ($rcloneVer -lt [version]'1.62') {
                Write-Warning "rclone $($Matches[0]) erkannt -- --metadata erfordert mindestens v1.62."
                Write-Warning "Bitte rclone aktualisieren: https://rclone.org/downloads/"
                Stop-Script
            }
        }
    } catch {
        Write-Warning "rclone-Version konnte nicht ermittelt werden: $_"
        Write-Warning "Bitte sicherstellen, dass rclone >= 1.62 installiert ist (--metadata)."
    }

    try {
        $remoteOutput = @(& $RcloneExe listremotes 2>&1)
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "rclone listremotes meldet Fehler (Exit-Code $LASTEXITCODE):"
            $remoteOutput | ForEach-Object { Write-Warning "  $_" }
            $configuredRemotes = @()
        } else {
            # Nur echte Remote-Namen uebernehmen (enden auf ':') -- haelt
            # evtl. zwischengemischte Warnzeilen aus der Liste heraus.
            $configuredRemotes = @($remoteOutput |
                ForEach-Object { "$_".Trim() } |
                Where-Object { $_ -match ':$' })
        }
    } catch {
        Write-Warning "rclone-Remotes konnten nicht geprueft werden: $_"
        $configuredRemotes = @()
    }

    $RcloneSettings = Get-OptimalRcloneSettings
    Write-Host ("  Erkannter RAM: {0} GB -> rclone-Profil: --transfers={1} --drive-chunk-size={2}" -f `
        $RcloneSettings.RamGB, $RcloneSettings.Transfers, $RcloneSettings.ChunkSize) -ForegroundColor DarkGray
}

# ==================================================================
# LOG-KONFIGURATION
# ==================================================================
$MethodLabel = if ($MethodChoice -eq "1") { "Robocopy" } else { "rclone" }
$scriptDir   = Get-ScriptDirectory
$LogName     = "8_verschieben_auf_Google_Drive_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss')_$MethodLabel.log"

# Vorher wurde ausschliesslich in das Skriptverzeichnis geschrieben. Lag das
# Skript auf einer schreibgeschuetzten Freigabe oder auf einem Netzlaufwerk
# ohne Schreibrecht, brach der Lauf hier ab - obwohl das Protokoll gar nicht
# dort liegen muss. Jetzt: Skriptordner, dann LOCALAPPDATA, dann TEMP. Der
# Schreibtest ist echt (Datei anlegen und loeschen), denn ein vorhandenes
# Verzeichnis sagt nichts ueber das Schreibrecht aus.
$LogFile = $null
foreach ($cand in @($scriptDir,
                    (Join-Path $env:LOCALAPPDATA 'GDriveMigration'),
                    $env:TEMP)) {
    if ([string]::IsNullOrWhiteSpace($cand)) { continue }
    try {
        if (-not (Test-Path -LiteralPath $cand)) {
            New-Item -ItemType Directory -Path $cand -Force -ErrorAction Stop | Out-Null
        }
        $probe = Join-Path $cand ".logtest_$([guid]::NewGuid().Guid).tmp"
        [System.IO.File]::WriteAllText($probe, 'x')
        Remove-Item -LiteralPath $probe -Force -ErrorAction Stop
        $LogFile = Join-Path $cand $LogName
        break
    } catch { continue }
}
if (-not $LogFile) {
    Write-Warning "Kein beschreibbares Verzeichnis fuer die Logdatei gefunden."
    Write-Warning "Geprueft: '$scriptDir', LOCALAPPDATA, TEMP."
    Write-Warning "Ohne Protokoll waere die Migration nicht nachvollziehbar - Abbruch."
    Stop-Script
}
$LogDir      = Split-Path -Path $LogFile -Parent
$LongLogFile = Add-LongPathPrefix $LogFile
if ($LogDir -ne $scriptDir) {
    Write-Host "Hinweis: Protokoll wird nach '$LogDir' geschrieben (Skriptordner nicht beschreibbar)." -ForegroundColor Yellow
}

$TimestampFile  = [System.IO.Path]::ChangeExtension($LogFile, '.timestamps.ndjson')
$ChecksumFile   = [System.IO.Path]::ChangeExtension($LogFile, '.checksums.csv')
$LightCheckFile = [System.IO.Path]::ChangeExtension($LogFile, '.lightcheck.csv')
$ReparseFile    = [System.IO.Path]::ChangeExtension($LogFile, '.uebersprungene-links.csv')

# ==================================================================
# OPTIONEN
# ==================================================================
$RestoreAccessTime = $false
$SkipReparsePoints = $true
$DryRun            = $false

Write-Host ""
Write-Host "Probelauf (Dry-Run) durchfuehren? [j/N]"
Write-Host "(Zeigt nur, was uebertragen wuerde -- es wird nichts kopiert,"
Write-Host " umbenannt oder geloescht.)"
$dryChoice = Read-Host
$DryRun = ($dryChoice -match '^[JjYy]$')

if ($MethodChoice -eq "1") {
    $CopyFlags = "/COPY:DT"
    if (-not $DryRun) {
        Write-Host ""
        Write-Host "Sollen Dateiattribute (Schreibgeschuetzt, Versteckt) ebenfalls uebertragen werden? [j/N]"
        $AttrChoice = Read-Host
        $CopyFlags  = if ($AttrChoice -match '^[JjYy]$') { "/COPY:DAT" } else { "/COPY:DT" }

        Write-Host ""
        Write-Host "Soll zusaetzlich zu Creation/Modified-Time auch der 'Letzter Zugriff'"
        Write-Host "(LastAccessTime) wiederhergestellt werden? [j/N]"
        Write-Host "(Bei 'N' werden nur Erstellungs- und Aenderungsdatum gesetzt -- empfohlen,"
        Write-Host " da dies die Drive-Sync-Last deutlich reduziert.)"
        $accChoice = Read-Host
        $RestoreAccessTime = ($accChoice -match '^[JjYy]$')
    }

    Write-Host ""
    Write-Host "Symlinks und Reparse Points beim Kopieren ueberspringen? [J/n]"
    Write-Host "(Empfohlen: J -- verhindert Endlosschleifen und Mehrfach-Kopien)"
    $rpChoice = Read-Host
    $SkipReparsePoints = -not ($rpChoice -match '^[Nn]$')
} else {
    Write-Host ""
    Write-Host "Symlinks und Reparse Points beim Kopieren ueberspringen? [J/n]"
    Write-Host "(Empfohlen: J)"
    $rpChoice = Read-Host
    $SkipReparsePoints = -not ($rpChoice -match '^[Nn]$')
}

# ==================================================================
# QUELLVERZEICHNIS-AUSWAHL
# ==================================================================
$CurrentUser    = $env:USERNAME
$DesktopPath    = Get-UserShellFolder -Name 'Desktop'
$DownloadsPath  = Get-UserShellFolder -Name 'Downloads'
$SourceIsUserFolder = $false

# Das Menue wird aus dieser Liste aufgebaut, statt die Nummern fest zu
# verdrahten. Vorher standen leere Konfigurationswerte als leere Zeilen im
# Menue, und wer sie waehlte, bekam die nichtssagende Meldung "Es wurde
# kein Quellpfad eingegeben". Nicht konfigurierte Praesets fallen jetzt
# ganz heraus, die Nummerierung bleibt trotzdem lueckenlos.
$SourcePresets = New-Object System.Collections.Generic.List[object]
$SourcePresets.Add(@{ Label = 'Q:\'; Path = 'Q:\'; IsUserFolder = $false })
$SourcePresets.Add(@{ Label = 'R:\'; Path = 'R:\'; IsUserFolder = $false })
if (-not [string]::IsNullOrWhiteSpace($UncBaseQ)) {
    $SourcePresets.Add(@{ Label = $UncBaseQ; Path = $UncBaseQ; IsUserFolder = $false })
}
if (-not [string]::IsNullOrWhiteSpace($UncBaseR)) {
    $SourcePresets.Add(@{ Label = $UncBaseR; Path = $UncBaseR; IsUserFolder = $false })
}
$SourcePresets.Add(@{ Label = $DesktopPath;   Path = $DesktopPath;   IsUserFolder = $true })
$SourcePresets.Add(@{ Label = $DownloadsPath; Path = $DownloadsPath; IsUserFolder = $true })

$manualChoice = $SourcePresets.Count + 1

Write-Host "`n================================================"
Write-Host "  QUELLVERZEICHNIS WAEHLEN"
Write-Host "================================================"
for ($i = 0; $i -lt $SourcePresets.Count; $i++) {
    Write-Host ("  [{0}] {1}" -f ($i + 1), $SourcePresets[$i].Label)
}
Write-Host ("  [{0}] Manuelle Eingabe" -f $manualChoice)
Write-Host "================================================`n"
$SourceChoice = (Read-Host ("Auswahl (1 bis {0})" -f $manualChoice)).Trim()

$choiceNum = 0
if (-not [int]::TryParse($SourceChoice, [ref]$choiceNum)) { $choiceNum = -1 }

if ($choiceNum -ge 1 -and $choiceNum -le $SourcePresets.Count) {
    $preset             = $SourcePresets[$choiceNum - 1]
    $InputPath          = $preset.Path
    $SourceIsUserFolder = $preset.IsUserFolder
} elseif ($choiceNum -eq $manualChoice) {
    $hintUnc = if ($UncBaseQ) { " oder $UncBaseQ\Ordner" } else { "" }
    Write-Host "Bitte Quellordner eingeben (z. B. Q:\Ordner, R:\Ordner, C:\Users\$CurrentUser\Desktop\Ordner$hintUnc):"
    $InputPath = Read-Host
} else {
    Write-Warning ("Ungueltige Auswahl. Bitte 1 bis {0} eingeben." -f $manualChoice)
    Stop-Script
}

if ([string]::IsNullOrWhiteSpace($InputPath)) {
    Write-Warning "Es wurde kein Quellpfad eingegeben."
    Stop-Script
}

$SourcePath = $InputPath.Trim().Trim('"').Trim("'")
$SourcePath = $SourcePath.TrimEnd("\")
if ($SourcePath.Length -eq 2 -and $SourcePath -match '^[A-Za-z]:$') {
    $SourcePath = $SourcePath + "\"
}

$longSourceTest = Add-LongPathPrefix $SourcePath
if (-not ([System.IO.Directory]::Exists($longSourceTest) -or [System.IO.File]::Exists($longSourceTest))) {
    Write-Warning "Der Pfad '$SourcePath' wurde nicht gefunden!"
    Write-Warning "Tipp fuer Administratoren: Bei erhoehten Rechten ('Als Administrator' gestartet) sind"
    Write-Warning "Netzlaufwerksbuchstaben (Q:, R:) oft ausgeblendet. Bitte Option 3 oder 4 (UNC-Pfad) nutzen."
    Stop-Script
}

# Vorher wurde nur auf Existenz geprueft. Zeigte die Eingabe auf eine
# DATEI, lief Robocopy in einen unverstaendlichen Fehler, und der
# anschliessende /MIR-Purge haette das Elternverzeichnis getroffen.
if (-not [System.IO.Directory]::Exists($longSourceTest)) {
    Write-Warning "'$SourcePath' ist keine Ordner-, sondern eine Dateiangabe."
    Write-Warning "Bitte den uebergeordneten Ordner angeben."
    Stop-Script
}

# Sicherheitsnetz gegen Systemlaufwerke, Profil- und Programmordner.
$srcUnsafeReason = Test-SourcePathSafe -Path $SourcePath
if ($srcUnsafeReason) {
    Write-Warning "Quellverzeichnis abgelehnt:"
    Write-Warning "  $srcUnsafeReason"
    Write-Warning "Am Ende des Vorgangs wird die Quelle geloescht - deshalb sind solche"
    Write-Warning "Pfade grundsaetzlich gesperrt. Bitte einen konkreten Datenordner waehlen."
    Stop-Script
}

if ($SourcePath -notmatch '^[A-Za-z]:\\' -and $SourcePath -notmatch '^\\\\[^\\]+\\[^\\]+') {
    Write-Warning "Ungueltiger Pfad: '$SourcePath'."
    Write-Warning "Unterstuetzt: lokale Laufwerkspfade (z.B. Q:\Ordner) und UNC-Pfade (z.B. $UncBaseQ\Ordner)."
    Stop-Script
}

$DriveLetter = $null
if ($SourcePath -match '^[A-Za-z]:') {
    # Split-Path -Qualifier wirft bei UNC-Pfaden einen Fehler -
    # daher nur fuer Laufwerkspfade aufrufen.
    $DriveLetter = Split-Path -Path $SourcePath -Qualifier
}

# Der Leer-Guard ist zwingend: bei leerem $UncBaseQ liefert StartsWith('')
# immer $true, wodurch JEDE Quelle als Geteilte Ablage behandelt und in das
# falsche Ziel migriert wuerde.
# Vergleich ueber Test-PathInside (Gleichheit oder Praefix MIT '\'):
# StartsWith ohne Trennzeichen hielt '\\server\dfsarchiv\...' fuer einen
# Pfad unter '\\server\dfs' und migrierte ihn in die Geteilte Ablage
# (nachgestellt).
$isQDrive = ($DriveLetter -eq "Q:") -or
            ($UncBaseQ -and (Test-PathInside -Child $SourcePath -Parent $UncBaseQ))
$isRDrive = ($DriveLetter -eq "R:") -or ($UncBaseR -and (Test-PathInside -Child $SourcePath -Parent $UncBaseR))

# ==================================================================
# ZIELPFAD-ERMITTLUNG
# ==================================================================
$LocalSharedDrives = if (Test-Path -LiteralPath "G:\Geteilte Ablagen") { "G:\Geteilte Ablagen" }
                     elseif (Test-Path -LiteralPath "G:\Shared drives")  { "G:\Shared drives" }
                     else { $null }
$LocalMyDrive      = if (Test-Path -LiteralPath "G:\Meine Ablage") { "G:\Meine Ablage" }
                     elseif (Test-Path -LiteralPath "G:\My Drive")  { "G:\My Drive" }
                     else { $null }

if ($isQDrive) {
    $DestBase = if ($MethodChoice -eq "1") { $LocalSharedDrives } else { $RcloneRemoteQ }
} elseif ($isRDrive -or $SourceIsUserFolder) {
    $DestBase = if ($MethodChoice -eq "1") { $LocalMyDrive } else { $RcloneRemoteMyDrive }
} else {
    $srcDesc = if ($DriveLetter) { "Laufwerk $DriveLetter" } else { "Quelle '$SourcePath'" }
    Write-Warning "$srcDesc ist nicht als Standard konfiguriert."
    if ($MethodChoice -eq "1") {
        $DestBase = (Read-Host "Bitte Ziel-Basispfad manuell eingeben (z.B. G:\Geteilte Ablagen\HR oder G:\Shared drives\HR)").Trim().Trim('"').Trim("'")
    } else {
        $DestBase = (Read-Host "Bitte rclone-Ziel manuell eingeben (z.B. HR_Drive:)").Trim().Trim('"').Trim("'")
    }
}

if ([string]::IsNullOrWhiteSpace($DestBase)) {
    Write-Warning "Zielpfad konnte nicht ermittelt werden. Ist Google Drive gestartet und als G: eingebunden?"
    Write-Warning "Abbruch."
    Stop-Script
}

$isRootPath = if ($DriveLetter) {
    ($SourcePath.TrimEnd('\') -eq $DriveLetter.TrimEnd('\'))
} else {
    # UNC-Quelle: Root liegt vor, wenn die Quelle exakt einer der
    # konfigurierten Share-Basen entspricht ODER wenn sie ueberhaupt nur
    # aus \\Server\Freigabe besteht. Der zweite Fall fehlte: bei einer
    # frei eingegebenen Freigabewurzel galt sie als Unterordner, und das
    # Skript versuchte anschliessend, die Freigabe selbst zu loeschen.
    (($UncBaseQ -and ($SourcePath.TrimEnd('\') -ieq $UncBaseQ.TrimEnd('\'))) -or
     ($UncBaseR -and ($SourcePath.TrimEnd('\') -ieq $UncBaseR.TrimEnd('\'))) -or
     ($SourcePath.TrimEnd('\') -match '^\\\\[^\\]+\\[^\\]+$'))
}

# ==================================================================
# ZIELPFAD-AUSWAHL
# ==================================================================
$SourceLeaf = if ($isRootPath) { "" } else { Split-Path $SourcePath -Leaf }

Write-Host "`n================================================"
Write-Host "  ZIELPFAD WAEHLEN"
Write-Host "================================================"
Write-Host "  [a] Automatisch: Quell-Ordnername als Zielname verwenden"
if ($SourceLeaf) {
    Write-Host "      -> $DestBase\$SourceLeaf"
} else {
    Write-Host "      -> $DestBase  (Stammverzeichnis)"
}
Write-Host "  [b] Individuell: eigenen Unterordner unter '$DestBase' angeben"
Write-Host "      Beispiel: 'g-d'  oder  'XYZ\ABC'"
Write-Host "================================================`n"
$DestChoice = (Read-Host "Auswahl (a oder b)").Trim().ToLower()

if ($DestChoice -eq "a") {
    $FinalDest = if ($MethodChoice -eq "1") {
        if ($SourceLeaf) { Join-Path $DestBase $SourceLeaf } else { $DestBase }
    } else {
        if ($SourceLeaf) {
            $cleanLeaf = $SourceLeaf -replace '\\', '/'
            if ($DestBase.EndsWith(':')) { $DestBase + $cleanLeaf }
            else { $DestBase.TrimEnd('/') + '/' + $cleanLeaf }
        } else { $DestBase }
    }
} elseif ($DestChoice -eq "b") {
    Write-Host "Unterordner unter '$DestBase' eingeben:"
    Write-Host "(z.B. 'g-d'  oder  'XYZ\ABC')"
    $SubInput = (Read-Host).Trim().Trim('"').Trim("'")
    if ([string]::IsNullOrWhiteSpace($SubInput)) {
        Write-Warning "Kein Zielpfad eingegeben."
        Stop-Script
    }
    if ($SubInput -match '\.\.' -or
        $SubInput -match '^[A-Za-z]:' -or
        $SubInput -match '^\\\\') {
        Write-Warning "Pfadwechsel (..) und absolute Pfade sind als Ziel-Unterordner blockiert."
        Stop-Script
    }
    $FinalDest = if ($MethodChoice -eq "1") {
        Join-Path $DestBase $SubInput
    } else {
        $cleanSub = $SubInput -replace '\\', '/'
        if ($DestBase.EndsWith(':')) { $DestBase + $cleanSub }
        else { $DestBase.TrimEnd('/') + '/' + $cleanSub }
    }
} else {
    Write-Warning "Ungueltige Auswahl. Bitte 'a' oder 'b' eingeben."
    Stop-Script
}

# ==================================================================
# SCHUTZ: QUELLE/ZIEL-UEBERLAPPUNG (nur Robocopy -- Dateisystem-Ziel)
# ==================================================================
# Liegt das Ziel innerhalb der Quelle (oder umgekehrt bzw. identisch),
# droht beim Kopieren eine Endlosrekursion und beim anschliessenden
# Loeschen der Quelle Datenverlust im Ziel.
if ($MethodChoice -eq "1") {
    $srcCmp = $SourcePath.TrimEnd('\') + '\'
    $dstCmp = $FinalDest.TrimEnd('\') + '\'
    if ($dstCmp.StartsWith($srcCmp, [System.StringComparison]::OrdinalIgnoreCase) -or
        $srcCmp.StartsWith($dstCmp, [System.StringComparison]::OrdinalIgnoreCase)) {
        Write-Warning "Quelle und Ziel ueberlappen sich (oder sind identisch):"
        Write-Warning "  Quelle: $SourcePath"
        Write-Warning "  Ziel:   $FinalDest"
        Write-Warning "Das Ziel darf nicht innerhalb der Quelle liegen (und umgekehrt)."
        Write-Warning "Abbruch -- es wurde nichts kopiert oder geloescht."
        Stop-Script
    }
}

# ==================================================================
# GOOGLE DRIVE BEREINIGUNG & KOMPATIBILITAETSPRUEFUNG
# ==================================================================
$gdProblems = 0
if ($MethodChoice -eq "1") {
    Write-Host ""
    Write-Host "================================================"
    Write-Host "  GOOGLE DRIVE DATEINAMEN-BEREINIGUNG"
    Write-Host "================================================"
    Write-Host ""
    Write-Host "  Der Google Drive Desktop Client (G:\) erzwingt Windows-"
    Write-Host "  Namenskonventionen. Folgende Probleme werden automatisch"
    Write-Host "  korrigiert:"
    Write-Host "    - Fuehrende Leerzeichen, abschliessende Leerzeichen/Punkte"
    Write-Host "    - Verbotene Zeichen (: * ? `" < > |) -> '-'"
    Write-Host "    - Reservierte Namen (CON, PRN, NUL, ...) -> '_CON'"
    Write-Host ""
    Write-Host "  Fuehrende Punkte (z.B. .gitignore, .env) bleiben erhalten."
    Write-Host "  Alle Umbenennungen werden in einer CSV-Datei protokolliert."
    Write-Host ""
    if ($DryRun) {
        Write-Host "Probelauf: Namensbereinigung wird uebersprungen -- es folgt nur die Pruefung."
    } else {
        Write-Host "Sollen unzulaessige Dateinamen automatisch korrigiert werden? [J/N]"
        Write-Host "(Bei 'N' wird nur geprueft, nicht geaendert.)"
        $repairChoice = Read-Host

        if ($repairChoice -match '^[JjYy]$') {
            Write-Host ""
            $repairCount = Repair-GDrivePaths -RootPath $SourcePath -LogFile $LogFile -SkipReparsePoints $SkipReparsePoints
            Stop-WennAbgebrochen -Phase 'Dateinamen-Bereinigung' -Encoding ([System.Text.Encoding]::Unicode)
            if ($repairCount -gt 0) {
                Write-Host ""
                Write-Host "Bereinigung abgeschlossen. Fuehre jetzt Kontrollpruefung durch ..."
            }
        }
    }

    Write-Host ""
    $gdProblems = Test-GDriveCompatibility -RootPath $SourcePath -LogFile $LogFile -SkipReparsePoints $SkipReparsePoints -DestRoot $FinalDest
    Stop-WennAbgebrochen -Phase 'Google-Drive-Kompatibilitaetspruefung' -Encoding ([System.Text.Encoding]::Unicode)
    if ($gdProblems -gt 0) {
        Write-Host ""
        Write-Host "Es wurden $gdProblems verbleibende Kompatibilitaetsprobleme gefunden." -ForegroundColor Red
        Write-Host ""
        Write-Host "Hinweis zu langen Pfaden und Dateinamen:" -ForegroundColor Yellow
        Write-Host "  Robocopy und Google Drive for Desktop koennen grundsaetzlich auch mit sehr" -ForegroundColor Yellow
        Write-Host "  langen Pfaden (> 260 Zeichen) und langen Dateinamen (> 255 Zeichen) umgehen," -ForegroundColor Yellow
        Write-Host "  sofern auf dem System 'LongPathsEnabled' aktiviert ist." -ForegroundColor Yellow
        Write-Host "  Ist diese Windows-Einstellung NICHT aktiv, koennen solche Dateien nicht" -ForegroundColor Yellow
        Write-Host "  kopiert werden und muessen manuell umbenannt oder in kuerzere Verzeichnisse" -ForegroundColor Yellow
        Write-Host "  verschoben werden." -ForegroundColor Yellow
        Write-Host "  Pruefung: HKLM\SYSTEM\CurrentControlSet\Control\FileSystem\LongPathsEnabled" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Sicherheit beim Loeschen:" -ForegroundColor Yellow
        Write-Host "  Nach dem Kopiervorgang findet ein Abgleich statt. Vom Quelllaufwerk werden" -ForegroundColor Yellow
        Write-Host "  ausschliesslich Dateien geloescht, die auch im Ziel nachweislich vorhanden sind." -ForegroundColor Yellow
        Write-Host "  Dateien, die nicht kopiert werden konnten, bleiben im Quellordner erhalten." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Details zu den gefundenen Problemen stehen in der CSV-Datei." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Trotzdem fortfahren? [j/N]" -ForegroundColor Yellow
        $gdContinue = Read-Host
        if ($gdContinue -notmatch '^[JjYy]$') {
            Write-Host "Vorgang abgebrochen. Bitte verbleibende Probleme manuell beheben."
            Stop-Script
        }
    }
}

# ==================================================================
# RCLONE-REMOTE VALIDIERUNG
# ==================================================================
if ($MethodChoice -eq "2") {
    if ($FinalDest -match '^([^:]+:)') {
        $usedRemote      = $Matches[1]
        $usedRemoteLower = $usedRemote.ToLowerInvariant()
        $remotesLower    = @($configuredRemotes | ForEach-Object { $_.ToLowerInvariant() })
        if ($remotesLower -notcontains $usedRemoteLower) {
            Write-Warning "Das ausgewaehlte rclone-Remote '$usedRemote' ist auf diesem System nicht konfiguriert!"
            Write-Warning ""
            Write-Warning "Aktuell konfigurierte Remotes:"
            if ($configuredRemotes.Count -eq 0) {
                Write-Warning "  (keine)"
            } else {
                foreach ($r in $configuredRemotes) { Write-Warning "  - $r" }
            }
            Write-Warning ""
            Write-Warning "Bitte mit 'rclone config' einrichten oder ein anderes Ziel waehlen."
            Stop-Script
        }
        Write-Host "rclone-Remote '$usedRemote' ist konfiguriert." -ForegroundColor DarkGray
    }
}

# ==================================================================
# VORAB-BESTAETIGUNG
# ==================================================================
Write-Host "`n------------------------------------------------"
Write-Host "METHODE:     $MethodLabel"
if ($DryRun) {
    Write-Host "MODUS:       PROBELAUF -- es wird nichts kopiert, umbenannt oder geloescht" -ForegroundColor Yellow
}
Write-Host "QUELLORDNER: $SourcePath"
Write-Host "ZIELORDNER:  $FinalDest"
Write-Host "LOG-DATEI:   $LogFile"
Write-Host "------------------------------------------------"
Write-Host ""
Write-Host "Ist das Ziel korrekt? [J/N]"
$PathConfirm = Read-Host
if ($PathConfirm -notmatch '^[JjYy]$') {
    Write-Host "Vorgang abgebrochen. Bitte Skript erneut starten und korrekten Pfad waehlen."
    Stop-Script
}
Write-Host ""

# ==================================================================
# LOG-HEADER
# ==================================================================
# Encoding je Methode: Robocopy-/UNILOG-Logs sind UTF-16 (Unicode),
# rclone-Logs UTF-8 mit BOM. Die Log-Pfad-Substitution der Robocopy-
# Methode laesst die Header-Zeilen gezielt unangetastet (Marker '===').
$hdrEnc  = if ($MethodChoice -eq "1") { [System.Text.Encoding]::Unicode } else { $Utf8Bom }
$hdrOpts = "SkipReparsePoints=$SkipReparsePoints, RestoreAccessTime=$RestoreAccessTime"
if ($MethodChoice -eq "1") {
    $hdrOpts += ", CopyFlags=$CopyFlags"
} else {
    $hdrOpts += ", Transfers=$($RcloneSettings.Transfers), ChunkSize=$($RcloneSettings.ChunkSize)"
}
if ($DryRun) { $hdrOpts += ", PROBELAUF" }
foreach ($hdrLine in @(
    "==================================================================",
    "MIGRATIONS-LOG  --  8_verschieben_auf_Google_Drive.ps1",
    "Gestartet : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
    "Benutzer  : $env:USERDOMAIN\$env:USERNAME auf $env:COMPUTERNAME",
    "Methode   : $MethodLabel",
    "Quelle    : $SourcePath",
    "Ziel      : $FinalDest",
    "Optionen  : $hdrOpts",
    "=================================================================="
)) {
    Add-LogLine -Path $LogFile -Message $hdrLine -Encoding $hdrEnc
}

# ==================================================================
# METHODE 1: ROBOCOPY
# ==================================================================
if ($MethodChoice -eq "1") {

    if (-not $DryRun) {
        Write-Host "WICHTIG ZUM CLOUD-SYNC-RISIKO:"
        Write-Host "  G:\ ist ein virtueller Google-Drive-File-Stream-Cache."
        Write-Host "  Robocopy meldet Erfolg sobald Daten im LOKALEN Cache liegen --"
        Write-Host "  nicht erst nach vollstaendiger Cloud-Synchronisation."
        Write-Host ""
        Write-Host "  ACHTUNG: Waehrend des gesamten Vorgangs (Kopieren + Cloud-Sync)" -ForegroundColor Red
        Write-Host "  darf kein Anwender Dateien auf dem Quelllaufwerk oeffnen," -ForegroundColor Red
        Write-Host "  veraendern oder neu anlegen. Nach dem Kopieren werden neue oder" -ForegroundColor Red
        Write-Host "  geaenderte Dateien von Robocopy nicht erfasst und beim Loeschen" -ForegroundColor Red
        Write-Host "  unwiederbringlich vernichtet." -ForegroundColor Red
        Write-Host ""
        Write-Host "ZEITSTEMPEL-HINWEIS:"
        Write-Host "  Robocopy /COPY:DT schreibt Zeitstempel in den lokalen GDrive-Cache."
        Write-Host "  Google Drive kann diese beim Cloud-Sync zuruecksetzen."
        Write-Host "  Die Original-Zeitstempel werden jetzt exportiert und nach der"
        Write-Host "  Cloud-Bestaetigung auf dem Ziel wiederhergestellt."
        Write-Host ""

        # --- Schreibrechte-Test ---
        Write-Host "Pruefe Schreibrechte auf Zielverzeichnis ..."
        try {
            $longFinalDest = Add-LongPathPrefix $FinalDest
            if (-not [System.IO.Directory]::Exists($longFinalDest)) {
                [void][System.IO.Directory]::CreateDirectory($longFinalDest)
            }
            $testFile     = Join-Path $FinalDest ".write_test_$([guid]::NewGuid().Guid).tmp"
            $longTestFile = Add-LongPathPrefix $testFile
            [System.IO.File]::WriteAllText($longTestFile, "rwtest")
            [System.IO.File]::Delete($longTestFile)
            Write-Host "  Schreibrechte OK." -ForegroundColor Green
        } catch {
            Write-Warning "Keine Schreibrechte auf Zielverzeichnis '$FinalDest':"
            Write-Warning "  $_"
            Write-Warning ""
            Write-Warning "Moegliche Ursachen:"
            Write-Warning "  - Google Drive ist nicht gestartet oder nicht als G: eingebunden"
            Write-Warning "  - Der angegebene Geteilte-Ablage-Ordner ist fuer den Nutzer schreibgeschuetzt"
            Write-Warning "  - Drive for Desktop laeuft in einem Cache-only-Modus"
            Add-LogLine -Path $LogFile -Message "[$(Get-Date)] ABBRUCH: Keine Schreibrechte auf '$FinalDest'. Fehler: $_" `
                        -Encoding ([System.Text.Encoding]::Unicode)
            Stop-Script
        }
        Write-Host ""

    }

    # --- Zeitstempel-Export + Quell-Statistik ---
    # Laeuft bewusst AUCH im Probelauf: Die Google-Limit-Warnungen
    # (500.000 Objekte, 100 Ordnerebenen, 5-TB-Dateien, Cache-Platz)
    # speisen sich aus dieser Statistik. Sie waren vorher im Probelauf
    # ausgeblendet - also genau dann nicht sichtbar, wenn man sie zur
    # Planung gebraucht haette.
    if ($DryRun) {
        Write-Host "Ermittle Quell-Statistik fuer die Limit-Pruefung ..."
    } else {
        Write-Host "Exportiere Zeitstempel ..."
    }
    $srcStats = Export-Timestamps -RootPath $SourcePath -OutputFile $TimestampFile `
                    -SkipReparsePoints $SkipReparsePoints -ReparseReportFile $ReparseFile
    Stop-WennAbgebrochen -Phase 'Zeitstempel-Export/Quell-Statistik' -Encoding ([System.Text.Encoding]::Unicode)

    # --- Google-Drive-Limits & Cache-Speicher pruefen ---
    if ($srcStats) {
        $limitWarnings = New-Object System.Collections.Generic.List[string]

        if ($isQDrive) {
            # Geteilte Ablagen: harte Google-Limits (Stand 2026, bitte bei
            # Bedarf gegen die aktuelle Google-Doku pruefen):
            #   ca. 500.000 Objekte und max. 100 Ordner-Ebenen pro Ablage.
            if ($srcStats.Count -gt 500000) {
                $limitWarnings.Add("Quelle enthaelt $($srcStats.Count) Objekte -- Geteilte Ablagen sind auf ca. 500.000 Objekte begrenzt. Die Migration wird sehr wahrscheinlich mittendrin scheitern; bitte auf mehrere Geteilte Ablagen aufteilen.")
            } elseif ($srcStats.Count -gt 400000) {
                $limitWarnings.Add("Quelle enthaelt $($srcStats.Count) Objekte -- nahe am Limit von ca. 500.000 Objekten pro Geteilter Ablage (bereits vorhandene Objekte im Ziel zaehlen mit).")
            }
            if ($srcStats.MaxDepth -gt 95) {
                $limitWarnings.Add("Maximale Ordnertiefe der Quelle: $($srcStats.MaxDepth) Ebenen -- Geteilte Ablagen erlauben max. 100 verschachtelte Ordner.")
            }
        }
        if ($srcStats.HugeFiles -gt 0) {
            $limitWarnings.Add("$($srcStats.HugeFiles) Datei(en) groesser als 5 TB -- Google Drive lehnt solche Dateien ab.")
        }

        # Methode Robocopy schreibt zunaechst in den lokalen Drive-Cache
        # (unter %LOCALAPPDATA%). Reicht der freie Platz dort nicht fuer
        # die Quellgroesse, kann der Cache volllaufen, bevor der Cloud-
        # Sync nachkommt.
        try {
            $cacheRoot  = [System.IO.Path]::GetPathRoot($env:LOCALAPPDATA)
            $cacheDrive = New-Object System.IO.DriveInfo($cacheRoot)
            if ($cacheDrive.AvailableFreeSpace -lt $srcStats.TotalBytes) {
                $srcGB  = [Math]::Round($srcStats.TotalBytes / 1GB, 1)
                $freeGB = [Math]::Round($cacheDrive.AvailableFreeSpace / 1GB, 1)
                $limitWarnings.Add("Quellgroesse ($srcGB GB) uebersteigt den freien Platz auf dem Drive-Cache-Laufwerk $cacheRoot ($freeGB GB frei). Der lokale Cache kann volllaufen, bevor der Cloud-Sync nachkommt -- ggf. in Tranchen migrieren.")
            }
        } catch { }

        if ($limitWarnings.Count -gt 0) {
            Write-Host ""
            foreach ($w in $limitWarnings) {
                Write-Warning $w
                Add-LogLine -Path $LogFile -Message "[$(Get-Date)] WARNUNG: $w" `
                            -Encoding ([System.Text.Encoding]::Unicode)
            }
            Write-Host ""
            Write-Host "Trotzdem fortfahren? [j/N]" -ForegroundColor Yellow
            $limitContinue = Read-Host
            if ($limitContinue -notmatch '^[JjYy]$') {
                Write-Host "Vorgang abgebrochen. Es wurde nichts kopiert oder geloescht."
                Add-LogLine -Path $LogFile -Message "[$(Get-Date)] ABBRUCH: Nutzer hat nach Limit-Warnungen abgebrochen." `
                            -Encoding ([System.Text.Encoding]::Unicode)
                Stop-Script
            }
        }
    }

    # --- Robocopy-Aufruf ---
    $RoboArgs = @(
        $SourcePath,
        $FinalDest,
        "/E",
        "/R:5", "/W:5",
        $CopyFlags,
        "/DCOPY:T",
        "/NP",
        "/FP",
        "/NC",   # keine Klassenspalte ("Neue Datei" etc.) -- der Fortschritts-Ticker unten erwartet "Groesse Pfad"
        "/NDL",
        "/XD") + $ExcludedDirNames + @(
        "/UNILOG+:$LongLogFile",
        "/TEE"
    )
    if ($SkipReparsePoints) { $RoboArgs += @("/XJ") }
    if ($DryRun)            { $RoboArgs += @("/L") }

    # Letzter Pruefpunkt vor dem Kopieren (Strg+C waehrend einer Rueckfrage).
    Stop-WennAbgebrochen -Phase 'Vorbereitung vor dem Kopieren' -Encoding ([System.Text.Encoding]::Unicode)

    Write-Host "Starte Kopiervorgang ..."
    Write-Host ""
    $colWidth = try { [Math]::Max(40, [Console]::WindowWidth - 6) } catch { 80 }
    & robocopy $RoboArgs | ForEach-Object {
        if ($_ -match '^\s+\d+\s+(?:[A-Za-z]:\\|\\\\)') {
            $disp = $_.TrimStart()
            if ($disp.Length -gt $colWidth) { $disp = '...' + $disp.Substring($disp.Length - $colWidth + 3) }
            Write-Host ("`r  " + $disp.PadRight($colWidth)) -NoNewline
        } else {
            $line = $_ -replace '\bInsgesamt\b', 'In Quelle' -replace 'KopiertÜbersprungenKeine Übereinstimmung', 'Kopiert  ber. vorh.   Namenskonflikt  ' -replace '\bExtras\b', 'Nur im Ziel'
            Write-Host ("`r" + (' ' * ($colWidth + 4)) + "`r" + $line)
        }
    }
    Write-Host ""
    $RoboExit = $LASTEXITCODE

    # --- Log-Pfad-Substitution ---
    Start-Sleep -Seconds 1
    if (Test-Path -LiteralPath $LogFile) {
        $sourceEscaped      = [regex]::Escape($SourcePath.TrimEnd('\'))
        $finalDestClean     = $FinalDest.TrimEnd('\')
        $finalDestCleanSafe = $finalDestClean.Replace('$', '$$')
        $tmpFile    = "$LogFile.tmp"
        $reader     = $null
        $writer     = $null
        $logSuccess = $false
        try {
            $reader  = [System.IO.StreamReader]::new($LogFile, [System.Text.Encoding]::Unicode)
            $writer  = [System.IO.StreamWriter]::new($tmpFile, $false, [System.Text.Encoding]::Unicode)
            # Zweck der Ersetzung: die Dateizeilen (/FP /NC: Einrueckung,
            # Groesse, TAB, voller Quellpfad) sollen zeigen, WOHIN die Datei
            # kopiert wurde. Nur diese Zeilen werden umgeschrieben. Vorher
            # traf die Ersetzung jeden Quellpfad hinter einem Leerzeichen -
            # auch Robocopys eigene Kopfzeile '   Quelle : ...' und die
            # FEHLER-Zeilen ('... FEHLER 32 (0x00000020) Folgende Datei wird
            # kopiert <Quellpfad>'). Im Log standen dann Quelle und Ziel mit
            # demselben Pfad, und die fehlgeschlagenen Quelldateien waren
            # nicht mehr auffindbar (nachgestellt mit robocopy zwischen zwei
            # Temp-Ordnern und einer gesperrten Datei). Groessenangabe wie
            # von robocopy geschrieben: '3', '3.0 m', je nach Sprache ',' .
            $regex   = [regex]::new('^(\s+[0-9][0-9.,]*(?:\s[kmgt])?\t)' + $sourceEscaped + '(?=\\|$)',
                                    [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            $ersatz  = '${1}' + $finalDestCleanSafe
            # Alles bis einschliesslich der zweiten '==='-Markerzeile ist der
            # Skript-Log-Header -- er nennt bewusst den Quellpfad und darf
            # nicht durch die Substitution verfaelscht werden.
            $headerSeps = 0
            while ($null -ne ($line = $reader.ReadLine())) {
                if ($headerSeps -lt 2) {
                    if ($line -match '^={20,}') { $headerSeps++ }
                    $writer.WriteLine($line)
                    continue
                }
                $writer.WriteLine($regex.Replace($line, $ersatz, 1))
            }
            $logSuccess = $true
        } catch {
            Write-Warning "Log konnte nicht formatiert werden: $_"
        } finally {
            if ($reader) { $reader.Close() }
            if ($writer) { $writer.Close() }
        }
        if ($logSuccess) {
            try {
                Move-Item -LiteralPath $tmpFile -Destination $LogFile -Force -ErrorAction Stop
            } catch {
                Write-Warning "Fehler beim Ersetzen der Log-Datei: $_"
            }
        }
        # Zwischendatei bei Fehlschlag entfernen. Sie blieb sonst als
        # '<log>.log.tmp' neben dem Protokoll liegen und wurde von keinem
        # Aufraeumpfad erfasst - bei jedem gescheiterten Lauf eine weitere.
        if (Test-Path -LiteralPath $tmpFile) {
            Remove-Item -LiteralPath $tmpFile -Force -ErrorAction SilentlyContinue
        }
    }

    # Strg+C waehrend robocopy: robocopy haengt an derselben Konsole und
    # endet mit 0xC000013A (STATUS_CONTROL_C_EXIT), in $LASTEXITCODE als
    # negative Zahl -1073741510 (gemessen). Vorher lief der Ablauf danach
    # als "Kopiervorgang abgeschlossen" weiter bis zu den Rueckfragen.
    Stop-WennAbgebrochen -Phase "Kopiervorgang (Robocopy, Exit-Code $RoboExit)" -Encoding ([System.Text.Encoding]::Unicode)

    # --- Probelauf: hier ist Schluss ---
    if ($DryRun) {
        Write-Host ""
        if ($RoboExit -ge 16 -or $RoboExit -lt 0) {
            Write-Warning "Robocopy meldet einen kritischen Fehler (Exit-Code $RoboExit) -- bitte Log pruefen."
        }
        Write-Host "PROBELAUF abgeschlossen -- es wurde nichts kopiert, umbenannt oder geloescht." -ForegroundColor Green
        Write-Host "Was uebertragen wuerde, steht im Log: $LogFile"
        Add-LogLine -Path $LogFile -Message "[$(Get-Date)] PROBELAUF beendet (Robocopy /L, Exit-Code $RoboExit). Keine Aenderungen." `
                    -Encoding ([System.Text.Encoding]::Unicode)
        Stop-Script
    }

    # --- Robocopy-Exit-Code-Bewertung ---
    # Negative Exit-Codes sind NTSTATUS-Werte eines gewaltsam beendeten
    # robocopy (0xC000013A nach Strg+C, gemessen -1073741510; ebenso ein
    # Absturz wie 0xC0000005). Sie fielen durch '-ge 16', '-ge 8' und
    # '-ge 4' hindurch und galten als fehlerfreier Lauf - endete robocopy
    # ohne gesetztes Strg+C-Flag, gab Option [3] danach das Komplett-
    # Loeschen der Quelle frei. Jetzt: kritischer Fehler.
    if ($RoboExit -ge 16 -or $RoboExit -lt 0) {
        Write-Host ""
        Write-Warning "Robocopy meldet einen KRITISCHEN FEHLER (Exit-Code $RoboExit)."
        Write-Warning "  Ursache: Laufwerk nicht erreichbar, ungueltiger Pfad oder schwerwiegendes Problem."
        Write-Warning "  Die Quelldateien werden NICHT geloescht."
        Write-Warning "  Bitte pruefen: Ist Google Drive gestartet? Ist G: eingebunden?"
        Write-Warning "  Log: $LogFile"
        Add-LogLine -Path $LogFile -Message "[$(Get-Date)] ABBRUCH: Robocopy Exit-Code $RoboExit (kritisch) -- Quelle wurde nicht geloescht." `
                    -Encoding ([System.Text.Encoding]::Unicode)
        Stop-Script
    }

    if ($RoboExit -ge 8) {
        Write-Host ""
        Write-Warning "Robocopy meldet Kopierfehler bei einzelnen Dateien (Exit-Code $RoboExit)."
        Write-Warning "  Ursache: Zugriff verweigert, Datei gesperrt oder Schreibfehler."
        Write-Warning "  Einige Dateien wurden moeglicherweise NICHT kopiert."
        Write-Warning "  Log: $LogFile"
        Write-Host ""
        Write-Host "Trotzdem mit Cloud-Pruefung und Loeschen fortfahren? [j/N]"
        Write-Host "(Bei 'N' bleibt die Quelle vollstaendig erhalten.)"
        $roboErrorContinue = Read-Host
        if ($roboErrorContinue -notmatch '^[JjYy]$') {
            Write-Host "Vorgang abgebrochen. Quelldateien bleiben erhalten: $SourcePath"
            Add-LogLine -Path $LogFile -Message "[$(Get-Date)] ABBRUCH: Nutzer hat bei Robocopy Exit-Code $RoboExit abgebrochen. Quelle bleibt erhalten." `
                        -Encoding ([System.Text.Encoding]::Unicode)
            Write-Host "`nVorgang abgeschlossen. Details: $LogFile" -ForegroundColor Green
            Stop-Script
        }
        Write-Host ""
        Write-Warning "Nutzer hat trotz Kopierfehlern die Fortsetzung bestaetigt."
        Add-LogLine -Path $LogFile -Message "[$(Get-Date)] WARNUNG: Nutzer hat bei Robocopy Exit-Code $RoboExit die Fortsetzung bestaetigt." `
                    -Encoding ([System.Text.Encoding]::Unicode)
    }

    if ($RoboExit -ge 4 -and $RoboExit -lt 8) {
        Write-Host ""
        Write-Host "Robocopy meldet Namenskonflikte oder Typ-/Groessen-Abweichungen (Exit-Code $RoboExit)." -ForegroundColor Yellow
        Write-Host "  Ursache: In der Quelle existiert eine Datei, aber im Ziel gibt es bereits" -ForegroundColor Yellow
        Write-Host "  einen Ordner mit exakt demselben Namen (oder umgekehrt), oder eine Datei" -ForegroundColor Yellow
        Write-Host "  existiert in beiden Seiten mit unterschiedlichem Typ/Groesse." -ForegroundColor Yellow
        Write-Host "  Diese betroffenen Elemente wurden NICHT kopiert, um die Verzeichnisstruktur" -ForegroundColor Yellow
        Write-Host "  nicht zu zerstoeren." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  --> EMPFEHLUNG: Pruefen Sie das Log auf die uebersprungenen Elemente." -ForegroundColor Yellow
        Write-Host "      (Details: $LogFile)" -ForegroundColor Yellow

        Add-LogLine -Path $LogFile -Message "[$(Get-Date)] HINWEIS: Robocopy Exit-Code $RoboExit (Konflikte/Mismatch). Einzelne Dateien wurden uebersprungen." `
                    -Encoding ([System.Text.Encoding]::Unicode)
    }

    Write-Host ""
    Write-Host "Kopiervorgang abgeschlossen (Robocopy Exit-Code $RoboExit)." -ForegroundColor Green
    Write-Host ""
    Write-Host "Bitte pruefen Sie jetzt in drive.google.com, ob alle Dateien"
    Write-Host "vollstaendig in der Cloud angekommen sind."
    Write-Host ""

    Write-Host "Haben Sie den Upload in der Cloud geprueft? [J/N]"
    $CloudConfirm = Read-Host
    if ($CloudConfirm -notmatch '^[JjYy]$') {
        Write-Host "Loeschen abgebrochen. Quelle bleibt erhalten: $SourcePath"
        Add-LogLine -Path $LogFile -Message "[$(Get-Date)] INFO: Nutzer hat Cloud-Pruefung nicht bestaetigt. Quelle '$SourcePath' bleibt erhalten." `
                    -Encoding ([System.Text.Encoding]::Unicode)
        Write-Host "`nVorgang abgeschlossen. Details: $LogFile" -ForegroundColor Green
        Stop-Script
    }

    # ==================================================================
    # VERIFIKATIONS-AUSWAHL (CRC / LIGHT / KEINE)
    # ==================================================================
    Write-Host ""
    Write-Host "================================================"
    Write-Host "  ZIEL-VERIFIKATION"
    Write-Host "================================================"
    Write-Host ""
    # Beide Verfahren pruefen das Ziel so, wie es LOKAL unter G:\ erscheint
    # (Drive-Cache), nicht den Stand in der Cloud. Frueher klangen die
    # Texte nach einer Cloud-Pruefung; massgeblich fuer den Upload bleibt
    # die Rueckfrage oben (Kontrolle in drive.google.com).
    Write-Host "  Hinweis: Beide Verfahren pruefen das Ziel unter G:\ (lokaler Drive-Cache)," -ForegroundColor Yellow
    Write-Host "  NICHT, ob der Upload in die Cloud abgeschlossen ist." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  [1] SHA-256 Pruefsumme (bit-genau, sehr sicher, langsamer)"
    Write-Host "      -- hashed jede Quell- und Zieldatei (Inhalt, wie er unter G:\ gelesen wird)"
    Write-Host "      -- Cloud-Platzhalter werden dabei automatisch heruntergeladen"
    Write-Host ""
    Write-Host "  [2] Light-Check (Existenz + Dateigroesse, schnell)"
    Write-Host "      -- prueft fuer jede Quelldatei, ob sie im Ziel liegt"
    Write-Host "      -- vergleicht nur die Dateigroesse, nicht den Inhalt"
    Write-Host "      -- Cloud-Platzhalter (Inhalt nicht lokal) gelten als NICHT geprueft"
    Write-Host "         und bleiben in der Quelle; ihre Groesse belegt den Inhalt nicht"
    Write-Host "      -- empfohlen, wenn der CRC-Check zu lange dauern wuerde"
    Write-Host ""
    Write-Host "  [3] Keine Verifikation (nur bei RoboExit < 4 erlaubtes Komplett-Loeschen)"
    Write-Host "      -- bei RoboExit >= 4 bleibt die Quelle komplett erhalten"
    Write-Host ""
    Write-Host "================================================"
    $verifyChoice = Read-Host "Auswahl (1, 2 oder 3)"

    $verifyMode = $null
    $csProblems = 0
    $csOkPaths  = $null
    $csCounters = $null

    switch ($verifyChoice) {
        "1" {
            $verifyMode = "CRC"
            Write-Host ""
            Write-Host "Starte Pruefsummen-Verifizierung ..."
            $csData     = Compare-FileChecksums -SourceRoot $SourcePath -DestRoot $FinalDest -ReportFile $ChecksumFile -LogFile $LogFile -SkipReparsePoints $SkipReparsePoints
            $csCounters = $csData.Counters
            $csOkPaths  = $csData.OkPaths
            # Platzhalter zaehlen als Problem: nur so bleiben sie beim
            # partiellen Loeschen in der Quelle erhalten.
            $csPlaceholder = if ($null -ne $csCounters.Placeholder) { $csCounters.Placeholder } else { 0 }
            $csProblems = $csCounters.Mismatch + $csCounters.Missing + $csCounters.Error + $csPlaceholder
        }
        "2" {
            $verifyMode = "LIGHT"
            Write-Host ""
            Write-Host "Starte Light-Check (Existenz + Groesse) ..."
            $csData     = Compare-FileExistence -SourceRoot $SourcePath -DestRoot $FinalDest -ReportFile $LightCheckFile -LogFile $LogFile -SkipReparsePoints $SkipReparsePoints
            $csCounters = $csData.Counters
            $csOkPaths  = $csData.OkPaths
            # Platzhalter zaehlen als Problem: nur so bleiben sie beim
            # partiellen Loeschen in der Quelle erhalten.
            $csPlaceholder = if ($null -ne $csCounters.Placeholder) { $csCounters.Placeholder } else { 0 }
            $csProblems = $csCounters.Mismatch + $csCounters.Missing + $csCounters.Error + $csPlaceholder
        }
        "3" {
            $verifyMode = "NONE"
            Write-Host ""
            if ($RoboExit -ge 4) {
                Write-Host "HINWEIS: Robocopy hat zuvor Fehler gemeldet (Exit-Code $RoboExit)." -ForegroundColor Yellow
                Write-Host "  Ohne Verifikation wird das Loeschen der Quelle deshalb gleich" -ForegroundColor Yellow
                Write-Host "  hart abgebrochen - die Quelle bleibt zum Schutz vollstaendig erhalten." -ForegroundColor Yellow
                Write-Host "  Fuer selektives Loeschen der erfolgreich kopierten Dateien bitte" -ForegroundColor Yellow
                Write-Host "  stattdessen Option [2] (Light-Check) waehlen." -ForegroundColor Yellow
                Write-Host ""
            }
            Write-Host "Verifikation uebersprungen."
            Add-LogLine -Path $LogFile -Message "[$(Get-Date)] INFO: Verifikation vom Nutzer uebersprungen." `
                        -Encoding ([System.Text.Encoding]::Unicode)
        }
        default {
            Write-Warning "Ungueltige Auswahl. Bitte 1, 2 oder 3 eingeben."
            Stop-Script
        }
    }

    # Wurde die Verifikation per Strg+C abgebrochen, ist die Liste der
    # geprueften Dateien unvollstaendig - $csProblems bleibt dabei 0 und
    # der Ablauf saehe wie ein sauberer Durchlauf aus. Deshalb hier
    # ausdruecklich abbrechen, statt auf Basis halber Daten zu loeschen.
    if ($verifyMode -in @("CRC", "LIGHT") -and (Test-ShouldStop)) {
        Write-Host ""
        Write-Warning "Die Verifikation wurde abgebrochen und ist unvollstaendig."
        Write-Warning "Geprueft wurden $($csCounters.Total) Datei(en). Es wird NICHTS geloescht."
        Add-LogLine -Path $LogFile -Message "[$(Get-Date)] ABBRUCH (Strg+C): Verifikation unvollstaendig ($($csCounters.Total) geprueft). Quelle unveraendert." `
                    -Encoding ([System.Text.Encoding]::Unicode)
        Write-Host "`nVorgang abgeschlossen. Details: $LogFile" -ForegroundColor Green
        Stop-Script
    }

    # --- Bei Abweichungen: Partielles Loeschen bestaetigen lassen ---
    if ($verifyMode -in @("CRC", "LIGHT") -and $csProblems -gt 0) {
        Write-Host ""
        $modeLabel = if ($verifyMode -eq "CRC") { "PRUEFSUMMEN" } else { "LIGHT-CHECK" }
        Write-Warning "$modeLabel MELDET ABWEICHUNGEN."
        Write-Warning "$($csCounters.Mismatch) Abweichung(en), $($csCounters.Missing) fehlend, $($csCounters.Error) Fehler."
        $reportPath = if ($verifyMode -eq "CRC") { $ChecksumFile } else { $LightCheckFile }
        Write-Warning "Details: $reportPath"
        Write-Host ""

        Write-Host "Moechten Sie das partielle Loeschen aktivieren?"
        Write-Host "  -> Nur die ERFOLGREICH verifizierten Dateien werden im Original geloescht."
        Write-Host "  -> Dateien mit Abweichung, Fehlern oder zu langen Pfaden bleiben erhalten."
        $csErrorContinue = Read-Host "[J] Partiell loeschen / [N] Nichts loeschen und abbrechen"

        if ($csErrorContinue -notmatch '^[JjYy]$') {
            Write-Host "Loeschen vollstaendig abgebrochen. Quelle bleibt erhalten: $SourcePath"
            Add-LogLine -Path $LogFile -Message "[$(Get-Date)] ABBRUCH: Nutzer hat partielles Loeschen abgelehnt." -Encoding ([System.Text.Encoding]::Unicode)
            Write-Host "`nVorgang abgeschlossen. Details: $LogFile" -ForegroundColor Green
            Stop-Script
        }
        Add-LogLine -Path $LogFile -Message "[$(Get-Date)] WARNUNG: Partielles Loeschen aktiviert ($verifyMode). $csProblems Dateien mit Problemen bleiben erhalten." -Encoding ([System.Text.Encoding]::Unicode)
    }

    # --- Zeitstempel wiederherstellen ---
    if (Test-Path -LiteralPath $TimestampFile) {
        Write-Host ""
        Write-Host "Stelle Zeitstempel auf Zieldateien wieder her ..."
        Import-Timestamps -DestPath $FinalDest -TimestampFile $TimestampFile -RestoreAccessTime $RestoreAccessTime
    }

    # ==================================================================
    # QUELLE LOESCHEN
    # ==================================================================
    # Strg+C wurde bisher nur INNERHALB der Schleifen gepollt. Wer waehrend
    # der Verifikation abbrach, landete danach trotzdem in der Loeschphase -
    # auf Basis einer unvollstaendigen Pruefliste. Hier ist Endstation.
    if (Test-ShouldStop) {
        Write-Host ""
        Write-Warning "Abbruch durch Anwender - es wird NICHTS geloescht."
        Write-Warning "Die Quelle bleibt vollstaendig erhalten: $SourcePath"
        Add-LogLine -Path $LogFile -Message "[$(Get-Date)] ABBRUCH (Strg+C): Loeschphase nicht erreicht. Quelle unveraendert." `
                    -Encoding ([System.Text.Encoding]::Unicode)
        Write-Host "`nVorgang abgeschlossen. Details: $LogFile" -ForegroundColor Green
        Stop-Script
    }

    Write-Host ""
    Write-Host "Loeschen der Quelle: $SourcePath ..."

    $DeleteErrors = New-Object System.Collections.Generic.List[string]

    if ($verifyMode -in @("CRC", "LIGHT")) {
        $delGB = [Math]::Round($csCounters.Bytes / 1GB, 2)
        Write-Host ""
        Write-Host "LETZTE BESTAETIGUNG:" -ForegroundColor Yellow
        Write-Host "  Es werden jetzt $($csOkPaths.Count) verifizierte Datei(en) ($delGB GB) aus" -ForegroundColor Yellow
        Write-Host "  '$SourcePath' endgueltig geloescht." -ForegroundColor Yellow
        $finalDeleteConfirm = Read-Host "Fortfahren? [J/N]"
        if ($finalDeleteConfirm -notmatch '^[JjYy]$') {
            Write-Host "Loeschen abgebrochen. Quelle bleibt erhalten: $SourcePath"
            Add-LogLine -Path $LogFile -Message "[$(Get-Date)] ABBRUCH: Nutzer hat letzte Loeschbestaetigung verweigert. Quelle bleibt erhalten." `
                        -Encoding ([System.Text.Encoding]::Unicode)
            Write-Host "`nVorgang abgeschlossen. Details: $LogFile" -ForegroundColor Green
            Stop-Script
        }

        $modeLabel = if ($verifyMode -eq "CRC") { "CRC-verifiziert" } else { "Light-verifiziert" }
        Write-Host "  Fuehre selektives Loeschen durch (nur $modeLabel und OK-Status)..."
        $delOk = 0
        foreach ($rel in $csOkPaths) {
            if (Test-ShouldStop) { break }

            $srcLong = Add-LongPathPrefix (Join-Path $SourcePath $rel)
            try {
                if (Test-FileIsLocked -Path $srcLong) {
                    [void](Wait-FileAvailable -Path $srcLong -MaxWaitSeconds 30)
                }
                # [System.IO.File]::Delete statt Remove-Item: die
                # Provider-Cmdlets von PowerShell 5.1 behandeln
                # '\\?\'-Praefixe unzuverlaessig. Hier waere die Folge,
                # dass eine bereits ins Ziel kopierte Datei in der Quelle
                # liegenbleibt und der Lauf sie als Fehler meldet -
                # ausgerechnet bei den langen Pfaden, fuer die der
                # Praefix ueberhaupt eingefuehrt wurde.
                # Schreibschutz VOR dem Loeschen abraeumen. Auf einer Datei mit
                # gesetztem ReadOnly-Attribut wirft [System.IO.File]::Delete
                # eine UnauthorizedAccessException (nachgestellt). Test-
                # FileIsLocked schlaegt nicht an, weil es nur mit
                # FileAccess::Read oeffnet - schreibgeschuetzte Dateien gelten
                # dort korrekt als 'nicht gesperrt'. Invoke-WithRetry
                # wiederholte den Aufruf dann fuenfmal mit 1+2+4+8 Sekunden
                # Pause, obwohl ein Attributproblem durch Warten nicht
                # verschwindet: rund 15 s Leerlauf je Datei, danach 'Zugriff
                # verweigert', und die im Ziel nachweislich angekommene Datei
                # blieb in der Quelle liegen. Der Vollloesch-Zweig ueber
                # robocopy /MIR ist davon nicht betroffen - beide Wege
                # verhielten sich also unterschiedlich, obwohl beide 'Quelle
                # loeschen' heissen. Import-Timestamps macht es an anderer
                # Stelle bereits genauso vor.
                try {
                    $srcAttr = [System.IO.File]::GetAttributes($srcLong)
                    if ($srcAttr -band [System.IO.FileAttributes]::ReadOnly) {
                        [System.IO.File]::SetAttributes(
                            $srcLong,
                            ($srcAttr -band (-bnot [System.IO.FileAttributes]::ReadOnly)))
                    }
                } catch {
                    Write-Verbose "Schreibschutz nicht abraeumbar: $srcLong - $_"
                }
                # MaxAttempts bewusst niedrig: Nachdem der Schreibschutz oben
                # abgeraeumt ist, deutet eine verbleibende
                # UnauthorizedAccessException auf ein echtes Rechteproblem hin,
                # das durch Warten nicht besser wird. Ein Versuch Wiederholung
                # bleibt fuer den kurzlebigen Fall (AV-Scanner haelt die Datei
                # noch); die vollen fuenf Versuche kosteten je Datei rund 15 s
                # Leerlauf - bei 5.000 betroffenen Dateien ueber 20 Stunden,
                # in denen nichts geloescht wird.
                Invoke-WithRetry -ScriptBlock {
                    [System.IO.File]::Delete($srcLong)
                } -MaxAttempts 2 | Out-Null
                $delOk++
                # Gemeinsames Laufprotokoll: bewusst NUR die Loeschungen
                # und Fehler, nicht jede gepruefte Datei. Bei einem
                # Bestand in Millionenhoehe waere ein Eintrag je Pruefung
                # eine Protokolldatei von hunderten Megabyte - der
                # Loeschvorgang dagegen ist der Schritt, den man
                # spaeter tatsaechlich nachvollziehen will.
                if ($script:GemeinsamGeladen) {
                    try {
                        Write-Laufprotokoll -Skript '8_verschieben_auf_Google_Drive' `
                            -Pfad $srcLong -Aktion 'Quelle geloescht' `
                            -Status 'OK' -Detail "verifiziert per $verifyMode"
                    } catch { }
                }
            } catch {
                $DeleteErrors.Add("Dateifehler: $_")
                if ($script:GemeinsamGeladen) {
                    try {
                        Write-Laufprotokoll -Skript '8_verschieben_auf_Google_Drive' `
                            -Pfad $srcLong -Aktion 'Quelle geloescht' `
                            -Status 'FEHLER' -Detail ([string]$_)
                    } catch { }
                }
            }
        }

        Write-Host "  Raeume leere Ordnerstrukturen auf..."
        $allDirs = New-Object System.Collections.Generic.List[string]
        Get-FilesStreaming -RootPath $SourcePath -SkipReparsePoints $SkipReparsePoints -IncludeDirectories |
            ForEach-Object { $allDirs.Add($_.FullName) }
        # Nach VERZEICHNISTIEFE absteigend, nicht nach Pfadlaenge. Die
        # Laenge korreliert meist mit der Tiefe, garantiert sie aber nicht:
        # ein tiefer Ordner mit kurzen Namen kam sonst vor seinem
        # flacheren Geschwister mit langem Namen an die Reihe. Da
        # Directory.Delete(...,$false) nur leere Ordner entfernt und
        # Fehler verschluckt werden, blieben in solchen Faellen einzelne
        # leere Ordner zurueck.
        $allDirs.Sort([System.Collections.Generic.Comparer[string]]::Create({
            param($a, $b)
            $da = ($a -split '\\').Count
            $db = ($b -split '\\').Count
            if ($da -ne $db) { return $db.CompareTo($da) }
            return $b.CompareTo($a)
        }))
        foreach ($dirFull in $allDirs) {
            if (Test-ShouldStop) { break }

            $dirLong = Add-LongPathPrefix (Remove-LongPathPrefix $dirFull)
            try {
                [System.IO.Directory]::Delete($dirLong, $false)
            } catch { }
        }
        if (-not $isRootPath) {
            try {
                [System.IO.Directory]::Delete((Add-LongPathPrefix $SourcePath), $false)
            } catch { }
        }

        # Strg+C in der Loeschschleife: vorher stand trotzdem "ERFOLG" im
        # Log, obwohl ein Teil der verifizierten Dateien noch in der
        # Quelle lag.
        if (Test-ShouldStop) {
            Write-Warning "Loeschen durch Anwender abgebrochen (Strg+C): $delOk von $($csOkPaths.Count) verifizierten Datei(en) geloescht."
            Add-LogLine -Path $LogFile -Message "[$(Get-Date)] ABBRUCH (Strg+C): Loeschen der Quelle unterbrochen - $delOk von $($csOkPaths.Count) verifizierten Datei(en) geloescht, der Rest bleibt in der Quelle." `
                        -Encoding ([System.Text.Encoding]::Unicode)
            $DeleteErrors.Add("Loeschen durch Anwender abgebrochen (Strg+C) - die Quelle wurde nur teilweise geloescht.")
        } elseif ($csProblems -gt 0) {
            Write-Host "  Erfolgreich geloescht: $delOk Dateien." -ForegroundColor Green
            Write-Host "  ACHTUNG: $csProblems nicht verifizierte Datei(en) wurden absichtlich im Quellordner behalten."
            Add-LogLine -Path $LogFile -Message "[$(Get-Date)] ERFOLG (PARTIELL/$verifyMode): $delOk Dateien geloescht. $csProblems Datei(en) absichtlich im Quellordner behalten." `
                        -Encoding ([System.Text.Encoding]::Unicode)
        } else {
            Write-Host "  Erfolgreich geloescht: $delOk Dateien." -ForegroundColor Green
            Add-LogLine -Path $LogFile -Message "[$(Get-Date)] ERFOLG ($verifyMode): $delOk Dateien geloescht." `
                        -Encoding ([System.Text.Encoding]::Unicode)
        }

    } else {
        if ($RoboExit -ge 4) {
            Write-Host ""
            Write-Host "SICHERHEITSABBRUCH: Komplettes Loeschen verweigert." -ForegroundColor Red
            Write-Host "  Robocopy hat Kopierfehler oder Konflikte gemeldet (Exit-Code $RoboExit)." -ForegroundColor Red
            Write-Host "  Da keine Verifikation durchgefuehrt wurde, kann nicht sichergestellt" -ForegroundColor Red
            Write-Host "  werden, dass alle Dateien im Ziel vorhanden sind." -ForegroundColor Red
            Write-Host "  Die Quelle '$SourcePath' bleibt zum Schutz vollstaendig erhalten." -ForegroundColor Red
            Write-Host ""
            Write-Host "  --> EMPFEHLUNG: Skript erneut starten und Option [2] Light-Check waehlen." -ForegroundColor Yellow
            Write-Host "      Damit werden die erfolgreich kopierten Dateien selektiv geloescht." -ForegroundColor Yellow
            $DeleteErrors.Add("Loeschen blockiert: Robocopy-Fehler + keine Verifikation (Datenverlust-Schutz).")
            Add-LogLine -Path $LogFile -Message "[$(Get-Date)] SICHERHEITSABBRUCH: Komplettes Loeschen verweigert. RoboExit=$RoboExit. Verifikation uebersprungen." `
                        -Encoding ([System.Text.Encoding]::Unicode)
        } else {
            Write-Host ""
            Write-Host "LETZTE BESTAETIGUNG:" -ForegroundColor Yellow
            Write-Host "  Der KOMPLETTE Quellordner '$SourcePath' wird jetzt endgueltig geloescht." -ForegroundColor Yellow
            Write-Host "  (Ohne Verifikation gibt es keine dateigenaue Pruefung -- Grundlage ist" -ForegroundColor Yellow
            Write-Host "  allein der fehlerfreie Robocopy-Lauf, Exit-Code $RoboExit.)" -ForegroundColor Yellow
            $finalDeleteConfirm = Read-Host "Fortfahren? [J/N]"
            if ($finalDeleteConfirm -notmatch '^[JjYy]$') {
                Write-Host "Loeschen abgebrochen. Quelle bleibt erhalten: $SourcePath"
                Add-LogLine -Path $LogFile -Message "[$(Get-Date)] ABBRUCH: Nutzer hat letzte Loeschbestaetigung verweigert. Quelle bleibt erhalten." `
                            -Encoding ([System.Text.Encoding]::Unicode)
                Write-Host "`nVorgang abgeschlossen. Details: $LogFile" -ForegroundColor Green
                Stop-Script
            }

            # Der Leerordner fuer den /MIR-Purge lag vorher in "Dokumente".
            # Das ist gleich doppelt falsch: Dokumente wird haeufig selbst
            # in die Cloud synchronisiert (jeder Lauf loest Sync-Traffic
            # aus), und wenn der Nutzer das Profilverzeichnis als Quelle
            # waehlt, liegt der Leerordner INNERHALB der Quelle - der
            # Purge loescht dann sein eigenes Arbeitsverzeichnis.
            $emptyBaseCandidates = @($env:TEMP,
                                     (Get-UserShellFolder -Name 'Documents'),
                                     $env:LOCALAPPDATA)
            $emptyDir = $null
            foreach ($cand in $emptyBaseCandidates) {
                if ([string]::IsNullOrWhiteSpace($cand)) { continue }
                if (Test-PathInside -Child $cand -Parent $SourcePath) { continue }
                $tryDir = Join-Path $cand "8_verschieben_auf_Google_Drive_empty_$([guid]::NewGuid().Guid)"
                try {
                    New-Item -ItemType Directory -Path $tryDir -Force -ErrorAction Stop | Out-Null
                    $emptyDir = $tryDir
                    break
                } catch { continue }
            }
            if (-not $emptyDir) {
                $DeleteErrors.Add("Kein Arbeitsverzeichnis fuer den Loeschvorgang verfuegbar (alle Kandidaten liegen in der Quelle oder sind nicht beschreibbar).")
            }
            if ($DeleteErrors.Count -eq 0) {
                # robocopy /MIR beachtet /XD beim Purge NUR fuer Verzeichnisse
                # unmittelbar in der Wurzel. Liegt ein ausgeschlossenes
                # Verzeichnis tiefer (z. B. Q:\Abteilung\Unterordner\~snapshot),
                # gilt der ganze Ast 'Unterordner' als ueberzaehlig und wird
                # samt Snapshot geloescht -- auch wenn der Vollpfad zusaetzlich
                # in /XD steht (beides nachgestellt). Diese Verzeichnisse wurden
                # per /XD nie kopiert, tauchen in keiner Verifikation auf und
                # waeren damit endgueltig verloren. Darum vorher pruefen und in
                # diesem Fall ueber die skripteigene Aufzaehlung loeschen, die
                # die Ausschlussliste auf JEDER Ebene beachtet.
                $srcRoot = $SourcePath.TrimEnd('\')
                $nestedExcluded = @(
                    Get-RemainingExcludedDir -RootPath $SourcePath |
                        Where-Object { (Split-Path $_ -Parent).TrimEnd('\') -ne $srcRoot }
                )
                if ($nestedExcluded.Count -gt 0) {
                    Write-Host "  $($nestedExcluded.Count) ausgeschlossene(r) Systemordner liegt/liegen unterhalb der Wurzel." -ForegroundColor Yellow
                    Write-Host "  Der schnelle Robocopy-Purge wuerde sie mitloeschen -- es wird stattdessen" -ForegroundColor Yellow
                    Write-Host "  dateiweise geloescht. Das dauert laenger, verschont sie aber zuverlaessig." -ForegroundColor Yellow
                    Add-LogLine -Path $LogFile -Message "[$(Get-Date)] HINWEIS: Robocopy-Purge uebersprungen, $($nestedExcluded.Count) verschachtelte(r) ausgeschlossene(r) Ordner: $($nestedExcluded -join '; ')" `
                                -Encoding ([System.Text.Encoding]::Unicode)
                    $purgeErrors = 0
                    Get-FilesStreaming -RootPath $SourcePath -SkipReparsePoints $SkipReparsePoints -IncludeFiles |
                        ForEach-Object {
                            if (Test-ShouldStop) { return }
                            try {
                                [System.IO.File]::Delete((Add-LongPathPrefix (Remove-LongPathPrefix $_.FullName)))
                            } catch { $purgeErrors++ }
                        }
                    if ($purgeErrors -gt 0) {
                        $DeleteErrors.Add("$purgeErrors Datei(en) konnten beim Loeschen der Quelle nicht entfernt werden.")
                    }
                } else {
                    $mirrorArgs = @($emptyDir, $SourcePath, "/MIR", "/R:5", "/W:5", "/NP", "/NFL", "/NDL", "/NJH", "/NJS", "/XD") + $ExcludedDirNames
                    if ($SkipReparsePoints) { $mirrorArgs += "/XJ" }
                    & robocopy $mirrorArgs | Out-Null
                    # Negativ = robocopy gewaltsam beendet (Strg+C: -1073741510,
                    # gemessen) - fiel vorher durch '-ge 8' und galt als Erfolg.
                    if ($LASTEXITCODE -ge 8 -or $LASTEXITCODE -lt 0) {
                        $label = if ($isRootPath) { "Root" } else { "Unterordner" }
                        $DeleteErrors.Add("Robocopy Mirror-Fehler bei $label-Loeschen (Exit-Code $LASTEXITCODE)")
                    }
                }
                if (-not $isRootPath) {
                    # Kurz auf Filesystem-Cache-Release warten (DFS/Netzwerk haben Handle-Latenz nach robocopy /MIR).
                    Start-Sleep -Milliseconds 750

                    # KEIN 'Remove-Item -Recurse -Force': der Purge oben hat die
                    # Verzeichnisse aus $ExcludedDirNames per /XD bewusst
                    # verschont. Ein rekursives Force-Delete wuerde genau diesen
                    # Bestand vernichten -- Daten, die nie ins Ziel kopiert
                    # wurden (Zeile mit /XD beim Kopieren) und die deshalb in
                    # keiner Verifikation auftauchen koennen. Stattdessen wie im
                    # selektiven Zweig tiefenzuerst nur LEERE Verzeichnisse
                    # entfernen; was uebrig bleibt, ist absichtlich uebrig.
                    $restDirs = New-Object System.Collections.Generic.List[string]
                    try {
                        Get-FilesStreaming -RootPath $SourcePath -SkipReparsePoints $SkipReparsePoints -IncludeDirectories |
                            ForEach-Object { $restDirs.Add($_.FullName) }
                    } catch { $DeleteErrors.Add("Aufraeumen der Quellstruktur: $_") }
                    $restDirs.Sort([System.Collections.Generic.Comparer[string]]::Create({
                        param($a, $b)
                        $da = ($a -split '\\').Count
                        $db = ($b -split '\\').Count
                        if ($da -ne $db) { return $db.CompareTo($da) }
                        return $b.CompareTo($a)
                    }))
                    foreach ($dirFull in $restDirs) {
                        if (Test-ShouldStop) { break }
                        try {
                            [System.IO.Directory]::Delete((Add-LongPathPrefix (Remove-LongPathPrefix $dirFull)), $false)
                        } catch { }
                    }

                    try {
                        [System.IO.Directory]::Delete((Add-LongPathPrefix $SourcePath), $false)
                    } catch {
                        $keptDirs = Get-RemainingExcludedDir -RootPath $SourcePath
                        if ($keptDirs.Count -gt 0) {
                            Write-Host "  $($keptDirs.Count) Systemordner absichtlich behalten (nie kopiert):" -ForegroundColor Yellow
                            foreach ($k in $keptDirs) { Write-Host "    $k" -ForegroundColor Yellow }
                            Write-Host "  Der Quellordner '$SourcePath' bleibt deshalb bestehen." -ForegroundColor Yellow
                            Add-LogLine -Path $LogFile -Message "[$(Get-Date)] HINWEIS: $($keptDirs.Count) ausgeschlossene(r) Systemordner absichtlich behalten: $($keptDirs -join '; '). Quellordner '$SourcePath' bleibt bestehen." `
                                        -Encoding ([System.Text.Encoding]::Unicode)
                        } else {
                            $DeleteErrors.Add("Quellordner konnte nicht entfernt werden: $_")
                        }
                    }
                }
            }
            if ($emptyDir) {
                Remove-Item -LiteralPath $emptyDir -Recurse -Force -ErrorAction SilentlyContinue
            }

            # Nachkontrolle, bevor unten "ERFOLG: Quelle ... geloescht" steht.
            # Vorher genuegte dafuer eine leere Fehlerliste: ein per Strg+C
            # abgebrochener Purge (robocopy negativ, die Schleifen oben per
            # break/return verlassen) hinterliess keinen Eintrag, und bei
            # einer Laufwerks-/Freigabewurzel wurde gar nicht nachgesehen.
            # Erfolg heisst jetzt: kein Abbruch und keine Datei mehr in der
            # Quelle ausser in den absichtlich behaltenen Ausschlussordnern
            # (die Get-FilesStreaming ueberspringt).
            if (Test-ShouldStop) {
                $DeleteErrors.Add("Abbruch durch Anwender (Strg+C) - die Quelle wurde nur teilweise geloescht.")
            } elseif ($DeleteErrors.Count -eq 0 -and
                      [System.IO.Directory]::Exists((Add-LongPathPrefix $SourcePath))) {
                $restDateien = 0
                Get-FilesStreaming -RootPath $SourcePath -SkipReparsePoints $SkipReparsePoints -IncludeFiles |
                    ForEach-Object { $restDateien++ }
                if (Test-ShouldStop) {
                    $DeleteErrors.Add("Abbruch durch Anwender (Strg+C) - Restbestand der Quelle nicht vollstaendig geprueft.")
                } elseif ($restDateien -gt 0) {
                    $DeleteErrors.Add("$restDateien Datei(en) liegen nach dem Loeschen weiterhin in der Quelle (ausserhalb der absichtlich behaltenen Ausschlussordner).")
                }
            }
        }
    }

    if ($DeleteErrors.Count -gt 0) {
        Write-Warning "Beim Loeschen sind Fehler aufgetreten:"
        $DeleteErrors | Select-Object -Unique | ForEach-Object { Write-Warning "  $_" }
        Add-LogLine -Path $LogFile -Message "[$(Get-Date)] WARNUNG: Fehler beim Loeschen der Quelle." -Encoding ([System.Text.Encoding]::Unicode)
    } elseif ($null -eq $csOkPaths) {
        Write-Host "Quelle erfolgreich vollstaendig geloescht." -ForegroundColor Green
        Add-LogLine -Path $LogFile -Message "[$(Get-Date)] ERFOLG: Quelle '$SourcePath' geloescht." -Encoding ([System.Text.Encoding]::Unicode)
    }

# ==================================================================
# METHODE 2: RCLONE
# ==================================================================
} else {

    $CloudDest      = $FinalDest -replace '\\', '/'
    $cloudDestSlash = $CloudDest.TrimEnd('/') + '/'

    $RcloneArgs = @(
        "move",
        $SourcePath,
        $CloudDest,
        "--create-empty-src-dirs",
        "--bwlimit", "08:00,15M 18:00,off",
        "--transfers=$($RcloneSettings.Transfers)",
        "--tpslimit=8",
        "--drive-chunk-size=$($RcloneSettings.ChunkSize)",
        "--drive-use-trash=false",
        "--drive-stop-on-upload-limit",
        "--retries=5",
        "--low-level-retries=15",
        "--metadata",
        "--verbose",
        "--stats=1m"
    )

    # Dieselben Verzeichnisse ausschliessen wie die Robocopy-Methode
    # (Papierkorb, Systemordner, NAS-Snapshots).
    $RcloneExcludeArgs = @()
    foreach ($exDir in $ExcludedDirNames) {
        # Zwei Muster sind noetig: '<dir>/**' haelt den Inhalt heraus,
        # '<dir>/' den Ordner selbst. Ohne das zweite Muster legt
        # --create-empty-src-dirs im Ziel leere Papierkorb- und
        # Snapshot-Ordner an.
        $RcloneExcludeArgs += @("--exclude", "$exDir/**")
        $RcloneExcludeArgs += @("--exclude", "$exDir/")
    }
    $RcloneArgs += $RcloneExcludeArgs

    if ($SkipReparsePoints) { $RcloneArgs += "--skip-links" }
    if (-not $isRootPath)   { $RcloneArgs += "--delete-empty-src-dirs" }
    if ($DryRun)            { $RcloneArgs += "--dry-run" }

    # 'rclone move' loescht jede Quelldatei unmittelbar nach dem
    # erfolgreichen Upload. Die Robocopy-Methode fragt dreimal nach, die
    # rclone-Methode fragte bisher gar nicht - dieselbe Tragweite, aber
    # ohne Rueckfrage. rclone prueft nach dem Upload die Pruefsumme, das
    # Loeschen ist also abgesichert; die Entscheidung bleibt trotzdem beim
    # Anwender.
    if (-not $DryRun) {
        Write-Host ""
        Write-Host "LETZTE BESTAETIGUNG (rclone move):" -ForegroundColor Yellow
        Write-Host "  rclone laedt jede Datei aus '$SourcePath' hoch und loescht sie" -ForegroundColor Yellow
        Write-Host "  danach sofort in der Quelle (nach Pruefsummen-Kontrolle)." -ForegroundColor Yellow
        Write-Host "  Ziel: $CloudDest" -ForegroundColor Yellow
        Write-Host "  Ein Zwischenschritt zum Zuruecknehmen existiert nicht." -ForegroundColor Yellow
        $rcloneConfirm = Read-Host "Fortfahren? [J/N]"
        if ($rcloneConfirm -notmatch '^[JjYy]$') {
            Write-Host "Vorgang abgebrochen. Es wurde nichts uebertragen oder geloescht."
            Add-LogLine -Path $LogFile -Message "[$(Get-Date)] ABBRUCH: Nutzer hat den rclone-Transfer vor dem Start abgelehnt."
            Stop-Script
        }
        Write-Host ""
    }

    # Strg+C waehrend der Rueckfragen: vor dem Transfer aussteigen.
    Stop-WennAbgebrochen -Phase 'Vorbereitung vor dem rclone-Transfer'

    Write-Host "Starte rclone-Transfer (direkt in die Cloud) ..."
    Write-Host ""
    $colWidth     = try { [Math]::Max(40, [Console]::WindowWidth - 6) } catch { 80 }
    $rcloneWriter = [System.IO.StreamWriter]::new($LogFile, $true, $Utf8Bom)
    try {
        & $RcloneExe $RcloneArgs 2>&1 | ForEach-Object {
            $line = "$_"
            $rcloneWriter.WriteLine($line)
            if ($line -match ':\s+(?:Copied|Moved)') {
                $file = ($line -replace '^.*?\d{2}:\d{2}:\d{2}\s+\S+\s+:\s+', '') `
                             -replace ':\s+(?:Copied|Moved).*$', ''
                $disp = if ($file.Length -gt $colWidth) { '...' + $file.Substring($file.Length - $colWidth + 3) } else { $file }
                Write-Host ("`r  " + $disp.PadRight($colWidth)) -NoNewline
            } elseif ($line -match 'ERROR\s*:') {
                Write-Host ("`r" + (' ' * ($colWidth + 4)) + "`r$line") -ForegroundColor Red
            }
        }
    } finally {
        $rcloneWriter.Close()
    }
    Write-Host ""
    $RcloneExit = $LASTEXITCODE

    if ($RcloneExit -eq 0 -and $isRootPath -and -not $DryRun) {
        # Dieselben Ausschluesse wie beim move. Ohne sie lief rmdirs auch
        # durch ~snapshot/.snapshot, 'System Volume Information' und
        # $RECYCLE.BIN (read-only bzw. gesperrt) - Fehler dort setzten den
        # Exit-Code und damit "rclone meldet Fehler", obwohl diese Ordner
        # absichtlich nie migriert werden. Die Filter sind globale Optionen;
        # laut rclone-Doku (v1.73.1, README.txt) gelten sie fuer alle
        # Befehle, nur "Rclone purge does not obey filters"; Changelog
        # v1.55: "rmdirs: Make --rmdirs obey the filters". Beim lokalen
        # Backend (kein ListR) steigt rclone in per '<dir>/' ausgeschlossene
        # Ordner nicht hinab ("Directory recursion optimisation").
        $RmdirArgs = @("rmdirs", $SourcePath, "--leave-root", "--log-file", $LongLogFile, "--log-level", "WARNING") + $RcloneExcludeArgs
        & $RcloneExe $RmdirArgs
        if ($LASTEXITCODE -ne 0) { $RcloneExit = $LASTEXITCODE }
    }

    # --- Log-Pfad-Substitution ---
    Start-Sleep -Seconds 1
    if (Test-Path -LiteralPath $LogFile) {
        $cloudDestSlashSafe = $cloudDestSlash.Replace('$', '$$')
        $tmpFile    = "$LogFile.tmp"
        $reader     = $null
        $writer     = $null
        $logSuccess = $false
        try {
            $reader  = [System.IO.StreamReader]::new($LogFile, $Utf8Bom)
            $writer  = [System.IO.StreamWriter]::new($tmpFile, $false, $Utf8Bom)
            while ($null -ne ($line = $reader.ReadLine())) {
                if ($line -match ':\s+(Transferred|Errors|Checks|Deleted|Renamed|Elapsed time|Server side copies)\s*:') {
                    $writer.WriteLine($line)
                } elseif ($line -match '^(.*?(?:ERROR|WARNING|NOTICE|INFO)\s*:\s+)(.+)(:\s+(?:Copied|Moved|Deleted|Renamed|Updated|Unchanged|Skipped|Removing|Multipart|Failed)\b)') {
                    $writer.WriteLine(
                        $line -replace '^(.*?(?:ERROR|WARNING|NOTICE|INFO)\s*:\s+)(.+)(:\s+(?:Copied|Moved|Deleted|Renamed|Updated|Unchanged|Skipped|Removing|Multipart|Failed)\b)', "`${1}$cloudDestSlashSafe`${2}`${3}"
                    )
                } elseif ($line -match '^(.*?(?:ERROR|WARNING|NOTICE|INFO)\s*:\s+)(.+?):\s') {
                    $writer.WriteLine(
                        $line -replace '^(.*?(?:ERROR|WARNING|NOTICE|INFO)\s*:\s+)(.+?:\s)', "`${1}$cloudDestSlashSafe`${2}"
                    )
                } else {
                    # Generische Status-Message (z.B. "INFO : There was nothing to transfer") - unveraendert lassen.
                    $writer.WriteLine($line)
                }
            }
            $logSuccess = $true
        } catch {
            Write-Warning "Log konnte nicht formatiert werden: $_"
        } finally {
            if ($reader) { $reader.Close() }
            if ($writer) { $writer.Close() }
        }
        if ($logSuccess) {
            try {
                Move-Item -LiteralPath $tmpFile -Destination $LogFile -Force -ErrorAction Stop
            } catch {
                Write-Warning "Fehler beim Ersetzen der Log-Datei: $_"
            }
        }
        # Zwischendatei bei Fehlschlag entfernen. Sie blieb sonst als
        # '<log>.log.tmp' neben dem Protokoll liegen und wurde von keinem
        # Aufraeumpfad erfasst - bei jedem gescheiterten Lauf eine weitere.
        if (Test-Path -LiteralPath $tmpFile) {
            Remove-Item -LiteralPath $tmpFile -Force -ErrorAction SilentlyContinue
        }
    }

    if ($DryRun) {
        Write-Host ""
        Write-Host "PROBELAUF abgeschlossen (rclone --dry-run, Exit-Code $RcloneExit)." -ForegroundColor Green
        Write-Host "Es wurde nichts uebertragen oder geloescht. Details: $LogFile"
        Add-LogLine -Path $LogFile -Message "[$(Get-Date)] PROBELAUF beendet (rclone --dry-run, Exit-Code $RcloneExit). Keine Aenderungen."
    } elseif ($RcloneExit -eq 0) {
        Write-Host ""
        Write-Host "Migration erfolgreich. Quelle wurde von rclone sicher geloescht." -ForegroundColor Green
        Add-LogLine -Path $LogFile -Message "[$(Get-Date)] ERFOLG: rclone-Transfer abgeschlossen. Quelle '$SourcePath', Ziel '$FinalDest'."
    } else {
        Write-Host ""
        Write-Warning "rclone meldet Fehler (Exit-Code $RcloneExit). Bitte Log pruefen: $LogFile"
        Add-LogLine -Path $LogFile -Message "[$(Get-Date)] FEHLER: rclone Exit-Code $RcloneExit. Quelle '$SourcePath'. Log pruefen."
    }
}

# ==================================================================
# ABSCHLUSS
# ==================================================================
Write-Host ""
Write-Host "Vorgang abgeschlossen. Details: $LogFile" -ForegroundColor Green
Stop-Script
