<#
.SYNOPSIS
    Gemeinsame Grundfunktionen der Migrations-Skripte (PowerShell-Teil).

.DESCRIPTION
    Die Sammlung ist durch Kopieren gewachsen. Dieselbe Aufgabe wurde
    dadurch mehrfach geloest - und die Korrekturen wanderten nicht mit.
    Konkrete Faelle aus der Durchsicht:

      * Resolve-ScriptDirectory (ps2exe-tauglicher Skriptordner) gab es in
        2a, 6, 7, 8 und 9, aber nicht in 2b und 2c - obwohl deren Koepfe
        die ps2exe-Uebersetzung ausdruecklich dokumentieren.
      * Show-OfficeRunningWarning fehlte ausgerechnet in 2b (Excel), dem
        Skript mit dem groessten Risiko fuer uebernommene Sitzungen.
      * Dieselbe Wartefunktion hiess an fuenf Stellen anders:
        Wait-FileAvailable, Wait-FileReady, Wait-FileStable,
        Test-FileReady, Wait-ForFileUnlock.

    Dieses Modul haelt die kanonische Fassung. Die Skripte binden es MIT
    RUECKFALL ein - fehlt die Datei, arbeiten sie mit ihrer eingebauten
    Kopie weiter. Damit bleibt jedes Skript einzeln lauffaehig und die
    ps2exe-Uebersetzung funktioniert unveraendert.

.EXAMPLE
    # Einbinden mit Rueckfall (empfohlenes Muster):
    $modul = Join-Path $PSScriptRoot '_gemeinsam.psm1'
    if (Test-Path -LiteralPath $modul) {
        try { Import-Module $modul -Force -ErrorAction Stop } catch {}
    }

.NOTES
    Stand: 14.08.2026
#>

Set-StrictMode -Version 2.0

$script:ModulVersion = '1.0.0'

# ==================================================================
# Skript- und Log-Verzeichnisse
# ==================================================================
function Resolve-ScriptDirectory {
    <#
        Verzeichnis des Skripts bzw. der EXE.

        $PSScriptRoot greift beim direkten Aufruf; bei einer ps2exe-EXE
        ist es leer - dann liefert die Entry-Assembly das EXE-Verzeichnis.
        $PWD nur als letzter Rueckfall: das ist das aktuelle
        Arbeitsverzeichnis und bei Doppelklick aus dem Explorer zufaellig.
    #>
    param([string]$AufruferPfad)

    if (-not [string]::IsNullOrWhiteSpace($AufruferPfad)) {
        return Split-Path -Parent $AufruferPfad
    }
    try {
        $entry = [System.Reflection.Assembly]::GetEntryAssembly()
        if ($entry -and $entry.Location) {
            return Split-Path -Parent $entry.Location
        }
    } catch {}
    return $PWD.Path
}

function Resolve-WritableDirectory {
    <#
        Erster Kandidat, in den sich WIRKLICH schreiben laesst.

        Der Test legt eine Datei an und loescht sie wieder. Ein
        existierendes Verzeichnis sagt nichts ueber das Schreibrecht aus -
        und genau daran scheiterten Laeufe erst ganz am Ende, nachdem
        stundenlang gearbeitet worden war.
    #>
    param([string[]]$Kandidaten)

    foreach ($kandidat in $Kandidaten) {
        if ([string]::IsNullOrWhiteSpace($kandidat)) { continue }
        try {
            if (-not [System.IO.Directory]::Exists($kandidat)) {
                New-Item -ItemType Directory -Path $kandidat -Force -ErrorAction Stop | Out-Null
            }
            $probe = Join-Path $kandidat (".writetest_{0}.tmp" -f ([Guid]::NewGuid().ToString('N')))
            [System.IO.File]::WriteAllText($probe, 'x')
            [System.IO.File]::Delete($probe)
            return $kandidat
        } catch { continue }
    }
    return $null
}

# ==================================================================
# Langpfade
# ==================================================================
function Add-LongPathPrefix {
    param([string]$Path)
    if ([string]::IsNullOrEmpty($Path))  { return $Path }
    if ($Path -like '\\?\*')             { return $Path }
    if ($Path -match '^[A-Za-z]:$')      { return '\\?\' + $Path + '\' }
    if ($Path -like '\\*')               { return '\\?\UNC\' + $Path.TrimStart('\') }
    if ($Path -match '^[A-Za-z]:\\')     { return '\\?\' + $Path }
    return $Path
}

function Remove-LongPathPrefix {
    param([string]$Path)
    if ([string]::IsNullOrEmpty($Path))  { return $Path }
    if ($Path -like '\\?\UNC\*')         { return '\\' + $Path.Substring(8) }
    if ($Path -like '\\?\*')             { return $Path.Substring(4) }
    return $Path
}

function Test-PathExistsLong {
    <#
        Existenzpruefung ueber [System.IO] statt Test-Path.

        Die Provider-Cmdlets von PowerShell 5.1 kommen mit '\\?\'-
        Praefixen nicht zuverlaessig zurecht. Ein falsches $false ist
        hier besonders folgenreich: der Aufrufer haelt ein Ziel fuer frei
        und ueberschreibt eine bestehende Datei.
    #>
    param([string]$Path)
    $lang = Add-LongPathPrefix $Path
    return ([System.IO.File]::Exists($lang) -or [System.IO.Directory]::Exists($lang))
}

# ==================================================================
# Laufende Office-Sitzungen
# ==================================================================
function Get-RunningOfficeSessions {
    <#
        Office-Prozesse mit sichtbarem Hauptfenster, also echte Sitzungen
        des Anwenders - im Unterschied zu unsichtbaren
        Automatisierungs-Instanzen.
    #>
    param([string[]]$ProcessNames = @('WINWORD', 'EXCEL', 'POWERPNT'))
    $gefunden = @()
    foreach ($name in $ProcessNames) {
        try {
            $gefunden += @(Get-Process -Name $name -ErrorAction SilentlyContinue |
                           Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero })
        } catch { }
    }
    return ,@($gefunden)
}

function Show-OfficeRunningWarning {
    <#
        Weist vor dem ersten COM-Zugriff auf laufende Sitzungen hin.

        'New-Object -ComObject Word.Application' startet KEINE neue
        Instanz, wenn Office bereits laeuft, sondern haengt sich an die
        vorhandene Sitzung. Visible=$false laesst sie fuer den Anwender
        wie abgestuerzt aussehen, und die eigene Instanz ist danach nicht
        mehr zuverlaessig von der fremden zu unterscheiden - womit auch
        die PID-Bindung beim Aufraeumen ins Leere greift.

        Rueckgabe: $true = fortfahren, $false = Anwender bricht ab.
    #>
    param(
        [string[]]$ProcessNames = @('WINWORD', 'EXCEL', 'POWERPNT'),
        [string]$AnwendungsName = 'Office',
        [switch]$Silent
    )

    $sitzungen = Get-RunningOfficeSessions -ProcessNames $ProcessNames
    if ($sitzungen.Count -eq 0) { return $true }

    $namen = ($sitzungen | ForEach-Object { $_.ProcessName } | Select-Object -Unique) -join ', '
    $pids  = ($sitzungen | ForEach-Object { $_.Id }) -join ', '

    if ($Silent) {
        Write-Warning "Laufende Office-Sitzungen erkannt ($namen, PID: $pids) - sie werden geschuetzt, aber nicht geschlossen."
        return $true
    }

    Write-Host ""
    Write-Host ("=" * 66) -ForegroundColor Yellow
    Write-Host "  WARNUNG: $AnwendungsName laeuft bereits" -ForegroundColor Yellow
    Write-Host ("=" * 66) -ForegroundColor Yellow
    Write-Host "  Gefundene Sitzungen: $namen (PID: $pids)"
    Write-Host ""
    Write-Host "  Ihre Sitzung wird vom Skript NICHT beendet. Waehrend des Laufs"
    Write-Host "  kann sie aber ausgeblendet werden und Warnhinweise sind"
    Write-Host "  abgeschaltet - das wirkt wie ein Absturz."
    Write-Host ""
    Write-Host "  EMPFEHLUNG: $AnwendungsName jetzt schliessen und neu starten." -ForegroundColor Yellow
    Write-Host ("=" * 66) -ForegroundColor Yellow
    Write-Host ""
    # Ohne echte Konsole NICHT fragen. Read-Host blockiert dort unbegrenzt
    # (Aufgabenplanung mit angehaengter Konsole) oder liefert sofort leer -
    # beides ist als Freigabe ungeeignet. Massgeblich ist die tatsaechliche
    # Eingabefaehigkeit, nicht der Hostname (ps2exe meldet 'PSRunspace-Host').
    $kannFragen = $false
    try { $kannFragen = -not [Console]::IsInputRedirected } catch { $kannFragen = $false }
    if (-not $kannFragen) {
        Write-Warning "Keine interaktive Konsole - Lauf wird zum Schutz laufender Sitzungen abgebrochen. Fuer den unbeaufsichtigten Betrieb -Silent bzw. -NoInteractive verwenden."
        return $false
    }

    $antwort = Read-Host "Trotzdem fortfahren? [j/N]"
    if ($antwort -notmatch '^[JjYy]') {
        Write-Host "Abgebrochen." -ForegroundColor Cyan
        return $false
    }
    return $true
}

function Stop-TrackedOfficeProcess {
    <#
        Beendet einen Office-Prozess NUR, wenn er nachweislich der eigene
        ist. Zwei Schutzstufen gegen PID-Recycling: Windows kann eine
        freigewordene PID binnen Millisekunden neu vergeben.
          1. Namenspruefung - ist es ueberhaupt noch ein Office-Programm?
          2. Startzeitpruefung - ist es DERSELBE Prozess?
    #>
    param(
        [Nullable[int]]$ProcessId,
        [Nullable[datetime]]$StartTime = $null
    )
    if ($null -eq $ProcessId -or $ProcessId -le 0) { return }
    try {
        $proc = Get-Process -Id $ProcessId -ErrorAction Stop
        if ($proc.Name -notmatch '^(WINWORD|EXCEL|POWERPNT)$') { return }
        if ($null -ne $StartTime) {
            try { if ($proc.StartTime -ne $StartTime) { return } } catch { return }
        }
        if (-not $proc.HasExited) {
            $proc | Stop-Process -Force -ErrorAction SilentlyContinue
        }
    } catch {}
}

# ==================================================================
# Dateizugriff
# ==================================================================
function Wait-FileAvailable {
    <#
        Wartet auf exklusiven Zugriff.

        EIN Name fuer diese Aufgabe. Bisher hiess sie je nach Skript
        Wait-FileAvailable (2a), Wait-FileReady (2b), Wait-FileStable
        (2c), Test-FileReady (7) oder Wait-ForFileUnlock (8) - was
        verhinderte, dass eine Korrektur an einer Stelle sich als
        Korrektur an den anderen wiedererkennen liess.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$MaxAttempts = 10,
        [int]$DelayMs = 300
    )
    $lang = Add-LongPathPrefix $Path
    for ($i = 1; $i -le $MaxAttempts; $i++) {
        try {
            $fs = [System.IO.File]::Open($lang, 'Open', 'Read', 'None')
            $fs.Close(); $fs.Dispose()
            return $true
        } catch {
            Start-Sleep -Milliseconds $DelayMs
        }
    }
    return $false
}

function Invoke-WithRetry {
    <#
        Wiederholt eine Operation mit wachsender Wartezeit. Gegen
        Virenscanner und Netzwerk-Aussetzer, die frisch geschriebene
        Dateien kurzzeitig exklusiv halten.
    #>
    param(
        [Parameter(Mandatory)][scriptblock]$Action,
        [int]$MaxAttempts = 4,
        [int]$DelayMs = 500
    )
    for ($i = 1; $i -le $MaxAttempts; $i++) {
        try {
            return (& $Action)
        } catch {
            if ($i -eq $MaxAttempts) { throw }
            Start-Sleep -Milliseconds ($DelayMs * $i)
        }
    }
}

function Get-FileTimestamps {
    param([string]$Path)
    $lang = Add-LongPathPrefix $Path
    try {
        return @{
            Creation   = [System.IO.File]::GetCreationTimeUtc($lang)
            LastWrite  = [System.IO.File]::GetLastWriteTimeUtc($lang)
            LastAccess = [System.IO.File]::GetLastAccessTimeUtc($lang)
        }
    } catch { return $null }
}

function Restore-FileTimestamps {
    <#
        Schreibt die gesicherten Zeitstempel zurueck. Ohne das traegt
        jede bearbeitete Datei das heutige Aenderungsdatum - was bei
        einem gewachsenen Bestand dessen ganze Chronologie zerstoert.
    #>
    param([string]$Path, $Timestamps)
    if ($null -eq $Timestamps) { return $false }
    $lang = Add-LongPathPrefix $Path
    try {
        [System.IO.File]::SetCreationTimeUtc($lang,   $Timestamps.Creation)
        [System.IO.File]::SetLastWriteTimeUtc($lang,  $Timestamps.LastWrite)
        [System.IO.File]::SetLastAccessTimeUtc($lang, $Timestamps.LastAccess)
        return $true
    } catch { return $false }
}

# ==================================================================
# CSV
# ==================================================================
function ConvertTo-CsvField {
    <#
        RFC-4180-konformes Quoting. Semikolon und Anfuehrungszeichen sind
        in Windows-Dateinamen zulaessig; ohne Maskierung verschieben sie
        in Excel die Spalten und zerstoeren den Rueckweg ueber einen
        Anwenden-Modus.
    #>
    param([string]$Text, [char]$Delimiter = ';')
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $wert = $Text -replace '\r?\n', ' '
    if ($wert.IndexOf($Delimiter) -ge 0 -or $wert.Contains('"')) {
        return '"' + ($wert -replace '"', '""') + '"'
    }
    return $wert
}

# ==================================================================
# Zentrale Konfiguration
# ==================================================================
$script:PfadeCache = $null

function Import-Konfiguration {
    <#
        Liest pfade.json neben dem Modul. Eine fehlende oder fehlerhafte
        Datei darf keinen Lauf verhindern - dann gelten die Vorgabewerte.
    #>
    param([switch]$NeuLaden)

    if ($null -ne $script:PfadeCache -and -not $NeuLaden) {
        return $script:PfadeCache
    }

    $standard = [ordered]@{
        presets = @('Q:\', 'R:\', 'G:\Geteilte Ablagen', 'G:\Meine Ablage', '\\server\dfs')
    }

    try {
        $datei = Join-Path $PSScriptRoot 'pfade.json'
        if (Test-Path -LiteralPath $datei) {
            $roh = Get-Content -LiteralPath $datei -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($roh.PSObject.Properties.Name -contains 'presets' -and $roh.presets) {
                $standard['presets'] = @($roh.presets)
            }
            $script:PfadeCache = $standard
            return $standard
        }
    } catch { }

    $script:PfadeCache = $standard
    return $standard
}

function Import-Presetliste {
    <#
        Die Verzeichnis-Presets als Liste, leere Eintraege entfernt.
    #>
    $konf = Import-Konfiguration
    return @($konf['presets'] | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

# ==================================================================
# Gemeinsames Laufprotokoll (JSON Lines)
# ==================================================================
function Write-Laufprotokoll {
    <#
        Eine Zeile je Datei und Schritt, ueber alle Skripte hinweg.
        Ergaenzt die skripteigenen Protokolle, ersetzt sie nicht.
        Angehaengt wird zeilenweise, damit parallele Laeufe und harte
        Abbrueche nichts zerstoeren.
    #>
    param(
        [Parameter(Mandatory)][string]$Skript,
        [Parameter(Mandatory)][string]$Pfad,
        [Parameter(Mandatory)][string]$Aktion,
        [Parameter(Mandatory)][string]$Status,
        [string]$Detail = '',
        [string]$Protokolldatei
    )

    if ([string]::IsNullOrWhiteSpace($Protokolldatei)) {
        $basis = Resolve-WritableDirectory @($PSScriptRoot, $env:LOCALAPPDATA, $env:TEMP)
        if (-not $basis) { return }
        $Protokolldatei = Join-Path $basis 'migration.jsonl'
    }

    $satz = [ordered]@{
        ts     = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')
        skript = $Skript
        pfad   = (Remove-LongPathPrefix $Pfad)
        aktion = $Aktion
        status = $Status
    }
    if ($Detail) { $satz['detail'] = $Detail }

    try {
        $zeile = ($satz | ConvertTo-Json -Compress -Depth 4)
        [System.IO.File]::AppendAllText($Protokolldatei, $zeile + [Environment]::NewLine,
                                        (New-Object System.Text.UTF8Encoding($false)))
    } catch {
        # Ein nicht schreibbares Gesamtprotokoll darf den Lauf nicht
        # anhalten - die skripteigenen Protokolle laufen weiter.
    }
}

Export-ModuleMember -Function `
    Resolve-ScriptDirectory, Resolve-WritableDirectory,
    Add-LongPathPrefix, Remove-LongPathPrefix, Test-PathExistsLong,
    Get-RunningOfficeSessions, Show-OfficeRunningWarning, Stop-TrackedOfficeProcess,
    Wait-FileAvailable, Invoke-WithRetry,
    Get-FileTimestamps, Restore-FileTimestamps,
    ConvertTo-CsvField, Import-Konfiguration, Import-Presetliste,
    Write-Laufprotokoll
