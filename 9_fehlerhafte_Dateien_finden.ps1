# =====================================================================
# Fehlerhafte Office-Dateien finden (Word / Excel / PowerPoint)
# Voraussetzung: Windows PowerShell 5.1, .NET Framework, STA-Modus
#
# Hinweise zur .exe-Kompilierung (PS2EXE/ps2exe): mit -sta kompilieren, -noConsole nicht verwenden
# (Read-Host und Fortschritt werden sonst unsichtbar). PresentationFramework wird nicht mehr
# benoetigt - die drei Rueckfragen laufen seit der Umstellung auf Confirm-YesNo ueber die Konsole.
# Invoke-ps2exe -inputFile "9_fehlerhafte_Dateien_finden.ps1" -outputFile \
# "9_fehlerhafte_Dateien_finden.exe" -iconFile "powershell_icon.ico" -sta
#
# Stand: 11.06.2026
# =====================================================================

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ---------------------------------------------------------------------
# Rueckfragen
# ---------------------------------------------------------------------
function Confirm-YesNo {
    <#
        Ja/Nein-Rueckfrage auf der Konsole.

        Dieses Skript war das einzige der Sammlung, das seine drei
        Rueckfragen ueber [System.Windows.MessageBox]::Show stellte. Das
        setzt eine Fensterstation voraus: unter 'ps2exe -noConsole', in
        einer Remote-Sitzung ohne Desktop oder in einem geplanten Task
        erscheint das Fenster nicht oder unsichtbar hinter anderen - der
        Lauf stand dann ohne erkennbaren Grund. Alle anderen Skripte
        benutzen Read-Host; jetzt auch dieses.

        Ohne interaktive Konsole wird NICHT fortgefahren: die drei
        Rueckfragen betreffen alle einen Zustand, in dem ein blindes
        Weiterlaufen Office-Sitzungen des Anwenders gefaehrdet oder
        stundenlang in Timeouts laeuft.
    #>
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Message,
        [switch]$DefaultYes
    )

    # Nicht am Hostnamen festmachen: ps2exe - der im Kopf dieser Datei
    # vorgeschriebene Auslieferungsweg - stellt einen eigenen PSHost bereit,
    # dessen Name 'PSRunspace-Host' ist (ps2exe 1.0.18, ps2exe.ps1 Z. 2431-2435;
    # eine testweise kompilierte .exe meldet genau das). Mit der frueheren
    # Pruefung auf 'ConsoleHost' war $interactive in der ausgelieferten .exe
    # IMMER $false: alle drei Rueckfragen lieferten ohne Zutun $false und das
    # Programm brach vor der ersten geprueften Datei ab, obwohl eine voll
    # funktionsfaehige Konsole samt Read-Host vorhanden war.
    # Massgeblich ist deshalb die tatsaechliche Eingabefaehigkeit: ist die
    # Standardeingabe NICHT umgeleitet, haengt eine echte Konsole daran und
    # Read-Host funktioniert - gleich, wie der Host heisst. Wirft der Zugriff
    # (gar keine Konsole, z. B. 'ps2exe -noConsole' oder ein Dienst), bleibt es
    # bei $false: dann waere die Rueckfrage unsichtbar und ein blindes
    # Weiterlaufen genau das, was der Kommentar oben ausschliesst.
    $interactive = $false
    try { $interactive = -not [Console]::IsInputRedirected } catch { $interactive = $false }

    Write-Host ""
    Write-Host ("=" * 70) -ForegroundColor Yellow
    Write-Host ("  " + $Title) -ForegroundColor Yellow
    Write-Host ("=" * 70) -ForegroundColor Yellow
    foreach ($zeile in ($Message -split "`n")) { Write-Host ("  " + $zeile.TrimEnd()) }
    Write-Host ("=" * 70) -ForegroundColor Yellow

    if (-not $interactive) {
        Write-Host "  Keine interaktive Konsole - es wird NICHT fortgefahren." -ForegroundColor Red
        return $false
    }

    $hint = if ($DefaultYes) { '[J/n]' } else { '[j/N]' }
    while ($true) {
        $answer = Read-Host ("  Fortfahren? " + $hint)
        if ([string]::IsNullOrWhiteSpace($answer)) { return [bool]$DefaultYes }
        switch -Regex ($answer.Trim()) {
            '^[JjYy]' { return $true }
            '^[Nn]'   { return $false }
            default   { Write-Host "  Bitte 'j' oder 'n' eingeben." -ForegroundColor DarkGray }
        }
    }
}

# ---------------------------------------------------------------------
# 1. Konfiguration
# ---------------------------------------------------------------------

$Config = @{
    MaxHyperlinkLength    = 255
    TimeoutSeconds        = 120
    RetryCount            = 1
    RetryDelayMs          = 2000
    GcIntervalFiles       = 500
    CsvCheckpointEvery    = 500
    ProgressInterval      = 25
    WatchdogPollMs        = 100
    OfficeStartDelayMs    = 800
    ZombieIdleSeconds     = 30
    WatchdogShutdownMs    = 2000
    TrustCenterTimeoutSec = 25
    DummyPassword         = ([Guid]::NewGuid().ToString() + "!Xq#")
    WordExtensions        = @(".docx",".doc",".docm",".dotx",".dotm")
    ExcelExtensions       = @(".xlsx",".xls",".xlsm",".xlsb",".xltx",".xltm")
    PptExtensions         = @(".pptx",".ppt",".pptm",".potx",".potm",".ppsx",".pps",".ppsm")
}

# Verzeichnis-Praesets fuer die Startauswahl. Leere oder nicht vorhandene
# Eintraege werden im Menue ausgeblendet, die Nummerierung bleibt
# trotzdem lueckenlos.
$script:DirectoryPresets = @(
    'Q:\'
    'R:\'
    'G:\Geteilte Ablagen'
    'G:\Meine Ablage'
    '\\server\dfs'
)


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


$officeExtensions = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)
@($Config.WordExtensions + $Config.ExcelExtensions + $Config.PptExtensions) |
    ForEach-Object { $officeExtensions.Add($_) | Out-Null }

# Verzeichnisse, die nicht durchsucht werden. Der wichtigste Eintrag ist
# der Papierkorb: dort liegen GELOESCHTE Office-Dateien unter Namen wie
# '$RIA3B7K.docx' - also mit Originalendung. Ohne Ausschluss oeffnete das
# Skript sie per COM und meldete sie als defekt. Die Folge waren
# Falschmeldungen ueber Dateien, die der Anwender bewusst geloescht hat,
# und bei blockierenden Dateien jeweils der volle Watchdog-Timeout.
$script:ExcludeDirNames = @(
    '$RECYCLE.BIN', 'System Volume Information', 'RECYCLER',
    '.git', '.svn', '.hg', 'node_modules', '__pycache__',
    '~snapshot', '.snapshot'
)

$script:LogFile = $null
$script:ShouldStop = $false
$script:ReportShownInExcel = $false

# Strg+C als Eingabe behandeln und in Such- und Pruefschleife pollen:
# ein add_CancelKeyPress-ScriptBlock wuerde auf dem Ctrl+C-Thread ohne
# PowerShell-Runspace laufen und dort fehlschlagen - der Abbruch kaeme
# dann doch hart und der Excel-Bericht ginge verloren.
$script:CtrlCAsInput = $false
try {
    # Ebenfalls nicht am Hostnamen festmachen (siehe Confirm-YesNo): unter
    # ps2exe heisst der Host 'PSRunspace-Host', der sanfte Strg+C-Abbruch
    # blieb in der ausgelieferten .exe daher wirkungslos. Ob eine Konsole da
    # ist, zeigt der Zugriff selbst am zuverlaessigsten.
    if (-not [Console]::IsInputRedirected) {
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

# ---------------------------------------------------------------------
# 2. Hilfsfunktionen
# ---------------------------------------------------------------------

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR","DEBUG")]
        [string]$Level = "INFO",
        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Write-Host $Message -ForegroundColor $Color
    if ($script:LogFile) {
        try {
            Add-Content -Path $script:LogFile -Value "[$ts] [$Level] $Message" -Encoding UTF8
        } catch {}
    }
}

function Wait-ForAnyKey {
    Write-Host "`nBeliebige Taste drücken zum Beenden..." -ForegroundColor Gray
    try { $null = [System.Console]::ReadKey($true); return } catch {}
    try { $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown"); return } catch {}
    try { Read-Host | Out-Null } catch {}
}

# Whitelist fuer den Windows-Temp-Cleanup: NUR Eintraege mit diesen
# Praefixen (Office/COM-Reste sowie eigene Skript-Artefakte) duerfen
# geloescht werden. NIEMALS pauschal leeren: Fremdprozesse legen
# aktive Daten ohne Lock in %TEMP% ab; blindes Loeschen zerstoert sie.
$script:WindowsTempWhitelistPrefixes = @(
    '~$', '~df', 'gen_py', 'vbe', 'excel8.0',
    '9_fehlerhafte_dateien_finden', 'trustcheck_'
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
                Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
            } else {
                Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
            }
        } catch {}
    }
}

function Stop-Script {
    <#
        Zentraler Ausstieg. Vorher stand an fuenf Stellen 'Wait-ForAnyKey;
        exit' - der Temp-Cleanup und das Zuruecksetzen von
        TreatControlCAsInput liefen dabei nie.
    #>
    param([int]$Code = 0)
    try { Invoke-WindowsTempCleanup } catch { }
    try { if ($script:CtrlCAsInput) { [Console]::TreatControlCAsInput = $false } } catch { }
    Wait-ForAnyKey
    exit $Code
}

function Get-RunningOfficeSessions {
    <#
        Liefert alle Office-Prozesse mit sichtbarem Hauptfenster, also
        echte Sitzungen des Anwenders (im Unterschied zu den unsichtbaren
        Automatisierungs-Zombies weiter unten).
    #>
    $found = @()
    foreach ($n in @("WINWORD","EXCEL","POWERPNT")) {
        try {
            $found += @(Get-Process -Name $n -ErrorAction SilentlyContinue |
                        Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero })
        } catch { }
    }
    return ,@($found)
}

function Remove-LongPathPrefix {
    param([string]$Path)
    if ([string]::IsNullOrEmpty($Path)) { return $Path }
    if ($Path -like "\\?\UNC\*") { return "\\" + $Path.Substring(8) }
    if ($Path -like "\\?\*")     { return $Path.Substring(4) }
    return $Path
}

function Test-IstVerschluesselt {
    <#
        Erkennt eine verschluesselte OOXML-Datei an der DATEISIGNATUR, bevor
        Office sie ueberhaupt zu Gesicht bekommt.

        Hintergrund: Der Passwortschutz wurde bisher allein aus der
        Office-Fehlermeldung abgeleitet ('passwort|password|kennwort|...').
        Excel meldet bei einer mit Oeffnungskennwort geschuetzten Datei aber
        den generischen HRESULT 0x800A03EC, der eine Zeile frueher bereits
        als 'Office-Fehler' abgefangen wird - die Kategorie 'Passwortschutz'
        war fuer Excel damit unerreichbar, und der Bericht nannte einen
        Anwendungsfehler statt der wahren Ursache.

        Eine verschluesselte OOXML-Datei ist kein ZIP ('PK'), sondern ein
        CFB-Container mit der Signatur D0 CF 11 E0 A1 B1 1A E1. Bei den
        MODERNEN Endungen (.docx/.xlsx/.pptx usw.) ist dieser Container ein
        eindeutiger Beleg fuer Verschluesselung. Fuer die ALTEN Formate
        (.doc/.xls/.ppt) sagt er nichts aus - sie sind immer CFB; dort bleibt
        es bei der bisherigen Auswertung.
    #>
    param([string]$Path)

    $modern = @('.docx','.docm','.dotx','.dotm',
                '.xlsx','.xlsm','.xltx','.xltm','.xlsb',
                '.pptx','.pptm','.potx','.potm','.ppsx','.ppsm')
    $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($modern -notcontains $ext) { return $false }

    try {
        $fs = [System.IO.File]::Open($Path, 'Open', 'Read', 'ReadWrite')
        try {
            $buf = New-Object byte[] 8
            if ($fs.Read($buf, 0, 8) -lt 8) { return $false }
        } finally { $fs.Dispose() }
    } catch {
        return $false   # nicht lesbar - andere Pruefungen melden das
    }

    $cfb = @(0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE7)
    $cfb[7] = 0xE1
    for ($i = 0; $i -lt 8; $i++) {
        if ($buf[$i] -ne $cfb[$i]) { return $false }
    }
    return $true
}

function Get-OfficeAppType {
    param([string]$Extension)
    $e = $Extension.ToLowerInvariant()
    if ($Config.WordExtensions  -contains $e) { return "Word" }
    if ($Config.ExcelExtensions -contains $e) { return "Excel" }
    if ($Config.PptExtensions   -contains $e) { return "PowerPoint" }
    return "Unbekannt"
}

function Get-FileMetadata {
    param([string]$Path)
    $clean = Remove-LongPathPrefix -Path $Path
    try {
        $fi = New-Object System.IO.FileInfo -ArgumentList $Path
        return @{
            SizeMB   = [math]::Round($fi.Length / 1MB, 2)
            Modified = $fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm")
        }
    } catch {
        try {
            $fi = New-Object System.IO.FileInfo -ArgumentList $clean
            return @{
                SizeMB   = [math]::Round($fi.Length / 1MB, 2)
                Modified = $fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm")
            }
        } catch {
            return @{ SizeMB = 0; Modified = "" }
        }
    }
}

function Get-EchtenHResult {
    <#
        Liefert den TATSAECHLICHEN HRESULT einer Ausnahme.

        Beim Aufruf einer .NET-Methode verpackt PowerShell die Ausnahme in
        eine System.Management.Automation.MethodInvocationException. Deren
        HResult ist konstant 0x80131501 - unabhaengig davon, was wirklich
        passiert ist. Damit war der komplette HRESULT-Switch der
        Fehlerkategorisierung auf diesem Weg tot: eine gesperrte Datei
        (0x80070020) und eine ohne Leserecht (0x80070005) kamen beide als
        derselbe Wrapper-Code an. Nachgestellt unter 5.1 mit einer exklusiv
        gesperrten Datei: aeusserer HResult 0x80131501, innerer 0x80070020.
        Deshalb die Kette der InnerExceptions durchgehen und den ersten
        Wert nehmen, der nicht der Wrapper-Code ist.
    #>
    param($Exception)
    $WRAPPER = 0x80131501
    $e = $Exception
    $tiefe = 0
    while ($null -ne $e -and $tiefe -lt 8) {
        if ($e.HResult -ne 0 -and $e.HResult -ne $WRAPPER) { return $e.HResult }
        $e = $e.InnerException
        $tiefe++
    }
    if ($null -ne $Exception) { return $Exception.HResult }
    return 0
}

function Test-FileAccessible {
    param([string]$Path)
    $fs = $null
    try {
        $fs = [System.IO.File]::Open($Path, 'Open', 'Read', 'ReadWrite')
        return @{ OK = $true; Error = $null; HResult = 0 }
    } catch {
        return @{ OK = $false; Error = $_.Exception.Message
                  HResult = (Get-EchtenHResult -Exception $_.Exception) }
    } finally {
        if ($null -ne $fs) { try { $fs.Close(); $fs.Dispose() } catch {} }
    }
}

function Test-OfficeAppAlive {
    param([object]$App)
    try { $null = $App.Version; return $true } catch { return $false }
}

function Close-ComObject {
    param([object]$ComObj, [string]$AppType)
    if ($null -eq $ComObj) { return }
    switch ($AppType) {
        "word"  { $ComObj.Close(0) }
        "excel" { $ComObj.Close($false) }
        "ppt"   { $ComObj.Close() }
    }
}

function New-TrackedOfficeApp {
    param([string]$ProgId, [string]$ProcessName)
    $before = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue).Id
    $app    = New-Object -ComObject $ProgId
    Start-Sleep -Milliseconds $Config.OfficeStartDelayMs
    $after  = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue).Id
    $newPid = $after | Where-Object { $_ -notin $before } | Select-Object -First 1
    $newStart = [DateTime]::MinValue
    if ($newPid) {
        try { $newStart = (Get-Process -Id $newPid -ErrorAction Stop).StartTime } catch {}
    }
    return @{ App = $app; Pid = $newPid; Start = $newStart }
}

function Set-OfficeAppSafe {
    param([object]$WordApp, [object]$ExcelApp, [object]$PptApp)
    try { $WordApp.Visible           = $false } catch {}
    try { $ExcelApp.Visible          = $false } catch {}
    try { $WordApp.DisplayAlerts     = 0 }      catch {}
    try { $ExcelApp.DisplayAlerts    = $false } catch {}
    try { $ExcelApp.AskToUpdateLinks = $false } catch {}
    try { $PptApp.DisplayAlerts      = 1 }      catch {}
    try { $WordApp.AutomationSecurity  = 3 } catch {}
    try { $ExcelApp.AutomationSecurity = 3 } catch {}
    try { $PptApp.AutomationSecurity   = 3 } catch {}
}

function Get-ErrorCategory {
    param([string]$Message, [int]$HResult, [string]$AppType)

    $msg = if ($Message) { $Message.ToLowerInvariant() } else { "" }
    $hex = "0x" + $HResult.ToString("X8")

    # Pfadlaengen-Limit: vom Aufrufer wird " [HINWEIS: Pfad N Zeichen -
    # moeglicherweise pfadlaengenbedingt.]" an die Fehlermeldung gehaengt,
    # wenn comPath > 259 Zeichen ist. Office-COM versteht keinen \\?\-Praefix,
    # also schlaegt Open auf langen Pfaden fehl - meist mit derselben
    # generischen 0x800A03EC-Meldung ("Open-Eigenschaft kann nicht
    # zugeordnet werden"), die ohne Hinweis als "Office-Fehler" verbucht
    # wuerde. Eigene Kategorie macht die Ursache im Bericht direkt
    # erkennbar (und bietet eine konkrete Massnahme an).
    if ($msg -match "pfadl.ngenbedingt|pfad \d+ zeichen") {
        return @{
            Kategorie = "Pfadlaenge"
            Details   = "Pfad ueber 260 Zeichen - Office-COM kann diese Datei nicht oeffnen (kein Long-Path-Support). Massnahme: Datei in kuerzeren Pfad kopieren oder in Windows LongPathsEnabled aktivieren (HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem\LongPathsEnabled=1, Neustart). Ursprungstext: $Message"
        }
    }

    # Namenskonflikt VOR HRESULT-Switch: dieser modale Excel-Dialog
    # ("Name darf integriertem Namen nicht gleichen", englisch "name must
    # not be the same as a built-in name") wirft 0x800A03EC, wuerde also
    # sonst als generischer "Office-Fehler" verbucht. Klare Klassifikation
    # hilft, das Symptom im Bericht direkt zu erkennen.
    if ($msg -match "integriertem namen|integrierter name|built-in name|name darf|reservierter name") {
        return @{
            Kategorie = "Namenskonflikt"
            Details   = "Modaler Excel-Dialog 'Name darf integriertem Namen nicht gleichen' blockiert das Open. Ursache: User-definedName in der Datei kollidiert mit Excel-Built-In (z.B. Print_Area, Print_Titles, _FilterDatabase, Druckbereich). Manuelles Wegklicken noetig, oder Datei vorher mit Skript 3b bereinigen (OOXML-Pre-Clean). Ursprungstext: $Message"
        }
    }

    switch ($hex) {
        "0x800A03EC" { return @{ Kategorie = "Office-Fehler";      Details = "Office meldet einen allgemeinen Anwendungsfehler (0x800A03EC). Ursprungstext: $Message" } }
        "0x80070020" { return @{ Kategorie = "Datei gesperrt";     Details = "Datei wird durch anderen Prozess verwendet (Sharing Violation, 0x80070020). Ursprungstext: $Message" } }
        "0x80070005" { return @{ Kategorie = "Zugriff verweigert"; Details = "Keine Leseberechtigung (0x80070005). Ursprungstext: $Message" } }
        "0x800706BA" { return @{ Kategorie = "Office-Absturz";     Details = "RPC-Server nicht verfügbar (0x800706BA) — Office-Prozess instabil. Ursprungstext: $Message" } }
        "0x800706BE" { return @{ Kategorie = "Office-Absturz";     Details = "RPC-Aufruf fehlgeschlagen (0x800706BE). Ursprungstext: $Message" } }
        "0x80030002" { return @{ Kategorie = "Datei fehlt";        Details = "Datei nicht gefunden beim Öffnen. Ursprungstext: $Message" } }
        "0x800A1066" { return @{ Kategorie = "Ungültige Daten";    Details = "Excel meldet ungültige Daten (0x800A1066). Ursprungstext: $Message" } }
    }

    if ($msg -match "passwort|password|kennwort|verschlüssel|encrypt") {
        return @{ Kategorie = "Passwortschutz"; Details = "Datei ist passwortgeschützt oder verschlüsselt. Ursprungstext: $Message" }
    }
    if ($msg -match "beschädig|corrupt|damaged|unlesbar|nicht lesbar|cannot be read|unreadable|invalid format") {
        return @{ Kategorie = "Korruption"; Details = "Datei scheint beschädigt zu sein. Ursprungstext: $Message" }
    }
    if ($msg -match "gesperrt|locked|in use|wird verwendet|wird bearbeitet|sharing violation") {
        return @{ Kategorie = "Datei gesperrt"; Details = "Datei wird gerade durch anderen Prozess oder Benutzer verwendet. Ursprungstext: $Message" }
    }
    if ($msg -match "zugriff|access denied|permission|berechtigung|unauthorized") {
        return @{ Kategorie = "Zugriff verweigert"; Details = "Keine Lese-/Öffnungsberechtigung. Ursprungstext: $Message" }
    }
    if ($msg -match "format|file type|dateityp|erweiterung|extension") {
        return @{ Kategorie = "Formatfehler"; Details = "Format nicht erkannt oder Erweiterung passt nicht zum Inhalt. Ursprungstext: $Message" }
    }
    if ($msg -match "rpc|remoteprozeduraufruf|server execution|server executes|call failed") {
        return @{ Kategorie = "Office-Absturz"; Details = "COM-/RPC-Fehler — Office-Instanz instabil. Ursprungstext: $Message" }
    }
    if ($msg -match "timeout|zeitüberschreitung") {
        return @{ Kategorie = "Timeout"; Details = $Message }
    }
    if ($msg -match "not found|nicht gefunden|existiert nicht|does not exist") {
        return @{ Kategorie = "Datei fehlt"; Details = "Datei konnte nicht gefunden werden (Laufwerk offline?). Ursprungstext: $Message" }
    }
    if ($msg -match "network|netzwerk|verbindung|connection") {
        return @{ Kategorie = "Netzwerkfehler"; Details = "Netzwerkzugriff fehlgeschlagen. Ursprungstext: $Message" }
    }

    return @{ Kategorie = "Unbekannt"; Details = "Nicht klassifizierbarer Fehler. Ursprungstext: $Message (HRESULT $hex)" }
}

function Test-IsTransientError {
    param([int]$HResult, [string]$Message)
    $hex = "0x" + $HResult.ToString("X8")
    # Namenskonflikt: retry hilft nicht, Dialog blockiert weiter
    if ($Message -match "integriertem namen|integrierter name|built-in name|name darf|reservierter name") { return $false }
    $transient = @("0x800706BA","0x800706BE","0x80070040","0x80070071","0x8007003B","0x80040154","0x80010108")
    if ($transient -contains $hex) { return $true }
    if ($Message -match "network|netzwerk|rpc|connection|verbindung|temporar|temporar|unavailable|not available") { return $true }
    return $false
}

function New-ErrorRecord {
    param(
        [string]$Ordner,
        [string]$DateiPfad,
        [string]$Kategorie,
        [string]$Details,
        [string]$Dateityp   = "",
        [double]$SizeMB     = 0,
        [string]$Modified   = ""
    )

    # Gemeinsames Laufprotokoll (migration.jsonl) - ergaenzt das
    # skripteigene Protokoll, ersetzt es nicht. Erst damit laesst sich
    # der Fortschritt ueber alle elf Schritte hinweg auswerten.
    if ($script:GemeinsamGeladen) {
        try {
            Write-Laufprotokoll -Skript '9_fehlerhafte_Dateien_finden' `
                -Pfad $DateiPfad -Aktion 'Pruefung' `
                -Status $Kategorie -Detail $Details
        } catch { }
    }
    return [PSCustomObject]@{
        Ordner           = $Ordner
        Datei_Pfad       = $DateiPfad
        Kategorie        = $Kategorie
        Details          = $Details
        Dateityp         = $Dateityp
        Dateigroesse_MB  = $SizeMB
        Letzte_Aenderung = $Modified
    }
}

function Get-FilesIterative {
    param(
        [string]$RootPath,
        [System.Collections.Generic.HashSet[string]]$Extensions,
        [System.Collections.Generic.List[string]]$FileList,
        [System.Collections.Generic.List[object]]$ErrorList,
        [System.Collections.Generic.HashSet[string]]$VisitedPaths,
        [scriptblock]$OnProgress = $null,  # Optional: alle N Dateien fuer Live-Feedback
        [int]$ProgressInterval   = 250
    )

    $stack = New-Object 'System.Collections.Generic.Stack[string]'
    $stack.Push($RootPath)

    while ($stack.Count -gt 0) {
        # Ohne diese Pruefung wuerde Strg+C in der Suchphase ignoriert: der
        # Handler setzt zwar $script:ShouldStop, aber bis zur Pruefschleife
        # rattert die Verzeichnis-Enumeration auf gigantischen DFS-Shares
        # potenziell 45+ Minuten weiter. Die bis hier gesammelten Daten in
        # $FileList/$ErrorList bleiben durch das saubere break erhalten.
        Test-AbortRequested
        if ($script:ShouldStop) { break }

        $current   = $stack.Pop()
        $clean     = Remove-LongPathPrefix -Path $current
        $canonical = $clean
        try { $canonical = [System.IO.Path]::GetFullPath($clean) } catch {}

        if (-not $VisitedPaths.Add($canonical)) {
            $ErrorList.Add((New-ErrorRecord -Ordner "ZIRKELBEZUG" -DateiPfad $clean `
                -Kategorie "Verzeichnis" -Details "Symlink-/DFS-Zyklus — übersprungen"))
            continue
        }

        $dirSeen = [System.Collections.Generic.HashSet[string]]::new(
            [System.StringComparer]::OrdinalIgnoreCase
        )

        # Einmaliger Enumerate-Aufruf pro Ordner, lokale Filterung gegen das
        # HashSet. Spart bei 19 Office-Endungen den 19-fachen Netzwerk-Roundtrip
        # pro Verzeichnis (kritisch auf DFS/SMB).
        try {
            foreach ($f in [System.IO.Directory]::EnumerateFiles($current)) {
                $fName = [System.IO.Path]::GetFileName($f)
                if ($fName.StartsWith("~`$")) { continue }
                $actualExt = [System.IO.Path]::GetExtension($f)
                if ([string]::IsNullOrEmpty($actualExt)) { continue }
                if (-not $Extensions.Contains($actualExt)) { continue }
                if ($dirSeen.Add($f)) {
                    # WICHTIG: Truthy-Check ($FileList) statt $null-Vergleich
                    # ist hier ein klassischer PowerShell-Stolperstein.
                    # Eine leere Generic.List<T> wird von PowerShell ueber
                    # die IList-Interface-Regel zu $false ausgewertet
                    # (LanguagePrimitives.IsTrue greift bei IList auf
                    # Count != 0 zurueck). Ergebnis: die ALLERERSTE Datei
                    # wuerde nie in die Liste hinzugefuegt - die Liste
                    # bliebe ewig leer, der Scan meldete "0 Dateien".
                    # $OnProgress ist ein ScriptBlock-Parameter, dort ist der
                    # Truthy-Check zwar harmlos, aber konsistent auf
                    # $null-Vergleich umgestellt.
                    if ($null -ne $FileList) {
                        $FileList.Add($f)
                        if (($null -ne $OnProgress) -and ($FileList.Count % $ProgressInterval -eq 0)) {
                            & $OnProgress $FileList.Count
                        }
                    }
                }
            }
        }
        catch [System.UnauthorizedAccessException] {
            $ErrorList.Add((New-ErrorRecord -Ordner "ZUGRIFF VERWEIGERT" -DateiPfad $clean `
                -Kategorie "Zugriff verweigert" -Details "Keine Leseberechtigung (Dateien in: $clean)"))
        }
        catch [System.ArgumentException] {
            # Long-Path-Praefix wird vom Framework nicht unterstuetzt: nach
            # aussen werfen, damit der Fallback ohne \\?\-Praefix greift.
            # ABER NUR wenn der Fehler direkt beim Root auftritt - das ist
            # das Symptom "Framework kann \\?\ generell nicht". Bei tieferen
            # Pfaden ist es ein einzelner kaputter Ordner (unzulaessige
            # Zeichen, Reserved Name, etc.); den nur loggen, sonst geht
            # der gesammelte Fortschritt aus tausenden Dateien verloren.
            if ($current -eq $RootPath -and $current -like "\\?\*") { throw }
            $ErrorList.Add((New-ErrorRecord -Ordner "ENUMERATIONSFEHLER" -DateiPfad $clean `
                -Kategorie "Pfadfehler" -Details "Ungültiger Pfad (Dateisuche): $($_.Exception.Message)"))
        }
        catch [System.NotSupportedException] {
            if ($current -eq $RootPath -and $current -like "\\?\*") { throw }
            $ErrorList.Add((New-ErrorRecord -Ordner "ENUMERATIONSFEHLER" -DateiPfad $clean `
                -Kategorie "Pfadfehler" -Details "Pfadformat nicht unterstützt: $($_.Exception.Message)"))
        }
        catch {
            $ErrorList.Add((New-ErrorRecord -Ordner "ENUMERATIONSFEHLER" -DateiPfad $clean `
                -Kategorie "Pfadfehler" -Details "Fehler bei Dateisuche: $($_.Exception.Message)"))
        }

        try {
            foreach ($dir in [System.IO.Directory]::EnumerateDirectories($current)) {
                $leaf = [System.IO.Path]::GetFileName($dir.TrimEnd('\'))
                if ($script:ExcludeDirNames -contains $leaf) { continue }
                # Junctions/Symlinks nicht folgen: sonst werden Dateien
                # doppelt geprueft oder die Suche laeuft im Kreis.
                try {
                    $dirAttr = [System.IO.File]::GetAttributes($dir)
                    if ($dirAttr -band [System.IO.FileAttributes]::ReparsePoint) { continue }
                } catch { continue }
                $stack.Push($dir)
            }
        }
        catch [System.UnauthorizedAccessException] {
            $ErrorList.Add((New-ErrorRecord -Ordner "ZUGRIFF VERWEIGERT" -DateiPfad $clean `
                -Kategorie "Zugriff verweigert" -Details "Keine Leseberechtigung (Unterordner von: $clean)"))
        }
        catch [System.ArgumentException] {
            if ($current -eq $RootPath -and $current -like "\\?\*") { throw }
            $ErrorList.Add((New-ErrorRecord -Ordner "ENUMERATIONSFEHLER" -DateiPfad $clean `
                -Kategorie "Pfadfehler" -Details "Ungültiger Pfad (Verzeichnissuche): $($_.Exception.Message)"))
        }
        catch [System.NotSupportedException] {
            if ($current -eq $RootPath -and $current -like "\\?\*") { throw }
            $ErrorList.Add((New-ErrorRecord -Ordner "ENUMERATIONSFEHLER" -DateiPfad $clean `
                -Kategorie "Pfadfehler" -Details "Pfadformat nicht unterstützt (Verzeichnissuche): $($_.Exception.Message)"))
        }
        catch {
            $ErrorList.Add((New-ErrorRecord -Ordner "ENUMERATIONSFEHLER" -DateiPfad $clean `
                -Kategorie "Pfadfehler" -Details "Fehler bei Verzeichnissuche: $($_.Exception.Message)"))
        }
    }
}

function Save-Checkpoint {
    param(
        [System.Collections.Generic.List[object]]$Results,
        [string]$Path
    )
    try {
        $Results | Export-Csv -Path $Path -Delimiter ";" -Encoding UTF8 -NoTypeInformation
    } catch {}
}

# ---------------------------------------------------------------------
# 3. STA-Modus prüfen
# ---------------------------------------------------------------------

if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Write-Host "Skript benötigt STA-Modus. Bitte mit 'powershell.exe -STA -File ...' starten." -ForegroundColor Red
    Stop-Script
}

# ---------------------------------------------------------------------
# 4. Skript-Begruessung und Trust-Center-Hinweise
# ---------------------------------------------------------------------

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " Fehlerhafte Office-Dateien finden (Word/Excel/PowerPoint)" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Was macht dieses Skript:" -ForegroundColor Yellow
Write-Host "  - Durchlaeuft rekursiv ein Verzeichnis (Word/Excel/PowerPoint)"
Write-Host "  - Versucht jede Datei stillschweigend per COM zu oeffnen"
Write-Host "  - Klassifiziert Fehler (Korruption, Passwortschutz, Sperre,"
Write-Host "    Office-Absturz, Namenskonflikt, etc.)"
Write-Host "  - Erzeugt einen Excel-Bericht und ein detailliertes Logfile"
Write-Host "    NEBEN dem Skript."
Write-Host ""
Write-Host "Vor dem Start:" -ForegroundColor Yellow
Write-Host "  Trust-Center-Einstellungen in WORD, EXCEL und POWERPOINT:"
Write-Host "    Datei -> Optionen -> Trust Center -> Einstellungen fuer das"
Write-Host "    Trust Center -> Geschuetzte Ansicht:"
Write-Host "      [ ] Geschuetzte Ansicht fuer Dateien aus dem Internet"
Write-Host "      [ ] Geschuetzte Ansicht fuer Dateien an unsicheren Orten"
Write-Host "      [ ] Geschuetzte Ansicht fuer Outlook-Anhaenge"
Write-Host "    (alle drei Haken DEAKTIVIEREN)"
Write-Host ""
Write-Host "    Trust Center -> Einstellungen fuer Makros:"
Write-Host "      (o) Alle Makros aktivieren"
Write-Host ""
Write-Host "  Sonst erscheinen modale Dialoge, die COM blockieren und das"
Write-Host "  Skript pro Datei in einen Timeout laufen lassen."
Write-Host ""

$tcTitle   = "Konfiguration prüfen"
$tcMessage = "Haben Sie die Einstellungen im Trust Center (Geschützte Ansicht DEAKTIVIERT, Makros AKTIVIERT) für Word, Excel und PowerPoint vorgenommen?"
$tcResult = Confirm-YesNo -Title $tcTitle -Message $tcMessage

if (-not $tcResult) {
    Write-Host "Bitte konfigurieren Sie zuerst das Trust Center in den Office-Optionen. Abbruch." -ForegroundColor Red
    Stop-Script
}

# ---------------------------------------------------------------------
# 4b. Hilfsfunktion: Desktop/Downloads ueber User-Shell-Folder aufloesen
# ---------------------------------------------------------------------

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
            if (Test-Path -LiteralPath $expanded) { return $expanded }
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

# ---------------------------------------------------------------------
# 5. Pfadauswahl
# ---------------------------------------------------------------------

$desktopDir   = Get-UserShellFolder -Name 'Desktop'
$downloadsDir = Get-UserShellFolder -Name 'Downloads'

# Menue aus den Praesets aufbauen, statt die Nummern fest zu verdrahten.
# Vorher standen nicht konfigurierte Praesets als tote Eintraege im Menue;
# wer sie waehlte, lief in einen Enumerationsfehler statt in eine
# verstaendliche Meldung.
$pathChoices = New-Object System.Collections.Generic.List[string]
foreach ($p in $script:DirectoryPresets) {
    if (-not [string]::IsNullOrWhiteSpace($p)) { $pathChoices.Add($p) }
}
$pathChoices.Add($desktopDir)
$pathChoices.Add($downloadsDir)
$manualIdx = $pathChoices.Count + 1

Write-Host "Zielpfad auswaehlen:" -ForegroundColor Yellow
for ($i = 0; $i -lt $pathChoices.Count; $i++) {
    $label = switch ($pathChoices[$i]) {
        $desktopDir   { "Desktop ($desktopDir)" }
        $downloadsDir { "Downloads ($downloadsDir)" }
        default       { $pathChoices[$i] }
    }
    Write-Host ("  [{0}] {1}" -f ($i + 1), $label) -ForegroundColor White
}
Write-Host ("  [{0}] Eigenen Pfad eingeben" -f $manualIdx) -ForegroundColor White
Write-Host ""
$choice = (Read-Host ("Auswahl [1-{0}]" -f $manualIdx)).Trim()

$choiceNum = 0
if (-not [int]::TryParse($choice, [ref]$choiceNum)) { $choiceNum = -1 }

if ($choiceNum -ge 1 -and $choiceNum -le $pathChoices.Count) {
    $rootPath = $pathChoices[$choiceNum - 1]
}
elseif ($choiceNum -eq $manualIdx) {
    do {
        $rawInput = Read-Host "Geben Sie den gewünschten Pfad ein"
        $rootPath = if ($null -ne $rawInput) { $rawInput.Trim().Trim([char[]]@('"',"'")) } else { '' }
        if ([string]::IsNullOrWhiteSpace($rootPath)) {
            Write-Host "Pfad darf nicht leer sein." -ForegroundColor Yellow
        }
        elseif (-not (Test-Path -LiteralPath $rootPath -PathType Container)) {
            Write-Host "Pfad '$rootPath' existiert nicht oder ist kein Verzeichnis." -ForegroundColor Yellow
            $rootPath = ""
        }
    } while ([string]::IsNullOrWhiteSpace($rootPath))
}
else {
    Write-Host ("Ungueltige Auswahl: '{0}'. Erlaubt: 1-{1}." -f $choice, $manualIdx) -ForegroundColor Red
    Stop-Script 1
}

# Preset-Pfade koennen auf diesem Rechner fehlen (Laufwerk nicht gemappt).
if (-not (Test-Path -LiteralPath $rootPath -PathType Container)) {
    Write-Host "Pfad '$rootPath' existiert nicht oder ist kein Verzeichnis." -ForegroundColor Red
    Write-Host "Hinweis: Bei erhoehten Rechten sind gemappte Netzlaufwerke oft ausgeblendet -" -ForegroundColor Yellow
    Write-Host "dann bitte den UNC-Pfad ueber die manuelle Eingabe verwenden." -ForegroundColor Yellow
    Stop-Script 1
}

# ---------------------------------------------------------------------
# 5b. Fortschrittsmodus
# ---------------------------------------------------------------------

Write-Host ""
# Hinweis zur Formulierung: Die Menuetexte versprachen frueher bei [3]
# einen "sofortigen Start" und bei [2] einen zusaetzlichen "Vorab-Scan".
# Beides traf nicht zu - alle drei Varianten sammeln erst die komplette
# Dateiliste und pruefen danach. Der Unterschied liegt allein in der
# Darstellung waehrend des Sammelns und der Pruefung.
Write-Host "Fortschrittsanzeige:" -ForegroundColor Yellow
Write-Host "  [1] Keine Fortschrittsleiste, laufende Zaehlung als Textzeilen [Standard]" -ForegroundColor White
Write-Host "  [2] Fortschrittsleiste mit Restzeit-Schaetzung (ETA)" -ForegroundColor White
Write-Host "  [3] Fortschrittsleiste ohne Restzeit-Schaetzung" -ForegroundColor White
Write-Host ""
Write-Host "  Hinweis: Bei allen drei Varianten wird zuerst die vollstaendige" -ForegroundColor DarkGray
Write-Host "  Dateiliste gesammelt. Auf UNC-Pfaden oder sehr grossen Freigaben" -ForegroundColor DarkGray
Write-Host "  kann das 10-30 Minuten dauern, bevor die erste Datei geprueft wird." -ForegroundColor DarkGray
Write-Host "  Variante [1] und [3] melden den Zwischenstand schon waehrend" -ForegroundColor DarkGray
Write-Host "  des Sammelns, [2] bleibt in dieser Phase still." -ForegroundColor DarkGray
Write-Host ""
$progressChoice = (Read-Host "Auswahl [1-3]").Trim()
switch ($progressChoice) {
    "2"     { $script:UseProgress = $true;  $script:SkipPreScan = $false }
    "3"     { $script:UseProgress = $true;  $script:SkipPreScan = $true  }
    default { $script:UseProgress = $false; $script:SkipPreScan = $true  }
}

# ---------------------------------------------------------------------
# 6. Pfad-Normalisierung und Long-Path-Behandlung
# ---------------------------------------------------------------------

if ($rootPath -match "^[A-Za-z]:$") { $rootPath += "\" }

$resolvedPath = $rootPath
if ($rootPath -match "^[A-Za-z]:\\") {
    try {
        $driveLetter = $rootPath.Substring(0, 1)
        $psDrive     = Get-PSDrive -Name $driveLetter -ErrorAction SilentlyContinue
        if ($psDrive -and $psDrive.DisplayRoot) {
            $resolvedPath = $psDrive.DisplayRoot + $rootPath.Substring(2)
            Write-Host "Gemapptes Laufwerk aufgelöst: $rootPath -> $resolvedPath" -ForegroundColor DarkGray
        }
    } catch {}
}

$longPathsEnabled = $false
try {
    $lpe = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem" `
        -Name "LongPathsEnabled" -ErrorAction SilentlyContinue
    if ($lpe -and $lpe.LongPathsEnabled -eq 1) { $longPathsEnabled = $true }
} catch {}

$longPath = $resolvedPath
if (-not $longPathsEnabled -and $longPath -notlike "\\?\*") {
    if ($longPath -like "\\*") {
        $longPath = "\\?\UNC\" + $longPath.TrimStart('\')
    } else {
        $longPath = "\\?\" + $longPath
    }
}

# ---------------------------------------------------------------------
# 7. Berichts- und Logpfade
# ---------------------------------------------------------------------

$timestamp       = (Get-Date).ToString("yyyy-MM-dd_HHmmss")
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
$scriptDir       = Resolve-ScriptDirectory

# Ausgabeverzeichnis mit Fallback-Kette. Vorher wurde ungeprueft neben das
# Skript geschrieben: lag es auf einer schreibgeschuetzten Freigabe, schlug
# am Ende sowohl SaveAs als auch der CSV-Notfall-Export fehl - und damit
# waren die Ergebnisse eines stundenlangen Scans verloren. Der Schreibtest
# ist echt (Datei anlegen und wieder loeschen), denn ein existierendes
# Verzeichnis sagt nichts ueber das Schreibrecht aus.
$outDir = $null
foreach ($cand in @($scriptDir,
                    (Join-Path $env:LOCALAPPDATA 'OfficeFileCheck'),
                    $env:TEMP)) {
    if ([string]::IsNullOrWhiteSpace($cand)) { continue }
    try {
        if (-not (Test-Path -LiteralPath $cand)) {
            New-Item -ItemType Directory -Path $cand -Force -ErrorAction Stop | Out-Null
        }
        $probe = Join-Path $cand (".writetest_{0}.tmp" -f ([Guid]::NewGuid().ToString("N")))
        [System.IO.File]::WriteAllText($probe, 'x')
        [System.IO.File]::Delete($probe)
        $outDir = $cand
        break
    } catch { continue }
}
if (-not $outDir) {
    Write-Host "Kein beschreibbares Verzeichnis fuer Bericht und Log gefunden." -ForegroundColor Red
    Write-Host "Geprueft: '$scriptDir', LOCALAPPDATA, TEMP." -ForegroundColor Red
    Stop-Script 1
}

$reportPath      = [System.IO.Path]::Combine($outDir, "9_fehlerhafte_Dateien_finden_$timestamp.xlsx")
$logPath         = [System.IO.Path]::Combine($outDir, "9_fehlerhafte_Dateien_finden_$timestamp.log")
$checkpointPath  = [System.IO.Path]::Combine($outDir, "9_fehlerhafte_Dateien_finden_$timestamp.checkpoint.csv")
# Arbeitsverzeichnis fuer die Smoke-Test-Dateien. Vorher lag es im
# Dokumente-Ordner: der wird auf vielen Rechnern in OneDrive/Google Drive
# synchronisiert, sodass jeder Lauf drei Testdateien durch die Cloud
# schickte - und bei einem Absturz zwischen Anlegen und Loeschen blieben
# sie dort liegen. %TEMP% ist der richtige Ort und wird ausserdem von
# Invoke-WindowsTempCleanup (Praefix 'trustcheck_') mit abgeraeumt.
$smokeTempDir = $env:TEMP
if ([string]::IsNullOrWhiteSpace($smokeTempDir) -or -not (Test-Path -LiteralPath $smokeTempDir)) {
    $smokeTempDir = [Environment]::GetFolderPath('MyDocuments')
}
$script:LogFile  = $logPath

Write-Log -Message "Bericht: $reportPath" -Color Gray
Write-Log -Message "Log:     $logPath"    -Color Gray
if ($outDir -ne $scriptDir) {
    Write-Log -Message "Hinweis: Skriptordner ist nicht beschreibbar - Ausgabe geht nach '$outDir'." -Level WARN -Color Yellow
}

# ---------------------------------------------------------------------
# 8. Zombie-Prozesse bereinigen
# ---------------------------------------------------------------------

# --- Offene Office-Sitzungen des Anwenders ---
# Das ist der gefaehrlichste Zustand fuer dieses Skript, und er wurde
# bisher gar nicht geprueft:
#   1. 'New-Object -ComObject Word.Application' startet KEINE neue
#      Instanz, wenn Word bereits laeuft - COM haengt sich an die
#      vorhandene. Set-OfficeAppSafe setzt danach Visible=$false und
#      DisplayAlerts=0, das Fenster des Anwenders verschwindet also.
#   2. Am Skriptende ruft der finally-Block $word.Quit() auf und beendet
#      damit die Sitzung des Anwenders samt ungespeicherter Dokumente.
#   3. New-TrackedOfficeApp findet in diesem Fall keinen NEUEN Prozess,
#      $wordPid bleibt $null - der Watchdog prueft 'TargetPid -gt 0' und
#      ist damit wirkungslos. Blockierende Dateien haengen endlos.
#   4. Greift der Watchdog doch (bei anderer App), wuerde Stop-Process
#      die Office-Sitzung des Anwenders hart abschiessen.
$openOfficeSessions = Get-RunningOfficeSessions
if ($openOfficeSessions.Count -gt 0) {
    $namen = ($openOfficeSessions | ForEach-Object { $_.ProcessName } | Select-Object -Unique) -join ', '
    Write-Log -Message "" -Color Yellow
    Write-Log -Message "ACHTUNG: Es laufen bereits Office-Sitzungen ($namen)." -Level WARN -Color Red
    Write-Log -Message "  Das Skript steuert Office ueber COM. Bei laufenden Sitzungen bedeutet das:" -Color Yellow
    Write-Log -Message "    - Ihre offenen Fenster werden ausgeblendet und Warnhinweise abgeschaltet." -Color Yellow
    Write-Log -Message "    - Ihre Sitzung wird am Ende nicht beendet, aber Warnhinweise bleiben abgeschaltet." -Color Yellow
    Write-Log -Message "    - Der Timeout-Watchdog ist wirkungslos, weil kein eigener Prozess erkannt wird." -Color Yellow
    Write-Log -Message "  Bitte Word, Excel und PowerPoint schliessen und das Skript neu starten." -Color Yellow
    Write-Log -Message "" -Color Yellow

    $sessMsg = "Es laufen bereits Office-Sitzungen ($namen).`n`n" +
               "Diese werden vom Skript uebernommen und waehrend des Laufs ausgeblendet. " +
               "Zusaetzlich ist der Timeout-Watchdog dann wirkungslos, weil kein eigener " +
               "Prozess erkannt wird - haengende Dateien blockieren den Lauf unbegrenzt.`n`n" +
               "Empfehlung: JETZT abbrechen, Office schliessen, Skript neu starten.`n`n" +
               "Trotzdem fortfahren?"
    $sessRes = Confirm-YesNo -Title "Office laeuft bereits" -Message $sessMsg
    if (-not $sessRes) {
        Write-Log -Message "Abbruch durch Benutzer - Office-Sitzungen laufen noch." -Color Yellow
        Stop-Script
    }
    Write-Log -Message "Benutzer setzt trotz laufender Office-Sitzungen fort." -Level WARN -Color Yellow
}

$now           = Get-Date
$idleThreshold = [TimeSpan]::FromSeconds($Config.ZombieIdleSeconds)
$zombieProcesses = @(
    @("WINWORD","EXCEL","POWERPNT") | ForEach-Object {
        Get-Process -Name $_ -ErrorAction SilentlyContinue |
            Where-Object {
                if ($_.MainWindowHandle -ne [IntPtr]::Zero) { $false }
                else {
                    $ok = $false
                    # Wenn StartTime nicht lesbar ist (Race im Init-Moment, fehlende
                    # Leserechte fuer fremde User-Sessions auf Terminalservern, AV-
                    # Hooking, etc.), bleibt $ok = $false: ein nicht verifizierbarer
                    # Prozess darf nicht hart beendet werden.
                    try { $ok = (($now - $_.StartTime) -gt $idleThreshold) } catch { $ok = $false }
                    $ok
                }
            }
    }
)

if ($zombieProcesses.Count -gt 0) {
    Write-Log -Message "`n$($zombieProcesses.Count) unsichtbare Office-Hintergrundprozesse gefunden (älter als $($Config.ZombieIdleSeconds)s)." -Level WARN -Color Yellow
    Write-Log -Message "ACHTUNG: Dies können auch laufende, aber minimierte oder gerade startende Office-Sitzungen sein." -Level WARN -Color Yellow
    $killChoice = Read-Host "Sollen diese vor dem Scan beendet werden? (j/n)"
    if ($killChoice -match "^[JjYy]") {
        foreach ($proc in $zombieProcesses) {
            try {
                $proc | Stop-Process -Force
                Write-Log -Message "  Beendet: $($proc.ProcessName) (PID $($proc.Id))" -Color DarkYellow
            } catch {
                Write-Log -Message "  Konnte nicht beendet werden: $($proc.ProcessName) (PID $($proc.Id))" -Level ERROR -Color Red
            }
        }
        Start-Sleep -Seconds 1
    }
}

# ---------------------------------------------------------------------
# 8b. Trust-Center Smoke-Test (Word / Excel / PowerPoint)
# ---------------------------------------------------------------------

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

function New-MinimalXlsx {
    param([string]$Path)

    if ([System.IO.File]::Exists($Path)) {
        [System.IO.File]::Delete($Path)
    }

    $contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
    $rels         = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    $workbook     = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'
    $wbRels       = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
    $sheet        = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>'

    $entries = @(
        @{ Name = "[Content_Types].xml";         Content = $contentTypes }
        @{ Name = "_rels/.rels";                 Content = $rels         }
        @{ Name = "xl/workbook.xml";             Content = $workbook     }
        @{ Name = "xl/_rels/workbook.xml.rels";  Content = $wbRels       }
        @{ Name = "xl/worksheets/sheet1.xml";    Content = $sheet        }
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
    param([string]$TempDir, [int]$TimeoutSec = 25)

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
            $doc = $word.Documents.Open($fp, $false, $true, $false)
            if (-not $doc) {
                return @{ Ok = $false; Msg = "Open lieferte kein Document"; AppPid = $myWordPid }
            }
            try { $doc.Close($false) } catch {}
            return @{ Ok = $true; AppPid = $myWordPid }
        } catch {
            return @{ Ok = $false; Msg = $_.Exception.Message; AppPid = $myWordPid }
        } finally {
            try { $word.Interactive = $true } catch {}
            # NUR eine selbst gestartete Instanz beenden. New-Object
            # -ComObject Word.Application startet KEINE neue Instanz, wenn
            # Word bereits laeuft - COM haengt sich an die vorhandene. Ein
            # bedingungsloses Quit() beendete damit die Sitzung des Anwenders
            # samt ungespeicherter Dokumente, und wegen DisplayAlerts = 0
            # sogar ohne Speichern-Rueckfrage. Der Aufraeumteil am Skriptende
            # macht es bereits richtig; die Warnung in Abschnitt 8 sichert dem
            # Anwender ausdruecklich zu 'Ihre Sitzung wird am Ende nicht
            # beendet' - direkt danach lief hier das Gegenteil.
            # $myWordPid ist oben bereits ermittelt und ist genau dann gesetzt,
            # wenn eine neue EXCEL/WINWORD-Instanz entstanden ist.
            if ($myWordPid) {
                try { $word.Quit() } catch {}
            } else {
                # Fremdsitzung uebernommen: nur wieder sichtbar machen und
                # die Referenz freigeben.
                try { $word.Visible = $true } catch {}
            }
            # Wenn $word null ist (Office nicht installiert, Lizenz, COM-
            # Init fehlgeschlagen), wirft FinalReleaseComObject sonst eine
            # ArgumentNullException und der Job crasht statt ein sauberes
            # Ok=$false zurueckzugeben - der Smoke-Test laeuft dann ins
            # Timeout statt direkt zu melden, dass Office nicht da ist.
            if ($null -ne $word) {
                try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null } catch {}
            }
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    } -ArgumentList $testDocx, $pidFile

    $completed = Wait-Job $job -Timeout $TimeoutSec

    if (-not $completed) {
        try { Stop-Job   $job -ErrorAction SilentlyContinue } catch {}
        try { Remove-Job $job -Force -ErrorAction SilentlyContinue } catch {}
        if ([System.IO.File]::Exists($pidFile)) {
            try {
                $savedPid = [int][System.IO.File]::ReadAllText($pidFile).Trim()
                if ($savedPid -gt 0) {
                    try {
                        $p = Get-Process -Id $savedPid -ErrorAction Stop
                        if ($p.Name -match '^(WINWORD|EXCEL|POWERPNT)$') {
                            Stop-Process -Id $savedPid -Force -ErrorAction SilentlyContinue
                        }
                    } catch {}
                }
            } catch {}
            try { [System.IO.File]::Delete($pidFile) } catch {}
        }
        try { [System.IO.File]::Delete($testDocx) } catch {}
        return @{ Ok = $false; Msg = "TIMEOUT nach ${TimeoutSec}s - Trust Center vermutlich nicht konfiguriert oder Office-Profil beschaedigt" }
    }

    $result = Receive-Job $job
    try { Remove-Job $job -Force -ErrorAction SilentlyContinue } catch {}
    if ([System.IO.File]::Exists($pidFile)) { try { [System.IO.File]::Delete($pidFile) } catch {} }
    try { [System.IO.File]::Delete($testDocx) } catch {}

    if ($result -and $result.Ok) { return @{ Ok = $true } }
    $m = if ($result) { $result.Msg } else { "Kein Ergebnis vom COM-Job" }
    return @{ Ok = $false; Msg = $m }
}

function Test-ExcelTrustCenter {
    param([string]$TempDir, [int]$TimeoutSec = 25)

    $testXlsx = Join-Path $TempDir ("trustcheck_{0}.xlsx" -f ([Guid]::NewGuid().ToString("N")))
    $pidFile  = Join-Path $TempDir ("trustcheck_{0}.pid"  -f ([Guid]::NewGuid().ToString("N")))

    try {
        New-MinimalXlsx -Path $testXlsx
    } catch {
        return @{ Ok = $false; Msg = "Test-XLSX konnte nicht erstellt werden: $_" }
    }

    $job = Start-Job -ScriptBlock {
        param($fp, $pidFile)

        $existingPids = @(Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })

        $ex = New-Object -ComObject Excel.Application
        $ex.Visible            = $false
        $ex.DisplayAlerts      = $false
        $ex.AskToUpdateLinks   = $false
        $ex.AutomationSecurity = 3
        try { $ex.Interactive = $false } catch {}

        Start-Sleep -Milliseconds 50
        $myExcelPid = $null
        foreach ($p in (Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue)) {
            if ($existingPids -notcontains $p.Id) { $myExcelPid = $p.Id; break }
        }
        if ($myExcelPid) { try { [System.IO.File]::WriteAllText($pidFile, "$myExcelPid") } catch {} }

        $wb = $null
        try {
            $wb = $ex.Workbooks.Open($fp, 0, $true)
            if (-not $wb) {
                return @{ Ok = $false; Msg = "Open lieferte kein Workbook"; AppPid = $myExcelPid }
            }
            try { $wb.Close($false) } catch {}
            return @{ Ok = $true; AppPid = $myExcelPid }
        } catch {
            return @{ Ok = $false; Msg = $_.Exception.Message; AppPid = $myExcelPid }
        } finally {
            if ($wb) { try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null } catch {} }
            try { $ex.Interactive = $true } catch {}
            try { $ex.Quit() } catch {}
            if ($null -ne $ex) {
                try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ex) | Out-Null } catch {}
            }
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    } -ArgumentList $testXlsx, $pidFile

    $completed = Wait-Job $job -Timeout $TimeoutSec

    if (-not $completed) {
        try { Stop-Job   $job -ErrorAction SilentlyContinue } catch {}
        try { Remove-Job $job -Force -ErrorAction SilentlyContinue } catch {}
        if ([System.IO.File]::Exists($pidFile)) {
            try {
                $savedPid = [int][System.IO.File]::ReadAllText($pidFile).Trim()
                if ($savedPid -gt 0) {
                    try {
                        $p = Get-Process -Id $savedPid -ErrorAction Stop
                        if ($p.Name -match '^(WINWORD|EXCEL|POWERPNT)$') {
                            Stop-Process -Id $savedPid -Force -ErrorAction SilentlyContinue
                        }
                    } catch {}
                }
            } catch {}
            try { [System.IO.File]::Delete($pidFile) } catch {}
        }
        try { [System.IO.File]::Delete($testXlsx) } catch {}
        return @{ Ok = $false; Msg = "TIMEOUT nach ${TimeoutSec}s - Trust Center vermutlich nicht konfiguriert oder Office-Profil beschaedigt" }
    }

    $result = Receive-Job $job
    try { Remove-Job $job -Force -ErrorAction SilentlyContinue } catch {}
    if ([System.IO.File]::Exists($pidFile)) { try { [System.IO.File]::Delete($pidFile) } catch {} }
    try { [System.IO.File]::Delete($testXlsx) } catch {}

    if ($result -and $result.Ok) { return @{ Ok = $true } }
    $m = if ($result) { $result.Msg } else { "Kein Ergebnis vom COM-Job" }
    return @{ Ok = $false; Msg = $m }
}

function Test-PowerPointTrustCenter {
    param([string]$TempDir, [int]$TimeoutSec = 25)

    $testPath = Join-Path $TempDir ("trustcheck_{0}.pptx" -f ([Guid]::NewGuid().ToString("N")))
    $pidFile  = Join-Path $TempDir ("trustcheck_{0}.pid"  -f ([Guid]::NewGuid().ToString("N")))

    $job = Start-Job -ScriptBlock {
        param($fp, $pidFile)

        $existingPids = @(Get-Process -Name "POWERPNT" -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })

        $ppt = New-Object -ComObject PowerPoint.Application
        try { $ppt.DisplayAlerts = 1 } catch {}
        try { $ppt.AutomationSecurity = 3 } catch {}

        Start-Sleep -Milliseconds 50
        $myPptPid = $null
        foreach ($p in (Get-Process -Name "POWERPNT" -ErrorAction SilentlyContinue)) {
            if ($existingPids -notcontains $p.Id) { $myPptPid = $p.Id; break }
        }
        if ($myPptPid) { try { [System.IO.File]::WriteAllText($pidFile, "$myPptPid") } catch {} }

        $pres1 = $null
        $pres2 = $null
        try {
            $pres1 = $ppt.Presentations.Add(-1)
            if (-not $pres1) {
                return @{ Ok = $false; Msg = "Presentations.Add lieferte keine Praesentation"; AppPid = $myPptPid }
            }
            $pres1.SaveAs($fp, 24)
            $pres1.Close()
            $pres1 = $null

            $pres2 = $ppt.Presentations.Open($fp, -1, -1, 0)
            if (-not $pres2) {
                return @{ Ok = $false; Msg = "Presentations.Open lieferte keine Praesentation"; AppPid = $myPptPid }
            }
            $pres2.Close()
            $pres2 = $null

            return @{ Ok = $true; AppPid = $myPptPid }
        } catch {
            return @{ Ok = $false; Msg = $_.Exception.Message; AppPid = $myPptPid }
        } finally {
            if ($pres1) { try { $pres1.Close() } catch {} }
            if ($pres2) { try { $pres2.Close() } catch {} }
            # Nur eine selbst gestartete Instanz beenden - Begruendung siehe
            # Word-Smoke-Test oben. Fuer PowerPoint gilt dasselbe, dort ist
            # DisplayAlerts = 1 (ppAlertsNone), also ebenfalls ohne Rueckfrage.
            if ($myPptPid) {
                try { $ppt.Quit() } catch {}
            } else {
                try { $ppt.Visible = $true } catch {}
            }
            if ($null -ne $ppt) {
                try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ppt) | Out-Null } catch {}
            }
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    } -ArgumentList $testPath, $pidFile

    $completed = Wait-Job $job -Timeout $TimeoutSec

    if (-not $completed) {
        try { Stop-Job   $job -ErrorAction SilentlyContinue } catch {}
        try { Remove-Job $job -Force -ErrorAction SilentlyContinue } catch {}
        if ([System.IO.File]::Exists($pidFile)) {
            try {
                $savedPid = [int][System.IO.File]::ReadAllText($pidFile).Trim()
                if ($savedPid -gt 0) {
                    try {
                        $p = Get-Process -Id $savedPid -ErrorAction Stop
                        if ($p.Name -match '^(WINWORD|EXCEL|POWERPNT)$') {
                            Stop-Process -Id $savedPid -Force -ErrorAction SilentlyContinue
                        }
                    } catch {}
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

    if ($result -and $result.Ok) { return @{ Ok = $true } }
    $m = if ($result) { $result.Msg } else { "Kein Ergebnis vom COM-Job" }
    return @{ Ok = $false; Msg = $m }
}

Write-Log -Message "Office-Smoke-Test (Word / Excel / PowerPoint)..." -Color Gray

$smokeResults = [ordered]@{
    Word       = (Test-WordTrustCenter       -TempDir $smokeTempDir -TimeoutSec $Config.TrustCenterTimeoutSec)
    Excel      = (Test-ExcelTrustCenter      -TempDir $smokeTempDir -TimeoutSec $Config.TrustCenterTimeoutSec)
    PowerPoint = (Test-PowerPointTrustCenter -TempDir $smokeTempDir -TimeoutSec $Config.TrustCenterTimeoutSec)
}

$smokeFailures = @()
foreach ($entry in $smokeResults.GetEnumerator()) {
    if ($entry.Value.Ok) {
        Write-Log -Message ("  {0,-12} OK" -f $entry.Key) -Color DarkGreen
    } else {
        Write-Log -Message ("  {0,-12} FEHLER: {1}" -f $entry.Key, $entry.Value.Msg) -Level WARN -Color Yellow
        $smokeFailures += $entry.Key
    }
}

if ($smokeFailures.Count -gt 0) {
    Write-Log -Message "" -Color Yellow
    Write-Log -Message "TRUST-CENTER-TEST FEHLGESCHLAGEN: $($smokeFailures -join ', ')" -Level WARN -Color Yellow
    Write-Log -Message "  Moegliche Ursachen:" -Color Gray
    Write-Log -Message "    - Quellordner nicht als vertrauenswuerdiger Speicherort eingetragen" -Color Gray
    Write-Log -Message "    - 'Geschuetzte Ansicht' fuer Netzwerk/Internet aktiv" -Color Gray
    Write-Log -Message "    - Office-Profil beschaedigt oder Erstinitialisierung steht aus" -Color Gray
    Write-Log -Message ("  Risiko: jede betroffene Datei laeuft in den {0}s-Watchdog-Timeout." -f $Config.TimeoutSeconds) -Color Gray
    Write-Log -Message "" -Color Yellow

    $smokeMsg     = "Der Office-Smoke-Test ist fuer folgende Office-Anwendung(en) fehlgeschlagen:`n`n" +
                    ($smokeFailures -join ", ") +
                    "`n`nDetails siehe Log. Risiko: betroffene Dateien laufen in den $($Config.TimeoutSeconds)s-Timeout.`n`nTrotzdem fortfahren?"
    $smokeResult = Confirm-YesNo -Title "Trust-Center-Test fehlgeschlagen" -Message $smokeMsg

    if (-not $smokeResult) {
        Write-Log -Message "Abbruch durch Benutzer nach Trust-Center-Warnung." -Color Yellow
        Stop-Script
    }
    Write-Log -Message "Trust-Center-Warnung vom Benutzer ignoriert - Fortsetzung." -Level WARN -Color Yellow
} else {
    Write-Log -Message "Office-Smoke-Test OK fuer alle drei Office-Anwendungen." -Color DarkGreen
}

# ---------------------------------------------------------------------
# 9. Office-COM-Instanzen starten
# ---------------------------------------------------------------------

Write-Log -Message "Office-Instanzen werden gestartet..." -Color Gray

$wordInfo  = New-TrackedOfficeApp -ProgId "Word.Application"       -ProcessName "WINWORD"
$excelInfo = New-TrackedOfficeApp -ProgId "Excel.Application"      -ProcessName "EXCEL"
$pptInfo   = New-TrackedOfficeApp -ProgId "PowerPoint.Application" -ProcessName "POWERPNT"

$word  = $wordInfo.App;  $wordPid  = $wordInfo.Pid;  $wordStart  = $wordInfo.Start
$excel = $excelInfo.App; $excelPid = $excelInfo.Pid; $excelStart = $excelInfo.Start
$pptn  = $pptInfo.App;   $pptPid   = $pptInfo.Pid;   $pptStart   = $pptInfo.Start

Set-OfficeAppSafe -WordApp $word -ExcelApp $excel -PptApp $pptn

# Ohne eigene Prozess-ID greift der Watchdog nicht (er prueft TargetPid
# -gt 0). Das passierte bisher stillschweigend - eine haengende Datei
# blockierte den Lauf dann unbegrenzt, ohne dass die Ursache erkennbar war.
foreach ($chk in @(@{N='Word'; P=$wordPid}, @{N='Excel'; P=$excelPid}, @{N='PowerPoint'; P=$pptPid})) {
    if (-not $chk.P -or $chk.P -le 0) {
        Write-Log -Message ("WARNUNG: Fuer {0} konnte kein eigener Prozess ermittelt werden." -f $chk.N) -Level WARN -Color Yellow
        Write-Log -Message ("  Der {0}s-Timeout-Watchdog ist fuer {1} deaktiviert - haengende Dateien blockieren den Lauf." -f $Config.TimeoutSeconds, $chk.N) -Level WARN -Color Yellow
        Write-Log -Message  "  Ursache ist meist eine bereits laufende Office-Sitzung." -Level WARN -Color Yellow
    }
}

try { Write-Log -Message "Office-Version Word:       $($word.Version)"  -Color DarkGray } catch {}
try { Write-Log -Message "Office-Version Excel:      $($excel.Version)" -Color DarkGray } catch {}
try { Write-Log -Message "Office-Version PowerPoint: $($pptn.Version)"  -Color DarkGray } catch {}

$results      = [System.Collections.Generic.List[object]]::new()
$allFiles     = [System.Collections.Generic.List[string]]::new()
$visitedPaths = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

Write-Log -Message "`nScan läuft in: $rootPath" -Color Yellow
Write-Log -Message "Timeout pro Datei: $($Config.TimeoutSeconds)s" -Color Gray

# ---------------------------------------------------------------------
# 10. Watchdog-Runspace
# ---------------------------------------------------------------------

$watchdog = [hashtable]::Synchronized(@{
    TargetPid   = 0
    TargetStart = [DateTime]::MinValue
    Deadline    = [DateTime]::MaxValue
    TimedOut    = $false
    Running     = $true
})

$watchdogRunspace = [runspacefactory]::CreateRunspace()
$watchdogRunspace.Open()

$watchdogPS = [powershell]::Create()
$watchdogPS.Runspace = $watchdogRunspace
$watchdogPS.AddScript({
    param($state, $pollMs)
    while ($state.Running) {
        if ($state.TargetPid -gt 0 -and [DateTime]::UtcNow -gt $state.Deadline) {
            $state.TimedOut = $true
            try {
                $p = Get-Process -Id $state.TargetPid -ErrorAction Stop
                if ($p.Name -match '^(WINWORD|EXCEL|POWERPNT)$' -and
                    ($state.TargetStart -eq [DateTime]::MinValue -or $p.StartTime -eq $state.TargetStart)) {
                    Stop-Process -Id $state.TargetPid -Force
                }
            } catch {}
            $state.TargetPid = 0
        }
        Start-Sleep -Milliseconds $pollMs
    }
}).AddArgument($watchdog).AddArgument($Config.WatchdogPollMs) | Out-Null

$watchdogHandle = $watchdogPS.BeginInvoke()

# ---------------------------------------------------------------------
# 11. Dateisuche
# ---------------------------------------------------------------------

$scanStart = Get-Date

try {
    Write-Log -Message "Dateien werden gesucht..." -Color Gray

    # Bei Option 1/3 (SkipPreScan): Live-Feedback waehrend des Sammelns,
    # damit kein blindes Warten entsteht. Bei Option 2: stille Suche -
    # ETA-Anzeige folgt nach Abschluss.
    $progressCallback = if ($script:SkipPreScan) {
        { param($cnt) Write-Host "  Bisher gefunden: $cnt Office-Dateien (laufend)..." -ForegroundColor DarkGray }
    } else { $null }

    try {
        Get-FilesIterative -RootPath $longPath -Extensions $officeExtensions `
            -FileList $allFiles -ErrorList $results -VisitedPaths $visitedPaths `
            -OnProgress $progressCallback
    }
    catch {
        Write-Log -Message "Long-Path-Enumeration fehlgeschlagen: $($_.Exception.Message)" -Level WARN -Color Yellow
        Write-Log -Message "Wiederhole Suche ohne Long-Path-Präfix. HINWEIS: Bisher gesammelte Enumerations-Fehler gehen verloren; es gilt nun das 260-Zeichen-Limit." -Level WARN -Color Yellow
        $allFiles.Clear(); $results.Clear(); $visitedPaths.Clear()
        # Der zweite Versuch war bisher ungeschuetzt. Scheiterte auch er
        # (Laufwerk offline, Freigabe weg, Pfad geloescht), flog die
        # Ausnahme aus dem aeusseren try - und das hat nur ein finally,
        # kein catch. Der Anwender bekam einen roten Stacktrace statt
        # einer Meldung, und der Bericht wurde nie erzeugt.
        try {
            Get-FilesIterative -RootPath $resolvedPath -Extensions $officeExtensions `
                -FileList $allFiles -ErrorList $results -VisitedPaths $visitedPaths `
                -OnProgress $progressCallback
        }
        catch {
            Write-Log -Message "Auch die Suche ohne Long-Path-Präfix ist fehlgeschlagen: $($_.Exception.Message)" -Level ERROR -Color Red
            Write-Log -Message "Pfad nicht erreichbar: $resolvedPath" -Level ERROR -Color Red
            Write-Log -Message "  Moegliche Ursachen: Netzlaufwerk getrennt, Freigabe offline, Pfad geloescht," -Color Gray
            Write-Log -Message "  oder bei erhoehten Rechten ausgeblendetes gemapptes Laufwerk (dann UNC-Pfad nutzen)." -Color Gray
            $results.Add((New-ErrorRecord -Ordner "ENUMERATIONSFEHLER" -DateiPfad $resolvedPath `
                -Kategorie "Pfadfehler" -Details "Verzeichnis nicht durchsuchbar: $($_.Exception.Message)"))
            $allFiles.Clear()
        }
    }

    $totalFiles = $allFiles.Count
    Write-Log -Message "$totalFiles Office-Dateien gefunden. Prüfung startet..." -Color Cyan

    # Office-Apps, deren Restart endgueltig fehlgeschlagen ist: weitere
    # Dateien dieses Typs werden direkt als Problem markiert, ohne erneut
    # einen Restart zu probieren (sonst Endlos-Crash-Schleife bei kaputter
    # Office-Installation oder vollem RAM).
    $script:AppFailed = @{ "word" = $false; "excel" = $false; "ppt" = $false }
    # Get-OfficeAppType liefert 'Word'/'Excel'/'PowerPoint', gesetzt wird
    # AppFailed aber ueber $targetApp mit 'word'/'excel'/'ppt'. Fuer Word und
    # Excel faellt das nicht auf, weil ein @{}-Hashtable in PowerShell
    # case-insensitiv ist - 'PowerPoint' hat mit 'ppt' aber keine
    # Aehnlichkeit. Nachgestellt: ContainsKey('Word')=True,
    # ContainsKey('Excel')=True, ContainsKey('PowerPoint')=False, und
    # AppFailed['PowerPoint'] ist $null, also auch als Wert falsy - die
    # Bremse gegen die Endlos-Neustartschleife war fuer PowerPoint doppelt
    # wirkungslos. Diese Zuordnung haelt Setzen und Pruefen deckungsgleich.
    $script:AppTypeToKey = @{ "Word" = "word"; "Excel" = "excel"; "PowerPoint" = "ppt" }

    # -----------------------------------------------------------------
    # 12. Dateiprüfung
    # -----------------------------------------------------------------

    $counter = 0
    foreach ($filePath in $allFiles) {
        Test-AbortRequested
        if ($script:ShouldStop) { break }

        $counter++
        $fileName = [System.IO.Path]::GetFileName($filePath)
        $dirName  = Remove-LongPathPrefix -Path ([System.IO.Path]::GetDirectoryName($filePath))
        $ext      = [System.IO.Path]::GetExtension($filePath)
        $comPath  = Remove-LongPathPrefix -Path $filePath
        $appType  = Get-OfficeAppType -Extension $ext

        # Wenn der Restart der zustaendigen Office-App in einem frueheren
        # Iterationsschritt endgueltig fehlgeschlagen ist, gar nicht erst
        # versuchen - direkt als Problem markieren und Bericht-Lauf nicht
        # gefaehrden (sonst Endlos-Crash beim naechsten Restart-Versuch).
        $appKey = if ($appType) { $script:AppTypeToKey[$appType] } else { $null }
        if ($appKey -and $script:AppFailed.ContainsKey($appKey) -and $script:AppFailed[$appKey]) {
            $meta = Get-FileMetadata -Path $filePath
            $results.Add((New-ErrorRecord -Ordner $dirName -DateiPfad $comPath `
                -Kategorie "Office-App nicht verfügbar" `
                -Details "$appType-Restart in dieser Session bereits fehlgeschlagen - Datei nicht geprueft." `
                -Dateityp $appType -SizeMB $meta.SizeMB -Modified $meta.Modified))
            Write-Log -Message "Problem: $fileName [Office-App nicht verfügbar ($appType)]" -Level ERROR -Color Red
            continue
        }

        if ($counter % $Config.ProgressInterval -eq 0 -or $counter -eq $totalFiles) {
            if (-not $script:UseProgress) {
                # Option 1: laufende Zaehlung inline, keine Progress-Bar
                Write-Host ("  [{0,5} / {1}] {2}" -f $counter, $totalFiles, $fileName) -ForegroundColor DarkGray
            }
            elseif ($script:SkipPreScan) {
                # Option 3: Progress-Bar ohne ETA - Total ist bekannt (Pre-Scan
                # lief trotzdem im Hintergrund mit Live-Feedback), aber wir
                # zeigen keine ETA-Berechnung, weil "sofortiger Start"-Semantik
                # bewahrt werden soll.
                Write-Progress -Activity "Office-Dateien prüfen" `
                    -Status "[$counter / $totalFiles] $fileName" `
                    -PercentComplete -1
            }
            else {
                # Option 2: Progress-Bar mit ETA
                $elapsed    = (Get-Date) - $scanStart
                $avgPerFile = if ($counter -gt 0) { $elapsed.TotalSeconds / $counter } else { 0 }
                $remaining  = [math]::Max(0, $totalFiles - $counter)
                $etaSec     = $avgPerFile * $remaining
                $etaStr = if     ($etaSec -gt 3600) { "{0:N1} Std." -f ($etaSec / 3600) }
                          elseif ($etaSec -gt 60)   { "{0:N0} Min." -f ($etaSec / 60) }
                          else                      { "{0:N0} Sek." -f $etaSec }
                $pct = [math]::Round(($counter / $totalFiles) * 100, 0)
                Write-Progress -Activity "Office-Dateien prüfen" `
                    -Status "[$counter / $totalFiles] ETA: $etaStr | $fileName" `
                    -PercentComplete $pct
            }
        }

        $targetApp   = $null
        $targetPid   = 0
        $targetStart = [DateTime]::MinValue
        $extLower    = $ext.ToLowerInvariant()

        if     ($Config.WordExtensions  -contains $extLower) { $targetApp = "word";  $targetPid = $wordPid;  $targetStart = $wordStart }
        elseif ($Config.ExcelExtensions -contains $extLower) { $targetApp = "excel"; $targetPid = $excelPid; $targetStart = $excelStart }
        elseif ($Config.PptExtensions   -contains $extLower) { $targetApp = "ppt";   $targetPid = $pptPid;   $targetStart = $pptStart }

        if (-not $targetApp) { continue }

        $preflight = Test-FileAccessible -Path $filePath
        if (-not $preflight.OK) {
            $cat  = Get-ErrorCategory -Message $preflight.Error -HResult $preflight.HResult -AppType $targetApp
            $meta = Get-FileMetadata -Path $filePath
            $results.Add((New-ErrorRecord -Ordner $dirName -DateiPfad $comPath `
                -Kategorie $cat.Kategorie -Details "[Preflight] $($cat.Details)" `
                -Dateityp $appType -SizeMB $meta.SizeMB -Modified $meta.Modified))
            Write-Log -Message "Preflight-Fehler: $fileName [$($cat.Kategorie)]" -Level WARN -Color Red
            continue
        }

        # Verschluesselung VOR dem COM-Open an der Dateisignatur erkennen.
        # Aus der Office-Fehlermeldung laesst sie sich fuer Excel nicht
        # ableiten: dort kommt der generische HRESULT 0x800A03EC, der weiter
        # oben bereits als 'Office-Fehler' abgefangen wird - die Kategorie
        # 'Passwortschutz' war damit unerreichbar. Nebenbei entfaellt fuer
        # diese Dateien der komplette COM-Weg samt Watchdog-Timeout.
        if (Test-IstVerschluesselt -Path $comPath) {
            $meta = Get-FileMetadata -Path $filePath
            $results.Add((New-ErrorRecord -Ordner $dirName -DateiPfad $comPath `
                -Kategorie "Passwortschutz" `
                -Details "Datei ist verschluesselt (OOXML mit Oeffnungskennwort - CFB-Container statt ZIP). Nicht pruefbar." `
                -Dateityp $appType -SizeMB $meta.SizeMB -Modified $meta.Modified))
            Write-Log -Message "Problem: $fileName [Passwortschutz]" -Level WARN -Color Yellow
            continue
        }

        $attempt      = 0
        $maxAttempts  = 1 + $Config.RetryCount
        $finalProblem = $false
        $finalReason  = ""
        $finalHResult = 0

        while ($attempt -lt $maxAttempts) {
            $attempt++
            $watchdog.Deadline    = [DateTime]::UtcNow.AddSeconds($Config.TimeoutSeconds)
            $watchdog.TimedOut    = $false
            $watchdog.TargetStart = $targetStart
            $watchdog.TargetPid   = $targetPid

            $comObj     = $null
            $collection = $null
            $isProblem  = $false
            $reason     = ""
            $hresult    = 0

            try {
                switch ($targetApp) {
                    "word" {
                        # Parameter: FileName, ConfirmConversions, ReadOnly,
                        # AddToRecentFiles, PasswordDocument, PasswordTemplate,
                        # Revert, WritePasswordDocument, WritePasswordTemplate.
                        # Bisher wurde nur PasswordDocument gesetzt. Dateien mit
                        # SCHREIB-Schutzpasswort zeigen dann trotzdem einen
                        # modalen Dialog und liefen jedes Mal in den vollen
                        # 120s-Watchdog-Timeout, statt sofort als
                        # "Passwortschutz" erkannt zu werden.
                        $collection = $word.Documents
                        $comObj     = $collection.Open($comPath, $false, $true, $false,
                                                       $Config.DummyPassword, "", $false,
                                                       $Config.DummyPassword)
                        Close-ComObject -ComObj $comObj -AppType "word"
                    }
                    "excel" {
                        # Parameter: Filename, UpdateLinks, ReadOnly, Format,
                        # Password, WriteResPassword, IgnoreReadOnlyRecommended.
                        # WriteResPassword deckt den Schreibschutz ab,
                        # IgnoreReadOnlyRecommended unterdrueckt zusaetzlich den
                        # Dialog "Schreibgeschuetzt oeffnen?" - beides sind
                        # modale Blocker, die vorher in den Timeout liefen.
                        $collection = $excel.Workbooks
                        $comObj     = $collection.Open($comPath, 0, $true, [System.Type]::Missing,
                                                       $Config.DummyPassword, $Config.DummyPassword,
                                                       $true)
                        Close-ComObject -ComObj $comObj -AppType "excel"
                    }
                    "ppt" {
                        $collection = $pptn.Presentations
                        $comObj     = $collection.Open($comPath, [int]-1, [int]0, [int]0)
                        Close-ComObject -ComObj $comObj -AppType "ppt"
                    }
                }
            }
            catch {
                $isProblem = $true
                $reason    = $_.Exception.Message
                # Echten HRESULT auspacken - PowerShell verpackt .NET- und
                # COM-Ausnahmen in eine MethodInvocationException mit dem
                # konstanten Code 0x80131501 (siehe Get-EchtenHResult).
                $hresult   = Get-EchtenHResult -Exception $_.Exception
                if ($comPath.Length -gt 259) {
                    $reason += " [HINWEIS: Pfad $($comPath.Length) Zeichen — möglicherweise pfadlängenbedingt.]"
                }
            }
            finally {
                $watchdog.TargetPid = 0
                if ($null -ne $comObj) {
                    if ($isProblem) { try { Close-ComObject -ComObj $comObj -AppType $targetApp } catch {} }
                    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObj) | Out-Null } catch {}
                    $comObj = $null
                }
                if ($null -ne $collection) {
                    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($collection) | Out-Null } catch {}
                    $collection = $null
                }
            }

            $needsRestart = $false

            if ($watchdog.TimedOut -and $isProblem) {
                $needsRestart = $true
                if ($targetApp -eq "ppt") {
                    $reason = "TIMEOUT nach $($Config.TimeoutSeconds)s — PowerPoint blockiert. " +
                              "Mögliche Ursachen: Passwortschutz (PowerPoint.Open unterstützt keine Passwort-Injection!), " +
                              "SmartArt-Schriftarthänger oder beschädigte Einbettungen."
                } else {
                    $reason = "TIMEOUT nach $($Config.TimeoutSeconds)s — Datei hat Office blockiert (Passwortschutz, defekte Einbettung oder schwere Korruption)."
                }
            }
            elseif ($isProblem) {
                $appRef = switch ($targetApp) { "word" { $word } "excel" { $excel } "ppt" { $pptn } }
                if (-not (Test-OfficeAppAlive -App $appRef)) {
                    $needsRestart = $true
                    $reason += " [OFFICE-ABSTURZ: $targetApp-Instanz wird neu gestartet.]"
                }
            }

            if ($needsRestart) {
                Write-Log -Message "  Office-Neustart nötig ($targetApp): $fileName" -Level WARN -Color Magenta
                if ($targetPid -gt 0) {
                    try {
                        $p = Get-Process -Id $targetPid -ErrorAction Stop
                        if ($p.Name -match '^(WINWORD|EXCEL|POWERPNT)$' -and
                            ($targetStart -eq [DateTime]::MinValue -or $p.StartTime -eq $targetStart)) {
                            Stop-Process -Id $targetPid -Force -ErrorAction SilentlyContinue
                        }
                    } catch {}
                }
                # New-TrackedOfficeApp ruft intern New-Object -ComObject auf,
                # was bei blockiertem COM-Server, voller RAM, kaputter Office-
                # Lizenz oder zerstoerter Registrierung eine COMException werfen
                # kann. Ohne try/catch hier wuerde die Exception bis zum
                # aeusseren try propagieren - aber das hat KEIN catch (nur
                # finally), also wird der Excel-Bericht nicht mehr erstellt
                # und Stunden an gesammelten Scandaten gehen verloren.
                try {
                    switch ($targetApp) {
                        "word" {
                            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
                            $wordInfo = New-TrackedOfficeApp -ProgId "Word.Application" -ProcessName "WINWORD"
                            $word = $wordInfo.App; $wordPid = $wordInfo.Pid; $wordStart = $wordInfo.Start; $targetPid = $wordPid; $targetStart = $wordStart
                        }
                        "excel" {
                            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
                            $excelInfo = New-TrackedOfficeApp -ProgId "Excel.Application" -ProcessName "EXCEL"
                            $excel = $excelInfo.App; $excelPid = $excelInfo.Pid; $excelStart = $excelInfo.Start; $targetPid = $excelPid; $targetStart = $excelStart
                        }
                        "ppt" {
                            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pptn) | Out-Null } catch {}
                            $pptInfo = New-TrackedOfficeApp -ProgId "PowerPoint.Application" -ProcessName "POWERPNT"
                            $pptn = $pptInfo.App; $pptPid = $pptInfo.Pid; $pptStart = $pptInfo.Start; $targetPid = $pptPid; $targetStart = $pptStart
                        }
                    }
                    Set-OfficeAppSafe -WordApp $word -ExcelApp $excel -PptApp $pptn
                } catch {
                    Write-Log -Message "  KRITISCH: Office-Neustart ($targetApp) fehlgeschlagen: $($_.Exception.Message)" -Level ERROR -Color Red
                    Write-Log -Message "  Weitere $targetApp-Dateien werden bis Skriptende uebersprungen. Bisheriger Fortschritt bleibt erhalten." -Level WARN -Color Yellow
                    $script:AppFailed[$targetApp] = $true
                    $finalProblem = $true
                    $finalReason  = "Office-Restart fehlgeschlagen nach Defekt durch $fileName : $($_.Exception.Message)"
                    $finalHResult = -1
                    break
                }
            }

            if (-not $isProblem) {
                $finalProblem = $false
                break
            }

            $finalProblem = $true
            $finalReason  = $reason
            $finalHResult = $hresult

            # Retry auch nach Office-Neustart erlaubt: $maxAttempts deckelt
            # die Anzahl Versuche, $needsRestart blockiert Retry nicht.
            # Bei echtem Datei-Defekt bricht der zweite Versuch wieder ab.
            if ($attempt -lt $maxAttempts -and (Test-IsTransientError -HResult $hresult -Message $reason)) {
                Write-Log -Message "  Transienter Fehler bei $fileName — Retry in $($Config.RetryDelayMs / 1000)s" -Color DarkYellow
                Start-Sleep -Milliseconds $Config.RetryDelayMs
                continue
            }

            break
        }

        if ($finalProblem) {
            $cat  = Get-ErrorCategory -Message $finalReason -HResult $finalHResult -AppType $targetApp
            $meta = Get-FileMetadata -Path $filePath
            $results.Add((New-ErrorRecord -Ordner $dirName -DateiPfad $comPath `
                -Kategorie $cat.Kategorie -Details $cat.Details `
                -Dateityp $appType -SizeMB $meta.SizeMB -Modified $meta.Modified))
            if ($cat.Kategorie -eq "Namenskonflikt") {
                Write-Log -Message "Problem: $fileName [Namenskonflikt - modaler Excel-Dialog ueber Built-In-Namen, manuelles Wegklicken erforderlich]" -Level ERROR -Color Red
            } else {
                Write-Log -Message "Problem: $fileName [$($cat.Kategorie)]" -Level ERROR -Color Red
            }
        }

        if ($counter % $Config.CsvCheckpointEvery -eq 0 -and $results.Count -gt 0) {
            Save-Checkpoint -Results $results -Path $checkpointPath
        }
        if ($counter % $Config.GcIntervalFiles -eq 0) {
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }

    Write-Progress -Activity "Office-Dateien prüfen" -Completed

    # -----------------------------------------------------------------
    # 13. Excel-Bericht erstellen
    # -----------------------------------------------------------------

    $scanEnd      = Get-Date
    $scanDuration = $scanEnd - $scanStart
    $durationStr  = "{0:N0} Min. {1:N0} Sek." -f $scanDuration.TotalMinutes, $scanDuration.Seconds

    if ($results.Count -gt 0) {

        # CSV-Fallback als zentraler Scriptblock: wird sowohl bei toter
        # Excel-Instanz (Pre-Check) als auch bei einem Crash mitten im
        # Excel-Bericht-Aufbau (z.B. unsichtbares Lizenz-Popup, COM-
        # Routing-Fehler, Workbooks.Add wirft) verwendet. Ohne diesen
        # zentralen Fallback gingen bei einem COM-Ausfall waehrend des
        # Bericht-Aufbaus alle in stundenlanger Arbeit gesammelten
        # Scan-Daten verloren (Checkpoint ist evtl. veraltet).
        $csvPath = [System.IO.Path]::ChangeExtension($reportPath, ".csv")
        $writeCsvFallback = {
            param([string]$Reason)
            try {
                $results | Export-Csv -Path $csvPath -Delimiter ";" -Encoding UTF8 -NoTypeInformation
                Write-Log -Message "CSV-Bericht gespeichert: $csvPath" -Color Green
                if (Test-Path -LiteralPath $checkpointPath) {
                    try { Remove-Item -LiteralPath $checkpointPath -Force -ErrorAction SilentlyContinue } catch {}
                }
            } catch {
                Write-Log -Message "CSV-Fallback fehlgeschlagen ($Reason): $($_.Exception.Message)" -Level ERROR -Color Red
                Write-Log -Message "Checkpoint bleibt erhalten: $checkpointPath" -Level WARN -Color Yellow
            }
        }

        # Merker fuer den finally-Block: nur wenn der Bericht am Ende
        # sichtbar in Excel steht, darf Excel offen bleiben. Vorher wurde
        # allein an $results.Count entschieden - nach jedem CSV-Fallback
        # blieb dadurch eine UNSICHTBARE EXCEL.EXE als Zombie zurueck.
        $script:ReportShownInExcel = $false

        if (-not (Test-OfficeAppAlive -App $excel)) {
            Write-Log -Message "`nExcel-Instanz nicht mehr verfügbar — Bericht wird als CSV gespeichert." -Level WARN -Color Yellow
            & $writeCsvFallback "Excel nicht verfügbar"
        }
        else {
            $wbReport = $null
            try {
                $wbReport = $excel.Workbooks.Add()
                $ws       = $wbReport.Worksheets.Item(1)
                $ws.Name  = "Fehlerhafte Dateien"

                $ws.Cells.Item(1, 1) = "Startpfad:";  $ws.Cells.Item(1, 2) = $resolvedPath
                $ws.Cells.Item(2, 1) = "Startzeit:";  $ws.Cells.Item(2, 2) = $scanStart.ToString("yyyy-MM-dd HH:mm:ss")
                $ws.Cells.Item(3, 1) = "Endzeit:";    $ws.Cells.Item(3, 2) = $scanEnd.ToString("yyyy-MM-dd HH:mm:ss")
                $ws.Cells.Item(4, 1) = "Dauer:";      $ws.Cells.Item(4, 2) = $durationStr

                # Ehrliche Auszaehlung: bei Abbruch durch Ctrl+C wurden nur
                # $counter Dateien tatsaechlich geprueft, aber $totalFiles
                # gefunden. Stur $totalFiles als "Gescannt" einzutragen wuerde
                # 99.995 ungeprueften Dateien faelschlich Sauberkeit attestieren.
                if ($script:ShouldStop -and $counter -lt $totalFiles) {
                    $ws.Cells.Item(5, 1) = "Status:";   $ws.Cells.Item(5, 2) = "ABGEBROCHEN durch Benutzer (Ctrl+C)"
                    $ws.Cells.Item(6, 1) = "Gefunden:"; $ws.Cells.Item(6, 2) = $totalFiles
                    $ws.Cells.Item(7, 1) = "Geprueft:"; $ws.Cells.Item(7, 2) = $counter
                    $ws.Cells.Item(8, 1) = "Probleme:"; $ws.Cells.Item(8, 2) = $results.Count
                    $headerOffset = 2
                    try { $ws.Range("B5").Font.Color = 255 } catch {}
                    try { $ws.Range("A1:A8").Font.Bold = $true } catch {}
                } else {
                    $ws.Cells.Item(5, 1) = "Gescannt:"; $ws.Cells.Item(5, 2) = $totalFiles
                    $ws.Cells.Item(6, 1) = "Probleme:"; $ws.Cells.Item(6, 2) = $results.Count
                    $headerOffset = 0
                    try { $ws.Range("A1:A6").Font.Bold = $true } catch {}
                }

                $headerRow = 8 + $headerOffset
                $ws.Cells.Item($headerRow, 1) = "Ordner (Link)"
                $ws.Cells.Item($headerRow, 2) = "Vollständiger Pfad (Link)"
                $ws.Cells.Item($headerRow, 3) = "Kategorie"
                $ws.Cells.Item($headerRow, 4) = "Details"
                $ws.Cells.Item($headerRow, 5) = "Dateityp"
                $ws.Cells.Item($headerRow, 6) = "Größe (MB)"
                $ws.Cells.Item($headerRow, 7) = "Letzte Änderung"
                try {
                    $ws.Rows.Item($headerRow).Font.Bold           = $true
                    $ws.Rows.Item($headerRow).Interior.ColorIndex = 15
                } catch {}

                $metaLabels = @("ZUGRIFF VERWEIGERT","ENUMERATIONSFEHLER","ZIRKELBEZUG")
                $row = $headerRow + 1

                # Performance-Optimierung: Spalten 3-7 (Kategorie, Details, Dateityp,
                # Groesse, Datum) werden in ein 2D-Array gesammelt und nach dem Loop
                # in EINEM COM-Aufruf an die Range geschrieben. Cell-by-Cell ueber
                # COM ist bei 5000+ Items qualvoll langsam (35.000 Marshaling-
                # Roundtrips). Hyperlinks (Spalten 1-2) muessen leider einzeln
                # bleiben - Hyperlinks.Add hat keine Bulk-API.
                $rowCount = $results.Count
                $dataArray = if ($rowCount -gt 0) {
                    New-Object 'object[,]' $rowCount, 5
                } else { $null }
                $arrIdx = 0

                foreach ($item in $results) {
                    $safeOrdner = "$($item.Ordner)"
                    $safeDatei  = "$($item.Datei_Pfad)"
                    if ($safeOrdner.Length -gt 32000) { $safeOrdner = $safeOrdner.Substring(0, 32000) + " [GEKÜRZT]" }
                    if ($safeDatei.Length  -gt 32000) { $safeDatei  = $safeDatei.Substring(0, 32000)  + " [GEKÜRZT]" }

                    try {
                        if ($metaLabels -notcontains $safeOrdner) {
                            if ($safeOrdner.Length -le $Config.MaxHyperlinkLength) {
                                $ws.Hyperlinks.Add($ws.Cells.Item($row, 1), $safeOrdner, "", "Ordner öffnen", $safeOrdner) | Out-Null
                            } else {
                                $ws.Cells.Item($row, 1) = $safeOrdner
                            }
                            if ($safeDatei.Length -le $Config.MaxHyperlinkLength) {
                                $ws.Hyperlinks.Add($ws.Cells.Item($row, 2), $safeDatei, "", "Datei öffnen", $safeDatei) | Out-Null
                            } else {
                                $ws.Cells.Item($row, 2) = $safeDatei
                            }
                        }
                        else {
                            $ws.Cells.Item($row, 1) = $safeOrdner
                            $ws.Cells.Item($row, 2) = $safeDatei
                        }
                    } catch {
                        try { $ws.Cells.Item($row, 1) = $safeOrdner } catch {}
                        try { $ws.Cells.Item($row, 2) = $safeDatei }  catch {}
                    }

                    $safeDetails = "$($item.Details)"
                    if ($safeDetails.Length -gt 32000) { $safeDetails = $safeDetails.Substring(0, 32000) + " [TEXT GEKÜRZT]" }

                    # 2D-Array befuellen (Spalten 3-7 entsprechen Array-Index 0-4).
                    # Apostroph-Praefix zwingt Excel zu Text-Interpretation, damit
                    # lange Zahlen-/Datums-Strings nicht in Wissenschaftsnotation
                    # umgewandelt werden.
                    $dataArray[$arrIdx, 0] = "'" + "$($item.Kategorie)"
                    $dataArray[$arrIdx, 1] = "'" + $safeDetails
                    $dataArray[$arrIdx, 2] = "'" + "$($item.Dateityp)"
                    $dataArray[$arrIdx, 3] = if ($null -ne $item.Dateigroesse_MB -and "$($item.Dateigroesse_MB)" -ne "") {
                        [double]$item.Dateigroesse_MB
                    } else { $null }
                    $dataArray[$arrIdx, 4] = "'" + "$($item.Letzte_Aenderung)"
                    $arrIdx++

                    $row++
                }

                # Bulk-Schreibung der Datenzellen 3-7. Bei Fehler (sehr selten -
                # z.B. wenn Excel-COM den Array-Typ nicht akzeptiert) Fallback
                # zellenweise, damit der Bericht nicht ganz verloren geht.
                if ($rowCount -gt 0 -and $null -ne $dataArray) {
                    $bulkStart = $headerRow + 1
                    $bulkEnd   = $bulkStart + $rowCount - 1
                    $bulkOk    = $false
                    try {
                        $bulkRange = $ws.Range($ws.Cells.Item($bulkStart, 3), $ws.Cells.Item($bulkEnd, 7))
                        $bulkRange.Value2 = $dataArray
                        $bulkOk = $true
                    } catch {
                        Write-Log -Message "Bulk-Schreibung fehlgeschlagen, Fallback zellenweise: $($_.Exception.Message)" -Level WARN -Color Yellow
                    }
                    if (-not $bulkOk) {
                        for ($i = 0; $i -lt $rowCount; $i++) {
                            $r = $bulkStart + $i
                            try { $ws.Cells.Item($r, 3) = $dataArray[$i, 0] } catch {}
                            try { $ws.Cells.Item($r, 4) = $dataArray[$i, 1] } catch {}
                            try { $ws.Cells.Item($r, 5) = $dataArray[$i, 2] } catch {}
                            try { $ws.Cells.Item($r, 6) = $dataArray[$i, 3] } catch {}
                            try { $ws.Cells.Item($r, 7) = $dataArray[$i, 4] } catch {}
                        }
                    }
                }

                try {
                    $lastRow = $row - 1
                    if ($lastRow -gt $headerRow) {
                        $filterRange = $ws.Range($ws.Cells.Item($headerRow, 1), $ws.Cells.Item($lastRow, 7))
                        $filterRange.AutoFilter() | Out-Null
                    }
                } catch {}

                try { $ws.Columns.AutoFit() | Out-Null } catch {}

                try {
                    $wbReport.SaveAs($reportPath, 51)
                    Write-Log -Message "`nBericht gespeichert: $reportPath" -Color Green
                    if (Test-Path -LiteralPath $checkpointPath) {
                        try { Remove-Item -LiteralPath $checkpointPath -Force -ErrorAction SilentlyContinue } catch {}
                    }
                } catch {
                    Write-Log -Message "`nWarnung: Bericht konnte nicht gespeichert werden: $($_.Exception.Message)" -Level WARN -Color Yellow
                    Write-Log -Message "Der Bericht ist weiterhin in Excel geöffnet." -Level WARN -Color Yellow
                }

                $excel.Visible = $true
                $script:ReportShownInExcel = $true
            }
            catch {
                # Excel-COM ist mitten im Bericht-Aufbau abgestuerzt
                # (z.B. unsichtbares Lizenz-Popup, COM-Routing-Fehler,
                # Workbooks.Add wirft, gen_py-Cache-Korruption). Ohne
                # diesen Notausgang gingen alle gesammelten Scan-Daten
                # verloren. Stub-Workbook still schliessen (best effort,
                # COM ist eh kaputt), dann CSV-Fallback erzwingen.
                $errMsg = $_.Exception.Message
                Write-Log -Message "`nExcel-Bericht-Erstellung fehlgeschlagen: $errMsg" -Level ERROR -Color Red
                Write-Log -Message "Wechsle auf CSV-Notfall-Export, um die gesammelten Daten zu retten." -Level WARN -Color Yellow

                if ($null -ne $wbReport) {
                    try { $wbReport.Close($false) } catch {}
                }

                & $writeCsvFallback "Excel-COM-Ausfall im Bericht: $errMsg"
            }
        }
        Write-Log -Message "Scan beendet. $($results.Count) Probleme in $durationStr gefunden." -Color Green
    }
    elseif ($script:ShouldStop -and $counter -lt $totalFiles) {
        Write-Log -Message "`nABGEBROCHEN durch Benutzer (Ctrl+C). $counter von $totalFiles Dateien geprueft, keine Probleme darunter." -Level WARN -Color Yellow
    }
    else {
        Write-Log -Message "`nSauber! Keine Probleme in $durationStr gefunden." -Color Cyan
    }
}
finally {
    # -----------------------------------------------------------------
    # 14. Aufräumen
    # -----------------------------------------------------------------

    $watchdog.Running = $false
    $waited = 0
    while ($null -ne $watchdogHandle -and -not $watchdogHandle.IsCompleted -and $waited -lt $Config.WatchdogShutdownMs) {
        Start-Sleep -Milliseconds 50
        $waited += 50
    }
    if ($null -ne $watchdogHandle -and $watchdogHandle.IsCompleted) {
        try { $watchdogPS.EndInvoke($watchdogHandle) } catch {}
    } else {
        try { $watchdogPS.Stop() } catch {}
    }
    try { $watchdogPS.Dispose() } catch {}
    try { $watchdogRunspace.Close() } catch {}
    try { $watchdogRunspace.Dispose() } catch {}

    # Quit() nur fuer selbst gestartete Instanzen. Konnte New-TrackedOfficeApp
    # keine neue Prozess-ID ermitteln, hat sich COM an eine bereits laufende
    # Sitzung des Anwenders gehaengt - ein Quit() darauf schliesst dessen
    # Fenster und verwirft ungespeicherte Dokumente. In dem Fall wird die
    # Referenz nur freigegeben; die Sitzung bleibt bestehen.
    if ($wordPid -gt 0) {
        try { $word.Quit() } catch {}
    } else {
        Write-Log -Message "Word wurde nicht beendet (bestehende Sitzung des Anwenders)." -Color DarkGray
        try { $word.Visible = $true } catch {}
    }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}

    if ($pptPid -gt 0) {
        try { $pptn.Quit() } catch {}
    } else {
        Write-Log -Message "PowerPoint wurde nicht beendet (bestehende Sitzung des Anwenders)." -Color DarkGray
    }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pptn) | Out-Null } catch {}

    # Excel bleibt nur dann offen, wenn der Bericht dort auch wirklich
    # angezeigt wird. Frueher entschied allein $results.Count - nach einem
    # CSV-Fallback (Excel-COM abgestuerzt, Lizenz-Popup, Workbooks.Add
    # gescheitert) blieb eine unsichtbare EXCEL.EXE dauerhaft im Speicher.
    if (-not $script:ReportShownInExcel) {
        if ($excelPid -gt 0) {
            try { $excel.Quit() } catch {}
        } else {
            Write-Log -Message "Excel wurde nicht beendet (bestehende Sitzung des Anwenders)." -Color DarkGray
            try { $excel.Visible = $true } catch {}
        }
    }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Stop-Script
