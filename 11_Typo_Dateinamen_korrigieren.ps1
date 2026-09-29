# ============================================================
# Typo-Scan mit Hunspell + interaktive Korrektur
# PowerShell 5.1 / .NET Framework
# Ablage: C:\tmp_skripte\
#
# Optionale Dateien in C:\tmp_skripte\hunspell\ (werden wenn vorhanden genutzt):
#
# ignore.txt – Firmen-/Kundennamen, die als korrekt gelten sollen
# skip_dirs.txt – Ordnernamen-Muster zum Überspringen (Wildcards via -like, z. B. .git, node_modules, *_Archiv)
# skip_extensions.txt – Endungen zum Überspringen (z. B. .tmp, .lock)
# hashes.txt – optional SHA256-Hashes im Format <HASH>  <Dateiname> für Integritätsprüfung der DLLs/Wörterbücher
#
# Parameter für Scheduled-Task/Silent-Mode:
# .\11_Typo_Dateinamen_korrigieren.ps1 -Mode Scan -RootPath "Q:\" -NoProgress
# .\11_Typo_Dateinamen_korrigieren.ps1 -Mode ApplyFromCsv -ApplyCsv "C:\Pfad\Scan.csv"
#
# Wichtig für die PS2EXE-Umwandlung: Nicht mit -noConsole kompilieren, 
# da Modi 2 und die Abschlussmeldung Console brauchen. Ansonsten 
# funktioniert das Skript out-of-the-box als .exe, alle Pfade 
# sind absolut, keine $PSScriptRoot-Abhängigkeiten.
# Invoke-ps2exe -inputFile "11_Typo_Dateinamen_korrigieren.ps1" \
# -outputFile "11_Typo_Dateinamen_korrigieren.exe" \
# -iconFile "powershell_icon.ico"
#
# Stand: 11.06.2026
# ============================================================

[CmdletBinding()]
param(
    [ValidateSet('Scan','Interactive','ApplyFromCsv')]
    [string]$Mode,
    [string]$RootPath,
    [string]$ApplyCsv,
    [switch]$NoProgress
)

$ErrorActionPreference = 'Stop'

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
try { $OutputEncoding            = [System.Text.Encoding]::UTF8 } catch {}


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

# ---------- Konfiguration ----------
# Ablage der Woerterbuecher und Filterlisten. 'C:\tmp_skripte' war fest
# verdrahtet - auf einem Rechner ohne Schreibrecht auf C:\ scheiterte
# bereits das New-Item in Initialize-Hunspell und das Skript brach ab,
# bevor irgendetwas geprueft wurde.
function Get-WritableDir {
    param([string[]]$Candidates)
    foreach ($cand in $Candidates) {
        if ([string]::IsNullOrWhiteSpace($cand)) { continue }
        try {
            if (-not (Test-Path -LiteralPath $cand)) {
                New-Item -ItemType Directory -Path $cand -Force -ErrorAction Stop | Out-Null
            }
            $probe = Join-Path $cand (".writetest_{0}.tmp" -f ([Guid]::NewGuid().ToString('N')))
            [System.IO.File]::WriteAllText($probe, 'x')
            [System.IO.File]::Delete($probe)
            return $cand
        } catch { continue }
    }
    return $null
}

$HunspellDir = Get-WritableDir @(
    'C:\tmp_skripte\hunspell',
    (Join-Path $env:LOCALAPPDATA 'TypoFix\hunspell'),
    (Join-Path $env:TEMP 'TypoFix_hunspell')
)
if (-not $HunspellDir) {
    Write-Host ""
    Write-Host "Kein beschreibbares Verzeichnis fuer die Woerterbuecher gefunden." -ForegroundColor Red
    Write-Host "Geprueft: C:\tmp_skripte, LOCALAPPDATA, TEMP." -ForegroundColor Red
    Write-Host "Bitte eines dieser Verzeichnisse beschreibbar machen." -ForegroundColor Yellow
    exit 1
}
# Bezugsquellen. Die Woerterbuecher haengen an einem COMMIT, nicht am
# Zweig 'master': ein Commit-SHA ist unveraenderlich, der Zweig aendert
# sich jederzeit - und damit aenderte sich unbemerkt, was als Tippfehler
# gilt und welche Dateien dieses Skript umbenennt.
# Stand: Commit 31bc2a11 vom 09.12.2024 (letzte Aenderung an de/).
# Ueberschreibbar ueber pfade.json, damit fuer eine Aktualisierung nicht
# das Skript editiert werden muss.
$HunspellCommit = '31bc2a1104a1cd175f900902f994c76dea35c763'
$NupkgUrl       = 'https://www.nuget.org/api/v2/package/NHunspell/1.2.5554.16953'
$DicUrl         = "https://raw.githubusercontent.com/LibreOffice/dictionaries/$HunspellCommit/de/de_DE_frami.dic"
$AffUrl         = "https://raw.githubusercontent.com/LibreOffice/dictionaries/$HunspellCommit/de/de_DE_frami.aff"

# pfade.json hat Vorrang, wenn sie die Schluessel mitbringt.
try {
    $konfDatei = Join-Path $PSScriptRoot 'pfade.json'
    if (Test-Path -LiteralPath $konfDatei) {
        $konf = Get-Content -LiteralPath $konfDatei -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($konf.hunspell) {
            if ($konf.hunspell.nupkg) { $NupkgUrl = [string]$konf.hunspell.nupkg }
            if ($konf.hunspell.dic)   { $DicUrl   = [string]$konf.hunspell.dic }
            if ($konf.hunspell.aff)   { $AffUrl   = [string]$konf.hunspell.aff }
        }
    }
} catch {
    # Kaputte oder fehlende Konfiguration: die Vorgaben oben gelten.
}
$MaxPathLen    = 259

# Verzeichnis-Praesets fuer die Startauswahl. Leere Eintraege werden im
# Menue ausgeblendet, die Nummerierung bleibt lueckenlos.
$DirectoryPresets = @(
    'Q:\'
    'R:\'
    'G:\Geteilte Ablagen\'
    'G:\Meine Ablage\'
    '\\server\dfs\'
)

# Fest eingebaute Ausschlussliste. Bisher gab es NUR die optionale Datei
# 'skip_dirs.txt' - fehlte sie (der Regelfall bei einer frischen
# Installation), lief der Scan durch Versionsverwaltung und Systemordner.
# Im Interaktiv-Modus wurden dort Ordner und Dateien tatsaechlich
# umbenannt: eine Umbenennung unterhalb von '.git' macht das Repository
# unbrauchbar. Die Datei skip_dirs.txt ERGAENZT diese Liste jetzt nur noch.
$DefaultSkipDirs = @(
    '.git', '.svn', '.hg', '.bzr',
    '$RECYCLE.BIN', 'System Volume Information', 'RECYCLER',
    'node_modules', '__pycache__', '.venv', 'venv',
    '.vs', '.idea', 'AppData',
    '~snapshot', '.snapshot'
)

$ReservedNames = @(
    'CON','PRN','AUX','NUL',
    'COM1','COM2','COM3','COM4','COM5','COM6','COM7','COM8','COM9',
    'LPT1','LPT2','LPT3','LPT4','LPT5','LPT6','LPT7','LPT8','LPT9'
)

# Korpus-basierter Zeichendreher-Pass: Tokens werden aus allen Dateinamen
# des Scan-Baums gesammelt, ein seltenes Token (<= MaxRareFreq) gilt als
# Tippfehler-Kandidat, wenn ein Damerau-Distanz-1-Partner mit Haeufigkeit
# >= MinDominantFreq existiert. Transposition (zwei benachbarte Buchstaben
# vertauscht) wird mit hoher Konfidenz markiert.
$CorpusMinDominantFreq = 5    # haeufiges Wort muss mind. 5x vorkommen
$CorpusMaxRareFreq     = 1    # seltenes Wort darf max. 1x vorkommen
$CorpusMinTokenLength  = 4    # nur Tokens >= 4 Zeichen ins Vokabular

# ---------- Bootstrap Hunspell ----------
# Dateien, deren Integritaet zwingend geprueft wird. Die drei DLLs werden
# als NATIVER CODE in den Prozess geladen; die Woerterbuecher steuern, was
# als Tippfehler gilt und damit, welche Dateien umbenannt werden.
$IntegrityFiles = @(
    'NHunspell.dll', 'Hunspellx64.dll', 'Hunspellx86.dll',
    'de_DE_frami.dic', 'de_DE_frami.aff'
)

function Test-DownloadHashes {
    <#
        Integritaetspruefung nach dem Prinzip "beim ersten Mal festnageln"
        (Trust On First Use).

        Vorher war die Pruefung faktisch wirkungslos: sie stieg bei
        fehlender hashes.txt sofort wieder aus - und diese Datei ist im
        Auslieferungszustand nicht vorhanden, der Skriptkopf fuehrt sie
        ausdruecklich als optional. Damit lief der Regelfall voellig
        ungeprueft, obwohl anschliessend drei native DLLs per Add-Type in
        den Prozess geladen werden.

        Jetzt gilt:
          - Fehlt hashes.txt, werden die Summen der frisch geladenen
            Dateien EINMALIG erfasst und geschrieben. Bei interaktivem
            Lauf wird das angezeigt und bestaetigt.
          - Existiert sie, muss JEDE dort genannte Datei passen. Eine
            Abweichung bricht ab, statt zu laden.
          - Eine Datei, die in hashes.txt fehlt, aber auf der Platte
            liegt, wird ergaenzt und gemeldet.

        Damit faellt jede spaetere Veraenderung auf - der Fall, um den es
        geht (manipulierte Neu-Ausgabe, ausgetauschte Datei im Cache).
        Wer feste Sollwerte vorgeben will, legt hashes.txt einfach vorab
        an: '<SHA256>  <Dateiname>' je Zeile.
    #>
    param(
        [string]$dir,
        [switch]$NonInteractive
    )
    $hashFile = Join-Path $dir 'hashes.txt'

    $expected = @{}
    if (Test-Path -LiteralPath $hashFile) {
        Get-Content -LiteralPath $hashFile -Encoding UTF8 |
            Where-Object { $_ -and -not $_.StartsWith('#') } |
            ForEach-Object {
                $parts = $_ -split '\s+', 2
                if ($parts.Count -eq 2) { $expected[$parts[1].Trim()] = $parts[0].Trim().ToUpper() }
            }
    }

    $vorhanden = @{}
    foreach ($fname in $IntegrityFiles) {
        $full = Join-Path $dir $fname
        if (Test-Path -LiteralPath $full) {
            $vorhanden[$fname] = (Get-FileHash -LiteralPath $full -Algorithm SHA256).Hash.ToUpper()
        }
    }
    if ($vorhanden.Count -eq 0) {
        throw "Keine der zu pruefenden Dateien liegt in '$dir' - Bootstrap fehlgeschlagen."
    }

    # 1) Abweichungen sind ein harter Abbruch.
    $abweichungen = @()
    foreach ($fname in $vorhanden.Keys) {
        if ($expected.ContainsKey($fname) -and $expected[$fname] -ne $vorhanden[$fname]) {
            $abweichungen += ("  {0}`n     erwartet: {1}`n     gefunden: {2}" -f `
                              $fname, $expected[$fname], $vorhanden[$fname])
        }
    }
    if ($abweichungen.Count -gt 0) {
        throw ("SHA256-Abweichung - es wird NICHTS geladen:`n" + ($abweichungen -join "`n") +
               "`n`nWenn die Aenderung beabsichtigt ist (neue Version), '$hashFile' loeschen und neu starten.")
    }

    # 2) Fehlende Eintraege ergaenzen bzw. erstmalig festnageln.
    $neu = @($vorhanden.Keys | Where-Object { -not $expected.ContainsKey($_) })
    if ($neu.Count -eq 0) {
        Write-Host "  Integritaet geprueft: $($vorhanden.Count) Datei(en) unveraendert." -ForegroundColor DarkGray
        return
    }

    if (-not (Test-Path -LiteralPath $hashFile)) {
        Write-Host ""
        Write-Host "  Erstmalige Integritaets-Festlegung" -ForegroundColor Yellow
        Write-Host "  Die folgenden Dateien werden als vertrauenswuerdig festgeschrieben." -ForegroundColor Yellow
        Write-Host "  Ab dem naechsten Lauf bricht jede Abweichung den Start ab." -ForegroundColor Yellow
        Write-Host ""
    }
    foreach ($fname in ($neu | Sort-Object)) {
        Write-Host ("    {0}  {1}" -f $vorhanden[$fname], $fname) -ForegroundColor DarkGray
        $expected[$fname] = $vorhanden[$fname]
    }

    if (-not $NonInteractive -and -not (Test-Path -LiteralPath $hashFile)) {
        $ok = Read-Host "  Diese Staende festschreiben und fortfahren? [J/n]"
        if ($ok -and $ok.Trim() -notmatch '^[JjYy]') {
            throw "Abgebrochen - Integritaets-Festlegung vom Anwender verweigert."
        }
    }

    $zeilen = @(
        "# SHA256-Sollwerte fuer 11_Typo_Dateinamen_korrigieren.ps1",
        "# Format: <SHA256>  <Dateiname>. Zeilen mit '#' werden ignoriert.",
        "# Eine Abweichung bricht den Start ab. Bei bewusstem Versionswechsel",
        "# diese Datei loeschen; sie wird dann neu erzeugt."
    ) + @($expected.Keys | Sort-Object | ForEach-Object { "{0}  {1}" -f $expected[$_], $_ })
    Set-Content -LiteralPath $hashFile -Value $zeilen -Encoding UTF8
    Write-Host "  Sollwerte geschrieben: $hashFile" -ForegroundColor DarkGray
}

function Save-Woerterbuch {
    <#
        Laedt ein Hunspell-Woerterbuch ueber eine Zwischendatei '<Ziel>.part'
        und benennt erst nach bestandener Plausibilitaetspruefung um. Wirft
        bei jedem Fehler; die Zwischendatei wird in jedem Fall entfernt.

        Plausibel heisst (gemessen an de_DE_frami, Commit 31bc2a11):
          .dic  4.356.903 Bytes, erste Zeile = Wortzahl 258200, 258220 Zeilen,
                letztes Byte LF  -> Mindestgroesse 1 MB, Kopfzeile numerisch,
                mindestens 90 % der angekuendigten Eintraege (die Zahl ist
                laut Hunspell nur ein Richtwert), Ende auf LF.
          .aff  19.067 Bytes, 'SET ISO8859-1' in Zeile 1, letztes Byte LF
                -> Mindestgroesse 4 KB, eine 'SET '-Zeile, Ende auf LF.
        Ein abgebrochener Download endet praktisch immer mitten in einer
        Zeile; eine Fehlerseite (Proxy, HTML) hat weder Wortzahl noch SET.
    #>
    param([string]$Uri, [string]$Ziel, [ValidateSet('dic','aff')][string]$Art)
    $teil = $Ziel + '.part'
    try {
        if (Test-Path -LiteralPath $teil) { Remove-Item -LiteralPath $teil -Force }
        Invoke-WebRequest -Uri $Uri -OutFile $teil -UseBasicParsing
        $bytes = [System.IO.File]::ReadAllBytes($teil)
        $mindest = if ($Art -eq 'dic') { 1MB } else { 4KB }
        if ($bytes.Length -lt $mindest) {
            throw "Download von '$Uri' unvollstaendig: nur $($bytes.Length) Bytes (erwartet mindestens $mindest)."
        }
        if ($bytes[$bytes.Length - 1] -ne 10) {
            throw "Download von '$Uri' endet mitten in einer Zeile - vermutlich abgebrochen."
        }
        # ISO-8859-1 bildet jedes Byte 1:1 ab - fuer die Strukturpruefung genuegt das.
        $text   = [System.Text.Encoding]::GetEncoding(28591).GetString($bytes)
        $zeilen = $text.Split([char]10)
        if ($Art -eq 'dic') {
            $kopf = $zeilen[0].Trim()
            $anzahl = 0
            if (-not [int]::TryParse($kopf, [ref]$anzahl) -or $anzahl -le 0) {
                throw "Download von '$Uri' ist keine Hunspell-.dic (erste Zeile '$kopf' ist keine Wortzahl)."
            }
            # -2: Kopfzeile und das leere Element hinter dem letzten LF.
            if (($zeilen.Length - 2) -lt [int]($anzahl * 0.9)) {
                throw "Download von '$Uri' unvollstaendig: $($zeilen.Length - 2) Zeilen, angekuendigt $anzahl."
            }
        } else {
            if (-not ($zeilen | Where-Object { $_ -match '^SET\s+\S' } | Select-Object -First 1)) {
                throw "Download von '$Uri' ist keine Hunspell-.aff (keine SET-Zeile)."
            }
        }
        Move-Item -LiteralPath $teil -Destination $Ziel -Force
    } finally {
        if (Test-Path -LiteralPath $teil) { Remove-Item -LiteralPath $teil -Force -ErrorAction SilentlyContinue }
    }
}

function Initialize-Hunspell {
    if (-not (Test-Path -LiteralPath $HunspellDir)) {
        New-Item -ItemType Directory -Path $HunspellDir -Force | Out-Null
    }

    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
    try { [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials } catch { }

    $dllNH   = Join-Path $HunspellDir 'NHunspell.dll'
    $dll64   = Join-Path $HunspellDir 'Hunspellx64.dll'
    $dll32   = Join-Path $HunspellDir 'Hunspellx86.dll'
    $dicPath = Join-Path $HunspellDir 'de_DE_frami.dic'
    $affPath = Join-Path $HunspellDir 'de_DE_frami.aff'

    if (-not ((Test-Path $dllNH) -and (Test-Path $dll64) -and (Test-Path $dll32))) {
        Write-Host "Lade NHunspell von nuget.org ..." -ForegroundColor Cyan
        $nupkg = Join-Path $HunspellDir 'nhunspell.zip'
        Invoke-WebRequest -Uri $NupkgUrl -OutFile $nupkg -UseBasicParsing
        $extract = Join-Path $HunspellDir '_nupkg'
        if (Test-Path $extract) { Remove-Item $extract -Recurse -Force }
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($nupkg, $extract)
        Get-ChildItem -Path $extract -Recurse -Include 'NHunspell.dll','Hunspellx64.dll','Hunspellx86.dll' |
            ForEach-Object { Copy-Item $_.FullName -Destination $HunspellDir -Force }
        Remove-Item $extract -Recurse -Force
        Remove-Item $nupkg -Force
    }

    # Woerterbuecher nie direkt in die Zieldatei laden. Vorher schrieb
    # Invoke-WebRequest -OutFile unmittelbar nach de_DE_frami.dic/.aff: ein
    # abgebrochener Download hinterliess eine Teildatei, der naechste Lauf
    # sah sie per Test-Path als vorhanden an und Test-DownloadHashes schrieb
    # ihren Hash per TOFU als Sollwert fest (nachgestellt mit einem
    # Download, der nach der Haelfte abbricht). Jetzt: in eine Zwischendatei
    # laden, Groesse und Aufbau pruefen, erst dann umbenennen.
    if (-not (Test-Path $dicPath)) {
        Write-Host "Lade de_DE_frami.dic ..." -ForegroundColor Cyan
        Save-Woerterbuch -Uri $DicUrl -Ziel $dicPath -Art 'dic'
    }
    if (-not (Test-Path $affPath)) {
        Write-Host "Lade de_DE_frami.aff ..." -ForegroundColor Cyan
        Save-Woerterbuch -Uri $AffUrl -Ziel $affPath -Art 'aff'
    }

    # Zwingend VOR dem Add-Type: danach ist der native Code bereits im
    # Prozess und eine Pruefung waere wirkungslos.
    Test-DownloadHashes -dir $HunspellDir -NonInteractive:(-not $script:InteractiveLaunch)

    try { Add-Type -Path $dllNH } catch { throw "NHunspell.dll konnte nicht geladen werden: $($_.Exception.Message)" }
    try { [NHunspell.Hunspell]::NativeDllPath = $HunspellDir } catch {
        throw "NativeDllPath konnte nicht gesetzt werden: $($_.Exception.Message)"
    }
    try {
        $script:Hunspell = New-Object NHunspell.Hunspell($affPath, $dicPath)
    } catch {
        throw "Hunspell-Initialisierung fehlgeschlagen (pruefe DLL-Architektur und Woerterbuecher): $($_.Exception.Message)"
    }
    if (-not $script:Hunspell) { throw "Hunspell-Instanz ist null." }

    $ignorePath = Join-Path $HunspellDir 'ignore.txt'
    if (Test-Path $ignorePath) {
        Get-Content $ignorePath -Encoding UTF8 | Where-Object { $_ -and -not $_.StartsWith('#') } |
            ForEach-Object {
                $w = $_.Trim()
                if ($w) { $script:Hunspell.Add($w) }
            }
    }
}

# ---------- Filter-Listen ----------
function Import-FilterList {
    param([string]$path)
    if (-not (Test-Path $path)) { return @() }
    @(Get-Content $path -Encoding UTF8 |
        Where-Object { $_ -and -not $_.StartsWith('#') } |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ })
}

function Test-SkipDir {
    param([string]$dirName)
    foreach ($p in $script:SkipDirs) {
        if ($dirName -like $p) { return $true }
    }
    return $false
}

function Test-SkipExt {
    param([string]$ext)
    if (-not $ext) { return $false }
    return ($script:SkipExts -contains $ext.ToLower())
}

# ---------- Damerau-Levenshtein ----------
function Get-EditDistance {
    param([string]$a, [string]$b)
    $n = $a.Length; $m = $b.Length
    if ($n -eq 0) { return $m }
    if ($m -eq 0) { return $n }
    $d = New-Object 'System.Int32[][]' ($n+1)
    for ($i=0; $i -le $n; $i++) { $d[$i] = New-Object 'System.Int32[]' ($m+1) }
    for ($i=0; $i -le $n; $i++) { $d[$i][0] = $i }
    for ($j=0; $j -le $m; $j++) { $d[0][$j] = $j }
    for ($i=1; $i -le $n; $i++) {
        for ($j=1; $j -le $m; $j++) {
            $cost = if ($a[$i-1] -eq $b[$j-1]) { 0 } else { 1 }
            $d[$i][$j] = [Math]::Min([Math]::Min($d[$i-1][$j]+1, $d[$i][$j-1]+1), $d[$i-1][$j-1]+$cost)
            if ($i -gt 1 -and $j -gt 1 -and $a[$i-1] -eq $b[$j-2] -and $a[$i-2] -eq $b[$j-1]) {
                $d[$i][$j] = [Math]::Min($d[$i][$j], $d[$i-2][$j-2]+1)
            }
        }
    }
    return $d[$n][$m]
}

# ---------- Spell-Cache ----------
# Ordinale Schluessel, Gross-/Kleinschreibung zaehlt. '@{}' vergleicht
# kulturabhaengig und ohne Gross-/Kleinschreibung: unter de-DE sind
# 'Straße' und 'Strasse' derselbe Schluessel (gemessen, PS 5.1). Hunspell
# prueft aber schreibweisengenau - der Cache lieferte dann das Ergebnis
# des jeweils anderen Worts, und ß/ss-Tippfehler fielen je nach
# Reihenfolge durch oder erzeugten Vorschlaege ohne Aenderung.
$script:SpellCache   = New-Object System.Collections.Hashtable ([StringComparer]::Ordinal)
$script:SuggestCache = New-Object System.Collections.Hashtable ([StringComparer]::Ordinal)

function Test-Spell {
    param([string]$word)
    if ($script:SpellCache.ContainsKey($word)) { return $script:SpellCache[$word] }
    $ok = $script:Hunspell.Spell($word)
    $script:SpellCache[$word] = $ok
    return $ok
}

function Get-SpellSuggest {
    param([string]$word)
    if ($script:SuggestCache.ContainsKey($word)) { return $script:SuggestCache[$word] }
    $s = @($script:Hunspell.Suggest($word))
    $script:SuggestCache[$word] = $s
    return $s
}

# ---------- Token-Analyse ----------

# Strikter Transpositions-Check: Damerau-Distanz=1 kann auch eine
# Substitution sein - diese Funktion akzeptiert NUR das Vertauschen
# zweier benachbarter Buchstaben (das klassische Zeichendreher-Muster).
function Test-IsTransposition {
    param([string]$a, [string]$b)
    if ($a.Length -ne $b.Length) { return $false }
    if ($a.Length -lt 2)         { return $false }
    $diffs = New-Object System.Collections.Generic.List[int]
    for ($i = 0; $i -lt $a.Length; $i++) {
        if ($a[$i] -ne $b[$i]) { $diffs.Add($i) | Out-Null }
        if ($diffs.Count -gt 2) { return $false }
    }
    if ($diffs.Count -ne 2)         { return $false }
    if ($diffs[1] -ne $diffs[0]+1)  { return $false }
    return ($a[$diffs[0]] -eq $b[$diffs[1]]) -and ($a[$diffs[1]] -eq $b[$diffs[0]])
}

# Sammelt alle Tokens aus den uebergebenen Pfaden (Datei-Stems und
# Ordnernamen), gibt eine Hashtable Token(lowercase) -> Haeufigkeit zurueck.
# Tokens kuerzer als $CorpusMinTokenLength werden ignoriert; ebenso reine
# Zahlen, ALL-CAPS-Akronyme und Tokens, die Hunspell als korrektes Wort
# kennt (dann ist die Korpus-Heuristik nicht informativer als das
# Woerterbuch und produziert tendenziell False-Positives).
function Build-CorpusVocabulary {
    param(
        [string[]]$Files,
        [string[]]$Dirs
    )
    # Ordinal statt '@{}': kulturabhaengig fielen 'strasse' und 'straße'
    # auf einen Schluessel und ihre Haeufigkeiten wurden zusammengezaehlt.
    $vocab = New-Object System.Collections.Hashtable ([StringComparer]::Ordinal)

    $addTokens = {
        param([string]$text)
        # Frueher hiess diese Variable $matches und ueberschrieb damit die
        # PowerShell-Automatikvariable, die einige Zeilen weiter vom
        # -cmatch-Operator gefuellt wird. Funktioniert hat es nur, weil
        # foreach den Enumerator vorher festhaelt - eine Falle bei jeder
        # spaeteren Aenderung.
        $tokenMatches = [regex]::Matches($text, '\p{L}+')
        foreach ($m in $tokenMatches) {
            $word = $m.Value
            if ($word.Length -lt $CorpusMinTokenLength) { continue }
            if ($word -cmatch '^[\p{Lu}0-9]+$')         { continue }
            $key = $word.ToLowerInvariant()
            if ($vocab.ContainsKey($key)) { $vocab[$key]++ }
            else                          { $vocab[$key] = 1 }
        }
    }

    foreach ($f in $Files) {
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($f)
        & $addTokens $stem
    }
    foreach ($d in $Dirs) {
        $leaf = Split-Path -Path $d -Leaf
        & $addTokens $leaf
    }
    return $vocab
}

# Baut zwei Indizes ueber das Vokabular auf. Ohne sie war die Suche nach
# einem Distanz-1-Partner eine vollstaendige Schleife ueber ALLE Tokens,
# mit einem Damerau-Levenshtein je Paar. Bei 100.000 Dateien und einem
# Vokabular von einigen zehntausend Tokens sind das schnell hunderte
# Millionen Levenshtein-Berechnungen in PowerShell - der Lauf war
# praktisch nicht zu Ende zu bringen.
#
#   AnagramIndex: sortierte Buchstabenfolge -> Tokenliste.
#     Eine Transposition vertauscht nur zwei benachbarte Buchstaben, das
#     Ergebnis ist ein Anagramm. Der Partner steht damit in einem
#     einzigen Hashtable-Zugriff fest.
#   LengthIndex: Wortlaenge -> Tokenliste.
#     Fuer die uebrigen Distanz-1-Faelle (Substitution, Einfuegung,
#     Loeschung) bleibt eine Schleife noetig, sie laeuft aber nur noch
#     ueber Tokens passender Laenge statt ueber das ganze Vokabular.
function Build-VocabIndex {
    param([hashtable]$Vocab)
    # Ordinal wie das Vokabular selbst (siehe Build-CorpusVocabulary).
    $anagram = New-Object System.Collections.Hashtable ([StringComparer]::Ordinal)
    $byLength = @{}
    foreach ($key in $Vocab.Keys) {
        $chars = $key.ToCharArray()
        [Array]::Sort($chars)
        $sig = -join $chars
        if (-not $anagram.ContainsKey($sig)) {
            $anagram[$sig] = New-Object System.Collections.Generic.List[string]
        }
        $anagram[$sig].Add($key)

        $len = $key.Length
        if (-not $byLength.ContainsKey($len)) {
            $byLength[$len] = New-Object System.Collections.Generic.List[string]
        }
        $byLength[$len].Add($key)
    }
    return @{ Anagram = $anagram; ByLength = $byLength }
}

# Pro Token im Stem: wenn das Token im Korpus selten ist
# (<= MaxRareFreq), suche im Vokabular einen Damerau-Distanz-1-Partner
# mit Haeufigkeit >= MinDominantFreq. Bevorzugt echte Transpositionen
# (zwei benachbarte Buchstaben vertauscht); akzeptiert bei eindeutigem
# Haeufigkeitsgefaelle auch Substitutionen/Insertionen.
function Test-NameForCorpusTransposition {
    param(
        [string]$stem,
        [hashtable]$Vocab,
        [hashtable]$Index = $null
    )
    if (-not $Vocab -or $Vocab.Count -eq 0) { return $null }

    $tokens = [regex]::Matches($stem, '\p{L}+')
    if ($tokens.Count -eq 0) { return $null }

    $candidates = New-Object System.Collections.Generic.List[hashtable]

    foreach ($t in $tokens) {
        $word = $t.Value
        if ($word.Length -lt $CorpusMinTokenLength) { continue }
        if ($word -cmatch '^[\p{Lu}0-9]+$')         { continue }

        $key    = $word.ToLowerInvariant()
        $myFreq = if ($Vocab.ContainsKey($key)) { $Vocab[$key] } else { 0 }
        if ($myFreq -gt $CorpusMaxRareFreq) { continue }

        $bestKey         = $null
        $bestFreq        = 0
        $bestIsTranspos  = $false

        # Kandidatenmenge stark eingrenzen, statt das ganze Vokabular zu
        # durchlaufen: Anagramme (Transpositionen) plus Tokens mit
        # Laenge n-1, n und n+1 (alle uebrigen Distanz-1-Faelle).
        $candidateKeys = $null
        if ($Index) {
            $chars = $key.ToCharArray()
            [Array]::Sort($chars)
            $sig = -join $chars
            $bucket = New-Object System.Collections.Generic.List[string]
            if ($Index.Anagram.ContainsKey($sig)) {
                foreach ($k in $Index.Anagram[$sig]) { $bucket.Add($k) }
            }
            # Klammern sind Pflicht: das Komma bindet staerker als '-', ohne
            # sie wurde "$key.Length - (1, $key.Length, ...)" gerechnet und
            # warf (op_Subtraction auf Object[], gemessen unter 5.1) - bei
            # fast jedem seltenen Wort landete die Datei als "Fehler".
            foreach ($len in @(($key.Length - 1), $key.Length, ($key.Length + 1))) {
                if ($Index.ByLength.ContainsKey($len)) {
                    foreach ($k in $Index.ByLength[$len]) { $bucket.Add($k) }
                }
            }
            # Ordinal entdoppeln. 'Select-Object -Unique' vergleicht unter
            # de-DE kulturabhaengig und warf 'straße' als Dublette von
            # 'strasse' hinaus (gemessen) - der Kandidat fehlte dann.
            $gesehen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
            $candidateKeys = New-Object System.Collections.Generic.List[string]
            foreach ($k in $bucket) { if ($gesehen.Add($k)) { $candidateKeys.Add($k) } }
        } else {
            $candidateKeys = $Vocab.Keys
        }

        foreach ($otherKey in $candidateKeys) {
            # Ordinal: '-eq' haelt unter de-DE 'straße' und 'strasse' fuer gleich.
            if ([string]::Equals($otherKey, $key, [System.StringComparison]::Ordinal)) { continue }
            if ($Vocab[$otherKey] -lt $CorpusMinDominantFreq) { continue }
            if ([Math]::Abs($otherKey.Length - $key.Length) -gt 1) { continue }

            # Billiger Vorfilter vor dem teuren Levenshtein: bei Distanz 1
            # unterscheidet sich hoechstens eine Stelle, also muss entweder
            # das erste oder das letzte Zeichen uebereinstimmen. Der Index
            # wird bewusst ausgeschrieben - negative String-Indizes sind in
            # PowerShell 5.1 nicht zuverlaessig.
            if ($otherKey[0] -ne $key[0] -and
                $otherKey[$otherKey.Length - 1] -ne $key[$key.Length - 1]) { continue }

            $dist = Get-EditDistance $key $otherKey
            if ($dist -ne 1) { continue }

            $isTrans = Test-IsTransposition $key $otherKey

            # Transposition schlaegt Nicht-Transposition; bei gleichem
            # Status gewinnt das haeufigere Vokabular-Wort.
            $isBetter = $false
            if ($isTrans -and -not $bestIsTranspos)            { $isBetter = $true }
            elseif ($isTrans -eq $bestIsTranspos -and $Vocab[$otherKey] -gt $bestFreq) {
                $isBetter = $true
            }

            if ($isBetter) {
                $bestKey        = $otherKey
                $bestFreq       = $Vocab[$otherKey]
                $bestIsTranspos = $isTrans
            }
        }

        if ($bestKey) {
            $newToken = Restore-Case -orig $word -new $bestKey
            $candidates.Add(@{
                Token     = $word
                Index     = $t.Index
                Length    = $t.Length
                New       = $newToken
                Dist      = 1
                Ambiguous = $false
                Source    = 'Bestand'
                FreqOld   = $myFreq
                FreqNew   = $bestFreq
                IsTrans   = $bestIsTranspos
            })
        }
    }

    if ($candidates.Count -eq 0)  { return $null }
    if ($candidates.Count -gt 1)  { return @{ Multi=$true; Tokens=$candidates; Source='Bestand' } }
    return $candidates[0]
}

function Test-NameForTypo {
    param([string]$stem)
    $tokens = [regex]::Matches($stem, '\p{L}+')
    if ($tokens.Count -eq 0) { return $null }

    $candidates = New-Object System.Collections.Generic.List[hashtable]
    foreach ($t in $tokens) {
        $word = $t.Value
        if ($word.Length -le 3)              { continue }
        if ($word -cmatch '^[\p{Lu}0-9]+$')  { continue }
        if (Test-Spell $word)                { continue }

        $sugg = Get-SpellSuggest $word
        if ($sugg.Count -eq 0) { continue }

        $top = $sugg[0]
        if ($top -match '\s')           { continue }
        if ($top -match '[\\/:*?"<>|]') { continue }
        if ($top.EndsWith('.'))         { continue }

        $dist = Get-EditDistance $word.ToLower() $top.ToLower()
        if ($dist -gt 2) { continue }

        $ambiguous = $false
        if ($sugg.Count -ge 2) {
            $dist2 = Get-EditDistance $word.ToLower() $sugg[1].ToLower()
            if (($dist + 1) -gt $dist2) { $ambiguous = $true }
        }

        $candidates.Add(@{
            Token     = $word
            Index     = $t.Index
            Length    = $t.Length
            New       = $top
            Dist      = $dist
            Ambiguous = $ambiguous
            Source    = 'Woerterbuch'
        })
    }

    if ($candidates.Count -eq 0)  { return $null }
    if ($candidates.Count -gt 1)  { return @{ Multi=$true; Tokens=$candidates; Source='Woerterbuch' } }
    if ($candidates[0].Ambiguous) { return @{ Multi=$true; Tokens=$candidates; Source='Woerterbuch' } }
    return $candidates[0]
}

# ---------- Validierung ----------
function Test-NameValid {
    param([string]$name, [string]$parentPath)
    if (-not $name)                                    { return @{ Ok=$false; Reason='Leer' } }
    if ($name -match '[\\/:*?"<>|]')                   { return @{ Ok=$false; Reason='Ungueltige Zeichen' } }
    if ($name.EndsWith('.') -or $name.EndsWith(' '))   { return @{ Ok=$false; Reason='Punkt/Leerzeichen am Ende' } }
    $base = [System.IO.Path]::GetFileNameWithoutExtension($name).ToUpper()
    if ($ReservedNames -contains $base)                { return @{ Ok=$false; Reason='Reservierter Name' } }
    # Der Zielpfad muss unter der Grenze bleiben - nicht wegen der
    # Windows-API (dafuer gibt es den '\\?\'-Praefix), sondern weil der
    # Bestand anschliessend nach Google Drive wandert. Eigener Status
    # statt 'Fehler': das ist ein bewusst uebersprungener Eintrag, kein
    # Fehlschlag, und wurde in der Statistik bisher falsch einsortiert.
    $full = Join-Path $parentPath $name
    if ($full.Length -gt $MaxPathLen) {
        return @{ Ok=$false; Status='ZuLang'; Reason="Zielpfad >$MaxPathLen Zeichen" }
    }
    return @{ Ok=$true }
}

# ---------- Case-Wiederherstellung ----------
function Restore-Case {
    param([string]$orig, [string]$new)
    if ($orig.Length -gt 1 -and $orig -ceq $orig.ToUpper()) { return $new.ToUpper() }
    if ($orig -ceq $orig.ToLower())                         { return $new.ToLower() }
    if ($orig[0] -cmatch '\p{Lu}') {
        if ($new.Length -le 1) { return $new.ToUpper() }
        return $new.Substring(0,1).ToUpper() + $new.Substring(1).ToLower()
    }
    return $new.ToLower()
}

# ---------- Ziel-Namen bauen ----------
function Format-NewName {
    param([string]$stem, [string]$ext, [hashtable]$res)
    $newTok  = Restore-Case -orig $res.Token -new $res.New
    $newStem = $stem.Substring(0, $res.Index) + $newTok + $stem.Substring($res.Index + $res.Length)
    return ($newStem + $ext)
}

# ---------- Rename ----------
function Invoke-SafeRename {
    param([string]$FullPath, [string]$NewName, [string]$Kind)
    # Intern konsequent mit '\\?\'-Praefix arbeiten. Vorher liefen
    # File.Move/Directory.Move und Test-Path auf den nackten Pfaden -
    # Eintraege jenseits von 259 Zeichen wurden dadurch gar nicht erst
    # umbenannt, sondern pauschal als 'Fehler' verbucht.
    $FullPath = Remove-LongPathPrefix $FullPath
    $parent   = Split-Path -Parent $FullPath
    $check    = Test-NameValid -name $NewName -parentPath $parent
    if (-not $check.Ok) {
        $st = if ($check.Status) { $check.Status } else { 'Fehler' }
        return @{ Status=$st; Details=$check.Reason }
    }

    $target = Join-Path $parent $NewName

    # Wirklich unveraendert nur, wenn auch die Gross-/Kleinschreibung
    # uebereinstimmt. Frueheres -ieq blockierte gewollte Case-Korrekturen
    # (z.B. "dokument.txt" -> "Dokument.txt") faelschlich als "Unveraendert".
    # Ordinal vergleichen: -ceq/-ieq/-cne vergleichen kulturabhaengig, und
    # unter de-DE ist 'Strasse' -ceq 'Straße' $true (gemessen, PS 5.1) -
    # jede ß/ss-Korrektur wurde so als "Unveraendert" verbucht und nie
    # ausgefuehrt. Fuer NTFS sind die beiden Namen verschieden.
    if ([string]::Equals($target, $FullPath, [System.StringComparison]::Ordinal)) { return @{ Status='Unveraendert' } }

    # Case-Only-Aenderung erkennen: alter und neuer Pfad sind auf einem
    # case-insensitiven Dateisystem (NTFS) "gleich", unterscheiden sich aber
    # in der Schreibweise. Direktes File.Move scheitert dann mit
    # "Source and destination path must be different" - Workaround: ueber
    # eindeutigen Zwischennamen umbenennen. OrdinalIgnoreCase statt -ieq
    # (kulturabhaengig, s.o.): sonst galt 'Straße' -> 'Strasse' als
    # Case-Only und lief unnoetig ueber den Zwischennamen.
    $isCaseOnly = [string]::Equals($target, $FullPath, [System.StringComparison]::OrdinalIgnoreCase)

    # Langpfad-Fassungen fuer alle Dateisystem-Zugriffe.
    $srcLong    = Add-LongPathPrefix $FullPath
    $targetLong = Add-LongPathPrefix $target

    if (-not $isCaseOnly -and
        ([System.IO.File]::Exists($targetLong) -or [System.IO.Directory]::Exists($targetLong))) {
        return @{ Status='Konflikt'; Details='Ziel existiert' }
    }

    $restoreRO = $false
    try {
        $attr = [System.IO.File]::GetAttributes($srcLong)
        if (($attr -band [System.IO.FileAttributes]::ReadOnly) -ne 0) {
            [System.IO.File]::SetAttributes($srcLong, $attr -bxor [System.IO.FileAttributes]::ReadOnly)
            $restoreRO = $true
        }
    } catch {}

    for ($i=1; $i -le 3; $i++) {
        try {
            if ($isCaseOnly) {
                # Two-Step ueber kurzen Zwischennamen, garantiert
                # kollisionsfrei. Bewusst OHNE den neuen Namen im
                # Zwischennamen, damit lange Dateinamen nicht das
                # Pfadlimit reissen. Scheitert der zweite Move, wird
                # der erste zurueckgerollt - sonst straendete der
                # Eintrag dauerhaft unter dem GUID-Zwischennamen.
                $intermediate = Add-LongPathPrefix (Join-Path $parent ("__caserename_" + [Guid]::NewGuid().ToString('N') + [System.IO.Path]::GetExtension($NewName)))
                if ($Kind -eq 'Datei') {
                    [System.IO.File]::Move($srcLong, $intermediate)
                    try {
                        [System.IO.File]::Move($intermediate, $targetLong)
                    } catch {
                        try { [System.IO.File]::Move($intermediate, $srcLong) } catch {}
                        throw
                    }
                } else {
                    [System.IO.Directory]::Move($srcLong, $intermediate)
                    try {
                        [System.IO.Directory]::Move($intermediate, $targetLong)
                    } catch {
                        try { [System.IO.Directory]::Move($intermediate, $srcLong) } catch {}
                        throw
                    }
                }
            } else {
                if ($Kind -eq 'Datei') { [System.IO.File]::Move($srcLong, $targetLong) }
                else                   { [System.IO.Directory]::Move($srcLong, $targetLong) }
            }
            if ($restoreRO) {
                try {
                    $a = [System.IO.File]::GetAttributes($targetLong)
                    [System.IO.File]::SetAttributes($targetLong, $a -bor [System.IO.FileAttributes]::ReadOnly)
                } catch {
                    Write-Warning "ReadOnly-Attribut auf '$target' konnte nicht wiederhergestellt werden: $($_.Exception.Message)"
                }
            }
            return @{ Status='Umbenannt' }
        } catch {
            if ($i -eq 3) {
                $fehlerText = $_.Exception.Message
                # Schreibschutz zuruecksetzen, wenn die Umbenennung endgueltig
                # gescheitert ist. Vorher blieb das oben entfernte ReadOnly-
                # Attribut dann dauerhaft weg (nachgestellt: gesperrte,
                # schreibgeschuetzte Datei -> 'Fehler', danach ohne ReadOnly).
                if ($restoreRO) {
                    try {
                        $a = [System.IO.File]::GetAttributes($srcLong)
                        [System.IO.File]::SetAttributes($srcLong, $a -bor [System.IO.FileAttributes]::ReadOnly)
                    } catch {
                        Write-Warning "ReadOnly-Attribut auf '$FullPath' konnte nicht wiederhergestellt werden: $($_.Exception.Message)"
                    }
                }
                return @{ Status='Fehler'; Details=$fehlerText }
            }
            Start-Sleep -Milliseconds (250*$i)
        }
    }
}

# ---------- Rekursive Enumeration (robust) ----------
function Add-LongPathPrefix {
    param([string]$Path)
    if ([string]::IsNullOrEmpty($Path)) { return $Path }
    if ($Path.StartsWith('\\?\'))      { return $Path }
    if ($Path -match '^[A-Za-z]:$')     { return '\\?\' + $Path + '\' }
    if ($Path.StartsWith('\\'))         { return '\\?\UNC\' + $Path.Substring(2) }
    return '\\?\' + $Path
}

function Remove-LongPathPrefix {
    param([string]$Path)
    if ([string]::IsNullOrEmpty($Path))   { return $Path }
    if ($Path.StartsWith('\\?\UNC\'))    { return '\\' + $Path.Substring(8) }
    if ($Path.StartsWith('\\?\'))        { return $Path.Substring(4) }
    return $Path
}

function Get-AllPaths {
    param([string]$root)
    $files = New-Object System.Collections.Generic.List[string]
    $dirs  = New-Object System.Collections.Generic.List[string]
    $stack = New-Object System.Collections.Generic.Stack[string]
    $stack.Push($root)
    $script:LongPathSkips = 0

    while ($stack.Count -gt 0) {
        $cur = $stack.Pop()

        # EnumerateFiles/-Directories liefern einen Iterator. Wirft er
        # mitten in der Aufzaehlung (typisch: PathTooLongException bei
        # einem einzelnen zu langen Eintrag), bricht die GESAMTE Schleife
        # fuer dieses Verzeichnis ab - alle nachfolgenden Eintraege
        # fehlten dann stillschweigend im Ergebnis. Deshalb: erst die
        # komplette Liste materialisieren, und bei einem Fehler ohne
        # Long-Path-Praefix denselben Ordner mit Praefix erneut versuchen.
        $fileList = $null
        try {
            $fileList = @([System.IO.Directory]::GetFiles($cur))
        } catch {
            try {
                $fileList = @([System.IO.Directory]::GetFiles((Add-LongPathPrefix $cur)))
            } catch {
                $script:LongPathSkips++
                Write-Warning "Dateien in '$(Remove-LongPathPrefix $cur)' nicht lesbar: $($_.Exception.Message)"
                $fileList = @()
            }
        }
        foreach ($f in $fileList) {
            try {
                $ext = [System.IO.Path]::GetExtension($f)
                if (Test-SkipExt $ext) { continue }
                $files.Add($f)
            } catch {}
        }

        $dirList = $null
        try {
            $dirList = @([System.IO.Directory]::GetDirectories($cur))
        } catch {
            try {
                $dirList = @([System.IO.Directory]::GetDirectories((Add-LongPathPrefix $cur)))
            } catch {
                $script:LongPathSkips++
                Write-Warning "Ordner in '$(Remove-LongPathPrefix $cur)' nicht lesbar: $($_.Exception.Message)"
                $dirList = @()
            }
        }
        foreach ($d in $dirList) {
            try {
                $name = Split-Path -Leaf $d
                if (Test-SkipDir $name) { continue }
                $attr = [System.IO.File]::GetAttributes($d)
                if (($attr -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
                $dirs.Add($d)
                $stack.Push($d)
            } catch {}
        }
    }

    return @{ Files=$files; Dirs=$dirs }
}

# ---------- Benutzerinteraktion ----------
function Read-UserChoice {
    param([string]$oldName, [string]$newName, [string]$details = $null)
    if ($script:autoYes) { return 'J' }
    if ($script:skipAll) { return 'N' }

    while ($true) {
        Write-Host ""
        Write-Host "  Alt: $oldName" -ForegroundColor Yellow
        Write-Host "  Neu: $newName" -ForegroundColor Green
        if ($details) {
            Write-Host "  Info: $details" -ForegroundColor DarkGray
        }
        $choice = Read-Host "  [J]a / [N]ein / [A]lle ja / [K]eine mehr / [B]earbeiten / [Q]uit"
        switch -Regex ($choice.Trim().ToUpper()) {
            '^J$' { return 'J' }
            '^N$' { return 'N' }
            '^A$' { $script:autoYes = $true; return 'J' }
            '^K$' { $script:skipAll = $true; return 'N' }
            '^Q$' { $script:abort   = $true; return 'N' }
            '^B$' {
                $custom = Read-Host "  Neuer Name"
                $custom = $custom.Trim().Trim('"').Trim("'")
                if ($custom) { return "B:$custom" }
                Write-Host "  Leer -> erneute Auswahl." -ForegroundColor DarkGray
            }
            default {
                Write-Host "  Ungueltige Eingabe." -ForegroundColor DarkGray
            }
        }
    }
}

function Read-SingleKey {
    param([string]$prompt)
    Write-Host $prompt -NoNewline
    try {
        $null = [System.Console]::ReadKey($true)
    } catch {
        $null = Read-Host
    }
    Write-Host ""
}

# ---------- Windows-Temp-Cleanup ----------
# Whitelist: NUR Eintraege mit diesen Praefixen (Office/COM-Reste sowie
# eigene Skript-Artefakte) duerfen geloescht werden. NIEMALS pauschal
# leeren: Fremdprozesse legen aktive Daten ohne Lock in %TEMP% ab;
# blindes Loeschen zerstoert sie.
$script:WindowsTempWhitelistPrefixes = @(
    '~$', '~df', 'gen_py', 'vbe', 'excel8.0',
    '11_typo_dateinamen_korrigieren'
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
    #
    # Protokolle dieses Skripts sind KEINE Reste. Landet das Protokoll im
    # TEMP-Rueckfall von Get-WritableDir, traegt es das Whitelist-Praefix
    # '11_typo_dateinamen_korrigieren' und wurde hier am Ende desselben
    # Laufs geloescht - die Zusammenfassung nannte danach einen Pfad, den
    # es nicht mehr gab (nachgestellt mit umgelenktem TEMP). Ebenso traf
    # es die Scan-CSV eines frueheren Laufs, die Eingabe fuer Modus 3.
    param([string[]]$Ausnehmen = @())
    $winTemp = $env:TEMP
    if (-not $winTemp) { $winTemp = $env:TMP }
    if (-not $winTemp -or -not (Test-Path -LiteralPath $winTemp)) { return }
    Get-ChildItem -LiteralPath $winTemp -Force -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            if (-not (Test-WhitelistedTempEntry $_.Name)) { return }
            $eintrag = $_.FullName
            if (@($Ausnehmen | Where-Object { $_ -and [string]::Equals($_, $eintrag, [System.StringComparison]::OrdinalIgnoreCase) }).Count -gt 0) { return }
            if (-not $_.PSIsContainer -and
                $_.Name -match '^11_Typo_Dateinamen_korrigieren_(Scan|Interactive|Apply)_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}\.csv$') { return }
            if ($_.PSIsContainer) {
                Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
            } else {
                Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
            }
        } catch {}
    }
}

# ---------- Log (Streaming) ----------
function New-LogWriter {
    param([string]$path)
    # UTF-8 MIT BOM. Vorher wurde ohne BOM geschrieben - Excel oeffnet eine
    # UTF-8-CSV ohne BOM als ANSI, wodurch alle Umlaute in den Pfaden
    # zerstoert dargestellt werden. Der dokumentierte Arbeitsablauf ist
    # aber genau dieser: Scan-CSV in Excel oeffnen, Spalte 'Uebernehmen'
    # pflegen, speichern, Modus 3. Beim Speichern wurden die kaputten
    # Umlaute festgeschrieben und der Anwenden-Modus fand die Dateien
    # anschliessend nicht mehr ("Quelle nicht mehr vorhanden").
    $enc = New-Object System.Text.UTF8Encoding($true)
    $sw  = New-Object System.IO.StreamWriter($path, $false, $enc)
    $sw.AutoFlush = $true
    $sw.WriteLine('Typ;AlterPfad;NeuerName;Status;Details;Uebernehmen')
    return $sw
}

function Write-LogRow {
    param([System.IO.StreamWriter]$writer, [hashtable]$row)

    # Gemeinsames Laufprotokoll (migration.jsonl) - ergaenzt das
    # skripteigene Protokoll, ersetzt es nicht. Erst damit laesst sich
    # der Fortschritt ueber alle elf Schritte hinweg auswerten.
    if ($script:GemeinsamGeladen) {
        try {
            Write-Laufprotokoll -Skript '11_Typo_Dateinamen_korrigieren' `
                -Pfad ([string]$row.AlterPfad) -Aktion 'Dateiname pruefen' `
                -Status ([string]$row.Status) -Detail ([string]$row.NeuerName)
        } catch { }
    }
    $d = if ($null -ne $row.Details) { [string]$row.Details } else { '' }
    $u = if ($row.ContainsKey('Uebernehmen') -and $null -ne $row.Uebernehmen) { [string]$row.Uebernehmen } else { '' }

    # RFC-4180-konformes CSV-Quoting: Werte mit Trennzeichen (;), Quote (")
    # oder Zeilenumbruch in "..." einschliessen und enthaltene " durch ""
    # escapen. Semikolons in Dateinamen sind unter Windows zulaessig
    # (z.B. "Protokoll;Meeting.docx") - frueheres Hart-Replace ';'->','
    # zerstoerte solche Pfade beim Round-Trip ueber den ApplyFromCsv-Modus.
    $quote = {
        param([string]$v)
        if ($null -eq $v) { return '' }
        if ($v -match '[;"\r\n]') {
            return '"' + ($v -replace '"','""') + '"'
        }
        return $v
    }

    # Details und Status enthalten keine vom User kontrollierten Pfade, aber
    # Status-Strings koennen evtl. spaeter erweitert werden - sicherer ist
    # einheitliches Quoting fuer alle Felder.
    $line = '{0};{1};{2};{3};{4};{5}' -f `
        (& $quote $row.Typ),
        (& $quote $row.AlterPfad),
        (& $quote $row.NeuerName),
        (& $quote $row.Status),
        (& $quote ($d -replace "`r?`n",' ')),
        (& $quote $u)
    $writer.WriteLine($line)
}

# ---------- Menues ----------
function Select-Mode {
    Write-Host ""
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host "Welcher Modus?"                                           -ForegroundColor Cyan
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  [1]  Scan        (nur Vorschlaege in CSV, keine Aenderung)" -ForegroundColor White
    Write-Host "  [2]  Interaktiv  (je Fund nachfragen, sofort umbenennen)"  -ForegroundColor White
    Write-Host "  [3]  Anwenden    (aus bearbeiteter Scan-CSV)"              -ForegroundColor White
    Write-Host ""
    $m = Read-Host "Ihre Wahl (1-3)"
    switch ($m.Trim()) {
        "1" { return 'Scan' }
        "2" { return 'Interactive' }
        "3" { return 'ApplyFromCsv' }
        default { return $null }
    }
}

function Select-RootPath {
    Write-Host ""
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host "Welches Laufwerk soll durchsucht werden?"                 -ForegroundColor Cyan
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host ""
    $presets = @($DirectoryPresets | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    for ($i = 0; $i -lt $presets.Count; $i++) {
        Write-Host ("  [{0}]  {1}" -f ($i + 1), $presets[$i]) -ForegroundColor White
    }
    $manualIdx = $presets.Count + 1
    Write-Host ("  [{0}]  Eigenen Pfad eingeben" -f $manualIdx) -ForegroundColor White
    Write-Host ""
    $c = (Read-Host ("Ihre Wahl (1-{0})" -f $manualIdx)).Trim()

    $num = 0
    if (-not [int]::TryParse($c, [ref]$num)) { $num = -1 }

    if ($num -ge 1 -and $num -le $presets.Count) {
        $sel = $presets[$num - 1]
        # Preset-Pfade koennen auf diesem Rechner fehlen. Vorher wurden sie
        # ungeprueft zurueckgegeben; der Abbruch kam spaeter ohne Hinweis
        # auf die haeufigste Ursache.
        if (-not (Test-Path -LiteralPath $sel)) {
            Write-Host "  '$sel' ist nicht erreichbar." -ForegroundColor Red
            Write-Host "  Hinweis: Bei erhoehten Rechten sind gemappte Netzlaufwerke oft" -ForegroundColor Yellow
            Write-Host "  ausgeblendet - dann bitte den UNC-Pfad manuell eingeben." -ForegroundColor Yellow
            return $null
        }
        return $sel
    }
    if ($num -eq $manualIdx) {
        $p = Read-Host "Bitte vollstaendigen Pfad eingeben"
        $p = $p.Trim().Trim('"').Trim("'")
        if ($p -match '^[A-Za-z]:$') { $p = $p + '\' }
        if ($p -notmatch '\\$')      { $p = $p + '\' }
        return $p
    }
    return $null
}

# ---------- Apply-CSV laden ----------
function Import-ApplyCsv {
    param([string]$path)
    if (-not (Test-Path -LiteralPath $path)) { throw "CSV nicht gefunden: $path" }
    $list = New-Object System.Collections.Generic.List[pscustomobject]
    # -LiteralPath statt -Path: ohne den Schalter interpretiert PowerShell
    # eckige Klammern im Pfad als Wildcard-Zeichenklassen (z.B. "[Archiv]"),
    # wodurch die CSV-Datei nicht gefunden wird und der Anwenden-Modus
    # komplett scheitert.
    $rows = Import-Csv -LiteralPath $path -Delimiter ';' -Encoding UTF8
    foreach ($r in $rows) {
        if (-not $r.AlterPfad)          { continue }
        if (-not $r.NeuerName)          { continue }
        if ($r.Uebernehmen -notmatch '^(J|j|Y|y|1)$') { continue }
        $list.Add([pscustomobject]@{
            Typ         = $r.Typ
            AlterPfad   = $r.AlterPfad
            NeuerName   = $r.NeuerName
            Uebernehmen = $r.Uebernehmen
        })
    }
    return $list
}

# ---------- Verarbeitung (Scan + Interactive) ----------
function Invoke-ItemProcessing {
    param(
        [string]$FullPath,
        [string]$Kind,
        [string]$ProcessMode,
        [System.IO.StreamWriter]$LogWriter,
        [hashtable]$Stats,
        [hashtable]$Vocab = $null,
        [hashtable]$VocabIndex = $null
    )
    if ($script:abort) { return }

    $name = Split-Path -Leaf $FullPath
    if ($name.StartsWith('~$')) { return }
    try {
        $attr = [System.IO.File]::GetAttributes($FullPath)
        if (($attr -band ([System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System)) -ne 0) { return }
    } catch { return }

    if ($Kind -eq 'Datei') {
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($name)
        $ext  = [System.IO.Path]::GetExtension($name)
    } else {
        $stem = $name
        $ext  = ''
    }

    # Erst Korpus-Pass (sehr spezifisch fuer Zeichendreher, niedrige
    # False-Positive-Rate dank Haeufigkeitsgefaelle). Bei Treffer den
    # nehmen, sonst Fallback auf Hunspell-Pass.
    $res = $null
    if ($Vocab) {
        $res = Test-NameForCorpusTransposition -stem $stem -Vocab $Vocab -Index $VocabIndex
    }
    if (-not $res) {
        $res = Test-NameForTypo $stem
    }
    if (-not $res) { return }

    if ($res.Multi) {
        $tokInfo  = ($res.Tokens | ForEach-Object { "$($_.Token)->$($_.New)" }) -join ', '
        $srcLabel = if ($res.Source -eq 'Bestand') { 'Bestand' } else { 'Woerterbuch' }
        $Stats['ZurPruefung']++
        Write-LogRow -writer $LogWriter -row @{
            Typ=$Kind; AlterPfad=$FullPath; NeuerName=''; Status='ZurPruefung';
            Details=("Mehrdeutig ($srcLabel): $tokInfo"); Uebernehmen='N'
        }
        if ($ProcessMode -eq 'Interactive') {
            Write-Host "[ZurPruefung] $FullPath" -ForegroundColor DarkYellow
        }
        return
    }

    $newName = Format-NewName -stem $stem -ext $ext -res $res

    # Details fuer Log + UI: Quelle und (bei Korpus) Haeufigkeitsgefaelle
    $srcLabel = if ($res.Source -eq 'Bestand') { 'Bestand' } else { 'Woerterbuch' }
    $detailStr = "Token $($res.Token)->$($res.New), Dist=$($res.Dist), Quelle=$srcLabel"
    if ($res.Source -eq 'Bestand') {
        $transTag = if ($res.IsTrans) { ', Transposition' } else { '' }
        $detailStr += " ($($res.FreqNew)x '$($res.New)' vs $($res.FreqOld)x '$($res.Token)'$transTag)"
    }

    if ($ProcessMode -eq 'Scan') {
        $Stats['Vorschlag']++
        Write-LogRow -writer $LogWriter -row @{
            Typ=$Kind; AlterPfad=$FullPath; NeuerName=$newName; Status='Vorschlag';
            Details=$detailStr; Uebernehmen='J'
        }
        return
    }

    $decision = Read-UserChoice -oldName $name -newName $newName -details $detailStr
    if ($decision.StartsWith('B:')) {
        $custom = $decision.Substring(2).Trim().TrimEnd('.').Trim()
        if (-not $custom) {
            $Stats['Uebersprungen']++
            Write-LogRow -writer $LogWriter -row @{
                Typ=$Kind; AlterPfad=$FullPath; NeuerName=''; Status='Uebersprungen';
                Details='Leerer Name bei [B]'; Uebernehmen='N'
            }
            return
        }
        if ($custom -match '[\\/:*?"<>|]') {
            $Stats['Uebersprungen']++
            Write-LogRow -writer $LogWriter -row @{
                Typ=$Kind; AlterPfad=$FullPath; NeuerName=$custom; Status='Uebersprungen';
                Details='Ungueltige Zeichen im eigenen Namen'; Uebernehmen='N'
            }
            return
        }
        if ($Kind -eq 'Datei' -and $ext -and ($custom -notmatch '\.[^.\\/]+$')) {
            $newName = $custom + $ext
        } else {
            $newName = $custom
        }
        $decision = 'J'
    }

    if ($decision -ne 'J') {
        $Stats['Uebersprungen']++
        Write-LogRow -writer $LogWriter -row @{
            Typ=$Kind; AlterPfad=$FullPath; NeuerName=$newName; Status='Uebersprungen';
            Details=''; Uebernehmen='N'
        }
        return
    }

    $r = Invoke-SafeRename -FullPath $FullPath -NewName $newName -Kind $Kind
    $Stats[$r.Status]++
    Write-LogRow -writer $LogWriter -row @{
        Typ=$Kind; AlterPfad=$FullPath; NeuerName=$newName;
        Status=$r.Status; Details=$r.Details; Uebernehmen='J'
    }
    $color = switch ($r.Status) { 'Umbenannt' {'Green'} 'Konflikt' {'Yellow'} 'Fehler' {'Red'} default {'Gray'} }
    Write-Host "  -> [$($r.Status)]" -ForegroundColor $color
}

# ---------- Apply-Modus ----------
function Invoke-ApplyMode {
    param([System.Collections.Generic.List[pscustomobject]]$Entries,
          [System.IO.StreamWriter]$LogWriter,
          [hashtable]$Stats)

    $files = @($Entries | Where-Object { $_.Typ -eq 'Datei' })
    $dirs  = @($Entries | Where-Object { $_.Typ -eq 'Ordner' } |
               Sort-Object -Property `
                   @{ Expression = { ($_.AlterPfad -split '\\').Count }; Descending = $true },
                   @{ Expression = { $_.AlterPfad }; Descending = $true })

    $total = $files.Count + $dirs.Count
    $idx   = 0

    foreach ($e in $files) {
        Test-AbortRequested
        if ($script:abort) { break }
        $idx++
        if (-not $NoProgress -and ($idx % 25 -eq 0 -or $idx -eq $total)) {
            Write-Progress -Activity "Anwenden (Dateien)" -Status "$idx / $total" `
                -PercentComplete ([int](($idx*100)/[Math]::Max(1,$total)))
        }
        if (-not (Test-Path -LiteralPath $e.AlterPfad)) {
            $Stats['Fehler']++
            Write-LogRow -writer $LogWriter -row @{
                Typ='Datei'; AlterPfad=$e.AlterPfad; NeuerName=$e.NeuerName;
                Status='Fehler'; Details='Quelle nicht mehr vorhanden'; Uebernehmen=$e.Uebernehmen
            }
            continue
        }
        $r = Invoke-SafeRename -FullPath $e.AlterPfad -NewName $e.NeuerName -Kind 'Datei'
        $Stats[$r.Status]++
        Write-LogRow -writer $LogWriter -row @{
            Typ='Datei'; AlterPfad=$e.AlterPfad; NeuerName=$e.NeuerName;
            Status=$r.Status; Details=$r.Details; Uebernehmen=$e.Uebernehmen
        }
    }

    foreach ($e in $dirs) {
        Test-AbortRequested
        if ($script:abort) { break }
        $idx++
        if (-not $NoProgress -and ($idx % 25 -eq 0 -or $idx -eq $total)) {
            Write-Progress -Activity "Anwenden (Ordner)" -Status "$idx / $total" `
                -PercentComplete ([int](($idx*100)/[Math]::Max(1,$total)))
        }
        if (-not (Test-Path -LiteralPath $e.AlterPfad)) {
            $Stats['Fehler']++
            Write-LogRow -writer $LogWriter -row @{
                Typ='Ordner'; AlterPfad=$e.AlterPfad; NeuerName=$e.NeuerName;
                Status='Fehler'; Details='Quelle nicht mehr vorhanden'; Uebernehmen=$e.Uebernehmen
            }
            continue
        }
        $r = Invoke-SafeRename -FullPath $e.AlterPfad -NewName $e.NeuerName -Kind 'Ordner'
        $Stats[$r.Status]++
        Write-LogRow -writer $LogWriter -row @{
            Typ='Ordner'; AlterPfad=$e.AlterPfad; NeuerName=$e.NeuerName;
            Status=$r.Status; Details=$r.Details; Uebernehmen=$e.Uebernehmen
        }
    }

    if (-not $NoProgress) { Write-Progress -Activity "Anwenden" -Completed }
}

# ---------- Scan/Interactive-Modus ----------
function Invoke-ScanOrInteractive {
    param([string]$Root,
          [string]$ProcessMode,
          [System.IO.StreamWriter]$LogWriter,
          [hashtable]$Stats)

    Write-Host "Sammle Dateisystem ..." -ForegroundColor Cyan
    Write-Host "  Hinweis: Bei Netzlaufwerken (UNC) und sehr vielen Dateien" -ForegroundColor DarkGray
    Write-Host "  (>100.000) kann das Sammeln 10-30 Minuten dauern." -ForegroundColor DarkGray
    $paths  = Get-AllPaths -root $Root
    $nFiles = $paths.Files.Count
    $nDirs  = $paths.Dirs.Count
    Write-Host "  Dateien: $nFiles, Ordner: $nDirs" -ForegroundColor DarkGray
    if ($script:LongPathSkips -gt 0) {
        Write-Host "  WARNUNG: $($script:LongPathSkips) Verzeichnis(se) waren nicht lesbar - deren Inhalt fehlt im Ergebnis." -ForegroundColor Yellow
    }

    Write-Host "Baue Korpus-Vokabular ..." -ForegroundColor Cyan
    $vocab = Build-CorpusVocabulary -Files $paths.Files -Dirs $paths.Dirs
    Write-Host "  Tokens (eindeutig): $($vocab.Count)" -ForegroundColor DarkGray
    $vocabIndex = Build-VocabIndex -Vocab $vocab
    Write-Host "  Index: $($vocabIndex.Anagram.Count) Anagramm-Gruppen, $($vocabIndex.ByLength.Count) Laengenklassen" -ForegroundColor DarkGray

    $idx = 0
    foreach ($f in $paths.Files) {
        Test-AbortRequested
        if ($script:abort) { break }
        $idx++
        if (-not $NoProgress -and ($idx % 100 -eq 0 -or $idx -eq $nFiles)) {
            Write-Progress -Activity "Dateien pruefen" -Status "$idx / $nFiles" `
                -PercentComplete ([int](($idx*100)/[Math]::Max(1,$nFiles)))
        }
        try {
            Invoke-ItemProcessing -FullPath $f -Kind 'Datei' -ProcessMode $ProcessMode `
                -LogWriter $LogWriter -Stats $Stats -Vocab $vocab -VocabIndex $vocabIndex
        } catch {
            $Stats['Fehler']++
            Write-LogRow -writer $LogWriter -row @{
                Typ='Datei'; AlterPfad=$f; NeuerName=''; Status='Fehler';
                Details=$_.Exception.Message; Uebernehmen='N'
            }
        }
    }
    if (-not $NoProgress) { Write-Progress -Activity "Dateien pruefen" -Completed }

    if ($script:abort) { return }

    $sortedDirs = $paths.Dirs | Sort-Object -Property `
        @{ Expression = { ($_ -split '\\').Count }; Descending = $true },
        @{ Expression = { $_ }; Descending = $true }

    $idx = 0
    $nTot = @($sortedDirs).Count
    foreach ($d in $sortedDirs) {
        Test-AbortRequested
        if ($script:abort) { break }
        $idx++
        if (-not $NoProgress -and ($idx % 50 -eq 0 -or $idx -eq $nTot)) {
            Write-Progress -Activity "Ordner pruefen" -Status "$idx / $nTot" `
                -PercentComplete ([int](($idx*100)/[Math]::Max(1,$nTot)))
        }
        if (-not [System.IO.Directory]::Exists($d)) { continue }
        try {
            Invoke-ItemProcessing -FullPath $d -Kind 'Ordner' -ProcessMode $ProcessMode `
                -LogWriter $LogWriter -Stats $Stats -Vocab $vocab -VocabIndex $vocabIndex
        } catch {
            $Stats['Fehler']++
            Write-LogRow -writer $LogWriter -row @{
                Typ='Ordner'; AlterPfad=$d; NeuerName=''; Status='Fehler';
                Details=$_.Exception.Message; Uebernehmen='N'
            }
        }
    }
    if (-not $NoProgress) { Write-Progress -Activity "Ordner pruefen" -Completed }
}

# ---------- Main ----------
$script:Hunspell = $null
$script:autoYes  = $false
$script:skipAll  = $false
$script:abort    = $false
$logWriter       = $null
# Bei Aufruf mit -Mode (Scheduled-Task/Silent) am Ende nicht auf einen
# Tastendruck warten - der geplante Task wuerde sonst dauerhaft haengen.
$script:InteractiveLaunch = -not $PSBoundParameters.ContainsKey('Mode')

# Strg+C als Eingabe behandeln und in den Schleifen pollen: ein
# add_CancelKeyPress-ScriptBlock wuerde auf dem Ctrl+C-Thread ohne
# PowerShell-Runspace laufen und dort fehlschlagen - der Abbruch kaeme
# dann doch hart mitten in einer Umbenennung.
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
                if (-not $script:abort) {
                    Write-Host "`n[ABBRUCH]   Strg+C erkannt - laufende Operation wird noch fertiggestellt..." -ForegroundColor Yellow
                }
                $script:abort = $true
            }
        }
    } catch {}
}

try {
    Initialize-Hunspell

    # Eingebaute Liste zuerst, Datei ergaenzt sie nur.
    $script:SkipDirs = @($DefaultSkipDirs) + @(Import-FilterList (Join-Path $HunspellDir 'skip_dirs.txt'))
    $script:SkipDirs = @($script:SkipDirs | Select-Object -Unique)
    $script:SkipExts = @(Import-FilterList (Join-Path $HunspellDir 'skip_extensions.txt') | ForEach-Object { $_.ToLower() })

    if (-not $Mode) { $Mode = Select-Mode }
    if (-not $Mode) {
        Write-Host "Ungueltige Auswahl." -ForegroundColor Red
        return
    }

    $ts      = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
    $suffix  = switch ($Mode) {
        'Scan'         { 'Scan' }
        'Interactive'  { 'Interactive' }
        'ApplyFromCsv' { 'Apply' }
    }
    # Ausgabeverzeichnis mit Fallback. Vorher ging das Protokoll ungeprueft
    # in "Eigene Dokumente"; ist der Ordner umgeleitet oder gesperrt, warf
    # New-LogWriter und der Lauf endete mit "Abbruch:", bevor auch nur eine
    # Datei geprueft war.
    $logDir = Get-WritableDir @(
        [Environment]::GetFolderPath('MyDocuments'),
        (Join-Path $env:LOCALAPPDATA 'TypoFix'),
        $env:TEMP
    )
    if (-not $logDir) {
        Write-Host "Kein beschreibbares Verzeichnis fuer das Protokoll gefunden." -ForegroundColor Red
        return
    }
    $logPath   = Join-Path $logDir "11_Typo_Dateinamen_korrigieren_${suffix}_$ts.csv"
    $logWriter = New-LogWriter -path $logPath

    $stats = @{
        Umbenannt     = 0
        Konflikt      = 0
        Fehler        = 0
        Unveraendert  = 0
        Uebersprungen = 0
        ZurPruefung   = 0
        Vorschlag     = 0
        ZuLang        = 0
    }

    Write-Host ""
    Write-Host "Modus : $Mode"     -ForegroundColor Cyan
    Write-Host "Log   : $logPath"  -ForegroundColor Cyan

    if ($Mode -eq 'ApplyFromCsv') {
        if (-not $ApplyCsv) {
            $ApplyCsv = Read-Host "Pfad zur bearbeiteten CSV"
            $ApplyCsv = $ApplyCsv.Trim().Trim('"').Trim("'")
        }
        Write-Host "CSV   : $ApplyCsv" -ForegroundColor Cyan
        Write-Host ""
        $entries = Import-ApplyCsv -path $ApplyCsv
        Write-Host "Anzuwendende Eintraege: $($entries.Count)" -ForegroundColor DarkGray
        Invoke-ApplyMode -Entries $entries -LogWriter $logWriter -Stats $stats
    }
    else {
        if (-not $RootPath) { $RootPath = Select-RootPath }
        if (-not $RootPath) {
            Write-Host "Ungueltige Auswahl." -ForegroundColor Red
            return
        }
        if (-not (Test-Path -LiteralPath $RootPath)) {
            Write-Host "Pfad nicht erreichbar: $RootPath" -ForegroundColor Red
            return
        }
        Write-Host "Scanne: $RootPath" -ForegroundColor Cyan
        Write-Host ""
        Invoke-ScanOrInteractive -Root $RootPath -ProcessMode $Mode -LogWriter $logWriter -Stats $stats
    }

    Write-Host ""
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host "Zusammenfassung"                                          -ForegroundColor Cyan
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host ("  Vorschlaege    : {0}" -f $stats['Vorschlag'])
    Write-Host ("  Umbenannt      : {0}" -f $stats['Umbenannt'])     -ForegroundColor Green
    Write-Host ("  Uebersprungen  : {0}" -f $stats['Uebersprungen'])
    Write-Host ("  Zur Pruefung   : {0}" -f $stats['ZurPruefung'])   -ForegroundColor DarkYellow
    Write-Host ("  Konflikte      : {0}" -f $stats['Konflikt'])      -ForegroundColor Yellow
    Write-Host ("  Zielpfad zu lang: {0}" -f $stats['ZuLang'])       -ForegroundColor DarkYellow
    Write-Host ("  Fehler         : {0}" -f $stats['Fehler'])        -ForegroundColor Red
    Write-Host ("  Unveraendert   : {0}" -f $stats['Unveraendert'])
    Write-Host ""
    Write-Host "Log: $logPath" -ForegroundColor Cyan
    if ($Mode -eq 'Scan') {
        Write-Host ""
        Write-Host "Spalte 'Uebernehmen' in der CSV pruefen (J/N), dann Modus 3 ausfuehren." -ForegroundColor DarkCyan
    }
}
catch {
    Write-Host ""
    Write-Host "Abbruch: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($logWriter) {
        try { $logWriter.Flush(); $logWriter.Close() } catch {}
    }
    if ($script:Hunspell) {
        try { $script:Hunspell.Dispose() } catch {}
    }
    Invoke-WindowsTempCleanup -Ausnehmen @($logPath)
    try { if ($script:CtrlCAsInput) { [Console]::TreatControlCAsInput = $false } } catch {}
    if ($script:InteractiveLaunch) {
        Read-SingleKey "Beliebige Taste druecken zum Beenden"
    }
}
