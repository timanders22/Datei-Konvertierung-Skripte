<#
.SYNOPSIS
    Sammel-Prüfung über beide Sprachen: PowerShell und Python.

.DESCRIPTION
    Ruft nacheinander auf:
      1. pruefe_syntax.ps1        - Parser-Lauf über alle .ps1
      2. python -m py_compile     - Übersetzungslauf über alle .py
      3. PSScriptAnalyzer         - falls installiert (optional)
      4. ruff                     - falls installiert (optional)
      5. Eigene Musterprüfungen   - siehe unten

    Warum zusätzlich zu pruefe_syntax.ps1: Der reine Parser-Lauf hätte den
    schwerwiegendsten Fehler der Durchsicht NICHT gefunden. In

        $p -replace [regex]::Escape($ext) + '$', $ziel

    ist syntaktisch nichts falsch - der Komma-Operator bindet in
    PowerShell nur stärker als '+', wodurch ein einarmiges -replace
    entsteht, das nichts ersetzt. Solche Fallen findet erst eine
    Musterprüfung. Die eingebauten Prüfungen decken genau die Klassen ab,
    die in dieser Sammlung schon einmal aufgetreten sind.

.EXAMPLE
    .\pruefe_alles.ps1

.EXAMPLE
    .\pruefe_alles.ps1 -Ausfuehrlich

.NOTES
    Rückgabewert 0 = alles sauber, 1 = mindestens ein Befund.
    Stand: 14.08.2026
#>

[CmdletBinding()]
param(
    [string]$Pfad,
    [switch]$Ausfuehrlich
)

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

if ([string]::IsNullOrWhiteSpace($Pfad)) {
    $Pfad = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
}
$Pfad = $Pfad.Trim().Trim('"').Trim("'")

if (-not (Test-Path -LiteralPath $Pfad -PathType Container)) {
    Write-Host "Verzeichnis nicht gefunden: $Pfad" -ForegroundColor Red
    exit 1
}

$befunde = 0

function Write-Abschnitt {
    param([string]$Titel)
    Write-Host ""
    Write-Host ("=" * 74) -ForegroundColor Cyan
    Write-Host "  $Titel" -ForegroundColor Cyan
    Write-Host ("=" * 74) -ForegroundColor Cyan
}

# ==================================================================
# 1. PowerShell-Syntax
# ==================================================================
Write-Abschnitt "1/5  PowerShell-Syntax"
$syntaxSkript = Join-Path $Pfad 'pruefe_syntax.ps1'
if (Test-Path -LiteralPath $syntaxSkript) {
    & $syntaxSkript -Path $Pfad
    if ($LASTEXITCODE -ne 0) { $befunde++ }
} else {
    Write-Host "  pruefe_syntax.ps1 nicht gefunden - übersprungen." -ForegroundColor DarkGray
}

# ==================================================================
# 2. Python-Übersetzung
# ==================================================================
Write-Abschnitt "2/5  Python-Übersetzung (py_compile)"
$python = $null
foreach ($kandidat in @('python', 'py')) {
    try {
        $null = & $kandidat --version 2>&1
        if ($LASTEXITCODE -eq 0) { $python = $kandidat; break }
    } catch {}
}

if (-not $python) {
    Write-Host "  Kein Python im PATH gefunden - übersprungen." -ForegroundColor DarkGray
} else {
    $pyDateien = @(Get-ChildItem -LiteralPath $Pfad -Filter '*.py' -File | Sort-Object Name)
    $pyFehler = 0
    foreach ($datei in $pyDateien) {
        $ausgabe = & $python -m py_compile $datei.FullName 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host ("  OK   {0}" -f $datei.Name) -ForegroundColor Green
        } else {
            $pyFehler++
            Write-Host ("  FEHL {0}" -f $datei.Name) -ForegroundColor Red
            $ausgabe | ForEach-Object { Write-Host "         $_" -ForegroundColor Yellow }
        }
    }
    Write-Host ("  {0} Datei(en) geprüft, {1} mit Fehlern." -f $pyDateien.Count, $pyFehler) `
        -ForegroundColor $(if ($pyFehler) { 'Red' } else { 'Green' })
    if ($pyFehler) { $befunde++ }
}

# ==================================================================
# 3. PSScriptAnalyzer (optional)
# ==================================================================
Write-Abschnitt "3/5  PSScriptAnalyzer (optional)"
if (Get-Module -ListAvailable -Name PSScriptAnalyzer) {
    Import-Module PSScriptAnalyzer -ErrorAction SilentlyContinue
    $regeln = @(
        'PSAvoidUsingCmdletAliases',
        'PSUseDeclaredVarsMoreThanAssignments',
        'PSAvoidUsingPositionalParameters',
        'PSPossibleIncorrectComparisonWithNull',
        'PSAvoidGlobalVars'
    )
    $treffer = @(Invoke-ScriptAnalyzer -Path $Pfad -IncludeRule $regeln -ErrorAction SilentlyContinue)
    if ($treffer.Count -eq 0) {
        Write-Host "  Keine Befunde." -ForegroundColor Green
    } else {
        $treffer | Group-Object RuleName | Sort-Object Count -Descending | ForEach-Object {
            Write-Host ("  {0,-45} {1,4}" -f $_.Name, $_.Count) -ForegroundColor Yellow
            if ($Ausfuehrlich) {
                $_.Group | ForEach-Object {
                    Write-Host ("      {0}:{1}" -f (Split-Path $_.ScriptPath -Leaf), $_.Line) -ForegroundColor DarkGray
                }
            }
        }
        Write-Host "  (Hinweise, kein Abbruchgrund. Mit -Ausfuehrlich für Fundstellen.)" -ForegroundColor DarkGray
    }
} else {
    Write-Host "  PSScriptAnalyzer nicht installiert - übersprungen." -ForegroundColor DarkGray
    Write-Host "  Installation:  Install-Module PSScriptAnalyzer -Scope CurrentUser" -ForegroundColor DarkGray
}

# ==================================================================
# 4. ruff (optional)
# ==================================================================
Write-Abschnitt "4/5  ruff (optional)"
$ruff = Get-Command ruff -ErrorAction SilentlyContinue
if ($ruff) {
    & ruff check $Pfad --quiet
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  Keine Befunde." -ForegroundColor Green
    } else {
        Write-Host "  (Hinweise, kein Abbruchgrund.)" -ForegroundColor DarkGray
    }
} else {
    Write-Host "  ruff nicht installiert - übersprungen." -ForegroundColor DarkGray
    Write-Host "  Installation:  pip install ruff" -ForegroundColor DarkGray
}

# ==================================================================
# 5. Eigene Musterprüfungen
# ==================================================================
Write-Abschnitt "5/5  Musterprüfungen (Fallen aus dieser Sammlung)"

$musterTreffer = 0

function Test-Muster {
    param(
        [string]$Name,
        [string]$Erklaerung,
        [string]$Dateimuster,
        [scriptblock]$Pruefung
    )
    $gefunden = @()
    foreach ($datei in (Get-ChildItem -LiteralPath $script:Pfad -Filter $Dateimuster -File)) {
        $zeilen = [System.IO.File]::ReadAllLines($datei.FullName)
        # Blockkommentare <# ... #> ueberspringen. Ohne das meldet die
        # Pruefung ihre eigenen Beispiele im Hilfetext als Befund.
        $imBlock = $false
        for ($i = 0; $i -lt $zeilen.Count; $i++) {
            $zeile = $zeilen[$i]
            if ($imBlock) {
                if ($zeile -match '#>') { $imBlock = $false }
                continue
            }
            if ($zeile -match '<#') {
                if ($zeile -notmatch '#>') { $imBlock = $true }
                continue
            }
            if (& $Pruefung $zeilen[$i]) {
                $gefunden += [pscustomobject]@{
                    Datei = $datei.Name; Zeile = $i + 1; Text = $zeilen[$i].Trim()
                }
            }
        }
    }
    if ($gefunden.Count -eq 0) {
        Write-Host ("  OK   {0}" -f $Name) -ForegroundColor Green
    } else {
        Write-Host ("  FUND {0}  ({1})" -f $Name, $gefunden.Count) -ForegroundColor Red
        Write-Host ("       {0}" -f $Erklaerung) -ForegroundColor DarkGray
        foreach ($g in $gefunden) {
            $text = if ($g.Text.Length -gt 84) { $g.Text.Substring(0, 84) + '...' } else { $g.Text }
            Write-Host ("       {0}:{1}" -f $g.Datei, $g.Zeile) -ForegroundColor Yellow
            Write-Host ("         > {0}" -f $text) -ForegroundColor DarkGray
        }
        $script:musterTreffer += $gefunden.Count
    }
}

$script:Pfad = $Pfad

# --- 5.1  -replace mit ungeklammerter Verkettung ------------------
# Der Fehler, der 7_Dateien_ohne_Makro_finden.ps1 das Gegenteil seiner
# Aufgabe tun liess. Syntaktisch einwandfrei, deshalb vom Parser nicht
# zu finden.
Test-Muster -Name '-replace mit ungeklammerter Verkettung' `
    -Erklaerung "Der Komma-Operator bindet stärker als '+'. Muster und Ersatz klammern!" `
    -Dateimuster '*.ps1' `
    -Pruefung {
        param($z)
        if ($z -match '^\s*#') { return $false }
        if ($z -notmatch '-replace') { return $false }
        # Zeichenketten-Literale entfernen, bevor nach '+' gesucht wird.
        # Sonst schlägt jedes Regex-Literal mit Quantor an ('\d+', '\s+'),
        # und die Prüfung wäre wegen Fehlalarmen wertlos.
        $ohneLiterale = $z -replace "'[^']*'", "''" -replace '"[^"]*"', '""'
        # Jetzt zählt nur noch ein '+' zwischen '-replace' und dem Komma,
        # das NICHT von einer Klammer eingefasst ist.
        if ($ohneLiterale -notmatch '-replace\s+([^,]*),') { return $false }
        $argument = $Matches[1]
        if ($argument -match '^\s*\(') { return $false }   # korrekt geklammert
        return ($argument -match '\+')
    }

# --- 5.2  Provider-Cmdlets auf Langpfad-Variablen -----------------
Test-Muster -Name 'Provider-Cmdlets auf Langpfad-Variablen' `
    -Erklaerung "'\\?\'-Pfade in Provider-Cmdlets (Test-Path, New-Item, Move-Item, Remove-Item ...) sind unter 5.1 nur lokal belegt, nicht für UNC/G: - [System.IO] verwenden." `
    -Dateimuster '*.ps1' `
    -Pruefung {
        param($z)
        # Vorher nur Variablen, die auf 'Long'/'LongPath' ENDEN ($srcLong):
        # '$longTarget', '$longSourceTest', '$longFinalDest', '$longTestFile'
        # in 8_verschieben_auf_Google_Drive.ps1 fielen durch, ebenso
        # Move-Item und New-Item. Jetzt zusaetzlich 'long'/'Long' als
        # Wortanfang, gefolgt von einem Grossbuchstaben ($longTarget,
        # $LongLogFile) - dieser Teil ist schreibweisengenau, damit etwa
        # '$longest' nicht anschlaegt.
        # Bewusst KEINE Ausnahme fuer die Fundstellen in 8: gemessen
        # (29.09., PS 5.1.26100) arbeiten Test-Path/New-Item/Move-Item/
        # Remove-Item mit '\\?\' lokal auf C: korrekt, auch bei 430
        # Zeichen - aber nur mit LongPathsEnabled=1; UNC ('\\?\UNC\'),
        # Netzlaufwerke und das Drive-Laufwerk G: sind nicht nachgemessen.
        # Mehrzeilige Aufrufe (Cmdlet und Variable auf zwei Zeilen) sieht
        # die zeilenweise Pruefung weiterhin nicht.
        ($z -cmatch '(?i:\b(?:Test-Path|Remove-Item|Rename-Item|Copy-Item|Move-Item|New-Item|Get-Item))\b[^#]*\$(?:[Ll]ong[A-Z0-9_]\w*|\w*(?i:Long)(?:Path)?\b)') -and ($z -notmatch '^\s*#')
    }

# --- 5.3  Python: Ausgabe von Sonderzeichen ohne UTF-8-Umstellung --
$pyOhneEncoding = @()
foreach ($datei in (Get-ChildItem -LiteralPath $Pfad -Filter '*.py' -File)) {
    $inhalt = [System.IO.File]::ReadAllText($datei.FullName)
    $hatUmstellung = $inhalt -match 'reconfigure\s*\(\s*encoding'
    $hatSonderzeichen = $false
    foreach ($zeile in ($inhalt -split "`n")) {
        if ($zeile -notmatch 'print\(') { continue }
        foreach ($zeichen in $zeile.ToCharArray()) {
            if ([int]$zeichen -gt 255) { $hatSonderzeichen = $true; break }
        }
        if ($hatSonderzeichen) { break }
    }
    if ($hatSonderzeichen -and -not $hatUmstellung) {
        $pyOhneEncoding += $datei.Name
    }
}
if ($pyOhneEncoding.Count -eq 0) {
    Write-Host "  OK   Python-Konsolen-Encoding" -ForegroundColor Green
} else {
    Write-Host ("  FUND Python-Konsolen-Encoding  ({0})" -f $pyOhneEncoding.Count) -ForegroundColor Red
    Write-Host "       Gibt Sonderzeichen aus, stellt stdout aber nicht auf UTF-8 -> UnicodeEncodeError." -ForegroundColor DarkGray
    $pyOhneEncoding | ForEach-Object { Write-Host "       $_" -ForegroundColor Yellow }
    $musterTreffer += $pyOhneEncoding.Count
}

# --- 5.4  Python: verschachtelte Anführungszeichen in f-Strings ----
$fstringTreffer = @()
foreach ($datei in (Get-ChildItem -LiteralPath $Pfad -Filter '*.py' -File)) {
    $zeilen = [System.IO.File]::ReadAllLines($datei.FullName)
    for ($i = 0; $i -lt $zeilen.Count; $i++) {
        if ($zeilen[$i] -match 'f"[^"]*\{[^}]*"[^}]*\}') {
            $fstringTreffer += ("{0}:{1}" -f $datei.Name, ($i + 1))
        }
    }
}
if ($fstringTreffer.Count -eq 0) {
    Write-Host "  OK   f-Strings ohne verschachtelte Anführungszeichen" -ForegroundColor Green
} else {
    Write-Host ("  FUND Verschachtelte Anführungszeichen in f-Strings  ({0})" -f $fstringTreffer.Count) -ForegroundColor Red
    Write-Host "       Erst ab Python 3.12 zulässig (PEP 701)." -ForegroundColor DarkGray
    $fstringTreffer | ForEach-Object { Write-Host "       $_" -ForegroundColor Yellow }
    $musterTreffer += $fstringTreffer.Count
}

# --- 5.5  Variable benutzt, bevor sie zugewiesen wird -------------
# Diese Prüfung hat den Fehler in 2c gefunden, den weder Parser noch
# Musterprüfung sehen konnten: '$srcLong' wurde in der Junk-Erkennung
# gelesen, aber erst 25 Zeilen später zugewiesen. In PowerShell ist das
# still $null - der Aufruf warf, der leere catch-Block schluckte es, und
# das Aufräumen verwaister Sperrdateien fand nie statt.
#
# Herausgefiltert werden die drei Klassen, die harmlos sind:
#   - $script:/$global:-Variablen (auf Skriptebene initialisiert)
#   - foreach-Laufvariablen
#   - param()-Variablen (auch in Start-Job-Blöcken)
$uninitTreffer = @()
foreach ($datei in (Get-ChildItem -LiteralPath $Pfad -Filter '*.ps1' -File)) {
    $tk = $null; $er = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($datei.FullName, [ref]$tk, [ref]$er)
    if ($er -and $er.Count) { continue }

    foreach ($fn in $ast.FindAll({ param($n)
            $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {

        # Namen, die nicht als "unzugewiesen" gelten dürfen
        $sicher = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)
        foreach ($n in @('_','null','true','false','PSItem','args','PSCmdlet','PSBoundParameters',
                         'MyInvocation','PSScriptRoot','Host','Matches','LASTEXITCODE','PWD','Error')) {
            [void]$sicher.Add($n)
        }
        foreach ($p in $fn.Body.FindAll({ param($n)
                $n -is [System.Management.Automation.Language.ParameterAst] }, $true)) {
            [void]$sicher.Add($p.Name.VariablePath.UserPath)
        }
        foreach ($fe in $fn.Body.FindAll({ param($n)
                $n -is [System.Management.Automation.Language.ForEachStatementAst] }, $true)) {
            [void]$sicher.Add($fe.Variable.VariablePath.UserPath)
        }

        $ersteZuweisung = @{}
        foreach ($as in $fn.Body.FindAll({ param($n)
                $n -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)) {
            # Linke Seite auspacken: bei '[long]$x = 0' steht dort ein
            # ConvertExpressionAst (der Typ-Cast), nicht die Variable
            # selbst. Ohne dieses Auspacken gilt die Zeile nicht als
            # Zuweisung - und die Prüfung meldet ihren eigenen Fehlalarm.
            $links = $as.Left
            if ($links -is [System.Management.Automation.Language.ConvertExpressionAst]) {
                $links = $links.Child
            }
            if ($links -is [System.Management.Automation.Language.VariableExpressionAst]) {
                $nm = $links.VariablePath.UserPath
                if (-not $ersteZuweisung.ContainsKey($nm)) {
                    $ersteZuweisung[$nm] = $as.Extent.StartLineNumber
                }
            }
        }

        $gemeldet = @{}
        foreach ($ve in $fn.Body.FindAll({ param($n)
                $n -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)) {
            $nm = $ve.VariablePath.UserPath
            if ($sicher.Contains($nm))                       { continue }
            if ($ve.VariablePath.IsGlobal -or $ve.VariablePath.IsScript) { continue }
            if ($nm -like '*:*')                             { continue }
            if (-not $ersteZuweisung.ContainsKey($nm))       { continue }
            if ($ve.Extent.StartLineNumber -ge $ersteZuweisung[$nm]) { continue }
            $schluessel = "$($datei.Name)|$($fn.Name)|$nm"
            if ($gemeldet.ContainsKey($schluessel))          { continue }
            $gemeldet[$schluessel] = $true
            $uninitTreffer += [pscustomobject]@{
                Datei = $datei.Name; Funktion = $fn.Name; Variable = $nm
                Zeile = $ve.Extent.StartLineNumber; Zuweisung = $ersteZuweisung[$nm]
            }
        }
    }
}
if ($uninitTreffer.Count -eq 0) {
    Write-Host "  OK   Keine Variable vor ihrer Zuweisung benutzt" -ForegroundColor Green
} else {
    Write-Host ("  FUND Variable vor Zuweisung benutzt  ({0})" -f $uninitTreffer.Count) -ForegroundColor Red
    Write-Host "       In PowerShell still `$null - der Fehler zeigt sich erst im Verhalten." -ForegroundColor DarkGray
    foreach ($t in $uninitTreffer) {
        Write-Host ("       {0}  {1}():  `${2} in Z.{3}, zugewiesen erst in Z.{4}" -f `
            $t.Datei, $t.Funktion, $t.Variable, $t.Zeile, $t.Zuweisung) -ForegroundColor Yellow
    }
    $musterTreffer += $uninitTreffer.Count
}

if ($musterTreffer -gt 0) { $befunde++ }

# ==================================================================
# Abschluss
# ==================================================================
Write-Host ""
Write-Host ("=" * 74) -ForegroundColor Cyan
if ($befunde -eq 0) {
    Write-Host "  Alle Prüfungen ohne Befund." -ForegroundColor Green
} else {
    Write-Host "  $befunde Prüfabschnitt(e) mit Befunden - bitte oben nachsehen." -ForegroundColor Red
}
Write-Host ("=" * 74) -ForegroundColor Cyan
Write-Host ""

exit ([Math]::Min($befunde, 1))
