<#
.SYNOPSIS
    Syntaxprüfung für PowerShell-Skripte
.DESCRIPTION
    Parst .ps1-Dateien mit dem eingebauten PowerShell-Parser und meldet
    Syntaxfehler mit Zeile, Spalte und Textstelle. Führt nichts aus –
    die Skripte werden ausschließlich geparst.

.EXAMPLE
    .\pruefe_syntax.ps1
    Prüft alle .ps1-Dateien im Ordner des Skripts.

.EXAMPLE
    .\pruefe_syntax.ps1 -Path '.\2a_entferne_schutz_word.ps1'
    Prüft eine einzelne Datei.

.EXAMPLE
    .\pruefe_syntax.ps1 -Path 'D:\Skripte' -Recurse
    Prüft einen Ordner samt Unterordnern.
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Path,

    [switch]$Recurse
)

function Test-PowerShellSyntax {
    param([string]$FilePath)

    # WICHTIG: ParseFile erwartet zwei [ref]-Parameter. Beide Variablen
    # muessen VORHER existieren - '[ref]$null' oder ein undefiniertes
    # '[ref]$errs' brechen mit "NonExistingVariableReference" ab.
    $tokens = $null
    $errors = $null

    try {
        $null = [System.Management.Automation.Language.Parser]::ParseFile(
                    $FilePath, [ref]$tokens, [ref]$errors)
    } catch {
        return @{
            File   = $FilePath
            Ok     = $false
            Errors = @([pscustomobject]@{
                Line    = 0
                Column  = 0
                Message = "Datei nicht lesbar: $_"
                Text    = ''
            })
        }
    }

    $list = @()
    foreach ($e in @($errors)) {
        $list += [pscustomobject]@{
            Line    = $e.Extent.StartLineNumber
            Column  = $e.Extent.StartColumnNumber
            Message = $e.Message
            Text    = ($e.Extent.Text -replace '\s+', ' ')
        }
    }

    return @{
        File   = $FilePath
        Ok     = ($list.Count -eq 0)
        Errors = $list
        Tokens = @($tokens).Count
    }
}

# --- Zu prüfende Dateien ermitteln ---------------------------------
if ([string]::IsNullOrWhiteSpace($Path)) {
    $Path = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
}
$Path = $Path.Trim().Trim('"').Trim("'")

$files = @()
if (Test-Path -LiteralPath $Path -PathType Leaf) {
    $files = @(Get-Item -LiteralPath $Path)
} elseif (Test-Path -LiteralPath $Path -PathType Container) {
    $gciArgs = @{ LiteralPath = $Path; Filter = '*.ps1'; File = $true; ErrorAction = 'SilentlyContinue' }
    if ($Recurse) { $gciArgs['Recurse'] = $true }
    $files = @(Get-ChildItem @gciArgs | Sort-Object Name)
} else {
    Write-Error "Pfad nicht gefunden: $Path"
    exit 1
}

if ($files.Count -eq 0) {
    Write-Host "Keine .ps1-Dateien gefunden in: $Path" -ForegroundColor Yellow
    exit 0
}

# --- Prüfen --------------------------------------------------------
Write-Host ""
Write-Host "PowerShell-Syntaxprüfung" -ForegroundColor Cyan
Write-Host "Quelle: $Path" -ForegroundColor DarkGray
Write-Host ("-" * 70) -ForegroundColor DarkGray

$fehlerhaft = 0

foreach ($f in $files) {
    $res = Test-PowerShellSyntax -FilePath $f.FullName

    if ($res.Ok) {
        Write-Host ("  OK   {0,-45} {1,7} Token" -f $f.Name, $res.Tokens) -ForegroundColor Green
    } else {
        $fehlerhaft++
        Write-Host ("  FEHL {0,-45} {1} Fehler" -f $f.Name, $res.Errors.Count) -ForegroundColor Red
        foreach ($e in $res.Errors) {
            Write-Host ("         Zeile {0}, Spalte {1}: {2}" -f $e.Line, $e.Column, $e.Message) -ForegroundColor Yellow
            if ($e.Text) {
                $snippet = if ($e.Text.Length -gt 90) { $e.Text.Substring(0, 90) + '...' } else { $e.Text }
                Write-Host ("         > {0}" -f $snippet) -ForegroundColor DarkGray
            }
        }
    }
}

Write-Host ("-" * 70) -ForegroundColor DarkGray
if ($fehlerhaft -eq 0) {
    Write-Host "$($files.Count) Datei(en) geprüft - keine Syntaxfehler." -ForegroundColor Green
} else {
    Write-Host "$($files.Count) Datei(en) geprüft - $fehlerhaft mit Syntaxfehlern." -ForegroundColor Red
}
Write-Host ""

exit ([Math]::Min($fehlerhaft, 1))
