# Build-Skript für TN-Doku-Konfigurator (primär in src)
# Erstellt EXE via PyInstaller und packt alles als ZIP ins release/-Verzeichnis.

param(
    [switch]$NoVersionBump,
    [switch]$SkipZip,
    [switch]$Help,
    [switch]$Quiet
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
if ($PSVersionTable.PSVersion.Major -ge 6) {
    $OutputEncoding = [System.Text.Encoding]::UTF8
}

$projectDir = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path

function Show-Usage {
    Microsoft.PowerShell.Utility\Write-Host "TN-Doku-Konfigurator Build-Skript - Optionen:" -ForegroundColor DarkCyan
    Microsoft.PowerShell.Utility\Write-Host "  -Help          : Nur diese Hilfe anzeigen und beenden" -ForegroundColor DarkGray
    Microsoft.PowerShell.Utility\Write-Host "  -NoVersionBump : Versionsnummer nicht erhöhen" -ForegroundColor DarkGray
    Microsoft.PowerShell.Utility\Write-Host "  -SkipZip       : ZIP-Erstellung überspringen" -ForegroundColor DarkGray
    Microsoft.PowerShell.Utility\Write-Host "  -Quiet         : Kompakte Ausgabe (nur Fehler + Kurzfazit)" -ForegroundColor DarkGray
}

if (-not $Quiet -or $Help) { Show-Usage }
if ($Help) { exit 0 }

$script:QuietMode = $Quiet
function Write-Host {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromRemainingArguments = $true)]
        [object[]]$Object,
        [ConsoleColor]$ForegroundColor,
        [switch]$NoNewline,
        [object]$Separator = ' '
    )
    if (-not $script:QuietMode) {
        Microsoft.PowerShell.Utility\Write-Host @PSBoundParameters
        return
    }
    $text = if ($null -eq $Object) { '' } else { ($Object -join [string]$Separator) }
    $isErrorColor = $PSBoundParameters.ContainsKey('ForegroundColor') -and $ForegroundColor -eq [ConsoleColor]::Red
    $isSummary = $text -match 'Build erfolgreich|fehlgeschlagen|^Ergebnis:|^ZIP:|^EXE:'
    if ($isErrorColor -or $isSummary) {
        Microsoft.PowerShell.Utility\Write-Host @PSBoundParameters
    }
}

function ConvertTo-MarkdownContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Content
    )

    # Auf LF normalisieren für verlässliche Verarbeitung
    $normalized = $Content -replace "`r`n", "`n" -replace "`r", "`n"
    $lines = $normalized -split "`n", -1

    $result = [System.Collections.Generic.List[string]]::new()
    $blankStreak = 0

    foreach ($line in $lines) {
        $isBlank = [string]::IsNullOrWhiteSpace($line)
        if ($isBlank) {
            $blankStreak++
            # Maximal eine Leerzeile in Folge (MD012)
            if ($blankStreak -le 1) {
                $result.Add("")
            }
            continue
        }

        $blankStreak = 0

        # Trailing Spaces entfernen, aber exakt zwei Spaces (Hard Line Break) beibehalten
        if ($line -match "  $") {
            $cleanLine = $line -replace "[\t ]{3,}$", "  "
        } else {
            $cleanLine = $line -replace "[\t ]+$", ""
        }
        $result.Add($cleanLine)
    }

    # Führende/abschließende Leerzeilen entfernen
    while ($result.Count -gt 0 -and [string]::IsNullOrWhiteSpace($result[0])) {
        $result.RemoveAt(0)
    }
    while ($result.Count -gt 0 -and [string]::IsNullOrWhiteSpace($result[$result.Count - 1])) {
        $result.RemoveAt($result.Count - 1)
    }

    $cleaned = ($result -join "`r`n")

    # Harte Absicherung gegen MD012: niemals mehr als eine Leerzeile in Folge
    $cleaned = [regex]::Replace($cleaned, '(\r?\n){3,}', "`r`n`r`n")

    # Genau einen abschließenden Zeilenumbruch sicherstellen
    return ($cleaned.TrimEnd("`r", "`n") + "`r`n")
}

function Invoke-MarkdownCleanup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectDir
    )

    Write-Host "Markdown-Bereinigung (MD-Standardfehler) ..." -ForegroundColor Yellow

    $ignoreRegex = '\\.venv\\|\\.git\\|\\dist\\|\\build\\|\\node_modules\\'
    $targets = Get-ChildItem -Path $ProjectDir -Recurse -File -Filter *.md |
        Where-Object { $_.FullName -notmatch $ignoreRegex } |
        Select-Object -ExpandProperty FullName

    $updated = 0
    $enc = [Text.UTF8Encoding]::new($false)

    $targetFiles = $targets | Sort-Object -Unique

    foreach ($file in $targetFiles) {
        try {
            $original = [IO.File]::ReadAllText($file, $enc)
            $cleaned = ConvertTo-MarkdownContent -Content $original
            if ($original -ne $cleaned) {
                [IO.File]::WriteAllText($file, $cleaned, $enc)
                $updated++
            }
        } catch {
            Write-Host "  Hinweis: Markdown-Bereinigung übersprungen für '$file' ($_)" -ForegroundColor Yellow
        }
    }

    Write-Host "  Markdown-Dateien geprüft: $($targetFiles.Count), angepasst: $updated" -ForegroundColor DarkGray
}

function New-ReleaseNotesContent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Version,
        [Parameter(Mandatory = $true)]
        [string]$ProjectDir,
        [Parameter(Mandatory = $true)]
        [string]$DistPath,
        [string]$ZipName
    )

    $today = (Get-Date).ToString('yyyy-MM-dd')
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("# Release Notes v$Version")
    $lines.Add("")
    $lines.Add("- Datum: $today")
    $lines.Add("- Version: $Version")
    $lines.Add("- Build-Ordner: $(Split-Path -Leaf $DistPath)")
    if ($ZipName) {
        $lines.Add("- ZIP-Archiv: $ZipName")
    } else {
        $lines.Add("- ZIP-Archiv: nicht erstellt (`-SkipZip)")
    }
    $lines.Add("")

    if ($ZipName) {
        $releaseUrl = "https://github.com/GoroTech-Tools/TN-Doku-Konfigurator/releases/tag/v$Version"
        $downloadUrl = "https://github.com/GoroTech-Tools/TN-Doku-Konfigurator/releases/download/v$Version/$ZipName"
        $lines.Add("## Download")
        $lines.Add("")
        $lines.Add("- [Release-Seite v$Version]($releaseUrl)")
        $lines.Add("- [ZIP direkt herunterladen]($downloadUrl)")
        $lines.Add("")
    }

    $changelogPath = Join-Path $ProjectDir 'docs\CHANGELOG.md'
    if (Test-Path $changelogPath) {
        try {
            $changelog = Get-Content $changelogPath -Raw -Encoding UTF8
            $escapedVersion = [regex]::Escape($Version)
            $pattern = "(?ms)^## \[$escapedVersion\].*?(?=^## \[|\z)"
            $match = [regex]::Match($changelog, $pattern)
            if ($match.Success) {
                $lines.Add("## Changes")
                $lines.Add("")
                $changeLines = $match.Value.Trim() -split "`r?`n"
                foreach ($line in ($changeLines | Select-Object -Skip 1)) {
                    if ($line.Trim() -eq '---') {
                        break
                    }
                    $lines.Add($line)
                }
                $lines.Add("")
                return ($lines -join "`r`n")
            }
        } catch {
            # Fallback unten
        }
    }

    $lines.Add("## Changes")
    $lines.Add("")
    $lines.Add("- Build successfully created.")
    $lines.Add("- Details: see docs/CHANGELOG.md.")
    $lines.Add("")

    return ($lines -join "`r`n")
}

# ---------------------------------------------------------------------------
# Schritt 0: Versionsnummer in build_info.py erhöhen
# ---------------------------------------------------------------------------
$buildInfoPath = Join-Path $PSScriptRoot 'build_info.py'
$newVersion = $null

if (Test-Path $buildInfoPath) {
    $content = Get-Content $buildInfoPath -Raw -Encoding UTF8
    if ($content -match "'version':\s*'([0-9]+)\.([0-9]+)\.([0-9]+)'") {
        $major = [int]$Matches[1]
        $minor = [int]$Matches[2]
        $patch = [int]$Matches[3]
        $currentVersion = "$major.$minor.$patch"

        if ($NoVersionBump) {
            $newVersion = $currentVersion
            Write-Host "Versionssprung übersprungen. Version: $newVersion" -ForegroundColor Cyan
        } else {
            $patch++
            if ($patch -ge 10) { $patch = 0; $minor++ }
            if ($minor -ge 10) { $minor = 0; $major++ }
            $newVersion = "$major.$minor.$patch"
            $newDate = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
            $content = $content -replace "'version':\s*'[0-9]+\.[0-9]+\.[0-9]+'", "'version': '$newVersion'"
            $content = $content -replace "'build_date':\s*'[^']+'", "'build_date': '$newDate'"
            $content = $content.TrimEnd() + [Environment]::NewLine
            Set-Content $buildInfoPath $content -Encoding UTF8
            Write-Host "Neue Version: $newVersion" -ForegroundColor Cyan

            # README.md: Version-Zeile aktualisieren
            $readmePath = Join-Path $projectDir 'README.md'
            if (Test-Path $readmePath) {
                $readme = Get-Content $readmePath -Raw -Encoding UTF8
                $dateForMd = (Get-Date).ToString('dd.MM.yyyy')
                $readme = [regex]::Replace(
                    $readme,
                    '\*Version: [0-9]+\.[0-9]+\.[0-9]+ \(Build: [0-9]{2}\.[0-9]{2}\.[0-9]{4}\)\*',
                    "*Version: $newVersion (Build: $dateForMd)*"
                )
                Set-Content $readmePath $readme -Encoding UTF8
                Write-Host "README.md aktualisiert." -ForegroundColor Cyan
            }

            # Weitere Docs: Versionsnummern ersetzen (currentVersion -> newVersion)
            $enc = [Text.UTF8Encoding]::new($false)
            $docFiles = @(
                'docs/DOKUMENTATION_ANWENDER.md',
                'docs/INSTALLATION.md',
                'docs/DOKUMENTATION_TECHNIK.md'
            )
            foreach ($f in $docFiles) {
                $fp = Join-Path $projectDir $f
                if (Test-Path $fp) {
                    $c = [IO.File]::ReadAllText($fp, $enc)
                    $u = $c.Replace($currentVersion, $newVersion)
                    if ($c -ne $u) {
                        [IO.File]::WriteAllText($fp, $u, $enc)
                        Write-Host "$f aktualisiert." -ForegroundColor Cyan
                    }
                }
            }
        }
    } else {
        Write-Host "Konnte Versionsnummer nicht erkennen in build_info.py!" -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host "src/build_info.py nicht gefunden!" -ForegroundColor Red
    exit 1
}

# Typische Markdown-Fehler vor dem Build bereinigen (u. a. MD009/MD012/EOF-Newline)
Invoke-MarkdownCleanup -ProjectDir $projectDir

Write-Host "`nTN-Doku-Konfigurator Build-Prozess" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Green

# ---------------------------------------------------------------------------
# Schritt 1: PyInstaller Build
# ---------------------------------------------------------------------------
$folderName = "TN-Doku-Konfigurator-v$newVersion"
$distPath   = Join-Path $projectDir "dist\$folderName"
$oneFileExePath = Join-Path $projectDir 'dist\TN-Doku-Konfigurator.exe'

# Alte Build-Artefakte ggf. entfernen
if (Test-Path $distPath) {
    Write-Host "`nEntferne alten Build-Ordner: $distPath" -ForegroundColor Yellow
    try {
        Remove-Item $distPath -Recurse -Force -ErrorAction Stop
    } catch {
        Write-Host "Konnte alten Build-Ordner nicht entfernen: $_" -ForegroundColor Red
        exit 1
    }
}
if (Test-Path $oneFileExePath) {
    Write-Host "Entferne alte Onefile-EXE: $oneFileExePath" -ForegroundColor Yellow
    try {
        Remove-Item $oneFileExePath -Force -ErrorAction Stop
    } catch {
        Write-Host "Konnte alte Onefile-EXE nicht entfernen: $_" -ForegroundColor Red
        exit 1
    }
}

Write-Host "`n1. PyInstaller Build ..." -ForegroundColor Yellow
$pyInstallerLog = Join-Path $projectDir 'build\last-pyinstaller.log'
New-Item -ItemType Directory -Path (Split-Path $pyInstallerLog -Parent) -Force | Out-Null

$specPath = Join-Path $PSScriptRoot 'TN-Doku-Konfigurator.spec'
$pythonLauncher = $null
$venvPython = Join-Path $projectDir '.venv\Scripts\python.exe'
if (Test-Path $venvPython) {
    $pythonLauncher = $venvPython
    Write-Host "Verwende Projekt-venv: $pythonLauncher" -ForegroundColor DarkGray
} else {
    $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
    if ($pythonCmd) {
        $pythonLauncher = $pythonCmd.Source
    } else {
        $pyCmd = Get-Command py -ErrorAction SilentlyContinue
        if ($pyCmd) {
            $pythonLauncher = $pyCmd.Source
        }
    }
}

if (-not $pythonLauncher) {
    Write-Host "Weder 'python' noch 'py' wurde im PATH gefunden." -ForegroundColor Red
    exit 1
}

if (Test-Path $venvPython) {
    & $pythonLauncher -c "import docx; print(docx.__file__)" *> $null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "python-docx fehlt in der Projekt-venv. Bitte 'pip install python-docx' ausführen." -ForegroundColor Red
        exit 1
    }
}

if ($Quiet) {
    & $pythonLauncher -m PyInstaller $specPath --noconfirm *> $pyInstallerLog
} else {
    & $pythonLauncher -m PyInstaller $specPath --noconfirm
}

if ($LASTEXITCODE -ne 0) {
    Write-Host "PyInstaller Build fehlgeschlagen!" -ForegroundColor Red
    if ($Quiet -and (Test-Path $pyInstallerLog)) {
        Write-Host "Details siehe: $pyInstallerLog" -ForegroundColor Yellow
        Get-Content $pyInstallerLog -Tail 30
    }
    exit 1
}

if (-not (Test-Path $oneFileExePath)) {
    Write-Host "Onefile-EXE wurde nicht erstellt: $oneFileExePath" -ForegroundColor Red
    exit 1
}

New-Item -ItemType Directory -Path $distPath -Force | Out-Null
Copy-Item -Path $oneFileExePath -Destination (Join-Path $distPath 'TN-Doku-Konfigurator.exe') -Force

Write-Host "  EXE erstellt in: $distPath" -ForegroundColor Green

# ---------------------------------------------------------------------------
# Schritt 2: Zusätzliche Dateien in dist kopieren
# ---------------------------------------------------------------------------
Write-Host "`n2. Verteilungsdateien kopieren ..." -ForegroundColor Yellow

$dataDestRoot = Join-Path $distPath 'data'
New-Item -ItemType Directory -Path $dataDestRoot -Force | Out-Null

# Ablagesystem (Template-Ordner)
$listenSrc = Join-Path $projectDir 'data\Ablagesystem'
if (Test-Path $listenSrc) {
    $listenDest = Join-Path $distPath 'data\Ablagesystem'
    Copy-Item -Path $listenSrc -Destination $listenDest -Recurse -Force
    Write-Host "  data\\Ablagesystem kopiert." -ForegroundColor DarkGray
} else {
    Write-Host "  Hinweis: data\\Ablagesystem-Ordner nicht gefunden - wird nicht mitkopiert." -ForegroundColor Yellow
}

# Teilnehmer-CSV (Beispiel/Vorlage)
$csvSrc = Join-Path $projectDir 'data\Teilnehmer_Beginn.CSV'
if (Test-Path $csvSrc) {
    Copy-Item -Path $csvSrc -Destination (Join-Path $distPath 'data\Teilnehmer_Beginn.CSV') -Force
    Write-Host "  data\\Teilnehmer_Beginn.CSV kopiert." -ForegroundColor DarkGray
}

# Dokumentation und Lizenz
foreach ($file in @('README.md')) {
    $src = Join-Path $projectDir $file
    if (Test-Path $src) {
        Copy-Item -Path $src -Destination (Join-Path $distPath $file) -Force
        Write-Host "  $file kopiert." -ForegroundColor DarkGray
    }
}

# docs/-Ordner (für integrierte Doku-Buttons in der GUI)
$docsSrc = Join-Path $projectDir 'docs'
if (-not (Test-Path $docsSrc)) {
    Write-Host "Pflichtordner fehlt: '$docsSrc'" -ForegroundColor Red
    Write-Host "Abbruch, da Release ohne Dokumentation nicht erstellt werden soll." -ForegroundColor Red
    exit 1
}

$docsDest = Join-Path $distPath 'docs'
Copy-Item -Path $docsSrc -Destination $docsDest -Recurse -Force
Write-Host "  docs/ kopiert." -ForegroundColor DarkGray

# Pflicht-Dokumente validieren
$requiredDocs = @(
    'DOKUMENTATION_ANWENDER.md',
    'DOKUMENTATION_TECHNIK.md',
    'INSTALLATION.md'
)
foreach ($doc in $requiredDocs) {
    $docPath = Join-Path $docsDest $doc
    if (-not (Test-Path $docPath)) {
        Write-Host "Pflichtdokument fehlt im Release: '$docPath'" -ForegroundColor Red
        Write-Host "Abbruch, damit kein unvollständiges Release erzeugt wird." -ForegroundColor Red
        exit 1
    }
}
Write-Host "  docs/ validiert." -ForegroundColor DarkGray

# ---------------------------------------------------------------------------
# Schritt 3: ZIP erstellen (+ Schritt 4: Release Notes erzeugen)
# ---------------------------------------------------------------------------
$releaseDir = Join-Path $projectDir 'release'
New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null

function Move-ReleaseItemToArchive {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,
        [Parameter(Mandatory = $true)]
        [string]$ArchivePath
    )

    $attempt = 0
    while ($attempt -lt 5) {
        try {
            if (Test-Path $ArchivePath) {
                Remove-Item $ArchivePath -Recurse -Force -ErrorAction SilentlyContinue
            }
            Move-Item -Path $SourcePath -Destination $ArchivePath -Force -ErrorAction Stop
            return
        } catch {
            $attempt++
            if ($attempt -ge 5) {
                Write-Host "  Warnung: Verschieben von '$SourcePath' nach '_Archiv' nach 5 Versuchen fehlgeschlagen: $_" -ForegroundColor Yellow
                return
            }
            Start-Sleep -Seconds 1
        }
    }
}

function Move-PreviousReleasesToArchive {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ReleaseDir,
        [Parameter(Mandatory = $true)]
        [string]$CurrentVersion
    )

    $archiveDir = Join-Path $ReleaseDir '_Archiv'
    New-Item -ItemType Directory -Path $archiveDir -Force | Out-Null

    $archivePattern = '^TN-Doku-Konfigurator|^RELEASE_NOTES_v'
    $items = Get-ChildItem -Path $ReleaseDir -Force | Where-Object {
        $_.Name -ne '_Archiv' -and ($_.Name -match $archivePattern)
    }

    foreach ($item in $items) {
        $currentZip = "TN-Doku-Konfigurator_v$CurrentVersion.zip"
        $currentNotes = "RELEASE_NOTES_v$CurrentVersion.md"
        if ($item.Name -eq $currentZip -or $item.Name -eq $currentNotes) {
            continue
        }

        $target = Join-Path $archiveDir $item.Name
        Move-ReleaseItemToArchive -SourcePath $item.FullName -ArchivePath $target
        if (Test-Path $target) {
            Write-Host "  Archiviert: $($item.Name) -> _Archiv" -ForegroundColor DarkGray
        }
    }
}

Move-PreviousReleasesToArchive -ReleaseDir $releaseDir -CurrentVersion $newVersion

$zipName = $null
$zipPath = $null

if (-not $SkipZip) {
    Write-Host "`n3. ZIP-Archiv erstellen ..." -ForegroundColor Yellow
    $zipName = "TN-Doku-Konfigurator_v$newVersion.zip"
    $zipPath = Join-Path $releaseDir $zipName

    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }

    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::CreateFromDirectory($distPath, $zipPath)
        $sizeMb = [math]::Round((Get-Item $zipPath).Length / 1MB, 1)
        Write-Host "ZIP: $zipName ($sizeMb MB)" -ForegroundColor Green
    } catch {
        Write-Host "System.IO.Compression fehlgeschlagen, verwende Compress-Archive-Fallback..." -ForegroundColor Yellow
        try {
            Compress-Archive -Path (Join-Path $distPath '*') -DestinationPath $zipPath -Force
            $sizeMb = [math]::Round((Get-Item $zipPath).Length / 1MB, 1)
            Write-Host "ZIP: $zipName ($sizeMb MB)" -ForegroundColor Green
        } catch {
            Write-Host "ZIP-Erstellung fehlgeschlagen: $_" -ForegroundColor Red
            exit 1
        }
    }
}

Write-Host "`n4. Release Notes erzeugen ..." -ForegroundColor Yellow
$releaseNotesName = "RELEASE_NOTES_v$newVersion.md"
$releaseNotesPath = Join-Path $releaseDir $releaseNotesName
$releaseNotesContent = New-ReleaseNotesContent -Version $newVersion -ProjectDir $projectDir -DistPath $distPath -ZipName $zipName
$releaseNotesContent = ConvertTo-MarkdownContent -Content $releaseNotesContent
$utf8NoBom = [Text.UTF8Encoding]::new($false)
[IO.File]::WriteAllText($releaseNotesPath, $releaseNotesContent, $utf8NoBom)
Write-Host "Release Notes: $releaseNotesName" -ForegroundColor Green

# ---------------------------------------------------------------------------
# Zusammenfassung
# ---------------------------------------------------------------------------
Write-Host "`nBuild erfolgreich abgeschlossen!" -ForegroundColor Green
Write-Host "Ergebnis: $distPath" -ForegroundColor Cyan
$exePath = Join-Path $distPath 'TN-Doku-Konfigurator.exe'
if (Test-Path $exePath) {
    $sizeMb = [math]::Round((Get-Item $exePath).Length / 1MB, 1)
    Write-Host "EXE: TN-Doku-Konfigurator.exe ($sizeMb MB)" -ForegroundColor Cyan
}
