param(
    [string]$WorkbookPath = (Join-Path $PSScriptRoot "..\Projet_VBA_BS.xlsm"),
    [switch]$NoBackup
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$sourceDir = Join-Path $repoRoot "src"
$resolvedWorkbook = (Resolve-Path $WorkbookPath).Path
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backupPath = [System.IO.Path]::Combine(
    [System.IO.Path]::GetDirectoryName($resolvedWorkbook),
    [System.IO.Path]::GetFileNameWithoutExtension($resolvedWorkbook) +
        ".backup-" + $timestamp + ".xlsm"
)
$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) (
    "vba-sync-" + [System.Guid]::NewGuid().ToString("N")
)

$excel = $null
$workbook = $null

function Write-AnsiCopy {
    param(
        [Parameter(Mandatory = $true)][string]$Source,
        [Parameter(Mandatory = $true)][string]$Destination
    )

    $utf8 = [System.Text.UTF8Encoding]::new($false)
    $windows1252 = [System.Text.Encoding]::GetEncoding(1252)
    $text = [System.IO.File]::ReadAllText($Source, $utf8)
    $text = $text -replace "`r?`n", "`r`n"
    [System.IO.File]::WriteAllText($Destination, $text, $windows1252)
}

function Remove-VbaComponentIfPresent {
    param(
        [Parameter(Mandatory = $true)]$Project,
        [Parameter(Mandatory = $true)][string]$Name
    )

    $component = $null
    try {
        $component = $Project.VBComponents.Item($Name)
    }
    catch {
        return
    }

    # Les modules document (Type 100) ne doivent jamais être supprimés.
    if ($component.Type -ne 100) {
        $Project.VBComponents.Remove($component)
    }
}

try {
    if (-not (Test-Path $sourceDir)) {
        throw "Le dossier src est introuvable : $sourceDir"
    }

    if (-not $NoBackup) {
        Copy-Item -LiteralPath $resolvedWorkbook -Destination $backupPath
        Write-Host "Sauvegarde créée : $backupPath"
    }

    New-Item -ItemType Directory -Path $tempDir | Out-Null

    foreach ($name in @("Classe1.cls", "Portfolio.bas", "DesignModule.bas", "frmPortfolio.frm")) {
        $sourcePath = Join-Path $sourceDir $name
        if (-not (Test-Path $sourcePath)) {
            throw "Fichier source manquant : $sourcePath"
        }
        Write-AnsiCopy -Source $sourcePath -Destination (Join-Path $tempDir $name)
    }

    Copy-Item -LiteralPath (Join-Path $sourceDir "frmPortfolio.frx") `
              -Destination (Join-Path $tempDir "frmPortfolio.frx")

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($resolvedWorkbook)

    try {
        $project = $workbook.VBProject
        $null = $project.VBComponents.Count
    }
    catch {
        throw @"
Excel bloque l'accès au projet VBA.

Dans Excel :
1. Fichier > Options > Centre de gestion de la confidentialité.
2. Paramètres du Centre > Paramètres des macros.
3. Cochez « Accès approuvé au modèle d'objet du projet VBA ».
4. Relancez ce script, puis décochez l'option après la synchronisation.
"@
    }

    foreach ($componentName in @("frmPortfolio", "DesignModule", "Portfolio", "Classe1")) {
        Remove-VbaComponentIfPresent -Project $project -Name $componentName
    }

    foreach ($fileName in @("Classe1.cls", "Portfolio.bas", "DesignModule.bas", "frmPortfolio.frm")) {
        $null = $project.VBComponents.Import((Join-Path $tempDir $fileName))
    }

    $excel.Run("'" + $workbook.Name + "'!LanceTout", $false)

    $workbook.Worksheets("Portfolio").Activate()
    $workbook.Save()
    Write-Host "Projet VBA synchronisé et classeur recalculé : $resolvedWorkbook"
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
    }
    if (Test-Path $tempDir) {
        Remove-Item -LiteralPath $tempDir -Recurse -Force
    }
    if ($null -ne $workbook) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    if ($null -ne $excel) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
