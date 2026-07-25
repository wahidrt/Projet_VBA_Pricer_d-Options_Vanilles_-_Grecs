param(
    [string]$WorkbookPath = (Join-Path $PSScriptRoot "..\Projet_VBA_BS.xlsm")
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$sourceDir = Join-Path $repoRoot "src"
$resolvedWorkbook = (Resolve-Path $WorkbookPath).Path
$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) (
    "vba-export-" + [System.Guid]::NewGuid().ToString("N")
)

$excel = $null
$workbook = $null

function Convert-AnsiFileToUtf8 {
    param(
        [Parameter(Mandatory = $true)][string]$Source,
        [Parameter(Mandatory = $true)][string]$Destination
    )

    $windows1252 = [System.Text.Encoding]::GetEncoding(1252)
    $utf8 = [System.Text.UTF8Encoding]::new($false)
    $text = [System.IO.File]::ReadAllText($Source, $windows1252)
    [System.IO.File]::WriteAllText($Destination, $text, $utf8)
}

try {
    New-Item -ItemType Directory -Force -Path $sourceDir | Out-Null
    New-Item -ItemType Directory -Path $tempDir | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($resolvedWorkbook, $null, $true)

    try {
        $project = $workbook.VBProject
        $null = $project.VBComponents.Count
    }
    catch {
        throw "Activez temporairement « Accès approuvé au modèle d'objet du projet VBA » dans Excel."
    }

    foreach ($component in $project.VBComponents) {
        $extension = switch ($component.Type) {
            1 { ".bas" }
            2 { ".cls" }
            3 { ".frm" }
            default { $null }
        }

        if ($null -eq $extension) {
            continue
        }

        $tempPath = Join-Path $tempDir ($component.Name + $extension)
        $component.Export($tempPath)
        Convert-AnsiFileToUtf8 -Source $tempPath `
                              -Destination (Join-Path $sourceDir ($component.Name + $extension))

        if ($extension -eq ".frm") {
            $frxPath = [System.IO.Path]::ChangeExtension($tempPath, ".frx")
            if (Test-Path $frxPath) {
                Copy-Item -LiteralPath $frxPath `
                          -Destination (Join-Path $sourceDir ($component.Name + ".frx")) `
                          -Force
            }
        }
    }

    Write-Host "Sources VBA exportées en UTF-8 vers : $sourceDir"
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
