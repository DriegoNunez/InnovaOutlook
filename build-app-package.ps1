param(
    [string]$OutputPath = ".\dist\InnovaAttachmentList-M365.zip",
    [string]$PackageRoot = ".\appPackage"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

function Resolve-PathValue([string]$pathValue) {
    if ([System.IO.Path]::IsPathRooted($pathValue)) {
        return [System.IO.Path]::GetFullPath($pathValue)
    }

    return [System.IO.Path]::GetFullPath((Join-Path $repoRoot $pathValue))
}

$packageRoot = Resolve-PathValue $PackageRoot
$resolvedOutputPath = Resolve-PathValue $OutputPath
$outputDirectory = Split-Path -Parent $resolvedOutputPath

$requiredFiles = @(
    "manifest.json",
    "outline.png",
    "color.png",
    "color32x32.png"
)

foreach ($file in $requiredFiles) {
    $path = Join-Path $packageRoot $file
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Missing required package file: $path"
    }
}

New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null

if (Test-Path -LiteralPath $resolvedOutputPath) {
    Remove-Item -LiteralPath $resolvedOutputPath -Force
}

Compress-Archive -Path (Join-Path $packageRoot "*") -DestinationPath $resolvedOutputPath

Write-Host "Created app package:" $resolvedOutputPath -ForegroundColor Green
