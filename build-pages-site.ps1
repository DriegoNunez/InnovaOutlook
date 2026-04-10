param(
    [string]$OutputPath = ".\dist\site"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

function Resolve-PathValue([string]$pathValue) {
    if ([System.IO.Path]::IsPathRooted($pathValue)) {
        return [System.IO.Path]::GetFullPath($pathValue)
    }

    return [System.IO.Path]::GetFullPath((Join-Path $repoRoot $pathValue))
}

$resolvedOutputPath = Resolve-PathValue $OutputPath

if (Test-Path -LiteralPath $resolvedOutputPath) {
    Remove-Item -LiteralPath $resolvedOutputPath -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $resolvedOutputPath | Out-Null

$pathsToCopy = @(
    "assets",
    "src",
    "index.html",
    "support.html",
    "privacy.html",
    "terms.html"
)

foreach ($relativePath in $pathsToCopy) {
    $sourcePath = Join-Path $repoRoot $relativePath
    $destinationPath = Join-Path $resolvedOutputPath $relativePath
    Copy-Item -LiteralPath $sourcePath -Destination $destinationPath -Recurse -Force
}

New-Item -ItemType File -Path (Join-Path $resolvedOutputPath ".nojekyll") -Force | Out-Null

Write-Host "Created GitHub Pages site at:" $resolvedOutputPath -ForegroundColor Green
