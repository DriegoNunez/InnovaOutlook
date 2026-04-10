param(
    [string]$BaseUrl = "https://driegonunez.github.io/InnovaOutlook",
    [string]$OutputPath = ".\dist\github-pages"
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
$siteUri = [System.Uri]$BaseUrl.TrimEnd("/")
$siteUrl = $siteUri.AbsoluteUri.TrimEnd("/")
$siteHost = $siteUri.Host

if (Test-Path -LiteralPath $resolvedOutputPath) {
    Remove-Item -LiteralPath $resolvedOutputPath -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $resolvedOutputPath | Out-Null

$manifestTemplatePath = Join-Path $repoRoot "manifest.xml"
$manifestOutputPath = Join-Path $resolvedOutputPath "manifest.xml"
$xmlContent = Get-Content -LiteralPath $manifestTemplatePath -Raw
$xmlContent = $xmlContent.Replace("https://localhost:3000", $siteUrl)
Set-Content -LiteralPath $manifestOutputPath -Value $xmlContent -Encoding UTF8

$packageOutputPath = Join-Path $resolvedOutputPath "appPackage"
New-Item -ItemType Directory -Force -Path $packageOutputPath | Out-Null

$jsonTemplatePath = Join-Path $repoRoot "appPackage\manifest.json"
$jsonText = Get-Content -LiteralPath $jsonTemplatePath -Raw
$jsonText = $jsonText.Replace("https://localhost:3000", $siteUrl)
$jsonText = $jsonText.Replace('"localhost"', '"' + $siteHost + '"')

$jsonOutputPath = Join-Path $packageOutputPath "manifest.json"
Set-Content -LiteralPath $jsonOutputPath -Value $jsonText -Encoding UTF8

Copy-Item -LiteralPath (Join-Path $repoRoot "appPackage\outline.png") -Destination (Join-Path $packageOutputPath "outline.png") -Force
Copy-Item -LiteralPath (Join-Path $repoRoot "appPackage\color.png") -Destination (Join-Path $packageOutputPath "color.png") -Force
Copy-Item -LiteralPath (Join-Path $repoRoot "appPackage\color32x32.png") -Destination (Join-Path $packageOutputPath "color32x32.png") -Force

& (Join-Path $repoRoot "build-pages-site.ps1") -OutputPath (Join-Path $resolvedOutputPath "site")
& (Join-Path $repoRoot "build-app-package.ps1") `
    -PackageRoot (Join-Path $resolvedOutputPath "appPackage") `
    -OutputPath (Join-Path $resolvedOutputPath "InnovaAttachmentList-M365.zip")

Set-Content -LiteralPath (Join-Path $resolvedOutputPath "site-url.txt") -Value $siteUrl -Encoding UTF8

Write-Host "Created GitHub Pages release assets at:" $resolvedOutputPath -ForegroundColor Green
Write-Host "Base URL:" $siteUrl -ForegroundColor Cyan
