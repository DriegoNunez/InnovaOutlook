param(
    [switch]$SkipCertificateCheck
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectPath = Join-Path $repoRoot "host\InnovaOutlook.Host.csproj"

if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    throw "The .NET SDK is required to run the local Outlook add-in host."
}

if (-not (Test-Path -LiteralPath $projectPath)) {
    throw "Could not find the host project at $projectPath."
}

if (-not $SkipCertificateCheck) {
    dotnet dev-certs https --check | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "No trusted HTTPS development certificate was found." -ForegroundColor Yellow
        Write-Host "Run 'dotnet dev-certs https --trust' once, then rerun this script." -ForegroundColor Yellow
        exit 1
    }
}

Write-Host "Starting Innova Outlook host on https://localhost:3000" -ForegroundColor Cyan
Set-Location -LiteralPath $repoRoot
dotnet run --project $projectPath
