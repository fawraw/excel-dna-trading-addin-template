#requires -Version 5.1
<#
    Builds the Excel-DNA add-in and writes the packaged .xll to ./dist/.

    Usage:
        ./scripts/build.ps1                 # Release build
        ./scripts/build.ps1 -Configuration Debug
#>
[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$Proj = Join-Path $Root 'src/TradingAddin/TradingAddin.csproj'
$Dist = Join-Path $Root 'dist'

Write-Host "Restoring..." -ForegroundColor Cyan
dotnet restore $Proj

Write-Host "Building $Configuration..." -ForegroundColor Cyan
dotnet build $Proj -c $Configuration --no-restore

Write-Host "Packaging .xll..." -ForegroundColor Cyan
$OutDir = Join-Path $Root "src/TradingAddin/bin/$Configuration/net6.0-windows"
New-Item -ItemType Directory -Force -Path $Dist | Out-Null
Copy-Item "$OutDir/TradingAddin-AddIn-packed.xll"    $Dist -Force -ErrorAction SilentlyContinue
Copy-Item "$OutDir/TradingAddin-AddIn64-packed.xll"  $Dist -Force -ErrorAction SilentlyContinue

Write-Host "`nBuild complete -- packaged files:" -ForegroundColor Green
Get-ChildItem $Dist
