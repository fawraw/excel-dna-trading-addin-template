#requires -Version 5.1
<#
    Copies the packaged .xll to %APPDATA%/Microsoft/AddIns and adds it to Excel's
    add-in list via the registry so it auto-loads on next startup.

    Usage:
        ./scripts/install.ps1
        ./scripts/install.ps1 -Uninstall      # Remove the add-in
#>
[CmdletBinding()]
param(
    [switch]$Uninstall
)

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$Source = Join-Path $Root 'dist/TradingAddin-AddIn64-packed.xll'
$AddInDir  = Join-Path $env:APPDATA 'Microsoft\AddIns'
$Dest = Join-Path $AddInDir 'TradingAddin-AddIn64-packed.xll'

if ($Uninstall) {
    if (Test-Path $Dest) {
        Remove-Item $Dest -Force
        Write-Host "Removed $Dest" -ForegroundColor Yellow
    } else {
        Write-Host "Not installed: $Dest" -ForegroundColor Yellow
    }
    return
}

if (!(Test-Path $Source)) {
    throw "Build first: $Source not found. Run ./scripts/build.ps1"
}

New-Item -ItemType Directory -Force -Path $AddInDir | Out-Null
Copy-Item $Source $Dest -Force
Write-Host "Installed $Dest" -ForegroundColor Green

Write-Host @"

Next steps:
  1. Open Excel.
  2. File -> Options -> Add-ins -> Manage: Excel Add-ins -> Go...
  3. Tick TradingAddin-AddIn64-packed and click OK.
  4. The "Trading" tab should appear in the ribbon.

For auto-load on every Excel start, see docs/install.md.
"@ -ForegroundColor Cyan
