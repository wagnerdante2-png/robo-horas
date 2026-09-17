Write-Host ""
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host " ROBO HORAS - PORTABLE v1.3" -ForegroundColor Cyan
Write-Host " ABA UNICA REAL + TEMPLATE COM DADOS DO BI" -ForegroundColor Cyan
Write-Host " SEM INSTALACAO | SEM PYTHON | SEM ACTIONS" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""
$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root
. (Join-Path $Root "src\bootstrap.ps1")
. (Join-Path $Root "src\recipients.ps1")
. (Join-Path $Root "src\excel.ps1")
. (Join-Path $Root "src\normalize_input.ps1")
. (Join-Path $Root "src\main.ps1")
