Write-Host ""
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host " ROBO HORAS - PORTABLE v0.8" -ForegroundColor Cyan
Write-Host " TESTE COM LISTA DE DESTINATARIOS | WHATSAPP WEB" -ForegroundColor Cyan
Write-Host " SEM INSTALACAO | SEM PYTHON | SEM ACTIONS" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""
$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root
. (Join-Path $Root "src\bootstrap.ps1")
. (Join-Path $Root "src\excel.ps1")
. (Join-Path $Root "src\normalize_input.ps1")
. (Join-Path $Root "src\main.ps1")
