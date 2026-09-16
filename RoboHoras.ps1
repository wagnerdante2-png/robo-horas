$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root
. (Join-Path $Root "src\bootstrap.ps1")
. (Join-Path $Root "src\excel.ps1")
. (Join-Path $Root "src\main.ps1")
