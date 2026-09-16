try {
    if (-not (Test-Path -LiteralPath $ConfigPath)) {
        Copy-Item -LiteralPath $ConfigExamplePath -Destination $ConfigPath -Force

        if (-not (Test-Path -LiteralPath $StorePath)) {
            Copy-Item -LiteralPath $StoreExamplePath -Destination $StorePath -Force
        }

        Write-Host ""
        Write-Host "Primeira execucao: config.json e data\lojas.csv foram criados." -ForegroundColor Yellow
        Write-Host "Edite os dois arquivos e execute RoboHoras.exe novamente." -ForegroundColor Yellow
        Write-Host ""

        Start-Process -FilePath notepad.exe -ArgumentList $ConfigPath
        Start-Process -FilePath notepad.exe -ArgumentList $StorePath
        Read-Host "Pressione ENTER para fechar"
        return
    }

    $config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json

    $chromePath = Get-RoboChrome
    if (-not $chromePath) {
        throw "Google Chrome nao encontrado neste computador."
    }

    $profilePath = Join-Path $env:LOCALAPPDATA "RoboHoras\ChromeProfile"
    Ensure-RoboDirectory $profilePath

    $snapshot = Get-RoboDownloadSnapshot
    $downloadsPath = [string]$snapshot.Directory
    $before = [hashtable]$snapshot.Files

    $biUrl = [string]$config.bi.url
    if ([string]::IsNullOrWhiteSpace($biUrl)) {
        throw "Defina bi.url em config.json."
    }

    Write-RoboLog ("Abrindo BI: " + $biUrl)

    $chromeArguments = @("--user-data-dir=$profilePath", "--start-maximized", $biUrl)
    Start-Process -FilePath $chromePath -ArgumentList $chromeArguments | Out-Null

    $excelPath = Wait-RoboExcel $downloadsPath $before ([int]$config.bi.exportWaitSeconds)
    $storeMap = Get-RoboStoreMap

    Invoke-RoboExcel $excelPath $config $storeMap $chromePath $profilePath

    Write-RoboLog "Processamento concluido."
    Write-Host ""
    Write-Host "Concluido. Consulte a pasta output." -ForegroundColor Green
}
catch {
    Write-RoboLog $_.Exception.Message "ERRO"
    Write-Host ""
    Write-Host ("[ERRO] " + $_.Exception.Message) -ForegroundColor Red
    Write-Host ""
    Read-Host "Pressione ENTER para fechar"
    exit 1
}
