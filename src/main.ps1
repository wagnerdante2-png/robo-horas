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

    Write-Host ""
    Write-Host "Escolha o modo de execucao:" -ForegroundColor Cyan
    Write-Host "  1 - Usar o Excel mais recente que ja esta em Downloads"
    Write-Host "  2 - Abrir o BI e aguardar uma nova exportacao"
    Write-Host ""
    $choice = Read-Host "Opcao [1]"
    if ([string]::IsNullOrWhiteSpace($choice)) {
        $choice = "1"
    }

    $excelPath = $null

    if ($choice -eq "1") {
        $latestExcel = Get-ChildItem -LiteralPath $downloadsPath -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -in @(".xlsx", ".xls") -and -not $_.Name.StartsWith("~$") } |
            Sort-Object LastWriteTimeUtc -Descending |
            Select-Object -First 1

        if (-not $latestExcel) {
            throw "Nenhum arquivo Excel foi encontrado em $downloadsPath."
        }

        $excelPath = $latestExcel.FullName
        Write-RoboLog ("Modo teste: usando Excel existente: " + $excelPath)
    }
    elseif ($choice -eq "2") {
        $biUrl = [string]$config.bi.url
        if ([string]::IsNullOrWhiteSpace($biUrl)) {
            throw "Defina bi.url em config.json."
        }

        Write-RoboLog ("Abrindo BI: " + $biUrl)

        $chromeArguments = @("--user-data-dir=$profilePath", "--start-maximized", $biUrl)
        Start-Process -FilePath $chromePath -ArgumentList $chromeArguments | Out-Null

        $waitSeconds = 900
        if ($config.bi.PSObject.Properties.Name -contains "exportWaitSeconds") {
            $waitSeconds = [int]$config.bi.exportWaitSeconds
        }

        $excelPath = Wait-RoboExcel $downloadsPath $before $waitSeconds
    }
    else {
        throw "Opcao invalida. Execute novamente e escolha 1 ou 2."
    }

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
