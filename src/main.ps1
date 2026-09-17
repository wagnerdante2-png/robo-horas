try {
    if (-not (Test-Path -LiteralPath $ConfigPath)) {
        Copy-Item -LiteralPath $ConfigExamplePath -Destination $ConfigPath -Force

        if (-not (Test-Path -LiteralPath $StorePath)) {
            Copy-Item -LiteralPath $StoreExamplePath -Destination $StorePath -Force
        }

        Write-Host ""
        Write-Host "Primeira execucao: config.json e data\lojas.csv foram criados." -ForegroundColor Yellow
        Write-Host "Edite os dois arquivos se desejar e execute RoboHoras.exe novamente." -ForegroundColor Yellow
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

    $testMode = $false
    if ($config.whatsapp.PSObject.Properties.Name -contains "testMode") {
        $testMode = [bool]$config.whatsapp.testMode
    }

    if ($testMode) {
        $testPhone = ""
        $selectedRecipientName = "Numero avulso"
        $recipients = @()

        if ($config.whatsapp.PSObject.Properties.Name -contains "testRecipients") {
            foreach ($recipient in @($config.whatsapp.testRecipients)) {
                $enabled = $true
                if ($recipient.PSObject.Properties.Name -contains "enabled") {
                    $enabled = [bool]$recipient.enabled
                }
                if ($enabled) {
                    $recipients += $recipient
                }
            }
        }

        Write-Host ""
        Write-Host "MODO DE TESTE DO WHATSAPP" -ForegroundColor Yellow

        if ($recipients.Count -gt 0) {
            Write-Host "Destinatarios pre-configurados:" -ForegroundColor Cyan
            for ($i = 0; $i -lt $recipients.Count; $i++) {
                $recipient = $recipients[$i]
                $name = [string]$recipient.name
                if ([string]::IsNullOrWhiteSpace($name)) { $name = "Teste $($i + 1)" }
                $phoneText = ConvertTo-RoboPhone ([string]$recipient.phone)
                if ([string]::IsNullOrWhiteSpace($phoneText)) { $phoneText = "nao configurado" }
                Write-Host ("  {0} - {1} | {2}" -f ($i + 1), $name, $phoneText)
            }
            Write-Host "  0 - Digitar outro numero"
            Write-Host ""

            $recipientChoice = Read-Host "Escolha o destinatario [1]"
            if ([string]::IsNullOrWhiteSpace($recipientChoice)) { $recipientChoice = "1" }

            $recipientIndex = 0
            if (-not [int]::TryParse($recipientChoice, [ref]$recipientIndex)) {
                throw "Opcao de destinatario invalida."
            }

            if ($recipientIndex -eq 0) {
                $testPhone = ConvertTo-RoboPhone (Read-Host "WhatsApp de teste (DDI + DDD + numero)")
            }
            elseif ($recipientIndex -ge 1 -and $recipientIndex -le $recipients.Count) {
                $selected = $recipients[$recipientIndex - 1]
                $selectedRecipientName = [string]$selected.name
                if ([string]::IsNullOrWhiteSpace($selectedRecipientName)) {
                    $selectedRecipientName = "Teste $recipientIndex"
                }
                $testPhone = ConvertTo-RoboPhone ([string]$selected.phone)

                if ([string]::IsNullOrWhiteSpace($testPhone)) {
                    Write-Host ("O numero de '{0}' ainda nao esta configurado." -f $selectedRecipientName) -ForegroundColor Yellow
                    $testPhone = ConvertTo-RoboPhone (Read-Host "Digite o numero para este teste")
                }
            }
            else {
                throw "Opcao de destinatario fora da lista."
            }
        }
        else {
            if ($config.whatsapp.PSObject.Properties.Name -contains "testPhone") {
                $testPhone = ConvertTo-RoboPhone ([string]$config.whatsapp.testPhone)
            }

            if ([string]::IsNullOrWhiteSpace($testPhone)) {
                Write-Host "Digite o numero que deve receber a unica mensagem de teste."
                Write-Host "Formato: DDI + DDD + numero, somente digitos. Ex.: 5511999999999"
                $testPhone = ConvertTo-RoboPhone (Read-Host "WhatsApp de teste")
            }
        }

        if ($testPhone.Length -lt 10) {
            throw "Telefone de teste invalido. Use DDI + DDD + numero."
        }

        $config.whatsapp.testPhone = $testPhone

        if ([bool]$config.whatsapp.dryRun) {
            Write-Host ""
            $realSend = Read-Host ("Deseja fazer UM envio real para {0} ({1}) agora? [s/N]" -f $selectedRecipientName, $testPhone)
            if ($realSend -match '^(s|sim|y|yes)$') {
                $config.whatsapp.dryRun = $false
            }
        }

        if (-not [bool]$config.whatsapp.dryRun) {
            Write-Host ""
            Write-Host "ATENCAO: sera enviada somente UMA mensagem pelo WhatsApp Web." -ForegroundColor Yellow
            Write-Host ("Destino: {0} | {1}" -f $selectedRecipientName, $testPhone) -ForegroundColor Yellow
            $confirm = Read-Host "Digite ENVIAR para confirmar"
            if ($confirm -ne "ENVIAR") {
                Write-Host "Envio cancelado. O robo continuara em simulacao." -ForegroundColor Yellow
                $config.whatsapp.dryRun = $true
            }
        }
    }

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

    $pauseOnFinish = $true
    if ($config.whatsapp.PSObject.Properties.Name -contains "pauseOnFinish") {
        $pauseOnFinish = [bool]$config.whatsapp.pauseOnFinish
    }
    if ($pauseOnFinish) {
        Write-Host ""
        Read-Host "Pressione ENTER para fechar"
    }
}
catch {
    Write-RoboLog $_.Exception.Message "ERRO"
    Write-Host ""
    Write-Host ("[ERRO] " + $_.Exception.Message) -ForegroundColor Red
    Write-Host ""
    Read-Host "Pressione ENTER para fechar"
    exit 1
}
