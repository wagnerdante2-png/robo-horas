try{
 if(-not(Test-Path $ConfigPath)){Copy-Item $ConfigExample $ConfigPath -Force;if(-not(Test-Path $StorePath)){Copy-Item $StoreExample $StorePath -Force};Write-Host '';Write-Host 'Primeira execucao: config.json e data\lojas.csv foram criados.' -ForegroundColor Yellow;Start-Process notepad.exe $ConfigPath;Start-Process notepad.exe $StorePath;Read-Host 'Edite os arquivos e pressione ENTER para fechar';return}
 $cfg=Get-Content $ConfigPath -Raw -Encoding UTF8|ConvertFrom-Json
 $chrome=FindChrome;if(-not$chrome){throw 'Google Chrome nao encontrado.'}
 $profile=Join-Path $env:LOCALAPPDATA 'RoboHoras\ChromeProfile';if(-not(Test-Path $profile)){New-Item -ItemType Directory $profile -Force|Out-Null}
 $snap=DownloadSnapshot;$downloads=$snap[0];$before=$snap[1]
 $url=[string]$cfg.bi.url;if(-not$url){throw 'Defina bi.url em config.json.'}
 Log("Abrindo BI: "+$url);Start-Process $chrome -ArgumentList("--user-data-dir=`"$profile`" --start-maximized `"$url`"")|Out-Null
 $file=WaitExcel $downloads $before ([int]$cfg.bi.exportWaitSeconds)
 $map=StoreMap
 ProcessExcel $file $cfg $map $chrome $profile
 Log 'Processamento concluido.';Write-Host '';Write-Host 'Concluido. Consulte a pasta output.' -ForegroundColor Green
}catch{Log $_.Exception.Message 'ERRO';Write-Host '';Write-Host('[ERRO] '+$_.Exception.Message) -ForegroundColor Red;Read-Host 'Pressione ENTER para fechar';exit 1}
