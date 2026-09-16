function New-RoboMessage {
    param(
        [string]$Store,
        $Worksheet,
        [int]$RowNumber,
        [array]$Headers,
        [hashtable]$HeaderIndex,
        $Config,
        [int]$RecordCount
    )

    $lines = New-Object "System.Collections.Generic.List[string]"
    $title = ([string]$Config.message.title).Replace("{loja}", $Store)

    $lines.Add($title)
    $lines.Add("")

    if ([bool]$Config.message.includeRecordCount -and $RecordCount -gt 1) {
        $lines.Add(("Registros no relatorio: {0}" -f $RecordCount))
        $lines.Add("")
    }

    $fields = @($Config.message.fields)

    if ($fields.Count -eq 0) {
        $fields = @()
        $maxFields = [int]$Config.message.maxAutoFields

        foreach ($header in $Headers) {
            if ($header -ne [string]$Config.excel.storeColumn) {
                $fields += $header
            }

            if ($fields.Count -ge $maxFields) {
                break
            }
        }
    }

    foreach ($field in $fields) {
        $fieldName = [string]$field

        if ($HeaderIndex.ContainsKey($fieldName)) {
            $columnNumber = [int]$HeaderIndex[$fieldName]
            $value = [string]$Worksheet.Cells.Item($RowNumber, $columnNumber).Text

            if ([string]::IsNullOrWhiteSpace($value)) {
                $value = "-"
            }

            $lines.Add(("{0}: {1}" -f $fieldName, $value))
        }
    }

    $footer = [string]$Config.message.footer
    if (-not [string]::IsNullOrWhiteSpace($footer)) {
        $lines.Add("")
        $lines.Add($footer.Trim())
    }

    return [string]::Join([Environment]::NewLine, $lines)
}

function Send-RoboWhatsApp {
    param(
        [string]$ChromePath,
        [string]$ProfilePath,
        [string]$Phone,
        [string]$Message,
        [int]$WaitSeconds
    )

    $encoded = [Uri]::EscapeDataString($Message)
    $url = "https://web.whatsapp.com/send?phone=$Phone&text=$encoded"
    $arguments = @("--user-data-dir=$ProfilePath", "--start-maximized", $url)

    Start-Process -FilePath $ChromePath -ArgumentList $arguments | Out-Null
    Start-Sleep -Seconds $WaitSeconds

    $shell = New-Object -ComObject WScript.Shell
    $activated = $false

    for ($attempt = 0; $attempt -lt 8; $attempt++) {
        if ($shell.AppActivate("WhatsApp")) {
            $activated = $true
            break
        }
        Start-Sleep -Seconds 1
    }

    if (-not $activated) {
        throw "Janela do WhatsApp nao encontrada."
    }

    Start-Sleep -Milliseconds 500
    $shell.SendKeys("{ENTER}")
}

function Invoke-RoboExcel {
    param(
        [string]$ExcelPath,
        $Config,
        [hashtable]$StoreMap,
        [string]$ChromePath,
        [string]$ProfilePath
    )

    $excelApp = $null
    $workbook = $null
    $worksheet = $null

    try {
        $excelApp = New-Object -ComObject Excel.Application
        $excelApp.Visible = $false
        $excelApp.DisplayAlerts = $false

        $workbook = $excelApp.Workbooks.Open($ExcelPath, 0, $true)

        $sheetName = [string]$Config.excel.worksheet
        if ([string]::IsNullOrWhiteSpace($sheetName)) {
            $worksheet = $workbook.Worksheets.Item(1)
        }
        else {
            $worksheet = $workbook.Worksheets.Item($sheetName)
        }

        $usedRange = $worksheet.UsedRange
        $rowCount = [int]$usedRange.Rows.Count
        $columnCount = [int]$usedRange.Columns.Count

        if ($rowCount -lt 2) {
            throw "O Excel exportado nao possui linhas de dados."
        }

        $headers = @()
        $headerIndex = @{}

        for ($column = 1; $column -le $columnCount; $column++) {
            $header = ([string]$worksheet.Cells.Item(1, $column).Text).Trim()

            if ([string]::IsNullOrWhiteSpace($header)) {
                $header = "Coluna$column"
            }

            $headers += $header
            $headerIndex[$header] = $column
        }

        $storeColumnName = [string]$Config.excel.storeColumn
        if (-not $headerIndex.ContainsKey($storeColumnName)) {
            throw "Coluna '$storeColumnName' nao encontrada. Colunas detectadas: $([string]::Join(', ', $headers))"
        }

        $storeColumnNumber = [int]$headerIndex[$storeColumnName]
        $groups = @{}

        for ($row = 2; $row -le $rowCount; $row++) {
            $store = ConvertTo-RoboStore $worksheet.Cells.Item($row, $storeColumnNumber).Value2

            if ([string]::IsNullOrWhiteSpace($store)) {
                continue
            }

            if (-not $groups.ContainsKey($store)) {
                $groups[$store] = New-Object System.Collections.ArrayList
            }

            [void]$groups[$store].Add($row)
        }

        Write-RoboLog ("Lojas encontradas no Excel: " + $groups.Count)
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

        foreach ($store in ($groups.Keys | Sort-Object)) {
            $sourceRows = $groups[$store]
            $storeFile = Join-Path $StoreOutputDirectory ("{0}_{1}.xlsx" -f $store, $timestamp)

            $outWorkbook = $null
            $outWorksheet = $null

            try {
                $outWorkbook = $excelApp.Workbooks.Add()
                $outWorksheet = $outWorkbook.Worksheets.Item(1)

                for ($column = 1; $column -le $columnCount; $column++) {
                    $outWorksheet.Cells.Item(1, $column).Value2 = $headers[$column - 1]
                }

                $destinationRow = 2

                foreach ($sourceRow in $sourceRows) {
                    for ($column = 1; $column -le $columnCount; $column++) {
                        $outWorksheet.Cells.Item($destinationRow, $column).Value2 = $worksheet.Cells.Item($sourceRow, $column).Value2
                    }
                    $destinationRow++
                }

                $outWorkbook.SaveAs($storeFile, 51)
            }
            finally {
                if ($outWorkbook) {
                    try { $outWorkbook.Close($false) } catch {}
                }
                if ($outWorksheet) {
                    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($outWorksheet) } catch {}
                }
                if ($outWorkbook) {
                    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($outWorkbook) } catch {}
                }
            }

            if (-not $StoreMap.ContainsKey($store)) {
                Write-RoboLog "$store sem telefone cadastrado; arquivo separado criado." "AVISO"
                continue
            }

            $phone = [string]$StoreMap[$store]
            $message = New-RoboMessage $store $worksheet ([int]$sourceRows[0]) $headers $headerIndex $Config $sourceRows.Count

            $previewFile = Join-Path $PreviewDirectory ("{0}_{1}.txt" -f $store, $timestamp)
            $nl = [Environment]::NewLine
            $previewText = "Loja: $store" + $nl + "Telefone: $phone" + $nl + $nl + $message
            Set-Content -LiteralPath $previewFile -Value $previewText -Encoding UTF8

            if ([bool]$Config.whatsapp.dryRun) {
                Write-Host ""
                Write-Host ("PREVIA {0} -> {1}" -f $store, $phone) -ForegroundColor Yellow
                Write-Host $message
                Write-RoboLog "$store preparado em DRY-RUN."
            }
            else {
                Send-RoboWhatsApp $ChromePath $ProfilePath $phone $message ([int]$Config.whatsapp.waitSeconds)
                Write-RoboLog "$store enviado."
                Start-Sleep -Seconds ([int]$Config.whatsapp.delayBetweenMessagesSeconds)
            }
        }
    }
    finally {
        if ($workbook) {
            try { $workbook.Close($false) } catch {}
        }

        if ($excelApp) {
            try { $excelApp.Quit() } catch {}
        }

        if ($worksheet) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) } catch {}
        }
        if ($workbook) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } catch {}
        }
        if ($excelApp) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp) } catch {}
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
