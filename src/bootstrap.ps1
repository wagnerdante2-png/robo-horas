$ConfigPath = Join-Path $Root "config.json"
$ConfigExample = Join-Path $Root "config.example.json"
$StorePath = Join-Path $Root "data\lojas.csv"
$StoreExample = Join-Path $Root "data\lojas.exemplo.csv"
$Output = Join-Path $Root "output"
$LogDir = Join-Path $Output "logs"
$PreviewDir = Join-Path $Output "previews"
$StoreDir = Join-Path $Output "lojas"
foreach ($p in @($Output,$LogDir,$PreviewDir,$StoreDir,(Join-Path $Root "data"))) { if(-not(Test-Path $p)){New-Item -ItemType Directory -Path $p -Force|Out-Null} }
$LogPath = Join-Path $LogDir ("run_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
function Log([string]$m,[string]$l="INFO"){$x="{0} | {1} | {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"),$l,$m;Add-Content $LogPath $x -Encoding UTF8;Write-Host $x}
function NormStore($v){if($null-eq$v){return ""};$t=([string]$v).Trim().ToUpper();if($t-match'^\d+([.,]0+)?$'){return("ML{0:D2}"-f[int]([double]($t-replace',','.')))};if($t-match'^ML\s*0*(\d+)$'){return("ML{0:D2}"-f[int]$Matches[1])};return($t-replace'\s+','')}
function NormPhone($v){if($null-eq$v){return ""};return(([string]$v)-replace'\D','')}
function FindChrome{$pf86=[Environment]::GetFolderPath('ProgramFilesX86');$a=@((Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe'),(Join-Path $pf86 'Google\Chrome\Application\chrome.exe'),(Join-Path $env:LOCALAPPDATA 'Google\Chrome\Application\chrome.exe'));foreach($x in $a){if($x-and(Test-Path $x)){return $x}};return $null}
function StoreMap{if(-not(Test-Path $StorePath)){throw "Cadastro de lojas ausente: $StorePath"};$h=@{};foreach($r in Import-Csv $StorePath){$active=$true;if($r.PSObject.Properties.Name-contains'ativo'){$active=@('1','true','sim','s','ativo')-contains(([string]$r.ativo).Trim().ToLower())};if(-not$active){continue};$s=NormStore $r.loja;$p=NormPhone $r.telefone;if($s-and$p){$h[$s]=$p}};if($h.Count-eq0){throw 'Nenhuma loja ativa com telefone valido.'};return $h}
function DownloadSnapshot{$d=Join-Path $env:USERPROFILE 'Downloads';if(-not(Test-Path $d)){$d=$env:USERPROFILE};$h=@{};Get-ChildItem $d -File -ErrorAction SilentlyContinue|?{$_.Extension-in@('.xlsx','.xls')-and-not$_.Name.StartsWith('~$')}|%{$h[$_.FullName]=$_.LastWriteTimeUtc.Ticks};return @($d,$h)}
function WaitExcel($d,$before,$sec){$end=(Get-Date).AddSeconds($sec);Write-Host '';Write-Host 'BI aberto. Exporte o relatorio para Excel.' -ForegroundColor Cyan;while((Get-Date)-lt$end){$f=Get-ChildItem $d -File -ErrorAction SilentlyContinue|?{$_.Extension-in@('.xlsx','.xls')-and-not$_.Name.StartsWith('~$')}|sort LastWriteTimeUtc -Descending;foreach($x in $f){if(-not$before.ContainsKey($x.FullName)-or$before[$x.FullName]-ne$x.LastWriteTimeUtc.Ticks){Start-Sleep 2;Log("Excel detectado: "+$x.FullName);return $x.FullName}};Start-Sleep 1};throw 'Tempo esgotado aguardando o Excel.'}
