param(
  [string]$WatchPath = "C:\Users\Roberto Moura\Downloads\Banco de Dados Operacional\00 - Entrada diaria"
)

$ErrorActionPreference = "Stop"
$script = "C:\Users\Roberto Moura\Downloads\Propostas do sistema\08 - Códigos\atualizar_tudo_ao_salvar.py"

Write-Host "Monitorando: $WatchPath"
Write-Host "Quando salvar arquivo novo, a atualização será disparada."

$fsw = New-Object System.IO.FileSystemWatcher
$fsw.Path = $WatchPath
$fsw.IncludeSubdirectories = $true
$fsw.EnableRaisingEvents = $true
$fsw.Filter = "*.*"

$action = {
  $path = $Event.SourceEventArgs.FullPath
  $name = $Event.SourceEventArgs.Name
  if ($name -like "~$*") { return }
  if ($path -match "\\99 - Logs\\" -or $path -match "\\99 - Processados\\") { return }
  $ext = [System.IO.Path]::GetExtension($path).ToLowerInvariant()
  if ($ext -notin @(".csv",".xlsx",".xls")) { return }
  Write-Host "Detectado: $path"
  Start-Process -FilePath "python" -ArgumentList @($script, "--arquivo", $path) -NoNewWindow
}

Register-ObjectEvent $fsw Created -Action $action | Out-Null
Register-ObjectEvent $fsw Changed -Action $action | Out-Null

while ($true) { Start-Sleep -Seconds 2 }
