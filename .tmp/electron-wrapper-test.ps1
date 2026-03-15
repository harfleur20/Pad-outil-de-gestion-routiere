$ErrorActionPreference = 'Stop'
$repo = 'c:\Users\harfl\OneDrive\Desktop\projet-PAD'
$log = Join-Path $repo '.tmp\electron-wrapper.log'
if (Test-Path $log) { Remove-Item $log -Force }
$p = Start-Process powershell -ArgumentList '-NoProfile','-Command',"Set-Location '$repo'; node app/electron/start-electron.cjs *> '$log'" -PassThru
Start-Sleep -Seconds 12
$wrapper = Get-Process -Id $p.Id -ErrorAction SilentlyContinue | Select-Object Id,ProcessName
$electron = Get-Process electron -ErrorAction SilentlyContinue | Select-Object -First 5 Id,ProcessName,MainWindowTitle
$logText = if (Test-Path $log) { Get-Content $log -Raw } else { '' }
[pscustomobject]@{
  wrapperRunning = [bool]$wrapper
  electronRunning = [bool]$electron
  electronProcesses = $electron
  preloadErrorDetected = [bool]($logText -match 'Unable to load preload script|module not found|sandbox_bundle|preload script')
  logTail = if (Test-Path $log) { (Get-Content $log -Tail 40) -join "`n" } else { '' }
} | ConvertTo-Json -Depth 4
if ($wrapper) { cmd /c "taskkill /PID $($p.Id) /T /F" | Out-Null }
