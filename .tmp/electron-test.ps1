$ErrorActionPreference = 'Stop'
$repo = 'c:\Users\harfl\OneDrive\Desktop\projet-PAD'
$log = Join-Path $repo '.tmp\electron-direct.log'
if (Test-Path $log) { Remove-Item $log -Force }
$p = Start-Process cmd.exe -ArgumentList '/c',"cd /d $repo && node_modules\\.bin\\electron.cmd app\\electron\\main.cjs > .tmp\\electron-direct.log 2>&1" -PassThru
Start-Sleep -Seconds 10
$electronProc = Get-Process electron -ErrorAction SilentlyContinue | Select-Object -First 5 Id,ProcessName,MainWindowTitle
$cmdProc = Get-Process -Id $p.Id -ErrorAction SilentlyContinue | Select-Object Id,ProcessName
$logText = if (Test-Path $log) { Get-Content $log -Tail 80 } else { @() }
[pscustomobject]@{
  cmdWrapperStarted = [bool]$cmdProc
  electronRunning = [bool]$electronProc
  electronProcesses = $electronProc
  logTail = $logText
} | ConvertTo-Json -Depth 4
if ($cmdProc) { cmd /c "taskkill /PID $($p.Id) /T /F" | Out-Null }
