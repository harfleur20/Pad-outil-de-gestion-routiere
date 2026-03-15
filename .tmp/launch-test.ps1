$ErrorActionPreference = 'Stop'
$repo = 'c:\Users\harfl\OneDrive\Desktop\projet-PAD'
$webLog = Join-Path $repo '.tmp\dev-web.log'
$electronLog = Join-Path $repo '.tmp\electron.log'
if (Test-Path $webLog) { Remove-Item $webLog -Force }
if (Test-Path $electronLog) { Remove-Item $electronLog -Force }
$web = Start-Process powershell -ArgumentList '-NoProfile','-Command',"Set-Location '$repo'; npm run dev:web *> '$webLog'" -PassThru
$ready = $false
for ($i = 0; $i -lt 60; $i++) {
  Start-Sleep -Seconds 1
  try {
    $resp = Invoke-WebRequest 'http://127.0.0.1:5173' -UseBasicParsing -TimeoutSec 2
    if ($resp.StatusCode -ge 200 -and $resp.StatusCode -lt 500) { $ready = $true; break }
  } catch {}
}
if (-not $ready) {
  cmd /c "taskkill /PID $($web.Id) /T /F" | Out-Null
  throw 'Vite ne demarre pas sur 5173.'
}
$electron = Start-Process powershell -ArgumentList '-NoProfile','-Command',"Set-Location '$repo'; node app/electron/start-electron.cjs *> '$electronLog'" -PassThru
Start-Sleep -Seconds 12
$electronProc = Get-Process electron -ErrorAction SilentlyContinue | Select-Object -First 3 Id,ProcessName,MainWindowTitle
$electronWrapper = Get-Process -Id $electron.Id -ErrorAction SilentlyContinue | Select-Object Id,ProcessName
$electronLogText = if (Test-Path $electronLog) { Get-Content $electronLog -Raw } else { '' }
$webLogTail = if (Test-Path $webLog) { (Get-Content $webLog -Tail 20) -join "`n" } else { '' }
$hasPreloadError = $false
if ($electronLogText -match 'Unable to load preload script|module not found|sandbox_bundle|preload script') { $hasPreloadError = $true }
$result = [pscustomobject]@{
  viteReady = $ready
  electronWrapperStarted = [bool]$electronWrapper
  electronProcessRunning = [bool]$electronProc
  electronProcesses = $electronProc
  hasPreloadError = $hasPreloadError
  electronLogTail = if (Test-Path $electronLog) { (Get-Content $electronLog -Tail 40) -join "`n" } else { '' }
  webLogTail = $webLogTail
}
$result | ConvertTo-Json -Depth 5
cmd /c "taskkill /PID $($electron.Id) /T /F" | Out-Null
cmd /c "taskkill /PID $($web.Id) /T /F" | Out-Null
