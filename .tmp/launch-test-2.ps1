$ErrorActionPreference = 'Stop'
$repo = 'c:\Users\harfl\OneDrive\Desktop\projet-PAD'
$webLog = Join-Path $repo '.tmp\dev-web-2.log'
$electronLog = Join-Path $repo '.tmp\electron-wrapper-2.log'
if (Test-Path $webLog) { Remove-Item $webLog -Force }
if (Test-Path $electronLog) { Remove-Item $electronLog -Force }
$web = Start-Process powershell -ArgumentList '-NoProfile','-Command',"Set-Location '$repo'; npm run dev:web *> '$webLog'" -PassThru
$ready = $false
for ($i = 0; $i -lt 45; $i++) {
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
$electronLogText = if (Test-Path $electronLog) { Get-Content $electronLog -Raw } else { '' }
$electronRunning = [bool](Get-Process electron -ErrorAction SilentlyContinue)
$result = [pscustomobject]@{
  viteReady = $ready
  electronRunning = $electronRunning
  preloadErrorDetected = [bool]($electronLogText -match 'Unable to load preload script|module not found|sandbox_bundle|preload script')
  refusedConnectionDetected = [bool]($electronLogText -match 'ERR_CONNECTION_REFUSED')
  electronLogTail = if (Test-Path $electronLog) { (Get-Content $electronLog -Tail 20) -join "`n" } else { '' }
  webLogTail = if (Test-Path $webLog) { (Get-Content $webLog -Tail 12) -join "`n" } else { '' }
}
$result | ConvertTo-Json -Depth 4
cmd /c "taskkill /PID $($electron.Id) /T /F" | Out-Null
cmd /c "taskkill /PID $($web.Id) /T /F" | Out-Null
