# Waits until Folio is closed, then installs the new build from dist\next.
# Only the app itself (resources\app.asar) changes; the Electron runtime stays the same.
$root = Split-Path -Parent $PSScriptRoot
$cur = Join-Path $root 'dist\Folio-win32-x64\resources\app.asar'
$new = Join-Path $root 'dist\next\Folio-win32-x64\resources\app.asar'
if (-not (Test-Path $new)) { 'nothing to update'; return }
while (Get-Process Folio -ErrorAction SilentlyContinue) { Start-Sleep -Seconds 2 }
$ok = $false
for ($i = 0; $i -lt 30; $i++) {
  try { Copy-Item $new $cur -Force -ErrorAction Stop; $ok = $true; break } catch { Start-Sleep -Seconds 1 }
}
if (-not $ok) { 'FAILED: file is locked'; exit 1 }
# New icon or other change in the exe itself (tools\set-exe-icon.mjs)
$newExe = Join-Path $root 'dist\next\Folio-win32-x64\Folio.exe'
if (Test-Path $newExe) {
  $curExe = Join-Path $root 'dist\Folio-win32-x64\Folio.exe'
  for ($i = 0; $i -lt 30; $i++) {
    try { Copy-Item $newExe $curExe -Force -ErrorAction Stop; break } catch { Start-Sleep -Seconds 1 }
  }
}
Remove-Item (Join-Path $root 'dist\next') -Recurse -Force -ErrorAction SilentlyContinue
"updated $(Get-Date -Format HH:mm:ss)"
