# Folio: ярлыки и пункт «Открыть с помощью» для PDF. Только текущий пользователь, без прав администратора.
param([switch]$Remove)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$dir = Join-Path $root 'dist\Folio-win32-x64'
$exe = Join-Path $dir 'Folio.exe'
$cls = 'HKCU:\Software\Classes'
$links = @(
  (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Folio.lnk'),
  (Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Folio.lnk')
)

if ($Remove) {
  Remove-Item "$cls\Folio.Document" -Recurse -Force -ErrorAction SilentlyContinue
  Remove-Item "$cls\Applications\Folio.exe" -Recurse -Force -ErrorAction SilentlyContinue
  Remove-ItemProperty "$cls\.pdf\OpenWithProgids" -Name 'Folio.Document' -ErrorAction SilentlyContinue
  foreach ($e in '.docx', '.xlsx', '.pptx') { Remove-ItemProperty "$cls\$e\OpenWithProgids" -Name 'Folio.Document' -ErrorAction SilentlyContinue }
  $links | ForEach-Object { Remove-Item $_ -Force -ErrorAction SilentlyContinue }
  Get-ChildItem $dir -Filter 'Folio-*.ico' -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
  Write-Host 'Ярлыки и пункт «Открыть с помощью» убраны. Папку D:\Folio можно удалить вручную.'
  return
}

if (-not (Test-Path $exe)) { Write-Host "Не найден $exe"; return }
$cmd = "`"$exe`" `"%1`""

# Значок берём из отдельного .ico, в имени которого — отпечаток его содержимого. Windows кэширует
# значки по пути файла, поэтому со старым путём (Folio.exe,0) после смены логотипа показывала бы прежний.
$icoSrc = Join-Path $root 'build\icon.ico'
$icon = "`"$exe`",0"
if (Test-Path $icoSrc) {
  $hash = (Get-FileHash $icoSrc -Algorithm SHA1).Hash.Substring(0, 8).ToLower()
  Get-ChildItem $dir -Filter 'Folio-*.ico' -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
  $ico = Join-Path $dir "Folio-$hash.ico"
  Copy-Item $icoSrc $ico -Force
  $icon = $ico
}

function Set-Key($path, $value) {
  if (-not (Test-Path $path)) { New-Item $path -Force | Out-Null }
  if ($null -ne $value) { Set-Item $path -Value $value }
}
Set-Key "$cls\Folio.Document" 'Документ'
Set-Key "$cls\Folio.Document\DefaultIcon" $icon
Set-Key "$cls\Folio.Document\shell\open\command" $cmd
Set-Key "$cls\.pdf\OpenWithProgids" $null
New-ItemProperty "$cls\.pdf\OpenWithProgids" -Name 'Folio.Document' -PropertyType None -Value ([byte[]]@()) -Force | Out-Null
foreach ($e in '.docx', '.xlsx', '.pptx') {
  Set-Key "$cls\$e\OpenWithProgids" $null
  New-ItemProperty "$cls\$e\OpenWithProgids" -Name 'Folio.Document' -PropertyType None -Value ([byte[]]@()) -Force | Out-Null
}
Set-Key "$cls\Applications\Folio.exe" $null
New-ItemProperty "$cls\Applications\Folio.exe" -Name 'FriendlyAppName' -Value 'Folio' -Force | Out-Null
Set-Key "$cls\Applications\Folio.exe\shell\open\command" $cmd
Set-Key "$cls\Applications\Folio.exe\SupportedTypes" $null
New-ItemProperty "$cls\Applications\Folio.exe\SupportedTypes" -Name '.pdf' -Value '' -Force | Out-Null
foreach ($e in '.docx', '.xlsx', '.pptx') { New-ItemProperty "$cls\Applications\Folio.exe\SupportedTypes" -Name $e -Value '' -Force | Out-Null }

$sh = New-Object -ComObject WScript.Shell
foreach ($l in $links) {
  $s = $sh.CreateShortcut($l)
  $s.TargetPath = $exe
  $s.WorkingDirectory = $dir
  $s.IconLocation = if ($icon -like '*.ico') { "$icon,0" } else { "$exe,0" }
  $s.Save()
}

# Просим Проводник перечитать значки (как после установки программы).
try {
  Add-Type -Namespace Folio -Name Shell -MemberDefinition '[DllImport("shell32.dll")] public static extern void SHChangeNotify(int e, uint f, IntPtr a, IntPtr b);'
  [Folio.Shell]::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)
} catch {}

Write-Host ''
Write-Host 'Готово. Ярлык Folio — на рабочем столе и в меню «Пуск».'
Write-Host 'Чтобы PDF, Word, Excel или PowerPoint всегда открывались в Folio: правой кнопкой по файлу → «Открыть с помощью» → «Выбрать другое приложение» → Folio → «Всегда».'
