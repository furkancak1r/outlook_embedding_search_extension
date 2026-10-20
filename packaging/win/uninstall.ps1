<#
HatirMail per-user uninstaller. Removes only what install.ps1 created.

Usage:
  powershell -NoProfile -ExecutionPolicy Bypass -File uninstall.ps1 [-InstallDir <dir>] [-KeepData] [-SkipRegistration]
  -KeepData          : keep data\ (the local search index) and settings.json
  -SkipRegistration  : do not touch the registry (for isolated test installs)
  -Machine           : also remove the per-machine COM class entries (HKLM; needs administrator rights)
The local index is deleted unless -KeepData is given. Mail in Outlook is never touched.
#>
param(
    [string]$InstallDir = (Join-Path $env:LOCALAPPDATA 'HatirMail'),
    [switch]$KeepData,
    [switch]$SkipRegistration,
    [switch]$Machine
)
$ErrorActionPreference = 'Stop'
$clsid = '{2B7C9E14-6A3F-4D85-A1E2-7F90C4B3D215}'
$progId = 'HatirMail.Connect'
if (-not $SkipRegistration) {
    if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) { throw 'Close Outlook first, then run the uninstaller again.' }
    foreach ($k in "HKCU:\Software\Microsoft\Office\Outlook\Addins\$progId", "HKCU:\Software\Classes\CLSID\$clsid",
                   "HKCU:\Software\Classes\$progId") {
        if (Test-Path $k) { Remove-Item -Path $k -Recurse -Force }
    }
    Write-Host '[HatirMail] registration removed (current user)'
    if ($Machine) {
        foreach ($view in 'Registry64', 'Registry32') {
            $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey('LocalMachine', $view)
            try { foreach ($k in "Software\Classes\CLSID\$clsid", "Software\Classes\$progId") { $base.DeleteSubKeyTree($k, $false) } }
            finally { $base.Close() }
        }
        Write-Host '[HatirMail] per-machine COM class entries removed'
    }
}
if (-not (Test-Path (Join-Path $InstallDir 'addin\HatirMail.dll'))) { throw "not a HatirMail install folder: $InstallDir" }
# Stop only HatirMail's own worker (python.exe inside this install folder), never other Python processes.
$ownPy = Join-Path $InstallDir 'runtime\python\python.exe'
Get-Process python -ErrorAction SilentlyContinue | Where-Object { $_.Path -and ($_.Path -ieq $ownPy) } | Stop-Process -Force
foreach ($d in 'app', 'addin', 'model', 'runtime', 'uninstall.ps1') {
    $p = Join-Path $InstallDir $d
    if (Test-Path $p) { Remove-Item -LiteralPath $p -Recurse -Force }
}
if (-not $KeepData) {
    foreach ($d in 'data', 'settings.json') {
        $p = Join-Path $InstallDir $d
        if (Test-Path $p) { Remove-Item -LiteralPath $p -Recurse -Force }
    }
    if (-not (Get-ChildItem -LiteralPath $InstallDir -Force)) { Remove-Item -LiteralPath $InstallDir -Force }
}
Write-Host '[HatirMail] removed' $(if ($KeepData) { '(index kept)' } else { '' })
