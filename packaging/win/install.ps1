<#
HatirMail per-user installer (no admin rights, no service, no scheduled task, no PATH change).

What it does:
  1. Copies the package to the install folder (default: %LOCALAPPDATA%\HatirMail). Existing user data under
     data\ is kept.
  2. Extracts the embedded Python and installs the worker wheels offline with pip --require-hashes.
  3. Checks every model file against hatirmail-model.json (sha256).
  4. Unless -SkipRegistration: registers the COM add-in for the CURRENT USER only (HKCU\Software\Classes and
     HKCU\Software\Microsoft\Office\Outlook\Addins). Close Outlook before you run this step.
  5. Checks that COM can create the add-in object in this process (the same check that Outlook does).

Usage:
  powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1 [-InstallDir <dir>] [-SkipRegistration] [-RegFile <path>] [-RegisterOnly]
  -SkipRegistration  : install files only (used for isolated tests; no registry write at all)
  -RegFile <path>    : write the HKCU registration as a .reg file for review instead of writing the registry
  -RegisterOnly      : skip the file steps; only (re)register an existing install
  -Machine           : ALSO register the COM class per machine (HKLM\Software\Classes, 64-bit and 32-bit views).
                       Needs administrator rights. Use it only when UAC is off and the user is an administrator:
                       then Outlook runs with a full administrator token and Windows COM ignores HKCU class entries.
  -CheckOnly         : change nothing; only report whether Outlook can create the add-in (COM activation check)
#>
param(
    [string]$InstallDir = (Join-Path $env:LOCALAPPDATA 'HatirMail'),
    [switch]$SkipRegistration,
    [string]$RegFile = '',
    [switch]$RegisterOnly,
    [switch]$Machine,
    [switch]$CheckOnly
)
$ErrorActionPreference = 'Stop'
$pkg = Split-Path -Parent $MyInvocation.MyCommand.Path
$clsid = '{2B7C9E14-6A3F-4D85-A1E2-7F90C4B3D215}'
$progId = 'HatirMail.Connect'

function Step($m) { Write-Host "[HatirMail] $m" }
function Test-ComActivation {
    # The same lookup that Outlook does: ProgID -> CLSID -> mscoree -> HatirMail.dll -> ribbon XML.
    $ok = $false
    try { $ok = (New-Object -ComObject $progId).GetCustomUI('Microsoft.Outlook.Explorer') -like '*HatirMailSearch*' } catch { }
    $admin = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
    $uacOff = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -ErrorAction SilentlyContinue).EnableLUA -eq 0
    if ($ok) {
        Step 'COM activation check: OK. Start Outlook: Home tab -> HatırMail.'
    } elseif ($uacOff -and $admin) {
        Write-Warning ('COM activation check FAILED. UAC is off and this account is an administrator: Outlook runs with a full ' +
                       'administrator token, and Windows COM ignores per-user (HKCU) class entries for such processes. The add-in ' +
                       'will not load. Options: run Outlook as a standard user (UAC on), or run this installer again as ' +
                       'administrator with -RegisterOnly -Machine (writes HKLM\Software\Classes).')
    } elseif ($admin) {
        Write-Warning "COM activation check FAILED in this elevated process. Run the check again without 'Run as administrator'."
    } else {
        Write-Warning 'COM activation check FAILED. Run the installer again; if it fails again, see KURULUM.md (Sorun giderme).'
    }
    return $ok
}
if ($CheckOnly) { [void](Test-ComActivation); return }
function Invoke-Native([string]$exe, [string[]]$argv) {
    # pip writes warnings to stderr; they must not stop the script. The exit code decides.
    $old = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
    try { & $exe @argv 2>&1 | ForEach-Object { Write-Host "  $_" } } finally { $ErrorActionPreference = $old }
    if ($LASTEXITCODE -ne 0) { throw ("{0} failed (exit code {1})" -f [IO.Path]::GetFileName($exe), $LASTEXITCODE) }
}
if (-not $RegisterOnly) {

# ---- 1. files --------------------------------------------------------------------------------------
New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
$InstallDir = (Resolve-Path $InstallDir).Path
foreach ($d in 'app', 'addin', 'model') {
    $dst = Join-Path $InstallDir $d
    if (Test-Path $dst) { Remove-Item -LiteralPath $dst -Recurse -Force }   # only our own folders
    Copy-Item -LiteralPath (Join-Path $pkg $d) -Destination $dst -Recurse
}
New-Item -ItemType Directory -Force -Path (Join-Path $InstallDir 'data') | Out-Null
Copy-Item -LiteralPath (Join-Path $pkg 'uninstall.ps1') -Destination $InstallDir -Force
if (-not (Test-Path (Join-Path $InstallDir 'settings.json'))) {
    Copy-Item -LiteralPath (Join-Path $pkg 'settings.json') -Destination $InstallDir
}
Step "files copied to $InstallDir"

# ---- 2. embedded Python + offline wheels ----------------------------------------------------------------
$py = Join-Path $InstallDir 'runtime\python'
if (Test-Path $py) { Remove-Item -LiteralPath $py -Recurse -Force }
Expand-Archive -LiteralPath (Join-Path $pkg 'runtime\python-embed-amd64.zip') -DestinationPath $py
$pth = Get-ChildItem -LiteralPath $py -Filter 'python3*._pth' | Select-Object -First 1
# Enable site-packages for the embedded interpreter (Lib\site-packages inside the install folder only).
Set-Content -LiteralPath $pth.FullName -Encoding ASCII -Value @(
    ($pth.BaseName + '.zip'), '.', 'Lib\site-packages', 'import site')
$exe = Join-Path $py 'python.exe'
$site = Join-Path $py 'Lib\site-packages'
New-Item -ItemType Directory -Force -Path $site | Out-Null
# pip itself: check the wheel hash, then unpack it (a wheel is a zip file). No network.
$pipWheel = Get-ChildItem -LiteralPath (Join-Path $pkg 'wheels') -Filter 'pip-*.whl' | Select-Object -First 1
$pipHash = ((Get-Content -LiteralPath (Join-Path $pkg 'runtime\pip.lock') -Raw) -split 'sha256:')[1].Trim()
if ((Get-FileHash -LiteralPath $pipWheel.FullName -Algorithm SHA256).Hash.ToLowerInvariant() -ne $pipHash) { throw 'pip wheel hash mismatch' }
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($pipWheel.FullName, $site)
$env:PYTHONNOUSERSITE = '1'; $env:PIP_NO_CACHE_DIR = '1'; $env:PIP_DISABLE_PIP_VERSION_CHECK = '1'
Invoke-Native $exe @('-I', '-X', 'utf8', '-m', 'pip', 'install', '--no-index', '--find-links', (Join-Path $pkg 'wheels'),
    '--require-hashes', '--target', $site, '-r', (Join-Path $pkg 'runtime\requirements-win.lock'), '--quiet')
Step 'embedded Python and worker dependencies installed (offline, hash-checked)'

# ---- 3. model files ----------------------------------------------------------------------------------
$manifest = Get-Content -LiteralPath (Join-Path $InstallDir 'model\hatirmail-model.json') -Raw | ConvertFrom-Json
foreach ($p in $manifest.files.PSObject.Properties) {
    $f = Join-Path $InstallDir ('model\' + $p.Name)
    if ($p.Name -eq 'hatirmail-model.json') { continue }
    $h = (Get-FileHash -LiteralPath $f -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($h -ne $p.Value.sha256) { throw "model file hash mismatch: $($p.Name)" }
}
Invoke-Native $exe @('-I', '-X', 'utf8', '-c', "import os, sys; os.environ['ORT_DISABLE_TELEMETRY'] = '1'; sys.path.insert(0, r'$InstallDir\app\worker'); import hm_worker, sqlite3, onnxruntime, tokenizers; print('worker import OK, sqlite', sqlite3.sqlite_version, 'onnxruntime', onnxruntime.__version__)")
Step 'model files verified (sha256)'
}

# ---- 4. per-user COM registration ------------------------------------------------------------------------
$dll = Join-Path $InstallDir 'addin\HatirMail.dll'
$codeBase = 'file:///' + ($dll -replace '\\', '/')
$asm = 'HatirMail, Version=0.1.0.0, Culture=neutral, PublicKeyToken=null'
$reg = @(
    @("Software\Classes\$progId", '', $progId),
    @("Software\Classes\$progId\CLSID", '', $clsid),
    @("Software\Classes\CLSID\$clsid", '', $progId),
    @("Software\Classes\CLSID\$clsid\ProgId", '', $progId),
    @("Software\Classes\CLSID\$clsid\InprocServer32", '', 'mscoree.dll'),
    @("Software\Classes\CLSID\$clsid\InprocServer32", 'ThreadingModel', 'Both'),
    @("Software\Classes\CLSID\$clsid\InprocServer32", 'Class', 'HatirMail.Connect'),
    @("Software\Classes\CLSID\$clsid\InprocServer32", 'Assembly', $asm),
    @("Software\Classes\CLSID\$clsid\InprocServer32", 'RuntimeVersion', 'v4.0.30319'),
    @("Software\Classes\CLSID\$clsid\InprocServer32", 'CodeBase', $codeBase),
    @("Software\Classes\CLSID\$clsid\InprocServer32\0.1.0.0", 'Class', 'HatirMail.Connect'),
    @("Software\Classes\CLSID\$clsid\InprocServer32\0.1.0.0", 'Assembly', $asm),
    @("Software\Classes\CLSID\$clsid\InprocServer32\0.1.0.0", 'RuntimeVersion', 'v4.0.30319'),
    @("Software\Classes\CLSID\$clsid\InprocServer32\0.1.0.0", 'CodeBase', $codeBase),
    @("Software\Classes\CLSID\$clsid\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}", '', ''),
    @("Software\Microsoft\Office\Outlook\Addins\$progId", 'FriendlyName', 'HatırMail'),
    @("Software\Microsoft\Office\Outlook\Addins\$progId", 'Description', 'Hatırladıklarınızla yerel e-posta araması')
)
if ($RegFile) {
    $lines = @('Windows Registry Editor Version 5.00', '')
    foreach ($g in ($reg | Group-Object { $_[0] })) {
        $lines += "[HKEY_CURRENT_USER\$($g.Name)]"
        foreach ($e in $g.Group) {
            $n = if ($e[1] -eq '') { '@' } else { '"' + $e[1] + '"' }
            $lines += ($n + '="' + ($e[2] -replace '\\', '\\') + '"')
        }
        $lines += ''
    }
    $lines += '[HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\HatirMail.Connect]'
    $lines += '"LoadBehavior"=dword:00000003'
    Set-Content -LiteralPath $RegFile -Value $lines -Encoding Unicode
    Step "registration written to $RegFile (registry not changed)"
} elseif (-not $SkipRegistration) {
    if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) { throw 'Close Outlook first, then run the installer again.' }
    foreach ($e in $reg) {
        $k = 'HKCU:\' + $e[0]
        if (-not (Test-Path $k)) { New-Item -Path $k -Force | Out-Null }
        if ($e[1] -eq '') { Set-Item -Path $k -Value $e[2] } else { New-ItemProperty -Path $k -Name $e[1] -Value $e[2] -PropertyType String -Force | Out-Null }
    }
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\Outlook\Addins\$progId" -Name LoadBehavior -Value 3 -PropertyType DWord -Force | Out-Null
    Step 'COM add-in registered for the current user.'
    $admin = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
    if ($Machine) {
        if (-not $admin) { throw '-Machine needs administrator rights.' }
        foreach ($view in 'Registry64', 'Registry32') {
            $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey('LocalMachine', $view)
            try {
                foreach ($e in ($reg | Where-Object { $_[0] -like 'Software\Classes\*' })) {
                    $k = $base.CreateSubKey($e[0])
                    try { $k.SetValue($e[1], $e[2], [Microsoft.Win32.RegistryValueKind]::String) } finally { $k.Close() }
                }
            } finally { $base.Close() }
        }
        Step 'COM class also registered per machine (HKLM\Software\Classes, 64-bit and 32-bit views).'
    }
    [void](Test-ComActivation)
} else {
    Step 'registration skipped (-SkipRegistration)'
}
Step 'done'
