<#
HatirMail firewall rule (Windows Defender Firewall): block all outbound connections of the HatirMail worker, the
embedded python.exe of the install folder. The add-in and the worker talk through pipes, not through the network:
HatirMail works the same with the rule. Outlook is not affected (the add-in runs inside Outlook).

Usage:
  powershell -NoProfile -ExecutionPolicy Bypass -File firewall.ps1 [-InstallDir <dir>] [-Remove] [-ParentWindow <hwnd>]
The setup program runs it after install.ps1; uninstall.ps1 runs it with -Remove. A second run with a rule present
changes nothing (an update asks for no approval).
A rule change needs administrator rights. Without a full administrator token, the script starts itself again with
"Run as administrator" (Windows asks: UAC). -ParentWindow: the setup window; then Windows shows the prompt in front
of it. The rule names the full path of python.exe: the rule of another install does not change.
After the change, python.exe tries one connection as the current user to 192.0.2.1 (TEST-NET-1, RFC 5737: a
documentation address that no host uses). Windows refuses a blocked connection at once with WSAEACCES (10013).
ponytail: any other result counts as "not verified" (firewall off, a policy that ignores local rules, a third-party
firewall, or no network route).

Exit codes: 0 = the rule is present and Windows refused the test connection (-Remove: no rule is left); 1 = no rule
(no administrator approval, or an error; -Remove: the rule is still present); 2 = the rule is present, but the block
was not verified.
This file has a UTF-8 BOM: Windows PowerShell 5.1 reads a file without a BOM in the ANSI code page.
#>
param([string]$InstallDir = $PSScriptRoot, [switch]$Remove, [long]$ParentWindow = 0, [switch]$Elevated)
$ErrorActionPreference = 'Stop'
$program = Join-Path $InstallDir 'runtime\python\python.exe'
$group = 'HatirMail'
function Say($m) { Write-Host "[HatirMail] $m" }
# A function unrolls its array output: one rule comes back as one CimInstance, and PowerShell 5.1 gives no .Count for
# it (2.19, 10-22: the rule was there, the count was empty). Callers count with @(Get-OwnRule).Count.
function Get-OwnRule {
    Get-NetFirewallRule -Group $group -ErrorAction SilentlyContinue |
      Where-Object { (Get-NetFirewallApplicationFilter -AssociatedNetFirewallRule $_).Program -eq $program }
}
function Set-Rule {
    # One rule at a time, by its unique name. Never pipe a list that can be empty into Remove-NetFirewallRule.
    foreach ($r in Get-OwnRule) { Remove-NetFirewallRule -Name $r.Name }
    if (-not $Remove) {
        New-NetFirewallRule -DisplayName 'HatırMail: arama işlemi internete çıkamaz' -Group $group -Direction Outbound `
            -Action Block -Program $program -Description 'HatırMail kurulumu ekledi. HatırMail kaldırılınca silinir.' | Out-Null
    }
}
if ($Elevated) { Set-Rule; exit 0 }

if ((@(Get-OwnRule).Count -gt 0) -eq [bool]$Remove) {   # a change is necessary
    try {
        $admin = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole(
            [Security.Principal.WindowsBuiltInRole]::Administrator)
        if ($admin) { Set-Rule }   # full token (UAC off, or "Run as administrator"): no prompt
        else {
            $argv = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" -InstallDir "{1}" -Elevated{2}' -f $PSCommandPath,
                $InstallDir, $(if ($Remove) { ' -Remove' } else { '' })
            $psi = New-Object Diagnostics.ProcessStartInfo((Join-Path $PSHOME 'powershell.exe'), $argv)
            $psi.Verb = 'runas'; $psi.UseShellExecute = $true; $psi.WindowStyle = 'Hidden'
            if ($ParentWindow) { $psi.ErrorDialog = $true; $psi.ErrorDialogParentHandle = [IntPtr]$ParentWindow }
            $p = [Diagnostics.Process]::Start($psi)   # throws when the user says No
            $p.WaitForExit()
            if ($p.ExitCode -ne 0) { throw "elevated firewall.ps1: exit code $($p.ExitCode)" }
        }
    } catch {
        Say "firewall rule not changed: $($_.Exception.Message)"
    }
}
$has = @(Get-OwnRule).Count -gt 0
if ($Remove) {
    if ($has) { Say 'firewall rule still present (it names only the removed python.exe)'; exit 1 }
    Say 'no firewall rule left'
    exit 0
}
if (-not $has) { Say 'no firewall rule: HatirMail works, but Windows does not block its network access'; exit 1 }
$code = & $program -I -c "import socket; s = socket.socket(); s.settimeout(3); print(s.connect_ex(('192.0.2.1', 443)))"
if ("$code" -eq '10013') { Say 'firewall rule present; check OK: Windows refused a connection of python.exe'; exit 0 }
$state = try { (Get-NetFirewallProfile -PolicyStore ActiveStore | ForEach-Object {
    "$($_.Name) enabled=$($_.Enabled) local_rules=$($_.AllowLocalFirewallRules)" }) -join ', ' } catch { 'profile state not readable' }
Say "firewall rule present, but the block was not verified (connect result $code; $state)"
exit 2
