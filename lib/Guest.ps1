# Guest OS steps of a template update: patch, run customer hooks, seal, shut down.
# Requires lib\Common.ps1. Windows guests: WinRM (HTTP 5985) from the admin host. Linux guests: Posh-SSH module.
#
# The template definition (config Templates.<name>) drives everything:
#   OS            Windows | Linux
#   GuestFiles    @(@{ Source = 'D:\Distr\agent.msi'; Destination = 'C:\Temp\agents' }, ...)   copied before the hooks
#   GuestScripts  @('templates\guest\windows-baseline.example.ps1', ...)                       run in order (repo-relative or absolute)
#   SkipUpdateTitle  regex of Windows update titles to skip (for example '^VMware')
#   MaxUpdateRounds  install/reboot rounds for Windows (default 3)

$script:RepoRoot = Split-Path -Parent $PSScriptRoot

function Resolve-RepoPath([string]$Path) {
    if ([IO.Path]::IsPathRooted($Path)) { return $Path }
    return (Join-Path $script:RepoRoot $Path)
}

function Wait-Until {
    # Polls a scriptblock until it returns something truthy or the timeout expires. Returns the value or $null.
    param([scriptblock]$Condition, [int]$TimeoutSec, [int]$IntervalSec = 15, [string]$Activity = 'waiting')
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    do {
        $r = & $Condition
        if ($r) { return $r }
        Write-Log "  $Activity..." -Level Debug
        Start-Sleep -Seconds $IntervalSec
    } while ((Get-Date) -lt $deadline)
    return $null
}

#region Windows

function Connect-WindowsGuest {
    # Opens a PSSession to a freshly deployed guest (by IP: TrustedHosts must allow it, or use HTTPS listeners).
    param([string]$IPAddress, [pscredential]$Credential, [int]$TimeoutSec = 1800)
    $s = Wait-Until -TimeoutSec $TimeoutSec -IntervalSec 30 -Activity "WinRM on $IPAddress" -Condition {
        New-PSSession -ComputerName $IPAddress -Credential $Credential -ErrorAction SilentlyContinue
    }
    if (-not $s) { throw "No WinRM session to $IPAddress within $TimeoutSec s." }
    return $s
}

function Restart-WindowsGuest {
    # Reboots the guest and returns a new session once LastBootUpTime has changed.
    param([System.Management.Automation.Runspaces.PSSession]$Session, [string]$IPAddress, [pscredential]$Credential, [int]$TimeoutSec = 3600)
    $before = Invoke-Command -Session $Session -ScriptBlock { (Get-CimInstance Win32_OperatingSystem).LastBootUpTime }
    Invoke-Command -Session $Session -ScriptBlock { Restart-Computer -Force } -ErrorAction SilentlyContinue
    Remove-PSSession $Session -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 60
    $new = Wait-Until -TimeoutSec $TimeoutSec -IntervalSec 30 -Activity "reboot of $IPAddress" -Condition {
        $s = New-PSSession -ComputerName $IPAddress -Credential $Credential -ErrorAction SilentlyContinue
        if ($s) {
            $boot = Invoke-Command -Session $s -ScriptBlock { (Get-CimInstance Win32_OperatingSystem).LastBootUpTime } -ErrorAction SilentlyContinue
            if ($boot -and $boot -gt $before) { return $s }
            Remove-PSSession $s
        }
    }
    if (-not $new) { throw "$IPAddress did not come back after reboot within $TimeoutSec s." }
    Write-Log "  $IPAddress rebooted" -Level Ok
    return $new
}

function Copy-GuestFile {
    param($Session, [object[]]$Files)
    foreach ($f in @($Files | Where-Object { $_ })) {
        $src = Resolve-RepoPath $f.Source
        if (-not (Test-Path -LiteralPath $src)) { throw "Guest file not found: $src" }
        Invoke-Command -Session $Session -ScriptBlock { param($d) New-Item -ItemType Directory -Path $d -Force | Out-Null } -ArgumentList $f.Destination
        Copy-Item -LiteralPath $src -Destination $f.Destination -ToSession $Session -Recurse -Force
        Write-Log "  copied $(Split-Path $src -Leaf) -> $($f.Destination)"
    }
}

function Invoke-WindowsTemplateUpdate {
    <#
    .SYNOPSIS
        Patches a Windows template VM over WinRM, runs hooks, cleans up and starts sysprep (the VM shuts down).
    .DESCRIPTION
        1. Windows Update cache reset and reboot (stale SoftwareDistribution is the usual cause of failed scans);
        2. GuestFiles copy and GuestScripts hooks;
        3. update rounds: PSWindowsUpdate runs in a one-time scheduled task (the Windows Update API refuses
           remote sessions), the task is polled, then the guest is rebooted while a reboot is pending;
        4. update history for the report, DISM component cleanup;
        5. sysprep /generalize /oobe /shutdown from a SYSTEM task, so it survives the WinRM session ending.
        PSWindowsUpdate is copied from -PSWindowsUpdatePath when the guest has no internet access to PSGallery.
    .OUTPUTS
        Update history lines.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$IPAddress,
        [Parameter(Mandatory = $true)][pscredential]$Credential,
        [Parameter(Mandatory = $true)][hashtable]$Template,
        [string]$PSWindowsUpdatePath
    )
    $session = Connect-WindowsGuest -IPAddress $IPAddress -Credential $Credential
    try {
        Write-Log 'Resetting Windows Update cache'
        Invoke-Command -Session $session -ScriptBlock {
            New-Item -ItemType Directory -Path C:\Temp -Force | Out-Null
            Stop-Service wuauserv, bits -Force -ErrorAction SilentlyContinue
            Remove-Item C:\Windows\SoftwareDistribution -Recurse -Force -ErrorAction SilentlyContinue
            Start-Service bits, wuauserv -ErrorAction SilentlyContinue
        }
        $session = Restart-WindowsGuest -Session $session -IPAddress $IPAddress -Credential $Credential

        $hasModule = Invoke-Command -Session $session -ScriptBlock { [bool](Get-Module -ListAvailable PSWindowsUpdate) }
        if (-not $hasModule) {
            if ($PSWindowsUpdatePath) {
                Copy-Item -LiteralPath $PSWindowsUpdatePath -Destination 'C:\Program Files\WindowsPowerShell\Modules\PSWindowsUpdate' -ToSession $session -Recurse -Force
                Write-Log '  PSWindowsUpdate copied'
            }
            else {
                Invoke-Command -Session $session -ScriptBlock {
                    Install-PackageProvider NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
                    Install-Module PSWindowsUpdate -Force -Scope AllUsers
                }
                Write-Log '  PSWindowsUpdate installed from PSGallery'
            }
        }

        Copy-GuestFile -Session $session -Files $Template.GuestFiles
        foreach ($hook in @($Template.GuestScripts | Where-Object { $_ })) {
            $path = Resolve-RepoPath $hook
            Write-Log "Running hook $(Split-Path $path -Leaf)"
            Invoke-Command -Session $session -FilePath $path | ForEach-Object { Write-Log "  | $_" }
        }

        $skip = if ($Template.SkipUpdateTitle) { $Template.SkipUpdateTitle } else { '' }
        $rounds = if ($Template.MaxUpdateRounds) { [int]$Template.MaxUpdateRounds } else { 3 }
        for ($round = 1; $round -le $rounds; $round++) {
            Write-Log "Update round $round of $rounds"
            Invoke-Command -Session $session -ArgumentList $skip -ScriptBlock {
                param($SkipTitle)
                $script = @"
Import-Module PSWindowsUpdate
`$log = 'C:\Temp\TemplateUpdate.log'
`$updates = @(Get-WindowsUpdate -MicrosoftUpdate | Where-Object { -not '$SkipTitle' -or `$_.Title -notmatch '$SkipTitle' })
"Found `$(`$updates.Count) update(s)" | Out-File `$log -Append
foreach (`$u in `$updates) {
    "Installing `$(`$u.Title)" | Out-File `$log -Append
    Install-WindowsUpdate -MicrosoftUpdate -UpdateID `$u.Identity.UpdateID -AcceptAll -IgnoreReboot -Confirm:`$false | Out-Null
}
"Round finished" | Out-File `$log -Append
"@
                Set-Content -Path C:\Temp\TemplateUpdate.ps1 -Value $script -Encoding UTF8
                $action = New-ScheduledTaskAction -Execute powershell.exe -Argument '-NoProfile -ExecutionPolicy Bypass -File C:\Temp\TemplateUpdate.ps1'
                $principal = New-ScheduledTaskPrincipal -UserId SYSTEM -LogonType ServiceAccount -RunLevel Highest
                $settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Hours 5)
                Register-ScheduledTask -TaskName TemplateUpdate -Action $action -Principal $principal -Settings $settings -Force | Out-Null
                Start-ScheduledTask -TaskName TemplateUpdate
            }
            $done = Wait-Until -TimeoutSec (5 * 3600) -IntervalSec 120 -Activity 'updates installing' -Condition {
                (Invoke-Command -Session $session -ScriptBlock { [string](Get-ScheduledTask -TaskName TemplateUpdate).State }) -ne 'Running'
            }
            if (-not $done) { throw 'Update task did not finish within 5 hours.' }
            $found = Invoke-Command -Session $session -ScriptBlock {
                Unregister-ScheduledTask -TaskName TemplateUpdate -Confirm:$false
                (Select-String -Path C:\Temp\TemplateUpdate.log -Pattern '^Found (\d+)' | Select-Object -Last 1).Matches.Groups[1].Value
            }
            $pending = Invoke-Command -Session $session -ScriptBlock {
                (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') -or
                (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending')
            }
            Write-Log "  $found update(s) in this round, reboot pending: $pending"
            if ($pending) { $session = Restart-WindowsGuest -Session $session -IPAddress $IPAddress -Credential $Credential }
            if ([int]$found -eq 0 -and -not $pending) { break }
        }

        $history = Invoke-Command -Session $session -ScriptBlock {
            Import-Module PSWindowsUpdate
            Get-WUHistory -MaxDate (Get-Date).AddDays(-1) | Sort-Object Date | ForEach-Object { '{0:yyyy-MM-dd HH:mm} {1} - {2}' -f $_.Date, $_.Result, $_.Title }
        }
        Write-Log 'Component store cleanup'
        Invoke-Command -Session $session -ScriptBlock { Dism.exe /Online /Cleanup-Image /StartComponentCleanup /ResetBase /Quiet | Out-Null }

        Write-Log 'Scheduling sysprep (generalize, oobe, shutdown)'
        Invoke-Command -Session $session -ScriptBlock {
            $payload = @'
Remove-Item C:\Temp\TemplateUpdate.* -Force -ErrorAction SilentlyContinue
Remove-Item C:\Windows\System32\Sysprep\Panther -Recurse -Force -ErrorAction SilentlyContinue
Get-NetAdapter | ForEach-Object { Set-NetIPInterface -InterfaceIndex $_.ifIndex -Dhcp Enabled -ErrorAction SilentlyContinue; Set-DnsClientServerAddress -InterfaceIndex $_.ifIndex -ResetServerAddresses -ErrorAction SilentlyContinue }
Start-Process C:\Windows\System32\Sysprep\sysprep.exe -ArgumentList '/generalize /oobe /shutdown /quiet' -Wait
'@
            Set-Content -Path C:\Windows\Temp\RunSysprep.ps1 -Value $payload -Encoding UTF8
            $action = New-ScheduledTaskAction -Execute powershell.exe -Argument '-NoProfile -ExecutionPolicy Bypass -File C:\Windows\Temp\RunSysprep.ps1'
            $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1)
            $principal = New-ScheduledTaskPrincipal -UserId SYSTEM -LogonType ServiceAccount -RunLevel Highest
            Register-ScheduledTask -TaskName TemplateSysprep -Action $action -Trigger $trigger -Principal $principal -Force | Out-Null
        }
        return $history
    }
    finally { if ($session) { Remove-PSSession $session -ErrorAction SilentlyContinue } }
}

#endregion

#region Linux

function Invoke-LinuxGuestCommand {
    param([int]$SessionId, [string]$Command, [int]$TimeoutSec = 3600, [switch]$AllowFailure)
    $r = Invoke-SSHCommand -SessionId $SessionId -Command $Command -TimeOut $TimeoutSec
    if ($r.ExitStatus -ne 0 -and -not $AllowFailure) { throw "Guest command failed ($($r.ExitStatus)): $Command`n$($r.Error)" }
    return $r.Output
}

function Connect-LinuxGuest {
    # SSH session with root or passwordless sudo. Returns @{ Id; Sudo }.
    param([string]$IPAddress, [pscredential]$Credential, [int]$TimeoutSec = 1800)
    if (-not (Get-Command New-SSHSession -ErrorAction SilentlyContinue)) { throw 'Posh-SSH is not installed: Install-Module Posh-SSH -Scope CurrentUser' }
    $s = Wait-Until -TimeoutSec $TimeoutSec -IntervalSec 30 -Activity "SSH on $IPAddress" -Condition {
        New-SSHSession -ComputerName $IPAddress -Credential $Credential -AcceptKey -Force -ConnectionTimeout 15 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    }
    if (-not $s) { throw "No SSH session to $IPAddress within $TimeoutSec s." }
    $sudo = if ($Credential.UserName -eq 'root') { '' } else { 'sudo -n ' }
    $who = (Invoke-SSHCommand -SessionId $s.SessionId -Command "${sudo}whoami" -TimeOut 30).Output
    if ($who -notcontains 'root') { Remove-SSHSession -SessionId $s.SessionId | Out-Null; throw "$($Credential.UserName) has no passwordless sudo on $IPAddress." }
    return @{ Id = $s.SessionId; Sudo = $sudo }
}

function Invoke-LinuxTemplateUpdate {
    <#
    .SYNOPSIS
        Patches a Linux template VM over SSH, runs hooks, seals it and schedules a shutdown.
    .DESCRIPTION
        apt or dnf/yum upgrade, reboot, GuestFiles upload (to /tmp, then moved with sudo), GuestScripts hooks
        (bash, as root), then sealing so clones get unique identities: cloud-init clean, empty machine-id,
        SSH host keys removed (regenerated on first boot), logs and shell history cleared. The shutdown is
        scheduled with systemd-run so the last SSH command returns cleanly.
    .OUTPUTS
        Package manager output of the upgrade.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$IPAddress,
        [Parameter(Mandatory = $true)][pscredential]$Credential,
        [Parameter(Mandatory = $true)][hashtable]$Template
    )
    $g = Connect-LinuxGuest -IPAddress $IPAddress -Credential $Credential
    $sudo = $g.Sudo
    try {
        $os = (Invoke-LinuxGuestCommand $g.Id "grep '^PRETTY_NAME=' /etc/os-release") -replace '^PRETTY_NAME="?|"$', ''
        Write-Log "Guest OS: $os"
        $upgrade = if ((Invoke-LinuxGuestCommand $g.Id 'command -v apt-get' -AllowFailure)) {
            "${sudo}env DEBIAN_FRONTEND=noninteractive sh -c 'apt-get update -q && apt-get -yq -o Dpkg::Options::=--force-confold upgrade && apt-get -yq autoremove && apt-get clean'"
        }
        elseif ((Invoke-LinuxGuestCommand $g.Id 'command -v dnf' -AllowFailure)) { "${sudo}sh -c 'dnf -y upgrade && dnf clean all'" }
        else { "${sudo}sh -c 'yum -y update && yum clean all'" }
        Write-Log 'Upgrading packages'
        $log = Invoke-LinuxGuestCommand $g.Id $upgrade -TimeoutSec 7200

        Write-Log 'Rebooting guest'
        Invoke-LinuxGuestCommand $g.Id "${sudo}systemd-run --on-active=5 /bin/systemctl reboot" -AllowFailure | Out-Null
        Remove-SSHSession -SessionId $g.Id | Out-Null
        Start-Sleep -Seconds 60
        $g = Connect-LinuxGuest -IPAddress $IPAddress -Credential $Credential

        foreach ($f in @($Template.GuestFiles | Where-Object { $_ })) {
            $src = Resolve-RepoPath $f.Source
            if (-not (Test-Path -LiteralPath $src)) { throw "Guest file not found: $src" }
            Set-SCPItem -ComputerName $IPAddress -Credential $Credential -Path $src -Destination /tmp -AcceptKey -Force
            $leaf = Split-Path $src -Leaf
            Invoke-LinuxGuestCommand $g.Id "${sudo}mkdir -p '$($f.Destination)' && ${sudo}mv '/tmp/$leaf' '$($f.Destination)/' && ${sudo}chown -R root:root '$($f.Destination)/$leaf'" | Out-Null
            Write-Log "  copied $leaf -> $($f.Destination)"
        }
        foreach ($hook in @($Template.GuestScripts | Where-Object { $_ })) {
            $path = Resolve-RepoPath $hook
            $leaf = Split-Path $path -Leaf
            Write-Log "Running hook $leaf"
            Set-SCPItem -ComputerName $IPAddress -Credential $Credential -Path $path -Destination /tmp -AcceptKey -Force
            Invoke-LinuxGuestCommand $g.Id "${sudo}bash '/tmp/$leaf'; rc=`$?; rm -f '/tmp/$leaf'; exit `$rc" | ForEach-Object { Write-Log "  | $_" }
        }

        Write-Log 'Sealing'
        $seal = @(
            'cloud-init clean --logs --seed 2>/dev/null || true'
            'rm -f /etc/netplan/*.BeforeVMwareCustomization 2>/dev/null || true'
            'truncate -s 0 /etc/machine-id; rm -f /var/lib/dbus/machine-id; ln -sf /etc/machine-id /var/lib/dbus/machine-id'
            'journalctl --rotate 2>/dev/null; journalctl --vacuum-time=1s 2>/dev/null; rm -rf /var/log/*.gz /var/log/*.[0-9] /var/tmp/* /tmp/*'
            "rm -f /root/.bash_history /home/*/.bash_history"
            'systemd-run --on-active=60 /bin/systemctl poweroff'
            'rm -f /etc/ssh/ssh_host_*'
        ) -join '; '
        Invoke-LinuxGuestCommand $g.Id "${sudo}sh -c `"$seal`"" -AllowFailure | Out-Null
        return $log
    }
    finally { Remove-SSHSession -SessionId $g.Id -ErrorAction SilentlyContinue | Out-Null }
}

#endregion

function Get-TemplateDefinition {
    param([hashtable]$Config, [string]$Name)
    $t = $Config.Templates[$Name]
    if (-not $t) { throw "Template '$Name' is not defined in the Templates section. Known: $($Config.Templates.Keys -join ', ')" }
    if ($t.OS -notin 'Windows', 'Linux') { throw "Template '$Name': OS must be Windows or Linux." }
    return $t
}
