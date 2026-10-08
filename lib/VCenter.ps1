# PowerCLI helpers: connect to configured vCenters and resolve VM names safely.
# Requires lib\Common.ps1 and VMware.PowerCLI (VMware.VimAutomation.Core).

function Assert-PowerCli {
    if (-not (Get-Command Connect-VIServer -ErrorAction SilentlyContinue)) {
        throw 'VMware PowerCLI is not installed: Install-Module VMware.PowerCLI -Scope CurrentUser'
    }
}

function Connect-ToolkitVCenter {
    <#
    .SYNOPSIS
        Connects to one or more configured vCenters (Multiple mode). Returns the VIServer objects.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Config, [Parameter(Mandatory = $true)][object[]]$VCenter)
    Assert-PowerCli
    $certAction = if ($Config.Settings.SkipCertificateCheck) { 'Ignore' } else { 'Fail' }
    $null = Set-PowerCLIConfiguration -Scope Session -DefaultVIServerMode Multiple -InvalidCertificateAction $certAction `
        -ParticipateInCeip $false -Confirm:$false -WarningAction SilentlyContinue
    foreach ($v in $VCenter) {
        $cred = Get-ToolkitCredential -Name $v.Credential -CredentialPath $Config.Settings.CredentialPath
        try {
            $server = Connect-VIServer -Server $v.Name -Credential $cred -ErrorAction Stop -WarningAction SilentlyContinue
            Write-Log "Connected to $($v.Name) ($($server.Version) build $($server.Build))" -Level Ok
            $server
        }
        catch { Write-Log "Cannot connect to $($v.Name): $($_.Exception.Message)" -Level Error }
    }
}

function Disconnect-ToolkitVCenter {
    param([object[]]$Server)
    foreach ($s in @($Server | Where-Object { $_ })) {
        Disconnect-VIServer -Server $s -Confirm:$false -Force -ErrorAction SilentlyContinue
    }
}

function Get-VMShadowReason {
    <#
    .SYNOPSIS
        Why a VM object is not a "real" VM: SRM placeholder, template or no disks. Empty string for a real VM.
    #>
    param([Parameter(Mandatory = $true)]$VM)
    $cfg = $VM.ExtensionData.Config
    if ($cfg -and $cfg.ManagedBy -and $cfg.ManagedBy.ExtensionKey -eq 'com.vmware.vcDr') {
        return "SRM placeholder (ManagedBy $($cfg.ManagedBy.ExtensionKey), type $($cfg.ManagedBy.Type))"
    }
    if ($cfg -and $cfg.Template) { return 'template' }
    $disks = @($cfg.Hardware.Device | Where-Object { $_ -is [VMware.Vim.VirtualDisk] })
    if (-not $disks) { return 'no virtual disks' }
    return ''
}

function Get-VMServerName {
    # vCenter of a PowerCLI object, from its Uid: /VIServer=user@vc-01.contoso.local:443/VirtualMachine=vm-1/
    param($Object)
    if ($Object.Uid -match '@([^:/]+):\d+/') { return $Matches[1] }
    return ''
}

function Resolve-ToolkitVM {
    <#
    .SYNOPSIS
        Finds a VM by exact name on all connected vCenters and returns exactly one real VM.
    .DESCRIPTION
        When the name matches several objects (typical with SRM: the protected VM and its placeholder at the
        recovery site), placeholders, templates and diskless shells are discarded. If more than one real VM
        remains, the function throws instead of guessing. The returned object is re-read by MoRef so later
        operations cannot hit a namesake.
    #>
    param([Parameter(Mandatory = $true)][string]$Name, [Parameter(Mandatory = $true)][object[]]$Server)
    $found = @(Get-VM -Server $Server -Name $Name -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $Name })
    if (-not $found) { throw "VM '$Name' not found on $(($Server | ForEach-Object Name) -join ', ')." }
    if ($found.Count -eq 1) { return $found[0] }
    $real = foreach ($vm in $found) {
        $reason = Get-VMShadowReason $vm
        $where = Get-VMServerName $vm
        if ($reason) { Write-Log "  ${Name}: skipped $($vm.Id) on ${where}: $reason" -Level Warn } else { $vm }
    }
    $real = @($real)
    if (-not $real) { throw "VM '$Name': all $($found.Count) matches are placeholders, templates or empty shells." }
    if ($real.Count -gt 1) {
        throw "VM '$Name' is ambiguous: $(($real | ForEach-Object { "$($_.Id) ($($_.PowerState))" }) -join ', '). Resolve it manually."
    }
    $srv = $Server | Where-Object { $_.Name -eq (Get-VMServerName $real[0]) } | Select-Object -First 1
    return Get-VM -Server $srv -Id $real[0].Id
}
