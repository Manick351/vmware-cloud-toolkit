<#
.SYNOPSIS
    Lists, creates and removes vSphere VM snapshots across several vCenters, safely resolving VM names.

.DESCRIPTION
    Typical use: snapshot before a change, remove it after the change window, find forgotten snapshots.

      List    snapshots of the given VMs (or of all VMs with -AllVMs), with age and size;
      Create  one snapshot per VM, without memory and without quiescing, named "<Prefix>-<yyyyMMdd-HHmm>";
      Remove  snapshots created by this tool (name starts with -Prefix); -AnyName removes every snapshot.
              -OlderThanDays limits both List and Remove to old snapshots.

    VM names are searched on all configured vCenters. SRM placeholders, templates and diskless shells are
    ignored; a name that still matches several real VMs is refused. Operations run against the VM's MoRef.
    Create and Remove support -WhatIf / -Confirm.

.PARAMETER Action
    List, Create or Remove.

.PARAMETER VMName
    VM names. Required for Create; for List and Remove either -VMName or -AllVMs.

.PARAMETER AllVMs
    List / Remove across every VM of the selected vCenters (Remove still honours -Prefix unless -AnyName).

.PARAMETER Description
    Snapshot description (Create), for example the change number.

.PARAMETER Prefix
    Name prefix of snapshots created by this tool. Default: vct.

.PARAMETER AnyName
    Remove: do not filter by prefix.

.PARAMETER OlderThanDays
    List / Remove: only snapshots older than this.

.PARAMETER VCenter
    Limit to these configured vCenters (FQDN or short name).

.PARAMETER ConfigPath
    Toolkit configuration (.psd1). Default: ..\config\toolkit.psd1.

.EXAMPLE
    .\Invoke-VMSnapshot.ps1 -Action Create -VMName srv-app-01, srv-db-01 -Description 'CHG-1234 before upgrade'

.EXAMPLE
    .\Invoke-VMSnapshot.ps1 -Action Remove -VMName srv-app-01, srv-db-01 -WhatIf

.EXAMPLE
    .\Invoke-VMSnapshot.ps1 -Action List -AllVMs -OlderThanDays 7 | Export-Csv old-snapshots.csv -NoTypeInformation
#>
#Requires -Version 5.1
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true)][ValidateSet('List', 'Create', 'Remove')][string]$Action,
    [string[]]$VMName,
    [switch]$AllVMs,
    [string]$Description,
    [ValidatePattern('^[\w.-]+$')][string]$Prefix = 'vct',
    [switch]$AnyName,
    [ValidateRange(0, 3650)][int]$OlderThanDays = 0,
    [string[]]$VCenter,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config\toolkit.psd1')
)

. (Join-Path $PSScriptRoot '..\lib\Common.ps1')
. (Join-Path $PSScriptRoot '..\lib\VCenter.ps1')

if (-not $VMName -and -not $AllVMs) { throw 'Specify -VMName or -AllVMs.' }
if ($Action -eq 'Create' -and $AllVMs) { throw 'Create works on named VMs only.' }

$config = Import-ToolkitConfig -Path $ConfigPath
$null = Start-ToolkitLog -LogPath $config.Settings.LogPath -Name 'Invoke-VMSnapshot'
$servers = @(Connect-ToolkitVCenter -Config $config -VCenter (Select-ConfigVCenter -Config $config -Name $VCenter))
if (-not $servers) { throw 'No vCenter connection.' }

function ConvertTo-SnapshotRow($Snapshot, $VM) {
    [PSCustomObject]@{
        VCenter   = Get-VMServerName $VM
        VM        = $VM.Name
        Snapshot  = $Snapshot.Name
        Created   = $Snapshot.Created
        AgeDays   = [Math]::Round(((Get-Date) - $Snapshot.Created).TotalDays, 1)
        SizeGB    = [Math]::Round($Snapshot.SizeGB, 2)
        PowerState = $Snapshot.PowerState
        Description = $Snapshot.Description
        IsCurrent = $Snapshot.IsCurrent
    }
}

$failed = 0
try {
    $vms = if ($AllVMs) {
        Get-VM -Server $servers | Where-Object { -not (Get-VMShadowReason $_) }
    }
    else {
        foreach ($n in $VMName) {
            try { Resolve-ToolkitVM -Name $n -Server $servers }
            catch { $failed++; Write-Log $_.Exception.Message -Level Error }
        }
    }
    $cutoff = (Get-Date).AddDays(-$OlderThanDays)

    foreach ($vm in @($vms)) {
        switch ($Action) {
            'List' {
                Get-Snapshot -VM $vm -ErrorAction SilentlyContinue | Where-Object { $_.Created -le $cutoff } |
                    ForEach-Object { ConvertTo-SnapshotRow $_ $vm }
            }
            'Create' {
                $name = '{0}-{1:yyyyMMdd-HHmm}' -f $Prefix, (Get-Date)
                $desc = if ($Description) { $Description } else { "Created by vmware-cloud-toolkit, $env:USERNAME" }
                $existing = @(Get-Snapshot -VM $vm -ErrorAction SilentlyContinue)
                if ($existing) { Write-Log "$($vm.Name) already has $($existing.Count) snapshot(s), oldest $(($existing | Sort-Object Created)[0].Created)" -Level Warn }
                if ($PSCmdlet.ShouldProcess("$($vm.Name) on $(Get-VMServerName $vm)", "Create snapshot $name")) {
                    try {
                        $snap = New-Snapshot -VM $vm -Name $name -Description $desc -Memory:$false -Quiesce:$false -Confirm:$false -ErrorAction Stop
                        Write-Log "$($vm.Name): snapshot '$name' created" -Level Ok
                        ConvertTo-SnapshotRow $snap $vm
                    }
                    catch { $failed++; Write-Log "$($vm.Name): $($_.Exception.Message)" -Level Error }
                }
            }
            'Remove' {
                $targets = @(Get-Snapshot -VM $vm -ErrorAction SilentlyContinue | Where-Object {
                        $_.Created -le $cutoff -and ($AnyName -or $_.Name -like "$Prefix-*") })
                if (-not $targets) { if (-not $AllVMs) { Write-Log "$($vm.Name): nothing to remove" }; continue }
                foreach ($s in ($targets | Sort-Object Created -Descending)) {
                    if (-not $PSCmdlet.ShouldProcess("$($vm.Name): '$($s.Name)' from $($s.Created) ($([Math]::Round($s.SizeGB, 1)) GB)", 'Remove snapshot')) { continue }
                    try {
                        Remove-Snapshot -Snapshot $s -Confirm:$false -ErrorAction Stop
                        Write-Log "$($vm.Name): snapshot '$($s.Name)' removed" -Level Ok
                    }
                    catch { $failed++; Write-Log "$($vm.Name): '$($s.Name)': $($_.Exception.Message)" -Level Error }
                }
                if ((Get-VM -Id $vm.Id -Server (Get-VMServerName $vm)).ExtensionData.Runtime.ConsolidationNeeded) {
                    Write-Log "$($vm.Name): disk consolidation is needed (vSphere Client: Snapshots > Consolidate)" -Level Warn
                }
            }
        }
    }
}
finally { Disconnect-ToolkitVCenter $servers }
if ($failed) { Write-Log "$failed operation(s) failed" -Level Error; exit 1 }
