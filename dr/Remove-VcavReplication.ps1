<#
.SYNOPSIS
    Removes Cloud Director Availability replications of VMs and optionally moves them out of the DR vApp.

.DESCRIPTION
    Finds every VM replication with the given VM name in the configured tenants (vCAV sessions of all tenants
    with VcavUrl), deletes it and waits for the vCAV task. With -MoveFromDrVApp a VM sitting in a vApp named
    "<name><Replication.DrVAppSuffix>" is moved back to "<name>" (powered-off VMs only).

    The VM itself and its replica disks' source are not touched; vCAV removes the replica at the destination.
    ConfirmImpact is High: PowerShell asks for confirmation unless -Confirm:$false is given.

.PARAMETER VMName
    VM names whose replications are removed.

.PARAMETER MoveFromDrVApp
    Move the VM back from the DR vApp.

.PARAMETER ConfigPath
    Toolkit configuration (.psd1). Default: ..\config\toolkit.psd1.

.EXAMPLE
    .\Remove-VcavReplication.ps1 -VMName srv-test-01 -WhatIf

.EXAMPLE
    .\Remove-VcavReplication.ps1 -VMName srv-test-01 -MoveFromDrVApp
#>
#Requires -Version 5.1
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $true)][string[]]$VMName,
    [switch]$MoveFromDrVApp,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config\toolkit.psd1')
)

. (Join-Path $PSScriptRoot '..\lib\Common.ps1')
. (Join-Path $PSScriptRoot '..\lib\VcdClient.ps1')
. (Join-Path $PSScriptRoot '..\lib\Replication.ps1')

$config = Import-ToolkitConfig -Path $ConfigPath
$null = Start-ToolkitLog -LogPath $config.Settings.LogPath -Name 'Remove-VcavReplication'
$suffix = if ($config.Replication.DrVAppSuffix) { $config.Replication.DrVAppSuffix } else { '-DR' }

$vcd = @{}; $vcav = @{}
$failed = 0
try {
    foreach ($t in $config.Tenants) {
        if (-not $t.VcavUrl -and -not $MoveFromDrVApp) { continue }
        $cred = Get-ToolkitCredential -Name $t.Credential -CredentialPath $config.Settings.CredentialPath
        if ($t.VcavUrl) { $vcav[$t.Name] = Connect-Vcav -Tenant $t -Config $config -Credential $cred }
        if ($MoveFromDrVApp) { $vcd[$t.Name] = Connect-VcdTenant -Tenant $t -Config $config -Credential $cred }
    }
    if (-not $vcav.Count) { throw 'No tenant with VcavUrl in the config.' }

    foreach ($name in $VMName) {
        try {
            $reps = @(Find-VcavVmReplication -Sessions $vcav -Name $name)
            if (-not $reps) { Write-Log "${name}: no replication found" -Level Warn }
            foreach ($rep in $reps) {
                $what = "$name ($($rep.source.site) -> $($rep.destination.site), health $($rep.overallHealth), id $($rep.id))"
                if (-not $PSCmdlet.ShouldProcess($what, 'Delete replication')) { continue }
                $task = Invoke-VcavApi -Session $rep.VcavSession -Path "vm-replications/$($rep.id)" -Method Delete
                $null = Wait-VcavTask -Session $rep.VcavSession -TaskId $task.id
                Write-Log "${name}: replication $($rep.id) deleted" -Level Ok
            }
            if ($MoveFromDrVApp) {
                $vm = Find-VcdVm -Sessions $vcd -Name $name
                $vapp = [string]$vm.Record.containerName
                if ($vapp.EndsWith($suffix) -and $vapp.Length -gt $suffix.Length) {
                    Move-VcdVmToVApp -Session $vm.Session -VmRecord $vm.Record -TargetVAppName $vapp.Substring(0, $vapp.Length - $suffix.Length)
                }
            }
        }
        catch {
            $failed++
            Write-Log "${name}: $(Get-RestErrorText $_)" -Level Error
        }
    }
}
finally {
    foreach ($s in $vcd.Values) { Disconnect-VcdTenant $s }
    foreach ($s in $vcav.Values) { Disconnect-Vcav $s }
}
if ($failed) { exit 1 }
