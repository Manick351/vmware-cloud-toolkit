<#
.SYNOPSIS
    Protects Cloud Director VMs with Cloud Director Availability replications using settings from the config.

.DESCRIPTION
    For every VM name:

      1. finds the VM in the configured tenants (exact name, must be unique);
      2. picks the replication pair whose SourceTenant is the VM's tenant (Replication.Pairs);
      3. resolves the destination VDC (pair's DestinationVdc, else the destination tenant's only VDC) and the
         destination storage policy (pair's StorageProfile, else the first policy of the VDC);
      4. opens a vCAV session in the destination tenant, extends it to the source site with the source
         tenant's credentials (POST /sessions/extend) and creates the replication (POST /vm-replications)
         with RPO, data connection type, quiescing and retention from Replication.Settings;
      5. optionally (-MoveToDrVApp) moves the VM into a vApp named "<current vApp><Replication.DrVAppSuffix>"
         so that protected VMs are visible in the tenant portal. Only powered-off VMs are moved.

    An existing replication of the VM is left alone unless -Replace is given (it is then deleted first).
    Supports -WhatIf and -Confirm.

.PARAMETER VMName
    VM names to protect.

.PARAMETER Replace
    Delete an existing replication of the VM and create a new one.

.PARAMETER MoveToDrVApp
    Move the VM into the DR vApp after the replication is created.

.PARAMETER ConfigPath
    Toolkit configuration (.psd1). Default: ..\config\toolkit.psd1.

.EXAMPLE
    .\New-VcavReplication.ps1 -VMName srv-app-01 -WhatIf

.EXAMPLE
    .\New-VcavReplication.ps1 -VMName srv-app-01, srv-app-02 -MoveToDrVApp -Confirm:$false
#>
#Requires -Version 5.1
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true)][string[]]$VMName,
    [switch]$Replace,
    [switch]$MoveToDrVApp,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config\toolkit.psd1')
)

. (Join-Path $PSScriptRoot '..\lib\Common.ps1')
. (Join-Path $PSScriptRoot '..\lib\VcdClient.ps1')
. (Join-Path $PSScriptRoot '..\lib\Replication.ps1')

$config = Import-ToolkitConfig -Path $ConfigPath
$null = Start-ToolkitLog -LogPath $config.Settings.LogPath -Name 'New-VcavReplication'
$pairs = @($config.Replication.Pairs)
if (-not $pairs) { throw 'Replication.Pairs is empty in the config.' }
$settings = @{ Rpo = 60; DataConnectionType = 'ENCRYPTED'; Quiesced = $false; RetentionInstances = 1; RetentionDistance = 60; InitialSyncTime = 0; IsMigration = $false }
if ($config.Replication.Settings) { foreach ($k in $config.Replication.Settings.Keys) { $settings[$k] = $config.Replication.Settings[$k] } }
$suffix = if ($config.Replication.DrVAppSuffix) { $config.Replication.DrVAppSuffix } else { '-DR' }

$tenantNames = @($pairs | ForEach-Object { $_.SourceTenant; $_.DestinationTenant } | Select-Object -Unique)
$tenants = Select-ConfigTenant -Config $config -Name $tenantNames
$creds = @{}; $vcd = @{}; $vcav = @{}
$failed = 0
try {
    foreach ($t in $tenants) {
        $creds[$t.Name] = Get-ToolkitCredential -Name $t.Credential -CredentialPath $config.Settings.CredentialPath
        $vcd[$t.Name] = Connect-VcdTenant -Tenant $t -Config $config -Credential $creds[$t.Name]
        if ($t.VcavUrl) { $vcav[$t.Name] = Connect-Vcav -Tenant $t -Config $config -Credential $creds[$t.Name] }
    }

    foreach ($name in $VMName) {
        try {
            $vm = Find-VcdVm -Sessions $vcd -Name $name
            $pair = $pairs | Where-Object { $_.SourceTenant -eq $vm.Tenant } | Select-Object -First 1
            if (-not $pair) { throw "no replication pair with SourceTenant '$($vm.Tenant)'" }
            $dstSession = $vcd[$pair.DestinationTenant]
            $dstVcav = $vcav[$pair.DestinationTenant]
            if (-not $dstVcav) { throw "destination tenant '$($pair.DestinationTenant)' has no VcavUrl" }
            Write-Log "$name ($($vm.Tenant), vApp $($vm.Record.containerName)) -> $($pair.DestinationTenant), site $($pair.SourceSite) -> $($pair.DestinationSite)"

            $existing = @(Find-VcavVmReplication -Sessions $vcav -Name $name)
            if ($existing -and -not $Replace) {
                Write-Log "${name}: already replicated ($($existing[0].source.site) -> $($existing[0].destination.site), health $($existing[0].overallHealth)); use -Replace to recreate" -Level Warn
                continue
            }

            $vdcRecords = @(Get-VcdQueryRecord -Session $dstSession -Type orgVdc)
            $dstTenant = $tenants | Where-Object { $_.Name -eq $pair.DestinationTenant }
            if ($dstTenant.ExcludeVdcPattern) { $vdcRecords = @($vdcRecords | Where-Object { $_.name -notmatch $dstTenant.ExcludeVdcPattern }) }
            $vdcRec = if ($pair.DestinationVdc) { $vdcRecords | Where-Object { $_.name -eq $pair.DestinationVdc } | Select-Object -First 1 }
                      elseif ($vdcRecords.Count -eq 1) { $vdcRecords[0] }
            if (-not $vdcRec) { throw "destination VDC not determined (set DestinationVdc in the pair; found: $($vdcRecords.name -join ', '))" }
            $vdc = (Invoke-VcdApi -Session $dstSession -Path $vdcRec.href).Vdc
            $profiles = @($vdc.VdcStorageProfiles.VdcStorageProfile | Where-Object { $_ })
            $sp = $profiles | Where-Object { $_.name -eq $pair.StorageProfile } | Select-Object -First 1
            if (-not $sp) {
                if ($pair.StorageProfile) { Write-Log "${name}: storage policy '$($pair.StorageProfile)' not in $($vdc.name), using '$($profiles[0].name)'" -Level Warn }
                $sp = $profiles[0]
            }

            $body = @{
                source          = @{ type = 'vm'; site = $pair.SourceSite; vmId = (Get-UrnUuid $vm.Record.href) }
                destination     = @{ type = 'vcloud'; site = $pair.DestinationSite; vdc = (Get-UrnUuid $vdc.id); storageProfile = (Get-UrnUuid $sp.id) }
                description     = 'Created by vmware-cloud-toolkit'
                rpo             = [int]$settings.Rpo
                dataConnectionType = $settings.DataConnectionType
                quiesced        = [bool]$settings.Quiesced
                retentionPolicy = @{ rules = @(@{ numberOfInstances = [int]$settings.RetentionInstances; distance = [int]$settings.RetentionDistance }) }
                initialSyncTime = [int64]$settings.InitialSyncTime
                isMigration     = [bool]$settings.IsMigration
            }
            $what = "$name -> $($pair.DestinationSite) / $($vdc.name) / $($sp.name), RPO $($settings.Rpo) min"
            if (-not $PSCmdlet.ShouldProcess($what, $(if ($existing) { 'Replace replication' } else { 'Create replication' }))) { continue }

            foreach ($old in $existing) {
                $task = Invoke-VcavApi -Session $old.VcavSession -Path "vm-replications/$($old.id)" -Method Delete
                $null = Wait-VcavTask -Session $old.VcavSession -TaskId $task.id
                Write-Log "${name}: previous replication $($old.id) deleted" -Level Ok
            }

            $srcCred = $creds[$vm.Tenant]
            $srcTenant = $tenants | Where-Object { $_.Name -eq $vm.Tenant }
            $srcUser = if ($srcCred.UserName -like '*@*' -or -not $srcTenant.Org) { $srcCred.UserName } else { "$($srcCred.UserName)@$($srcTenant.Org)" }
            $null = Invoke-VcavApi -Session $dstVcav -Path 'sessions/extend' -Method Post -Body @{
                type = 'credentials'; site = $pair.SourceSite; vcdUsername = $srcUser; vcdPassword = $srcCred.GetNetworkCredential().Password
            }
            $task = Invoke-VcavApi -Session $dstVcav -Path "vm-replications?site=$([Uri]::EscapeDataString($pair.DestinationSite))" -Method Post -Body $body
            $null = Wait-VcavTask -Session $dstVcav -TaskId $task.id
            Write-Log "${name}: replication created ($what)" -Level Ok

            if ($MoveToDrVApp -and $vm.Record.containerName -and -not $vm.Record.containerName.EndsWith($suffix)) {
                Move-VcdVmToVApp -Session $vm.Session -VmRecord $vm.Record -TargetVAppName ($vm.Record.containerName + $suffix) -Confirm:$false
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
