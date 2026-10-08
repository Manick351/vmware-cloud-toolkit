# Helpers shared by dr\New-VcavReplication.ps1 and dr\Remove-VcavReplication.ps1.
# Requires lib\Common.ps1 and lib\VcdClient.ps1.

function Find-VcdVm {
    <#
    .SYNOPSIS
        Finds a VM by exact name in the given tenant sessions. Returns @{ Tenant; Session; Record }.
        Throws when the name is not found or is ambiguous.
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Sessions, [Parameter(Mandatory = $true)][string]$Name)
    $hits = foreach ($tenantName in $Sessions.Keys) {
        $s = $Sessions[$tenantName]
        foreach ($r in (Get-VcdQueryRecord -Session $s -Type vm -Filter "isVAppTemplate==false;name==$Name")) {
            if ($r.name -eq $Name) { [PSCustomObject]@{ Tenant = $tenantName; Session = $s; Record = $r } }
        }
    }
    $hits = @($hits)
    if (-not $hits) { throw "VM '$Name' not found in tenants $($Sessions.Keys -join ', ')." }
    if ($hits.Count -gt 1) {
        throw "VM name '$Name' is ambiguous: $(($hits | ForEach-Object { "$($_.Tenant)/$($_.Record.containerName)" }) -join ', ')."
    }
    return $hits[0]
}

function Find-VcavVmReplication {
    # All VM replications with this VM name visible to the given vCAV sessions (deduplicated by id).
    param([Parameter(Mandatory = $true)][hashtable]$Sessions, [Parameter(Mandatory = $true)][string]$Name)
    $seen = @{}
    foreach ($tenantName in $Sessions.Keys) {
        $s = $Sessions[$tenantName]
        foreach ($it in (Get-VcavItem -Session $s -Path "vm-replications?vmName=$([Uri]::EscapeDataString($Name))")) {
            if ($it.vmName -ne $Name -or $seen.ContainsKey($it.id)) { continue }
            $seen[$it.id] = $true
            $it | Add-Member -NotePropertyName VcavSession -NotePropertyValue $s -Force
            $it | Add-Member -NotePropertyName VcavTenant -NotePropertyValue $tenantName -Force -PassThru
        }
    }
}

function Get-UrnUuid([string]$Urn) {
    if ($Urn -match '([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})') { return $Matches[1].ToLower() }
    return $null
}

function Move-VcdVmToVApp {
    <#
    .SYNOPSIS
        Moves a powered-off VM into another vApp of the same VDC, creating the target vApp with the same
        vApp networks when it does not exist yet (composeVApp + recomposeVApp with sourceDelete).
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]$Session,
        [Parameter(Mandatory = $true)]$VmRecord,
        [Parameter(Mandatory = $true)][string]$TargetVAppName
    )
    if ($VmRecord.containerName -eq $TargetVAppName) { return }
    if ($VmRecord.status -ne 'POWERED_OFF') {
        Write-Log "VM $($VmRecord.name) is $($VmRecord.status); moving into '$TargetVAppName' needs it powered off, skipped" -Level Warn
        return
    }
    if (-not $PSCmdlet.ShouldProcess("$($VmRecord.name): vApp '$($VmRecord.containerName)' -> '$TargetVAppName'", 'Move VM')) { return }

    $source = (Invoke-VcdApi -Session $Session -Path $VmRecord.container).VApp
    $target = @(Get-VcdQueryRecord -Session $Session -Type vApp -Filter "name==$TargetVAppName;vdc==$($VmRecord.vdc)") | Select-Object -First 1
    if (-not $target) {
        $networks = (@($source.NetworkConfigSection.NetworkConfig) | Where-Object { $_ } | ForEach-Object { $_.OuterXml }) -join "`n"
        $name = [System.Security.SecurityElement]::Escape($TargetVAppName)
        $compose = @"
<?xml version="1.0" encoding="UTF-8"?>
<ComposeVAppParams name="$name" xmlns="http://www.vmware.com/vcloud/v1.5" xmlns:ovf="http://schemas.dmtf.org/ovf/envelope/1">
  <InstantiationParams>
    <NetworkConfigSection>
      <ovf:Info>vApp networks copied from $([System.Security.SecurityElement]::Escape($VmRecord.containerName))</ovf:Info>
      $networks
    </NetworkConfigSection>
  </InstantiationParams>
  <AllEULAsAccepted>true</AllEULAsAccepted>
</ComposeVAppParams>
"@
        $resp = Invoke-VcdApi -Session $Session -Path "$($VmRecord.vdc)/action/composeVApp" -Method Post -Body $compose `
            -ContentType 'application/vnd.vmware.vcloud.composeVAppParams+xml'
        $task = Get-VcdTaskFromResponse $resp
        if ($task) { $null = Wait-VcdTask -Session $Session -Task $task -Activity "Create vApp $TargetVAppName" }
        $target = @(Get-VcdQueryRecord -Session $Session -Type vApp -Filter "name==$TargetVAppName;vdc==$($VmRecord.vdc)") | Select-Object -First 1
        if (-not $target) { throw "vApp '$TargetVAppName' was not created." }
        Write-Log "vApp '$TargetVAppName' created" -Level Ok
    }

    $vm = (Invoke-VcdApi -Session $Session -Path $VmRecord.href).Vm
    $recompose = @"
<?xml version="1.0" encoding="UTF-8"?>
<RecomposeVAppParams xmlns="http://www.vmware.com/vcloud/v1.5" xmlns:ovf="http://schemas.dmtf.org/ovf/envelope/1">
  <SourcedItem sourceDelete="true">
    <Source href="$($VmRecord.href)"/>
    <InstantiationParams>
      $($vm.NetworkConnectionSection.OuterXml)
    </InstantiationParams>
    <StorageProfile href="$($vm.StorageProfile.href)"/>
  </SourcedItem>
  <AllEULAsAccepted>true</AllEULAsAccepted>
</RecomposeVAppParams>
"@
    $resp = Invoke-VcdApi -Session $Session -Path "$($target.href)/action/recomposeVApp" -Method Post -Body $recompose `
        -ContentType 'application/vnd.vmware.vcloud.recomposeVAppParams+xml'
    $null = Wait-VcdTask -Session $Session -Task (Get-VcdTaskFromResponse $resp) -Activity "Move $($VmRecord.name)"
    Write-Log "$($VmRecord.name) moved into vApp '$TargetVAppName'" -Level Ok
}
