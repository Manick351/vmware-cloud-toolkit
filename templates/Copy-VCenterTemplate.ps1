<#
.SYNOPSIS
    Copies a VM template from one vCenter to another (or within one) and converts the copy to a template.

.DESCRIPTION
    Resolves and validates everything before copying: the source template (it must be a template, not a VM),
    target folder, host (must be connected) and datastore (must be visible from the host; free space is
    compared with the template's committed size). An object with the new name must not exist on the target.

    The copy is made with New-VM -Template across the two vCenter connections, which requires both vCenters to
    be in the same SSO domain (Enhanced Linked Mode) or Advanced Cross vCenter vMotion support. If that is not
    available, use Export-VApp on the source and templates\Import-OvfTemplate.ps1 on the target.

    Supports -WhatIf: everything is resolved and printed, nothing is copied.

.PARAMETER SourceVCenter
    Source vCenter (configured name or short name).

.PARAMETER SourceTemplate
    Template name or MoRef id (VirtualMachine-vm-123).

.PARAMETER TargetVCenter
    Target vCenter.

.PARAMETER TargetHost
    ESXi host on the target.

.PARAMETER TargetFolder
    VM folder name (or Folder-group-v123 id). Default: the datacenter's VM root.

.PARAMETER TargetDatastore
    Datastore name. Default: the datastore with most free space visible from the host.

.PARAMETER NewName
    Name of the copy. Default: "<source name>_copy".

.PARAMETER ConfigPath
    Toolkit configuration (.psd1). Default: ..\config\toolkit.psd1.

.EXAMPLE
    .\Copy-VCenterTemplate.ps1 -SourceVCenter vc-02 -SourceTemplate W2K22STD -TargetVCenter vc-01 -TargetHost esx-01.contoso.local -TargetFolder Templates -WhatIf
#>
#Requires -Version 5.1
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)][string]$SourceVCenter,
    [Parameter(Mandatory = $true)][string]$SourceTemplate,
    [Parameter(Mandatory = $true)][string]$TargetVCenter,
    [Parameter(Mandatory = $true)][string]$TargetHost,
    [string]$TargetFolder,
    [string]$TargetDatastore,
    [string]$NewName,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config\toolkit.psd1')
)

. (Join-Path $PSScriptRoot '..\lib\Common.ps1')
. (Join-Path $PSScriptRoot '..\lib\VCenter.ps1')

$config = Import-ToolkitConfig -Path $ConfigPath
$null = Start-ToolkitLog -LogPath $config.Settings.LogPath -Name 'Copy-VCenterTemplate'
$srcCfg = Select-ConfigVCenter -Config $config -Name $SourceVCenter
$dstCfg = Select-ConfigVCenter -Config $config -Name $TargetVCenter
$vcList = @($srcCfg[0])
if ($dstCfg[0].Name -ne $srcCfg[0].Name) { $vcList += $dstCfg[0] }
$servers = @(Connect-ToolkitVCenter -Config $config -VCenter $vcList)
try {
    $src = $servers | Where-Object Name -eq $srcCfg[0].Name
    $dst = $servers | Where-Object Name -eq $dstCfg[0].Name
    if (-not $src -or -not $dst) { throw 'Source or target vCenter is not connected.' }

    $tpl = if ($SourceTemplate -like 'VirtualMachine-*') { Get-Template -Id $SourceTemplate -Server $src -ErrorAction SilentlyContinue }
           else { Get-Template -Name $SourceTemplate -Server $src -ErrorAction SilentlyContinue }
    $tpl = @($tpl)
    if ($tpl.Count -ne 1) { throw "Source template '$SourceTemplate' not found or not unique on $($src.Name) (is it currently a VM?)." }
    $tpl = $tpl[0]
    $sizeGB = [Math]::Round([double]$tpl.ExtensionData.Summary.Storage.Committed / 1GB, 1)
    Write-Log "Source: $($tpl.Name) ($($tpl.Id)), $sizeGB GB, tools $($tpl.ExtensionData.Guest.ToolsVersion), note '$($tpl.ExtensionData.Config.Annotation)'"
    if ([int]$tpl.ExtensionData.Guest.ToolsVersion -eq 0) { Write-Log 'The source reports no VMware Tools; the copy will have the same problem.' -Level Warn }

    $vmHost = Get-VMHost -Name $TargetHost -Server $dst -ErrorAction SilentlyContinue
    if (-not $vmHost) { throw "Host '$TargetHost' not found on $($dst.Name)." }
    if ($vmHost.ConnectionState -ne 'Connected') { throw "Host '$TargetHost' is $($vmHost.ConnectionState)." }

    $folder = if (-not $TargetFolder) { Get-Folder -Name vm -Type VM -Server $dst -Location (Get-Datacenter -VMHost $vmHost -Server $dst) -NoRecursion }
              elseif ($TargetFolder -like 'Folder-*') { Get-Folder -Id $TargetFolder -Server $dst -ErrorAction SilentlyContinue }
              else { Get-Folder -Name $TargetFolder -Type VM -Server $dst -ErrorAction SilentlyContinue }
    $folder = @($folder)
    if ($folder.Count -ne 1) { throw "Folder '$TargetFolder' not found or not unique on $($dst.Name)." }

    $stores = @(Get-Datastore -RelatedObject $vmHost -Server $dst | Sort-Object FreeSpaceGB -Descending)
    $ds = if ($TargetDatastore) { $stores | Where-Object Name -eq $TargetDatastore } else { $stores | Select-Object -First 1 }
    if (-not $ds) {
        $list = ($stores | ForEach-Object { '{0} ({1:N0} GB free)' -f $_.Name, $_.FreeSpaceGB }) -join ', '
        throw "Datastore '$TargetDatastore' is not visible from $TargetHost. Available: $list"
    }
    $level = if ($ds.FreeSpaceGB -gt 2 * $sizeGB) { 'Ok' } else { 'Warn' }
    Write-Log "Target: $($dst.Name), host $($vmHost.Name), folder $($folder[0].Name), datastore $($ds.Name) ($([Math]::Round($ds.FreeSpaceGB)) GB free)" -Level $level

    if (-not $NewName) { $NewName = "$($tpl.Name)_copy" }
    if (@(Get-VM -Name $NewName -Server $dst -ErrorAction SilentlyContinue) + @(Get-Template -Name $NewName -Server $dst -ErrorAction SilentlyContinue)) {
        throw "An object named '$NewName' already exists on $($dst.Name)."
    }
    foreach ($same in @(Get-Template -Name $tpl.Name -Server $dst -ErrorAction SilentlyContinue)) {
        Write-Log "Note: $($dst.Name) already has a template '$($same.Name)' ($($same.Id), note '$($same.ExtensionData.Config.Annotation)')" -Level Warn
    }

    if (-not $PSCmdlet.ShouldProcess("$($tpl.Name) -> $($dst.Name) as '$NewName'", "Copy template ($sizeGB GB)")) { return }
    $started = Get-Date
    $vm = New-VM -Name $NewName -Template $tpl -VMHost $vmHost -Datastore $ds -Location $folder[0] -Server $dst -ErrorAction Stop
    Write-Log "Copied in $([int]((Get-Date) - $started).TotalMinutes) min" -Level Ok
    $null = Set-VM -VM $vm -ToTemplate -Confirm:$false -ErrorAction Stop
    $new = Get-Template -Name $NewName -Server $dst
    Write-Log "Template '$($new.Name)' created: $($new.Id). Put this id into VCenters[].Templates in the config." -Level Ok
    $new
}
catch {
    Write-Log $_.Exception.Message -Level Error
    if ($_.Exception.Message -match 'cross|vCenter|not supported') {
        Write-Log 'If cross-vCenter cloning is not available, export with Export-VApp and use Import-OvfTemplate.ps1.' -Level Warn
    }
    exit 1
}
finally { Disconnect-ToolkitVCenter $servers }
