<#
.SYNOPSIS
    Imports an OVF/OVA package into vCenter with PowerCLI and optionally converts it to a template.

.DESCRIPTION
    Uses Import-VApp, so the vCenter session comes from the configured credential and no password ever appears
    in a command line or an ovftool vi:// URL. Before importing it checks the package (descriptor plus the
    disk files it references), the host, datastore free space against the package size, the folder and that
    the name is free. Network mappings from the OVF can be pointed to one port group with -Network.

    Supports -WhatIf: everything is validated and printed, nothing is uploaded.

.PARAMETER Path
    .ovf descriptor (with its .vmdk files next to it) or .ova archive.

.PARAMETER VCenter
    Target vCenter (configured name or short name).

.PARAMETER TargetHost
    ESXi host that receives the import.

.PARAMETER Datastore
    Target datastore. Default: the one with most free space visible from the host.

.PARAMETER Folder
    VM folder. Default: the datacenter's VM root.

.PARAMETER Name
    VM/template name. Default: the package file name.

.PARAMETER DiskStorageFormat
    Thin, Thick or EagerZeroedThick. Default: Thin.

.PARAMETER Network
    Port group to connect every OVF network to.

.PARAMETER ConvertToTemplate
    Convert the imported VM to a template.

.PARAMETER Notes
    Notes (annotation) to set, for example "Updated 2026-10-08".

.PARAMETER ConfigPath
    Toolkit configuration (.psd1). Default: ..\config\toolkit.psd1.

.EXAMPLE
    .\Import-OvfTemplate.ps1 -Path D:\export\W2K22STD.ovf -VCenter vc-01 -TargetHost esx-01.contoso.local -Folder Templates -ConvertToTemplate -WhatIf
#>
#Requires -Version 5.1
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [Parameter(Mandatory = $true)][string]$VCenter,
    [Parameter(Mandatory = $true)][string]$TargetHost,
    [string]$Datastore,
    [string]$Folder,
    [string]$Name,
    [ValidateSet('Thin', 'Thick', 'EagerZeroedThick')][string]$DiskStorageFormat = 'Thin',
    [string]$Network,
    [switch]$ConvertToTemplate,
    [string]$Notes,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config\toolkit.psd1')
)

. (Join-Path $PSScriptRoot '..\lib\Common.ps1')
. (Join-Path $PSScriptRoot '..\lib\VCenter.ps1')

$config = Import-ToolkitConfig -Path $ConfigPath
$null = Start-ToolkitLog -LogPath $config.Settings.LogPath -Name 'Import-OvfTemplate'

$package = Get-Item -LiteralPath $Path -ErrorAction Stop
if ($package.Extension -notin '.ovf', '.ova') { throw 'Path must be an .ovf or .ova file.' }
if (-not $Name) { $Name = $package.BaseName }
$sizeBytes = $package.Length
if ($package.Extension -eq '.ovf') {
    [xml]$ovf = Get-Content -LiteralPath $package.FullName -Raw
    $files = @($ovf.Envelope.References.File | ForEach-Object { $_.href })
    $missing = @($files | Where-Object { -not (Test-Path -LiteralPath (Join-Path $package.DirectoryName $_)) })
    if ($missing) { throw "Package incomplete, missing next to the descriptor: $($missing -join ', ')" }
    $sizeBytes = ($files | ForEach-Object { (Get-Item -LiteralPath (Join-Path $package.DirectoryName $_)).Length } | Measure-Object -Sum).Sum
    Write-Log "Package: $($package.Name) + $($files.Count) file(s), $([Math]::Round($sizeBytes / 1GB, 1)) GB"
}
else { Write-Log "Package: $($package.Name), $([Math]::Round($sizeBytes / 1GB, 1)) GB" }

$servers = @(Connect-ToolkitVCenter -Config $config -VCenter (Select-ConfigVCenter -Config $config -Name $VCenter))
if (-not $servers) { throw 'No vCenter connection.' }
$server = $servers[0]
try {
    $vmHost = Get-VMHost -Name $TargetHost -Server $server -ErrorAction SilentlyContinue
    if (-not $vmHost -or $vmHost.ConnectionState -ne 'Connected') { throw "Host '$TargetHost' not found or not connected." }
    $stores = @(Get-Datastore -RelatedObject $vmHost -Server $server | Sort-Object FreeSpaceGB -Descending)
    $ds = if ($Datastore) { $stores | Where-Object Name -eq $Datastore } else { $stores | Select-Object -First 1 }
    if (-not $ds) { throw "Datastore '$Datastore' is not visible from $TargetHost. Available: $($stores.Name -join ', ')" }
    if ($ds.FreeSpaceGB * 1GB -lt $sizeBytes * 1.2) { Write-Log "Datastore $($ds.Name) has only $([Math]::Round($ds.FreeSpaceGB)) GB free" -Level Warn }
    $location = if ($Folder) { @(Get-Folder -Name $Folder -Type VM -Server $server -ErrorAction SilentlyContinue) } else { @() }
    if ($Folder -and $location.Count -ne 1) { throw "Folder '$Folder' not found or not unique." }
    if (@(Get-VM -Name $Name -Server $server -ErrorAction SilentlyContinue) + @(Get-Template -Name $Name -Server $server -ErrorAction SilentlyContinue)) {
        throw "An object named '$Name' already exists on $($server.Name)."
    }

    $ovfConfig = Get-OvfConfiguration -Ovf $package.FullName
    if ($Network -and $ovfConfig.NetworkMapping) {
        $pg = Get-VirtualPortGroup -VMHost $vmHost -Name $Network -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $pg) { throw "Port group '$Network' not found on $TargetHost." }
        foreach ($map in $ovfConfig.NetworkMapping.PSObject.Properties) { $map.Value.Value = $pg }
    }

    Write-Log "Target: $($server.Name), host $($vmHost.Name), datastore $($ds.Name), folder $(if ($location) { $location[0].Name } else { '(root)' }), disks $DiskStorageFormat"
    if (-not $PSCmdlet.ShouldProcess("$Name on $($server.Name)", "Import $($package.Name)")) { return }

    $import = @{ Source = $package.FullName; Name = $Name; VMHost = $vmHost; Datastore = $ds; DiskStorageFormat = $DiskStorageFormat; Server = $server; ErrorAction = 'Stop' }
    if ($location) { $import.InventoryLocation = $location[0] }
    if ($ovfConfig) { $import.OvfConfiguration = $ovfConfig }
    $started = Get-Date
    $vm = Import-VApp @import
    Write-Log "Imported in $([int]((Get-Date) - $started).TotalMinutes) min: $($vm.Name) ($($vm.Id))" -Level Ok
    if ($Notes) { $vm = Set-VM -VM $vm -Notes $Notes -Confirm:$false }
    if ($ConvertToTemplate) {
        $null = Set-VM -VM $vm -ToTemplate -Confirm:$false
        $vm = Get-Template -Name $Name -Server $server
        Write-Log "Converted to template $($vm.Id). Put this id into VCenters[].Templates in the config." -Level Ok
    }
    $vm
}
catch {
    Write-Log $_.Exception.Message -Level Error
    Write-Log "If the upload broke, check for a half-created VM named '$Name' and remove it before retrying." -Level Warn
    exit 1
}
finally { Disconnect-ToolkitVCenter $servers }
