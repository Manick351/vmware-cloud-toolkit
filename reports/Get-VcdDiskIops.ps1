<#
.SYNOPSIS
    Per-disk IOPS limits of all VMs in Cloud Director tenants, with a per-storage-policy summary.

.DESCRIPTION
    Reads every VM once (one paged query per tenant, then the VM description) and lists each disk with its size,
    storage policy and IOPS limit (vcloud:iops of the disk's HostResource). The summary sheet aggregates by
    tenant and policy: number of disks, total GB, sum of IOPS limits, disks without a limit and the policy's own
    IOPS settings when the API exposes them (Cloud Director 10.x: IopsSettings of VdcStorageProfile).

    Useful before moving disks between policies or when the provider enforces IOPS per GB.

.PARAMETER ConfigPath
    Toolkit configuration (.psd1). Default: ..\config\toolkit.psd1.

.PARAMETER Tenant
    Names of configured tenants. Default: all.

.PARAMETER OutputPath
    Folder for the workbook. Default: Settings.ReportPath.

.PARAMETER PassThru
    Also return the disk rows to the pipeline.

.EXAMPLE
    .\Get-VcdDiskIops.ps1 -Tenant tenant-a

.EXAMPLE
    .\Get-VcdDiskIops.ps1 -PassThru | Where-Object { -not $_.'IOPS limit' }
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config\toolkit.psd1'),
    [string[]]$Tenant,
    [string]$OutputPath,
    [switch]$PassThru
)

. (Join-Path $PSScriptRoot '..\lib\Common.ps1')
. (Join-Path $PSScriptRoot '..\lib\VcdClient.ps1')
. (Join-Path $PSScriptRoot '..\lib\Xlsx.ps1')

$config = Import-ToolkitConfig -Path $ConfigPath
$tenants = Select-ConfigTenant -Config $config -Name $Tenant
if (-not $OutputPath) { $OutputPath = $config.Settings.ReportPath }
if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
$null = Start-ToolkitLog -LogPath $config.Settings.LogPath -Name 'Get-VcdDiskIops'

$disks = New-Object System.Collections.Generic.List[object]
$policies = @{}   # "<tenant>|<policy>" -> IOPS settings

foreach ($t in $tenants) {
    Write-Log "Tenant $($t.Name)"
    $vcd = $null
    try {
        $vcd = Connect-VcdTenant -Tenant $t -Config $config
        $vdcs = @(Get-VcdQueryRecord -Session $vcd -Type orgVdc)
        if ($t.ExcludeVdcPattern) { $vdcs = @($vdcs | Where-Object { $_.name -notmatch $t.ExcludeVdcPattern }) }
        $vdcName = @{}
        foreach ($v in $vdcs) { $vdcName[$v.href] = $v.name }

        $profileName = @{}
        foreach ($sp in (Get-VcdQueryRecord -Session $vcd -Type orgVdcStorageProfile)) {
            if (-not $vdcName.ContainsKey($sp.vdc)) { continue }
            $profileName[[string]$sp.href] = $sp.name
            try {
                $p = (Invoke-VcdApi -Session $vcd -Path $sp.href).VdcStorageProfile
                $io = $p.IopsSettings
                $policies["$($t.Name)|$($sp.name)"] = [PSCustomObject]@{
                    Enabled      = Get-XmlText $io.Enabled
                    DiskMax      = Get-XmlText $io.DiskIopsMax
                    DiskDefault  = Get-XmlText $io.DiskIopsDefault
                    PerGbMax     = Get-XmlText $io.DiskIopsPerGbMax
                    PolicyLimit  = Get-XmlText $io.StorageProfileIopsLimit
                }
            }
            catch { Write-Log "  policy $($sp.name): $(Get-RestErrorText $_)" -Level Warn }
        }

        $records = @(Get-VcdQueryRecord -Session $vcd -Type vm -Filter 'isVAppTemplate==false' |
            Where-Object { $vdcName.ContainsKey($_.vdc) })
        Write-Log "  $($records.Count) VM(s)"
        $i = 0
        foreach ($r in $records) {
            $i++
            Write-Progress -Activity "Tenant $($t.Name)" -Status $r.name -PercentComplete ($i * 100 / [Math]::Max(1, $records.Count))
            try { $vmXml = (Invoke-VcdApi -Session $vcd -Path $r.href).Vm }
            catch { Write-Log "  $($r.name): $(Get-RestErrorText $_)" -Level Warn; continue }
            $index = 0
            foreach ($d in (Get-VcdVmHardware -VmXml $vmXml -StorageProfileNames $profileName).Disks) {
                $iops = if ($d.Iops -and [int]$d.Iops -gt 0) { [int]$d.Iops } else { $null }
                $disks.Add([PSCustomObject]@{
                    Tenant         = $t.Name
                    VDC            = $vdcName[$r.vdc]
                    VM             = $r.name
                    Power          = $r.status
                    Disk           = $index
                    'Size GB'      = $d.CapacityGB
                    Policy         = $d.StorageProfile
                    'IOPS limit'   = $iops
                    'IOPS per GB'  = if ($iops -and $d.CapacityGB) { [Math]::Round($iops / $d.CapacityGB, 2) } else { $null }
                })
                $index++
            }
        }
        Write-Progress -Activity "Tenant $($t.Name)" -Completed
    }
    catch { Write-Log "Tenant $($t.Name) failed: $(Get-RestErrorText $_)" -Level Error }
    finally { if ($vcd) { Disconnect-VcdTenant $vcd } }
}

$summary = foreach ($g in ($disks | Group-Object Tenant, Policy | Sort-Object Name)) {
    $first = $g.Group[0]
    $limited = @($g.Group | Where-Object { $_.'IOPS limit' })
    $set = $policies["$($first.Tenant)|$($first.Policy)"]
    [PSCustomObject]@{
        Tenant                 = $first.Tenant
        Policy                 = $first.Policy
        Disks                  = $g.Count
        'Total GB'             = [Math]::Round(($g.Group | Measure-Object 'Size GB' -Sum).Sum, 1)
        'Sum of IOPS limits'   = ($limited | Measure-Object 'IOPS limit' -Sum).Sum
        'Disks without limit'  = $g.Count - $limited.Count
        'Policy IOPS enabled'  = if ($set) { $set.Enabled } else { $null }
        'Policy disk max'      = if ($set) { $set.DiskMax } else { $null }
        'Policy IOPS per GB'   = if ($set) { $set.PerGbMax } else { $null }
        'Policy total limit'   = if ($set) { $set.PolicyLimit } else { $null }
    }
}

$file = Join-Path $OutputPath ('vcd-disk-iops_{0:yyyyMMdd_HHmm}.xlsx' -f (Get-Date))
Export-Xlsx -Path $file -Sheets @(
    (New-XlsxSheet -Name 'Disks' -Rows $disks -Widths @{ VM = 28; VDC = 22; Policy = 22 })
    (New-XlsxSheet -Name 'By policy' -Rows @($summary) -Widths @{ Policy = 22 })
)
Write-Log "Report: $file ($($disks.Count) disks)" -Level Ok
if ($PassThru) { $disks }
else { $file }
