<#
.SYNOPSIS
    VM and resource report for one or more Cloud Director tenants, written to a single XLSX workbook.

.DESCRIPTION
    For every configured tenant (organization) the script reads, through the tenant-side Cloud Director API:

      - all VMs (one paged query per tenant) and, for each VM, its full description: vCPU, RAM, every disk with
        its storage policy, IP addresses, guest OS, creation date, snapshot date;
      - every Org VDC: allocated / used / free vCPU and RAM, limit / used / free per storage policy;
      - optionally (tenants with VcavUrl) Cloud Director Availability replications: which VMs are protected and
        how much vCPU / RAM a failover would need in each destination VDC.

    Sheets: "VMs", "VDC resources", "Disks by policy", and "Errors" when a tenant could not be read.
    A tenant that fails does not stop the report; the error is listed and optionally mailed separately.

.PARAMETER ConfigPath
    Toolkit configuration (.psd1). Default: ..\config\toolkit.psd1.

.PARAMETER Tenant
    Names of configured tenants to include. Default: all.

.PARAMETER OutputPath
    Folder for the workbook. Default: Settings.ReportPath.

.PARAMETER SkipReplication
    Do not query Cloud Director Availability.

.PARAMETER Fast
    Use only query records (no per-VM request): disks are summed per VM, IP / snapshot / OS columns stay empty.
    Much faster for large tenants.

.PARAMETER SendMail
    Mail the workbook to Mail.To; API errors go to Mail.ErrorTo (if set) as a separate HTML message.

.EXAMPLE
    .\Get-VcdVmReport.ps1

.EXAMPLE
    .\Get-VcdVmReport.ps1 -Tenant tenant-a, tenant-b -SendMail

.OUTPUTS
    Path of the generated workbook.
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config\toolkit.psd1'),
    [string[]]$Tenant,
    [string]$OutputPath,
    [switch]$SkipReplication,
    [switch]$Fast,
    [switch]$SendMail
)

. (Join-Path $PSScriptRoot '..\lib\Common.ps1')
. (Join-Path $PSScriptRoot '..\lib\VcdClient.ps1')
. (Join-Path $PSScriptRoot '..\lib\Xlsx.ps1')

$config = Import-ToolkitConfig -Path $ConfigPath
$tenants = Select-ConfigTenant -Config $config -Name $Tenant
if (-not $OutputPath) { $OutputPath = $config.Settings.ReportPath }
if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
$log = Start-ToolkitLog -LogPath $config.Settings.LogPath -Name 'Get-VcdVmReport'

$vmRows = New-Object System.Collections.Generic.List[object]
$diskRows = New-Object System.Collections.Generic.List[object]
$vdcInfo = New-Object System.Collections.Generic.List[object]   # capacity objects with tenant name
$errors = New-Object System.Collections.Generic.List[object]
$vmById = @{}                                                   # VM uuid -> @{ Cpu; RamGB }  (all tenants, for DR sums)
$replicatedIds = New-Object 'System.Collections.Generic.HashSet[string]'
$replications = New-Object System.Collections.Generic.List[object]

function Get-UuidFromHref([string]$Href) {
    if ($Href -match '([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})') { return $Matches[1].ToLower() }
    return $null
}

foreach ($t in $tenants) {
    Write-Log "Tenant $($t.Name): $($t.VcdUrl)"
    $vcd = $null
    try {
        $vcd = Connect-VcdTenant -Tenant $t -Config $config

        $vdcs = @(Get-VcdQueryRecord -Session $vcd -Type orgVdc)
        if ($t.ExcludeVdcPattern) { $vdcs = @($vdcs | Where-Object { $_.name -notmatch $t.ExcludeVdcPattern }) }
        $vdcName = @{}
        foreach ($v in $vdcs) { $vdcName[$v.href] = $v.name }

        $profileName = @{}
        foreach ($sp in (Get-VcdQueryRecord -Session $vcd -Type orgVdcStorageProfile)) { $profileName[[string]$sp.href] = $sp.name }

        foreach ($v in $vdcs) {
            $cap = Get-VcdVdcCapacity -Session $vcd -VdcHref $v.href
            $cap | Add-Member -NotePropertyName Tenant -NotePropertyValue $t.Name
            $cap | Add-Member -NotePropertyName Id -NotePropertyValue (Get-UuidFromHref $v.href)
            $vdcInfo.Add($cap)
        }

        $records = @(Get-VcdQueryRecord -Session $vcd -Type vm -Filter 'isVAppTemplate==false' |
            Where-Object { $vdcName.ContainsKey($_.vdc) })
        Write-Log "  $($vdcs.Count) VDC(s), $($records.Count) VM(s)"

        $i = 0
        foreach ($r in $records) {
            $i++
            Write-Progress -Activity "Tenant $($t.Name)" -Status $r.name -PercentComplete ($i * 100 / [Math]::Max(1, $records.Count))
            $id = Get-UuidFromHref $r.href
            $row = [ordered]@{
                Tenant      = $t.Name
                VDC         = $vdcName[$r.vdc]
                vApp        = $r.containerName
                VM          = $r.name
                Power       = $r.status
                vCPU        = [int]$r.numberOfCpus
                'RAM GB'    = [Math]::Round([double]$r.memoryMB / 1024, 1)
                'Disk GB'   = [Math]::Round([double]$r.totalStorageAllocatedMb / 1024, 1)
                Disks       = ''
                'IP address' = $r.ipAddress
                'Guest OS'  = $r.guestOs
                Created     = $null
                Snapshot    = $null
                Replicated  = $null
                'VM id'     = $id
            }
            $perPolicy = [ordered]@{}
            if ($Fast) {
                $perPolicy[[string]$r.storageProfileName] = $row['Disk GB']
            }
            else {
                try {
                    $vmXml = (Invoke-VcdApi -Session $vcd -Path $r.href).Vm
                    $hw = Get-VcdVmHardware -VmXml $vmXml -StorageProfileNames $profileName
                    $row.vCPU = $hw.Cpu
                    $row['RAM GB'] = $hw.RamGB
                    $row['Disk GB'] = [Math]::Round(($hw.Disks | Measure-Object CapacityGB -Sum).Sum, 1)
                    $row.Disks = ($hw.Disks | ForEach-Object { '{0} GB ({1})' -f $_.CapacityGB, $_.StorageProfile }) -join "`n"
                    foreach ($d in $hw.Disks) {
                        $key = [string]$d.StorageProfile
                        if (-not $perPolicy.Contains($key)) { $perPolicy[$key] = 0 }
                        $perPolicy[$key] += $d.CapacityGB
                    }
                    $ips = @($vmXml.NetworkConnectionSection.NetworkConnection | ForEach-Object { Get-XmlText $_.IpAddress } | Where-Object { $_ })
                    if ($ips) { $row['IP address'] = $ips -join "`n" }
                    $os = Get-XmlText $vmXml.OperatingSystemSection.Description
                    if ($os) { $row['Guest OS'] = $os }
                    $created = Get-XmlText $vmXml.DateCreated
                    if ($created) { $row.Created = [datetime]$created }
                    $snap = @($vmXml.SnapshotSection.Snapshot | Where-Object { $_ } | ForEach-Object { $_.created } | Sort-Object)
                    if ($snap) { $row.Snapshot = [datetime]$snap[-1] }
                }
                catch {
                    Write-Log "  $($r.name): details not read: $(Get-RestErrorText $_)" -Level Warn
                }
            }
            $vmById[$id] = @{ Cpu = $row.vCPU; RamGB = $row['RAM GB'] }
            $vmRows.Add([PSCustomObject]$row)

            $disk = [ordered]@{ Tenant = $t.Name; VDC = $row.VDC; VM = $r.name; Power = $r.status }
            foreach ($k in $perPolicy.Keys) { $disk["$k GB"] = [Math]::Round($perPolicy[$k], 1) }
            $diskRows.Add([PSCustomObject]$disk)
        }
        Write-Progress -Activity "Tenant $($t.Name)" -Completed

        if ($t.VcavUrl -and -not $SkipReplication) {
            $vcav = $null
            try {
                $vcav = Connect-Vcav -Tenant $t -Config $config
                foreach ($rep in (Get-VcavItem -Session $vcav -Path 'vm-replications')) {
                    $replications.Add($rep)
                    if ($rep.vmId) { [void]$replicatedIds.Add(([string]$rep.vmId).ToLower()) }
                }
            }
            catch {
                Write-Log "  vCAV: $(Get-RestErrorText $_)" -Level Warn
                $errors.Add([PSCustomObject]@{ Tenant = $t.Name; Source = 'vCAV'; Time = Get-Date; Error = (Get-RestErrorText $_) })
            }
            finally { if ($vcav) { Disconnect-Vcav $vcav } }
        }
        Write-Log "  done" -Level Ok
    }
    catch {
        $msg = Get-RestErrorText $_
        Write-Log "Tenant $($t.Name) failed: $msg" -Level Error
        $errors.Add([PSCustomObject]@{ Tenant = $t.Name; Source = 'Cloud Director'; Time = Get-Date; Error = $msg })
    }
    finally { if ($vcd) { Disconnect-VcdTenant $vcd } }
}

# Replication flags and failover demand per destination VDC. The same replication can be listed by both
# the source and the destination tenant, so it is counted once by id.
$demand = @{}
$counted = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($rep in $replications) {
    if (-not $counted.Add([string]$rep.id)) { continue }
    $dst = if ($rep.destination.vdcId) { Get-UuidFromHref ([string]$rep.destination.vdcId) } else { $null }
    if (-not $dst) { continue }
    if (-not $demand.ContainsKey($dst)) { $demand[$dst] = @{ Cpu = 0; RamGB = 0.0; Count = 0; Unknown = 0 } }
    $vm = $vmById[([string]$rep.vmId).ToLower()]
    if ($vm) { $demand[$dst].Cpu += $vm.Cpu; $demand[$dst].RamGB += $vm.RamGB; $demand[$dst].Count++ }
    else { $demand[$dst].Unknown++ }
}
if ($replications.Count) {
    foreach ($row in $vmRows) { $row.Replicated = if ($replicatedIds.Contains([string]$row.'VM id')) { 'Yes' } else { 'No' } }
}

$resourceRows = foreach ($c in $vdcInfo) {
    $d = $demand[$c.Id]
    $note = if ($d -and $d.Unknown) { "$($d.Unknown) replicated VM(s) outside the configured tenants are not counted" } else { '' }
    $cpuNeed = if ($d) { $d.Cpu } else { $null }
    $ramNeed = if ($d) { [Math]::Round($d.RamGB, 1) } else { $null }
    $cpuAfter = if ($d) { $c.CpuFree - $d.Cpu } else { $null }
    $ramAfter = if ($d) { [Math]::Round($c.RamFreeGB - $d.RamGB, 1) } else { $null }
    [PSCustomObject]@{ Tenant = $c.Tenant; VDC = $c.Name; Model = $c.AllocationModel; Resource = 'vCPU'; Allocated = $c.CpuAllocated; Used = $c.CpuUsed; Free = $c.CpuFree
        'DR demand' = $cpuNeed; 'Free after DR' = $cpuAfter; Note = $note; __Highlight = ($null -ne $cpuAfter -and $cpuAfter -lt 0) }
    [PSCustomObject]@{ Tenant = $c.Tenant; VDC = $c.Name; Model = $c.AllocationModel; Resource = 'RAM GB'; Allocated = $c.RamAllocatedGB; Used = $c.RamUsedGB; Free = $c.RamFreeGB
        'DR demand' = $ramNeed; 'Free after DR' = $ramAfter; Note = $note; __Highlight = ($null -ne $ramAfter -and $ramAfter -lt 0) }
    foreach ($sp in $c.StorageProfiles) {
        $state = if ($sp.Enabled) { '' } else { 'policy disabled' }
        [PSCustomObject]@{ Tenant = $c.Tenant; VDC = $c.Name; Model = $c.AllocationModel; Resource = "$($sp.Name) GB"; Allocated = $sp.LimitGB; Used = $sp.UsedGB; Free = $sp.FreeGB
            'DR demand' = $null; 'Free after DR' = $null; Note = $(if ($sp.LimitGB -eq 0) { 'unlimited' } else { $state }); __Highlight = ($null -ne $sp.FreeGB -and $sp.FreeGB -lt 0) }
    }
}

$sheets = @(
    (New-XlsxSheet -Name 'VMs' -Rows $vmRows -Widths @{ VM = 28; vApp = 24; Disks = 26; 'IP address' = 16; 'Guest OS' = 30; Created = 17; Snapshot = 17; 'VM id' = 38 })
    (New-XlsxSheet -Name 'VDC resources' -Rows @($resourceRows) -Widths @{ VDC = 24; Resource = 24; Note = 40 })
    (New-XlsxSheet -Name 'Disks by policy' -Rows $diskRows -Widths @{ VM = 28 })
)
if ($errors.Count) { $sheets += New-XlsxSheet -Name 'Errors' -Rows $errors -Widths @{ Error = 80; Time = 17 } }
$file = Join-Path $OutputPath ('vcd-vm-report_{0:yyyyMMdd_HHmm}.xlsx' -f (Get-Date))
Export-Xlsx -Path $file -Sheets $sheets
Write-Log "Report: $file ($($vmRows.Count) VMs, $($vdcInfo.Count) VDCs, $($errors.Count) error(s))" -Level Ok

if ($SendMail) {
    $body = "<p>Cloud Director report: $($vmRows.Count) VMs in $($vdcInfo.Count) VDCs of $(@($tenants).Count) tenant(s).</p>"
    if ($errors.Count) { $body += "<p><b>$($errors.Count) tenant(s) could not be read completely</b>, see the Errors sheet.</p>" }
    Send-ToolkitMail -Config $config -Subject ('Cloud Director VM report {0:yyyy-MM-dd}' -f (Get-Date)) -Body $body -BodyAsHtml -Attachment $file
    if ($errors.Count -and $config.Mail.ErrorTo) {
        $html = '<p>The following tenants returned API errors while building the report:</p>' + (ConvertTo-HtmlTable -Rows $errors) +
            '<p>Check API availability, the tenant credential and network access to the Cloud Director endpoint.</p>'
        Send-ToolkitMail -Config $config -To $config.Mail.ErrorTo -Subject ('Cloud Director API errors {0:yyyy-MM-dd}' -f (Get-Date)) -Body $html -BodyAsHtml
    }
}
$file
