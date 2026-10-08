<#
.SYNOPSIS
    One inventory of all VMs from vCenters and Cloud Director tenants, cross-checked with Active Directory.

.DESCRIPTION
    Collects VMs from:
      - every configured vCenter (PowerCLI): SRM placeholders and templates are excluded; location comes from
        VCenters[].Locations rules (ESXi host name pattern -> location) or VCenters[].Location;
      - every configured Cloud Director tenant (query records, no per-VM requests); location from Tenants[].Location.

    Then, with -CompareAD, enabled computer accounts matching -ADFilter are compared by short name with the VM
    list: servers that exist in AD but in no hypervisor are physical machines, VMs on unmanaged platforms or
    stale accounts; lastLogonTimestamp helps to tell them apart.

    Workbook sheets: "VMs", "By location", and "AD without VM" with -CompareAD.

.PARAMETER ConfigPath
    Toolkit configuration (.psd1). Default: ..\config\toolkit.psd1.

.PARAMETER SkipVCenter
    Do not query vCenters.

.PARAMETER SkipCloudDirector
    Do not query Cloud Director tenants.

.PARAMETER CompareAD
    Compare with Active Directory computer accounts (ADSI, no module needed).

.PARAMETER ADSearchBase
    LDAP path to search, for example 'OU=Servers,DC=contoso,DC=local'. Default: the current domain.

.PARAMETER ADFilter
    Extra LDAP filter for computers. Default: '(operatingSystem=*Server*)'.

.PARAMETER OutputPath
    Folder for the workbook. Default: Settings.ReportPath.

.PARAMETER SendMail
    Mail the workbook with a per-location summary.

.EXAMPLE
    .\Get-VMInventory.ps1 -CompareAD -ADSearchBase 'OU=Servers,DC=contoso,DC=local'
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config\toolkit.psd1'),
    [switch]$SkipVCenter,
    [switch]$SkipCloudDirector,
    [switch]$CompareAD,
    [string]$ADSearchBase,
    [string]$ADFilter = '(operatingSystem=*Server*)',
    [string]$OutputPath,
    [switch]$SendMail
)

. (Join-Path $PSScriptRoot '..\lib\Common.ps1')
. (Join-Path $PSScriptRoot '..\lib\VcdClient.ps1')
. (Join-Path $PSScriptRoot '..\lib\Xlsx.ps1')

$config = Import-ToolkitConfig -Path $ConfigPath
if (-not $OutputPath) { $OutputPath = $config.Settings.ReportPath }
if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
$null = Start-ToolkitLog -LogPath $config.Settings.LogPath -Name 'Get-VMInventory'

$rows = New-Object System.Collections.Generic.List[object]
$failures = New-Object System.Collections.Generic.List[string]

function Get-HostLocation([hashtable]$VCenter, [string]$HostName) {
    foreach ($rule in @($VCenter.Locations)) {
        if ($rule -and $HostName -like $rule.HostPattern) { return $rule.Location }
    }
    if ($VCenter.Location) { return $VCenter.Location }
    return ($VCenter.Name -split '\.')[0]
}

# --- vCenter ----------------------------------------------------------------------------------------------
if (-not $SkipVCenter -and @($config.VCenters).Count) {
    . (Join-Path $PSScriptRoot '..\lib\VCenter.ps1')
    foreach ($vc in $config.VCenters) {
        $server = Connect-ToolkitVCenter -Config $config -VCenter $vc
        if (-not $server) { $failures.Add("vCenter $($vc.Name)"); continue }
        try {
            $views = Get-View -Server $server -ViewType VirtualMachine -Property Name, Runtime.PowerState, Runtime.Host, Config.Template,
                Config.ManagedBy, Config.Hardware.NumCPU, Config.Hardware.MemoryMB, Config.GuestFullName, Config.Uuid,
                Guest.IpAddress, Guest.HostName, Guest.GuestFullName, Summary.Storage.Committed, Summary.Storage.Uncommitted, Network
            $hostNames = @{}
            foreach ($h in (Get-View -Server $server -ViewType HostSystem -Property Name)) { $hostNames[$h.MoRef.Value] = $h.Name }
            $netNames = @{}
            foreach ($n in (Get-View -Server $server -ViewType Network -Property Name)) { $netNames[$n.MoRef.Value] = $n.Name }
            $count = 0
            foreach ($v in $views) {
                if ($v.Config.Template) { continue }
                if ($v.Config.ManagedBy -and $v.Config.ManagedBy.ExtensionKey -eq 'com.vmware.vcDr') { continue }
                $esx = $hostNames[$v.Runtime.Host.Value]
                $row = [PSCustomObject]@{
                    Source     = 'vCenter'
                    Platform   = $vc.Name
                    Location   = Get-HostLocation $vc $esx
                    Name       = $v.Name
                    'Guest hostname' = $v.Guest.HostName
                    OS         = if ($v.Guest.GuestFullName) { $v.Guest.GuestFullName } else { $v.Config.GuestFullName }
                    Power      = ([string]$v.Runtime.PowerState) -replace '^powered', ''
                    'IP address' = $v.Guest.IpAddress
                    vCPU       = $v.Config.Hardware.NumCPU
                    'RAM GB'   = [Math]::Round($v.Config.Hardware.MemoryMB / 1024, 1)
                    'Disk GB'  = [Math]::Round(($v.Summary.Storage.Committed + $v.Summary.Storage.Uncommitted) / 1GB, 1)
                    Networks   = (@($v.Network | ForEach-Object { $netNames[$_.Value] }) -join ', ')
                    'Host / VDC' = $esx
                    Id         = $v.MoRef.Value
                }
                $rows.Add($row)
                $count++
            }
            Write-Log "  $($vc.Name): $count VM(s)" -Level Ok
        }
        catch { $failures.Add("vCenter $($vc.Name): $($_.Exception.Message)"); Write-Log $_.Exception.Message -Level Error }
        finally { Disconnect-ToolkitVCenter $server }
    }
}

# --- Cloud Director ---------------------------------------------------------------------------------------
if (-not $SkipCloudDirector) {
    foreach ($t in $config.Tenants) {
        $vcd = $null
        try {
            $vcd = Connect-VcdTenant -Tenant $t -Config $config
            $vdcName = @{}
            foreach ($v in (Get-VcdQueryRecord -Session $vcd -Type orgVdc)) { $vdcName[$v.href] = $v.name }
            $count = 0
            foreach ($r in (Get-VcdQueryRecord -Session $vcd -Type vm -Filter 'isVAppTemplate==false')) {
                $row = [PSCustomObject]@{
                    Source     = 'Cloud Director'
                    Platform   = $t.Name
                    Location   = if ($t.Location) { $t.Location } else { $t.Name }
                    Name       = $r.name
                    'Guest hostname' = $null
                    OS         = $r.guestOs
                    Power      = (Get-Culture).TextInfo.ToTitleCase(([string]$r.status -replace '^POWERED_', '').ToLower())
                    'IP address' = $r.ipAddress
                    vCPU       = [int]$r.numberOfCpus
                    'RAM GB'   = [Math]::Round([double]$r.memoryMB / 1024, 1)
                    'Disk GB'  = [Math]::Round([double]$r.totalStorageAllocatedMb / 1024, 1)
                    Networks   = $r.networkName
                    'Host / VDC' = $vdcName[$r.vdc]
                    Id         = ($r.href -split '/vm-')[-1]
                }
                $rows.Add($row)
                $count++
            }
            Write-Log "  $($t.Name): $count VM(s)" -Level Ok
        }
        catch { $failures.Add("tenant $($t.Name): $(Get-RestErrorText $_)"); Write-Log "$($t.Name): $(Get-RestErrorText $_)" -Level Error }
        finally { if ($vcd) { Disconnect-VcdTenant $vcd } }
    }
}

# --- Active Directory -------------------------------------------------------------------------------------
$adRows = @()
if ($CompareAD) {
    # VM names often carry suffixes ("srv-app-01 (old)", "srv-app-01_restore"): compare by the first token.
    $known = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($r in $rows) {
        [void]$known.Add((($r.Name -split '[\s_(]')[0]))
        if ($r.'Guest hostname') { [void]$known.Add(($r.'Guest hostname' -split '\.')[0]) }
    }
    $root = if ($ADSearchBase) { New-Object DirectoryServices.DirectoryEntry("LDAP://$ADSearchBase") } else { New-Object DirectoryServices.DirectoryEntry }
    $searcher = New-Object DirectoryServices.DirectorySearcher($root)
    $searcher.Filter = "(&(objectCategory=computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2))$ADFilter)"
    $searcher.PageSize = 1000
    foreach ($p in 'name', 'dnshostname', 'operatingsystem', 'lastlogontimestamp', 'distinguishedname') { [void]$searcher.PropertiesToLoad.Add($p) }
    $adRows = foreach ($c in $searcher.FindAll()) {
        $name = [string]$c.Properties['name'][0]
        if ($known.Contains($name)) { continue }
        $llt = if ($c.Properties['lastlogontimestamp'].Count) { [DateTime]::FromFileTime([int64]$c.Properties['lastlogontimestamp'][0]) } else { $null }
        [PSCustomObject]@{
            Name          = $name
            'DNS name'    = [string]$c.Properties['dnshostname'][0]
            OS            = [string]$c.Properties['operatingsystem'][0]
            'Last logon'  = $llt
            'Days since logon' = if ($llt) { [int]((Get-Date) - $llt).TotalDays } else { $null }
            OU            = ([string]$c.Properties['distinguishedname'][0]) -replace '^CN=[^,]+,', ''
            __Highlight   = ($null -eq $llt -or ((Get-Date) - $llt).TotalDays -gt 60)
        }
    }
    $adRows = @($adRows | Sort-Object 'Days since logon')
    Write-Log "AD: $($adRows.Count) enabled computer(s) without a VM in the inventory"
}

$byLocation = foreach ($g in ($rows | Group-Object Location | Sort-Object Name)) {
    [PSCustomObject]@{
        Location   = $g.Name
        VMs        = $g.Count
        'Powered on' = @($g.Group | Where-Object Power -eq 'On').Count
        vCPU       = ($g.Group | Measure-Object vCPU -Sum).Sum
        'RAM GB'   = [Math]::Round(($g.Group | Measure-Object 'RAM GB' -Sum).Sum)
        'Disk GB'  = [Math]::Round(($g.Group | Measure-Object 'Disk GB' -Sum).Sum)
        Sources    = (($g.Group.Platform | Select-Object -Unique) -join ', ')
    }
}

$sheets = @(
    (New-XlsxSheet -Name 'VMs' -Rows ($rows | Sort-Object Location, Name) -Widths @{ Name = 26; OS = 32; Networks = 26; 'Host / VDC' = 24; 'Guest hostname' = 26; Id = 38 })
    (New-XlsxSheet -Name 'By location' -Rows @($byLocation) -Widths @{ Sources = 40 })
)
if ($CompareAD) { $sheets += New-XlsxSheet -Name 'AD without VM' -Rows $adRows -Widths @{ 'DNS name' = 30; OS = 30; OU = 50; 'Last logon' = 17 } }
$file = Join-Path $OutputPath ('vm-inventory_{0:yyyyMMdd_HHmm}.xlsx' -f (Get-Date))
Export-Xlsx -Path $file -Sheets $sheets
Write-Log "Inventory: $file ($($rows.Count) VMs)" -Level Ok
foreach ($f in $failures) { Write-Log "Not included: $f" -Level Warn }

if ($SendMail) {
    $body = '<p>VM inventory by location:</p>' + (ConvertTo-HtmlTable -Rows @($byLocation))
    if ($failures.Count) { $body += '<p><b>Not included:</b> ' + [System.Net.WebUtility]::HtmlEncode($failures -join '; ') + '</p>' }
    if ($CompareAD) { $body += "<p>$($adRows.Count) AD computer account(s) without a VM, see the 'AD without VM' sheet.</p>" }
    Send-ToolkitMail -Config $config -Subject ('VM inventory {0:yyyy-MM-dd}' -f (Get-Date)) -Body $body -BodyAsHtml -Attachment $file
}
$file
