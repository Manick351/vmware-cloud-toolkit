<#
.SYNOPSIS
    Checks whether a set of VMs can be failed over to the DR site: replication health and destination capacity.

.DESCRIPTION
    For the given VM names the script:

      1. finds each VM in the configured Cloud Director tenants (one filtered query per tenant); duplicates are
         reported and the powered-on copy wins;
      2. finds its Cloud Director Availability replication and checks it: overall health, direction (a VM that
         now runs at the destination needs a reverse first), paused state, recovery state, last error, age of
         the latest instance against the RPO, other VMs in the same vApp replication that are not in the list;
      3. determines the destination tenant (-DestinationTenant, else the replication's destination org, else
         DR.TenantPairs from the config) and the destination VDC (the replication's VDC, else the tenant's
         DrVdc, else the only VDC);
      4. per destination VDC compares free vCPU, RAM and storage per policy with what the VMs need, keeping
         DR.ReservePercent of each limit free. Used capacity is the larger of the VDC counters and the sum of
         powered-on VMs.

    Read-only. Exit code 0 when nothing critical was found, 1 otherwise, 2 when the check could not run.

.PARAMETER VMName
    VM names to check (comma/space separated strings are split).

.PARAMETER ListPath
    Text file with VM names, one per line (# comments allowed). Combined with -VMName.

.PARAMETER Exclude
    Names to drop from the list.

.PARAMETER DestinationTenant
    Force the destination tenant for all VMs.

.PARAMETER ReservePercent
    Share of each limit that must stay free after failover. Default: DR.ReservePercent or 5.

.PARAMETER ConfigPath
    Toolkit configuration (.psd1). Default: ..\config\toolkit.psd1.

.PARAMETER PassThru
    Return an object with VMs, replications, capacity rows and findings.

.EXAMPLE
    .\Test-DRCapacity.ps1 -VMName srv-app-01, srv-app-02, srv-db-01

.EXAMPLE
    .\Test-DRCapacity.ps1 -ListPath .\billing-system.txt -Exclude srv-test-01 -ReservePercent 10
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [string[]]$VMName,
    [string]$ListPath,
    [string[]]$Exclude,
    [string]$DestinationTenant,
    [ValidateRange(0, 50)][int]$ReservePercent = -1,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config\toolkit.psd1'),
    [switch]$PassThru
)

. (Join-Path $PSScriptRoot '..\lib\Common.ps1')
. (Join-Path $PSScriptRoot '..\lib\VcdClient.ps1')

$config = Import-ToolkitConfig -Path $ConfigPath
$null = Start-ToolkitLog -LogPath $config.Settings.LogPath -Name 'Test-DRCapacity'
if ($ReservePercent -lt 0) { $ReservePercent = if ($null -ne $config.DR.ReservePercent) { [int]$config.DR.ReservePercent } else { 5 } }
$pairs = if ($config.DR.TenantPairs) { $config.DR.TenantPairs } else { @{} }

function Split-Names([string[]]$Text) {
    @($Text -split '[,;\s]+' | Where-Object { $_ } | ForEach-Object { $_.Trim().ToLower() } | Select-Object -Unique)
}
$names = @(Split-Names $VMName)
if ($ListPath) { $names += Split-Names (Get-Content -LiteralPath $ListPath | Where-Object { $_ -notmatch '^\s*#' }) }
$names = @($names | Select-Object -Unique)
if ($Exclude) { $ex = Split-Names $Exclude; $names = @($names | Where-Object { $ex -notcontains $_ }) }
if (-not $names) { Write-Log 'No VM names given (-VMName or -ListPath).' -Level Error; exit 2 }
if ($DestinationTenant -and -not ($config.Tenants | Where-Object { $_.Name -eq $DestinationTenant })) {
    Write-Log "Unknown destination tenant '$DestinationTenant'." -Level Error; exit 2
}

$findings = New-Object System.Collections.Generic.List[object]
function Add-Finding([ValidateSet('Critical', 'Warn')][string]$Severity, [string]$Text, [string]$Area = 'Capacity') {
    $findings.Add([PSCustomObject]@{ Severity = $Severity; Area = $Area; Text = $Text })
    Write-Log $Text -Level $(if ($Severity -eq 'Critical') { 'Error' } else { 'Warn' })
}
function Get-Uuid([string]$Text) {
    if ($Text -match '([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})') { return $Matches[1].ToLower() }
    return $null
}

Write-Log "DR check for $($names.Count) VM(s): $($names -join ', '); reserve $ReservePercent%"

# --- sessions -------------------------------------------------------------------------------------------
$ctx = @{}   # tenant name -> @{ Tenant; Vcd; Vcav }
foreach ($t in $config.Tenants) {
    $entry = @{ Tenant = $t; Vcd = $null; Vcav = $null }
    try { $entry.Vcd = Connect-VcdTenant -Tenant $t -Config $config }
    catch { Add-Finding Warn "$($t.Name): Cloud Director login failed: $(Get-RestErrorText $_)" 'Access' }
    if ($t.VcavUrl) {
        try { $entry.Vcav = Connect-Vcav -Tenant $t -Config $config }
        catch { Add-Finding Warn "$($t.Name): vCAV login failed: $(Get-RestErrorText $_)" 'Access' }
    }
    $ctx[$t.Name] = $entry
}
$byOrg = @{}
foreach ($e in $ctx.Values) { if ($e.Vcd -and $e.Vcd.Org) { $byOrg[$e.Vcd.Org.ToLower()] = $e.Tenant.Name } }

$exitCode = 2
try {
    if (-not @($ctx.Values | Where-Object { $_.Vcd })) { Write-Log 'No tenant session could be opened.' -Level Error; exit 2 }

    # --- 1. find VMs ------------------------------------------------------------------------------------
    $hits = New-Object System.Collections.Generic.List[object]
    foreach ($e in @($ctx.Values | Where-Object { $_.Vcd })) {
        $filter = 'isVAppTemplate==false;(' + (($names | ForEach-Object { "name==$_" }) -join ',') + ')'
        try {
            foreach ($r in (Get-VcdQueryRecord -Session $e.Vcd -Type vm -Filter $filter)) {
                if ($names -contains $r.name.ToLower()) { $hits.Add([PSCustomObject]@{ Name = $r.name.ToLower(); Tenant = $e.Tenant.Name; Rec = $r }) }
            }
        }
        catch { Add-Finding Warn "$($e.Tenant.Name): VM query failed: $(Get-RestErrorText $_)" 'Access' }
    }

    $vms = New-Object System.Collections.Generic.List[object]
    foreach ($name in $names) {
        $found = @($hits | Where-Object { $_.Name -eq $name })
        if (-not $found) { Add-Finding Critical "${name}: VM not found in any tenant"; continue }
        $h = $found | Sort-Object { if ($_.Rec.status -eq 'POWERED_ON') { 0 } else { 1 } } | Select-Object -First 1
        if ($found.Count -gt 1) {
            Add-Finding Warn ("${name}: {0} VMs with this name ({1}), using the one in {2}" -f $found.Count,
                (($found | ForEach-Object { "$($_.Tenant) $($_.Rec.status)" }) -join ', '), $h.Tenant)
        }
        $vms.Add([PSCustomObject]@{
            Name = $name; Tenant = $h.Tenant; Status = $h.Rec.status; VApp = $h.Rec.containerName
            Cpu = [int]$h.Rec.numberOfCpus; RamMb = [int64]$h.Rec.memoryMB
            DiskMb = [int64]$h.Rec.totalStorageAllocatedMb; Policy = $h.Rec.storageProfileName
            Others = @($found | Where-Object { $_ -ne $h }); Repl = $null; DstTenant = $null; DstHow = ''; DstPolicy = $null
        })
    }

    # --- 2. replications --------------------------------------------------------------------------------
    function Find-Replication([string]$Name) {
        foreach ($e in @($ctx.Values | Where-Object { $_.Vcav })) {
            try { $page = Invoke-VcavApi -Session $e.Vcav -Path "vm-replications?vmName=$([Uri]::EscapeDataString($Name))" }
            catch { continue }
            $it = @($page.items | Where-Object { $_ -and ([string]$_.vmName).ToLower() -eq $Name }) | Select-Object -First 1
            if (-not $it) { continue }
            $ts = $it.destinationState.latestInstance.timestamp
            $last = if ($ts) { [DateTimeOffset]::FromUnixTimeMilliseconds([int64]$ts).LocalDateTime } else { $null }
            return [PSCustomObject]@{
                Session = $e.Vcav; Id = $it.id
                SrcTenant = $byOrg[([string]$it.source.org).ToLower()]; SrcOrg = $it.source.org
                DstTenant = $byOrg[([string]$it.destination.org).ToLower()]; DstOrg = $it.destination.org
                DstVdcId = Get-Uuid ([string]$it.destination.vdcId); DstPolicy = $it.storageProfileName
                Health = [string]$it.overallHealth; Recovery = [string]$it.destinationState.recoveryInfo.recoveryState
                Paused = [bool]$it.isPaused; Rpo = [int]$it.settings.rpo; Last = $last
                AgeMin = if ($last) { [int][Math]::Round(((Get-Date) - $last).TotalMinutes) } else { $null }
                LastError = $it.destinationState.lastError; VAppReplicationId = $it.vappReplicationId
            }
        }
        return $null
    }

    foreach ($vm in $vms) {
        $vm.Repl = Find-Replication $vm.Name
        $fromRepl = if ($vm.Repl -and $vm.Repl.SrcTenant -eq $vm.Tenant) { $vm.Repl.DstTenant } else { $null }
        $vm.DstTenant = @($DestinationTenant, $fromRepl, $pairs[$vm.Tenant]) | Where-Object { $_ } | Select-Object -First 1
        $vm.DstHow = if ($DestinationTenant) { 'parameter' } elseif ($fromRepl) { 'replication' } elseif ($vm.DstTenant) { 'tenant pairs' } else { '' }
        $vm.DstPolicy = if ($vm.Repl -and $vm.Repl.DstPolicy) { $vm.Repl.DstPolicy } else { $vm.Policy }
        Write-Log ('  {0,-18} {1} {2}, vApp {3}: {4} vCPU / {5} GB RAM / {6} GB on {7}' -f $vm.Name, $vm.Tenant, $vm.Status, $vm.VApp,
            $vm.Cpu, [Math]::Round($vm.RamMb / 1024), [Math]::Round($vm.DiskMb / 1024), $vm.Policy)
        if ($DestinationTenant -and $fromRepl -and $fromRepl -ne $DestinationTenant) {
            Add-Finding Warn "$($vm.Name): destination forced to $DestinationTenant, but the replication goes to $fromRepl"
        }
        if (-not $vm.DstTenant) { Add-Finding Critical "$($vm.Name): destination tenant unknown, use -DestinationTenant or DR.TenantPairs" }
        elseif ($vm.DstTenant -eq $vm.Tenant) { Add-Finding Critical "$($vm.Name): destination tenant equals the current one"; $vm.DstTenant = $null }
        elseif (@($vm.Others | Where-Object { $_.Tenant -eq $vm.DstTenant })) {
            Add-Finding Warn "$($vm.Name): $($vm.DstTenant) already has a VM with this name"
        }
    }

    $replRows = New-Object System.Collections.Generic.List[object]
    $vappChecked = @{}
    foreach ($vm in $vms) {
        $r = $vm.Repl
        if (-not $r) {
            Add-Finding Critical "$($vm.Name): no VM replication found" 'Replication'
            $replRows.Add([PSCustomObject]@{ VM = $vm.Name; Direction = '-'; Health = '-'; Recovery = '-'; 'RPO min' = $null; 'Latest instance' = $null; 'Age min' = $null; Result = 'MISSING' })
            continue
        }
        $critBefore = @($findings | Where-Object { $_.Severity -eq 'Critical' }).Count
        $allBefore = $findings.Count
        if ($r.Health -eq 'RED') { Add-Finding Critical "$($vm.Name): replication health RED" 'Replication' }
        elseif ($r.Health -ne 'GREEN') { Add-Finding Warn "$($vm.Name): replication health $($r.Health)" 'Replication' }
        if ($r.SrcTenant -and $r.SrcTenant -ne $vm.Tenant) {
            Add-Finding Critical "$($vm.Name): replication runs from $($r.SrcTenant) but the VM is in $($vm.Tenant); reverse it first" 'Replication'
        }
        if ($r.Recovery -and $r.Recovery -ne 'NOT_STARTED') { Add-Finding Warn "$($vm.Name): recovery state $($r.Recovery)" 'Replication' }
        if ($r.Paused) { Add-Finding Warn "$($vm.Name): replication is paused" 'Replication' }
        if ($r.LastError) { Add-Finding Warn "$($vm.Name): last error $($r.LastError | ConvertTo-Json -Compress -Depth 5)" 'Replication' }
        if (-not $r.Last) { Add-Finding Critical "$($vm.Name): no replica instance yet" 'Replication' }
        elseif ($r.Rpo -gt 0 -and $r.AgeMin -gt 2 * $r.Rpo) { Add-Finding Critical "$($vm.Name): latest instance is $($r.AgeMin) min old, RPO $($r.Rpo) min" 'Replication' }
        elseif ($r.Rpo -gt 0 -and $r.AgeMin -gt $r.Rpo) { Add-Finding Warn "$($vm.Name): latest instance is $($r.AgeMin) min old, above RPO $($r.Rpo) min" 'Replication' }
        if ($r.VAppReplicationId -and -not $vappChecked.ContainsKey($r.VAppReplicationId)) {
            $vappChecked[$r.VAppReplicationId] = $true
            try {
                $vr = Invoke-VcavApi -Session $r.Session -Path "vapp-replications/$($r.VAppReplicationId)"
                $others = @($vr.vmReplications | ForEach-Object { ([string]$_.vmName).ToLower() } | Where-Object { $_ -and $names -notcontains $_ })
                if ($others) { Add-Finding Warn "vApp replication of $($vm.Name) also contains VMs outside this check: $($others -join ', ')" 'Replication' }
            }
            catch { }
        }
        $critAfter = @($findings | Where-Object { $_.Severity -eq 'Critical' }).Count
        $result = if ($critAfter -gt $critBefore) { 'PROBLEM' } elseif ($findings.Count -gt $allBefore) { 'WARNING' } else { 'OK' }
        $replRows.Add([PSCustomObject]@{
            VM = $vm.Name; Direction = "$($r.SrcOrg) -> $($r.DstOrg)"; Health = $r.Health; Recovery = $r.Recovery
            'RPO min' = $r.Rpo; 'Latest instance' = $r.Last; 'Age min' = $r.AgeMin; Result = $result
        })
    }

    # --- 3. destination VDCs ----------------------------------------------------------------------------
    function Resolve-DestinationVdc([string]$TenantName, [string]$VdcId) {
        $e = $ctx[$TenantName]
        if (-not $e -or -not $e.Vcd) { return $null }
        if ($VdcId) { return [PSCustomObject]@{ Ctx = $e; Href = "$($e.Vcd.BaseUrl)/api/vdc/$VdcId"; How = 'from replication' } }
        $recs = @(Get-VcdQueryRecord -Session $e.Vcd -Type orgVdc)
        if ($e.Tenant.ExcludeVdcPattern) { $recs = @($recs | Where-Object { $_.name -notmatch $e.Tenant.ExcludeVdcPattern }) }
        if ($e.Tenant.DrVdc) {
            $pick = @($recs | Where-Object { $_.name -eq $e.Tenant.DrVdc })
            if ($pick) { return [PSCustomObject]@{ Ctx = $e; Href = $pick[0].href; How = 'DrVdc from config' } }
        }
        if (-not $recs) { return $null }
        $how = if ($recs.Count -gt 1) { "first of $($recs.Count) VDCs, set DrVdc in the config" } else { 'only VDC' }
        return [PSCustomObject]@{ Ctx = $e; Href = $recs[0].href; How = $how }
    }

    $groups = [ordered]@{}
    foreach ($vm in @($vms | Where-Object { $_.DstTenant })) {
        $vdcId = if ($vm.Repl -and $vm.Repl.DstTenant -eq $vm.DstTenant) { $vm.Repl.DstVdcId } else { $null }
        $dest = Resolve-DestinationVdc $vm.DstTenant $vdcId
        if (-not $dest) { Add-Finding Critical "$($vm.Name): destination VDC in $($vm.DstTenant) not determined"; continue }
        if (-not $groups.Contains($dest.Href)) { $groups[$dest.Href] = [PSCustomObject]@{ Dest = $dest; Vms = New-Object System.Collections.Generic.List[object] } }
        $groups[$dest.Href].Vms.Add($vm)
    }

    $capacityRows = New-Object System.Collections.Generic.List[object]
    function Test-Resource([string]$Who, [string]$Resource, $Cap, $Used, $Need) {
        if ($null -eq $Cap -or $Cap -le 0) {
            $capacityRows.Add([PSCustomObject]@{ Destination = $Who; Resource = $Resource; Limit = 'unlimited'; Used = $Used; Free = $null; Need = $Need; 'Left after DR' = $null; Result = 'CHECK' })
            Add-Finding Warn "${Who}: $Resource has no limit (allocation model), confirm capacity with the provider"
            return
        }
        $reserve = [Math]::Ceiling($Cap * $ReservePercent / 100)
        $free = $Cap - $Used
        $left = $free - $Need
        $result = if ($left -lt 0) { 'NOT ENOUGH' } elseif ($left -lt $reserve) { 'BELOW RESERVE' } else { 'OK' }
        if ($result -eq 'NOT ENOUGH') { Add-Finding Critical "${Who}: $Resource short by $(-$left)" }
        elseif ($result -eq 'BELOW RESERVE') { Add-Finding Warn "${Who}: $Resource left after DR $left, below the $ReservePercent% reserve ($reserve)" }
        else { Write-Log ('  {0,-22} limit {1}, used {2}, need {3}, left {4}' -f $Resource, $Cap, $Used, $Need, $left) -Level Ok }
        $capacityRows.Add([PSCustomObject]@{ Destination = $Who; Resource = $Resource; Limit = $Cap; Used = $Used; Free = $free; Need = $Need; 'Left after DR' = $left; Result = $result })
    }

    foreach ($g in $groups.Values) {
        $e = $g.Dest.Ctx
        $vdc = (Invoke-VcdApi -Session $e.Vcd -Path $g.Dest.Href).Vdc
        $who = "$($e.Tenant.Name) / $($vdc.name)"
        $in = $g.Vms.ToArray()
        Write-Log "$who ($($vdc.AllocationModel), VDC $($g.Dest.How)): $($in.Name -join ', ')"
        if ($g.Dest.How -like 'first of*') { Add-Finding Warn "${who}: VDC picked automatically ($($g.Dest.How))" }

        $on = @(Get-VcdQueryRecord -Session $e.Vcd -Type vm -Filter "isVAppTemplate==false;status==POWERED_ON;vdc==$($g.Dest.Href)")
        $onCpu = ($on | ForEach-Object { [int]$_.numberOfCpus } | Measure-Object -Sum).Sum
        $onRam = [Math]::Ceiling((($on | ForEach-Object { [int64]$_.memoryMB } | Measure-Object -Sum).Sum) / 1024)

        $mhz = [double]$vdc.VCpuInMhz2
        if ($mhz -le 0) { $mhz = [double]$vdc.VCpuInMhz }
        $cpu = $vdc.ComputeCapacity.Cpu
        $cpuCap = [double]$cpu.Allocated
        if ($cpuCap -le 0) { $cpuCap = [double]$cpu.Limit }
        $capV = if ($cpuCap -gt 0 -and $mhz -gt 0) { [int64][Math]::Floor($cpuCap / $mhz) } else { $null }
        $usedV = if ($mhz -gt 0) { [int64][Math]::Ceiling([double]$cpu.Used / $mhz) } else { 0 }
        Test-Resource $who 'vCPU' $capV ([Math]::Max($usedV, [int64]$onCpu)) ([int]($in | Measure-Object Cpu -Sum).Sum)

        $mem = $vdc.ComputeCapacity.Memory
        $memCap = [double]$mem.Allocated
        if ($memCap -le 0) { $memCap = [double]$mem.Limit }
        $capG = if ($memCap -gt 0) { [int64][Math]::Floor($memCap / 1024) } else { $null }
        $usedG = [int64][Math]::Ceiling([double]$mem.Used / 1024)
        Test-Resource $who 'RAM GB' $capG ([Math]::Max($usedG, [int64]$onRam)) ([int64][Math]::Ceiling((($in | Measure-Object RamMb -Sum).Sum) / 1024))

        $profiles = @{}
        foreach ($sp in @($vdc.VdcStorageProfiles.VdcStorageProfile | Where-Object { $_ })) { $profiles[$sp.name] = $sp.href }
        foreach ($pg in ($in | Group-Object DstPolicy)) {
            $need = [int64][Math]::Ceiling((($pg.Group | Measure-Object DiskMb -Sum).Sum) / 1024)
            $label = "Storage '$($pg.Name)' GB"
            if (-not $profiles.ContainsKey($pg.Name)) {
                Add-Finding Critical "${who}: storage policy '$($pg.Name)' is not available, $need GB needed"
                $capacityRows.Add([PSCustomObject]@{ Destination = $who; Resource = $label; Limit = 'no policy'; Used = $null; Free = $null; Need = $need; 'Left after DR' = $null; Result = 'NOT ENOUGH' })
                continue
            }
            $p = (Invoke-VcdApi -Session $e.Vcd -Path $profiles[$pg.Name]).VdcStorageProfile
            $limit = [int64][Math]::Floor([double]$p.Limit / 1024)
            $used = [int64][Math]::Ceiling([double]$p.StorageUsedMB / 1024)
            $countBefore = $findings.Count
            Test-Resource $who $label $(if ($limit -gt 0) { $limit } else { $null }) $used $need
            # Replica disks already live at the destination and are part of "used"; the shortfall may be overstated.
            $replicated = @($pg.Group | Where-Object { $_.Repl -and $_.Repl.DstTenant -eq $_.DstTenant })
            $row = $capacityRows[$capacityRows.Count - 1]
            if ($replicated -and $row.Result -in 'NOT ENOUGH', 'BELOW RESERVE') {
                for ($k = $countBefore; $k -lt $findings.Count; $k++) { $findings[$k].Severity = 'Warn' }
                $row.Result = 'CHECK'
                Write-Log "  replicas of $($replicated.Name -join ', ') are already stored there and may be counted in 'used'; verify in Cloud Director" -Level Warn
            }
        }
    }

    # --- summary ----------------------------------------------------------------------------------------
    $vmTable = $vms | Select-Object Name, Tenant, Status, Cpu, @{ n = 'RamGB'; e = { [Math]::Round($_.RamMb / 1024) } },
        @{ n = 'DiskGB'; e = { [Math]::Round($_.DiskMb / 1024) } }, Policy, @{ n = 'Destination'; e = { if ($_.DstTenant) { "$($_.DstTenant) ($($_.DstHow))" } } }
    $text = ($vmTable | Format-Table -AutoSize | Out-String -Width 220) + ($replRows | Format-Table -AutoSize | Out-String -Width 220) +
        ($capacityRows | Format-Table -AutoSize | Out-String -Width 220)
    Write-Host $text
    if ($script:ToolkitLogFile) { Add-Content -LiteralPath $script:ToolkitLogFile -Value $text -Encoding UTF8 }

    $crit = @($findings | Where-Object { $_.Severity -eq 'Critical' })
    $capCrit = @($crit | Where-Object { $_.Area -eq 'Capacity' })
    $replCrit = @($crit | Where-Object { $_.Area -eq 'Replication' })
    if ($capCrit) { Write-Log 'CAPACITY: NOT ENOUGH or incomplete data' -Level Error }
    elseif (@($capacityRows | Where-Object Result -ne 'OK')) { Write-Log 'CAPACITY: enough, some items need a manual check' -Level Warn }
    else { Write-Log 'CAPACITY: ENOUGH' -Level Ok }
    if ($replCrit) { Write-Log 'REPLICATION: PROBLEMS - failover now would lose data or is impossible' -Level Error }
    elseif (@($findings | Where-Object { $_.Area -eq 'Replication' })) { Write-Log 'REPLICATION: current, with remarks' -Level Warn }
    else { Write-Log 'REPLICATION: CURRENT' -Level Ok }
    Write-Log "Log: $script:ToolkitLogFile"
    $exitCode = if ($crit) { 1 } else { 0 }

    if ($PassThru) {
        [PSCustomObject]@{ VMs = $vmTable; Replications = $replRows.ToArray(); Capacity = $capacityRows.ToArray(); Findings = $findings.ToArray(); ExitCode = $exitCode }
    }
}
finally {
    foreach ($e in $ctx.Values) {
        if ($e.Vcd) { Disconnect-VcdTenant $e.Vcd }
        if ($e.Vcav) { Disconnect-Vcav $e.Vcav }
    }
}
if (-not $PassThru) { exit $exitCode }
