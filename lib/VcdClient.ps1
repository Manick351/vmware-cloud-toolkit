# Minimal REST clients for VMware Cloud Director (tenant side) and VMware Cloud Director Availability.
# Requires lib\Common.ps1. Windows PowerShell 5.1 and PowerShell 7 compatible.
#
#   $vcd = Connect-VcdTenant -Tenant $tenant -Config $config
#   Get-VcdQueryRecord -Session $vcd -Type vm -Filter 'isVAppTemplate==false'
#   Disconnect-VcdTenant $vcd

function Get-HeaderValue {
    # Response header value as a string (PowerShell 7 returns string[]).
    param($Headers, [string]$Name)
    $v = $Headers[$Name]
    if ($v -is [array]) { $v = $v[0] }
    return [string]$v
}

#region Cloud Director

function Connect-VcdTenant {
    <#
    .SYNOPSIS
        Opens a Cloud Director API session for a configured tenant (POST /cloudapi/1.0.0/sessions).
    .DESCRIPTION
        The credential user name may be "user" (the tenant's Org is appended) or "user@org".
        Returns a session object used by the other Vcd* functions.
    #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Tenant,
        [Parameter(Mandatory = $true)][hashtable]$Config,
        [System.Management.Automation.PSCredential]$Credential
    )
    if (-not $Credential) {
        $Credential = Get-ToolkitCredential -Name $Tenant.Credential -CredentialPath $Config.Settings.CredentialPath
    }
    $user = $Credential.UserName
    if ($user -notlike '*@*') {
        if (-not $Tenant.Org) { throw "Tenant '$($Tenant.Name)': user name has no @org and Org is not set." }
        $user = "$user@$($Tenant.Org)"
    }
    $pair = '{0}:{1}' -f $user, $Credential.GetNetworkCredential().Password
    $basic = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($pair))
    $version = if ($Tenant.ApiVersion) { $Tenant.ApiVersion } else { $Config.Settings.VcdApiVersion }
    $skip = [bool]$Config.Settings.SkipCertificateCheck
    $endpoint = if ($Tenant.Provider) { 'sessions/provider' } else { 'sessions' }

    $resp = Invoke-ToolkitRest -Uri "$($Tenant.VcdUrl)/cloudapi/1.0.0/$endpoint" -Method Post -SkipCertificateCheck:$skip `
        -Headers @{ Authorization = "Basic $basic"; Accept = "application/json;version=$version" }
    $token = Get-HeaderValue $resp.Headers 'X-VMWARE-VCLOUD-ACCESS-TOKEN'
    if (-not $token) { throw "Tenant '$($Tenant.Name)': no access token in the session response." }
    $info = $resp.Content | ConvertFrom-Json
    return [PSCustomObject]@{
        Tenant     = $Tenant.Name
        BaseUrl    = $Tenant.VcdUrl
        Token      = $token
        ApiVersion = $version
        Org        = $info.org.name
        User       = $info.user.name
        SessionId  = $info.id
        SkipCert   = $skip
    }
}

function Disconnect-VcdTenant {
    param([Parameter(Mandatory = $true, Position = 0)]$Session)
    if (-not $Session.SessionId) { return }
    try {
        Invoke-ToolkitRest -Uri "$($Session.BaseUrl)/cloudapi/1.0.0/sessions/$($Session.SessionId)" -Method Delete `
            -Headers (Get-VcdHeader $Session 'application/json') -SkipCertificateCheck:$Session.SkipCert -Retries 0 | Out-Null
    }
    catch { }
}

function Get-VcdHeader {
    param($Session, [string]$Accept = 'application/*+xml', [string]$ContentType)
    $h = @{ Authorization = "Bearer $($Session.Token)"; Accept = "$Accept;version=$($Session.ApiVersion)" }
    if ($ContentType) { $h['Content-Type'] = $ContentType }
    return $h
}

function Invoke-VcdApi {
    <#
    .SYNOPSIS
        Calls the legacy XML API. -Path is relative to /api or a full href. Returns [xml] (or $null for empty bodies).
    #>
    param(
        [Parameter(Mandatory = $true)]$Session,
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$Method = 'Get',
        [string]$Body,
        [string]$ContentType
    )
    $uri = if ($Path -match '^https?://') { $Path } else { "$($Session.BaseUrl)/api/$($Path.TrimStart('/'))" }
    $request = @{
        Uri = $uri; Method = $Method; Headers = (Get-VcdHeader $Session 'application/*+xml' $ContentType)
        SkipCertificateCheck = $Session.SkipCert
    }
    if ($Body) { $request.Body = [Text.Encoding]::UTF8.GetBytes($Body) }
    $resp = Invoke-ToolkitRest @request
    if (-not $resp.Content) { return $null }
    $content = if ($resp.Content -is [byte[]]) { [Text.Encoding]::UTF8.GetString($resp.Content) } else { $resp.Content }
    return [xml]$content
}

function Get-VcdQueryRecord {
    <#
    .SYNOPSIS
        Runs a typed query (/api/query?type=...) and returns all record elements across pages.
    .EXAMPLE
        Get-VcdQueryRecord -Session $vcd -Type vm -Filter "isVAppTemplate==false;vdc==$vdcHref"
    #>
    param(
        [Parameter(Mandatory = $true)]$Session,
        [Parameter(Mandatory = $true)][string]$Type,
        [string]$Filter,
        [int]$PageSize = 128
    )
    $page = 1
    $seen = 0
    do {
        $query = "query?type=$Type&format=records&pageSize=$PageSize&page=$page"
        if ($Filter) { $query += '&filter=' + [Uri]::EscapeDataString("($Filter)") }
        $xml = Invoke-VcdApi -Session $Session -Path $query
        $root = $xml.QueryResultRecords
        $total = [int]$root.total
        $records = @($root.ChildNodes | Where-Object { $_.LocalName -like '*Record' })
        foreach ($r in $records) { $r }
        $seen += $records.Count
        $page++
    } while ($records.Count -gt 0 -and $seen -lt $total)
}

function Wait-VcdTask {
    <#
    .SYNOPSIS
        Waits for a Cloud Director task (XML Task element or href) and throws if it does not succeed.
    #>
    param(
        [Parameter(Mandatory = $true)]$Session,
        [Parameter(Mandatory = $true)]$Task,
        [string]$Activity = 'task',
        [int]$TimeoutSec = 3600
    )
    $href = if ($Task -is [string]) { $Task } else { $Task.href }
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    do {
        $t = (Invoke-VcdApi -Session $Session -Path $href).Task
        if ($t.status -notin 'queued', 'preRunning', 'running') { break }
        if ((Get-Date) -gt $deadline) { throw "$Activity did not finish within $TimeoutSec s (task $href)." }
        Start-Sleep -Seconds 5
    } while ($true)
    if ($t.status -ne 'success') {
        $msg = if ($t.Error) { $t.Error.message } else { $t.status }
        throw "$Activity failed: $msg"
    }
    return $t
}

function Get-VcdTaskFromResponse {
    # Returns the first Task element of a response: either the root Task or <Tasks><Task> of an entity.
    param([xml]$Response)
    if (-not $Response) { return $null }
    if ($Response.Task) { return $Response.Task }
    $root = $Response.DocumentElement
    $task = $root.SelectSingleNode("*[local-name()='Tasks']/*[local-name()='Task']")
    return $task
}

function Get-XmlText {
    # Element value whether PowerShell surfaced it as a string (no attributes) or as an XmlElement.
    param($Node)
    if ($null -eq $Node) { return $null }
    if ($Node -is [System.Xml.XmlElement]) { return $Node.InnerText }
    return [string]$Node
}

function Get-VcdVmHardware {
    <#
    .SYNOPSIS
        CPU, RAM and disks of a VM from its full XML (VirtualHardwareSection: ResourceType 3 CPU, 4 RAM, 17 disk).
    #>
    param([Parameter(Mandatory = $true)]$VmXml, [hashtable]$StorageProfileNames = @{})
    $items = @($VmXml.VirtualHardwareSection.Item)
    $cpu = ($items | Where-Object { $_.ResourceType -eq '3' } | Select-Object -First 1).VirtualQuantity
    $ramMb = ($items | Where-Object { $_.ResourceType -eq '4' } | Select-Object -First 1).VirtualQuantity
    $disks = foreach ($d in ($items | Where-Object { $_.ResourceType -eq '17' })) {
        $hr = $d.HostResource
        $profileHref = $hr.storageProfileHref
        [PSCustomObject]@{
            Name           = Get-XmlText $d.ElementName
            CapacityGB     = [Math]::Round([double]$hr.capacity / 1024, 1)
            StorageProfile = if ($StorageProfileNames.ContainsKey([string]$profileHref)) { $StorageProfileNames[[string]$profileHref] } else { $profileHref }
            BusType        = $hr.busSubType
            Iops           = $hr.iops
        }
    }
    return [PSCustomObject]@{
        Cpu   = [int](Get-XmlText $cpu)
        RamGB = [Math]::Round([double](Get-XmlText $ramMb) / 1024, 1)
        Disks = @($disks)
    }
}

function Get-VcdVdcCapacity {
    <#
    .SYNOPSIS
        Allocated / used / free CPU (vCPU), RAM (GB) and storage per policy (GB) of an Org VDC.
    .DESCRIPTION
        CPU in vCPU = MHz / VCpuInMhz2. Storage policies are read one by one (Limit and StorageUsedMB).
        A Limit of 0 means unlimited; Free is then empty.
    #>
    param([Parameter(Mandatory = $true)]$Session, [Parameter(Mandatory = $true)][string]$VdcHref)
    $vdc = (Invoke-VcdApi -Session $Session -Path $VdcHref).Vdc
    $mhz = [double]$vdc.VCpuInMhz2
    if (-not $mhz) { $mhz = 1 }
    $cc = $vdc.ComputeCapacity
    $cpuAlloc = [double]$cc.Cpu.Allocated / $mhz
    $cpuUsed = [double]$cc.Cpu.Used / $mhz
    $ramAlloc = [double]$cc.Memory.Allocated / 1024
    $ramUsed = [double]$cc.Memory.Used / 1024
    $profiles = foreach ($sp in @($vdc.VdcStorageProfiles.VdcStorageProfile)) {
        if (-not $sp) { continue }
        $p = (Invoke-VcdApi -Session $Session -Path $sp.href).VdcStorageProfile
        $limit = [double]$p.Limit / 1024
        $used = [double]$p.StorageUsedMB / 1024
        [PSCustomObject]@{
            Name    = $sp.name
            Href    = $sp.href
            Enabled = $p.Enabled -eq 'true'
            LimitGB = [Math]::Round($limit, 1)
            UsedGB  = [Math]::Round($used, 1)
            FreeGB  = if ($limit -gt 0) { [Math]::Round($limit - $used, 1) } else { $null }
        }
    }
    return [PSCustomObject]@{
        Name            = $vdc.name
        Href            = $VdcHref
        AllocationModel = $vdc.AllocationModel
        CpuAllocated    = [Math]::Round($cpuAlloc, 1)
        CpuUsed         = [Math]::Round($cpuUsed, 1)
        CpuFree         = [Math]::Round($cpuAlloc - $cpuUsed, 1)
        RamAllocatedGB  = [Math]::Round($ramAlloc, 1)
        RamUsedGB       = [Math]::Round($ramUsed, 1)
        RamFreeGB       = [Math]::Round($ramAlloc - $ramUsed, 1)
        StorageProfiles = @($profiles)
    }
}

#endregion

#region Cloud Director Availability

function Connect-Vcav {
    <#
    .SYNOPSIS
        Opens a Cloud Director Availability session with Cloud Director credentials (POST /sessions, type vcdCredentials).
    #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Tenant,
        [Parameter(Mandatory = $true)][hashtable]$Config,
        [System.Management.Automation.PSCredential]$Credential
    )
    if (-not $Tenant.VcavUrl) { throw "Tenant '$($Tenant.Name)' has no VcavUrl." }
    if (-not $Credential) {
        $Credential = Get-ToolkitCredential -Name $Tenant.Credential -CredentialPath $Config.Settings.CredentialPath
    }
    $user = $Credential.UserName
    if ($user -notlike '*@*' -and $Tenant.Org) { $user = "$user@$($Tenant.Org)" }
    $accept = "application/vnd.vmware.h4-v$($Config.Settings.VcavApiVersion)+json;charset=UTF-8"
    $skip = [bool]$Config.Settings.SkipCertificateCheck
    $body = @{ type = 'vcdCredentials'; vcdUser = $user; vcdPassword = $Credential.GetNetworkCredential().Password } | ConvertTo-Json
    $resp = Invoke-ToolkitRest -Uri "$($Tenant.VcavUrl)/sessions" -Method Post -Body $body -ContentType 'application/json' `
        -Headers @{ Accept = $accept } -SkipCertificateCheck:$skip
    $token = Get-HeaderValue $resp.Headers 'X-VCAV-Auth'
    if (-not $token) { throw "Tenant '$($Tenant.Name)': no X-VCAV-Auth token in the session response." }
    return [PSCustomObject]@{
        Tenant   = $Tenant.Name
        BaseUrl  = $Tenant.VcavUrl
        Token    = $token
        Accept   = $accept
        SkipCert = $skip
    }
}

function Disconnect-Vcav {
    param([Parameter(Mandatory = $true, Position = 0)]$Session)
    try { Invoke-VcavApi -Session $Session -Path 'sessions' -Method Delete -Retries 0 | Out-Null } catch { }
}

function Invoke-VcavApi {
    <#
    .SYNOPSIS
        Calls the Cloud Director Availability REST API. Returns parsed JSON.
    #>
    param(
        [Parameter(Mandatory = $true)]$Session,
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$Method = 'Get',
        $Body,
        [int]$Retries = 2
    )
    $request = @{
        Uri = "$($Session.BaseUrl)/$($Path.TrimStart('/'))"; Method = $Method; Retries = $Retries
        Headers = @{ 'X-VCAV-Auth' = $Session.Token; Accept = $Session.Accept }
        SkipCertificateCheck = $Session.SkipCert
    }
    if ($null -ne $Body) {
        $request.Body = [Text.Encoding]::UTF8.GetBytes(($Body | ConvertTo-Json -Depth 10))
        $request.ContentType = 'application/json'
    }
    $resp = Invoke-ToolkitRest @request
    if (-not $resp.Content) { return $null }
    $content = if ($resp.Content -is [byte[]]) { [Text.Encoding]::UTF8.GetString($resp.Content) } else { $resp.Content }
    return $content | ConvertFrom-Json
}

function Get-VcavItem {
    <#
    .SYNOPSIS
        Reads a paged collection (offset/limit, { items, total }) completely.
    .EXAMPLE
        Get-VcavItem -Session $vcav -Path 'vm-replications'
    #>
    param([Parameter(Mandatory = $true)]$Session, [Parameter(Mandatory = $true)][string]$Path, [int]$Limit = 100)
    $offset = 0
    $sep = if ($Path.Contains('?')) { '&' } else { '?' }
    do {
        $page = Invoke-VcavApi -Session $Session -Path "$Path${sep}offset=$offset&limit=$Limit"
        $items = @($page.items)
        foreach ($i in $items) { $i }
        $offset += $Limit
    } while ($items.Count -gt 0 -and $offset -lt [int]$page.total)
}

function Wait-VcavTask {
    # Polls /tasks/<id> until the task leaves RUNNING/QUEUED; throws on failure.
    param([Parameter(Mandatory = $true)]$Session, [Parameter(Mandatory = $true)][string]$TaskId, [int]$TimeoutSec = 1800)
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    do {
        $t = Invoke-VcavApi -Session $Session -Path "tasks/$TaskId"
        if ($t.state -notin 'RUNNING', 'QUEUED', 'STARTED') { break }
        if ((Get-Date) -gt $deadline) { throw "vCAV task $TaskId did not finish within $TimeoutSec s." }
        Start-Sleep -Seconds 5
    } while ($true)
    if ($t.state -ne 'SUCCEEDED') {
        $msg = if ($t.error) { $t.error.msg } else { $t.state }
        throw "vCAV task $TaskId failed: $msg"
    }
    return $t
}

#endregion
