<#
.SYNOPSIS
    Determines the Cloud Director version of an endpoint from the tenant side.

.DESCRIPTION
    Tenants usually cannot see the exact build of the provider's Cloud Director, but it can be inferred:

      1. GET /api/versions (no authentication): the highest supported API version maps to a product release;
      2. if /api/versions is redirected to the portal, the login page markup is searched for "version=NN.N";
      3. if nothing was returned, Accept headers with decreasing API versions are sent to /api/session until
         one is answered with 401 (version accepted, not logged in) instead of 406;
      4. with -Authenticated (configured tenant only): /cloudapi/1.0.0/site and /api/session are checked for
         version fields.

    The result is the product release that introduced the highest API version. A later patch of the same
    release reports the same API version, so treat the patch level as a lower bound.

.PARAMETER Url
    Cloud Director address(es). Any URL on the host works (for example a tenant portal link); it is reduced
    to scheme://host.

.PARAMETER Tenant
    Configured tenant name(s) instead of -Url.

.PARAMETER ConfigPath
    Toolkit configuration (.psd1), used with -Tenant. Default: ..\config\toolkit.psd1.

.PARAMETER Authenticated
    Also log in (with -Tenant) and query authenticated endpoints.

.PARAMETER SkipCertificateCheck
    Accept self-signed certificates (lab endpoints).

.EXAMPLE
    .\Get-VcdVersion.ps1 -Url https://vcd.example.com/tenant/contoso

.EXAMPLE
    .\Get-VcdVersion.ps1 -Tenant tenant-a, tenant-b -Authenticated | Format-Table
#>
#Requires -Version 5.1
[CmdletBinding(DefaultParameterSetName = 'Url')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Url', Position = 0)][string[]]$Url,
    [Parameter(Mandatory = $true, ParameterSetName = 'Tenant')][string[]]$Tenant,
    [Parameter(ParameterSetName = 'Tenant')][string]$ConfigPath = (Join-Path $PSScriptRoot '..\config\toolkit.psd1'),
    [Parameter(ParameterSetName = 'Tenant')][switch]$Authenticated,
    [switch]$SkipCertificateCheck
)

. (Join-Path $PSScriptRoot '..\lib\Common.ps1')
. (Join-Path $PSScriptRoot '..\lib\VcdClient.ps1')

# Highest API version -> release that introduced it.
$apiToProduct = [ordered]@{
    '39.1' = '10.6.1'; '39.0' = '10.6'
    '38.1' = '10.5.1'; '38.0' = '10.5'
    '37.3' = '10.4.3'; '37.2' = '10.4.2'; '37.1' = '10.4.1'; '37.0' = '10.4'
    '36.3' = '10.3.3'; '36.2' = '10.3.2'; '36.1' = '10.3.1'; '36.0' = '10.3'
    '35.2' = '10.2.2'; '35.0' = '10.2'
    '34.0' = '10.1'; '33.0' = '10.0'; '32.0' = '9.7'; '31.0' = '9.5'; '30.0' = '9.1'; '29.0' = '9.0'; '27.0' = '8.20'
}

function Get-ProductForApi([string]$Api) {
    if ($apiToProduct.Contains($Api)) { return $apiToProduct[$Api] }
    $lower = @($apiToProduct.Keys | Where-Object { [version]$_ -le [version]$Api } | Sort-Object { [version]$_ })
    if ($lower) { return "$($apiToProduct[$lower[-1]]) or later" }
    return $null
}

function Invoke-Probe([string]$Uri, [hashtable]$Headers = @{}) {
    try {
        $r = Invoke-ToolkitRest -Uri $Uri -Headers $Headers -SkipCertificateCheck:$SkipCertificateCheck -Retries 0 -TimeoutSec 30
        return [PSCustomObject]@{ Ok = $true; Code = [int]$r.StatusCode; Content = [string]$r.Content }
    }
    catch {
        $code = 0
        if ($_.Exception.Response) { try { $code = [int]$_.Exception.Response.StatusCode } catch { } }
        return [PSCustomObject]@{ Ok = $false; Code = $code; Content = $_.Exception.Message }
    }
}

function Test-VcdEndpoint([string]$BaseUrl, $Session) {
    $findings = New-Object System.Collections.Generic.List[string]
    $maxApi = $null

    $resp = Invoke-Probe "$BaseUrl/api/versions"
    if ($resp.Ok -and $resp.Content -match '<SupportedVersions') {
        $xml = [xml]$resp.Content
        $versions = @($xml.SupportedVersions.VersionInfo | ForEach-Object { Get-XmlText $_.Version } | Where-Object { $_ -match '^\d+\.\d+$' })
        if ($versions) {
            $maxApi = ($versions | Sort-Object { [version]$_ })[-1]
            $findings.Add("/api/versions: $($versions.Count) versions, highest $maxApi")
        }
    }
    elseif ($resp.Ok) {
        $hits = @([regex]::Matches($resp.Content, '(?i)version=(\d{2}\.\d)') | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
        if ($hits) {
            $maxApi = ($hits | Sort-Object { [version]$_ })[-1]
            $findings.Add("portal markup: version=$($hits -join ', ')")
        }
        else { $findings.Add('/api/versions returned HTML without version markers') }
    }
    else { $findings.Add("/api/versions: HTTP $($resp.Code)") }

    if (-not $maxApi) {
        foreach ($v in @($apiToProduct.Keys | Sort-Object { [version]$_ } -Descending)) {
            $p = Invoke-Probe "$BaseUrl/api/session" @{ Accept = "application/*+xml;version=$v" }
            # 406 = version not supported; 401/403/200 = version accepted (authentication is a separate matter)
            if ($p.Code -in 200, 401, 403) {
                $maxApi = $v
                $findings.Add("Accept header probe: $v accepted (HTTP $($p.Code))")
                break
            }
        }
    }

    $reported = $null
    if ($Session) {
        foreach ($ep in @(
                @{ Path = 'cloudapi/1.0.0/site'; Accept = 'application/json' }
                @{ Path = 'api/session'; Accept = 'application/*+xml' })) {
            $p = Invoke-Probe "$BaseUrl/$($ep.Path)" @{ Authorization = "Bearer $($Session.Token)"; Accept = "$($ep.Accept);version=$($Session.ApiVersion)" }
            if ($p.Ok) {
                $m = [regex]::Match($p.Content, '(?i)"?(product|build)Version"?\s*[:=]\s*"?(\d+\.\d+[\.\d]*)')
                if ($m.Success) { $reported = $m.Groups[2].Value; $findings.Add("$($ep.Path): $($m.Groups[1].Value)Version $reported") }
                else { $findings.Add("$($ep.Path): no version fields") }
            }
            else { $findings.Add("$($ep.Path): HTTP $($p.Code)") }
        }
    }

    [PSCustomObject]@{
        Url             = $BaseUrl
        MaxApiVersion   = $maxApi
        Product         = if ($reported) { $reported } elseif ($maxApi) { Get-ProductForApi $maxApi } else { $null }
        Basis           = if ($reported) { 'reported by API' } elseif ($maxApi) { 'highest API version' } else { 'not determined' }
        Details         = $findings -join '; '
    }
}

if ($PSCmdlet.ParameterSetName -eq 'Tenant') {
    $config = Import-ToolkitConfig -Path $ConfigPath
    if ($config.Settings.SkipCertificateCheck) { $SkipCertificateCheck = $true }
    foreach ($t in (Select-ConfigTenant -Config $config -Name $Tenant)) {
        $session = $null
        if ($Authenticated) {
            try { $session = Connect-VcdTenant -Tenant $t -Config $config }
            catch { Write-Warning "$($t.Name): login failed: $(Get-RestErrorText $_)" }
        }
        try { Test-VcdEndpoint -BaseUrl $t.VcdUrl -Session $session | Add-Member -NotePropertyName Tenant -NotePropertyValue $t.Name -PassThru }
        finally { if ($session) { Disconnect-VcdTenant $session } }
    }
}
else {
    foreach ($u in $Url) {
        $uri = [uri]$u
        Test-VcdEndpoint -BaseUrl ('{0}://{1}' -f $uri.Scheme, $uri.Authority)
    }
}
