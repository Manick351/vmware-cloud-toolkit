# Shared helpers for vmware-cloud-toolkit scripts. Dot-source: . "$PSScriptRoot\..\lib\Common.ps1"
# Windows PowerShell 5.1 and PowerShell 7 compatible.

#region Configuration

function Import-ToolkitConfig {
    <#
    .SYNOPSIS
        Loads the toolkit configuration (.psd1) and fills in defaults.
    .DESCRIPTION
        See examples\toolkit.example.psd1. Sections: Settings, Mail, Tenants (vCloud Director organizations),
        VCenters, Templates, Replication, DR. Scripts only read the sections they need.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Config file not found: $Path. Copy examples\toolkit.example.psd1 to config\toolkit.psd1 and adjust it."
    }
    $config = Import-PowerShellDataFile -LiteralPath $Path
    foreach ($section in 'Settings', 'Mail', 'Templates', 'Replication', 'DR') {
        if (-not $config.ContainsKey($section)) { $config[$section] = @{} }
    }
    foreach ($section in 'Tenants', 'VCenters') {
        if (-not $config.ContainsKey($section)) { $config[$section] = @() }
    }
    $root = Split-Path -Parent $PSScriptRoot
    $defaults = @{
        LogPath              = (Join-Path $root 'logs')
        ReportPath           = (Join-Path $root 'out')
        CredentialPath       = (Join-Path $root 'config\credentials')
        VcdApiVersion        = '36.0'
        VcavApiVersion       = '4.5'
        SkipCertificateCheck = $false
    }
    foreach ($k in $defaults.Keys) {
        if (-not $config.Settings.ContainsKey($k)) { $config.Settings[$k] = $defaults[$k] }
    }
    foreach ($t in $config.Tenants) {
        foreach ($required in 'Name', 'VcdUrl', 'Credential') {
            if (-not $t[$required]) { throw "Tenant entry without '$required' in $Path." }
        }
        $t.VcdUrl = $t.VcdUrl.TrimEnd('/') -replace '/api$', ''
        if ($t.VcavUrl) { $t.VcavUrl = $t.VcavUrl.TrimEnd('/') }
        if (-not $t.ContainsKey('ExcludeVdcPattern')) { $t.ExcludeVdcPattern = '' }
    }
    foreach ($v in $config.VCenters) {
        foreach ($required in 'Name', 'Credential') {
            if (-not $v[$required]) { throw "vCenter entry without '$required' in $Path." }
        }
    }
    return $config
}

function Select-ConfigTenant {
    # Filters configured tenants by name (all when -Name is empty).
    param([hashtable]$Config, [string[]]$Name)
    $all = @($Config.Tenants)
    if (-not $Name) { return $all }
    $selected = @($all | Where-Object { $Name -contains $_.Name })
    $unknown = @($Name | Where-Object { $all.Name -notcontains $_ })
    if ($unknown) { throw "Unknown tenant(s): $($unknown -join ', '). Known: $($all.Name -join ', ')" }
    return $selected
}

function Select-ConfigVCenter {
    # Filters configured vCenters by FQDN or short name (all when -Name is empty).
    param([hashtable]$Config, [string[]]$Name)
    $all = @($Config.VCenters)
    if (-not $Name) { return $all }
    $selected = @($all | Where-Object { $Name -contains $_.Name -or $Name -contains ($_.Name -split '\.')[0] })
    if (-not $selected) { throw "No configured vCenter matches '$($Name -join ', ')'. Known: $($all.Name -join ', ')" }
    return $selected
}

#endregion

#region Credentials

function Get-ToolkitCredential {
    <#
    .SYNOPSIS
        Resolves a credential by its logical name from the config.
    .DESCRIPTION
        Lookup order:
          1. SecretManagement vault (Get-Secret -Name <name>), if the module is installed;
          2. <CredentialPath>\<name>.xml created with  Get-Credential | Export-Clixml  (DPAPI, same user and machine);
          3. interactive Get-Credential prompt, when the session is interactive.
        Passwords are never read from the config file or the command line.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [string]$CredentialPath,
        [string]$Vault
    )
    if (Get-Command Get-Secret -ErrorAction SilentlyContinue) {
        $secretArgs = @{ Name = $Name; ErrorAction = 'SilentlyContinue' }
        if ($Vault) { $secretArgs.Vault = $Vault }
        $secret = Get-Secret @secretArgs
        if ($secret -is [System.Management.Automation.PSCredential]) { return $secret }
        if ($secret) { throw "Secret '$Name' exists but is not a PSCredential." }
    }
    if ($CredentialPath) {
        $file = Join-Path $CredentialPath "$Name.xml"
        if (Test-Path -LiteralPath $file) { return Import-Clixml -LiteralPath $file }
    }
    if ([Environment]::UserInteractive -and $Host.Name -ne 'Default Host') {
        $cred = Get-Credential -Message "Credential '$Name'"
        if ($cred) { return $cred }
    }
    throw "Credential '$Name' not found in SecretManagement or $CredentialPath."
}

#endregion

#region Logging

$script:ToolkitLogFile = $null

function Start-ToolkitLog {
    # Opens a plain-text log next to the console output. Returns the log path.
    param([string]$LogPath, [string]$Name)
    if (-not (Test-Path -LiteralPath $LogPath)) { New-Item -ItemType Directory -Path $LogPath -Force | Out-Null }
    $script:ToolkitLogFile = Join-Path $LogPath ('{0}_{1:yyyyMMdd_HHmmss}.log' -f $Name, (Get-Date))
    return $script:ToolkitLogFile
}

function Write-Log {
    param(
        [Parameter(Position = 0)][string]$Message,
        [ValidateSet('Info', 'Warn', 'Error', 'Ok', 'Debug')][string]$Level = 'Info'
    )
    $line = '{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}' -f (Get-Date), $Level.ToUpper(), $Message
    $color = @{ Info = 'Gray'; Warn = 'Yellow'; Error = 'Red'; Ok = 'Green'; Debug = 'DarkGray' }[$Level]
    if ($Level -ne 'Debug' -or $VerbosePreference -ne 'SilentlyContinue') { Write-Host $line -ForegroundColor $color }
    if ($script:ToolkitLogFile) {
        try { Add-Content -LiteralPath $script:ToolkitLogFile -Value $line -Encoding UTF8 -WhatIf:$false } catch { }
    }
}

#endregion

#region REST

function Invoke-ToolkitRest {
    <#
    .SYNOPSIS
        Invoke-WebRequest wrapper: TLS 1.2, optional certificate check bypass on 5.1 and 7, retries on 5xx / timeouts.
    .OUTPUTS
        The raw response (use .Content and .Headers). Callers parse JSON or XML themselves.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [string]$Method = 'Get',
        [hashtable]$Headers = @{},
        $Body,
        [string]$ContentType,
        [switch]$SkipCertificateCheck,
        [int]$Retries = 2,
        [int]$TimeoutSec = 120
    )
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $request = @{ Uri = $Uri; Method = $Method; Headers = $Headers; UseBasicParsing = $true; TimeoutSec = $TimeoutSec; ErrorAction = 'Stop' }
    if ($null -ne $Body) { $request.Body = $Body }
    if ($ContentType) { $request.ContentType = $ContentType }
    if ($SkipCertificateCheck) {
        if ($PSVersionTable.PSVersion.Major -ge 6) { $request.SkipCertificateCheck = $true }
        else { Enable-InsecureCertificatePolicy }
    }
    for ($attempt = 0; ; $attempt++) {
        try { return Invoke-WebRequest @request }
        catch {
            $status = 0
            if ($_.Exception.Response) { try { $status = [int]$_.Exception.Response.StatusCode } catch { } }
            $transient = ($status -ge 500 -or $status -eq 0) -and $status -ne 501
            if (-not $transient -or $attempt -ge $Retries) { throw }
            Start-Sleep -Seconds (5 * ($attempt + 1))
        }
    }
}

function Enable-InsecureCertificatePolicy {
    # Windows PowerShell 5.1 has no -SkipCertificateCheck; accept any certificate for this process.
    if ('ToolkitTrustAllPolicy' -as [type]) { return }
    Add-Type -TypeDefinition @'
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class ToolkitTrustAllPolicy : ICertificatePolicy {
    public bool CheckValidationResult(ServicePoint sp, X509Certificate cert, WebRequest req, int problem) { return true; }
}
'@
    [Net.ServicePointManager]::CertificatePolicy = New-Object ToolkitTrustAllPolicy
}

function Get-RestErrorText {
    # Extracts a readable message from a failed web request (vCD returns XML Error, vCAV JSON).
    param($ErrorRecord)
    $text = $ErrorRecord.Exception.Message
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        $detail = $ErrorRecord.ErrorDetails.Message
        if ($detail -match 'message="([^"]+)"') { $detail = $Matches[1] }
        elseif ($detail -match '"msg"\s*:\s*"([^"]+)"') { $detail = $Matches[1] }
        $text = "$text $detail"
    }
    return $text.Trim()
}

#endregion

#region Mail and output

function Send-ToolkitMail {
    <#
    .SYNOPSIS
        Sends an e-mail using the Mail section of the config. Does nothing when Mail.SmtpServer is empty.
    #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Config,
        [Parameter(Mandatory = $true)][string]$Subject,
        [string]$Body = '',
        [string[]]$To,
        [string[]]$Attachment,
        [switch]$BodyAsHtml
    )
    $mail = $Config.Mail
    if (-not $mail.SmtpServer) { Write-Log 'Mail.SmtpServer is not set, e-mail skipped' -Level Warn; return }
    if (-not $To) { $To = @($mail.To) }
    if (-not $To) { Write-Log 'No mail recipients configured, e-mail skipped' -Level Warn; return }
    $send = @{
        SmtpServer = $mail.SmtpServer; From = $mail.From; To = $To; Subject = $Subject; Body = $Body
        Encoding = [System.Text.Encoding]::UTF8; BodyAsHtml = [bool]$BodyAsHtml
    }
    if ($mail.Port) { $send.Port = $mail.Port }
    if ($mail.UseSsl) { $send.UseSsl = $true }
    if ($mail.Credential) {
        $send.Credential = Get-ToolkitCredential -Name $mail.Credential -CredentialPath $Config.Settings.CredentialPath
    }
    $files = @($Attachment | Where-Object { $_ -and (Test-Path -LiteralPath $_) })
    if ($files) { $send.Attachments = $files }
    Send-MailMessage @send -WarningAction SilentlyContinue
    Write-Log "Mail '$Subject' sent to $($To -join ', ')" -Level Ok
}

function ConvertTo-HtmlTable {
    # Small HTML table for mail bodies; values are HTML-encoded.
    param([object[]]$Rows, [string[]]$Columns)
    if (-not $Rows) { return '' }
    if (-not $Columns) { $Columns = @($Rows[0].PSObject.Properties.Name) }
    $enc = { param($v) [System.Net.WebUtility]::HtmlEncode([string]$v) }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<table style="border-collapse:collapse;font-family:Segoe UI,Arial;font-size:12px">')
    [void]$sb.Append('<tr>' + (($Columns | ForEach-Object { "<th style=""background:#2e75b6;color:#fff;padding:4px 8px;text-align:left"">$(& $enc $_)</th>" }) -join '') + '</tr>')
    foreach ($r in $Rows) {
        [void]$sb.Append('<tr>' + (($Columns | ForEach-Object { "<td style=""border:1px solid #ddd;padding:4px 8px"">$(& $enc $r.$_)</td>" }) -join '') + '</tr>')
    }
    [void]$sb.Append('</table>')
    return $sb.ToString()
}

function New-XlsxSheet {
    <#
    .SYNOPSIS
        Builds a sheet definition for Export-Xlsx from objects: columns in property order (union over all rows),
        type 'auto', width from -Widths or from the header length.
    #>
    param([string]$Name, [object[]]$Rows, [hashtable]$Widths = @{})
    $keys = New-Object System.Collections.Generic.List[string]
    foreach ($r in $Rows) {
        foreach ($p in $r.PSObject.Properties.Name) { if ($p -ne '__Highlight' -and -not $keys.Contains($p)) { $keys.Add($p) } }
    }
    $columns = foreach ($k in $keys) {
        $w = if ($Widths.ContainsKey($k)) { $Widths[$k] } else { [Math]::Max(10, [Math]::Min(40, $k.Length + 4)) }
        @{ Key = $k; Width = $w; Type = 'auto' }
    }
    return @{ Name = $Name; Columns = @($columns); Rows = @($Rows); Freeze = $true; Filter = $true }
}

function Get-SafeSheetName {
    # Excel sheet names: max 31 chars, no []:*?/\ ; made unique against $Used.
    param([string]$Name, [System.Collections.Generic.HashSet[string]]$Used)
    $base = ($Name -replace '[\[\]:\*\?/\\]', '_').Trim("'")
    if (-not $base) { $base = 'Sheet' }
    if ($base.Length -gt 31) { $base = $base.Substring(0, 31) }
    $candidate = $base
    $i = 2
    while ($Used -and -not $Used.Add($candidate)) {
        $suffix = "_$i"
        $candidate = $base.Substring(0, [Math]::Min($base.Length, 31 - $suffix.Length)) + $suffix
        $i++
    }
    return $candidate
}

#endregion
