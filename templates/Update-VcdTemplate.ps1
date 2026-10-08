<#
.SYNOPSIS
    Updates a Cloud Director catalog template: deploy, patch, seal, capture back, keep the previous version as "<name>_Old".

.DESCRIPTION
    For each tenant whose TemplateDeploy.Templates lists the template:

      1. resolves the VDC, storage policy, catalog item, the VM inside the vApp template and the Org VDC
         network; stops if a vApp named like the template already exists (leftover of a failed run);
      2. creates a vApp on the network (composeVApp) and adds the template VM with guest customization and the
         static deploy IP (recomposeVApp), then powers it on;
      3. guest credentials: Windows templates with no GuestCredential get an auto-generated administrator
         password from guest customization (read back from the VM; needs the right to view it); otherwise
         GuestCredential is used;
      4. runs the guest part (lib\Guest.ps1) and waits for the VM to power itself off;
      5. undeploys the vApp, deletes "<name>_Old" from the catalog, renames the current catalog item to
         "<name>_Old", captures the vApp into the catalog as "<name>" with "Updated <date>" and deletes the vApp.

    The catalog is only modified after the guest has shut down cleanly; on failure the vApp is left for
    inspection and the catalog is untouched. Without -Execute only step 1 runs (dry run).

.PARAMETER Template
    Template name (Templates section, and the catalog item name).

.PARAMETER Tenant
    Limit to these tenants. Default: every tenant whose TemplateDeploy lists the template.

.PARAMETER Execute
    Perform the update.

.PARAMETER SendMail
    Mail the result with the log attached.

.PARAMETER ConfigPath
    Toolkit configuration (.psd1). Default: ..\config\toolkit.psd1.

.EXAMPLE
    .\Update-VcdTemplate.ps1 -Template W2K22STD

.EXAMPLE
    .\Update-VcdTemplate.ps1 -Template UBUNTU2404 -Tenant tenant-a -Execute -SendMail
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Template,
    [string[]]$Tenant,
    [switch]$Execute,
    [switch]$SendMail,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config\toolkit.psd1')
)

. (Join-Path $PSScriptRoot '..\lib\Common.ps1')
. (Join-Path $PSScriptRoot '..\lib\VcdClient.ps1')
. (Join-Path $PSScriptRoot '..\lib\Guest.ps1')

$config = Import-ToolkitConfig -Path $ConfigPath
$def = Get-TemplateDefinition -Config $config -Name $Template
$log = Start-ToolkitLog -LogPath $config.Settings.LogPath -Name "Update-VcdTemplate_$Template"
$targets = @(Select-ConfigTenant -Config $config -Name $Tenant | Where-Object { $_.TemplateDeploy -and @($_.TemplateDeploy.Templates) -contains $Template })
if (-not $targets) { throw "No selected tenant lists '$Template' in TemplateDeploy.Templates." }
if (-not $Execute) { Write-Log 'DRY RUN: objects are resolved and checked, nothing is changed (use -Execute).' -Level Warn }
$esc = { param($s) [System.Security.SecurityElement]::Escape([string]$s) }

$results = New-Object System.Collections.Generic.List[object]
foreach ($t in $targets) {
    $d = $t.TemplateDeploy
    $started = Get-Date
    $result = [ordered]@{ Tenant = $t.Name; Template = $Template; Catalog = $d.Catalog; Result = ''; Minutes = 0; Detail = '' }
    $vcd = $null
    try {
        $vcd = Connect-VcdTenant -Tenant $t -Config $config

        # --- 1. resolve -------------------------------------------------------------------------------
        $vdcRecs = @(Get-VcdQueryRecord -Session $vcd -Type orgVdc)
        $vdcRec = if ($d.Vdc) { $vdcRecs | Where-Object name -eq $d.Vdc | Select-Object -First 1 } elseif ($vdcRecs.Count -eq 1) { $vdcRecs[0] }
        if (-not $vdcRec) { throw "VDC not determined (TemplateDeploy.Vdc; found: $($vdcRecs.name -join ', '))" }
        $vdc = (Invoke-VcdApi -Session $vcd -Path $vdcRec.href).Vdc
        $profiles = @($vdc.VdcStorageProfiles.VdcStorageProfile | Where-Object { $_ })
        $sp = if ($d.StorageProfile) { $profiles | Where-Object name -eq $d.StorageProfile | Select-Object -First 1 } else { $profiles[0] }
        if (-not $sp) { throw "Storage policy '$($d.StorageProfile)' not in VDC $($vdc.name)" }
        $item = @(Get-VcdQueryRecord -Session $vcd -Type vAppTemplate -Filter "name==$Template;catalogName==$($d.Catalog)") | Select-Object -First 1
        if (-not $item) { throw "vApp template '$Template' not found in catalog '$($d.Catalog)'" }
        $tplVm = @(Get-VcdQueryRecord -Session $vcd -Type vm -Filter "isVAppTemplate==true;container==$($item.href)") | Select-Object -First 1
        if (-not $tplVm) { throw "No VM inside vApp template '$Template'" }
        $net = @(Get-VcdQueryRecord -Session $vcd -Type orgVdcNetwork -Filter "name==$($d.Network)") | Select-Object -First 1
        if (-not $net) { throw "Org VDC network '$($d.Network)' not found" }
        if (-not $d.DeployIp) { throw 'TemplateDeploy.DeployIp is not set' }
        $leftover = @(Get-VcdQueryRecord -Session $vcd -Type vApp -Filter "name==$Template;vdc==$($vdcRec.href)")
        if ($leftover) { throw "vApp '$Template' already exists (leftover of a failed run): inspect and delete it first" }
        $catalog = @(Get-VcdQueryRecord -Session $vcd -Type catalog -Filter "name==$($d.Catalog)") | Select-Object -First 1
        Write-Log "$($t.Name): VDC $($vdc.name), policy $($sp.name), catalog $($d.Catalog), network $($d.Network), deploy IP $($d.DeployIp)"
        if (-not $Execute) { $result.Result = 'DryRun'; $result.Detail = 'all objects resolved'; continue }

        # --- 2. deploy --------------------------------------------------------------------------------
        $compose = @"
<?xml version="1.0" encoding="UTF-8"?>
<ComposeVAppParams name="$(& $esc $Template)" xmlns="http://www.vmware.com/vcloud/v1.5" xmlns:ovf="http://schemas.dmtf.org/ovf/envelope/1">
  <Description>Template update in progress</Description>
  <InstantiationParams>
    <NetworkConfigSection>
      <ovf:Info>Template update network</ovf:Info>
      <NetworkConfig networkName="$(& $esc $d.Network)">
        <Configuration><ParentNetwork href="$($net.href)"/><FenceMode>bridged</FenceMode></Configuration>
      </NetworkConfig>
    </NetworkConfigSection>
  </InstantiationParams>
  <AllEULAsAccepted>true</AllEULAsAccepted>
</ComposeVAppParams>
"@
        $resp = Invoke-VcdApi -Session $vcd -Path "$($vdcRec.href)/action/composeVApp" -Method Post -Body $compose -ContentType 'application/vnd.vmware.vcloud.composeVAppParams+xml'
        $vappHref = $resp.VApp.href
        $null = Wait-VcdTask -Session $vcd -Task (Get-VcdTaskFromResponse $resp) -Activity 'Create vApp'
        Write-Log "vApp created" -Level Ok

        $autoPassword = ($def.OS -eq 'Windows' -and -not $def.GuestCredential)
        $computerName = if ($def.ComputerName) { $def.ComputerName } else { ($Template -replace '[^A-Za-z0-9-]', '').Substring(0, [Math]::Min(15, ($Template -replace '[^A-Za-z0-9-]', '').Length)) }
        $recompose = @"
<?xml version="1.0" encoding="UTF-8"?>
<RecomposeVAppParams xmlns="http://www.vmware.com/vcloud/v1.5" xmlns:ovf="http://schemas.dmtf.org/ovf/envelope/1">
  <SourcedItem>
    <Source href="$($tplVm.href)"/>
    <VmGeneralParams><Name>$(& $esc $Template)</Name><NeedsCustomization>true</NeedsCustomization></VmGeneralParams>
    <InstantiationParams>
      <NetworkConnectionSection ovf:required="false">
        <ovf:Info/>
        <PrimaryNetworkConnectionIndex>0</PrimaryNetworkConnectionIndex>
        <NetworkConnection network="$(& $esc $d.Network)">
          <NetworkConnectionIndex>0</NetworkConnectionIndex>
          <IpAddress>$($d.DeployIp)</IpAddress>
          <IsConnected>true</IsConnected>
          <IpAddressAllocationMode>MANUAL</IpAddressAllocationMode>
        </NetworkConnection>
      </NetworkConnectionSection>
      <GuestCustomizationSection ovf:required="false">
        <ovf:Info/>
        <Enabled>true</Enabled>
        <ChangeSid>true</ChangeSid>
        <AdminPasswordEnabled>$($autoPassword.ToString().ToLower())</AdminPasswordEnabled>
        <AdminPasswordAuto>$($autoPassword.ToString().ToLower())</AdminPasswordAuto>
        <ResetPasswordRequired>false</ResetPasswordRequired>
        <ComputerName>$(& $esc $computerName)</ComputerName>
      </GuestCustomizationSection>
    </InstantiationParams>
    <StorageProfile href="$($sp.href)"/>
  </SourcedItem>
  <AllEULAsAccepted>true</AllEULAsAccepted>
</RecomposeVAppParams>
"@
        $resp = Invoke-VcdApi -Session $vcd -Path "$vappHref/action/recomposeVApp" -Method Post -Body $recompose -ContentType 'application/vnd.vmware.vcloud.recomposeVAppParams+xml'
        $null = Wait-VcdTask -Session $vcd -Task (Get-VcdTaskFromResponse $resp) -Activity 'Add VM from template'
        $vmHref = Wait-Until -TimeoutSec 300 -IntervalSec 15 -Activity 'VM in vApp' -Condition {
            @((Invoke-VcdApi -Session $vcd -Path $vappHref).VApp.Children.Vm)[0].href
        }
        if (-not $vmHref) { throw 'VM did not appear in the vApp' }
        $resp = Invoke-VcdApi -Session $vcd -Path "$vmHref/power/action/powerOn" -Method Post
        $null = Wait-VcdTask -Session $vcd -Task (Get-VcdTaskFromResponse $resp) -Activity 'Power on'
        Write-Log 'VM deployed and powered on' -Level Ok
        Start-Sleep -Seconds 120

        # --- 3-4. guest -------------------------------------------------------------------------------
        $guestCred = if ($def.GuestCredential) { Get-ToolkitCredential -Name $def.GuestCredential -CredentialPath $config.Settings.CredentialPath }
        else {
            $adminPassword = Wait-Until -TimeoutSec 900 -IntervalSec 30 -Activity 'generated admin password' -Condition {
                Get-XmlText (Invoke-VcdApi -Session $vcd -Path "$vmHref/guestCustomizationSection").GuestCustomizationSection.AdminPassword
            }
            if (-not $adminPassword) { throw 'Generated administrator password not readable (missing right "View VM password"?)' }
            New-Object pscredential ("$computerName\Administrator", (ConvertTo-SecureString $adminPassword -AsPlainText -Force))
        }
        $guestLog = if ($def.OS -eq 'Windows') {
            Invoke-WindowsTemplateUpdate -IPAddress $d.DeployIp -Credential $guestCred -Template $def -PSWindowsUpdatePath $config.Settings.PSWindowsUpdatePath
        }
        else { Invoke-LinuxTemplateUpdate -IPAddress $d.DeployIp -Credential $guestCred -Template $def }
        $guestLog | Set-Content -LiteralPath (Join-Path (Split-Path $log) "${Template}_$($t.Name)_guest.txt") -Encoding UTF8

        $off = Wait-Until -TimeoutSec 5400 -IntervalSec 60 -Activity 'guest shutdown' -Condition {
            (Invoke-VcdApi -Session $vcd -Path $vmHref).Vm.status -eq '8'   # 8 = POWERED_OFF
        }
        if (-not $off) { throw "VM did not power off within 90 minutes after sealing; vApp '$Template' left for inspection, catalog untouched" }

        # --- 5. capture -------------------------------------------------------------------------------
        if ((Invoke-VcdApi -Session $vcd -Path $vappHref).VApp.deployed -eq 'true') {
            $resp = Invoke-VcdApi -Session $vcd -Path "$vappHref/action/undeploy" -Method Post -ContentType 'application/vnd.vmware.vcloud.undeployVAppParams+xml' `
                -Body '<UndeployVAppParams xmlns="http://www.vmware.com/vcloud/v1.5"><UndeployPowerAction>powerOff</UndeployPowerAction></UndeployVAppParams>'
            $null = Wait-VcdTask -Session $vcd -Task (Get-VcdTaskFromResponse $resp) -Activity 'Undeploy vApp'
        }
        $oldName = "${Template}_Old"
        foreach ($old in @(Get-VcdQueryRecord -Session $vcd -Type catalogItem -Filter "name==$oldName;catalogName==$($d.Catalog)")) {
            $null = Invoke-VcdApi -Session $vcd -Path $old.href -Method Delete
            Write-Log "Deleted previous '$oldName'"
        }
        $current = (Invoke-VcdApi -Session $vcd -Path $item.href).VAppTemplate
        $current.SetAttribute('name', $oldName)
        $resp = Invoke-VcdApi -Session $vcd -Path $item.href -Method Put -Body $current.OuterXml -ContentType 'application/vnd.vmware.vcloud.vAppTemplate+xml'
        $null = Wait-VcdTask -Session $vcd -Task (Get-VcdTaskFromResponse $resp) -Activity "Rename to $oldName"
        $capture = @"
<?xml version="1.0" encoding="UTF-8"?>
<CaptureVAppParams xmlns="http://www.vmware.com/vcloud/v1.5" name="$(& $esc $Template)">
  <Description>Updated $('{0:yyyy-MM-dd}' -f (Get-Date))</Description>
  <Source href="$vappHref"/>
</CaptureVAppParams>
"@
        $resp = Invoke-VcdApi -Session $vcd -Path "$($catalog.href)/action/captureVApp" -Method Post -Body $capture -ContentType 'application/vnd.vmware.vcloud.captureVAppParams+xml'
        $null = Wait-VcdTask -Session $vcd -Task (Get-VcdTaskFromResponse $resp) -Activity 'Capture to catalog' -TimeoutSec 7200
        $resp = Invoke-VcdApi -Session $vcd -Path $vappHref -Method Delete
        $null = Wait-VcdTask -Session $vcd -Task (Get-VcdTaskFromResponse $resp) -Activity 'Delete vApp'
        $result.Result = 'Updated'
        $result.Detail = "previous version kept as $oldName"
        Write-Log "$($t.Name): $Template captured to '$($d.Catalog)'" -Level Ok
    }
    catch {
        $result.Result = 'Failed'
        $result.Detail = Get-RestErrorText $_
        Write-Log "$($t.Name): $($result.Detail)" -Level Error
    }
    finally {
        $result.Minutes = [int]((Get-Date) - $started).TotalMinutes
        $results.Add([PSCustomObject]$result)
        if ($vcd) { Disconnect-VcdTenant $vcd }
    }
}

$results | Format-Table -AutoSize | Out-String | Write-Host
if ($SendMail -and $Execute) {
    $failed = @($results | Where-Object Result -eq 'Failed').Count
    $subject = "Cloud Director template $Template - " + $(if ($failed) { "$failed failure(s)" } else { 'updated' })
    $files = @($log) + @(Get-ChildItem -LiteralPath (Split-Path $log) -Filter "${Template}_*_guest.txt" | Where-Object LastWriteTime -gt (Get-Date).AddHours(-12) | ForEach-Object FullName)
    Send-ToolkitMail -Config $config -Subject $subject -Body (ConvertTo-HtmlTable -Rows $results) -BodyAsHtml -Attachment $files
}
$results
if (@($results | Where-Object Result -eq 'Failed')) { exit 1 }
