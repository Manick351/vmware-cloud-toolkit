<#
.SYNOPSIS
    Updates a vCenter VM template in place: convert to VM, boot with a customization spec, patch, seal, convert back.

.DESCRIPTION
    For each selected vCenter that has the template in its Templates map:

      1. resolves the template (MoRef id or name), the OS customization spec and the deploy IP; stops if the
         template is currently a VM (a previous run did not finish, see Repair-VCenterTemplate.ps1), the spec
         is missing or the deploy IP already answers to ping;
      2. converts the template to a VM, applies the customization spec (it must give the guest the static
         deploy IP), connects the NIC at power on and starts the VM;
      3. runs the guest part (lib\Guest.ps1): Windows - updates in rounds, hooks, cleanup, sysprep;
         Linux - package upgrade, hooks, sealing;
      4. waits for the guest to power itself off (sysprep / scheduled poweroff), with one graceful shutdown
         attempt as a fallback, writes "Updated <date>" into the notes and converts the VM back to a template.

    If anything fails after step 2, the script tries to leave the object as a template again (power off,
    convert) and reports what was done. vCenters are processed one after another.

    Without -Execute the script only resolves and validates everything (dry run).

.PARAMETER Template
    Template name from the Templates section of the config.

.PARAMETER VCenter
    Limit to these configured vCenters. Default: every vCenter whose Templates map contains the template.

.PARAMETER Execute
    Perform the update. Without it the run is a dry run.

.PARAMETER SendMail
    Mail the result with the log attached (Mail section).

.PARAMETER ConfigPath
    Toolkit configuration (.psd1). Default: ..\config\toolkit.psd1.

.EXAMPLE
    .\Update-VcenterTemplate.ps1 -Template W2K22STD

.EXAMPLE
    .\Update-VcenterTemplate.ps1 -Template UBUNTU2404 -VCenter vc-01 -Execute -SendMail
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Template,
    [string[]]$VCenter,
    [switch]$Execute,
    [switch]$SendMail,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config\toolkit.psd1')
)

. (Join-Path $PSScriptRoot '..\lib\Common.ps1')
. (Join-Path $PSScriptRoot '..\lib\VCenter.ps1')
. (Join-Path $PSScriptRoot '..\lib\Guest.ps1')

$config = Import-ToolkitConfig -Path $ConfigPath
$def = Get-TemplateDefinition -Config $config -Name $Template
$log = Start-ToolkitLog -LogPath $config.Settings.LogPath -Name "Update-VcenterTemplate_$Template"
$targets = @(Select-ConfigVCenter -Config $config -Name $VCenter | Where-Object { $_.Templates -and $_.Templates.ContainsKey($Template) })
if (-not $targets) { throw "No selected vCenter has '$Template' in its Templates map." }
if (-not $def.CustomizationSpec) { throw "Template '$Template' has no CustomizationSpec." }
if (-not $Execute) { Write-Log 'DRY RUN: objects are resolved and checked, nothing is changed (use -Execute).' -Level Warn }
$guestCred = if ($def.GuestCredential) { Get-ToolkitCredential -Name $def.GuestCredential -CredentialPath $config.Settings.CredentialPath }
if (-not $guestCred) { throw "Template '$Template' has no GuestCredential." }

$results = New-Object System.Collections.Generic.List[object]
foreach ($vc in $targets) {
    $started = Get-Date
    $result = [ordered]@{ VCenter = $vc.Name; Template = $Template; Result = ''; Minutes = 0; Detail = '' }
    $server = $null; $vm = $null; $converted = $false
    try {
        $server = Connect-ToolkitVCenter -Config $config -VCenter $vc
        if (-not $server) { throw 'not connected' }
        $ref = $vc.Templates[$Template]
        $byId = $ref -like 'VirtualMachine-*'
        $tpl = @(if ($byId) { Get-Template -Id $ref -Server $server -ErrorAction SilentlyContinue } else { Get-Template -Name $ref -Server $server -ErrorAction SilentlyContinue })
        if ($tpl.Count -ne 1) {
            $asVm = @(if ($byId) { Get-VM -Id $ref -Server $server -ErrorAction SilentlyContinue } else { Get-VM -Name $ref -Server $server -ErrorAction SilentlyContinue })
            if ($asVm) { throw "'$ref' is currently a VM ($($asVm[0].PowerState)), not a template: a previous run did not finish. Run Repair-VCenterTemplate.ps1." }
            throw "Template '$ref' not found (or not unique). Recreated templates get a new MoRef: update the config."
        }
        $tpl = $tpl[0]
        $spec = Get-OSCustomizationSpec -Name $def.CustomizationSpec -Server $server -ErrorAction SilentlyContinue
        if (-not $spec) { throw "Customization spec '$($def.CustomizationSpec)' not found." }
        $ip = $vc.DeployIp
        if (-not $ip) { throw "vCenter $($vc.Name) has no DeployIp." }
        Write-Log "$($vc.Name): $($tpl.Name) ($($tpl.Id)), note '$($tpl.ExtensionData.Config.Annotation)', tools $($tpl.ExtensionData.Guest.ToolsVersion); spec $($spec.Name); deploy IP $ip"
        if ([int]$tpl.ExtensionData.Guest.ToolsVersion -eq 0) { Write-Log 'VMware Tools never reported: customization and graceful shutdown may fail' -Level Warn }
        if (Test-Connection -ComputerName $ip -Count 2 -Quiet -ErrorAction SilentlyContinue) { throw "Deploy IP $ip already answers to ping." }
        if (-not $Execute) { $result.Result = 'DryRun'; $result.Detail = 'all objects resolved'; continue }

        Write-Log 'Converting template to VM and starting it'
        $vm = Set-Template -Template $tpl -ToVM -Confirm:$false -ErrorAction Stop
        $converted = $true
        $vm = Set-VM -VM $vm -OSCustomizationSpec $spec -Confirm:$false -ErrorAction Stop
        $null = Get-NetworkAdapter -VM $vm | Set-NetworkAdapter -StartConnected:$true -Confirm:$false
        $vm = Start-VM -VM $vm -Confirm:$false -ErrorAction Stop
        Start-Sleep -Seconds 180

        $guestLog = if ($def.OS -eq 'Windows') {
            Invoke-WindowsTemplateUpdate -IPAddress $ip -Credential $guestCred -Template $def -PSWindowsUpdatePath $config.Settings.PSWindowsUpdatePath
        }
        else { Invoke-LinuxTemplateUpdate -IPAddress $ip -Credential $guestCred -Template $def }
        $guestFile = Join-Path (Split-Path $log) ("{0}_{1}_guest.txt" -f $Template, ($vc.Name -split '\.')[0])
        $guestLog | Set-Content -LiteralPath $guestFile -Encoding UTF8

        Write-Log 'Waiting for the guest to power off'
        $sealedAt = Get-Date
        $off = Wait-Until -TimeoutSec 5400 -IntervalSec 30 -Activity 'guest shutdown' -Condition { (Get-VM -Id $vm.Id -Server $server).PowerState -eq 'PoweredOff' }
        if (-not $off) {
            Write-Log 'Guest did not power off in 90 minutes, trying a graceful shutdown' -Level Warn
            $null = Stop-VMGuest -VM (Get-VM -Id $vm.Id -Server $server) -Confirm:$false -ErrorAction SilentlyContinue
            $off = Wait-Until -TimeoutSec 600 -IntervalSec 30 -Condition { (Get-VM -Id $vm.Id -Server $server).PowerState -eq 'PoweredOff' }
            if (-not $off) { throw 'VM is still running; it was left as a VM for inspection.' }
            $result.Detail = 'sealing NOT confirmed (shutdown forced by timeout handler); verify a clone before use; '
        }
        else { $result.Detail = "guest powered off $([int]((Get-Date) - $sealedAt).TotalMinutes) min after sealing; " }

        $vm = Set-VM -VM (Get-VM -Id $vm.Id -Server $server) -Notes ('Updated {0:yyyy-MM-dd}' -f (Get-Date)) -Confirm:$false
        $null = Set-VM -VM $vm -ToTemplate -Confirm:$false -ErrorAction Stop
        $converted = $false
        $result.Result = 'Updated'
        Write-Log "$($vc.Name): $Template updated" -Level Ok
    }
    catch {
        $result.Result = 'Failed'
        $result.Detail += $_.Exception.Message
        Write-Log "$($vc.Name): $($_.Exception.Message)" -Level Error
        if ($converted -and $vm) {
            try {
                $cur = Get-VM -Id $vm.Id -Server $server
                if ($cur.PowerState -ne 'PoweredOff') { $null = Stop-VM -VM $cur -Confirm:$false -ErrorAction Stop }
                $null = Set-VM -VM (Get-VM -Id $vm.Id -Server $server) -ToTemplate -Confirm:$false -ErrorAction Stop
                $result.Detail += ' | converted back to a template without the update'
                Write-Log 'Rolled back: the object is a template again (not updated)' -Level Warn
            }
            catch { $result.Detail += ' | left as a VM, run Repair-VCenterTemplate.ps1'; Write-Log 'Could not convert back; run Repair-VCenterTemplate.ps1' -Level Error }
        }
    }
    finally {
        $result.Minutes = [int]((Get-Date) - $started).TotalMinutes
        $results.Add([PSCustomObject]$result)
        if ($server) { Disconnect-ToolkitVCenter $server }
    }
}

$results | Format-Table -AutoSize | Out-String | Write-Host
if ($SendMail -and $Execute) {
    $failed = @($results | Where-Object Result -eq 'Failed').Count
    $subject = "vCenter template $Template - " + $(if ($failed) { "$failed failure(s)" } else { 'updated' })
    $files = @($log) + @(Get-ChildItem -LiteralPath (Split-Path $log) -Filter "${Template}_*_guest.txt" | Where-Object LastWriteTime -gt (Get-Date).AddHours(-12) | ForEach-Object FullName)
    Send-ToolkitMail -Config $config -Subject $subject -Body (ConvertTo-HtmlTable -Rows $results) -BodyAsHtml -Attachment $files
}
$results
if (@($results | Where-Object Result -eq 'Failed')) { exit 1 }
