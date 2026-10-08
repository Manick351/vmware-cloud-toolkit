<#
.SYNOPSIS
    Finds managed vCenter templates that were left as VMs (for example by an interrupted update) and converts them back.

.DESCRIPTION
    Checks every template listed under VCenters[].Templates in the config (MoRef id or name):

      OK     the object is a template;
      STUCK  the object exists as a VM: power state, size, notes and snapshots are shown;
      LOST   neither a template nor a VM with this id/name exists (recreated template: update the config).

    Without -Fix nothing is changed. With -Fix every STUCK object without snapshots is shut down gracefully
    (VMware Tools; waits up to -ShutdownTimeoutSec) and converted to a template. VMs with snapshots are skipped:
    a template with a snapshot chain is usually not what was intended, remove them first.

.PARAMETER Fix
    Convert stuck VMs back to templates.

.PARAMETER VCenter
    Limit to these configured vCenters.

.PARAMETER ShutdownTimeoutSec
    How long to wait for a guest shutdown. Default: 600.

.PARAMETER ConfigPath
    Toolkit configuration (.psd1). Default: ..\config\toolkit.psd1.

.EXAMPLE
    .\Repair-VCenterTemplate.ps1

.EXAMPLE
    .\Repair-VCenterTemplate.ps1 -Fix -VCenter vc-02
#>
#Requires -Version 5.1
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$Fix,
    [string[]]$VCenter,
    [ValidateRange(60, 3600)][int]$ShutdownTimeoutSec = 600,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config\toolkit.psd1')
)

. (Join-Path $PSScriptRoot '..\lib\Common.ps1')
. (Join-Path $PSScriptRoot '..\lib\VCenter.ps1')

$config = Import-ToolkitConfig -Path $ConfigPath
$null = Start-ToolkitLog -LogPath $config.Settings.LogPath -Name 'Repair-VCenterTemplate'
$targets = @(Select-ConfigVCenter -Config $config -Name $VCenter | Where-Object { $_.Templates })
if (-not $targets) { throw 'No vCenter with a Templates map in the config.' }
if (-not $Fix) { Write-Log 'Detect mode: nothing will be changed (use -Fix to repair).' -Level Warn }

function Get-ManagedObject([string]$Ref, $Server) {
    # Template map values may be MoRef ids (VirtualMachine-vm-123) or names.
    $byId = $Ref -like 'VirtualMachine-*'
    $t = if ($byId) { Get-Template -Id $Ref -Server $Server -ErrorAction SilentlyContinue } else { Get-Template -Name $Ref -Server $Server -ErrorAction SilentlyContinue }
    if ($t) { return @{ State = 'OK'; Object = @($t)[0] } }
    $v = if ($byId) { Get-VM -Id $Ref -Server $Server -ErrorAction SilentlyContinue } else { Get-VM -Name $Ref -Server $Server -ErrorAction SilentlyContinue }
    if ($v) { return @{ State = 'STUCK'; Object = @($v)[0] } }
    return @{ State = 'LOST'; Object = $null }
}

$servers = @()
$rows = New-Object System.Collections.Generic.List[object]
try {
    foreach ($vc in $targets) {
        $server = Connect-ToolkitVCenter -Config $config -VCenter $vc
        if (-not $server) { continue }
        $servers += $server
        foreach ($name in $vc.Templates.Keys) {
            $ref = $vc.Templates[$name]
            $found = Get-ManagedObject $ref $server
            $row = [PSCustomObject]@{ VCenter = $vc.Name; Template = $name; Ref = $ref; State = $found.State; Power = $null; Snapshots = $null; Notes = $null; Action = '' }
            switch ($found.State) {
                'OK' { $row.Notes = $found.Object.ExtensionData.Config.Annotation }
                'STUCK' {
                    $vm = $found.Object
                    $row.Power = $vm.PowerState
                    $row.Snapshots = @(Get-Snapshot -VM $vm -ErrorAction SilentlyContinue).Count
                    $row.Notes = $vm.Notes
                    Write-Log "STUCK $($vc.Name) / ${name}: VM $($vm.Name), $($vm.PowerState), $([Math]::Round($vm.UsedSpaceGB, 1)) GB, $($row.Snapshots) snapshot(s)" -Level Warn
                    if (-not $Fix) { break }
                    if ($row.Snapshots) { $row.Action = 'skipped: has snapshots'; Write-Log '  skipped: remove the snapshots first' -Level Warn; break }
                    if (-not $PSCmdlet.ShouldProcess("$($vc.Name) / $($vm.Name)", 'Shut down and convert to template')) { $row.Action = 'skipped'; break }
                    try {
                        if ($vm.PowerState -ne 'PoweredOff') {
                            $null = Stop-VMGuest -VM $vm -Confirm:$false -ErrorAction Stop
                            $deadline = (Get-Date).AddSeconds($ShutdownTimeoutSec)
                            do { Start-Sleep -Seconds 15; $vm = Get-VM -Id $vm.Id -Server $server } while ($vm.PowerState -ne 'PoweredOff' -and (Get-Date) -lt $deadline)
                            if ($vm.PowerState -ne 'PoweredOff') { throw "guest did not shut down within $ShutdownTimeoutSec s; power it off manually" }
                        }
                        $null = Set-VM -VM $vm -ToTemplate -Confirm:$false -ErrorAction Stop
                        if (-not (Get-Template -Id $vm.Id -Server $server -ErrorAction SilentlyContinue)) { throw 'conversion reported success but the object is not a template' }
                        $row.Action = 'converted'; $row.State = 'FIXED'
                        Write-Log "  converted back to a template" -Level Ok
                    }
                    catch { $row.Action = "failed: $($_.Exception.Message)"; Write-Log "  $($_.Exception.Message)" -Level Error }
                }
                'LOST' { Write-Log "LOST  $($vc.Name) / ${name}: '$ref' not found; update the Templates map" -Level Error }
            }
            $rows.Add($row)
        }
    }
}
finally { Disconnect-ToolkitVCenter $servers }
$rows
if (@($rows | Where-Object { $_.State -in 'STUCK', 'LOST' })) { exit 1 }
