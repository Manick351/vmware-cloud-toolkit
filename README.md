# vmware-cloud-toolkit

PowerShell tools for VMware Cloud Director (tenant side), Cloud Director Availability and vCenter:
VM and capacity reports, DR readiness checks, replication management, snapshot hygiene and automated
updates of VM templates in vCenter and in Cloud Director catalogs.

> **🇷🇺 Кратко.** Инструменты PowerShell для VMware Cloud Director (со стороны тенанта), Cloud Director
> Availability и vCenter: отчёты по ВМ и ресурсам VDC, проверка готовности к DR (свежесть реплик и хватит ли
> ресурсов на резервной площадке), создание и удаление репликаций, снапшоты по нескольким vCenter с защитой
> от SRM-плейсхолдеров, автоматическое обновление шаблонов ВМ в vCenter и в каталогах vCD. Все параметры
> площадок — в одном конфиге, пароли — только через SecretManagement или DPAPI.

## Contents

| Area | Script | What it does | Changes anything? |
|---|---|---|---|
| Reports | [reports/Get-VcdVmReport.ps1](reports/Get-VcdVmReport.ps1) | VMs of all tenants (vCPU, RAM, disks per storage policy, IP, OS, snapshot, replicated), VDC capacity with failover demand from vCAV, XLSX + mail | No |
| | [reports/Get-VcdDiskIops.ps1](reports/Get-VcdDiskIops.ps1) | Per-disk IOPS limits and a per-policy summary with the policy's IOPS settings | No |
| | [reports/Get-VMInventory.ps1](reports/Get-VMInventory.ps1) | One inventory from vCenters and Cloud Director, locations by host pattern, AD computers without a VM | No |
| | [reports/Get-VcdVersion.ps1](reports/Get-VcdVersion.ps1) | Cloud Director version of a provider endpoint, inferred from the tenant side | No |
| DR | [dr/Test-DRCapacity.ps1](dr/Test-DRCapacity.ps1) | Can these VMs fail over now? Replication health, RPO age, direction; destination vCPU / RAM / storage with a reserve | No |
| | [dr/New-VcavReplication.ps1](dr/New-VcavReplication.ps1) | Creates vCAV replications from replication pairs and settings in the config, optional move into a DR vApp | `-WhatIf` / `-Confirm` |
| | [dr/Remove-VcavReplication.ps1](dr/Remove-VcavReplication.ps1) | Removes replications, optional move back from the DR vApp | `-WhatIf` / `-Confirm` (High) |
| Snapshots | [snapshots/Invoke-VMSnapshot.ps1](snapshots/Invoke-VMSnapshot.ps1) | List / create / remove snapshots across vCenters; ignores SRM placeholders, refuses ambiguous names, removes only its own snapshots by default | `-WhatIf` / `-Confirm` |
| Templates | [templates/Update-VcenterTemplate.ps1](templates/Update-VcenterTemplate.ps1) | Template -> VM -> customization spec -> patch -> seal -> template, with rollback | Dry run unless `-Execute` |
| | [templates/Update-VcdTemplate.ps1](templates/Update-VcdTemplate.ps1) | Deploy from catalog -> patch -> seal -> capture back, previous version kept as `<name>_Old` | Dry run unless `-Execute` |
| | [templates/Repair-VCenterTemplate.ps1](templates/Repair-VCenterTemplate.ps1) | Finds templates left as VMs and converts them back | Only with `-Fix` |
| | [templates/Copy-VCenterTemplate.ps1](templates/Copy-VCenterTemplate.ps1) | Copies a template to another vCenter with full pre-checks | `-WhatIf` |
| | [templates/Import-OvfTemplate.ps1](templates/Import-OvfTemplate.ps1) | OVF/OVA import with PowerCLI `Import-VApp` (no password in an ovftool URL) | `-WhatIf` |

Libraries in [lib/](lib): a small Cloud Director / vCAV REST client (`VcdClient.ps1`), PowerCLI helpers with safe
VM name resolution (`VCenter.ps1`), guest OS steps of template updates (`Guest.ps1`), an XLSX writer that needs
neither Excel nor ImportExcel (`Xlsx.ps1`), configuration, credentials, logging and mail (`Common.ps1`).

## Requirements

- Windows PowerShell 5.1 or PowerShell 7. The REST scripts also run on Linux / macOS with PowerShell 7
  (the AD comparison of `Get-VMInventory.ps1` needs Windows).
- Cloud Director 10.x tenant account (API version is configurable, default 36.0), optional Cloud Director
  Availability 4.x on the same credentials.
- vCenter scripts: [VMware PowerCLI](https://developer.broadcom.com/powercli) (`Install-Module VMware.PowerCLI`).
- Template updates: WinRM from the admin host to the deploy IP (Windows guests, PSWindowsUpdate is installed or
  copied automatically), [Posh-SSH](https://github.com/darkoperator/Posh-SSH) for Linux guests.
- Optional: `Microsoft.PowerShell.SecretManagement` with any vault for credentials.

## Setup

```powershell
Copy-Item examples\toolkit.example.psd1 config\toolkit.psd1      # config\ is git-ignored
notepad config\toolkit.psd1

# credentials are referenced by name; either a SecretManagement secret ...
Set-Secret -Name vcd-tenant-a -Secret (Get-Credential svc-reports)
# ... or a DPAPI-protected file for the same user on the same machine
New-Item config\credentials -ItemType Directory
Get-Credential svc-reports | Export-Clixml config\credentials\vcd-tenant-a.xml
```

The configuration ([examples/toolkit.example.psd1](examples/toolkit.example.psd1)) holds tenants (Cloud Director
URL, org, vCAV URL, DR VDC, template deploy settings), vCenters (deploy IP, template MoRef map, location rules),
template definitions with guest hooks, replication pairs and settings, the DR reserve. Every script has
comment-based help: `Get-Help .\dr\Test-DRCapacity.ps1 -Full`.

## Examples

```powershell
# monthly report for all tenants, mailed to Mail.To; API errors go to Mail.ErrorTo
.\reports\Get-VcdVmReport.ps1 -SendMail

# before a DR test: are replicas fresh and is there room at the destination?
.\dr\Test-DRCapacity.ps1 -ListPath .\systems\billing.txt -ReservePercent 10

# protect two VMs with the settings from the config (preview first)
.\dr\New-VcavReplication.ps1 -VMName srv-app-01, srv-app-02 -WhatIf

# snapshot before a change, remove after it
.\snapshots\Invoke-VMSnapshot.ps1 -Action Create -VMName srv-app-01, srv-db-01 -Description 'CHG-1234'
.\snapshots\Invoke-VMSnapshot.ps1 -Action Remove -VMName srv-app-01, srv-db-01

# forgotten snapshots older than a week, everywhere
.\snapshots\Invoke-VMSnapshot.ps1 -Action List -AllVMs -OlderThanDays 7

# template update: dry run, then the real run (schedule it monthly with Register-ScriptTask from windows-server-ops)
.\templates\Update-VcenterTemplate.ps1 -Template W2K22STD
.\templates\Update-VcenterTemplate.ps1 -Template W2K22STD -Execute -SendMail
```

`Test-DRCapacity.ps1` output (illustrative):

```text
VM          Direction              Health Recovery    RPO min Latest instance   Age min Result
--          ---------              ------ --------    ------- ---------------   ------- ------
srv-app-01  contoso-a -> contoso-b GREEN  NOT_STARTED      40 2026-10-08 14:09       12 OK
srv-db-01   contoso-a -> contoso-b YELLOW NOT_STARTED      40 2026-10-08 10:59      202 PROBLEM
srv-test-01 -                      -      -                                             MISSING

Destination           Resource                  Limit Used Free Need Left after DR Result
-----------           --------                  ----- ---- ---- ---- ------------- ------
tenant-b / vdc-b-main vCPU                         12    8    4   11            -7 NOT ENOUGH
tenant-b / vdc-b-main RAM GB                       64   40   24   38           -14 NOT ENOUGH
tenant-b / vdc-b-main Storage 'Fast' GB     unlimited   98       350               CHECK
tenant-b / vdc-b-main Storage 'Standard' GB       400  293  107  160           -53 CHECK

[ERROR] CAPACITY: NOT ENOUGH or incomplete data
[ERROR] REPLICATION: PROBLEMS - failover now would lose data or is impossible
```

## How the template update works

```
vCenter:          template ──► VM + customization spec ──► power on ──► guest steps ──► guest powers off ──► notes ──► template
Cloud Director:   catalog item ──► vApp (compose + recompose, static IP) ──► guest steps ──► powered off
                  ──► undeploy ──► "<name>" renamed to "<name>_Old" ──► capture vApp as "<name>" ──► delete vApp
```

Guest steps ([lib/Guest.ps1](lib/Guest.ps1)):

- **Windows**: reset the Windows Update cache, copy `GuestFiles`, run `GuestScripts` hooks, install updates in
  rounds (PSWindowsUpdate inside a one-time SYSTEM scheduled task, because the Windows Update API refuses remote
  sessions), reboot while a reboot is pending, DISM component cleanup, then sysprep `/generalize /oobe /shutdown`
  from a SYSTEM task.
- **Linux**: apt or dnf upgrade, reboot, `GuestFiles`, `GuestScripts` (bash as root), then sealing: cloud-init
  clean, empty machine-id, SSH host keys removed, logs and history cleared, poweroff scheduled with systemd-run.

Organisation-specific steps (agents, CA certificates, time servers, repositories) live in hook scripts; see
[templates/guest/windows-baseline.example.ps1](templates/guest/windows-baseline.example.ps1) and
[templates/guest/linux-baseline.example.sh](templates/guest/linux-baseline.example.sh).

Safety: nothing changes without `-Execute`; a template that is currently a VM, a missing customization spec or a
deploy IP that already answers stop the run before any change; on failure the vCenter object is converted back
to a template (or `Repair-VCenterTemplate.ps1` is suggested), and the Cloud Director catalog is only touched
after the guest has shut down cleanly.

## Notes

- **Name resolution.** With SRM the same VM name exists twice (the VM and its placeholder at the recovery site);
  snapshot and template tools skip placeholders, templates and diskless shells and refuse a name that still
  matches several real VMs. Operations then use the MoRef, never the name.
- **Replication accounting.** A replication is visible from both the source and the destination tenant; reports
  count it once by id. Failover demand per destination VDC is the sum of vCPU / RAM of the replicated VMs.
- **Storage after failover.** Replica disks already occupy the destination policy, so a storage shortfall for
  replicated VMs is reported as CHECK, not as a failure.
- **Tested** offline against a mock Cloud Director / vCAV API (reports, DR check, replication create/remove,
  vApp moves, template dry run). The PowerCLI and guest parts need a real vCenter and are not covered by it.

## Limitations

- Tenant-side Cloud Director API only (no provider / system administrator operations).
- Moving a VM into or out of the DR vApp requires the VM to be powered off; otherwise the move is skipped with a warning.
- `Copy-VCenterTemplate.ps1` needs vCenters that allow cross-vCenter cloning (same SSO domain); otherwise export
  and import the template as OVF.
- Windows template updates use WinRM over HTTP by IP: add the deploy IP to TrustedHosts or use an HTTPS listener.
- `Get-VcdVersion.ps1` infers the release from the highest API version; patch level is a lower bound.

## License

[MIT](LICENSE)
