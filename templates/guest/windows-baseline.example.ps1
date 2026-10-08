# Example guest hook for Windows templates (runs inside the guest as the template administrator,
# via Invoke-Command -FilePath, before the update rounds). Copy, adjust and list it in Templates.<name>.GuestScripts.
# Files it uses are copied by Templates.<name>.GuestFiles.

$ErrorActionPreference = 'Stop'

# Time source for clones that are not domain members yet.
$ntpPeers = 'ntp1.contoso.local ntp2.contoso.local'
w32tm /config /manualpeerlist:"$ntpPeers" /syncfromflags:manual /update | Out-Null
Restart-Service W32Time
Write-Output "NTP peers: $ntpPeers"

# Internal CA certificates (GuestFiles -> C:\Temp\certs).
foreach ($cert in Get-ChildItem C:\Temp\certs -Include *.cer, *.crt -Recurse -ErrorAction SilentlyContinue) {
    Import-Certificate -FilePath $cert.FullName -CertStoreLocation Cert:\LocalMachine\Root | Out-Null
    Write-Output "Imported $($cert.Name)"
}

# Agents shipped as MSI (GuestFiles -> C:\Temp\agents), installed quietly once.
foreach ($msi in Get-ChildItem C:\Temp\agents -Filter *.msi -ErrorAction SilentlyContinue) {
    $p = Start-Process msiexec.exe -ArgumentList "/i `"$($msi.FullName)`" /qn /norestart" -Wait -PassThru
    Write-Output "$($msi.Name): exit code $($p.ExitCode)"
}

Remove-Item C:\Temp\certs, C:\Temp\agents -Recurse -Force -ErrorAction SilentlyContinue
