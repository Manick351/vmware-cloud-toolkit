# vmware-cloud-toolkit configuration example.
# Copy to config\toolkit.psd1 (ignored by git) and adjust. No passwords here: every "Credential" value is the
# NAME of a credential resolved by Get-ToolkitCredential (SecretManagement secret, or
# config\credentials\<name>.xml created with:  Get-Credential | Export-Clixml config\credentials\<name>.xml).
@{
    Settings = @{
        LogPath              = 'C:\Ops\vmware-cloud-toolkit\logs'
        ReportPath           = 'C:\Ops\vmware-cloud-toolkit\out'
        # CredentialPath     = 'C:\Ops\vmware-cloud-toolkit\config\credentials'
        VcdApiVersion        = '36.0'          # Cloud Director 10.3+; lower it for older endpoints
        VcavApiVersion       = '4.5'
        SkipCertificateCheck = $false          # $true only for lab endpoints with self-signed certificates
        # PSWindowsUpdatePath = 'D:\Distr\Modules\PSWindowsUpdate'   # copied into Windows guests without PSGallery access
    }

    Mail = @{
        SmtpServer = 'smtp.contoso.local'
        Port       = 587
        UseSsl     = $true
        From       = 'vmware-reports@contoso.local'
        To         = @('it-ops@contoso.local')
        ErrorTo    = @('cloud-team@contoso.local')     # API errors of Get-VcdVmReport
        Credential = 'smtp-relay'                       # omit for an anonymous relay
    }

    # Cloud Director organizations (tenant side).
    Tenants = @(
        @{
            Name              = 'tenant-a'
            Org               = 'contoso-a'                       # appended to the user name as user@org
            VcdUrl            = 'https://vcd.example.com'
            VcavUrl           = 'https://vcav.example.com'        # Cloud Director Availability (optional)
            Credential        = 'vcd-tenant-a'
            ExcludeVdcPattern = 'service'                         # VDCs to ignore in reports
            Location          = 'SiteA'                           # for Get-VMInventory
            DrVdc             = 'vdc-a-main'                      # destination VDC when this tenant is a DR target
            TemplateDeploy    = @{                                # Update-VcdTemplate
                Templates      = @('W2K22STD', 'UBUNTU2404')
                Catalog        = 'Contoso templates'
                Vdc            = 'vdc-a-main'
                StorageProfile = 'Standard'
                Network        = 'net-a-templates'
                DeployIp       = '10.0.10.250'
            }
        }
        @{
            Name       = 'tenant-b'
            Org        = 'contoso-b'
            VcdUrl     = 'https://vcd2.example.com'
            VcavUrl    = 'https://vcav2.example.com'
            Credential = 'vcd-tenant-b'
            Location   = 'SiteB'
            DrVdc      = 'vdc-b-main'
        }
    )

    # vCenters (PowerCLI).
    VCenters = @(
        @{
            Name       = 'vc-01.contoso.local'
            Credential = 'vcenter-01'
            Location   = 'SiteA'
            Locations  = @(                                       # ESXi host pattern -> location (first match wins)
                @{ HostPattern = 'esx-b*'; Location = 'SiteB' }
                @{ HostPattern = 'esx-edge*'; Location = 'Branch-01' }
            )
            DeployIp   = '10.0.20.250'                            # static IP given by the customization spec
            Templates  = @{                                       # template name -> MoRef id (or VM template name)
                'W2K22STD'   = 'VirtualMachine-vm-1001'
                'UBUNTU2404' = 'VirtualMachine-vm-1002'
            }
        }
        @{
            Name       = 'vc-02.contoso.local'
            Credential = 'vcenter-02'
            Location   = 'SiteB'
            DeployIp   = '10.0.30.250'
            Templates  = @{ 'W2K22STD' = 'W2K22STD' }
        }
    )

    # Template definitions shared by Update-VcenterTemplate and Update-VcdTemplate.
    Templates = @{
        'W2K22STD' = @{
            OS                = 'Windows'
            ComputerName      = 'W2K22STD'
            CustomizationSpec = 'TemplateUpdate-Windows'          # vCenter only; must set the static DeployIp
            GuestCredential   = 'template-windows'                # omit for vCD: auto-generated admin password is used
            SkipUpdateTitle   = '^VMware'                         # VMware Tools come with the platform
            MaxUpdateRounds   = 3
            GuestFiles        = @(
                @{ Source = 'D:\Distr\certs'; Destination = 'C:\Temp' }
                @{ Source = 'D:\Distr\agents'; Destination = 'C:\Temp' }
            )
            GuestScripts      = @('templates\guest\windows-baseline.example.ps1')
        }
        'UBUNTU2404' = @{
            OS                = 'Linux'
            CustomizationSpec = 'TemplateUpdate-Linux'
            GuestCredential   = 'template-linux'                  # root, or a user with NOPASSWD sudo
            GuestFiles        = @(@{ Source = 'D:\Distr\certs'; Destination = '/root' })
            GuestScripts      = @('templates\guest\linux-baseline.example.sh')
        }
    }

    # Cloud Director Availability: New-VcavReplication / Remove-VcavReplication.
    Replication = @{
        Pairs = @(
            @{ SourceTenant = 'tenant-a'; DestinationTenant = 'tenant-b'; SourceSite = 'SiteA'; DestinationSite = 'SiteB'; DestinationVdc = 'vdc-b-main'; StorageProfile = 'Standard' }
            @{ SourceTenant = 'tenant-b'; DestinationTenant = 'tenant-a'; SourceSite = 'SiteB'; DestinationSite = 'SiteA'; DestinationVdc = 'vdc-a-main'; StorageProfile = 'Standard' }
        )
        Settings = @{
            Rpo                = 40               # minutes
            DataConnectionType = 'ENCRYPTED'
            Quiesced           = $false
            RetentionInstances = 8                # keep 8 point-in-time instances ...
            RetentionDistance  = 120              # ... 120 minutes apart
            InitialSyncTime    = 0                # 0 = start now (epoch ms otherwise)
            IsMigration        = $false
        }
        DrVAppSuffix = '-DR'
    }

    # Test-DRCapacity.
    DR = @{
        ReservePercent = 5
        TenantPairs    = @{ 'tenant-a' = 'tenant-b'; 'tenant-b' = 'tenant-a' }   # used when a VM has no replication
    }
}
