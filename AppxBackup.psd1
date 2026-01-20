@{
    # Script module or binary module file associated with this manifest.
    RootModule = 'AppxBackup.psm1'

    # Version number of this module.
    ModuleVersion = '2.0.1'

    # Supported PSEditions
    CompatiblePSEditions = @('Desktop', 'Core')

    # ID used to uniquely identify this module
    GUID = 'a3f2c8d1-9b4e-4a7c-8f6d-2e1b9c4a7f3e'

    # Author of this module
    Author = 'DeltaGa'

    # Company or vendor of this module
    CompanyName = 'DeltaGa Systems Engineering'

    # Copyright statement for this module
    Copyright = '(c) DeltaGa. All rights reserved.'

    # Description of the functionality provided by this module
    Description = @'
Enterprise-grade Windows Application Package (APPX/MSIX) backup and restoration toolkit.

This module provides comprehensive functionality for:
- Backing up installed Windows Store/MSIX applications to portable packages
- Creating ZIP-based dependency packages (.appxpack) with orchestrated installation
- Managing certificates with modern security practices
- Handling dependencies, resources, and multi-architecture packages
- Validating package integrity and compatibility
- Automated rollback and error recovery

Replaces 2016-era tooling with modern PowerShell 7+ native capabilities
while maintaining backward compatibility with Windows 10/11 APPX infrastructure.

Key Features:
[CHECK] Comprehensive error handling and logging
[CHECK] Pipeline support for batch operations
[CHECK] Progress indication for long-running operations
[CHECK] Secure certificate management with HSM support
[CHECK] MSIX sparse package support
[CHECK] ZIP-based dependency packaging with installation orchestration
[CHECK] Extensive validation and compatibility checking
'@

    # Minimum version of the PowerShell engine required by this module
    PowerShellVersion = '5.1'

    # Minimum version of the common language runtime (CLR) required by this module
    # CLRVersion = '4.0'

    # Processor architecture (None, X86, Amd64) required by this module
    ProcessorArchitecture = 'None'

    # Modules that must be imported into the global environment prior to importing this module
    RequiredModules = @()

    # Assemblies that must be loaded prior to importing this module
    RequiredAssemblies = @()

    # Script files (.ps1) that are run in the caller's environment prior to importing this module.
    ScriptsToProcess = @()

    # Type files (.ps1xml) to be loaded when importing this module
    TypesToProcess = @()

    # Format files (.ps1xml) to be loaded when importing this module
    FormatsToProcess = @()

    # Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
    NestedModules = @()

    # Functions to export from this module, for best performance, do not use wildcards and do not delete the entry
    FunctionsToExport = @(
        'Backup-AppxPackage',
        'Install-AppxBackup',
        'New-AppxBackupCertificate',
        'Test-AppxPackageIntegrity',
        'Get-AppxBackupInfo',
        'Export-AppxDependencies',
        'Get-AppxToolPath',
        'Test-AppxBackupCompatibility'
    )

    # Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry
    CmdletsToExport = @()

    # Variables to export from this module
    VariablesToExport = @()

    # Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry
    AliasesToExport = @(
        'Backup-AppX',
        'Export-AppX',
        'Save-AppxPackage',
        'Restore-AppxPackage'
    )

    # DSC resources to export from this module
    DscResourcesToExport = @()

    # List of all modules packaged with this module
    ModuleList = @()

    # List of all files packaged with this module
    FileList = @(
        'AppxBackup.psd1',
        'AppxBackup.psm1',
        'Public\Backup-AppxPackage.ps1',
        'Public\Install-AppxBackup.ps1',
        'Public\New-AppxBackupCertificate.ps1',
        'Public\Test-AppxPackageIntegrity.ps1',
        'Public\Get-AppxBackupInfo.ps1',
        'Public\Export-AppxDependencies.ps1',
        'Public\Get-AppxToolPath.ps1',
        'Public\Test-AppxBackupCompatibility.ps1',
        'Private\Invoke-ProcessSafely.ps1',
        'Private\Write-AppxLog.ps1',
        'Private\Get-AppxManifestData.ps1',
        'Private\New-AppxPackageInternal.ps1',
        'Private\Test-AppxToolAvailability.ps1',
        'Private\Resolve-AppxDependencies.ps1',
        'Private\ConvertTo-SecureFilePath.ps1',
        'Private\Get-AppxConfiguration.ps1',
        'Private\Get-AppxDefault.ps1',
        'Config\MimeTypes.json',
        'Config\ToolConfiguration.json',
        'Config\WindowsReservedNames.json',
        'Config\PackageConfiguration.json',
        'Config\ModuleDefaults.json'
    )

    # Private data to pass to the module specified in RootModule/ModuleToProcess
    PrivateData = @{
        PSData = @{
            # Tags applied to this module for discoverability
            Tags = @(
                'Windows',
                'APPX',
                'MSIX',
                'WindowsStore',
                'Backup',
                'Package',
                'Certificate',
                'UWP',
                'Store',
                'Application'
            )

            # A URL to the license for this module.
            LicenseUri = ''

            # A URL to the main website for this project.
            ProjectUri = ''

            # A URL to an icon representing this module.
            IconUri = ''

            # ReleaseNotes of this module
            ReleaseNotes = @'
            Version 2.0.1 (January 20, 2026)
            ================================

            RELEASE SUMMARY:
            Complete rewrite of AppxBackup module from 2016 deprecated implementation to modern PowerShell 7+ native capabilities with comprehensive feature set.

            MAJOR FEATURES:
            - ZIP-based dependency packaging (.appxpack) with structured metadata and orchestrated installation
            - Manifest parsing with multi-tier fallback strategies for namespace resolution
            - External JSON configuration system (MimeTypes, ToolConfiguration, ModuleDefaults, etc.)
            - Automatic certificate installation to Trusted Root store with privilege escalation fallback
            - SDK tool validation with intelligent fallback and comprehensive error diagnostics
            - Signature validation for backup archives using SignTool integration
            - Dynamic tool path resolution returning actual paths instead of boolean values
            - Process safety framework with tool-specific timeouts and output buffer management

            COMPATIBILITY:
            - Windows 10 1809+ (all editions)
            - Windows 11 (all versions)
            - Windows Server 2019+
            - PowerShell 5.1+ (7.4+ recommended)
            - MSIX Packaging Tool compatible
            - Legacy APPX fully supported
            '@

            # Prerelease string of this module
            Prerelease = ''

            # Flag to indicate whether the module requires explicit user acceptance for install/update/save
            RequireLicenseAcceptance = $false

            # External dependent modules of this module
            ExternalModuleDependencies = @()
        }
    }

    # HelpInfo URI of this module
    HelpInfoURI = ''

    # Default prefix for commands exported from this module
    DefaultCommandPrefix = ''
}