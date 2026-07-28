@{
    # Script module or binary module file associated with this manifest.
    RootModule = 'AppxBackup.psm1'

    # Version number of this module.
    ModuleVersion = '2.1.0'

    # Supported PSEditions
    # Core only: the Desktop edition tops out at Windows PowerShell 5.1, which can
    # never satisfy the PowerShellVersion = '7.4' requirement below. Advertising
    # Desktop made the module look installable on Windows PowerShell when it is not.
    CompatiblePSEditions = @('Core')

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
    PowerShellVersion = '7.4'

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
        'Private\New-AppxBackupZipArchive.ps1',
        'Private\New-AppxBackupManifest.ps1',
        'Private\New-AppxDependencyCertificate.ps1',
        'Private\ConvertTo-HtmlEncodedString.ps1',
        'Private\Copy-AppxSourceDirectory.ps1',
        'Private\Find-AppxSdkTool.ps1',
        'Private\Get-AppxMakeAppxErrorAnalysis.ps1',
        'Private\Install-AppxCertificateToStore.ps1',
        'Private\Invoke-AppxSignTool.ps1',
        'Private\Remove-AppxItemWithRetry.ps1',
        'Private\Resolve-AppxManifestNode.ps1',
        'Private\Test-AppxArchitectureCompatibility.ps1',
        'Private\Test-AppxDiskSpace.ps1',
        'Private\Test-AppxPackagingPrerequisites.ps1',
        'Config\ToolConfiguration.json',
        'Config\WindowsReservedNames.json',
        'Config\PackageConfiguration.json',
        'Config\ModuleDefaults.json',
        'Config\ZipPackagingConfiguration.json',
        'Examples\UsageExamples.md',
        'Import-AppxBackup.ps1',
        'Test-AppxBackupModule.ps1',
        'build.ps1',
        'PSScriptAnalyzerSettings.psd1',
        'tests\Module.Tests.ps1',
        'tests\ConvertTo-SecureFilePath.Tests.ps1',
        'tests\Invoke-ProcessSafely.Tests.ps1',
        'tests\Test-AppxToolAvailability.Tests.ps1',
        'tests\Configuration.Tests.ps1'
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
                'Application',
                'Restore',
                'Migration',
                'WindowsPackage',
                'AppxPackage',
                'MsixPackage',
                'Dependency'
            )

            # A URL to the license for this module.
            LicenseUri = 'https://github.com/DeltaGa/AppxBackup.Module/blob/main/LICENSE'

            # A URL to the main website for this project.
            ProjectUri = 'https://github.com/DeltaGa/AppxBackup.Module'

            # A URL to an icon representing this module.
            IconUri = 'https://raw.githubusercontent.com/DeltaGa/AppxBackup.Module/main/Assets/icon.png'

            # ReleaseNotes of this module
            ReleaseNotes = @'
Version 2.1.0 (July 27, 2026)
=============================

RELEASE SUMMARY:
Correctness release. Six defects are fixed in which a guard compared a boolean
result against $null and therefore never ran, including one that could hang a
backup indefinitely and one that corrupted every captured tool output. The
module gains its first automated test suite, a static-analysis configuration
and a CI pipeline.

FIXED:

Process execution
- Invoke-ProcessSafely never enforced its timeout. WaitForExit([int]) returns
  $false on timeout but was compared against $null, so the kill branch was
  unreachable and control fell through to an unbounded WaitForExit(). A tool
  that stopped responding blocked the caller with no recourse; MakeAppx is
  invoked with a 1800 second timeout. Verified: a 3s timeout previously ran a
  30s process to completion and reported success, and now aborts at 3s.
- Invoke-ProcessSafely never detected a failed Process.Start().
- Captured stdout and stderr were prefixed with the capture buffer's own
  capacity ("16384" and "4096"). The buffers were built from an Int64 supplied
  by ConvertFrom-Json, and StringBuilder has (Int32) and (String) overloads but
  no Int64, so PowerShell bound the String one. This corrupted MakeAppx error
  analysis and the diagnostics printed on failure.

Path validation
- ConvertTo-SecureFilePath ignored -MustExist, so a missing path produced
  "The property 'PSIsContainer' cannot be found" from a later Get-Item instead
  of a clear error.
- ConvertTo-SecureFilePath ignored -CreateIfMissing, and creation was also
  gated behind -MustExist. Backup-AppxPackage passes -CreateIfMissing alone,
  so -OutputPath was never created.

Tool discovery
- Test-AppxToolAvailability returned a two-element array on every cache hit.
  The cache check used return inside begin{}, which exits only that block, so
  the cached value was emitted and the full search then emitted a second one.

Diagnostics
- Backup-AppxPackage reported "variable cannot be retrieved" in place of the
  real failure when an error occurred before the output path was computed.
- Write-AppxLog: removed a dead preference guard and corrected an empty-string
  check against $null.

CHANGED:

- Aliases are exported module members instead of being created with
  -Scope Global. Remove-Module now withdraws them rather than leaving them in
  the caller's session. The four alias names are unchanged.
- The loader discovers Public/ and Private/ from disk instead of using
  hand-maintained ordered lists. Every file there defines only functions, so
  load order carries no meaning. Adding a file no longer requires registering
  it in two places.
- CompatiblePSEditions is now Core only. Desktop tops out at Windows
  PowerShell 5.1 and could never satisfy PowerShellVersion 7.4.
- AppxBackup.psd1 and Write-AppxLog.ps1 contain non-ASCII characters and now
  carry a UTF-8 BOM so they are read correctly by ANSI-defaulting hosts.

ADDED:

- tests/: 71 Pester tests covering path validation, process execution, tool
  discovery, configuration loading and the manifest/loader/disk contract.
  Written in the Pester 5 and 6 compatible subset. Every defect above has a
  regression test.
- build.ps1: single Build/Analyze/Test entry point shared by contributors
  and CI, with NUnit and CSV output under -CI.
- PSScriptAnalyzerSettings.psd1: correctness-scoped rule set; each exclusion
  documents why the rule does not apply to this project.
- .github/workflows/ci.yml: Windows CI across PowerShell 7.4 and 7.5.
- .gitignore and CHANGELOG.md.

COMPATIBILITY:
- Public contract unchanged: the same 8 functions and 4 aliases.
- Windows 10 1809+, Windows 11, Windows Server 2019+
- PowerShell 7.4+

Version 2.0.2 (February 13, 2026)
=================================

RELEASE SUMMARY:
Code quality and maintainability improvements with 11 new private helper functions extracted from core logic.
Refactored Backup-AppxPackage and New-AppxPackageInternal removing ~600 lines of duplicate/embedded functionality.
Comprehensive test harness added for all 8 public functions with CI/CD integration support.

WHAT'S NEW:

Code Organization
- 11 new private helper functions (ConvertTo-HtmlEncodedString, Copy-AppxSourceDirectory, Find-AppxSdkTool, Get-AppxMakeAppxErrorAnalysis, Install-AppxCertificateToStore, Invoke-AppxSignTool, Remove-AppxItemWithRetry, Resolve-AppxManifestNode, Test-AppxArchitectureCompatibility, Test-AppxDiskSpace, Test-AppxPackagingPrerequisites)
- Refactored core functions by removing ~600 lines of duplicate logic
- Improved code testability and reusability

Directory Operations & Validation
- Three-tier fallback directory copying (Robocopy → Copy-Item → .NET APIs)
- Disk space validation with configurable thresholds
- System prerequisite checking (SDK tools, Windows version)
- CPU architecture compatibility validation

Certificate & Signing
- Certificate installation to Windows certificate stores
- Robust SignTool execution wrapper with error handling
- MakeAppx error message parsing and interpretation

Testing & Documentation
- Test-AppxBackupModule.ps1 comprehensive test harness for all 8 public functions
- Automatic app discovery and transcript logging
- Structured results with pass/fail/skip/error tracking for CI/CD
- Expanded UsageExamples.md with detailed examples for all functions
- Updated README with Private Functions reference section

COMPATIBILITY:
- Windows 10 1809+ (all editions)
- Windows 11 (all versions)
- Windows Server 2019+
- PowerShell 7.4+
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