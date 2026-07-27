<img src="https://raw.githubusercontent.com/DeltaGa/AppxBackup.Module/main/Assets/social_preview_no_star_count.jpg" alt="AppxBackup Social Preview">

[![CodeFactor](https://www.codefactor.io/repository/github/deltaga/appxbackup.module/badge)](https://www.codefactor.io/repository/github/deltaga/appxbackup.module)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

# AppxBackup PowerShell Module v2.0.2

## Windows Application Package Backup & Restoration Toolkit

**Version:** 2.0.2  
**Release Date:** February 13, 2026  
**PowerShell:** 7.4+

---

## Table of Contents

- [Overview](#overview)
- [System Requirements](#system-requirements)
- [Installation](#installation)
- [Workflow Example](#workflow-example)
- [Available Commands](#available-commands)
- [Usage Examples](#usage-examples)
- [Troubleshooting](#troubleshooting)
- [Limitations](#limitations)
- [Module Architecture](#module-architecture)
- [Testing Infrastructure](#testing-infrastructure)
- [Support](#support)
- [Changelog](#changelog)
- [Citations](#citations)

---

## Overview

AppxBackup is a **complete 2026 rewrite** of the 2016 original APPX backup script, transformed into a production-grade PowerShell module. This toolkit provides **end-to-end** Windows Store/MSIX application backup and restoration.

### Key Capabilities

 **Complete Backup-to-Installation Pipeline** - One-command backup, one-command restore  
 **Automatic Certificate Management** - Self-signed cert creation and system trust  
 **Native Windows SDK Integration** - MakeAppx and SignTool automation with tool-specific optimization  
 **Modern MSIX Support** - Full support for MSIX + legacy APPX formats  
 **Externalized Configuration** - Configuration values in JSON for easy customization

---

## System Requirements

### Minimum Requirements

- **PowerShell:** 7.4
- **OS:** Windows 10 1809 / Windows Server 2019
- **Windows SDK:** 10.0.18362.0 (includes MakeAppx, SignTool)
- **Disk Space:** 100 MB temporary storage
- **Memory:** 512 MB available RAM

### Recommended Configuration

- **PowerShell:** 7.5+
- **OS:** Windows 11 24H2 / Windows Server 2022
- **Windows SDK:** 10.0.26100.0+ (includes MakeAppx, SignTool)
- **Disk Space:** 1 GB+ for large packages
- **Memory:** 2 GB+ available RAM

### Administrator Privileges

Required for:
- Installing certificates to LocalMachine store
- Accessing WindowsApps folder on some systems

Not required for:
- Basic backup operations (CurrentUser cert store)
- Package creation
- Most module functions

---

## Installation

### Quick Start (Developer Mode)

```powershell
# 1. Clone or extract the module
cd "C:\Path\To\AppxBackup.Module"

# 2. Import the module
.\Import-AppxBackup.ps1

# 3. Verify installation
Get-Command -Module AppxBackup
```

### Production Installation

```powershell
# Install to PowerShell Modules directory
$modulePath = "$env:USERPROFILE\Documents\PowerShell\Modules\AppxBackup"
Copy-Item -Path ".\AppxBackup.Module" -Destination $modulePath -Recurse -Force

# Import globally
Import-Module AppxBackup -Force

# Verify
Get-Module AppxBackup
```

---

## Workflow Example

For more detailed workflow examples, see [Common Workflows](https://github.com/DeltaGa/AppxBackup.Module/blob/main/Examples/UsageExamples.md#common-workflows) in [Examples/UsageExamples.md](https://github.com/DeltaGa/AppxBackup.Module/blob/main/Examples/UsageExamples.md).

### Scenario: Backup and Restore WorkMate App

#### Step 1: Backup the Application

```powershell
# Find the installed app
$app = Get-AppxPackage -Name "*WorkMate*"

# Create complete backup (package + certificate)
Backup-AppxPackage -PackagePath $app.InstallLocation -OutputPath "D:\Backups"
```

**Output:**
```
=== Backup-AppxPackage v2.0 ===
Package: 2242AppsTeam.WorkMate v7.5.4.0

[Stage 1/6] Validating inputs...                         ✓
[Stage 2/6] Creating package...                          ✓ (6.4s, 47.95 MB)
[Stage 3/6] Creating certificate...                      ✓
[Stage 4/6] Installing certificate to Trusted Root...    ✓
[Stage 5/6] Signing package...                           ✓

✓ Certificate installed successfully - package is ready to install

To install the package, run:
  Install-AppxBackup -PackagePath 'D:\Backups\WorkMate_7.5.4.0_x64.appx'
```

**Files Created:**
```
D:\Backups\
├── 2242AppsTeam.WorkMate_7.5.4.0_x64__v313ts49xh8we.appx  (47.95 MB, signed)
└── 2242AppsTeam.WorkMate_7.5.4.0_x64__v313ts49xh8we.cer  (certificate)
```

#### Step 2: Install on Another Machine

```powershell
# Single command - auto-detects certificate
Install-AppxBackup -PackagePath "D:\Backups\WorkMate_7.5.4.0_x64.appx"
```

**Output:**
```
=== APPX Package Installation Script ===

[1/3] Package Analysis
  Package: WorkMate_7.5.4.0_x64.appx
  Certificate: Auto-detected (WorkMate_7.5.4.0_x64.cer)

[2/3] Certificate Installation
  Target Store: Cert:\LocalMachine\Root
  Status: Installed successfully
  Thumbprint: 439F89C615A7468BF196D8951BBA4640C10D7D3F

[3/3] Package Installation
  Installing package...
  Status: Installed successfully
  Name: 2242AppsTeam.WorkMate
  Version: 7.5.4.0

=== Installation Complete ===
```

---

## Available Commands

### Public Functions

| Function | Description |
|----------|-------------|
| **`Backup-AppxPackage`** | Complete backup with cert creation and signing |
| **`Install-AppxBackup`** | Installation with auto-cert handling and dependency orchestration |
| **`New-AppxBackupCertificate`** | Create self-signed certificates (4096-bit RSA) |
| **`Get-AppxBackupInfo`** | Analyze backup packages without installing |
| **`Export-AppxDependencies`** | Extract and document package dependencies |
| **`Test-AppxPackageIntegrity`** | Validate package structure and signatures |
| **`Test-AppxBackupCompatibility`** | Check system compatibility before restore |
| **`Get-AppxToolPath`** | Locate Windows SDK tools (MakeAppx, SignTool) |

### Private Functions (Internal)

- `Get-AppxConfiguration` - Configuration loader from JSON files
- `Get-AppxDefault` - Configuration value accessor with hierarchical fallback
- `Invoke-ProcessSafely` - Robust external process execution with timeout
- `Get-AppxManifestData` - XML manifest parsing with namespace handling
- `New-AppxPackageInternal` - Core packaging logic with multi-tier fallback
- `Test-AppxToolAvailability` - SDK tool validation and caching
- `Resolve-AppxDependencies` - Recursive dependency graph analysis
- `ConvertTo-SecureFilePath` - Path validation and sanitization
- `Write-AppxLog` - Structured logging with level filtering
- `New-AppxBackupZipArchive` - ZIP archive creation for dependency packages
- `New-AppxBackupManifest` - Installation manifest generation
- `New-AppxDependencyCertificate` - Dependency-specific certificate creation
- `ConvertTo-HtmlEncodedString` - HTML entity encoding for safe output
- `Copy-AppxSourceDirectory` - Robust directory copying with multi-tier fallback
- `Find-AppxSdkTool` - Locate SDK tools in Program Files and registry
- `Get-AppxMakeAppxErrorAnalysis` - Parse and interpret MakeAppx error messages
- `Install-AppxCertificateToStore` - Install certificates to certificate stores
- `Invoke-AppxSignTool` - SignTool wrapper with error handling
- `Remove-AppxItemWithRetry` - Retry-based item removal with escalation
- `Resolve-AppxManifestNode` - Resolve XML nodes with namespace fallback
- `Test-AppxArchitectureCompatibility` - Check CPU architecture compatibility
- `Test-AppxDiskSpace` - Validate available disk space
- `Test-AppxPackagingPrerequisites` - Verify SDK tools and system requirements

---

## Usage Examples

For the backup command: see [Backup-AppxPackage](https://github.com/DeltaGa/AppxBackup.Module/blob/main/Examples/UsageExamples.md#2-backup-appxpackage) in [Examples/UsageExamples.md](https://github.com/DeltaGa/AppxBackup.Module/blob/main/Examples/UsageExamples.md).

For the install command: see [Install-AppxPackage](https://github.com/DeltaGa/AppxBackup.Module/blob/main/Examples/UsageExamples.md#8-install-appxbackup) in [Examples/UsageExamples.md](https://github.com/DeltaGa/AppxBackup.Module/blob/main/Examples/UsageExamples.md).

For more detailed usage examples, see the full [Examples/UsageExamples.md](https://github.com/DeltaGa/AppxBackup.Module/blob/main/Examples/UsageExamples.md).

---

## Troubleshooting

### Common Errors and Solutions

#### Error: 0x800B0109 (Certificate Not Trusted)

**Cause:** Certificate not in Trusted Root store  
**Solution:**
```powershell
# Run as Administrator
Import-Certificate -FilePath "path\to\cert.cer" `
    -CertStoreLocation "Cert:\LocalMachine\Root"
```

#### Error: 0x80073CF9 (Already Installed)

**Cause:** Package already installed  
**Solution:**
```powershell
Install-AppxBackup -PackagePath "package.appx" -Force
```

#### Error: MakeAppx Not Found

**Cause:** Windows SDK not installed  
**Solution:**
1. Install Windows SDK from Microsoft

#### Error: Access Denied (WindowsApps)

**Cause:** Insufficient permissions  
**Solution:**
- Module automatically uses multi-tier copy fallback
- Robocopy → Copy-Item → .NET APIs
- No manual intervention needed

---

## Limitations

### Restoration Behavior

Some applications cannot be successfully restored after being repackaged. If a repackaged application fails to install, the application cannot be used from the backup. Reinstall from the Microsoft Store.

---

## Module Architecture

```
AppxBackup.Module/
├── AppxBackup.psd1           # Module manifest
├── AppxBackup.psm1           # Module loader
├── Test-AppxBackupModule.ps1 # Module test suite
├── Import-AppxBackup.ps1     # Quick import helper
│
├── Config/                   # Externalized configuration
│   ├── ModuleDefaults.json           # Core module constants
│   ├── ToolConfiguration.json        # Tool-specific settings
│   ├── WindowsReservedNames.json     # Reserved filenames
│   ├── PackageConfiguration.json     # Package-related constants
│   └── ZipPackagingConfiguration.json # ZIP archive settings
│
├── Public/                   # Exported functions
│   ├── Backup-AppxPackage.ps1          # Main backup function
│   ├── Install-AppxBackup.ps1          # Package installer (Restore-AppxPackage alias)
│   ├── New-AppxBackupCertificate.ps1   # Certificate creation
│   ├── Get-AppxBackupInfo.ps1          # Package analysis
│   ├── Export-AppxDependencies.ps1     # Dependency export
│   ├── Test-AppxPackageIntegrity.ps1   # Integrity validation
│   ├── Test-AppxBackupCompatibility.ps1 # Compatibility check
│   └── Get-AppxToolPath.ps1            # Tool locator
│
├── Private/                  # Internal functions
│   ├── Get-AppxConfiguration.ps1       # Configuration loader
│   ├── Get-AppxDefault.ps1             # Configuration value accessor
│   ├── Invoke-ProcessSafely.ps1        # Process execution
│   ├── Get-AppxManifestData.ps1        # Manifest parsing
│   ├── New-AppxPackageInternal.ps1     # Core packaging logic
│   ├── Test-AppxToolAvailability.ps1   # Tool validation
│   ├── Resolve-AppxDependencies.ps1    # Dependency resolution
│   ├── ConvertTo-SecureFilePath.ps1    # Path sanitization
│   ├── Write-AppxLog.ps1               # Logging system
│   ├── New-AppxBackupZipArchive.ps1    # ZIP archive creation
│   ├── New-AppxBackupManifest.ps1      # Installation manifest generation
│   ├── New-AppxDependencyCertificate.ps1 # Dependency certificate creation
│   ├── ConvertTo-HtmlEncodedString.ps1 # HTML encoding
│   ├── Copy-AppxSourceDirectory.ps1    # Directory copying
│   ├── Find-AppxSdkTool.ps1            # SDK tool locator
│   ├── Get-AppxMakeAppxErrorAnalysis.ps1 # MakeAppx error parser
│   ├── Install-AppxCertificateToStore.ps1 # Certificate installation
│   ├── Invoke-AppxSignTool.ps1         # SignTool wrapper
│   ├── Remove-AppxItemWithRetry.ps1    # Retry-based removal
│   ├── Resolve-AppxManifestNode.ps1    # XML node resolution
│   ├── Test-AppxArchitectureCompatibility.ps1 # Architecture check
│   ├── Test-AppxDiskSpace.ps1          # Disk space validation
│   └── Test-AppxPackagingPrerequisites.ps1 # Prerequisites check
│
└── Examples/                 # Usage examples
    └── UsageExamples.md
```

---

## Testing Infrastructure

### Test-AppxBackupModule.ps1

A comprehensive test harness that validates all 8 exported public functions with multiple parameter combinations.

**Automatic app discovery** via Get-AppxPackage pattern matching eliminates manual configuration.
Full console transcript and structured results file saved to timestamped test run directories.

**Usage:**
```powershell
# Test with custom app and output folder
.\Test-AppxBackupModule.ps1 -TestFolder "C:\Temp\AppxTest\" -AppName "Netflix"

# Include destructive installation test (optional)
.\Test-AppxBackupModule.ps1 -TestFolder "C:\Temp\" -AppName "Spotify" -TestInstall
```

Results include pass/fail/skip/error counts, detailed timing, and structured output suitable for CI/CD integration.

---

### Configuration System

All hardcoded values are externalized to JSON configuration files in the `Config/` directory:

- **ModuleDefaults.json** - Path limits, timeouts, buffer sizes, disk space thresholds, etc.
- **ToolConfiguration.json** - Tool-specific timeouts, async wait times, exit code interpretation
- **WindowsReservedNames.json** - Windows reserved filenames for validation
- **PackageConfiguration.json** - Package extensions, signature files, compression levels, namespaces
- **ZipPackagingConfiguration.json** - ZIP archive structure, compression settings, manifest defaults, system requirements

---

## Support

### Getting Help

- **Documentation:** Enter `help Command` (e. g. `help Backup-AppxPackage`)
- **Examples:** See `/Examples/UsageExamples.md`
- **Issues:** Check logs in `$env:TEMP\AppxBackup_*.log`

### Reporting Issues

When reporting issues, include:
1. PowerShell version (`$PSVersionTable`)
2. Windows version (`Get-ComputerInfo | Select OSName, OSVersion`)
3. Error message and stack trace
4. Log file from `$env:TEMP\AppxBackup_*.log`

---

## Changelog

### Version 2.0.2 (February 13, 2026)

#### Code Quality & Maintainability
- Extracted 11 new private helper functions to improve code organization and testability.
- Refactored core functions (Backup-AppxPackage, New-AppxPackageInternal) by removing ~600 lines of duplicate/embedded logic.

#### Testing Infrastructure
- New `Test-AppxBackupModule.ps1` comprehensive test harness validates all 8 public functions with multiple parameter combinations.
- Automatic app discovery, full transcript logging, and structured results with pass/fail/skip/error tracking for CI/CD integration.

#### Usage Examples & Documentation
- Expanded `Examples/UsageExamples.md` with detailed examples for all public functions covering parameter usage and real-world scenarios.
- Updated README with new Private Functions reference section documenting 11 helper functions and their purposes.

#### Disk Space & Package Validation
- New `Test-AppxDiskSpace()` function validates disk space and warns about large packages (>1GB or >10k files by default).
- Configurable size thresholds in PackageConfiguration for environment-specific requirements.

#### Directory Copying & File Operations
- New `Copy-AppxSourceDirectory()` uses three-tier fallback strategy: Robocopy → Copy-Item → .NET APIs.
- Handles WindowsApps folder permission issues transparently without user intervention.

#### System Requirements Validation
- New `Test-AppxPackagingPrerequisites()` validates SDK tools and Windows version before packaging.
- New `Test-AppxArchitectureCompatibility()` checks CPU architecture compatibility for installation.

#### Certificate & Signing Tools
- New `Install-AppxCertificateToStore()` manages certificate installation to Windows certificate stores.
- New `Invoke-AppxSignTool()` provides robust SignTool execution with error handling.

#### Diagnostic Improvements
- New `Get-AppxMakeAppxErrorAnalysis()` parses MakeAppx error messages for better troubleshooting.
- New `Find-AppxSdkTool()` locates SDK tools in Program Files and registry paths.

#### XML & Output Processing
- New `Resolve-AppxManifestNode()` resolves XML nodes with namespace fallback strategies.
- New `ConvertTo-HtmlEncodedString()` prevents XSS attacks in HTML dependency reports.

#### Installation & Cleanup
- New `Remove-AppxItemWithRetry()` provides retry-based item removal with escalation for locked files.
- Improved dependency resolution with metadata-driven installation orchestration.

#### Future Changes (Planned)
  *Nothing explicitly planned for future versions.*

---

## Citations

### Author

**Tchicdje Kouojip Joram Smith (DeltaGa)**  
Email: dev.github.tkjoramsmith@outlook.com  
GitHub: [https://github.com/DeltaGa](https://github.com/DeltaGa)

### Original Repository

[**mjmeans/Appx-Backup**](https://github.com/mjmeans/Appx-Backup): PowerShell script to backup an installed Windows Store App to an installable Appx file. (2016). *GitHub*.

---

© 2026 DeltaGa. All rights reserved.
