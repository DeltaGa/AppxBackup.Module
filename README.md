<img src="https://raw.githubusercontent.com/DeltaGa/AppxBackup.Module/main/Assets/icon.png" alt="AppxBackup Icon" width="100" height="100">

# AppxBackup PowerShell Module v2.0.1

## Windows Application Package Backup & Restoration Toolkit

**Version:** 2.0.1  
**Release Date:** January 20, 2026  
**PowerShell:** 5.1+ (7.5+ Recommended)

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

- **PowerShell:** 5.1+
- **OS:** Windows 10 1809 / Windows Server 2019
- **Windows SDK:** 10.0.19041.0+ (includes MakeAppx, SignTool)
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

---

## Usage Examples

### Basic Operations

#### 1. Simple Backup
```powershell
$app = Get-AppxPackage -Name "Microsoft.WindowsCalculator"
Backup-AppxPackage -PackagePath $app.InstallLocation -OutputPath "C:\Backups"
```

#### 2. Backup with Custom Certificate
```powershell
Backup-AppxPackage -PackagePath $app.InstallLocation `
    -OutputPath "C:\Backups" `
    -CertificateSubject "CN=MyCompany Code Signing" `
    -CertificatePassword (ConvertTo-SecureString "MyPassword" -AsPlainText -Force)
```

#### 3. Skip Certificate (Use Existing)
```powershell
Backup-AppxPackage -PackagePath $app.InstallLocation `
    -OutputPath "C:\Backups" `
    -NoCertificate
```

### Dependency Operations

#### 4. Backup with Dependencies (ZIP Archive)
```powershell
# Creates .appxpack file with main package and all dependencies
Backup-AppxPackage -PackagePath $app.InstallLocation `
    -OutputPath "C:\Backups" `
    -IncludeDependencies
```

#### 5. Dependency Analysis Only (No Backup)
```powershell
# Creates PackageName_Dependencies.json without creating backup
Backup-AppxPackage -PackagePath $app.InstallLocation `
    -OutputPath "C:\Backups" `
    -DependencyReportOnly
```

### Batch Operations

#### 6. Backup All Apps from a Publisher
```powershell
Get-AppxPackage -Publisher "*Microsoft*" | 
    ForEach-Object {
        Backup-AppxPackage -PackagePath $_.InstallLocation `
            -OutputPath "D:\Backups\Microsoft" `
            -Verbose
    }
```

### Installation Operations

#### 7. Install Standalone Package
```powershell
Install-AppxBackup -PackagePath "C:\Backups\MyApp.appx"
```

#### 8. Install ZIP Archive with Dependencies
```powershell
Install-AppxBackup -PackagePath "C:\Backups\MyApp.appxpack"
```

#### 9. Install for Current User Only (No Admin)
```powershell
Install-AppxBackup -PackagePath "C:\Backups\MyApp.appx" `
    -CertStoreLocation CurrentUser
```

#### 10. Force Reinstall
```powershell
Install-AppxBackup -PackagePath "C:\Backups\MyApp.appx" -Force
```

#### 11. Skip Certificate (Already Trusted)
```powershell
Install-AppxBackup -PackagePath "C:\Backups\MyApp.appx" -SkipCertificate
```

#### 12. Allow Unsigned Packages (Developer Mode)
```powershell
Install-AppxBackup -PackagePath "C:\Backups\MyApp.appx" -AllowUnsigned
```

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
│   └── New-AppxDependencyCertificate.ps1 # Dependency certificate creation
│
└── Examples/                 # Usage examples
    └── UsageExamples.md
```
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

### Version 2.0.1 (January 20, 2026)

#### Dependency Packaging
- Replaced bundle system with ZIP archives (.appxpack) containing packages, certificates, and metadata
- Metadata-driven installation orchestration with ordered dependency resolution

#### Configuration System
- Introduced external JSON-based configuration system for module extensibility

#### Manifest Parsing
- Rewrote Get-AppxManifestData with multi-tier fallback strategies for namespace resolution
- Handles malformed and non-standard manifests without hard dependency on Microsoft schemas

#### Certificate Management
- Automatic installation to Trusted Root store immediately after creation
- Privilege escalation fallback with intelligent warnings
- Individual certificates for each dependency in ZIP archives

#### SDK Tool Validation
- Test-AppxToolAvailability now returns tool path string instead of boolean
- Mandatory validation in Backup-AppxPackage with installation diagnostics and PATH analysis

#### Process Execution
- Unified process safety with tool-specific timeouts
- Output buffer limit per stream to prevent memory exhaustion

#### Installation
- Extended Install-AppxBackup to support ZIP archives with nested package handling
- Improved progress reporting with separate extraction, validation, and installation stages

#### Future Changes (Planned)
  *Nothing explicitly planned for future versions.*

---

## Citations

### Original Repository

[**mjmeans/Appx-Backup**](https://github.com/mjmeans/Appx-Backup): PowerShell script to backup an installed Windows Store App to an installable Appx file. (2016). *GitHub*.

---

© 2026 DeltaGa. All rights reserved.
