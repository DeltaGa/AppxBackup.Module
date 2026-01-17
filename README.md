# AppxBackup PowerShell Module v2.0.0

## Windows Application Package Backup & Restoration Toolkit

**Version:** 2.0.0  
**Release Date:** January 15, 2026  
**PowerShell:** 5.1+ (7.5+ Recommended)

---

## 🎯 Overview

AppxBackup is a **complete 2026 rewrite** of the 2016 amateur APPX backup script, transformed into a production-grade PowerShell module. This toolkit provides **zero-dependency**, **end-to-end** Windows Store/MSIX application backup and restoration.

### Key Capabilities

✅ **Complete Backup-to-Installation Pipeline** - One-command backup, one-command restore  
✅ **Automatic Certificate Management** - Self-signed cert creation and system trust  
✅ **Native Windows SDK Integration** - MakeAppx and SignTool automation  
✅ **Intelligent Fallback Strategies** - Multi-tier copy and package creation  
✅ **Professional Error Diagnostics** - Actionable error messages with solutions  
✅ **Zero External Dependencies** - Pure PowerShell with optional SDK enhancement  
✅ **Production-Ready Logging** - Comprehensive structured logs with rotation  
✅ **Modern MSIX Support** - Full support for MSIX + legacy APPX formats

---

## 📊 2016 vs 2026: The Transformation

| Aspect | 2016 Version | 2026 Version |
|--------|---------------------|------------------------|
| **Lines of Code** | 145 lines | 3,259 lines (15 functions) |
| **Architecture** | Monolithic script | Modular with Public/Private separation |
| **Tool Dependencies** | Requires VS 2015 SDK (MakeCert.exe, Pvk2Pfx.exe) | Zero dependencies, native PowerShell |
| **Certificate Creation** | Deprecated MakeCert.exe | Native `New-SelfSignedCertificate` with 4096-bit RSA |
| **Package Signing** | Manual SignTool.exe invocation | Automated SignTool with proper error capture |
| **Error Handling** | String matching, no recovery | Try/catch with rollback + timeout protection |
| **Path Security** | Vulnerable to injection | Full sanitization, validation, `-LiteralPath` everywhere |
| **Manifest Parsing** | Ignores XML namespaces | Dynamic namespace detection, all schema versions |
| **Copy Operations** | Basic `Copy-Item` | Three-tier strategy: Robocopy → Copy-Item → .NET APIs |
| **Progress Indication** | None | Full 6-stage progress bars |
| **Logging** | None | Structured logs with levels, rotation, timestamps |
| **Dependency Resolution** | Ignored | Recursive dependency graph analysis |
| **Content_Types.xml** | Assumed present | Auto-generated with 30+ MIME types if missing |
| **Certificate Installation** | Manual | Automatic installation to Trusted Root |
| **Installation Script** | None | Standalone `Install-AppxBackup.ps1` |
| **MSIX Support** | No | Full MSIX + APPX support |
| **Pipeline Support** | None | `ValueFromPipeline` on all major functions |
| **Testing** | None | Unit testable, extensive validation |
| **Documentation** | 10 lines | 500+ lines with examples |

---

## 🚀 Installation

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

## 📖 Complete Workflow Example

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

## 🛠️ Available Commands

### Public Functions

| Function | Description | Status |
|----------|-------------|--------|
| **`Backup-AppxPackage`** | Complete backup with cert creation and signing | ✅ Production Ready |
| **`Install-AppxBackup`** | Standalone installation with auto-cert handling | ✅ Production Ready |
| **`New-AppxBackupCertificate`** | Create self-signed certificates (4096-bit RSA) | ✅ Production Ready |
| **`Restore-AppxPackage`** | Restore from backup with dependency resolution | ✅ Production Ready |
| **`Get-AppxBackupInfo`** | Analyze backup packages without installing | ✅ Production Ready |
| **`Export-AppxDependencies`** | Extract and document package dependencies | ✅ Production Ready |
| **`Test-AppxPackageIntegrity`** | Validate package structure and signatures | ✅ Production Ready |
| **`Test-AppxBackupCompatibility`** | Check system compatibility before restore | ✅ Production Ready |
| **`Get-AppxToolPath`** | Locate Windows SDK tools (MakeAppx, SignTool) | ✅ Production Ready |

### Private Functions (Internal)

- `Invoke-ProcessSafely` - Robust external process execution with timeout
- `Get-AppxManifestData` - XML manifest parsing with namespace handling
- `New-AppxPackageInternal` - Core packaging logic with multi-tier fallback
- `Test-AppxToolAvailability` - SDK tool validation and caching
- `Resolve-AppxDependencies` - Recursive dependency graph analysis
- `ConvertTo-SecureFilePath` - Path validation and sanitization
- `Write-AppxLog` - Structured logging with level filtering

---

## 📚 Usage Examples

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

### Batch Operations

#### 4. Backup All Apps from a Publisher
```powershell
Get-AppxPackage -Publisher "*Microsoft*" | 
    ForEach-Object {
        Backup-AppxPackage -PackagePath $_.InstallLocation `
            -OutputPath "D:\Backups\Microsoft" `
            -Verbose
    }
```

#### 5. Backup with Dependency Analysis
```powershell
$result = Backup-AppxPackage -PackagePath $app.InstallLocation `
    -OutputPath "C:\Backups" `
    -IncludeDependencies

# Export dependency report
$result.DependencyInfo | Export-Csv "dependencies.csv"
```

### Installation Operations

#### 6. Basic Installation
```powershell
Install-AppxBackup -PackagePath "C:\Backups\MyApp.appx"
```

#### 7. Install for Current User Only (No Admin)
```powershell
Install-AppxBackup -PackagePath "C:\Backups\MyApp.appx" `
    -CertStoreLocation CurrentUser
```

#### 8. Force Reinstall
```powershell
Install-AppxBackup -PackagePath "C:\Backups\MyApp.appx" -Force
```

#### 9. Manual Certificate Specification
```powershell
Install-AppxBackup -PackagePath "C:\Backups\MyApp.appx" `
    -CertificatePath "C:\Certs\MyCompanyCert.cer"
```

#### 10. Skip Certificate (Already Trusted)
```powershell
Install-AppxBackup -PackagePath "C:\Backups\MyApp.appx" -SkipCertificate
```

---

## 🐛 Troubleshooting

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
2. Module will fall back to .NET compression automatically

#### Error: Access Denied (WindowsApps)

**Cause:** Insufficient permissions  
**Solution:**
- Module automatically uses multi-tier copy fallback
- Robocopy → Copy-Item → .NET APIs
- No manual intervention needed

---

## 📋 System Requirements

### Minimum Requirements

- **PowerShell:** 5.1+
- **OS:** Windows 10 1809 / Windows Server 2019
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

## 📝 Module Architecture

```
AppxBackup.Module/
├── AppxBackup.psd1           # Module manifest
├── AppxBackup.psm1           # Module loader
├── Import-AppxBackup.ps1     # Quick import helper
│
├── Public/                   # Exported functions (9 total)
│   ├── Backup-AppxPackage.ps1          # Main backup function
│   ├── Install-AppxBackup.ps1          # Standalone installer
│   ├── New-AppxBackupCertificate.ps1   # Certificate creation
│   ├── Restore-AppxPackage.ps1         # Package restoration
│   ├── Get-AppxBackupInfo.ps1          # Package analysis
│   ├── Export-AppxDependencies.ps1     # Dependency export
│   ├── Test-AppxPackageIntegrity.ps1   # Integrity validation
│   ├── Test-AppxBackupCompatibility.ps1 # Compatibility check
│   └── Get-AppxToolPath.ps1            # Tool locator
│
├── Private/                  # Internal functions (7 total)
│   ├── Invoke-ProcessSafely.ps1        # Process execution
│   ├── Get-AppxManifestData.ps1        # Manifest parsing
│   ├── New-AppxPackageInternal.ps1     # Core packaging logic
│   ├── Test-AppxToolAvailability.ps1   # Tool validation
│   ├── Resolve-AppxDependencies.ps1    # Dependency resolution
│   ├── ConvertTo-SecureFilePath.ps1    # Path sanitization
│   └── Write-AppxLog.ps1               # Logging system
│
├── Docs/                     # Documentation
│   ├── QUICK-START.md
│   ├── SUMMARY.md
│   ├── 2016-vs-2026-Comparison.md
│   └── IMPLEMENTATION-STATUS.md
│
└── Examples/                 # Usage examples
    └── UsageExamples.md
```
---

## 📞 Support & Contributing

### Getting Help

- **Documentation:** See `/Docs` folder for detailed guides
- **Examples:** See `/Examples/UsageExamples.md`
- **Issues:** Check logs in `$env:TEMP\AppxBackup_*.log`

### Reporting Issues

When reporting issues, include:
1. PowerShell version (`$PSVersionTable`)
2. Windows version (`Get-ComputerInfo | Select OSName, OSVersion`)
3. Error message and stack trace
4. Log file from `$env:TEMP\AppxBackup_*.log`

---
© 2026 DeltaGa. All rights reserved.