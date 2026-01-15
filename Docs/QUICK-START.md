# AppxBackup - Quick Start Guide

## [CHAR_128640] How to Use This Module

### Step 1: Import the Module

**From PowerShell, navigate to the module directory and run:**

```powershell
# Navigate to module directory
cd C:\Path\To\AppxBackup.Module

# Import the module
.\Import-AppxBackup.ps1
```

**OR manually import:**

```powershell
Import-Module "C:\Path\To\AppxBackup.Module\AppxBackup.psd1" -Force
```

---

### Step 2: Verify Module Loaded

```powershell
# Check loaded commands
Get-Command -Module AppxBackup

# Should show 8 commands:
# - Backup-AppxPackage
# - New-AppxBackupCertificate
# - Get-AppxBackupInfo
# - Restore-AppxPackage
# - Export-AppxDependencies
# - Test-AppxBackupCompatibility
# - Test-AppxPackageIntegrity
# - Get-AppxToolPath
```

---

### Step 3: Use the Commands

#### Example 1: Backup an App

```powershell
# Find your app
$app = Get-AppxPackage -Name "*GameName*"

# Backup it
Backup-AppxPackage -PackagePath $app.InstallLocation -OutputPath "C:\GameBackups" -Verbose
```

#### Example 2: Backup with Full Path

```powershell
# If you know the exact path
Backup-AppxPackage -PackagePath "C:\Program Files\WindowsApps\Disney.Wreck-itRalph_1.0.0.12_x86__6rarf9sa4v8jt" -OutputPath "C:\GameBackups"
```

#### Example 3: Pipeline Usage

```powershell
# Backup all Adobe apps
Get-AppxPackage -Name "*Adobe*" | Backup-AppxPackage -OutputPath "C:\Backups"
```

---

## [WARNING] Common Mistakes

### [X] WRONG: Running script directly
```powershell
# This will NOT work:
.\Public\Backup-AppxPackage.ps1 -PackagePath ...
```

**Why:** Private functions not loaded, dependencies missing.

### [OK] CORRECT: Import module first
```powershell
# Step 1: Import
Import-Module .\AppxBackup.psd1

# Step 2: Use command
Backup-AppxPackage -PackagePath ...
```

---

## [CHAR_128295] Troubleshooting

### Error: "Command not recognized"

**Solution:**
```powershell
# Ensure module is imported
Get-Module AppxBackup

# If empty, import it
Import-Module .\AppxBackup.psd1 -Force
```

### Error: "Execution policy"

**Solution:**
```powershell
# Allow script execution (run as Administrator)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Error: "Module not found"

**Solution:**
```powershell
# Use full path
Import-Module "C:\Full\Path\To\AppxBackup.Module\AppxBackup.psd1" -Force
```

---

## [CHAR_128230] Installation (Permanent)

To make the module available system-wide:

```powershell
# Copy module to PowerShell modules directory
$modulePath = "$env:USERPROFILE\Documents\WindowsPowerShell\Modules\AppxBackup"
Copy-Item -Path "C:\Path\To\AppxBackup.Module" -Destination $modulePath -Recurse -Force

# Now you can import from anywhere
Import-Module AppxBackup
```

---

## [CHAR_128161] Complete Workflow Example

```powershell
# 1. Navigate to module
cd C:\Downloads\AppxBackup.Module

# 2. Import module
.\Import-AppxBackup.ps1

# 3. Find app to backup
$app = Get-AppxPackage -Name "*Wreck*"
$app.Name          # Verify it's the right app
$app.InstallLocation  # Check the path

# 4. Create backup directory
New-Item -Path "C:\GameBackups" -ItemType Directory -Force

# 5. Backup the app
Backup-AppxPackage -PackagePath $app.InstallLocation `
    -OutputPath "C:\GameBackups" `
    -Verbose

# 6. Verify backup created
Get-ChildItem "C:\GameBackups" -Filter "*.appx"
Get-ChildItem "C:\GameBackups" -Filter "*.cer"

# 7. Done! Files are ready for restoration
```

---

## [INFO] All Available Commands

```powershell
# Backup a package
Backup-AppxPackage -PackagePath <path> -OutputPath <path>

# Get package info
Get-AppxBackupInfo -PackagePath <path_to_appx>

# Restore a package
Restore-AppxPackage -PackagePath <path_to_appx> -CertificatePath <path_to_cer>

# Export dependencies
Export-AppxDependencies -PackagePath <path> -OutputPath <path> -Format HTML

# Test compatibility
Test-AppxBackupCompatibility -PackagePath <path_to_appx> -Detailed

# Test integrity
Test-AppxPackageIntegrity -PackagePath <path_to_appx> -VerifySignature

# Create certificate
New-AppxBackupCertificate -Subject "CN=MyCompany" -OutputPath <path>

# Get tool path
Get-AppxToolPath -ToolName MakeAppx
```

---

## [CHAR_127919] Key Points

1. [OK] **Always import the module first**
2. [OK] Use `Import-AppxBackup.ps1` helper script
3. [OK] Don't run individual .ps1 files directly
4. [OK] Module loads all dependencies automatically
5. [OK] Check logs in `$env:TEMP\AppxBackup_*.log`

---

**Need Help?**

```powershell
# Get help for any command
Get-Help Backup-AppxPackage -Full
Get-Help Backup-AppxPackage -Examples
```

---

**Module ready to use after import!** [CHAR_128640]
