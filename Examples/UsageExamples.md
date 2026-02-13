# AppxBackup Module - Usage Examples V2

**Complete console reference for the AppxBackup v2.0.1 module.**

All examples assume you've imported the module:
```powershell
cd "C:\Path\AppxBackup.Module"
.\Import-AppxBackup.ps1
```

And have identified your target app:
```powershell
$app = Get-AppxPackage -Name "*AppName*"
```

---

## 1. Get-AppxToolPath

**Purpose:** Locate Windows SDK tools needed for packaging operations.

**Basic usage:**
```powershell
Get-AppxToolPath -ToolName MakeAppx
```

**What it returns:** Full path to the tool executable, or `$null` if not found.

**Use case:** Verify that Windows SDK is installed and accessible before running backups.
```powershell
if (Get-AppxToolPath -ToolName MakeAppx) {
    Write-Host "MakeAppx is available"
} else {
    Write-Host "Install Windows SDK to use this module"
}
```

**With -Refresh:** Force the module to rediscover tools (clears cache).
```powershell
Get-AppxToolPath -ToolName SignTool -Refresh
```

---

## 2. Backup-AppxPackage

**Purpose:** Create a backup of an installed application with certificate and signing.

### Basic Backup (Single File)

```powershell
Backup-AppxPackage -PackagePath $app.InstallLocation -OutputPath "C:\Backups"
```

Creates:
- `AppName_Version_Arch__PublisherID.appx` - The signed package
- `AppName_Version_Arch__PublisherID.cer` - The certificate for installation

### Backup with Dependencies (ZIP Archive)

```powershell
Backup-AppxPackage -PackagePath $app.InstallLocation -OutputPath "C:\Backups" -IncludeDependencies
```

Creates:
- `AppName_Version_Arch__PublisherID.appxpack` - ZIP file containing:
  - Main package (.appx)
  - All dependency packages (.appx files)
  - Individual certificates for each package
  - `AppxBackupManifest.json` - Installation orchestration metadata

**Use this when:** Migrating to another system or ensuring all dependencies are preserved.

### Dependency Analysis Only (No Backup)

```powershell
Backup-AppxPackage -PackagePath $app.InstallLocation -OutputPath "C:\Reports" -DependencyReportOnly
```

Creates: `AppName_Dependencies.json` with dependency information only.

**Use this when:** You just need to document what dependencies an app has.

### Backup without Certificate (For Manual Signing)

```powershell
Backup-AppxPackage -PackagePath $app.InstallLocation -OutputPath "C:\Backups" -NoCertificate
```

**Use this when:** You'll sign the package with an existing certificate separately.

### Overwrite Existing Files

```powershell
Backup-AppxPackage -PackagePath $app.InstallLocation -OutputPath "C:\Backups" -Force
```

Without `-Force`, the command will fail if output files already exist.

### With Custom Compression

```powershell
Backup-AppxPackage -PackagePath $app.InstallLocation -OutputPath "C:\Backups" -CompressionLevel Maximum
```

Valid values: `None`, `Fast`, `Normal`, `Maximum`. Default is `Normal`.

### Pipeline Example

```powershell
Get-AppxPackage -Name "MyApp" | Backup-AppxPackage -OutputPath "C:\Backups" -IncludeDependencies
```

---

## 3. Get-AppxBackupInfo

**Purpose:** Examine a backed-up package without installing it.

### Basic Info

```powershell
Get-AppxBackupInfo -PackagePath "C:\Backups\MyApp.appx"
```

Returns:
- PackageName
- PackageVersion
- PackageArchitecture (x86, x64, ARM, ARM64)
- PackageSizeMB
- Publisher

### With File List

```powershell
Get-AppxBackupInfo -PackagePath "C:\Backups\MyApp.appx" -IncludeFileList
```

Adds `.FileList` property - array of all files in the package.

### With Signature Info

```powershell
Get-AppxBackupInfo -PackagePath "C:\Backups\MyApp.appx" -IncludeSignatureInfo
```

Adds `.SignatureInfo` property with certificate details.

### With Raw Manifest

```powershell
Get-AppxBackupInfo -PackagePath "C:\Backups\MyApp.appx" -IncludeManifestXml
```

Adds `.ManifestXml` property - the raw AppxManifest.xml as a string.

### All Info Combined

```powershell
Get-AppxBackupInfo -PackagePath "C:\Backups\MyApp.appx" -IncludeFileList -IncludeSignatureInfo -IncludeManifestXml
```

### For ZIP Archives (.appxpack)

```powershell
Get-AppxBackupInfo -PackagePath "C:\Backups\MyApp.appxpack"
```

Returns:
- Same as above, plus
- IsZipArchive: $true
- ContainedPackages: Array of all packages in archive
- ContainedCertificates: Array of all certificates
- TotalArchiveSize: Combined size

---

## 4. Test-AppxPackageIntegrity

**Purpose:** Verify that a package is not corrupted and is properly signed.

### Basic Check

```powershell
Test-AppxPackageIntegrity -PackagePath "C:\Backups\MyApp.appx"
```

Returns `.IsValid` - Boolean indicating if archive is intact.

### Check Signature

```powershell
Test-AppxPackageIntegrity -PackagePath "C:\Backups\MyApp.appx" -VerifySignature
```

Returns:
- `.IsValid` - Archive is intact
- `.SignatureValid` - Certificate signature is valid

### Check Manifest

```powershell
Test-AppxPackageIntegrity -PackagePath "C:\Backups\MyApp.appx" -CheckManifest
```

Returns:
- `.IsValid` - Archive is intact
- `.ManifestValid` - Manifest XML is well-formed

### Full Validation

```powershell
Test-AppxPackageIntegrity -PackagePath "C:\Backups\MyApp.appx" -VerifySignature -CheckManifest
```

### For ZIP Archives

```powershell
Test-AppxPackageIntegrity -PackagePath "C:\Backups\MyApp.appxpack" -VerifySignature -CheckManifest
```

Additionally validates:
- Archive structure (Packages/, Certificates/, manifest)
- AppxBackupManifest.json schema
- All contained package signatures

---

## 5. Test-AppxBackupCompatibility

**Purpose:** Check if a package can be installed on the current system.

### Basic Check

```powershell
Test-AppxBackupCompatibility -PackagePath "C:\Backups\MyApp.appx"
```

Returns:
- `.IsCompatible` - Can this package be installed?
- `.ArchitectureCompatible` - Does CPU architecture match?
- `.SystemArchitecture` - System's architecture (x64, ARM64, etc.)
- `.PackageArchitecture` - Package's required architecture

### With Dependency Check

```powershell
Test-AppxBackupCompatibility -PackagePath "C:\Backups\MyApp.appx" -CheckDependencies
```

Additionally checks if all required dependencies are installed on the system.

### With Detailed Report

```powershell
Test-AppxBackupCompatibility -PackagePath "C:\Backups\MyApp.appx" -Detailed
```

Returns detailed information and recommendations for incompatibilities.

### Full Analysis

```powershell
Test-AppxBackupCompatibility -PackagePath "C:\Backups\MyApp.appx" -CheckDependencies -Detailed
```

### For ZIP Archives

```powershell
Test-AppxBackupCompatibility -PackagePath "C:\Backups\MyApp.appxpack" -CheckDependencies -Detailed
```

Additionally analyzes all contained packages and dependencies.

---

## 6. Export-AppxDependencies

**Purpose:** Document all dependencies for an application in various formats. Use -IncludeOptional to ensure that dependencies are listed.

### Export to JSON

```powershell
Export-AppxDependencies -PackagePath $app.InstallLocation -OutputPath "C:\Reports\deps.json" -Format JSON
```

### Export to HTML

```powershell
Export-AppxDependencies -PackagePath $app.InstallLocation -OutputPath "C:\Reports\deps.html" -Format HTML
```

Best for viewing in a browser. Creates a formatted report table.

### Export to XML

```powershell
Export-AppxDependencies -PackagePath $app.InstallLocation -OutputPath "C:\Reports\deps.xml" -Format XML
```

### Export to CSV

```powershell
Export-AppxDependencies -PackagePath $app.InstallLocation -OutputPath "C:\Reports\deps.csv" -Format CSV
```

Best for importing into Excel or analysis tools.

### Auto-detect Format from Filename

```powershell
Export-AppxDependencies -PackagePath $app.InstallLocation -OutputPath "C:\Reports\deps.json"
```

Format is automatically detected from the `.json` extension (no `-Format` needed).

### Recursive Analysis

```powershell
Export-AppxDependencies -PackagePath $app.InstallLocation -OutputPath "C:\Reports\deps.json" -Recursive
```

Analyzes dependencies of dependencies (nested).

### Include Optional Dependencies

```powershell
Export-AppxDependencies -PackagePath $app.InstallLocation -OutputPath "C:\Reports\deps.json" -IncludeOptional
```

Includes optional framework packages.

### With Depth Limit

```powershell
Export-AppxDependencies -PackagePath $app.InstallLocation -OutputPath "C:\Reports\deps.json" -Recursive -MaxDepth 2
```

Limits recursive analysis to 2 levels deep. Default is 3.

### From Backup File

```powershell
Export-AppxDependencies -PackagePath "C:\Backups\MyApp.appx" -OutputPath "C:\Reports\deps.json"
```

Works on both installed packages and backup .appx files.

---

## 7. New-AppxBackupCertificate

**Purpose:** Create a self-signed certificate for code signing packages.

### Basic Certificate

```powershell
New-AppxBackupCertificate -Subject "CN=MyCompany" -OutputPath "C:\Certs\cert.cer"
```

Creates:
- Certificate in certificate store (CurrentUser\My)
- Exports public key to `.cer` file

Returns: `System.Security.Cryptography.X509Certificates.X509Certificate2` object with certificate details.

### With Custom Validity Period

```powershell
New-AppxBackupCertificate -Subject "CN=MyCompany" -OutputPath "C:\Certs\cert.cer" -ValidityYears 5
```

Default is 3 years. Valid range: 1-10 years.

### With Custom Key Length

```powershell
New-AppxBackupCertificate -Subject "CN=MyCompany" -OutputPath "C:\Certs\cert.cer" -KeyLength 2048
```

Valid values: 2048, 3072, 4096 bits. Default is 4096 (maximum security).

### Export Private Key (PFX)

```powershell
$pwd = ConvertTo-SecureString "YourPassword!" -AsPlainText -Force
New-AppxBackupCertificate -Subject "CN=MyCompany" -OutputPath "C:\Certs\cert.cer" `
    -Password $pwd -ExportPrivateKey
```

Creates both `.cer` (public) and `.pfx` (private key protected with password).

### Replace Existing Certificate

```powershell
New-AppxBackupCertificate -Subject "CN=MyCompany" -OutputPath "C:\Certs\cert.cer" -ReplaceExisting
```

Removes any existing certificates with the same subject before creating a new one.

### Overwrite File Without Prompting

```powershell
New-AppxBackupCertificate -Subject "CN=MyCompany" -OutputPath "C:\Certs\cert.cer" -Force
```

Overwrites existing `.cer` file if it exists.

### Different Certificate Store

```powershell
New-AppxBackupCertificate -Subject "CN=MyCompany" -OutputPath "C:\Certs\cert.cer" -StoreLocation "LocalMachine\My"
```

Default is `CurrentUser\My`. Use `LocalMachine\My` for system-wide certificates (requires admin).

---

## 8. Install-AppxBackup

**Purpose:** Install a backed-up package with automatic certificate trust.

### Basic Install (Auto-detect Certificate)

```powershell
Install-AppxBackup -PackagePath "C:\Backups\MyApp.appx"
```

Automatically looks for `MyApp.cer` in the same directory as the .appx file.

### With Explicit Certificate

```powershell
Install-AppxBackup -PackagePath "C:\Backups\MyApp.appx" -CertificatePath "C:\Certs\MyApp.cer"
```

Uses the specified certificate file.

### For Current User Only (No Admin)

```powershell
Install-AppxBackup -PackagePath "C:\Backups\MyApp.appx" -CertStoreLocation CurrentUser
```

Installs certificate to current user's store (no administrator required). Default is `LocalMachine` (system-wide, requires admin).

### Skip Certificate Installation

```powershell
Install-AppxBackup -PackagePath "C:\Backups\MyApp.appx" -SkipCertificate
```

**Use when:** Certificate is already trusted on the system.

### Force Reinstall

```powershell
Install-AppxBackup -PackagePath "C:\Backups\MyApp.appx" -Force
```

Reinstalls even if package is already installed.

### From ZIP Archive with Dependencies

```powershell
Install-AppxBackup -PackagePath "C:\Backups\MyApp.appxpack"
```

Automatically:
1. Extracts the .appxpack file
2. Installs all certificates
3. Installs packages in correct dependency order

### Extract to Custom Location

```powershell
Install-AppxBackup -PackagePath "C:\Backups\MyApp.appxpack" -ExtractPath "D:\Temp\Extract"
```

Extracts .appxpack contents to specified directory instead of temp.

### Skip Dependencies (ZIP Only)

```powershell
Install-AppxBackup -PackagePath "C:\Backups\MyApp.appxpack" -SkipDependencies
```

Installs only the main package from the .appxpack, skipping all dependencies.

### Allow Unsigned Packages

```powershell
Install-AppxBackup -PackagePath "C:\Backups\Unsigned.appx" -AllowUnsigned
```

**Requires:** Windows Developer Mode enabled. Use with caution.

### Pipeline Example

```powershell
Get-ChildItem "C:\Backups\*.appx" | Install-AppxBackup
```

Install multiple packages from a directory.

---

## Common Workflows

### Workflow 1: Backup and Restore Single App

```powershell
# Step 1: Backup
$app = Get-AppxPackage -Name "*Tivi*"
$result = Backup-AppxPackage -PackagePath $app.InstallLocation -OutputPath "D:\Backups" -Force

# Step 2: Transfer files to another computer

# Step 3: Install on new computer
Install-AppxBackup -PackagePath "D:\Backups\TiviMate_1.5.0.0_x64.appx"
```

### Workflow 2: Complete Migration with Dependencies

```powershell
# Step 1: Backup with all dependencies as single ZIP
$app = Get-AppxPackage -Name "*ComplexApp*"
$result = Backup-AppxPackage -PackagePath $app.InstallLocation `
    -OutputPath "D:\CompleteBackup" `
    -IncludeDependencies `
    -Force

# Step 2: Transfer single .appxpack file to another computer

# Step 3: Install everything (all packages + dependencies) from one file
Install-AppxBackup -PackagePath "D:\CompleteBackup\ComplexApp.appxpack"
```

### Workflow 3: Verify Backup Before Deployment

```powershell
$backupFile = "C:\Backups\MyApp.appx"

# Check integrity
$integrity = Test-AppxPackageIntegrity -PackagePath $backupFile -VerifySignature -CheckManifest
if (-not $integrity.IsValid) {
    Write-Host "Backup is corrupted - do not deploy"
    exit 1
}

# Check compatibility
$compat = Test-AppxBackupCompatibility -PackagePath $backupFile -CheckDependencies
if (-not $compat.IsCompatible) {
    Write-Host "Package not compatible with this system"
    exit 1
}

# Safe to install
Install-AppxBackup -PackagePath $backupFile
```

### Workflow 4: Document Dependencies for Compliance

```powershell
$app = Get-AppxPackage -Name "*MyApp*"

# Generate HTML report
Export-AppxDependencies -PackagePath $app.InstallLocation `
    -OutputPath "C:\Reports\MyApp_Dependencies.html" `
    -Format HTML `
    -Recursive `
    -IncludeOptional

# Open report
Invoke-Item "C:\Reports\MyApp_Dependencies.html"
```

### Workflow 5: Batch Backup Multiple Apps

```powershell
$backupDir = "D:\AllBackups"

Get-AppxPackage | Where-Object { $_.Publisher -like "*Microsoft*" } | ForEach-Object {
    Write-Host "Backing up: $($_.Name)"
    
    Backup-AppxPackage -PackagePath $_.InstallLocation `
        -OutputPath $backupDir `
        -Force -IncludeDependencies
}

Write-Host "All backups complete in: $backupDir"
```

---

## Tips

- **Always backup with `-IncludeDependencies`** when migrating to ensure nothing is missing.
- **Use `-Force`** in scripts to avoid interactive prompts.
- **Test compatibility** with `-CheckDependencies -Detailed` before deploying to new systems.
- **Export dependencies to HTML** for easy visual review and documentation.
- **Verify integrity** with `-VerifySignature -CheckManifest` before distributing backups.
- **Use `-DependencyReportOnly`** for lightweight analysis without creating large archives.
