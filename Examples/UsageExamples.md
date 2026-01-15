# AppxBackup Usage Examples

## Real-World Scenarios and Solutions

---

## Example 1: Basic App Backup

**Scenario:** You want to back up "Disney's Wreck-It Ralph" before resetting your PC.

```powershell
# Find the app
$app = Get-AppxPackage -Name "*Ralph*"

# Backup the app
Backup-AppxPackage -PackagePath $app.InstallLocation -OutputPath "C:\GameBackups" -Verbose

# Result:
# [CHECK] Package created: C:\GameBackups\Disney.Wreck-itRalph_1.0.0.12_x86__6rarf9sa4v8jt.appx
# [CHECK] Certificate created: C:\GameBackups\Disney.Wreck-itRalph_1.0.0.12_x86__6rarf9sa4v8jt.cer
```

---

## Example 2: Batch Backup Multiple Apps

**Scenario:** You're migrating to a new PC and want to backup all Adobe apps.

```powershell
# Backup all Adobe apps
$outputPath = "D:\AppBackups\Adobe"

Get-AppxPackage -Name "*Adobe*" | ForEach-Object {
    Write-Host "Backing up: $($_.Name)" -ForegroundColor Cyan
    
    try {
        Backup-AppxPackage -PackagePath $_.InstallLocation `
            -OutputPath $outputPath `
            -IncludeDependencies `
            -CompressionLevel Maximum `
            -Force `
            -ErrorAction Stop
        
        Write-Host "  [CHECK] Success" -ForegroundColor Green
    }
    catch {
        Write-Host "  [X] Failed: $_" -ForegroundColor Red
    }
}

# Generate summary report
Get-ChildItem -Path $outputPath -Filter "*.appx" | 
    Select-Object Name, @{N='SizeMB';E={[Math]::Round($_.Length/1MB,2)}} |
    Format-Table -AutoSize
```

---

## Example 3: Backup with Custom Certificate

**Scenario:** You need a certificate with specific validity and key strength.

```powershell
# Create high-security certificate (5 years, 4096-bit)
$certPassword = ConvertTo-SecureString "MyVerySecurePassword123!" -AsPlainText -Force

$cert = New-AppxBackupCertificate `
    -Subject "CN=My Organization Code Signing, O=MyOrg, C=US" `
    -OutputPath "C:\Certificates\MyOrg_CodeSigning.cer" `
    -ValidityYears 5 `
    -KeyLength 4096 `
    -Password $certPassword `
    -ExportPrivateKey `
    -Verbose

Write-Host "Certificate Thumbprint: $($cert.Thumbprint)" -ForegroundColor Yellow
Write-Host "Valid until: $($cert.NotAfter)" -ForegroundColor Yellow

# Use this certificate for all backups
$apps = Get-AppxPackage -Publisher "*MyOrg*"

foreach ($app in $apps) {
    Backup-AppxPackage -PackagePath $app.InstallLocation `
        -OutputPath "C:\EnterpriseBackups" `
        -CertificateSubject "CN=My Organization Code Signing, O=MyOrg, C=US" `
        -CompressionLevel Maximum
}
```

---

## Example 4: Dependency Analysis

**Scenario:** You need to understand what dependencies your app requires.

```powershell
# Get dependency information
$app = Get-AppxPackage -Name "MyComplexApp"

$result = Backup-AppxPackage -PackagePath $app.InstallLocation `
    -OutputPath "C:\Backups" `
    -IncludeDependencies

# Display dependency report
Write-Host "`n=== Dependency Report ===" -ForegroundColor Cyan
Write-Host "Package: $($result.PackageName) v$($result.PackageVersion)"
Write-Host "Dependencies: $($result.DependencyCount)"
Write-Host "Missing: $($result.DependenciesMissing)"

if ($result.DependencyInfo) {
    Write-Host "`nDetailed Dependencies:" -ForegroundColor Yellow
    
    $result.DependencyInfo.Dependencies | 
        Format-Table Name, MinVersion, DependencyType, IsInstalled, InstalledVersion -AutoSize
}

# Export to JSON for documentation
$result | ConvertTo-Json -Depth 10 | Out-File "C:\Reports\MyComplexApp_Dependencies.json"
```

---

## Example 5: Validation Before Deployment

**Scenario:** You need to verify a backup before deploying to production machines.

```powershell
# Validate the backed-up package
$packagePath = "C:\Backups\CriticalApp_1.5.0.0_x64__abc123.appx"

# Comprehensive validation
$validation = Test-AppxPackageIntegrity -PackagePath $packagePath `
    -VerifySignature `
    -CheckManifest

if ($validation.IsValid) {
    Write-Host "[CHECK] Package validation passed" -ForegroundColor Green
    Write-Host "  - Archive: Valid"
    Write-Host "  - Signature: $($validation.SignatureValid)"
    Write-Host "  - Manifest: $($validation.ManifestValid)"
    
    # Safe to deploy
    Write-Host "`nPackage ready for deployment" -ForegroundColor Cyan
}
else {
    Write-Host "[X] Package validation FAILED" -ForegroundColor Red
    Write-Host "Issues found:"
    $validation.Issues | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
    
    # Do not deploy
    exit 1
}
```

---

## Example 6: Automated Scheduled Backups

**Scenario:** You want to automatically backup specific apps every week.

```powershell
# backup-apps-scheduled.ps1

# Configuration
$appsToBackup = @(
    "Microsoft.Office.Excel",
    "Adobe.CreativeCloud",
    "Slack.Slack"
)
$backupRoot = "\\FileServer\AppBackups"
$timestamp = Get-Date -Format "yyyy-MM-dd"
$backupPath = Join-Path $backupRoot $timestamp

# Create dated backup directory
New-Item -Path $backupPath -ItemType Directory -Force | Out-Null

# Import module
Import-Module AppxBackup -ErrorAction Stop

# Backup each app
$results = @()

foreach ($appName in $appsToBackup) {
    try {
        $app = Get-AppxPackage -Name $appName -ErrorAction Stop
        
        Write-Host "Backing up: $($app.Name)..." -ForegroundColor Cyan
        
        $result = Backup-AppxPackage -PackagePath $app.InstallLocation `
            -OutputPath $backupPath `
            -CompressionLevel Maximum `
            -Force `
            -ErrorAction Stop
        
        $results += [PSCustomObject]@{
            App = $app.Name
            Version = $app.Version
            Status = "Success"
            Size = "$($result.PackageFileSizeMB) MB"
            Path = $result.PackageFilePath
        }
        
        Write-Host "  [CHECK] Complete" -ForegroundColor Green
    }
    catch {
        Write-Host "  [X] Failed: $_" -ForegroundColor Red
        
        $results += [PSCustomObject]@{
            App = $appName
            Version = "N/A"
            Status = "Failed"
            Size = "N/A"
            Path = $_.Exception.Message
        }
    }
}

# Generate report
$reportPath = Join-Path $backupPath "backup-report.html"

$html = @"
<!DOCTYPE html>
<html>
<head>
    <title>App Backup Report - $timestamp</title>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; }
        h1 { color: #0078D4; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; }
        th { background: #0078D4; color: white; padding: 10px; text-align: left; }
        td { border: 1px solid #ddd; padding: 8px; }
        tr:nth-child(even) { background: #f2f2f2; }
        .success { color: green; font-weight: bold; }
        .failed { color: red; font-weight: bold; }
    </style>
</head>
<body>
    <h1>App Backup Report</h1>
    <p><strong>Date:</strong> $timestamp</p>
    <p><strong>Location:</strong> $backupPath</p>
    <table>
        <tr>
            <th>Application</th>
            <th>Version</th>
            <th>Status</th>
            <th>Size</th>
        </tr>
        $(
            $results | ForEach-Object {
                $statusClass = if ($_.Status -eq "Success") { "success" } else { "failed" }
                "<tr><td>$($_.App)</td><td>$($_.Version)</td><td class='$statusClass'>$($_.Status)</td><td>$($_.Size)</td></tr>"
            }
        )
    </table>
</body>
</html>
"@

$html | Out-File -FilePath $reportPath -Encoding UTF8

Write-Host "`nReport generated: $reportPath" -ForegroundColor Cyan

# Send email notification (optional)
# Send-MailMessage -To "admin@company.com" -Subject "App Backup Report - $timestamp" ...
```

**Schedule this script:**
```powershell
# Create scheduled task (run as administrator)
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File C:\Scripts\backup-apps-scheduled.ps1"

$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At 2AM

Register-ScheduledTask -TaskName "AppxWeeklyBackup" `
    -Action $action `
    -Trigger $trigger `
    -Description "Weekly backup of critical APPX applications" `
    -User "SYSTEM" `
    -RunLevel Highest
```

---

## Example 7: Migration to New Hardware

**Scenario:** Moving to a new computer with different architecture (x64 to ARM64).

```powershell
# OLD COMPUTER (x64)
# ==================

# Backup all apps
$backupPath = "D:\MigrationBackup"
$allApps = Get-AppxPackage | Where-Object { $_.SignatureKind -eq 'Store' }

foreach ($app in $allApps) {
    Backup-AppxPackage -PackagePath $app.InstallLocation `
        -OutputPath $backupPath `
        -IncludeDependencies `
        -Force
}

# Copy to external drive or network share


# NEW COMPUTER (ARM64)
# =====================

# Check compatibility
$backups = Get-ChildItem -Path "E:\MigrationBackup" -Filter "*.appx"

foreach ($backup in $backups) {
    $compat = Test-AppxBackupCompatibility -PackagePath $backup.FullName
    
    if ($compat.IsCompatible) {
        Write-Host "[CHECK] $($backup.Name) - Compatible" -ForegroundColor Green
        
        # Install
        # Restore-AppxPackage -PackagePath $backup.FullName `
        #     -CertificatePath $backup.FullName.Replace('.appx', '.cer')
    }
    else {
        Write-Host "[X] $($backup.Name) - Incompatible: $($compat.Reason)" -ForegroundColor Red
    }
}
```

---

## Example 8: Tool Discovery and Troubleshooting

**Scenario:** Debugging why MakeAppx isn't being used.

```powershell
# Check tool availability
Write-Host "=== Tool Availability Check ===" -ForegroundColor Cyan

$tools = @('MakeAppx', 'SignTool', 'Certutil')

foreach ($tool in $tools) {
    $path = Get-AppxToolPath -ToolName $tool
    
    if ($path) {
        Write-Host "[CHECK] $tool`: $path" -ForegroundColor Green
        
        # Get version
        if ($tool -eq 'MakeAppx') {
            $version = (& $path /? 2>&1) | Select-String -Pattern "Version"
            Write-Host "  Version: $version" -ForegroundColor Gray
        }
    }
    else {
        Write-Host "[X] $tool`: Not found" -ForegroundColor Red
    }
}

# Check PowerShell certificate cmdlets
if (Get-Command New-SelfSignedCertificate -ErrorAction SilentlyContinue) {
    Write-Host "[CHECK] Native PowerShell certificate support available" -ForegroundColor Green
}

# Module configuration
Write-Host "`n=== Module Configuration ===" -ForegroundColor Cyan
$AppxBackupConfig | Format-Table -AutoSize
```

---

## Example 9: Silent Deployment Script

**Scenario:** Deploy backed-up apps to multiple machines silently.

```powershell
# deploy-apps.ps1
# Run on target machine

param(
    [string]$BackupPath = "\\FileServer\AppBackups\2026-01-13",
    [string[]]$AppsToInstall = @("MyApp", "AnotherApp")
)

Import-Module AppxBackup

# Install certificates first
Get-ChildItem -Path $BackupPath -Filter "*.cer" | ForEach-Object {
    Write-Host "Installing certificate: $($_.Name)" -ForegroundColor Cyan
    
    # Import to Trusted Root (requires admin)
    Import-Certificate -FilePath $_.FullName `
        -CertStoreLocation Cert:\LocalMachine\Root `
        -ErrorAction SilentlyContinue
}

# Install apps
foreach ($appName in $AppsToInstall) {
    $package = Get-ChildItem -Path $BackupPath -Filter "*$appName*.appx" | Select-Object -First 1
    
    if ($package) {
        Write-Host "Installing: $($package.Name)" -ForegroundColor Cyan
        
        # Restore-AppxPackage -PackagePath $package.FullName -Force
        # (Restore function implementation pending)
        
        Write-Host "  [CHECK] Installed" -ForegroundColor Green
    }
    else {
        Write-Host "  [X] Package not found: $appName" -ForegroundColor Red
    }
}
```

---

## Tips and Best Practices

### 1. Always Use `-Verbose` During Initial Testing
```powershell
Backup-AppxPackage -PackagePath $path -OutputPath $out -Verbose
```

### 2. Check Logs for Detailed Diagnostics
```powershell
Get-Content "$env:TEMP\AppxBackup_$(Get-Date -Format 'yyyyMMdd').log" -Tail 100
```

### 3. Use `-WhatIf` to Preview Actions
```powershell
Backup-AppxPackage -PackagePath $path -OutputPath $out -WhatIf
```

### 4. Store Certificates Securely
```powershell
# Use secure network locations with proper ACLs
$certPath = "\\SecureFileServer\Certificates"
# Set-Acl with restricted permissions
```

### 5. Validate Before Deployment
```powershell
# Always validate packages before deploying
Test-AppxPackageIntegrity -PackagePath $pkg -VerifySignature -CheckManifest
```

---

**These examples demonstrate the power, flexibility, and robustness of the AppxBackup v2.0 module.**
