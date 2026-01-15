<#
.SYNOPSIS
    Installs an APPX/MSIX package with automatic certificate trust configuration.

.DESCRIPTION
    This script automates the installation of APPX/MSIX packages by:
    1. Installing the associated certificate to the trusted store
    2. Installing the APPX package
    
    Designed as a companion to the AppxBackup module for seamless package restoration.

.PARAMETER PackagePath
    Path to the .appx or .msix package file to install.

.PARAMETER CertificatePath
    Path to the .cer certificate file. If not specified, looks for a .cer file
    with the same base name as the package in the same directory.

.PARAMETER CertStoreLocation
    Certificate store location. Valid values:
    - 'LocalMachine' (default, requires Administrator, system-wide trust)
    - 'CurrentUser' (no admin needed, current user only)

.PARAMETER SkipCertificate
    Skip certificate installation. Use if certificate is already trusted.

.PARAMETER Force
    Force installation even if package is already installed (reinstall).

.EXAMPLE
    Install-AppxBackup -PackagePath "C:\Backups\MyApp.appx"
    
    Installs MyApp.appx and automatically finds MyApp.cer in the same directory.

.EXAMPLE
    Install-AppxBackup -PackagePath "C:\Backups\MyApp.appx" -CertificatePath "C:\Certs\MyApp.cer"
    
    Installs MyApp.appx using the specified certificate file.

.EXAMPLE
    Install-AppxBackup -PackagePath "C:\Backups\MyApp.appx" -CertStoreLocation CurrentUser
    
    Installs for current user only (no Administrator required).

.EXAMPLE
    Install-AppxBackup -PackagePath "C:\Backups\MyApp.appx" -SkipCertificate
    
    Installs package without installing certificate (assumes already trusted).

.NOTES
    Author: CAN (Code Anything Now)
    Version: 1.0.0
    Requires: PowerShell 5.1+, Windows 10 1809+
#>

function Install-AppxBackup {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Path', 'AppxPath', 'PackageFile')]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            if (-not (Test-Path -LiteralPath $_)) {
                throw "Package file not found: $_"
            }
            if ($_ -notmatch '\.(appx|msix|appxbundle|msixbundle)$') {
                throw "File must be an APPX/MSIX package (.appx, .msix, .appxbundle, .msixbundle)"
            }
            $true
        })]
        [string]$PackagePath,

        [Parameter()]
        [Alias('Cert', 'CertPath', 'Certificate')]
        [ValidateScript({
            if ($_ -and -not (Test-Path -LiteralPath $_)) {
                throw "Certificate file not found: $_"
            }
            if ($_ -and $_ -notmatch '\.cer$') {
                throw "Certificate file must be a .cer file"
            }
        $true
    })]
    [string]$CertificatePath,

    [Parameter()]
    [ValidateSet('LocalMachine', 'CurrentUser')]
    [string]$CertStoreLocation = 'LocalMachine',

    [Parameter()]
    [switch]$SkipCertificate,

    [Parameter()]
    [switch]$Force
)

begin {
    $ErrorActionPreference = 'Stop'
    
    # Check PowerShell version
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        throw "This script requires PowerShell 5.1 or later. Current version: $($PSVersionTable.PSVersion)"
    }
    
    Write-Host "`n=== APPX Package Installation Script ===" -ForegroundColor Cyan
    Write-Host "Version: 1.0.0`n" -ForegroundColor Gray
}

process {
    try {
        # Resolve full paths
        $packageFullPath = Resolve-Path -LiteralPath $PackagePath | Select-Object -ExpandProperty Path
        $packageFileName = Split-Path -Path $packageFullPath -Leaf
        
        Write-Host "[1/3] Package Analysis" -ForegroundColor Yellow
        Write-Host "  Package: $packageFileName" -ForegroundColor White
        Write-Host "  Location: $packageFullPath" -ForegroundColor Gray
        
        # Auto-detect certificate if not specified
        if (-not $CertificatePath -and -not $SkipCertificate.IsPresent) {
            $packageDir = Split-Path -Path $packageFullPath -Parent
            $packageBaseName = [System.IO.Path]::GetFileNameWithoutExtension($packageFullPath)
            $autoCertPath = Join-Path $packageDir "$packageBaseName.cer"
            
            if (Test-Path -LiteralPath $autoCertPath) {
                $CertificatePath = $autoCertPath
                Write-Host "  Certificate: Auto-detected ($packageBaseName.cer)" -ForegroundColor Green
            }
            else {
                Write-Host "  Certificate: Not found (will attempt installation without cert)" -ForegroundColor Yellow
                Write-Host "    Searched: $autoCertPath" -ForegroundColor Gray
                $SkipCertificate = $true
            }
        }
        
        # Certificate Installation
        if (-not $SkipCertificate.IsPresent -and $CertificatePath) {
            Write-Host "`n[2/3] Certificate Installation" -ForegroundColor Yellow
            
            $certFullPath = Resolve-Path -LiteralPath $CertificatePath | Select-Object -ExpandProperty Path
            $certFileName = Split-Path -Path $certFullPath -Leaf
            
            Write-Host "  Certificate: $certFileName" -ForegroundColor White
            Write-Host "  Target Store: Cert:\$CertStoreLocation\Root" -ForegroundColor Gray
            
            # Determine store path
            $certStorePath = if ($CertStoreLocation -eq 'LocalMachine') {
                'Cert:\LocalMachine\Root'
            } else {
                'Cert:\CurrentUser\Root'
            }
            
            # Check if running as Administrator for LocalMachine
            if ($CertStoreLocation -eq 'LocalMachine') {
                $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
                
                if (-not $isAdmin) {
                    Write-Host "  WARNING: Not running as Administrator" -ForegroundColor Red
                    Write-Host "  Cannot install to LocalMachine store. Switching to CurrentUser..." -ForegroundColor Yellow
                    $certStorePath = 'Cert:\CurrentUser\Root'
                    $CertStoreLocation = 'CurrentUser'
                }
            }
            
            if ($PSCmdlet.ShouldProcess($certFileName, "Install certificate to $certStorePath")) {
                try {
                    # Read certificate to check if already installed
                    $newCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certFullPath)
                    $existingCerts = Get-ChildItem -Path $certStorePath -ErrorAction SilentlyContinue | 
                        Where-Object { $_.Thumbprint -eq $newCert.Thumbprint }
                    
                    if ($existingCerts) {
                        Write-Host "  Status: Already installed (Thumbprint: $($newCert.Thumbprint))" -ForegroundColor Green
                    }
                    else {
                        Import-Certificate -FilePath $certFullPath -CertStoreLocation $certStorePath -ErrorAction Stop | Out-Null
                        Write-Host "  Status: Installed successfully" -ForegroundColor Green
                        Write-Host "  Thumbprint: $($newCert.Thumbprint)" -ForegroundColor Gray
                    }
                }
                catch {
                    Write-Host "  ERROR: Certificate installation failed" -ForegroundColor Red
                    Write-Host "  Details: $_" -ForegroundColor Red
                    throw "Certificate installation failed: $_"
                }
            }
        }
        else {
            Write-Host "`n[2/3] Certificate Installation" -ForegroundColor Yellow
            Write-Host "  Status: Skipped" -ForegroundColor Gray
        }
        
        # Package Installation
        Write-Host "`n[3/3] Package Installation" -ForegroundColor Yellow
        
        # Check if package is already installed
        $packageBaseName = [System.IO.Path]::GetFileNameWithoutExtension($packageFullPath)
        $installedPackage = Get-AppxPackage | Where-Object { $_.Name -like "*$packageBaseName*" } | Select-Object -First 1
        
        if ($installedPackage -and -not $Force.IsPresent) {
            Write-Host "  Status: Package already installed" -ForegroundColor Yellow
            Write-Host "  Name: $($installedPackage.Name)" -ForegroundColor Gray
            Write-Host "  Version: $($installedPackage.Version)" -ForegroundColor Gray
            Write-Host "  Use -Force to reinstall" -ForegroundColor Gray
            return
        }
        
        if ($PSCmdlet.ShouldProcess($packageFileName, "Install APPX package")) {
            try {
                Write-Host "  Installing package..." -ForegroundColor White
                
                # Install package with appropriate parameters
                $installParams = @{
                    Path = $packageFullPath
                    ErrorAction = 'Stop'
                }
                
                if ($Force.IsPresent) {
                    $installParams['ForceUpdateFromAnyVersion'] = $true
                }
                
                Add-AppxPackage @installParams
                
                Write-Host "  Status: Installed successfully" -ForegroundColor Green
                
                # Get installed package info
                $newPackage = Get-AppxPackage | Where-Object { $_.Name -like "*$packageBaseName*" } | Select-Object -First 1
                if ($newPackage) {
                    Write-Host "  Name: $($newPackage.Name)" -ForegroundColor Gray
                    Write-Host "  Version: $($newPackage.Version)" -ForegroundColor Gray
                    Write-Host "  Install Location: $($newPackage.InstallLocation)" -ForegroundColor Gray
                }
            }
            catch {
                Write-Host "  ERROR: Package installation failed" -ForegroundColor Red
                Write-Host "  Details: $_" -ForegroundColor Red
                
                # Provide helpful error messages
                if ($_ -match '0x800B0109') {
                    Write-Host "`n  CAUSE: Certificate not trusted" -ForegroundColor Yellow
                    Write-Host "  SOLUTION: Run this script as Administrator to install certificate to LocalMachine store" -ForegroundColor Yellow
                }
                elseif ($_ -match '0x80073CF9') {
                    Write-Host "`n  CAUSE: Package already installed with different version" -ForegroundColor Yellow
                    Write-Host "  SOLUTION: Use -Force parameter to reinstall" -ForegroundColor Yellow
                }
                elseif ($_ -match '0x80073CF6') {
                    Write-Host "`n  CAUSE: Package signature invalid or tampered" -ForegroundColor Yellow
                    Write-Host "  SOLUTION: Re-create package from backup" -ForegroundColor Yellow
                }
                
                throw
            }
        }
        
        # Success summary
        Write-Host "`n=== Installation Complete ===" -ForegroundColor Green
        Write-Host "Package: $packageFileName" -ForegroundColor White
        if (-not $SkipCertificate.IsPresent -and $CertificatePath) {
            Write-Host "Certificate: Installed to $CertStoreLocation store" -ForegroundColor White
        }
        Write-Host ""
    }
    catch {
        Write-Host "`n=== Installation Failed ===" -ForegroundColor Red
        Write-Host "Error: $_" -ForegroundColor Red
        Write-Host ""
        throw
    }
}

end {
    # Cleanup
}

} # End function Install-AppxBackup