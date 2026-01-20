<#
.SYNOPSIS
    Installs an APPX/MSIX package or ZIP archive with automatic certificate trust configuration.

.DESCRIPTION
    This script automates the installation of APPX/MSIX packages by:
    1. Detecting package type (.appx/.msix standalone or .appxpack ZIP archive)
    2. For ZIP archives: extracting, parsing manifest, orchestrating installation
    3. Installing certificates to the trusted store
    4. Installing packages in correct dependency order
    
    Designed as a companion to the AppxBackup module for seamless package restoration.

.PARAMETER PackagePath
    Path to the package file to install. Supports:
    - .appx/.msix (standalone packages)
    - .appxpack (ZIP archive with dependencies)
    - .appxbundle/.msixbundle (bundle packages)

.PARAMETER CertificatePath
    Path to the .cer certificate file. If not specified, looks for a .cer file
    with the same base name as the package in the same directory.
    For .appxpack files, certificates are automatically extracted from the archive.

.PARAMETER CertStoreLocation
    Certificate store location. Valid values:
    - 'LocalMachine' (default, requires Administrator, system-wide trust)
    - 'CurrentUser' (no admin needed, current user only)

.PARAMETER SkipCertificate
    Skip certificate installation. Use if certificate is already trusted.

.PARAMETER Force
    Force installation even if package is already installed (reinstall).

.PARAMETER ExtractPath
    For .appxpack files, specifies where to extract the archive.
    Default: Temporary directory

.PARAMETER SkipDependencies
    For .appxpack files, install only the main package, skipping dependencies.

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
    Author: DeltaGa
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
            if ($_ -notmatch '\.(appx|msix|appxbundle|msixbundle|appxpack|zip)$') {
                throw "File must be an APPX/MSIX package (.appx, .msix, .appxbundle, .msixbundle, .appxpack)"
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
    [switch]$Force,
    
    [Parameter()]
    [string]$ExtractPath,
    
    [Parameter()]
    [switch]$SkipDependencies
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
        $packageExtension = [System.IO.Path]::GetExtension($packageFullPath).ToLower()
        
        # Detect if this is a ZIP archive
        $isZipArchive = ($packageExtension -eq '.appxpack' -or $packageExtension -eq '.zip')
        
        Write-Host "[1/4] Package Analysis" -ForegroundColor Yellow
        Write-Host "  Package: $packageFileName" -ForegroundColor White
        Write-Host "  Location: $packageFullPath" -ForegroundColor Gray
        Write-Host "  Type: $(if ($isZipArchive) { 'ZIP Archive (.appxpack)' } else { 'Standalone Package' })" -ForegroundColor $(if ($isZipArchive) { 'Cyan' } else { 'Gray' })
        
        # Variables for cleanup
        $extractedPath = $null
        $manifestData = $null
        $packagesToInstall = @()
        
        if ($isZipArchive) {
            # ===== ZIP ARCHIVE HANDLING =====
            Write-Host "`n[2/4] ZIP Archive Extraction" -ForegroundColor Yellow
            
            # Determine extraction path
            if (-not $ExtractPath) {
                $extractPrefix = Get-AppxDefault -Category 'temporaryDirectories' -Key 'extractTempPrefix' -ConfigName 'ZipPackagingConfiguration' -FallbackValue 'AppxExtract_'
                $guidFormat = Get-AppxDefault -Category 'temporaryDirectories' -Key 'guidFormat' -ConfigName 'ZipPackagingConfiguration' -FallbackValue 'D'
                $extractedPath = [System.IO.Path]::Combine($env:TEMP, "$extractPrefix$((New-Guid).ToString($guidFormat))")
            }
            else {
                $extractedPath = $ExtractPath
            }
            
            Write-Host "  Extracting to: $extractedPath" -ForegroundColor Gray
            
            # Load System.IO.Compression for ZIP operations
            Add-Type -AssemblyName 'System.IO.Compression.FileSystem'
            
            try {
                # Extract ZIP archive
                [System.IO.Compression.ZipFile]::ExtractToDirectory($packageFullPath, $extractedPath)
                Write-Host "  Status: Extracted successfully" -ForegroundColor Green
            }
            catch {
                throw "Failed to extract ZIP archive: $_"
            }
            
            # Get directory names from configuration
            $packagesDirName = Get-AppxDefault -Category 'archiveStructure' -Key 'packagesDirectory' -ConfigName 'ZipPackagingConfiguration' -FallbackValue 'Packages'
            $certsDirName = Get-AppxDefault -Category 'archiveStructure' -Key 'certificatesDirectory' -ConfigName 'ZipPackagingConfiguration' -FallbackValue 'Certificates'
            $manifestFileName = Get-AppxDefault -Category 'archiveStructure' -Key 'manifestFileName' -ConfigName 'ZipPackagingConfiguration' -FallbackValue 'AppxBackupManifest.json'
            
            # Verify structure
            $packagesDir = [System.IO.Path]::Combine($extractedPath, $packagesDirName)
            $certsDir = [System.IO.Path]::Combine($extractedPath, $certsDirName)
            $manifestPath = [System.IO.Path]::Combine($extractedPath, $manifestFileName)
            
            if (-not (Test-Path -LiteralPath $manifestPath)) {
                throw "Invalid ZIP archive: $manifestFileName not found"
            }
            
            Write-Host "  Reading installation manifest..." -ForegroundColor Gray
            
            # Parse manifest
            try {
                $manifestJson = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8
                $manifestData = $manifestJson | ConvertFrom-Json
                Write-Host "  Manifest: Version $($manifestData.Version)" -ForegroundColor Gray
                Write-Host "  Packages: $($manifestData.TotalPackages) total" -ForegroundColor Gray
                Write-Host "  Main: $($manifestData.MainPackage.Name) v$($manifestData.MainPackage.Version)" -ForegroundColor Gray
            }
            catch {
                throw "Failed to parse manifest: $_"
            }
            
            # Build installation order
            if (-not $SkipDependencies.IsPresent -and $manifestData.Dependencies.Count -gt 0) {
                Write-Host "  Dependencies: $($manifestData.Dependencies.Count) found" -ForegroundColor Cyan
                
                # Install in order from manifest
                foreach ($depIdentifier in $manifestData.InstallationOrder) {
                    # Skip main package (will be added last)
                    if ($depIdentifier -like "*$($manifestData.MainPackage.Name)*$($manifestData.MainPackage.Version)*") {
                        continue
                    }
                    
                    # Find dependency in manifest
                    $dep = $manifestData.Dependencies | Where-Object {
                        "$($_.Name)_$($_.Version)_$($_.Architecture)" -eq $depIdentifier
                    } | Select-Object -First 1
                    
                    if ($dep) {
                        $depPackagePath = [System.IO.Path]::Combine($extractedPath, $dep.PackageFile)
                        $depCertPath = if ($dep.CertificateFile) {
                            [System.IO.Path]::Combine($extractedPath, $dep.CertificateFile)
                        } else { $null }
                        
                        $packagesToInstall += [PSCustomObject]@{
                            Name            = $dep.Name
                            Version         = $dep.Version
                            PackagePath     = $depPackagePath
                            CertificatePath = $depCertPath
                            Type            = 'Dependency'
                        }
                    }
                }
            }
            else {
                Write-Host "  Dependencies: Skipped" -ForegroundColor Gray
            }
            
            # Add main package last
            $mainPackagePath = [System.IO.Path]::Combine($extractedPath, $manifestData.MainPackage.PackageFile)
            $mainCertPath = if ($manifestData.MainPackage.CertificateFile) {
                [System.IO.Path]::Combine($extractedPath, $manifestData.MainPackage.CertificateFile)
            } else { $null }
            
            $packagesToInstall += [PSCustomObject]@{
                Name            = $manifestData.MainPackage.Name
                Version         = $manifestData.MainPackage.Version
                PackagePath     = $mainPackagePath
                CertificatePath = $mainCertPath
                Type            = 'Main'
            }
            
            Write-Host "`n  Installation Order:" -ForegroundColor Cyan
            $orderNum = 1
            foreach ($pkg in $packagesToInstall) {
                Write-Host "    $orderNum. $($pkg.Name) v$($pkg.Version) [$($pkg.Type)]" -ForegroundColor Gray
                $orderNum++
            }
        }
        else {
            # ===== STANDALONE PACKAGE HANDLING =====
            Write-Host "`n[2/4] Standalone Package Mode" -ForegroundColor Yellow
            
            # Auto-detect certificate if not specified
            if (-not $CertificatePath -and -not $SkipCertificate.IsPresent) {
                $packageDir = Split-Path -Path $packageFullPath -Parent
                $packageBaseName = [System.IO.Path]::GetFileNameWithoutExtension($packageFullPath)
                $autoCertPath = [System.IO.Path]::Combine($packageDir, "$packageBaseName.cer")
                
                if (Test-Path -LiteralPath $autoCertPath) {
                    $CertificatePath = $autoCertPath
                    Write-Host "  Certificate: Auto-detected ($packageBaseName.cer)" -ForegroundColor Green
                }
                else {
                    Write-Host "  Certificate: Not found (will attempt installation without cert)" -ForegroundColor Yellow
                    Write-Host "    Searched: $autoCertPath" -ForegroundColor Gray
                }
            }
            
            # Add single package to install list
            $packagesToInstall = @([PSCustomObject]@{
                Name            = [System.IO.Path]::GetFileNameWithoutExtension($packageFullPath)
                Version         = 'Unknown'
                PackagePath     = $packageFullPath
                CertificatePath = $CertificatePath
                Type            = 'Standalone'
            })
        }
        
        # ===== CERTIFICATE INSTALLATION =====
        Write-Host "`n[3/4] Certificate Installation" -ForegroundColor Yellow
        
        if (-not $SkipCertificate.IsPresent) {
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
            
            Write-Host "  Target Store: $certStorePath" -ForegroundColor Gray
            
            # Collect all unique certificates
            $uniqueCerts = @{}
            foreach ($pkg in $packagesToInstall) {
                if ($pkg.CertificatePath -and (Test-Path -LiteralPath $pkg.CertificatePath)) {
                    $certName = [System.IO.Path]::GetFileName($pkg.CertificatePath)
                    if (-not $uniqueCerts.ContainsKey($certName)) {
                        $uniqueCerts[$certName] = $pkg.CertificatePath
                    }
                }
            }
            
            Write-Host "  Certificates to install: $($uniqueCerts.Count)" -ForegroundColor Gray
            
            # Install each certificate
            $certCount = 0
            foreach ($certEntry in $uniqueCerts.GetEnumerator()) {
                $certCount++
                $certPath = $certEntry.Value
                $certName = $certEntry.Key
                
                if ($PSCmdlet.ShouldProcess($certName, "Install certificate to $certStorePath")) {
                    try {
                        # Read certificate
                        $newCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certPath)
                        $existingCerts = Get-ChildItem -Path $certStorePath -ErrorAction SilentlyContinue | 
                            Where-Object { $_.Thumbprint -eq $newCert.Thumbprint }
                        
                        if ($existingCerts) {
                            Write-Host "  [$certCount/$($uniqueCerts.Count)] $certName - Already installed" -ForegroundColor Green
                        }
                        else {
                            Import-Certificate -FilePath $certPath -CertStoreLocation $certStorePath -ErrorAction Stop | Out-Null
                            Write-Host "  [$certCount/$($uniqueCerts.Count)] $certName - Installed (Thumbprint: $($newCert.Thumbprint))" -ForegroundColor Green
                        }
                    }
                    catch {
                        Write-Host "  [$certCount/$($uniqueCerts.Count)] $certName - FAILED: $_" -ForegroundColor Red
                        Write-AppxLog -Message "Certificate installation failed for $certName`: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'
                        throw "Certificate installation failed for $certName`: $_"
                    }
                }
            }
        }
        else {
            Write-Host "  Status: Skipped" -ForegroundColor Gray
        }
        
        # ===== PACKAGE INSTALLATION =====
        Write-Host "`n[4/4] Package Installation" -ForegroundColor Yellow
        Write-Host "  Installing $($packagesToInstall.Count) package(s)..." -ForegroundColor Gray
        
        $installedCount = 0
        $failedPackages = @()
        
        foreach ($pkg in $packagesToInstall) {
            $installedCount++
            Write-Host "`n  [$installedCount/$($packagesToInstall.Count)] $($pkg.Name)" -ForegroundColor Cyan
            
            # Check if package exists
            if (-not (Test-Path -LiteralPath $pkg.PackagePath)) {
                Write-Host "    ERROR: Package file not found: $($pkg.PackagePath)" -ForegroundColor Red
                $failedPackages += $pkg.Name
                continue
            }
            
            # Check if already installed
            $installedPackage = Get-AppxPackage | Where-Object { $_.Name -eq $pkg.Name } | Select-Object -First 1
            
            if ($installedPackage -and -not $Force.IsPresent) {
                Write-Host "    Status: Already installed (v$($installedPackage.Version))" -ForegroundColor Yellow
                Write-Host "    Use -Force to reinstall" -ForegroundColor Gray
                continue
            }
            
            if ($PSCmdlet.ShouldProcess($pkg.Name, "Install APPX package")) {
                try {
                    Write-Host "    Installing..." -ForegroundColor White
                    
                    # Install package with appropriate parameters
                    $installParams = @{
                        Path = $pkg.PackagePath
                        ErrorAction = 'Stop'
                    }
                    
                    if ($Force.IsPresent) {
                        $installParams['ForceUpdateFromAnyVersion'] = $true
                    }
                    
                    Add-AppxPackage @installParams
                    
                    Write-Host "    Status: Installed successfully" -ForegroundColor Green
                    
                    # Get installed package info
                    $newPackage = Get-AppxPackage | Where-Object { $_.Name -eq $pkg.Name } | Select-Object -First 1
                    if ($newPackage) {
                        Write-Host "    Version: $($newPackage.Version)" -ForegroundColor Gray
                    }
                }
                catch {
                    Write-Host "    ERROR: Installation failed" -ForegroundColor Red
                    Write-Host "    Details: $_" -ForegroundColor Red
                    
                    # Provide helpful error messages
                    if ($_ -match '0x800B0109') {
                        Write-Host "    CAUSE: Certificate not trusted" -ForegroundColor Yellow
                        Write-Host "    SOLUTION: Run as Administrator or install certificate manually" -ForegroundColor Yellow
                    }
                    elseif ($_ -match '0x80073CF9') {
                        Write-Host "    CAUSE: Package conflict or already installed" -ForegroundColor Yellow
                        Write-Host "    SOLUTION: Use -Force parameter to reinstall" -ForegroundColor Yellow
                    }
                    elseif ($_ -match '0x80073CF6') {
                        Write-Host "    CAUSE: Package signature invalid" -ForegroundColor Yellow
                        Write-Host "    SOLUTION: Verify certificate is installed and trusted" -ForegroundColor Yellow
                    }
                    
                    $failedPackages += $pkg.Name
                    Write-AppxLog -Message "Package installation failed for $($pkg.Name): $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'
                    
                    # For ZIP archives, continue with remaining packages; for standalone, throw
                    if (-not $isZipArchive) {
                        throw
                    }
                }
            }
        }
        
        # Success summary
        Write-Host "`n=== Installation Complete ===" -ForegroundColor Green
        Write-Host "Packages Installed: $($packagesToInstall.Count - $failedPackages.Count) of $($packagesToInstall.Count)" -ForegroundColor White
        
        if ($failedPackages.Count -gt 0) {
            Write-Host "Failed Packages: $($failedPackages.Count)" -ForegroundColor Red
            foreach ($failed in $failedPackages) {
                Write-Host "  - $failed" -ForegroundColor Red
            }
        }
        
        if (-not $SkipCertificate.IsPresent -and $uniqueCerts -and $uniqueCerts.Count -gt 0) {
            Write-Host "Certificates: $($uniqueCerts.Count) installed to $CertStoreLocation store" -ForegroundColor White
        }
        Write-Host ""
    }
    catch {
        Write-Host "`n=== Installation Failed ===" -ForegroundColor Red
        Write-Host "Error: $_" -ForegroundColor Red
        Write-Host ""
        throw
    }
    finally {
        # Cleanup extracted files
        if ($extractedPath -and (Test-Path -LiteralPath $extractedPath) -and -not $ExtractPath) {
            try {
                Remove-Item -Path $extractedPath -Recurse -Force -ErrorAction SilentlyContinue
                Write-AppxLog -Message "Cleaned up extracted files: $extractedPath" -Level 'Debug'
            }
            catch {
                Write-AppxLog -Message "Failed to cleanup extracted files: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
            }
        }
    }
}

end {
    # Cleanup
}

} # End function Install-AppxBackup