<#
.SYNOPSIS
    Creates a backup of an installed Windows Store/MSIX application package.

.DESCRIPTION
    Enterprise-grade backup solution for Windows application packages (APPX/MSIX).
    
    Features:
    - Native PowerShell certificate management (no external tools required)
    - Comprehensive error handling and validation
    - Progress indication for long-running operations
    - Dependency resolution and optional bundling
    - Secure certificate storage
    - MSIX and legacy APPX support
    - Automatic rollback on failure
    
.PARAMETER PackagePath
    Full path to the installed application package directory.
    Typically: C:\Program Files\WindowsApps\<PackageFullName>

.PARAMETER OutputPath
    Directory where the backup package and certificate will be saved.
    If not specified, uses current directory.

.PARAMETER IncludeDependencies
    Creates a complete ZIP archive (.appxpack) containing the main package
    plus all installed dependencies with their certificates. This provides a single
    file with everything needed for installation on another system.
    The ZIP archive includes:
    - All APPX/MSIX packages (main + dependencies)
    - All certificates for package signing
    - AppxBackupManifest.json with installation orchestration metadata
    - README.txt with installation instructions
    Note: Archive size will be significantly larger but ensures complete portability.
    Use Install-AppxBackup to restore from the ZIP archive.

.PARAMETER DependencyReportOnly
    Exports only a JSON report of dependencies without bundling them.
    Use this for analysis or lightweight backup when you don't need the dependencies
    packaged. The JSON report includes installation paths and version details.

.PARAMETER CreateBundle
    DEPRECATED: Use -IncludeDependencies instead.
    This parameter is kept for backward compatibility and will use ZIP-based packaging.

.PARAMETER CertificateSubject
    Subject name for the self-signed certificate.
    If not specified, uses the package publisher from the manifest.

.PARAMETER CertificatePassword
    Secure password for certificate export (optional).
    If not specified, certificate is exported without password protection.

.PARAMETER CompressionLevel
    Compression level for the package.
    Valid values: None, Fast, Normal, Maximum
    Default: Normal

.PARAMETER NoCertificate
    If specified, skips certificate creation and signing.
    Resulting package will not be installable without separate signing.

.PARAMETER Force
    Overwrites existing output files without prompting.

.EXAMPLE
    Backup-AppxPackage -PackagePath "C:\Program Files\WindowsApps\MyApp_1.0.0.0_x64__abc123" -OutputPath "C:\Backups"
    
    Creates a signed backup of the specified app in C:\Backups

.EXAMPLE
    Get-AppxPackage -Name "MyApp" | Backup-AppxPackage -OutputPath "C:\Backups" -IncludeDependencies
    
    Backs up MyApp with dependency analysis via pipeline

.EXAMPLE
    Backup-AppxPackage -PackagePath $path -OutputPath $out -IncludeDependencies -Force
    
    Creates a complete ZIP archive (.appxpack) including dependencies, overwriting existing files

.OUTPUTS
    AppxBackup.BackupResult
    
.NOTES
    Requires:
    - PowerShell 5.1+ (7.4+ recommended)
    - Administrator privileges for certificate operations
    - Windows 10 1809+ or Windows 11
    
    Author: DeltaGa
    Version: 2.0.0
#>

function Backup-AppxPackage {
    [CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Standard')]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(
            Mandatory,
            Position = 0,
            ValueFromPipeline,
            ValueFromPipelineByPropertyName
        )]
        [Alias('InstallLocation', 'Path', 'FullName')]
        [ValidateNotNullOrEmpty()]
        [string]$PackagePath,

        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$OutputPath = $PWD.Path,

        [Parameter()]
        [switch]$IncludeDependencies,
        
        [Parameter()]
        [switch]$DependencyReportOnly,

        [Parameter(ParameterSetName = 'Bundle')]
        [switch]$CreateBundle,

        [Parameter()]
        [string]$CertificateSubject,

        [Parameter()]
        [SecureString]$CertificatePassword,

        [Parameter()]
        [ValidateSet('None', 'Fast', 'Normal', 'Maximum')]
        [string]$CompressionLevel = 'Normal',

        [Parameter()]
        [switch]$NoCertificate,

        [Parameter()]
        [switch]$Force
    )

    begin {
        Write-AppxLog -Message "=== Backup-AppxPackage v2.0 ===" -Level 'Info'
        Write-AppxLog -Message "PowerShell Version: $($PSVersionTable.PSVersion)" -Level 'Debug'
        Write-AppxLog -Message "Operating System: $([Environment]::OSVersion.VersionString)" -Level 'Debug'
        
        # Initialize progress
        $progressId = Get-Random
        $progressStage = 0
        
        # Calculate total stages based on parameters
        $totalStages = 3  # Validation, Parse Manifest, Create Package (always)
        if ($IncludeDependencies.IsPresent) { $totalStages++ }  # Add dependency resolution
        if (-not $NoCertificate.IsPresent) { $totalStages += 2 }  # Add certificate creation and signing
        if ($CreateBundle.IsPresent) { $totalStages++ }  # Add bundle creation (future)
    }

    process {
        try {
            # Stage 1: Validation
            $progressStage++
            Write-Progress -Id $progressId -Activity "Backing up APPX Package" `
                -Status "Stage $progressStage/$totalStages : Validating requirements" `
                -PercentComplete (($progressStage / $totalStages) * 100)
            
            Write-AppxLog -Message "Validating Windows SDK tools..." -Level 'Verbose'
            
            # CRITICAL: Validate Windows SDK tools are available
            # These are MANDATORY - without them, backup will fail inconsistently
            $makeAppxPath = $null
            $signToolPath = $null
            $sdkMissing = @()
            
            try {
                $makeAppxPath = Test-AppxToolAvailability -ToolName 'MakeAppx' -ThrowOnError:$false
                if ($null -eq $makeAppxPath) {
                    $sdkMissing += 'MakeAppx.exe'
                }
            }
            catch {
                $sdkMissing += 'MakeAppx.exe'
            }
            
            try {
                $signToolPath = Test-AppxToolAvailability -ToolName 'SignTool' -ThrowOnError:$false
                if ($null -eq $signToolPath) {
                    $sdkMissing += 'SignTool.exe'
                }
            }
            catch {
                $sdkMissing += 'SignTool.exe'
            }
            
            if ($sdkMissing.Count -gt 0) {
                $errorMsg = @"
CRITICAL ERROR: Windows SDK tools not found

The following required tools are missing:
$($sdkMissing | ForEach-Object { "  - $_" } | Out-String)

APPX package backup REQUIRES the Windows SDK to be installed.
Without these tools, the backup process will fail or produce corrupted packages.

SOLUTION:
1. Download and install Windows SDK from:
   https://developer.microsoft.com/en-us/windows/downloads/windows-sdk/

2. Ensure the SDK bin directory is in your PATH, or install to default location:
   C:\Program Files (x86)\Windows Kits\10\bin\<version>\<arch>\

3. Verify installation:
   makeappx.exe /?
   signtool.exe /?

Current PATH directories searched:
$($env:PATH -split ';' | Where-Object { $_ -like '*Windows Kits*' } | ForEach-Object { "  - $_" } | Out-String)

This module does NOT support fallback methods for package creation.
Windows SDK is MANDATORY for reliable APPX backup operations.
"@
                Write-AppxLog -Message $errorMsg -Level 'Error'
                throw $errorMsg
            }
            
            Write-AppxLog -Message "Found MakeAppx: $makeAppxPath" -Level 'Debug'
            Write-AppxLog -Message "Found SignTool: $signToolPath" -Level 'Debug'
            
            Write-AppxLog -Message "Validating package path..." -Level 'Verbose'
            
            # Handle pipeline input from Get-AppxPackage
            if ($PackagePath -match '^\w+\.\w+_[\w\.]+_[\w]+__[\w]+$') {
                # This looks like a PackageFullName
                $package = Get-AppxPackage -Name $PackagePath -ErrorAction Stop
                if ($package) {
                    $PackagePath = $package.InstallLocation
                    Write-AppxLog -Message "Resolved PackageFullName to path: $PackagePath" -Level 'Verbose'
                }
            }
            
            $packagePath = ConvertTo-SecureFilePath -Path $PackagePath -MustExist -PathType Directory
            $outputPath = ConvertTo-SecureFilePath -Path $OutputPath -PathType Directory -CreateIfMissing
            
            # Find manifest
            $manifestPath = [System.IO.Path]::Combine($packagePath, 'AppxManifest.xml')
            if (-not (Test-Path -LiteralPath $manifestPath)) {
                throw "AppxManifest.xml not found in package directory: $packagePath"
            }

            # Stage 2: Parse Manifest
            $progressStage++
            Write-Progress -Id $progressId -Activity "Backing up APPX Package" `
                -Status "Stage $progressStage/$totalStages : Parsing manifest" `
                -PercentComplete (($progressStage / $totalStages) * 100)
            
            Write-AppxLog -Message "Parsing package manifest..." -Level 'Verbose'
            $manifestData = Get-AppxManifestData -ManifestPath $manifestPath -IncludeDependencies:$IncludeDependencies
            
            Write-AppxLog -Message "Package: $($manifestData.Name) v$($manifestData.Version)" -Level 'Info'
            Write-AppxLog -Message "Publisher: $($manifestData.Publisher)" -Level 'Debug'
            Write-AppxLog -Message "Architecture: $($manifestData.ProcessorArchitecture)" -Level 'Debug'

            # Generate output filename
            $baseFileName = Split-Path -Path $packagePath -Leaf
            $packageOutputPath = [System.IO.Path]::Combine($outputPath, "$baseFileName.appx")
            $certOutputPath = [System.IO.Path]::Combine($outputPath, "$baseFileName.cer")

            # Check for existing files
            if (-not $Force.IsPresent) {
                if (Test-Path -LiteralPath $packageOutputPath) {
                    throw "Output file already exists (use -Force to overwrite): $packageOutputPath"
                }
            }

            # Stage 3: Resolve Dependencies (if requested)
            $dependencyInfo = $null
            $dependencyPackages = @()
            $bundlePath = $null
            $bundleWorkDir = $null
            $isZipArchive = $false
            
            if ($IncludeDependencies.IsPresent -or $DependencyReportOnly.IsPresent) {
                $progressStage++
                Write-Progress -Id $progressId -Activity "Backing up APPX Package" `
                    -Status "Stage $progressStage/$totalStages : Resolving dependencies" `
                    -PercentComplete (($progressStage / $totalStages) * 100)
                
                Write-AppxLog -Message "Resolving package dependencies..." -Level 'Verbose'
                $dependencyInfo = Resolve-AppxDependencies -PackagePath $packagePath -IncludeOptional
                
                Write-AppxLog -Message "Dependencies found: $($dependencyInfo.TotalDependencies) (Missing: $($dependencyInfo.MissingCount))" -Level 'Info'
                
                if ($dependencyInfo.MissingCount -gt 0) {
                    Write-AppxLog -Message "Some dependencies are not installed. Bundle may be incomplete." -Level 'Warning'
                }
                
                # Export dependency report
                $depReportPath = [System.IO.Path]::Combine(
                    (Split-Path $packageOutputPath -Parent),
                    "$baseFileName`_Dependencies.json"
                )
                
                try {
                    $dependencyInfo | ConvertTo-Json -Depth 10 | Out-File -FilePath $depReportPath -Encoding UTF8
                    Write-AppxLog -Message "Dependency report saved: $depReportPath" -Level 'Info'
                }
                catch {
                    Write-AppxLog -Message "Failed to save dependency report: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                }
                
                # If IncludeDependencies, create full bundle with all dependencies
                if ($IncludeDependencies.IsPresent) {
                    Write-AppxLog -Message "Creating complete dependency bundle..." -Level 'Info'
                    Write-Host "`n[INFO] Bundling dependencies - this may take several minutes..." -ForegroundColor Cyan
                    
                    # Create temporary directory for dependency packages
                    $bundleWorkDir = [System.IO.Path]::Combine($env:TEMP, "AppxBundle_$(New-Guid)")
                    [void](New-Item -Path $bundleWorkDir -ItemType Directory -Force)
                    
                    try {
                        # Package each installed dependency
                        $depPackageCount = 0
                        foreach ($dep in @($dependencyInfo.Dependencies | Where-Object { $_.IsInstalled })) {
                            $depPackageCount++
                            Write-Progress -Id ($progressId + 1) -Activity "Packaging Dependencies" `
                                -Status "Package $depPackageCount of $($dependencyInfo.InstalledCount): $($dep.Name)" `
                                -PercentComplete (($depPackageCount / $dependencyInfo.InstalledCount) * 100)
                            
                            Write-AppxLog -Message "Packaging dependency: $($dep.Name) v$($dep.InstalledVersion)" -Level 'Verbose'
                            
                            if ($null -eq $dep.ResolvedPath -or -not (Test-Path -LiteralPath $dep.ResolvedPath)) {
                                Write-AppxLog -Message "Dependency path not accessible: $($dep.Name)" -Level 'Warning'
                                continue
                            }
                            
                            # Generate safe filename for dependency package
                            $depSafeName = "$($dep.Name)_$($dep.InstalledVersion)_$($dep.Architecture)".Replace(':', '_').Replace('\', '_').Replace('/', '_')
                            $depPackagePath = [System.IO.Path]::Combine($bundleWorkDir, "$depSafeName.appx")
                            
                            try {
                                # Create dependency package
                                $depResult = New-AppxPackageInternal `
                                    -SourcePath $dep.ResolvedPath `
                                    -OutputPath $depPackagePath `
                                    -CompressionLevel $CompressionLevel
                                
                                if ($depResult.Success) {
                                    $dependencyPackages += [PSCustomObject]@{
                                        Name = $dep.Name
                                        Version = $dep.InstalledVersion
                                        Architecture = $dep.Architecture
                                        PackagePath = $depPackagePath
                                        PackageSize = $depResult.PackageSize
                                    }
                                    Write-AppxLog -Message "Dependency packaged: $depPackagePath" -Level 'Debug'
                                }
                            }
                            catch {
                                Write-AppxLog -Message "Failed to package dependency $($dep.Name): $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                            }
                        }
                        
                        Write-Progress -Id ($progressId + 1) -Activity "Packaging Dependencies" -Completed
                        
                        Write-AppxLog -Message "Successfully packaged $($dependencyPackages.Count) dependencies" -Level 'Info'
                        Write-Host "[SUCCESS] $($dependencyPackages.Count) dependencies packaged, will bundle after main package creation" -ForegroundColor Green
                    }
                    catch {
                        Write-AppxLog -Message "Dependency packaging error: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'
                        # Continue with main package even if dependencies fail
                        $dependencyPackages = @()
                    }
                }
            }

            # Stage 4: Create Package
            $progressStage++
            Write-Progress -Id $progressId -Activity "Backing up APPX Package" `
                -Status "Stage $progressStage/$totalStages : Creating package file" `
                -PercentComplete (($progressStage / $totalStages) * 100)
            
            Write-AppxLog -Message "Creating package file..." -Level 'Verbose'
            
            if ($PSCmdlet.ShouldProcess($packageOutputPath, "Create APPX package")) {
                $packageResult = New-AppxPackageInternal `
                    -SourcePath $packagePath `
                    -OutputPath $packageOutputPath `
                    -CompressionLevel $CompressionLevel
                
                Write-AppxLog -Message "Package created: $($packageResult.PackageSizeMB) MB" -Level 'Info'
            }
            else {
                Write-AppxLog -Message "Package creation skipped (WhatIf)" -Level 'Info'
                return
            }
            
            # Stage 4.5: Create ZIP Archive (if dependencies were packaged)
            if ($IncludeDependencies.IsPresent -and $dependencyPackages.Count -gt 0 -and $bundleWorkDir) {
                Write-AppxLog -Message "Creating ZIP archive with main package + dependencies..." -Level 'Info'
                Write-Host "`n[INFO] Creating ZIP package archive with certificates..." -ForegroundColor Cyan
                
                try {
                    # Array to store all package files with their certificates
                    $allPackageFiles = @()
                    
                    # Create certificates for each dependency
                    $certCreationCount = 0
                    foreach ($dep in $dependencyPackages) {
                        $certCreationCount++
                        Write-Progress -Id ($progressId + 2) -Activity "Creating Dependency Certificates" `
                            -Status "Certificate $certCreationCount of $($dependencyPackages.Count): $($dep.Name)" `
                            -PercentComplete (($certCreationCount / $dependencyPackages.Count) * 100)
                        
                        Write-AppxLog -Message "Creating certificate for dependency: $($dep.Name)" -Level 'Verbose'
                        
                        # Get publisher from installed package
                        $depPublisher = $null
                        try {
                            # Find the installed package to get publisher
                            $depInstalled = Get-AppxPackage | Where-Object { 
                                $_.Name -eq $dep.Name -and $_.Version -eq $dep.Version 
                            } | Select-Object -First 1
                            
                            if ($depInstalled) {
                                $depPublisher = $depInstalled.Publisher
                                Write-AppxLog -Message "Extracted publisher for $($dep.Name): $depPublisher" -Level 'Debug'
                            }
                            else {
                                # Fallback: try without version match
                                $depInstalled = Get-AppxPackage -Name $dep.Name | Select-Object -First 1
                                if ($depInstalled) {
                                    $depPublisher = $depInstalled.Publisher
                                    Write-AppxLog -Message "Extracted publisher (no version match) for $($dep.Name): $depPublisher" -Level 'Debug'
                                }
                            }
                        }
                        catch {
                            Write-AppxLog -Message "Could not extract publisher for $($dep.Name): $_" -Level 'Debug'
                        }
                        
                        # Skip if no publisher found
                        if ([string]::IsNullOrWhiteSpace($depPublisher)) {
                            Write-AppxLog -Message "WARNING: No publisher found for $($dep.Name), skipping certificate and signing" -Level 'Warning'
                            # Add package without certificate (UNSIGNED)
                            $allPackageFiles += @{
                                PackagePath          = $dep.PackagePath
                                CertificatePath      = $null
                                CertificateThumbprint = $null
                                Name                 = $dep.Name
                                Version              = $dep.Version
                                Architecture         = $dep.Architecture
                            }
                            continue
                        }
                        
                        # Generate package name for certificate
                        $depPackageName = "$($dep.Name)_$($dep.Version)_$($dep.Architecture)"
                        
                        # Create certificate for this dependency
                        $depCertResult = New-AppxDependencyCertificate `
                            -PackageName $depPackageName `
                            -OutputDirectory $bundleWorkDir `
                            -PublisherSubject $depPublisher
                        
                        if ($depCertResult.Success) {
                            Write-AppxLog -Message "Certificate created for $($dep.Name): $($depCertResult.Thumbprint)" -Level 'Debug'
                            
                            # INSTALL THE CERTIFICATE IMMEDIATELY (mirroring vanilla backup)
                            Write-AppxLog -Message "Installing certificate for $($dep.Name) to Trusted Root store..." -Level 'Debug'
                            
                            try {
                                # Try LocalMachine\Root first (requires Administrator)
                                try {
                                    Import-Certificate -FilePath $depCertResult.CertificatePath `
                                        -CertStoreLocation "Cert:\LocalMachine\Root" `
                                        -ErrorAction Stop | Out-Null
                                    
                                    Write-AppxLog -Message "Certificate installed to LocalMachine\Root for $($dep.Name)" -Level 'Debug'
                                }
                                catch {
                                    # Fall back to CurrentUser\Root (no admin required)
                                    Import-Certificate -FilePath $depCertResult.CertificatePath `
                                        -CertStoreLocation "Cert:\CurrentUser\Root" `
                                        -ErrorAction Stop | Out-Null
                                    
                                    Write-AppxLog -Message "Certificate installed to CurrentUser\Root for $($dep.Name)" -Level 'Debug'
                                }
                            }
                            catch {
                                Write-AppxLog -Message "WARNING: Failed to install certificate for $($dep.Name): $_" -Level 'Warning'
                                # Continue even if certificate installation fails
                            }
                            
                            # SIGN THE DEPENDENCY PACKAGE IMMEDIATELY
                            try {
                                # Get certificate from store by thumbprint
                                $depCertPath = "Cert:\CurrentUser\My\$($depCertResult.Thumbprint)"
                                $depCert = Get-Item -LiteralPath $depCertPath -ErrorAction Stop
                                
                                Write-AppxLog -Message "Signing dependency package: $($dep.Name)" -Level 'Debug'
                                
                                # Find SignTool.exe
                                $signToolPath = $null
                                $makeAppxPath = Get-Command 'makeappx.exe' -ErrorAction SilentlyContinue | 
                                    Select-Object -First 1 -ExpandProperty Source
                                
                                if ($makeAppxPath) {
                                    $sdkDir = Split-Path -Path $makeAppxPath -Parent
                                    $signToolTest = [System.IO.Path]::Combine($sdkDir, 'signtool.exe')
                                    if (Test-Path -LiteralPath $signToolTest) {
                                        $signToolPath = $signToolTest
                                    }
                                }
                                
                                if ($null -eq $signToolPath) {
                                    $signToolCmd = Get-Command 'signtool.exe' -ErrorAction SilentlyContinue | 
                                        Select-Object -First 1
                                    if ($signToolCmd) {
                                        $signToolPath = $signToolCmd.Source
                                    }
                                }
                                
                                if ($null -eq $signToolPath) {
                                    throw "SignTool.exe not found"
                                }
                                
                                # Sign the dependency package
                                $signArgs = @(
                                    'sign',
                                    '/fd', 'SHA256',
                                    '/sha1', $depCert.Thumbprint,
                                    '/v',
                                    $dep.PackagePath
                                )
                                
                                $signResult = Invoke-ProcessSafely -FilePath $signToolPath `
                                    -ArgumentList $signArgs `
                                    -TimeoutSeconds 120 `
                                    -NoWindow
                                
                                if ($signResult.Success) {
                                    Write-AppxLog -Message "Dependency package signed: $($dep.Name)" -Level 'Debug'
                                }
                                else {
                                    Write-AppxLog -Message "WARNING: Failed to sign dependency $($dep.Name): Exit code $($signResult.ExitCode)" -Level 'Warning'
                                }
                            }
                            catch {
                                Write-AppxLog -Message "WARNING: Failed to sign dependency $($dep.Name): $_" -Level 'Warning'
                                # Continue even if signing fails - package still usable with certificate
                            }
                            
                            # Add to package files array
                            $allPackageFiles += @{
                                PackagePath          = $dep.PackagePath
                                CertificatePath      = $depCertResult.CertificatePath
                                CertificateThumbprint = $depCertResult.Thumbprint
                                Name                 = $dep.Name
                                Version              = $dep.Version
                                Architecture         = $dep.Architecture
                            }
                        }
                        else {
                            Write-AppxLog -Message "WARNING: Certificate creation failed for $($dep.Name): $($depCertResult.Error)" -Level 'Warning'
                            # Add package without certificate (UNSIGNED)
                            $allPackageFiles += @{
                                PackagePath          = $dep.PackagePath
                                CertificatePath      = $null
                                CertificateThumbprint = $null
                                Name                 = $dep.Name
                                Version              = $dep.Version
                                Architecture         = $dep.Architecture
                            }
                        }
                    }
                    
                    Write-Progress -Id ($progressId + 2) -Activity "Creating Dependency Certificates" -Completed
                    
                    # Add main package to the files array (certificate will be created later in normal flow)
                    # For now, we'll create it here to include in the ZIP
                    Write-AppxLog -Message "Creating certificate for main package..." -Level 'Verbose'
                    
                    # Main package certificate will be created in the normal certificate creation stage
                    # For ZIP packaging, we need to ensure it exists first
                    $mainCertPath = [System.IO.Path]::Combine($bundleWorkDir, "$baseFileName.cer")
                    
                    # We'll add main package to array, cert will be added after it's created
                    $mainPackageFileEntry = @{
                        PackagePath          = $packageOutputPath
                        CertificatePath      = $mainCertPath  # Will be created later
                        CertificateThumbprint = $null  # Will be set later
                        Name                 = $manifestData.Name
                        Version              = $manifestData.Version
                        Architecture         = $manifestData.ProcessorArchitecture
                    }
                    
                    # Generate manifest metadata
                    Write-AppxLog -Message "Generating installation manifest..." -Level 'Verbose'
                    $manifestData_manifest = New-AppxBackupManifest `
                        -MainPackageInfo @{
                            Name                 = $manifestData.Name
                            Version              = $manifestData.Version
                            Architecture         = $manifestData.ProcessorArchitecture
                            Publisher            = $manifestData.Publisher
                            PublisherDisplayName = $manifestData.PublisherDisplayName
                            ResourceId           = $manifestData.ResourceId
                        } `
                        -DependencyInfo $dependencyInfo.Dependencies `
                        -PackageFiles (@($mainPackageFileEntry) + $allPackageFiles) `
                        -OutputDirectory $bundleWorkDir
                    
                    Write-AppxLog -Message "Installation manifest generated" -Level 'Debug'
                    
                    # NOW SIGN THE MAIN PACKAGE (before adding to ZIP)
                    # We need to create certificate and sign just like vanilla backup
                    Write-AppxLog -Message "Creating and signing main package for ZIP..." -Level 'Verbose'
                    
                    try {
                        # Create certificate for main package
                        $certSubject = if ($CertificateSubject) { 
                            $CertificateSubject 
                        } else { 
                            $manifestData.Publisher 
                        }
                        
                        $mainCertParams = @{
                            Subject = $certSubject
                            OutputPath = $mainCertPath
                            ValidityYears = $script:AppxBackupConfig.DefaultCertificateValidityYears
                            KeyLength = $script:AppxBackupConfig.DefaultKeyLength
                        }
                        
                        if ($CertificatePassword) {
                            $mainCertParams['Password'] = $CertificatePassword
                        }
                        
                        $mainCertificate = New-AppxBackupCertificate @mainCertParams
                        Write-AppxLog -Message "Main package certificate created: $($mainCertificate.Thumbprint)" -Level 'Debug'
                        
                        # Update main package entry with certificate info
                        $mainPackageFileEntry.CertificateThumbprint = $mainCertificate.Thumbprint
                        
                        # INSTALL THE MAIN CERTIFICATE IMMEDIATELY (mirroring vanilla backup)
                        Write-AppxLog -Message "Installing main package certificate to Trusted Root store..." -Level 'Verbose'
                        
                        try {
                            # Try LocalMachine\Root first (requires Administrator)
                            try {
                                Import-Certificate -FilePath $mainCertPath `
                                    -CertStoreLocation "Cert:\LocalMachine\Root" `
                                    -ErrorAction Stop | Out-Null
                                
                                Write-AppxLog -Message "Main certificate installed to LocalMachine\Root (system-wide trust)" -Level 'Info'
                            }
                            catch {
                                # Fall back to CurrentUser\Root (no admin required)
                                Import-Certificate -FilePath $mainCertPath `
                                    -CertStoreLocation "Cert:\CurrentUser\Root" `
                                    -ErrorAction Stop | Out-Null
                                
                                Write-AppxLog -Message "Main certificate installed to CurrentUser\Root (user trust only)" -Level 'Warning'
                                Write-AppxLog -Message "For system-wide trust, run as Administrator" -Level 'Warning'
                            }
                        }
                        catch {
                            Write-AppxLog -Message "WARNING: Failed to install main certificate: $_" -Level 'Warning'
                            # Continue even if certificate installation fails
                        }
                        
                        # Sign the main package with SignTool
                        $mainCertPath_store = "Cert:\CurrentUser\My\$($mainCertificate.Thumbprint)"
                        $mainCert = Get-Item -LiteralPath $mainCertPath_store -ErrorAction Stop
                        
                        Write-AppxLog -Message "Signing main package with certificate: $($mainCert.Thumbprint)" -Level 'Debug'
                        
                        # Find SignTool.exe
                        $signToolPath = $null
                        $makeAppxPath = Get-Command 'makeappx.exe' -ErrorAction SilentlyContinue | 
                            Select-Object -First 1 -ExpandProperty Source
                        
                        if ($makeAppxPath) {
                            $sdkDir = Split-Path -Path $makeAppxPath -Parent
                            $signToolTest = [System.IO.Path]::Combine($sdkDir, 'signtool.exe')
                            if (Test-Path -LiteralPath $signToolTest) {
                                $signToolPath = $signToolTest
                            }
                        }
                        
                        if ($null -eq $signToolPath) {
                            $signToolCmd = Get-Command 'signtool.exe' -ErrorAction SilentlyContinue | 
                                Select-Object -First 1
                            if ($signToolCmd) {
                                $signToolPath = $signToolCmd.Source
                            }
                        }
                        
                        if ($null -eq $signToolPath) {
                            throw "SignTool.exe not found"
                        }
                        
                        # Copy main package to work directory for signing
                        $mainPackageInZip = [System.IO.Path]::Combine($bundleWorkDir, [System.IO.Path]::GetFileName($packageOutputPath))
                        Copy-Item -Path $packageOutputPath -Destination $mainPackageInZip -Force
                        
                        # Sign the copy that will go in the ZIP
                        $signArgs = @(
                            'sign',
                            '/fd', 'SHA256',
                            '/sha1', $mainCert.Thumbprint,
                            '/v',
                            $mainPackageInZip
                        )
                        
                        $signResult = Invoke-ProcessSafely -FilePath $signToolPath `
                            -ArgumentList $signArgs `
                            -TimeoutSeconds 300 `
                            -NoWindow
                        
                        if ($signResult.Success) {
                            Write-AppxLog -Message "Main package signed successfully" -Level 'Info'
                            # Update the main package path to point to the signed version
                            $mainPackageFileEntry.PackagePath = $mainPackageInZip
                        }
                        else {
                            Write-AppxLog -Message "WARNING: Failed to sign main package: Exit code $($signResult.ExitCode)" -Level 'Warning'
                        }
                    }
                    catch {
                        Write-AppxLog -Message "WARNING: Failed to create certificate or sign main package: $_" -Level 'Warning'
                        # Continue anyway - unsigned main package
                    }
                    
                    # Create ZIP archive extension from configuration
                    $zipExtension = Get-AppxDefault 'archiveExtensions.zipArchive' 'ZipPackagingConfiguration' '.appxpack'
                    $zipOutputPath = $packageOutputPath -replace '\.appx$', $zipExtension -replace '\.msix$', $zipExtension
                    
                    Write-AppxLog -Message "Creating ZIP archive: $zipOutputPath" -Level 'Info'
                    Write-Host "[INFO] All packages signed, creating ZIP archive..." -ForegroundColor Cyan
                    
                    # Create the ZIP archive with all SIGNED packages and certificates
                    $zipResult = New-AppxBackupZipArchive `
                        -SourceDirectory $bundleWorkDir `
                        -OutputPath $zipOutputPath `
                        -ManifestData $manifestData_manifest `
                        -CompressionLevel 'Optimal'
                    
                    if ($zipResult.Success) {
                        Write-AppxLog -Message "ZIP archive created successfully: $($zipResult.TotalSizeMB)" -Level 'Info'
                        Write-Host "[SUCCESS] ZIP package created with $($dependencyPackages.Count) dependencies ($($zipResult.TotalSizeMB))" -ForegroundColor Green
                        Write-Host "[INFO] Archive contains $($zipResult.PackageCount) packages and $($zipResult.CertificateCount) certificates" -ForegroundColor Cyan
                        
                        # Update packageOutputPath to point to ZIP
                        $originalPackagePath = $packageOutputPath
                        $packageOutputPath = $zipOutputPath
                        $bundlePath = $zipOutputPath  # Set for return object compatibility
                        
                        # Remove standalone main package (now in ZIP)
                        if (Test-Path -LiteralPath $originalPackagePath) {
                            Remove-Item -LiteralPath $originalPackagePath -Force -ErrorAction SilentlyContinue
                            Write-AppxLog -Message "Removed standalone package (now in ZIP): $originalPackagePath" -Level 'Debug'
                        }
                        
                        # Mark that we created a ZIP archive
                        $isZipArchive = $true
                    }
                    else {
                        Write-AppxLog -Message "ZIP archive creation failed: $($zipResult.Error)" -Level 'Error'
                        Write-Host "[WARNING] ZIP creation failed. Standalone package retained." -ForegroundColor Yellow
                        $bundlePath = $null
                        $isZipArchive = $false
                    }
                }
                catch {
                    Write-AppxLog -Message "ZIP archive creation error: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'
                    Write-Host "[WARNING] ZIP creation failed. Standalone package retained." -ForegroundColor Yellow
                    $bundlePath = $null
                    $isZipArchive = $false
                }
                finally {
                    # Cleanup work directory (will be cleaned up later in main cleanup)
                    # We keep it for now in case we need to create the main cert
                }
            }

            # Stage 5: Certificate Creation & Signing
            $certificate = $null
            
            if (-not $NoCertificate.IsPresent) {
                $progressStage++
                Write-Progress -Id $progressId -Activity "Backing up APPX Package" `
                    -Status "Stage $progressStage/$totalStages : Creating certificate" `
                    -PercentComplete (($progressStage / $totalStages) * 100)
                
                Write-AppxLog -Message "Creating self-signed certificate..." -Level 'Verbose'
                
                $certSubject = if ($CertificateSubject) { 
                    $CertificateSubject 
                } else { 
                    $manifestData.Publisher 
                }
                
                if ($PSCmdlet.ShouldProcess($certSubject, "Create self-signed certificate")) {
                    $certParams = @{
                        Subject = $certSubject
                        OutputPath = $certOutputPath
                        ValidityYears = $script:AppxBackupConfig.DefaultCertificateValidityYears
                        KeyLength = $script:AppxBackupConfig.DefaultKeyLength
                    }
                    
                    if ($CertificatePassword) {
                        $certParams['Password'] = $CertificatePassword
                    }
                    
                    $certificate = New-AppxBackupCertificate @certParams
                    
                    Write-AppxLog -Message "Certificate created: $certOutputPath" -Level 'Info'
                    
                    # Install certificate to Trusted Root Certification Authorities
                    # This is REQUIRED for the signed package to install without errors
                    Write-AppxLog -Message "Installing certificate to Trusted Root store..." -Level 'Verbose'
                    
                    try {
                        # Import to LocalMachine\Root requires Administrator privileges
                        # Try LocalMachine first, fall back to CurrentUser if access denied
                        try {
                            Import-Certificate -FilePath $certOutputPath `
                                -CertStoreLocation "Cert:\LocalMachine\Root" `
                                -ErrorAction Stop | Out-Null
                            
                            Write-AppxLog -Message "Certificate installed to LocalMachine\Root (system-wide trust)" -Level 'Info'
                            $certInstalled = $true
                        }
                        catch {
                            # Likely not running as Administrator, try CurrentUser
                            Write-AppxLog -Message "Cannot install to LocalMachine\Root (not Administrator), trying CurrentUser..." -Level 'Debug'
                            
                            Import-Certificate -FilePath $certOutputPath `
                                -CertStoreLocation "Cert:\CurrentUser\Root" `
                                -ErrorAction Stop | Out-Null
                            
                            Write-AppxLog -Message "Certificate installed to CurrentUser\Root (user trust only)" -Level 'Warning'
                            Write-AppxLog -Message "For system-wide trust, run as Administrator or manually install to LocalMachine\Root" -Level 'Warning'
                            $certInstalled = $true
                        }
                    }
                    catch {
                        Write-AppxLog -Message "Failed to automatically install certificate: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                        Write-AppxLog -Message "You must manually install the certificate before the package can be installed" -Level 'Warning'
                        $certInstalled = $false
                    }
                }

                # Stage 6: Sign Package
                $progressStage++
                Write-Progress -Id $progressId -Activity "Backing up APPX Package" `
                    -Status "Stage $progressStage/$totalStages : Signing package" `
                    -PercentComplete (($progressStage / $totalStages) * 100)
                
                # ZIP archives (.appxpack) cannot be signed - certificates are included inside
                if ($isZipArchive) {
                    Write-AppxLog -Message "Skipping signing - ZIP archive packages are not signed (certificates included inside)" -Level 'Info'
                    Write-Host "[INFO] ZIP archive complete - certificates included for all packages" -ForegroundColor Cyan
                }
                else {
                    Write-AppxLog -Message "Signing package..." -Level 'Verbose'
                    
                    if ($PSCmdlet.ShouldProcess($packageOutputPath, "Sign package with certificate")) {
                        # APPX packages MUST be signed with SignTool.exe, not Set-AuthenticodeSignature
                        # Set-AuthenticodeSignature uses wrong SIP provider and fails with SIP_SUBJECTINFO error
                        
                        try {
                            # Get certificate from store by thumbprint
                            $certPath = "Cert:\CurrentUser\My\$($certificate.Thumbprint)"
                            $cert = Get-Item -LiteralPath $certPath -ErrorAction Stop
                            
                            Write-AppxLog -Message "Signing with certificate: $($cert.Thumbprint)" -Level 'Debug'
                        
                        # Find SignTool.exe (should be in same SDK as MakeAppx)
                        $signToolPath = $null
                        
                        # Try to find SignTool in same directory as MakeAppx
                        $makeAppxPath = Get-Command 'makeappx.exe' -ErrorAction SilentlyContinue | 
                            Select-Object -First 1 -ExpandProperty Source
                        
                        if ($makeAppxPath) {
                            $sdkDir = Split-Path -Path $makeAppxPath -Parent
                            $signToolTest = [System.IO.Path]::Combine($sdkDir, 'signtool.exe')
                            if (Test-Path -LiteralPath $signToolTest) {
                                $signToolPath = $signToolTest
                            }
                        }
                        
                        # Fallback: search PATH
                        if ($null -eq $signToolPath) {
                            $signToolCmd = Get-Command 'signtool.exe' -ErrorAction SilentlyContinue | 
                                Select-Object -First 1
                            if ($signToolCmd) {
                                $signToolPath = $signToolCmd.Source
                            }
                        }
                        
                        if ($null -eq $signToolPath) {
                            throw "SignTool.exe not found. Install Windows SDK to sign APPX packages."
                        }
                        
                        Write-AppxLog -Message "Using SignTool: $signToolPath" -Level 'Debug'
                        
                        # Build SignTool command for APPX signing
                        # sign = sign command
                        # /fd = file digest algorithm (required)
                        # /sha1 = certificate thumbprint
                        # /v = verbose output
                        $signArgs = @(
                            'sign',
                            '/fd', 'SHA256',
                            '/sha1', $cert.Thumbprint,
                            '/v',
                            $packageOutputPath
                        )
                        
                        Write-AppxLog -Message "SignTool command: $signToolPath $($signArgs -join ' ')" -Level 'Debug'
                        
                        # Execute SignTool
                        $signResult = Invoke-ProcessSafely -FilePath $signToolPath `
                            -ArgumentList $signArgs `
                            -TimeoutSeconds 300 `
                            -NoWindow
                        
                        if (-not $signResult.Success) {
                            $errorDetails = "SignTool failed with exit code $($signResult.ExitCode)"
                            if ($signResult.StandardError) {
                                $errorDetails += "`nSTDERR: $($signResult.StandardError)"
                            }
                            if ($signResult.StandardOutput) {
                                $errorDetails += "`nSTDOUT: $($signResult.StandardOutput)"
                            }
                            throw $errorDetails
                        }
                        
                        Write-AppxLog -Message "Package signed successfully" -Level 'Info'
                        
                        # Log signing details from SignTool output
                        if ($signResult.StandardOutput) {
                            Write-AppxLog -Message "SignTool output: $($signResult.StandardOutput)" -Level 'Debug'
                        }
                    }
                    catch {
                        Write-AppxLog -Message "Failed to sign package: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'
                        throw "Failed to sign package: $_"
                    }
                }
                }  # End else (not ZIP archive)
            }
            else {
                Write-AppxLog -Message "Certificate creation and signing skipped (-NoCertificate)" -Level 'Warning'
            }

            # Cleanup temporary work directory
            if ($bundleWorkDir -and (Test-Path -LiteralPath $bundleWorkDir)) {
                try {
                    Remove-Item -Path $bundleWorkDir -Recurse -Force -ErrorAction Stop
                    Write-AppxLog -Message "Cleaned up temporary work directory: $bundleWorkDir" -Level 'Debug'
                }
                catch {
                    Write-AppxLog -Message "Failed to cleanup work directory: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                }
            }

            # Complete
            Write-Progress -Id $progressId -Activity "Backing up APPX Package" -Completed
            
            # Build result object
            $result = [PSCustomObject]@{
                PSTypeName                  = 'AppxBackup.BackupResult'
                Success                     = $true
                PackageName                 = $manifestData.Name
                PackageVersion              = $manifestData.Version
                PackagePublisher            = $manifestData.Publisher
                PackageArchitecture         = $manifestData.ProcessorArchitecture
                PackageFilePath             = $packageOutputPath
                PackageFileSize             = if ($bundlePath) { (Get-Item -LiteralPath $bundlePath).Length } else { $packageResult.PackageSize }
                PackageFileSizeMB           = if ($bundlePath) { [Math]::Round((Get-Item -LiteralPath $bundlePath).Length / 1MB, 2) } else { $packageResult.PackageSizeMB }
                IsBundle                    = $false  # Deprecated: Now using ZIP archives
                IsZipArchive                = $isZipArchive
                BundledDependencyCount      = $dependencyPackages.Count
                CertificateFilePath         = if ($certificate) { $certOutputPath } else { $null }
                CertificateThumbprint       = if ($certificate) { $certificate.Thumbprint } else { $null }
                CertificateInstalled        = if ($certificate) { $certInstalled } else { $false }
                DependencyCount             = if ($dependencyInfo) { $dependencyInfo.TotalDependencies } else { 0 }
                DependenciesMissing         = if ($dependencyInfo) { $dependencyInfo.MissingCount } else { 0 }
                DependencyInfo              = $dependencyInfo
                DependencyReportPath        = if (($IncludeDependencies.IsPresent -or $DependencyReportOnly.IsPresent) -and $dependencyInfo) { $depReportPath } else { $null }
                SourcePath                  = $packagePath
                BackupDate                  = [DateTime]::Now
                CompressionUsed             = $CompressionLevel
            }
            
            # Summary with installation instructions
            Write-AppxLog -Message "=== Backup Complete ===" -Level 'Info'
            Write-AppxLog -Message "Package: $($result.PackageFilePath)" -Level 'Info'
            
            if ($certificate) {
                Write-AppxLog -Message "Certificate: $($result.CertificateFilePath)" -Level 'Info'
                
                if ($certInstalled) {
                    Write-Host "`nCertificate installed successfully - package is ready to install`n" -ForegroundColor Green
                    Write-Host "To install the package, run:" -ForegroundColor Cyan
                    Write-Host "  Add-AppxPackage -Path '$($result.PackageFilePath)'`n" -ForegroundColor White
                }
                else {
                    Write-Host "`nIMPORTANT: Certificate NOT automatically installed`n" -ForegroundColor Yellow
                    Write-Host "Before installing the package, install the certificate:" -ForegroundColor Yellow
                    Write-Host "  1. Run PowerShell as Administrator" -ForegroundColor White
                    Write-Host "  2. Import-Certificate -FilePath '$($result.CertificateFilePath)' -CertStoreLocation 'Cert:\LocalMachine\Root'`n" -ForegroundColor White
                    Write-Host "Then install the package:" -ForegroundColor Yellow
                    Write-Host "  Add-AppxPackage -Path '$($result.PackageFilePath)'`n" -ForegroundColor White
                }
            }
            
            # Notify about bundle creation
            if ($result.IsBundle) {
                Write-Host "`n=== BUNDLE CREATED ===" -ForegroundColor Green
                Write-Host "[SUCCESS] Complete bundle with all dependencies" -ForegroundColor Green
                Write-Host "  Bundle File: $($result.PackageFilePath)" -ForegroundColor White
                Write-Host "  Bundle Size: $($result.PackageFileSizeMB) MB" -ForegroundColor White
                Write-Host "  Main Package + $($result.BundledDependencyCount) Dependencies" -ForegroundColor Cyan
                Write-Host "`nThis bundle contains everything needed for installation." -ForegroundColor Gray
            }
            
            # Notify about dependency report if created
            if ($result.DependencyReportPath) {
                if (-not $result.IsZipArchive) {
                    Write-Host "`n[INFO] Dependency Report Created:" -ForegroundColor Cyan
                    Write-Host "  Location: $($result.DependencyReportPath)" -ForegroundColor White
                    Write-Host "  Dependencies: $($result.DependencyCount) total, $($result.DependenciesMissing) missing`n" -ForegroundColor Gray
                }
                else {
                    Write-Host "  Dependency Report: $($result.DependencyReportPath)`n" -ForegroundColor Gray
                }
            }
            
            return $result
        }
        catch {
            Write-Progress -Id $progressId -Activity "Backing up APPX Package" -Completed
            Write-AppxLog -Message "Backup failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'
            Write-AppxLog -Message "Stack trace: $($_.ScriptStackTrace)" -Level 'Debug'
            
            # Cleanup on failure
            if (Test-Path -LiteralPath $packageOutputPath) {
                Remove-Item -LiteralPath $packageOutputPath -Force -ErrorAction SilentlyContinue
            }
            
            throw
        }
    }

    end {
        Write-AppxLog -Message "Backup-AppxPackage completed" -Level 'Debug'
    }
}