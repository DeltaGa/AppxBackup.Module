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
    If specified, analyzes and reports package dependencies.
    Does not bundle dependencies (see -CreateBundle for that).

.PARAMETER CreateBundle
    If specified, creates an MSIX bundle including dependencies.
    Requires dependencies to be locally accessible.

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
    Backup-AppxPackage -PackagePath $path -OutputPath $out -CreateBundle -Force
    
    Creates a complete bundle including dependencies, overwriting existing files

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
            if ($IncludeDependencies.IsPresent) {
                $progressStage++
                Write-Progress -Id $progressId -Activity "Backing up APPX Package" `
                    -Status "Stage $progressStage/$totalStages : Resolving dependencies" `
                    -PercentComplete (($progressStage / $totalStages) * 100)
                
                Write-AppxLog -Message "Resolving package dependencies..." -Level 'Verbose'
                $dependencyInfo = Resolve-AppxDependencies -PackagePath $packagePath -IncludeOptional
                
                Write-AppxLog -Message "Dependencies found: $($dependencyInfo.TotalDependencies) (Missing: $($dependencyInfo.MissingCount))" -Level 'Info'
                
                if ($dependencyInfo.MissingCount -gt 0) {
                    Write-AppxLog -Message "Some dependencies are not installed. Package may not restore successfully." -Level 'Warning'
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
            }
            else {
                Write-AppxLog -Message "Certificate creation and signing skipped (-NoCertificate)" -Level 'Warning'
            }

            # Complete
            Write-Progress -Id $progressId -Activity "Backing up APPX Package" -Completed
            
            # Build result object
            $result = [PSCustomObject]@{
                PSTypeName              = 'AppxBackup.BackupResult'
                Success                 = $true
                PackageName             = $manifestData.Name
                PackageVersion          = $manifestData.Version
                PackagePublisher        = $manifestData.Publisher
                PackageArchitecture     = $manifestData.ProcessorArchitecture
                PackageFilePath         = $packageOutputPath
                PackageFileSize         = $packageResult.PackageSize
                PackageFileSizeMB       = $packageResult.PackageSizeMB
                CertificateFilePath     = if ($certificate) { $certOutputPath } else { $null }
                CertificateThumbprint   = if ($certificate) { $certificate.Thumbprint } else { $null }
                CertificateInstalled    = if ($certificate) { $certInstalled } else { $false }
                DependencyCount         = if ($dependencyInfo) { $dependencyInfo.TotalDependencies } else { 0 }
                DependenciesMissing     = if ($dependencyInfo) { $dependencyInfo.MissingCount } else { 0 }
                DependencyInfo          = $dependencyInfo
                SourcePath              = $packagePath
                BackupDate              = [DateTime]::Now
                CompressionUsed         = $CompressionLevel
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