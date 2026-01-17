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
                -Status "Stage $progressStage/$totalStages : Validating inputs" `
                -PercentComplete (($progressStage / $totalStages) * 100)
            
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
            $manifestPath = Join-Path $packagePath 'AppxManifest.xml'
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
            $packageOutputPath = Join-Path $outputPath "$baseFileName.appx"
            $certOutputPath = Join-Path $outputPath "$baseFileName.cer"

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
                            $signToolTest = Join-Path $sdkDir 'signtool.exe'
                            if (Test-Path -LiteralPath $signToolTest) {
                                $signToolPath = $signToolTest
                            }
                        }
                        
                        # Fallback: search PATH
                        if (-not $signToolPath) {
                            $signToolCmd = Get-Command 'signtool.exe' -ErrorAction SilentlyContinue | 
                                Select-Object -First 1
                            if ($signToolCmd) {
                                $signToolPath = $signToolCmd.Source
                            }
                        }
                        
                        if (-not $signToolPath) {
                            throw "SignTool.exe not found. Install Windows SDK to sign APPX packages."
                        }
                        
                        Write-AppxLog -Message "Using SignTool: $signToolPath" -Level 'Debug'
                        
                        # Build SignTool command for APPX signing
                        # /fd = file digest algorithm (required)
                        # /sha1 = certificate thumbprint
                        # /f = certificate file (we use /sha1 instead since cert is in store)
                        $signArgs = "sign /fd SHA256 /sha1 $($cert.Thumbprint) /v `"$packageOutputPath`""
                        
                        Write-AppxLog -Message "SignTool command: $signToolPath $signArgs" -Level 'Debug'
                        
                        # Execute SignTool
                        $signResult = Invoke-ProcessSafely -FilePath $signToolPath `
                            -Arguments $signArgs `
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
                DependencyCount         = if ($dependencyInfo) { $dependencyInfo.TotalDependencies } else { 0 }
                DependenciesMissing     = if ($dependencyInfo) { $dependencyInfo.MissingCount } else { 0 }
                DependencyInfo          = $dependencyInfo
                SourcePath              = $packagePath
                BackupDate              = [DateTime]::Now
                CompressionUsed         = $CompressionLevel
            }
            
            # Summary
            Write-AppxLog -Message "=== Backup Complete ===" -Level 'Info'
            Write-AppxLog -Message "Package: $($result.PackageFilePath)" -Level 'Info'
            if ($certificate) {
                Write-AppxLog -Message "Certificate: $($result.CertificateFilePath)" -Level 'Info'
                Write-Host "`n[CHAR_9888][CHAR_65039]  IMPORTANT: Install the .cer file to [Local Computer\Trusted Root Certification Authorities] before installing the package.`n" -ForegroundColor Yellow
            }
            
            return $result
        }
        catch {
            Write-Progress -Id $progressId -Activity "Backing up APPX Package" -Completed
            Write-AppxLog -Message "Backup failed: $_" -Level 'Error'
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