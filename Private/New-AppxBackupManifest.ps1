<#
.SYNOPSIS
    Generates AppxBackupManifest.json metadata for ZIP-based package archives.

.DESCRIPTION
    Internal helper function that creates the AppxBackupManifest.json file
    containing all metadata required for Install-AppxBackup to orchestrate
    proper package and certificate installation.
    
    The manifest includes:
    - Main package details (name, version, architecture, publisher)
    - Dependency information with installation order
    - Certificate thumbprints and file paths
    - Installation orchestration data
    - Archive metadata (creation date, version, total size)
    
    This metadata enables intelligent installation that:
    - Installs dependencies in the correct order
    - Matches certificates to packages
    - Validates architecture compatibility
    - Provides installation progress tracking

.PARAMETER MainPackageInfo
    Hashtable containing main package information from Get-AppxManifestData.

.PARAMETER DependencyInfo
    Array of hashtables containing dependency information from Resolve-AppxDependencies.

.PARAMETER PackageFiles
    Array of hashtables with PackagePath and CertificatePath for all packages.

.PARAMETER OutputDirectory
    Directory where packages and certificates are stored (for calculating relative paths).

.OUTPUTS
    [hashtable]
    Structured manifest data ready for ConvertTo-Json and ZIP inclusion.

.NOTES
    Author: DeltaGa
    Version: 2.0.0
    
    The manifest schema is versioned to support future format evolution.
    Version 1.0 is the initial implementation with ZIP-based packaging.

.EXAMPLE
    $manifest = New-AppxBackupManifest -MainPackageInfo $mainPkg -DependencyInfo $deps -PackageFiles $files
    $manifest | ConvertTo-Json -Depth 10 | Out-File "manifest.json"
#>

function New-AppxBackupManifest {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$MainPackageInfo,
        
        [Parameter()]
        [AllowEmptyCollection()]
        [array]$DependencyInfo = @(),
        
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [array]$PackageFiles,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$OutputDirectory
    )
    
    Write-AppxLog -Message "=== New-AppxBackupManifest ===" -Level 'Debug'
    Write-AppxLog -Message "Main Package: $($MainPackageInfo.Name) v$($MainPackageInfo.Version)" -Level 'Debug'
    Write-AppxLog -Message "Dependencies: $($DependencyInfo.Count)" -Level 'Debug'
    Write-AppxLog -Message "Package Files: $($PackageFiles.Count)" -Level 'Debug'
    
    try {
        # Calculate total archive size
        $totalSize = 0
        foreach ($file in $PackageFiles) {
            if (Test-Path -LiteralPath $file.PackagePath) {
                $totalSize += (Get-Item -LiteralPath $file.PackagePath).Length
            }
            if ($file.CertificatePath -and (Test-Path -LiteralPath $file.CertificatePath)) {
                $totalSize += (Get-Item -LiteralPath $file.CertificatePath).Length
            }
        }
        
        Write-AppxLog -Message "Total archive size: $([math]::Round($totalSize / 1MB, 2)) MB" -Level 'Debug'
        
        # Find main package in PackageFiles array
        $mainPackageFile = $PackageFiles | Where-Object { 
            [System.IO.Path]::GetFileNameWithoutExtension($_.PackagePath) -like "*$($MainPackageInfo.Name)*$($MainPackageInfo.Version)*"
        } | Select-Object -First 1
        
        if (-not $mainPackageFile) {
            Write-AppxLog -Message "WARNING: Could not locate main package file in PackageFiles array" -Level 'Warning'
            # Fallback to first package file
            $mainPackageFile = $PackageFiles[0]
        }
        
        # Build main package metadata
        $mainPackageMetadata = @{
            Name                  = $MainPackageInfo.Name
            Version               = $MainPackageInfo.Version
            Architecture          = $MainPackageInfo.Architecture
            Publisher             = $MainPackageInfo.Publisher
            PackageFile           = "Packages/$([System.IO.Path]::GetFileName($mainPackageFile.PackagePath))"
            CertificateFile       = if ($mainPackageFile.CertificatePath) { 
                "Certificates/$([System.IO.Path]::GetFileName($mainPackageFile.CertificatePath))" 
            } else { 
                $null 
            }
            CertificateThumbprint = $mainPackageFile.CertificateThumbprint
            PublisherDisplayName  = $MainPackageInfo.PublisherDisplayName
            ResourceId            = $MainPackageInfo.ResourceId
            IsBundle              = $false
            IsDevelopmentMode     = $false
        }
        
        Write-AppxLog -Message "Main package metadata created" -Level 'Debug'
        
        # Build dependencies metadata
        $dependenciesMetadata = @()
        $installationOrder = @()
        
        if ($DependencyInfo.Count -gt 0) {
            $installOrder = 1
            
            foreach ($dep in $DependencyInfo) {
                # Find corresponding package file
                $depPackageFile = $PackageFiles | Where-Object {
                    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($_.PackagePath)
                    $fileName -like "*$($dep.Name)*$($dep.Version)*"
                } | Select-Object -First 1
                
                if ($depPackageFile) {
                    $depMetadata = @{
                        Name                  = $dep.Name
                        Version               = $dep.Version
                        Architecture          = $dep.Architecture
                        MinVersion            = $dep.MinVersion
                        Publisher             = $dep.Publisher
                        PackageFile           = "Packages/$([System.IO.Path]::GetFileName($depPackageFile.PackagePath))"
                        CertificateFile       = if ($depPackageFile.CertificatePath) { 
                            "Certificates/$([System.IO.Path]::GetFileName($depPackageFile.CertificatePath))" 
                        } else { 
                            $null 
                        }
                        CertificateThumbprint = $depPackageFile.CertificateThumbprint
                        InstallOrder          = $installOrder
                        IsOptional            = $dep.IsOptional
                        DependencyType        = $dep.Type
                        IsInstalled           = $dep.IsInstalled
                    }
                    
                    $dependenciesMetadata += $depMetadata
                    
                    # Add to installation order (architecture-specific naming)
                    $depIdentifier = "$($dep.Name)_$($dep.Version)_$($dep.Architecture)"
                    $installationOrder += $depIdentifier
                    
                    $installOrder++
                    Write-AppxLog -Message "Added dependency: $($dep.Name) v$($dep.Version)" -Level 'Debug'
                }
                else {
                    Write-AppxLog -Message "WARNING: Could not locate package file for dependency: $($dep.Name)" -Level 'Warning'
                }
            }
        }
        
        # Add main package to end of installation order
        $mainIdentifier = "$($MainPackageInfo.Name)_$($MainPackageInfo.Version)_$($MainPackageInfo.Architecture)"
        $installationOrder += $mainIdentifier
        
        Write-AppxLog -Message "Installation order: $($installationOrder -join ', ')" -Level 'Debug'
        
        # Get manifest defaults from configuration
        $manifestVersion = Get-AppxDefault -Category 'manifestDefaults' -Key 'version' -ConfigName 'ZipPackagingConfiguration' -FallbackValue '1.0'
        $createdByString = Get-AppxDefault -Category 'manifestDefaults' -Key 'createdByString' -ConfigName 'ZipPackagingConfiguration' -FallbackValue 'AppxBackup v2.0'
        $defaultCompression = Get-AppxDefault -Category 'manifestDefaults' -Key 'defaultCompression' -ConfigName 'ZipPackagingConfiguration' -FallbackValue 'Optimal'
        $requiresElevation = Get-AppxDefault -Category 'manifestDefaults' -Key 'requiresElevation' -ConfigName 'ZipPackagingConfiguration' -FallbackValue $true
        $minOSVersion = Get-AppxDefault -Category 'systemRequirements' -Key 'minimumOSVersion' -ConfigName 'ZipPackagingConfiguration' -FallbackValue '10.0.17763'
        $minPSVersion = Get-AppxDefault -Category 'systemRequirements' -Key 'minimumPowerShellVersion' -ConfigName 'ZipPackagingConfiguration' -FallbackValue '5.1'
        
        # Build complete manifest structure
        $manifest = @{
            Version               = $manifestVersion
            CreatedBy             = $createdByString
            CreatedDate           = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
            MainPackage           = $mainPackageMetadata
            Dependencies          = $dependenciesMetadata
            InstallationOrder     = $installationOrder
            TotalPackages         = $PackageFiles.Count
            TotalSize             = $totalSize
            TotalSizeMB           = [math]::Round($totalSize / 1MB, 2)
            Compression           = $defaultCompression
            FormatVersion         = $manifestVersion
            RequiresElevation     = $requiresElevation
            MinimumOSVersion      = $minOSVersion
            MinimumPowerShellVersion = $minPSVersion
        }
        
        Write-AppxLog -Message "Manifest generation complete" -Level 'Info'
        Write-AppxLog -Message "Total packages: $($manifest.TotalPackages)" -Level 'Debug'
        Write-AppxLog -Message "Total dependencies: $($dependenciesMetadata.Count)" -Level 'Debug'
        
        return $manifest
    }
    catch {
        Write-AppxLog -Message "Manifest generation failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'
        throw
    }
}
