<#
.SYNOPSIS
    Gets comprehensive information about a backed-up APPX/MSIX package or .appxpack archive.

.DESCRIPTION
    Extracts and analyzes metadata from a package file or ZIP archive without installation.
    Provides detailed information about the package structure, manifest, signature, and compatibility.
    
    Supports:
    - Individual packages (.appx, .msix)
    - ZIP-based dependency archives (.appxpack)
    - Manifest extraction and parsing
    - Dependency enumeration
    - Certificate analysis
    - Archive structure validation

.PARAMETER PackagePath
    Path to the .appx, .msix, or .appxpack file to analyze.

.PARAMETER IncludeFileList
    If specified, includes a list of all files in the package.
    For .appxpack archives, lists all contained packages.

.PARAMETER IncludeSignatureInfo
    If specified, includes digital signature details.
    For .appxpack archives, includes signature info for all packages.

.PARAMETER IncludeManifestXml
    If specified, includes the raw manifest XML.
    For .appxpack archives, includes the main package manifest.

.EXAMPLE
    Get-AppxBackupInfo -PackagePath "C:\Backups\MyApp.appx"
    
    Gets basic package information

.EXAMPLE
    Get-AppxBackupInfo -PackagePath "C:\Backups\MyApp.appx" -IncludeFileList -IncludeSignatureInfo
    
    Gets comprehensive package information including files and signature

.EXAMPLE
    Get-AppxBackupInfo -PackagePath "C:\Backups\MyApp.appxpack" -IncludeSignatureInfo
    
    Analyzes a ZIP-based dependency archive with all packages and certificates

.OUTPUTS
    AppxBackup.PackageInfo
    For .appxpack archives, includes additional properties:
    - IsZipArchive: $true
    - ArchiveManifest: AppxBackupManifest.json contents
    - ContainedPackages: Array of all packages in archive
    - ContainedCertificates: Array of all certificates in archive
    - TotalArchiveSize: Combined size of all packages
    
.NOTES
    Does not require the package to be installed.
    Extracts manifest temporarily to analyze structure.
    For .appxpack files, extracts and analyzes AppxBackupManifest.json.
#>

function Get-AppxBackupInfo {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [Alias('FullName', 'Path')]
        [string]$PackagePath,

        [Parameter()]
        [switch]$IncludeFileList,

        [Parameter()]
        [switch]$IncludeSignatureInfo,

        [Parameter()]
        [switch]$IncludeManifestXml
    )

    begin {
        Write-AppxLog -Message "Analyzing package information" -Level 'Verbose'
        
        # Load compression assembly
        Add-Type -AssemblyName 'System.IO.Compression.FileSystem'
        
        # Load ZIP archive extension from configuration
        try {
            $zipConfig = Get-AppxConfiguration -ConfigName 'ZipPackagingConfiguration'
            $zipArchiveExtension = $zipConfig.archiveExtensions.zipArchive
        }
        catch {
            Write-AppxLog -Message "Failed to load ZIP configuration, using fallback: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
            $zipArchiveExtension = '.appxpack'
        }
    }

    process {
        try {
            # Validate package path
            $packagePath = ConvertTo-SecureFilePath -Path $PackagePath -MustExist -PathType File
            
            $packageFile = Get-Item -LiteralPath $packagePath
            
            Write-AppxLog -Message "Analyzing: $($packageFile.Name)" -Level 'Verbose'
            
            # Detect if this is a .appxpack ZIP archive
            $isZipArchive = $packageFile.Extension -eq $zipArchiveExtension
            
            if ($isZipArchive) {
                Write-AppxLog -Message "Detected .appxpack ZIP archive - using archive analysis path" -Level 'Debug'
                
                # Analyze as ZIP-based dependency archive
                $result = Get-AppxpackArchiveInfo -PackagePath $packagePath `
                    -PackageFile $packageFile `
                    -IncludeFileList:$IncludeFileList `
                    -IncludeSignatureInfo:$IncludeSignatureInfo `
                    -IncludeManifestXml:$IncludeManifestXml
                
                return $result
            }
            
            # Continue with standard package analysis for individual .appx/.msix files
            Write-AppxLog -Message "Detected individual package - using standard analysis path" -Level 'Debug'
            
            # Validate extension - load from configuration
            try {
                $pkgConfig = Get-AppxConfiguration -ConfigName 'PackageConfiguration'
                $validExtensions = $pkgConfig.packageExtensions.valid
            }
            catch {
                Write-AppxLog -Message "Failed to load package configuration, using fallback: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                $validExtensions = @('.appx', '.msix', '.appxbundle', '.msixbundle')
            }
            
            if ($packageFile.Extension -notin $validExtensions) {
                throw "Invalid package extension: $($packageFile.Extension). Expected: $($validExtensions -join ', ')"
            }

            # Create temp directory for extraction
            $tempDir = [System.IO.Path]::Combine($env:TEMP, "AppxInfo_$(New-Guid)")
            [void](New-Item -Path $tempDir -ItemType Directory -Force)
            
            try {
                # Open archive
                $archive = [System.IO.Compression.ZipFile]::OpenRead($packagePath)
                
                # Find manifest
                $manifestEntry = $archive.Entries | Where-Object { $_.Name -eq 'AppxManifest.xml' } | Select-Object -First 1
                
                if ($null -eq $manifestEntry) {
                    throw "AppxManifest.xml not found in package"
                }

                # Extract manifest
                $manifestPath = [System.IO.Path]::Combine($tempDir, 'AppxManifest.xml')
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($manifestEntry, $manifestPath, $true)
                
                # Parse manifest
                $manifestData = Get-AppxManifestData -ManifestPath $manifestPath -IncludeDependencies -IncludeCapabilities
                
                # Build basic result
                $result = [PSCustomObject]@{
                    PSTypeName              = 'AppxBackup.PackageInfo'
                    PackageFilePath         = $packagePath
                    PackageFileName         = $packageFile.Name
                    PackageExtension        = $packageFile.Extension
                    PackageSizeBytes        = $packageFile.Length
                    PackageSizeMB           = [Math]::Round($packageFile.Length / 1MB, 2)
                    PackageCreatedDate      = $packageFile.CreationTime
                    PackageModifiedDate     = $packageFile.LastWriteTime
                    PackageName             = $manifestData.Name
                    PackageVersion          = $manifestData.Version
                    PackagePublisher        = $manifestData.Publisher
                    PackageArchitecture     = $manifestData.ProcessorArchitecture
                    DisplayName             = $manifestData.DisplayName
                    PublisherDisplayName    = $manifestData.PublisherDisplayName
                    Description             = $manifestData.Description
                    IsBundle                = $manifestData.IsBundle
                    IsMSIX                  = $manifestData.IsMSIX
                    ApplicationCount        = $manifestData.Applications.Count
                    Applications            = $manifestData.Applications
                    DependencyCount         = $manifestData.Dependencies.Count
                    Dependencies            = $manifestData.Dependencies
                    CapabilityCount         = $manifestData.Capabilities.Count
                    Capabilities            = $manifestData.Capabilities
                    TargetDeviceFamilies    = $manifestData.TargetDeviceFamilies
                    EntryCount              = $archive.Entries.Count
                    FileList                = $null
                    SignatureInfo           = $null
                    ManifestXml             = $null
                }

                # Include file list if requested
                if ($IncludeFileList.IsPresent) {
                    $result.FileList = $archive.Entries | ForEach-Object {
                        [PSCustomObject]@{
                            FullName = $_.FullName
                            Name = $_.Name
                            Length = $_.Length
                            CompressedLength = $_.CompressedLength
                            CompressionRatio = if ($_.Length -gt 0) { 
                                [Math]::Round((1 - ($_.CompressedLength / $_.Length)) * 100, 1) 
                            } else { 0 }
                        }
                    } | Sort-Object FullName
                    
                    Write-AppxLog -Message "Included file list: $($result.FileList.Count) files" -Level 'Debug'
                }

                # Include signature info if requested
                if ($IncludeSignatureInfo.IsPresent) {
                    try {
                        $signature = Get-AuthenticodeSignature -FilePath $packagePath
                        
                        $result.SignatureInfo = [PSCustomObject]@{
                            Status = $signature.Status
                            StatusMessage = $signature.StatusMessage
                            SignerCertificate = if ($signature.SignerCertificate) {
                                [PSCustomObject]@{
                                    Subject = $signature.SignerCertificate.Subject
                                    Issuer = $signature.SignerCertificate.Issuer
                                    Thumbprint = $signature.SignerCertificate.Thumbprint
                                    NotBefore = $signature.SignerCertificate.NotBefore
                                    NotAfter = $signature.SignerCertificate.NotAfter
                                }
                            } else { $null }
                            TimeStamperCertificate = if ($signature.TimeStamperCertificate) {
                                [PSCustomObject]@{
                                    Subject = $signature.TimeStamperCertificate.Subject
                                    Issuer = $signature.TimeStamperCertificate.Issuer
                                }
                            } else { $null }
                        }
                        
                        Write-AppxLog -Message "Signature status: $($signature.Status)" -Level 'Debug'
                    }
                    catch {
                        Write-AppxLog -Message "Failed to get signature info: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                        $result.SignatureInfo = $null
                    }
                }

                # Include manifest XML if requested
                if ($IncludeManifestXml.IsPresent) {
                    $result.ManifestXml = Get-Content -LiteralPath $manifestPath -Raw
                    Write-AppxLog -Message "Included raw manifest XML" -Level 'Debug'
                }

                Write-AppxLog -Message "Package analysis complete: $($result.PackageName) v$($result.PackageVersion)" -Level 'Info'
                
                return $result
            }
            finally {
                # Cleanup
                if ($archive) { $archive.Dispose() }
                if (Test-Path -LiteralPath $tempDir) {
                    Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
        }
        catch {
            Write-AppxLog -Message "Failed to get package info: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'
            Write-AppxLog -Message "StackTrace: $($_.ScriptStackTrace)" -Level 'Debug'
            throw
        }
    }
}

<#
.SYNOPSIS
    Internal helper to analyze .appxpack ZIP archive contents.

.DESCRIPTION
    Extracts and analyzes AppxBackupManifest.json and package structure
    from a .appxpack ZIP-based dependency archive.

.OUTPUTS
    PSCustomObject with PackageInfo type containing archive metadata
#>
function Get-AppxpackArchiveInfo {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)]
        [string]$PackagePath,
        
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$PackageFile,
        
        [Parameter()]
        [switch]$IncludeFileList,
        
        [Parameter()]
        [switch]$IncludeSignatureInfo,
        
        [Parameter()]
        [switch]$IncludeManifestXml
    )
    
    Write-AppxLog -Message "Analyzing .appxpack archive: $($PackageFile.Name)" -Level 'Info'
    
    # Create temp extraction directory
    $tempDir = [System.IO.Path]::Combine(
        $env:TEMP,
        "AppxpackInfo_$(New-Guid)"
    )
    [void](New-Item -Path $tempDir -ItemType Directory -Force)
    
    try {
        # Open ZIP archive
        $archive = [System.IO.Compression.ZipFile]::OpenRead($PackagePath)
        
        try {
            # Load directory names from configuration
            $zipConfig = Get-AppxConfiguration -ConfigName 'ZipPackagingConfiguration'
            $manifestFileName = $zipConfig.archiveStructure.manifestFileName
            $packagesDirName = $zipConfig.archiveStructure.packagesDirectory
            $certsDirName = $zipConfig.archiveStructure.certificatesDirectory
            $readmeFileName = $zipConfig.archiveStructure.readmeFileName
            
            Write-AppxLog -Message "Archive structure config: Manifest=$manifestFileName, Packages=$packagesDirName, Certs=$certsDirName" -Level 'Debug'
            
            # Find and extract AppxBackupManifest.json
            $manifestEntry = $archive.Entries | Where-Object { $_.Name -eq $manifestFileName } | Select-Object -First 1
            
            if (-not $manifestEntry) {
                throw "AppxBackupManifest.json not found in archive. This may not be a valid .appxpack file."
            }
            
            $manifestPath = [System.IO.Path]::Combine($tempDir, $manifestFileName)
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($manifestEntry, $manifestPath, $true)
            
            # Parse manifest JSON
            $manifestContent = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
            Write-AppxLog -Message "Parsed AppxBackupManifest.json - Version: $($manifestContent.Version)" -Level 'Debug'
            
            # Extract main package info from manifest
            $mainPkg = $manifestContent.MainPackage
            
            # Enumerate all packages in archive
            $packageEntries = $archive.Entries | Where-Object { 
                $_.FullName -like "$packagesDirName/*" -and 
                $_.Name -match '\.(appx|msix)$' 
            }
            
            Write-AppxLog -Message "Found $($packageEntries.Count) packages in archive" -Level 'Debug'
            
            # Build contained packages metadata
            $containedPackages = @()
            foreach ($pkgEntry in $packageEntries) {
                $pkgInfo = [PSCustomObject]@{
                    FileName          = $pkgEntry.Name
                    FullPath          = $pkgEntry.FullName
                    SizeBytes         = $pkgEntry.Length
                    SizeMB            = [Math]::Round($pkgEntry.Length / 1MB, 2)
                    CompressedBytes   = $pkgEntry.CompressedLength
                    CompressionRatio  = if ($pkgEntry.Length -gt 0) { 
                        [Math]::Round((1 - ($pkgEntry.CompressedLength / $pkgEntry.Length)) * 100, 1) 
                    } else { 0 }
                }
                
                # Try to match with manifest data
                $isMainPackage = $pkgEntry.Name -eq [System.IO.Path]::GetFileName($mainPkg.PackageFile)
                $pkgInfo | Add-Member -NotePropertyName 'IsMainPackage' -NotePropertyValue $isMainPackage
                
                if ($isMainPackage) {
                    $pkgInfo | Add-Member -NotePropertyName 'PackageName' -NotePropertyValue $mainPkg.Name
                    $pkgInfo | Add-Member -NotePropertyName 'Version' -NotePropertyValue $mainPkg.Version
                    $pkgInfo | Add-Member -NotePropertyName 'Architecture' -NotePropertyValue $mainPkg.Architecture
                    $pkgInfo | Add-Member -NotePropertyName 'Publisher' -NotePropertyValue $mainPkg.Publisher
                }
                else {
                    # Match against dependencies
                    $matchedDep = $manifestContent.Dependencies | Where-Object {
                        $pkgEntry.Name -eq [System.IO.Path]::GetFileName($_.PackageFile)
                    } | Select-Object -First 1
                    
                    if ($matchedDep) {
                        $pkgInfo | Add-Member -NotePropertyName 'PackageName' -NotePropertyValue $matchedDep.Name
                        $pkgInfo | Add-Member -NotePropertyName 'Version' -NotePropertyValue $matchedDep.Version
                        $pkgInfo | Add-Member -NotePropertyName 'Architecture' -NotePropertyValue $matchedDep.Architecture
                        $pkgInfo | Add-Member -NotePropertyName 'Publisher' -NotePropertyValue $matchedDep.Publisher
                        $pkgInfo | Add-Member -NotePropertyName 'InstallOrder' -NotePropertyValue $matchedDep.InstallOrder
                        $pkgInfo | Add-Member -NotePropertyName 'IsDependency' -NotePropertyValue $true
                    }
                    else {
                        $pkgInfo | Add-Member -NotePropertyName 'IsDependency' -NotePropertyValue $false
                    }
                }
                
                # Add signature info if requested
                if ($IncludeSignatureInfo.IsPresent) {
                    # Extract package to temp for signature analysis
                    $tempPkgPath = [System.IO.Path]::Combine($tempDir, $pkgEntry.Name)
                    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($pkgEntry, $tempPkgPath, $true)
                    
                    try {
                        $signature = Get-AuthenticodeSignature -FilePath $tempPkgPath
                        
                        $sigInfo = [PSCustomObject]@{
                            Status            = $signature.Status
                            StatusMessage     = $signature.StatusMessage
                            SignerThumbprint  = if ($signature.SignerCertificate) { 
                                $signature.SignerCertificate.Thumbprint 
                            } else { $null }
                        }
                        
                        $pkgInfo | Add-Member -NotePropertyName 'SignatureInfo' -NotePropertyValue $sigInfo
                        
                        Write-AppxLog -Message "Signature for $($pkgEntry.Name): $($signature.Status)" -Level 'Debug'
                    }
                    catch {
                        Write-AppxLog -Message "Failed to get signature for $($pkgEntry.Name): $_" -Level 'Warning'
                        $pkgInfo | Add-Member -NotePropertyName 'SignatureInfo' -NotePropertyValue $null
                    }
                    finally {
                        # Cleanup temp package
                        if (Test-Path -LiteralPath $tempPkgPath) {
                            Remove-Item -LiteralPath $tempPkgPath -Force -ErrorAction SilentlyContinue
                        }
                    }
                }
                
                $containedPackages += $pkgInfo
            }
            
            # Enumerate certificates
            $certEntries = $archive.Entries | Where-Object { 
                $_.FullName -like "$certsDirName/*" -and 
                $_.Name -match '\.cer$' 
            }
            
            Write-AppxLog -Message "Found $($certEntries.Count) certificates in archive" -Level 'Debug'
            
            $containedCertificates = @()
            foreach ($certEntry in $certEntries) {
                $certInfo = [PSCustomObject]@{
                    FileName        = $certEntry.Name
                    FullPath        = $certEntry.FullName
                    SizeBytes       = $certEntry.Length
                }
                
                # Try to extract certificate details
                if ($IncludeSignatureInfo.IsPresent) {
                    $tempCertPath = [System.IO.Path]::Combine($tempDir, $certEntry.Name)
                    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($certEntry, $tempCertPath, $true)
                    
                    try {
                        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($tempCertPath)
                        
                        $certInfo | Add-Member -NotePropertyName 'Subject' -NotePropertyValue $cert.Subject
                        $certInfo | Add-Member -NotePropertyName 'Thumbprint' -NotePropertyValue $cert.Thumbprint
                        $certInfo | Add-Member -NotePropertyName 'NotBefore' -NotePropertyValue $cert.NotBefore
                        $certInfo | Add-Member -NotePropertyName 'NotAfter' -NotePropertyValue $cert.NotAfter
                        
                        Write-AppxLog -Message "Certificate $($certEntry.Name): $($cert.Subject)" -Level 'Debug'
                        
                        $cert.Dispose()
                    }
                    catch {
                        Write-AppxLog -Message "Failed to parse certificate $($certEntry.Name): $_" -Level 'Warning'
                    }
                    finally {
                        if (Test-Path -LiteralPath $tempCertPath) {
                            Remove-Item -LiteralPath $tempCertPath -Force -ErrorAction SilentlyContinue
                        }
                    }
                }
                
                $containedCertificates += $certInfo
            }
            
            # Calculate total archive size
            $totalArchiveSize = ($containedPackages | Measure-Object -Property SizeBytes -Sum).Sum
            $totalArchiveSizeMB = [Math]::Round($totalArchiveSize / 1MB, 2)
            
            # Extract main package manifest if requested
            $manifestXml = $null
            if ($IncludeManifestXml.IsPresent) {
                # Find main package in archive
                $mainPkgEntry = $packageEntries | Where-Object { 
                    $_.Name -eq [System.IO.Path]::GetFileName($mainPkg.PackageFile) 
                } | Select-Object -First 1
                
                if ($mainPkgEntry) {
                    # Extract main package to temp
                    $tempMainPkgPath = [System.IO.Path]::Combine($tempDir, $mainPkgEntry.Name)
                    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($mainPkgEntry, $tempMainPkgPath, $true)
                    
                    try {
                        # Open main package as ZIP
                        $mainPkgArchive = [System.IO.Compression.ZipFile]::OpenRead($tempMainPkgPath)
                        
                        try {
                            $appxManifestEntry = $mainPkgArchive.Entries | Where-Object { 
                                $_.Name -eq 'AppxManifest.xml' 
                            } | Select-Object -First 1
                            
                            if ($appxManifestEntry) {
                                $appxManifestPath = [System.IO.Path]::Combine($tempDir, 'AppxManifest.xml')
                                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($appxManifestEntry, $appxManifestPath, $true)
                                
                                $manifestXml = Get-Content -LiteralPath $appxManifestPath -Raw
                                Write-AppxLog -Message "Extracted AppxManifest.xml from main package" -Level 'Debug'
                            }
                        }
                        finally {
                            $mainPkgArchive.Dispose()
                        }
                    }
                    catch {
                        Write-AppxLog -Message "Failed to extract AppxManifest.xml from main package: $_" -Level 'Warning'
                    }
                    finally {
                        if (Test-Path -LiteralPath $tempMainPkgPath) {
                            Remove-Item -LiteralPath $tempMainPkgPath -Force -ErrorAction SilentlyContinue
                        }
                    }
                }
            }
            
            # Build file list if requested
            $archiveFileList = $null
            if ($IncludeFileList.IsPresent) {
                $archiveFileList = $archive.Entries | ForEach-Object {
                    [PSCustomObject]@{
                        FullName         = $_.FullName
                        Name             = $_.Name
                        Length           = $_.Length
                        CompressedLength = $_.CompressedLength
                        CompressionRatio = if ($_.Length -gt 0) { 
                            [Math]::Round((1 - ($_.CompressedLength / $_.Length)) * 100, 1) 
                        } else { 0 }
                    }
                } | Sort-Object FullName
                
                Write-AppxLog -Message "Included archive file list: $($archiveFileList.Count) entries" -Level 'Debug'
            }
            
            # Build result object
            $result = [PSCustomObject]@{
                PSTypeName                = 'AppxBackup.PackageInfo'
                PackageFilePath           = $PackagePath
                PackageFileName           = $PackageFile.Name
                PackageExtension          = $PackageFile.Extension
                PackageSizeBytes          = $PackageFile.Length
                PackageSizeMB             = [Math]::Round($PackageFile.Length / 1MB, 2)
                PackageCreatedDate        = $PackageFile.CreationTime
                PackageModifiedDate       = $PackageFile.LastWriteTime
                IsZipArchive              = $true
                ArchiveVersion            = $manifestContent.Version
                ArchiveCreatedBy          = $manifestContent.CreatedBy
                ArchiveCreatedDate        = $manifestContent.CreatedDate
                PackageName               = $mainPkg.Name
                PackageVersion            = $mainPkg.Version
                PackagePublisher          = $mainPkg.Publisher
                PackageArchitecture       = $mainPkg.Architecture
                DisplayName               = $mainPkg.PublisherDisplayName
                PublisherDisplayName      = $mainPkg.PublisherDisplayName
                TotalPackagesInArchive    = $manifestContent.TotalPackages
                DependencyCount           = $manifestContent.Dependencies.Count
                ContainedPackages         = $containedPackages
                ContainedCertificates     = $containedCertificates
                TotalArchiveContentSize   = $totalArchiveSize
                TotalArchiveContentSizeMB = $totalArchiveSizeMB
                InstallationOrder         = $manifestContent.InstallationOrder
                RequiresElevation         = $manifestContent.RequiresElevation
                MinimumOSVersion          = $manifestContent.MinimumOSVersion
                MinimumPowerShellVersion  = $manifestContent.MinimumPowerShellVersion
                ArchiveManifest           = $manifestContent
                EntryCount                = $archive.Entries.Count
                FileList                  = $archiveFileList
                SignatureInfo             = $null  # Not applicable to ZIP container
                ManifestXml               = $manifestXml
            }
            
            Write-AppxLog -Message "Archive analysis complete: $($mainPkg.Name) v$($mainPkg.Version) with $($manifestContent.TotalPackages) packages" -Level 'Info'
            
            return $result
        }
        finally {
            if ($archive) { $archive.Dispose() }
        }
    }
    catch {
        Write-AppxLog -Message "Failed to analyze .appxpack archive: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'
        throw
    }
    finally {
        # Cleanup temp directory
        if (Test-Path -LiteralPath $tempDir) {
            Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}