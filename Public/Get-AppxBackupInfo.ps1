<#
.SYNOPSIS
    Gets comprehensive information about a backed-up APPX/MSIX package.

.DESCRIPTION
    Extracts and analyzes metadata from a package file without installation.
    Provides detailed information about the package structure, manifest,
    signature, and compatibility.

.PARAMETER PackagePath
    Path to the .appx or .msix file to analyze.

.PARAMETER IncludeFileList
    If specified, includes a list of all files in the package.

.PARAMETER IncludeSignatureInfo
    If specified, includes digital signature details.

.PARAMETER IncludeManifestXml
    If specified, includes the raw manifest XML.

.EXAMPLE
    Get-AppxBackupInfo -PackagePath "C:\Backups\MyApp.appx"
    
    Gets basic package information

.EXAMPLE
    Get-AppxBackupInfo -PackagePath "C:\Backups\MyApp.appx" -IncludeFileList -IncludeSignatureInfo
    
    Gets comprehensive package information including files and signature

.OUTPUTS
    AppxBackup.PackageInfo
    
.NOTES
    Does not require the package to be installed.
    Extracts manifest temporarily to analyze structure.
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
    }

    process {
        try {
            # Validate package path
            $packagePath = ConvertTo-SecureFilePath -Path $PackagePath -MustExist -PathType File
            
            $packageFile = Get-Item -LiteralPath $packagePath
            
            Write-AppxLog -Message "Analyzing: $($packageFile.Name)" -Level 'Verbose'
            
            # Validate extension
            $validExtensions = @('.appx', '.msix', '.appxbundle', '.msixbundle')
            if ($packageFile.Extension -notin $validExtensions) {
                throw "Invalid package extension: $($packageFile.Extension). Expected: $($validExtensions -join ', ')"
            }

            # Create temp directory for extraction
            $tempDir = Join-Path $env:TEMP "AppxInfo_$(New-Guid)"
            [void](New-Item -Path $tempDir -ItemType Directory -Force)
            
            try {
                # Open archive
                $archive = [System.IO.Compression.ZipFile]::OpenRead($packagePath)
                
                # Find manifest
                $manifestEntry = $archive.Entries | Where-Object { $_.Name -eq 'AppxManifest.xml' } | Select-Object -First 1
                
                if (-not $manifestEntry) {
                    throw "AppxManifest.xml not found in package"
                }

                # Extract manifest
                $manifestPath = Join-Path $tempDir 'AppxManifest.xml'
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
                        Write-AppxLog -Message "Failed to get signature info: $_" -Level 'Warning'
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
            Write-AppxLog -Message "Failed to get package info: $_" -Level 'Error'
            throw
        }
    }
}
