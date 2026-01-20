<#
.SYNOPSIS
    Creates a ZIP archive containing APPX packages, certificates, and metadata.

.DESCRIPTION
    Internal helper function that creates a structured ZIP archive (.appxpack)
    containing:
    - All APPX packages (main + dependencies)
    - All certificates (.cer files)
    - AppxBackupManifest.json (installation metadata)
    - README.txt (installation instructions)
    
    This function implements the ZIP-based dependency packaging architecture,
    replacing the deprecated MakeAppx bundle approach which was incompatible
    with dependency packaging.
    
    The ZIP structure follows a consistent format:
    /Packages/          - All .appx/.msix files
    /Certificates/      - All .cer files
    AppxBackupManifest.json - Installation orchestration metadata
    README.txt          - Human-readable installation guide

.PARAMETER SourceDirectory
    Temporary directory containing all packages and certificates to be archived.

.PARAMETER OutputPath
    Full path for the output .appxpack file (ZIP archive).

.PARAMETER ManifestData
    Hashtable containing the AppxBackupManifest.json structure.

.PARAMETER CompressionLevel
    ZIP compression level. Valid values: NoCompression, Fastest, Optimal.
    Default: Optimal

.OUTPUTS
    [PSCustomObject]
    Properties:
    - Success: Boolean indicating operation success
    - ZipPath: Full path to created ZIP file
    - TotalSize: Total size in bytes
    - TotalSizeMB: Total size formatted in MB
    - PackageCount: Number of packages included
    - CertificateCount: Number of certificates included

.NOTES
    Author: DeltaGa
    Version: 2.0.0
    
    This is a critical component of the ZIP-based dependency packaging system.
    It ensures proper structure and metadata for Install-AppxBackup orchestration.

.EXAMPLE
    $result = New-AppxBackupZipArchive -SourceDirectory $tempDir -OutputPath $zipPath -ManifestData $manifest
    if ($result.Success) {
        Write-Host "ZIP created: $($result.TotalSizeMB)"
    }
#>

function New-AppxBackupZipArchive {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$SourceDirectory,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$OutputPath,
        
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$ManifestData,
        
        [Parameter()]
        [ValidateSet('NoCompression', 'Fastest', 'Optimal')]
        [string]$CompressionLevel = 'Optimal'
    )
    
    Write-AppxLog -Message "=== New-AppxBackupZipArchive ===" -Level 'Debug'
    Write-AppxLog -Message "Source: $SourceDirectory" -Level 'Debug'
    Write-AppxLog -Message "Output: $OutputPath" -Level 'Debug'
    Write-AppxLog -Message "Compression: $CompressionLevel" -Level 'Debug'
    
    try {
        # Validate source directory exists
        if (-not (Test-Path -LiteralPath $SourceDirectory -PathType Container)) {
            throw "Source directory does not exist: $SourceDirectory"
        }
        
        # Ensure output directory exists
        $outputDir = [System.IO.Path]::GetDirectoryName($OutputPath)
        if (-not (Test-Path -LiteralPath $outputDir)) {
            New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
            Write-AppxLog -Message "Created output directory: $outputDir" -Level 'Debug'
        }
        
        # Remove existing ZIP if present
        if (Test-Path -LiteralPath $OutputPath) {
            Remove-Item -LiteralPath $OutputPath -Force
            Write-AppxLog -Message "Removed existing ZIP: $OutputPath" -Level 'Debug'
        }
        
        # Load System.IO.Compression for ZIP operations
        Add-Type -AssemblyName 'System.IO.Compression.FileSystem'
        
        # Map compression level parameter to .NET enum
        $compressionMap = @{
            'NoCompression' = [System.IO.Compression.CompressionLevel]::NoCompression
            'Fastest'       = [System.IO.Compression.CompressionLevel]::Fastest
            'Optimal'       = [System.IO.Compression.CompressionLevel]::Optimal
        }
        $compressionEnum = $compressionMap[$CompressionLevel]
        
        Write-AppxLog -Message "Creating ZIP archive with $CompressionLevel compression..." -Level 'Info'
        
        # Create the ZIP archive
        # Note: We use CreateFromDirectory for simplicity, then add metadata files separately
        $tempZipDir = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "AppxZipTemp_$(New-Guid)")
        New-Item -Path $tempZipDir -ItemType Directory -Force | Out-Null
        
        try {
            # Get directory names from configuration
            $packagesDirName = Get-AppxDefault 'archiveStructure.packagesDirectory' 'ZipPackagingConfiguration' 'Packages'
            $certsDirName = Get-AppxDefault 'archiveStructure.certificatesDirectory' 'ZipPackagingConfiguration' 'Certificates'
            
            # Create subdirectories in temp location
            $packagesDir = [System.IO.Path]::Combine($tempZipDir, $packagesDirName)
            $certsDir = [System.IO.Path]::Combine($tempZipDir, $certsDirName)
            New-Item -Path $packagesDir -ItemType Directory -Force | Out-Null
            New-Item -Path $certsDir -ItemType Directory -Force | Out-Null
            
            # Copy packages
            $packageFiles = Get-ChildItem -LiteralPath $SourceDirectory -Filter '*.appx' -File -ErrorAction SilentlyContinue
            $packageFiles += Get-ChildItem -LiteralPath $SourceDirectory -Filter '*.msix' -File -ErrorAction SilentlyContinue
            
            $packageCount = 0
            foreach ($pkg in $packageFiles) {
                $destPath = [System.IO.Path]::Combine($packagesDir, $pkg.Name)
                Copy-Item -LiteralPath $pkg.FullName -Destination $destPath -Force
                $packageCount++
                Write-AppxLog -Message "Added package: $($pkg.Name)" -Level 'Debug'
            }
            
            # Copy certificates
            $certFiles = Get-ChildItem -LiteralPath $SourceDirectory -Filter '*.cer' -File -ErrorAction SilentlyContinue
            $certCount = 0
            foreach ($cert in $certFiles) {
                $destPath = [System.IO.Path]::Combine($certsDir, $cert.Name)
                Copy-Item -LiteralPath $cert.FullName -Destination $destPath -Force
                $certCount++
                Write-AppxLog -Message "Added certificate: $($cert.Name)" -Level 'Debug'
            }
            
            # Create AppxBackupManifest.json
            $manifestFileName = Get-AppxDefault 'archiveStructure.manifestFileName' 'ZipPackagingConfiguration' 'AppxBackupManifest.json'
            $manifestPath = [System.IO.Path]::Combine($tempZipDir, $manifestFileName)
            $manifestJson = $ManifestData | ConvertTo-Json -Depth 10
            [System.IO.File]::WriteAllText($manifestPath, $manifestJson, [System.Text.Encoding]::UTF8)
            Write-AppxLog -Message "Created $manifestFileName" -Level 'Debug'
            
            # Create README.txt with installation instructions
            $readmeFileName = Get-AppxDefault 'archiveStructure.readmeFileName' 'ZipPackagingConfiguration' 'README.txt'
            $projectUrl = Get-AppxDefault 'readmeTemplate.projectUrl' 'ZipPackagingConfiguration' 'https://github.com/DeltaGa/AppxBackup.Module'
            $readmePath = [System.IO.Path]::Combine($tempZipDir, $readmeFileName)
            $readmeContent = @"
==============================================================================
AppxBackup Package Archive - Installation Guide
==============================================================================

Package: $($ManifestData.MainPackage.Name)
Version: $($ManifestData.MainPackage.Version)
Architecture: $($ManifestData.MainPackage.Architecture)
Created: $($ManifestData.CreatedDate)
Total Packages: $($ManifestData.TotalPackages)

==============================================================================
INSTALLATION INSTRUCTIONS
==============================================================================

METHOD 1: Automated Installation (Recommended)
-----------------------------------------------
1. Use the Install-AppxBackup PowerShell function:

   Install-AppxBackup -PackagePath "$([System.IO.Path]::GetFileName($OutputPath))"

   This will:
   - Extract all packages and certificates
   - Install certificates to Trusted Root store
   - Install dependencies in correct order
   - Install the main package
   - Verify installation


METHOD 2: Manual Installation
------------------------------
1. Extract this ZIP archive to a temporary location

2. Install all certificates in the Certificates/ folder:
   
   Get-ChildItem ".\Certificates\*.cer" | ForEach-Object {
       Import-Certificate -FilePath `$_.FullName -CertStoreLocation Cert:\LocalMachine\Root
   }

3. Install packages in the order specified in AppxBackupManifest.json:
   
   # Dependencies first (in InstallationOrder), then main package
   Add-AppxPackage -Path ".\Packages\<dependency>.appx"
   Add-AppxPackage -Path ".\Packages\<main-package>.appx"


METHOD 3: Using Windows Package Manager (GUI)
----------------------------------------------
1. Extract this ZIP archive
2. Double-click each certificate in Certificates/ and install to Trusted Root
3. Right-click the main package in Packages/ and select "Install"
   (Dependencies must be installed first if present)


==============================================================================
TROUBLESHOOTING
==============================================================================

Error: Certificate not trusted (0x800B0109)
-------------------------------------------
Solution: Ensure all certificates are installed to Trusted Root store
Run PowerShell as Administrator to install to LocalMachine store


Error: Package already installed (0x80073CF9)
---------------------------------------------
Solution: Uninstall existing version first
Get-AppxPackage -Name "$($ManifestData.MainPackage.Name)" | Remove-AppxPackage


Error: Dependencies missing
---------------------------
Solution: Install dependencies in the order specified in AppxBackupManifest.json
All required dependencies are included in the Packages/ folder


==============================================================================
PACKAGE CONTENTS
==============================================================================

AppxBackupManifest.json - Installation metadata and package information
README.txt              - This file
Packages/               - All APPX/MSIX packages (main + dependencies)
Certificates/           - Self-signed certificates for package installation


==============================================================================
For more information, visit: $projectUrl
==============================================================================
"@
            [System.IO.File]::WriteAllText($readmePath, $readmeContent, [System.Text.Encoding]::UTF8)
            Write-AppxLog -Message "Created $readmeFileName" -Level 'Debug'
            
            # Create the final ZIP archive
            Write-AppxLog -Message "Compressing to ZIP archive..." -Level 'Info'
            [System.IO.Compression.ZipFile]::CreateFromDirectory(
                $tempZipDir,
                $OutputPath,
                $compressionEnum,
                $false  # Don't include base directory
            )
            
            # Get final file size
            $zipInfo = Get-Item -LiteralPath $OutputPath
            $totalSize = $zipInfo.Length
            $totalSizeMB = [math]::Round($totalSize / 1MB, 2)
            
            Write-AppxLog -Message "ZIP archive created successfully: $totalSizeMB MB" -Level 'Info'
            
            # Return result object
            return [PSCustomObject]@{
                Success          = $true
                ZipPath          = $OutputPath
                TotalSize        = $totalSize
                TotalSizeMB      = "$totalSizeMB MB"
                PackageCount     = $packageCount
                CertificateCount = $certCount
            }
        }
        finally {
            # Clean up temp directory
            if (Test-Path -LiteralPath $tempZipDir) {
                Remove-Item -Path $tempZipDir -Recurse -Force -ErrorAction SilentlyContinue
                Write-AppxLog -Message "Cleaned up temp ZIP directory" -Level 'Debug'
            }
        }
    }
    catch {
        Write-AppxLog -Message "ZIP archive creation failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'
        
        return [PSCustomObject]@{
            Success          = $false
            ZipPath          = $null
            TotalSize        = 0
            TotalSizeMB      = "0 MB"
            PackageCount     = 0
            CertificateCount = 0
            Error            = $_.Exception.Message
        }
    }
}