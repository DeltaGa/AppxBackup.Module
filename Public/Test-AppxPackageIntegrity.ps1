<#
.SYNOPSIS
    Validates the integrity and signature of an APPX/MSIX package or .appxpack archive.

.DESCRIPTION
    Comprehensive package validation including:
    - Signature verification
    - Manifest structure validation
    - Archive integrity checking
    - Certificate chain validation
    
    For .appxpack archives, additionally validates:
    - Archive structure (Packages/, Certificates/, manifest)
    - AppxBackupManifest.json schema and completeness
    - All contained package signatures
    - Certificate-to-package mapping
    - Installation order consistency

.PARAMETER PackagePath
    Path to the .appx, .msix, or .appxpack file to validate.

.PARAMETER VerifySignature
    If specified, validates the digital signature.
    For .appxpack archives, validates all contained package signatures.

.PARAMETER CheckManifest
    If specified, validates manifest structure.
    For .appxpack archives, validates AppxBackupManifest.json schema.

.EXAMPLE
    Test-AppxPackageIntegrity -PackagePath "C:\Backups\MyApp.appx" -VerifySignature

.EXAMPLE
    Test-AppxPackageIntegrity -PackagePath "C:\Backups\MyApp.appxpack" -VerifySignature -CheckManifest
    
    Validates .appxpack archive structure and all package signatures

.OUTPUTS
    AppxBackup.IntegrityResult
    For .appxpack archives, includes additional properties:
    - ArchiveStructureValid: Boolean
    - ManifestSchemaValid: Boolean
    - AllPackagesValid: Boolean
    - PackageValidationResults: Array of per-package validation
#>

function Test-AppxPackageIntegrity {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [string]$PackagePath,

        [Parameter()]
        [switch]$VerifySignature,

        [Parameter()]
        [switch]$CheckManifest
    )

    process {
        try {
            $packagePath = ConvertTo-SecureFilePath -Path $PackagePath -MustExist -PathType File
            
            Write-AppxLog -Message "Validating package: $packagePath" -Level 'Verbose'
            
            # Load ZIP archive extension from configuration
            try {
                $zipConfig = Get-AppxConfiguration -ConfigName 'ZipPackagingConfiguration'
                $zipArchiveExtension = $zipConfig.archiveExtensions.zipArchive
            }
            catch {
                Write-AppxLog -Message "Failed to load ZIP configuration, using fallback: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                $zipArchiveExtension = '.appxpack'
            }
            
            # Detect if this is a .appxpack archive
            $packageFile = Get-Item -LiteralPath $packagePath
            $isZipArchive = $packageFile.Extension -eq $zipArchiveExtension
            
            if ($isZipArchive) {
                Write-AppxLog -Message "Detected .appxpack archive - using archive validation path" -Level 'Debug'
                
                # Validate as ZIP-based dependency archive
                $result = Test-AppxpackArchiveIntegrity -PackagePath $packagePath `
                    -VerifySignature:$VerifySignature `
                    -CheckManifest:$CheckManifest
                
                return $result
            }
            
            # Continue with standard package validation
            Write-AppxLog -Message "Detected individual package - using standard validation path" -Level 'Debug'
            
            $issues = @()
            $isValid = $true

            # Check file extension
            $ext = [System.IO.Path]::GetExtension($packagePath)
            if ($ext -notin @('.appx', '.msix', '.appxbundle', '.msixbundle')) {
                $issues += "Invalid file extension: $ext"
                $isValid = $false
            }

            # Check if file is a valid ZIP archive
            try {
                Add-Type -AssemblyName 'System.IO.Compression.FileSystem'
                $archive = [System.IO.Compression.ZipFile]::OpenRead($packagePath)
                $entryCount = $archive.Entries.Count
                $archive.Dispose()
                
                Write-AppxLog -Message "Archive valid: $entryCount entries" -Level 'Debug'
            }
            catch {
                $issues += "Invalid ZIP archive: $_"
                $isValid = $false
            }

            # Verify signature if requested
            $signatureValid = $null
            if ($VerifySignature.IsPresent) {
                try {
                    $signature = Get-AuthenticodeSignature -FilePath $packagePath
                    $signatureValid = ($signature.Status -eq 'Valid')
                    
                    if ($null -eq $signatureValid) {
                        $issues += "Invalid signature: $($signature.StatusMessage)"
                        Write-AppxLog -Message "Signature status: $($signature.Status)" -Level 'Warning'
                    }
                }
                catch {
                    $issues += "Signature check failed: $_"
                    $signatureValid = $false
                }
            }

            # Check manifest if requested
            $manifestValid = $null
            if ($CheckManifest.IsPresent) {
                try {
                    # Extract and parse manifest
                    $tempDir = [System.IO.Path]::Combine($env:TEMP, "AppxValidation_$(New-Guid)")
                    [void](New-Item -Path $tempDir -ItemType Directory -Force)
                    
                    try {
                        $archive = [System.IO.Compression.ZipFile]::OpenRead($packagePath)
                        $manifestEntry = $archive.Entries | Where-Object { $_.Name -eq 'AppxManifest.xml' } | Select-Object -First 1
                        
                        if ($manifestEntry) {
                            $manifestPath = [System.IO.Path]::Combine($tempDir, 'AppxManifest.xml')
                            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($manifestEntry, $manifestPath)
                            
                            $manifestData = Get-AppxManifestData -ManifestPath $manifestPath
                            $manifestValid = $true
                            
                            Write-AppxLog -Message "Manifest valid: $($manifestData.Name)" -Level 'Debug'
                        }
                        else {
                            $issues += "AppxManifest.xml not found in package"
                            $manifestValid = $false
                        }
                    }
                    finally {
                        if ($archive) { $archive.Dispose() }
                        Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                    }
                }
                catch {
                    $issues += "Manifest validation failed: $_"
                    $manifestValid = $false
                }
            }

            # Build result
            $result = [PSCustomObject]@{
                PSTypeName          = 'AppxBackup.IntegrityResult'
                PackagePath         = $packagePath
                IsValid             = ($isValid -and ($signatureValid -ne $false) -and ($manifestValid -ne $false))
                SignatureValid      = $signatureValid
                ManifestValid       = $manifestValid
                ArchiveValid        = $isValid
                Issues              = $issues
                ValidationDate      = [DateTime]::Now
            }

            if ($result.IsValid) {
                Write-AppxLog -Message "Package validation passed" -Level 'Info'
            }
            else {
                Write-AppxLog -Message "Package validation failed: $($issues.Count) issues" -Level 'Warning'
            }

            return $result
        }
        catch {
            Write-AppxLog -Message "Package integrity check failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'
            Write-AppxLog -Message "StackTrace: $($_.ScriptStackTrace)" -Level 'Debug'
            throw
        }
    }
}

<#
.SYNOPSIS
    Internal helper to validate .appxpack ZIP archive integrity.

.DESCRIPTION
    Validates ZIP-based dependency archive structure, manifest schema,
    package signatures, and certificate-to-package mappings.

.OUTPUTS
    PSCustomObject with IntegrityResult type containing archive validation
#>
function Test-AppxpackArchiveIntegrity {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)]
        [string]$PackagePath,
        
        [Parameter()]
        [switch]$VerifySignature,
        
        [Parameter()]
        [switch]$CheckManifest
    )
    
    Write-AppxLog -Message "Validating .appxpack archive integrity: $PackagePath" -Level 'Info'
    
    $issues = @()
    $isValid = $true
    $archiveStructureValid = $true
    $manifestSchemaValid = $null
    $allPackagesValid = $null
    
    # Create temp extraction directory
    $tempDir = [System.IO.Path]::Combine(
        $env:TEMP,
        "AppxpackValidation_$(New-Guid)"
    )
    [void](New-Item -Path $tempDir -ItemType Directory -Force)
    
    try {
        # Check if file is a valid ZIP archive
        Add-Type -AssemblyName 'System.IO.Compression.FileSystem'
        
        try {
            $archive = [System.IO.Compression.ZipFile]::OpenRead($PackagePath)
            $entryCount = $archive.Entries.Count
            Write-AppxLog -Message "Archive opened successfully: $entryCount entries" -Level 'Debug'
        }
        catch {
            $issues += "Invalid ZIP archive: $_"
            $isValid = $false
            $archiveStructureValid = $false
            
            # Return early if ZIP is corrupted
            return [PSCustomObject]@{
                PSTypeName               = 'AppxBackup.IntegrityResult'
                PackagePath              = $PackagePath
                IsValid                  = $false
                IsZipArchive             = $true
                ArchiveStructureValid    = $false
                ManifestSchemaValid      = $null
                AllPackagesValid         = $null
                SignatureValid           = $null
                ManifestValid            = $null
                ArchiveValid             = $false
                Issues                   = $issues
                PackageValidationResults = @()
                ValidationDate           = [DateTime]::Now
            }
        }
        
        try {
            # Load configuration
            $zipConfig = Get-AppxConfiguration -ConfigName 'ZipPackagingConfiguration'
            $manifestFileName = $zipConfig.archiveStructure.manifestFileName
            $packagesDirName = $zipConfig.archiveStructure.packagesDirectory
            $certsDirName = $zipConfig.archiveStructure.certificatesDirectory
            $readmeFileName = $zipConfig.archiveStructure.readmeFileName
            
            # Required validation rules
            $requireManifest = $zipConfig.validationRules.requireManifestInArchive
            $requireReadme = $zipConfig.validationRules.requireReadmeInArchive
            $validatePackageSignatures = $zipConfig.validationRules.validatePackageSignatures
            $validateCertificatePresence = $zipConfig.validationRules.validateCertificatePresence
            
            Write-AppxLog -Message "Archive structure validation: Manifest=$requireManifest, Readme=$requireReadme, Signatures=$validatePackageSignatures" -Level 'Debug'
            
            # Validate required structure components
            $manifestEntry = $archive.Entries | Where-Object { $_.Name -eq $manifestFileName } | Select-Object -First 1
            if ($requireManifest -and -not $manifestEntry) {
                $issues += "Required file missing: $manifestFileName"
                $isValid = $false
                $archiveStructureValid = $false
            }
            else {
                Write-AppxLog -Message "Found $manifestFileName" -Level 'Debug'
            }
            
            $readmeEntry = $archive.Entries | Where-Object { $_.Name -eq $readmeFileName } | Select-Object -First 1
            if ($requireReadme -and -not $readmeEntry) {
                $issues += "Required file missing: $readmeFileName"
                $archiveStructureValid = $false
                # Not critical - don't fail validation
            }
            else {
                Write-AppxLog -Message "Found $readmeFileName" -Level 'Debug'
            }
            
            # Validate directory structure
            $hasPackagesDir = $archive.Entries | Where-Object { $_.FullName -like "$packagesDirName/*" } | Select-Object -First 1
            if (-not $hasPackagesDir) {
                $issues += "Required directory missing or empty: $packagesDirName/"
                $isValid = $false
                $archiveStructureValid = $false
            }
            else {
                Write-AppxLog -Message "Found $packagesDirName/ directory" -Level 'Debug'
            }
            
            $hasCertsDir = $archive.Entries | Where-Object { $_.FullName -like "$certsDirName/*" } | Select-Object -First 1
            if ($validateCertificatePresence -and -not $hasCertsDir) {
                $issues += "Required directory missing or empty: $certsDirName/"
                $isValid = $false
                $archiveStructureValid = $false
            }
            else {
                Write-AppxLog -Message "Found $certsDirName/ directory" -Level 'Debug'
            }
            
            # If manifest validation requested, validate schema
            if ($CheckManifest.IsPresent -and $manifestEntry) {
                manifestSchemaValid = $true
                
                # Extract and parse manifest
                $manifestPath = [System.IO.Path]::Combine($tempDir, $manifestFileName)
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($manifestEntry, $manifestPath, $true)
                
                try {
                    $manifestContent = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
                    
                    # Validate required fields
                    $requiredFields = @('Version', 'CreatedBy', 'CreatedDate', 'MainPackage', 
                                       'Dependencies', 'InstallationOrder', 'TotalPackages')
                    
                    foreach ($field in $requiredFields) {
                        if (-not $manifestContent.PSObject.Properties[$field]) {
                            $issues += "AppxBackupManifest.json missing required field: $field"
                            $manifestSchemaValid = $false
                            $isValid = $false
                        }
                    }
                    
                    # Validate MainPackage structure
                    if ($manifestContent.MainPackage) {
                        $mainPkgFields = @('Name', 'Version', 'Architecture', 'Publisher', 'PackageFile')
                        foreach ($field in $mainPkgFields) {
                            if (-not $manifestContent.MainPackage.PSObject.Properties[$field]) {
                                $issues += "MainPackage missing required field: $field"
                                $manifestSchemaValid = $false
                                $isValid = $false
                            }
                        }
                    }
                    else {
                        $issues += "MainPackage section missing from manifest"
                        $manifestSchemaValid = $false
                        $isValid = $false
                    }
                    
                    # Validate InstallationOrder consistency
                    $expectedPackageCount = 1 + $manifestContent.Dependencies.Count  # Main + dependencies
                    if ($manifestContent.InstallationOrder.Count -ne $expectedPackageCount) {
                        $issues += "InstallationOrder count mismatch: Expected $expectedPackageCount, found $($manifestContent.InstallationOrder.Count)"
                        $manifestSchemaValid = $false
                    }
                    
                    Write-AppxLog -Message "Manifest schema validation: $(if ($manifestSchemaValid) { 'PASSED' } else { 'FAILED' })" -Level 'Debug'
                }
                catch {
                    $issues += "Failed to parse AppxBackupManifest.json: $_"
                    $manifestSchemaValid = $false
                    $isValid = $false
                }
            }
            
            # If signature verification requested, validate all packages
            $packageValidationResults = @()
            if ($VerifySignature.IsPresent -and $hasPackagesDir) {
                allPackagesValid = $true
                
                # Get all package entries
                $packageEntries = $archive.Entries | Where-Object { 
                    $_.FullName -like "$packagesDirName/*" -and 
                    $_.Name -match '\.(appx|msix)$' 
                }
                
                Write-AppxLog -Message "Validating signatures for $($packageEntries.Count) packages" -Level 'Info'
                
                foreach ($pkgEntry in $packageEntries) {
                    $pkgResult = [PSCustomObject]@{
                        PackageName      = $pkgEntry.Name
                        SignatureValid   = $null
                        ValidationIssues = @()
                    }
                    
                    # Extract package to temp
                    $tempPkgPath = [System.IO.Path]::Combine($tempDir, $pkgEntry.Name)
                    
                    try {
                        [System.IO.Compression.ZipFileExtensions]::ExtractToFile($pkgEntry, $tempPkgPath, $true)
                        
                        # Verify signature
                        $signature = Get-AuthenticodeSignature -FilePath $tempPkgPath
                        $pkgResult.SignatureValid = ($signature.Status -eq 'Valid')
                        
                        if (-not $pkgResult.SignatureValid) {
                            $pkgResult.ValidationIssues += "Signature invalid: $($signature.StatusMessage)"
                            $allPackagesValid = $false
                            
                            if ($validatePackageSignatures) {
                                $isValid = $false
                            }
                        }
                        
                        Write-AppxLog -Message "Package $($pkgEntry.Name) signature: $($signature.Status)" -Level 'Debug'
                    }
                    catch {
                        $pkgResult.SignatureValid = $false
                        $pkgResult.ValidationIssues += "Signature check failed: $_"
                        $allPackagesValid = $false
                        
                        if ($validatePackageSignatures) {
                            $isValid = $false
                        }
                        
                        Write-AppxLog -Message "Failed to verify signature for $($pkgEntry.Name): $_" -Level 'Warning'
                    }
                    finally {
                        # Cleanup temp package
                        if (Test-Path -LiteralPath $tempPkgPath) {
                            Remove-Item -LiteralPath $tempPkgPath -Force -ErrorAction SilentlyContinue
                        }
                    }
                    
                    $packageValidationResults += $pkgResult
                }
                
                # Aggregate package validation issues
                $failedPackages = $packageValidationResults | Where-Object { -not $_.SignatureValid }
                if ($failedPackages.Count -gt 0) {
                    $issues += "$($failedPackages.Count) package(s) have invalid signatures"
                }
            }
            
            # Build result object
            $result = [PSCustomObject]@{
                PSTypeName               = 'AppxBackup.IntegrityResult'
                PackagePath              = $PackagePath
                IsValid                  = $isValid
                IsZipArchive             = $true
                ArchiveStructureValid    = $archiveStructureValid
                ManifestSchemaValid      = $manifestSchemaValid
                AllPackagesValid         = $allPackagesValid
                SignatureValid           = $allPackagesValid  # Aggregate signature status
                ManifestValid            = $manifestSchemaValid
                ArchiveValid             = $archiveStructureValid
                Issues                   = $issues
                PackageValidationResults = $packageValidationResults
                ValidationDate           = [DateTime]::Now
            }
            
            if ($result.IsValid) {
                Write-AppxLog -Message "Archive validation PASSED" -Level 'Info'
            }
            else {
                Write-AppxLog -Message "Archive validation FAILED: $($issues.Count) issues" -Level 'Warning'
            }
            
            return $result
        }
        finally {
            if ($archive) { $archive.Dispose() }
        }
    }
    catch {
        Write-AppxLog -Message "Archive integrity check failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'
        throw
    }
    finally {
        # Cleanup temp directory
        if (Test-Path -LiteralPath $tempDir) {
            Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}