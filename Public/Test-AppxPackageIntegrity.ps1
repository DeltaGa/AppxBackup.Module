<#
.SYNOPSIS
    Validates the integrity and signature of an APPX/MSIX package.

.DESCRIPTION
    Comprehensive package validation including:
    - Signature verification
    - Manifest structure validation
    - Archive integrity checking
    - Certificate chain validation

.PARAMETER PackagePath
    Path to the .appx or .msix file to validate.

.PARAMETER VerifySignature
    If specified, validates the digital signature.

.PARAMETER CheckManifest
    If specified, validates manifest structure.

.EXAMPLE
    Test-AppxPackageIntegrity -PackagePath "C:\Backups\MyApp.appx" -VerifySignature

.OUTPUTS
    AppxBackup.IntegrityResult
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
                    
                    if (-not $signatureValid) {
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
                    $tempDir = Join-Path $env:TEMP "AppxValidation_$(New-Guid)"
                    [void](New-Item -Path $tempDir -ItemType Directory -Force)
                    
                    try {
                        $archive = [System.IO.Compression.ZipFile]::OpenRead($packagePath)
                        $manifestEntry = $archive.Entries | Where-Object { $_.Name -eq 'AppxManifest.xml' } | Select-Object -First 1
                        
                        if ($manifestEntry) {
                            $manifestPath = Join-Path $tempDir 'AppxManifest.xml'
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
            Write-AppxLog -Message "Package integrity check failed: $_" -Level 'Error'
            throw
        }
    }
}
