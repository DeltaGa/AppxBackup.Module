<#
.SYNOPSIS
    Creates a self-signed certificate for a dependency package.

.DESCRIPTION
    Internal helper function that creates a self-signed certificate specifically
    for dependency packages during ZIP-based dependency packaging.
    
    This function is called for each dependency when -IncludeDependencies is used,
    ensuring that every package in the ZIP archive has its corresponding certificate
    for installation.
    
    Key differences from New-AppxBackupCertificate:
    - Streamlined for batch dependency processing
    - No user interaction or confirmations
    - No certificate installation (done during Install-AppxBackup)
    - Optimized for speed in loops
    - Publisher must be provided (obtained from Get-AppxPackage)

.PARAMETER PackageName
    Name of the dependency package, used for certificate filename generation.
    Example: "Microsoft.UI.Xaml.2.8_8.2501.31001.0_x64"

.PARAMETER OutputDirectory
    Directory where the certificate (.cer) file will be saved.

.PARAMETER PublisherSubject
    Publisher subject from Get-AppxPackage (CN=...).
    Required parameter - must be obtained from the installed package.

.OUTPUTS
    [PSCustomObject]
    Properties:
    - Success: Boolean indicating operation success
    - CertificatePath: Full path to the .cer file
    - Thumbprint: Certificate thumbprint
    - Subject: Certificate subject
    - ValidUntil: Certificate expiration date
    - StoreLocation: Certificate store path

.NOTES
    Author: DeltaGa
    Version: 2.0.1
    
    This function is designed for high-volume dependency processing.
    It uses cached configuration values for performance.

.EXAMPLE
    $cert = New-AppxDependencyCertificate `
        -PackageName "Microsoft.UI.Xaml.2.8_8.2501.31001.0_x64" `
        -OutputDirectory $outDir `
        -PublisherSubject $dep.Publisher
    if ($cert.Success) {
        Write-Host "Certificate: $($cert.Thumbprint)"
    }
#>

function New-AppxDependencyCertificate {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$PackageName,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$OutputDirectory,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$PublisherSubject
    )
    
    Write-AppxLog -Message "=== New-AppxDependencyCertificate ===" -Level 'Debug'
    Write-AppxLog -Message "Package: $PackageName" -Level 'Debug'
    Write-AppxLog -Message "Publisher: $PublisherSubject" -Level 'Debug'
    
    try {
        # Ensure publisher subject starts with CN=
        if (-not $PublisherSubject.StartsWith('CN=')) {
            $PublisherSubject = "CN=$PublisherSubject"
            Write-AppxLog -Message "Normalized publisher to: $PublisherSubject" -Level 'Debug'
        }
        
        # Generate certificate filename based on package name
        $certFileName = "$PackageName.cer"
        $certOutputPath = [System.IO.Path]::Combine($OutputDirectory, $certFileName)
        
        Write-AppxLog -Message "Certificate output: $certOutputPath" -Level 'Debug'
        
        # Get certificate validity period from configuration
        $validityYears = Get-AppxDefault 'certificateDefaults.defaultValidityYears' 'ModuleDefaults' 3
        $notAfter = (Get-Date).AddYears($validityYears)
        
        # Get certificate store location from configuration
        $certStoreLocation = Get-AppxDefault 'certificateSettings.defaultStoreLocation' 'ZipPackagingConfiguration' 'Cert:\CurrentUser\My'
        
        Write-AppxLog -Message "Creating certificate: $PublisherSubject" -Level 'Info'
        Write-AppxLog -Message "Validity: $validityYears years (until $($notAfter.ToString('yyyy-MM-dd')))" -Level 'Debug'
        
        # Create the self-signed certificate
        # Using CurrentUser store to avoid requiring elevation for dependency certificates
        $cert = New-SelfSignedCertificate `
            -Type CodeSigningCert `
            -Subject $PublisherSubject `
            -KeyLength 4096 `
            -HashAlgorithm SHA256 `
            -NotAfter $notAfter `
            -CertStoreLocation $certStoreLocation `
            -ErrorAction Stop
        
        if (-not $cert) {
            throw "Certificate creation returned null"
        }
        
        $thumbprint = $cert.Thumbprint
        Write-AppxLog -Message "Certificate created: $thumbprint" -Level 'Debug'
        
        # Export public certificate to .cer file
        $certBytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
        [System.IO.File]::WriteAllBytes($certOutputPath, $certBytes)
        
        Write-AppxLog -Message "Certificate exported: $certOutputPath" -Level 'Debug'
        
        # Verify export
        if (-not (Test-Path -LiteralPath $certOutputPath)) {
            throw "Certificate export verification failed: $certOutputPath"
        }
        
        $certFileSize = (Get-Item -LiteralPath $certOutputPath).Length
        Write-AppxLog -Message "Certificate size: $certFileSize bytes" -Level 'Debug'
        
        Write-AppxLog -Message "Dependency certificate created successfully" -Level 'Info'
        
        # Return result object
        return [PSCustomObject]@{
            Success         = $true
            CertificatePath = $certOutputPath
            Thumbprint      = $thumbprint
            Subject         = $PublisherSubject
            ValidUntil      = $notAfter
            StoreLocation   = "$certStoreLocation\$thumbprint"
        }
    }
    catch {
        Write-AppxLog -Message "Dependency certificate creation failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'
        
        return [PSCustomObject]@{
            Success         = $false
            CertificatePath = $null
            Thumbprint      = $null
            Subject         = $PublisherSubject
            ValidUntil      = $null
            StoreLocation   = $null
            Error           = $_.Exception.Message
        }
    }
}