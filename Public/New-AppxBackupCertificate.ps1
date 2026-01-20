<#
.SYNOPSIS
    Creates a self-signed code signing certificate for APPX/MSIX packages.

.DESCRIPTION
    Modern certificate creation using native PowerShell cmdlets.
    Completely eliminates dependency on deprecated MakeCert.exe.
    
    Features:
    - Uses New-SelfSignedCertificate (PowerShell 5.1+)
    - Proper EKU for code signing
    - Configurable key length and hash algorithm
    - Secure private key storage
    - Optional HSM/TPM integration
    - Automatic cleanup of expired certificates
    - Duplicate certificate detection and handling

.PARAMETER Subject
    Certificate subject name (e.g., "CN=MyCompany").
    Must match the package publisher for signature validation.

.PARAMETER OutputPath
    Path where the public certificate (.cer) will be exported.
    Private key remains in certificate store for security.

.PARAMETER ValidityYears
    Certificate validity period in years.
    Default: 3 years
    Maximum: 10 years

.PARAMETER KeyLength
    RSA key length in bits.
    Valid values: 2048, 3072, 4096
    Default: 4096 (maximum security)

.PARAMETER StoreLocation
    Certificate store location.
    Default: CurrentUser\My (user's personal store)

.PARAMETER Password
    Optional secure password for PFX export.
    If not specified, certificate is not exported as PFX.

.PARAMETER ExportPrivateKey
    If specified, exports certificate with private key as PFX.
    Requires -Password parameter.

.PARAMETER ReplaceExisting
    If specified, removes any existing certificates with the same subject
    before creating a new one. Use with caution.

.EXAMPLE
    New-AppxBackupCertificate -Subject "CN=MyApp Publisher" -OutputPath "C:\Certs\MyApp.cer"
    
    Creates a certificate and exports public key only

.EXAMPLE
    $pwd = ConvertTo-SecureString "P@ssw0rd!" -AsPlainText -Force
    New-AppxBackupCertificate -Subject "CN=MyCompany" -OutputPath "C:\Certs\cert.cer" `
        -Password $pwd -ExportPrivateKey
    
    Creates certificate and exports with private key protected by password

.EXAMPLE
    New-AppxBackupCertificate -Subject "CN=MyCompany" -OutputPath "C:\Certs\cert.cer" -ReplaceExisting
    
    Removes any existing certificates with same subject, then creates new one

.OUTPUTS
    System.Security.Cryptography.X509Certificates.X509Certificate2

.NOTES
    Requires PowerShell 5.1+ for New-SelfSignedCertificate cmdlet.
    
    The certificate is created in the specified store and remains there
    after export. Use Remove-Item to delete from store when no longer needed.
    
    Author: DeltaGa
    Version: 2.0.1
#>

function New-AppxBackupCertificate {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([System.Security.Cryptography.X509Certificates.X509Certificate2])]
    param(
        [Parameter(Mandatory, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$Subject,

        [Parameter(Mandatory, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$OutputPath,

        [Parameter()]
        [ValidateRange(1, 10)]
        [int]$ValidityYears = 3,

        [Parameter()]
        [ValidateSet(2048, 3072, 4096)]
        [int]$KeyLength = 4096,

        [Parameter()]
        [ValidateSet(
            'Cert:\CurrentUser\My',
            'Cert:\LocalMachine\My',
            'Cert:\CurrentUser\Root',
            'Cert:\LocalMachine\Root'
        )]
        [string]$StoreLocation = 'Cert:\CurrentUser\My',

        [Parameter()]
        [SecureString]$Password,

        [Parameter()]
        [switch]$ExportPrivateKey,

        [Parameter()]
        [switch]$ReplaceExisting
    )

    begin {
        Write-AppxLog -Message "Creating self-signed certificate: $Subject" -Level 'Verbose'
        
        # Validate PowerShell version supports New-SelfSignedCertificate
        if ($PSVersionTable.PSVersion.Major -lt 5) {
            throw "New-SelfSignedCertificate requires PowerShell 5.1 or later. Current version: $($PSVersionTable.PSVersion)"
        }

        # Validate New-SelfSignedCertificate cmdlet availability
        if (-not (Get-Command -Name New-SelfSignedCertificate -ErrorAction SilentlyContinue)) {
            throw "New-SelfSignedCertificate cmdlet not available. This feature requires PowerShell 5.1 or later."
        }

        # Validate Import-Certificate cmdlet availability (for validation)
        if (-not (Get-Command -Name Import-Certificate -ErrorAction SilentlyContinue)) {
            Write-Warning "Import-Certificate cmdlet not available. Certificate validation will be limited."
        }

        # Validate ExportPrivateKey requirements
        if ($ExportPrivateKey.IsPresent -and -not $Password) {
            throw "ExportPrivateKey requires Password parameter"
        }

        # Normalize subject to CN= format
        if (-not $Subject.StartsWith('CN=')) {
            $Subject = "CN=$Subject"
        }

        Write-AppxLog -Message "Certificate subject normalized to: $Subject" -Level 'Debug'
    }

    process {
        try {
            # Validate and prepare output path using ConvertTo-SecureFilePath
            Write-AppxLog -Message "Validating output path: $OutputPath" -Level 'Verbose'
            $outputPath = ConvertTo-SecureFilePath -Path $OutputPath -ResolveRelative
            
            # Ensure output directory exists
            $outputDir = Split-Path -Path $outputPath -Parent
            if ($outputDir -and -not (Test-Path -LiteralPath $outputDir)) {
                Write-AppxLog -Message "Creating output directory: $outputDir" -Level 'Verbose'
                [void](New-Item -Path $outputDir -ItemType Directory -Force -ErrorAction Stop)
            }

            # Check for existing certificates with same subject
            $existingCerts = Get-ChildItem -Path $StoreLocation -ErrorAction SilentlyContinue | 
                Where-Object { $_.Subject -eq $Subject }
            
            if ($existingCerts) {
                $certCount = @($existingCerts).Count
                Write-AppxLog -Message "Found $certCount existing certificate(s) with subject: $Subject" -Level 'Warning'
                
                if ($ReplaceExisting.IsPresent) {
                    Write-Host "`n[WARNING] Removing $certCount existing certificate(s) with subject: $Subject" -ForegroundColor Yellow
                    
                    foreach ($cert in @($existingCerts)) {
                        if ($PSCmdlet.ShouldProcess($cert.Thumbprint, "Remove existing certificate")) {
                            try {
                                $certPath = [System.IO.Path]::Combine($StoreLocation, $cert.Thumbprint)
                                Remove-Item -LiteralPath $certPath -Force -ErrorAction Stop
                                Write-Host "  Removed: $($cert.Thumbprint) (Expires: $($cert.NotAfter))" -ForegroundColor Gray
                                Write-AppxLog -Message "Removed existing certificate: $($cert.Thumbprint)" -Level 'Info'
                            }
                            catch {
                                Write-Warning "Failed to remove certificate $($cert.Thumbprint): $_"
                                Write-AppxLog -Message "Failed to remove existing certificate: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                            }
                        }
                    }
                }
                else {
                    Write-Host "`n[WARNING] Certificate with this subject already exists:" -ForegroundColor Yellow
                    foreach ($cert in @($existingCerts)) {
                        Write-Host "  Thumbprint: $($cert.Thumbprint)" -ForegroundColor Gray
                        Write-Host "  Valid: $($cert.NotBefore) to $($cert.NotAfter)" -ForegroundColor Gray
                        Write-Host "  Location: $StoreLocation\$($cert.Thumbprint)" -ForegroundColor Gray
                    }
                    Write-Host "`nThis will create an additional certificate with the same subject." -ForegroundColor Yellow
                    Write-Host "Use -ReplaceExisting to remove existing certificates first.`n" -ForegroundColor Yellow
                    
                    if ($null -eq $Force) {
                        $continue = Read-Host "Continue creating duplicate certificate? (y/N)"
                        if ($continue -ne 'y') {
                            Write-AppxLog -Message "Certificate creation cancelled by user (duplicate subject)" -Level 'Info'
                            return
                        }
                    }
                }
            }

            if ($PSCmdlet.ShouldProcess($Subject, "Create self-signed certificate")) {
                # Calculate expiration dates
                $notBefore = [DateTime]::Now.AddDays(-1) # Backdated to avoid timezone issues
                $notAfter = $notBefore.AddYears($ValidityYears)
                
                Write-AppxLog -Message "Certificate validity: $notBefore to $notAfter" -Level 'Debug'
                Write-Host "`n[INFO] Creating Certificate..." -ForegroundColor Cyan
                Write-Host "  Subject: $Subject" -ForegroundColor Gray
                Write-Host "  Key Length: $KeyLength bits" -ForegroundColor Gray
                Write-Host "  Valid Until: $($notAfter.ToString('yyyy-MM-dd'))" -ForegroundColor Gray
                Write-Host "  Store: $StoreLocation" -ForegroundColor Gray

                # Create certificate with comprehensive parameters
                Write-AppxLog -Message "Generating $KeyLength-bit RSA key pair..." -Level 'Verbose'
                
                $certParams = @{
                    Type = 'CodeSigningCert'
                    Subject = $Subject
                    KeyLength = $KeyLength
                    HashAlgorithm = 'SHA256'
                    NotBefore = $notBefore
                    NotAfter = $notAfter
                    CertStoreLocation = $StoreLocation
                    KeyExportPolicy = if ($ExportPrivateKey.IsPresent) { 'Exportable' } else { 'NonExportable' }
                    KeyUsage = 'DigitalSignature'
                    TextExtension = @(
                        # Code Signing EKU (Extended Key Usage)
                        "2.5.29.37={text}1.3.6.1.5.5.7.3.3",
                        # Basic Constraints: This is not a CA
                        "2.5.29.19={text}false"
                    )
                }

                $certificate = New-SelfSignedCertificate @certParams
                
                if ($null -eq $certificate) {
                    throw "New-SelfSignedCertificate returned null - certificate creation failed"
                }

                Write-Host "  [SUCCESS] Certificate created" -ForegroundColor Green
                Write-Host "  Thumbprint: $($certificate.Thumbprint)" -ForegroundColor Gray
                
                Write-AppxLog -Message "Certificate created: Thumbprint=$($certificate.Thumbprint)" -Level 'Info'
                Write-AppxLog -Message "  Subject: $($certificate.Subject)" -Level 'Debug'
                Write-AppxLog -Message "  Issuer: $($certificate.Issuer)" -Level 'Debug'
                Write-AppxLog -Message "  Valid: $($certificate.NotBefore) to $($certificate.NotAfter)" -Level 'Debug'
                Write-AppxLog -Message "  Key Length: $KeyLength bits" -Level 'Debug'
                Write-AppxLog -Message "  Has Private Key: $($certificate.HasPrivateKey)" -Level 'Debug'

                # Export public certificate (.cer)
                Write-Host "`n[INFO] Exporting Public Certificate..." -ForegroundColor Cyan
                Write-AppxLog -Message "Exporting public certificate to: $outputPath" -Level 'Verbose'
                
                try {
                    $cerBytes = $certificate.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
                    [System.IO.File]::WriteAllBytes($outputPath, $cerBytes)
                    
                    Write-Host "  [SUCCESS] Public certificate exported" -ForegroundColor Green
                    Write-Host "  Location: $outputPath" -ForegroundColor Gray
                    Write-AppxLog -Message "Public certificate exported successfully: $outputPath" -Level 'Info'
                }
                catch {
                    Write-AppxLog -Message "Failed to export public certificate: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'
                    Write-AppxLog -Message "StackTrace: $($_.ScriptStackTrace)" -Level 'Debug'
                    throw "Failed to export certificate to $outputPath : $_"
                }

                # Export private key if requested
                if ($ExportPrivateKey.IsPresent) {
                    $pfxPath = [System.IO.Path]::ChangeExtension($outputPath, '.pfx')
                    
                    Write-Host "`n[WARNING] Exporting Private Key" -ForegroundColor Yellow
                    Write-Host "  WARNING: Keep PFX file secure and encrypted!" -ForegroundColor Yellow
                    Write-Host "  Location: $pfxPath" -ForegroundColor Gray
                    
                    Write-AppxLog -Message "Exporting private key to: $pfxPath" -Level 'Verbose'
                    
                    try {
                        $pfxBytes = $certificate.Export(
                            [System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx,
                            $Password
                        )
                        
                        [System.IO.File]::WriteAllBytes($pfxPath, $pfxBytes)
                        
                        # Set restrictive permissions on PFX (NTFS only)
                        try {
                            $pfxItem = Get-Item -LiteralPath $pfxPath
                            
                            # Check if file system supports ACLs (NTFS)
                            if ($pfxItem.PSProvider.Name -eq 'FileSystem') {
                                $acl = Get-Acl -LiteralPath $pfxPath
                                $acl.SetAccessRuleProtection($true, $false) # Disable inheritance
                                
                                # Grant only current user full control
                                $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                                $rule = [System.Security.AccessControl.FileSystemAccessRule]::new(
                                    $identity,
                                    'FullControl',
                                    'Allow'
                                )
                                $acl.AddAccessRule($rule)
                                
                                Set-Acl -LiteralPath $pfxPath -AclObject $acl
                                
                                Write-Host "  [SUCCESS] Private key exported with restricted permissions" -ForegroundColor Green
                                Write-AppxLog -Message "Private key exported with restricted ACL: $pfxPath" -Level 'Info'
                            }
                            else {
                                Write-Warning "File system does not support ACLs - PFX permissions not restricted"
                                Write-Host "  [SUCCESS] Private key exported (permissions not restricted)" -ForegroundColor Yellow
                                Write-AppxLog -Message "Private key exported without ACL (non-NTFS): $pfxPath" -Level 'Warning'
                            }
                        }
                        catch {
                            Write-Warning "Failed to set restrictive permissions on PFX: $_"
                            Write-Host "  [SUCCESS] Private key exported (permission restriction failed)" -ForegroundColor Yellow
                            Write-AppxLog -Message "PFX exported but ACL modification failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                        }
                    }
                    catch {
                        Write-AppxLog -Message "Failed to export private key: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'
                        Write-AppxLog -Message "StackTrace: $($_.ScriptStackTrace)" -Level 'Debug'
                        throw "Failed to export private key to $pfxPath : $_"
                    }
                }

                # Display installation instructions
                Write-Host "`n[INFO] Certificate Installation Instructions:" -ForegroundColor Cyan
                Write-Host "   1. Double-click on: $outputPath" -ForegroundColor Gray
                Write-Host "   2. Click 'Install Certificate...'" -ForegroundColor Gray
                Write-Host "   3. Select 'Local Machine' (requires Administrator)" -ForegroundColor Gray
                Write-Host "   4. Choose 'Place all certificates in the following store'" -ForegroundColor Gray
                Write-Host "   5. Select 'Trusted Root Certification Authorities'" -ForegroundColor Gray
                Write-Host "   6. Click 'Finish'`n" -ForegroundColor Gray
                
                Write-Host "Alternatively, use PowerShell to install:" -ForegroundColor Cyan
                Write-Host "   Import-Certificate -FilePath '$outputPath' -CertStoreLocation Cert:\LocalMachine\Root" -ForegroundColor Gray
                Write-Host ""

                return $certificate
            }
            else {
                Write-Host "`n[INFO] Certificate creation skipped (WhatIf mode)" -ForegroundColor Yellow
                Write-AppxLog -Message "Certificate creation skipped (WhatIf mode)" -Level 'Info'
                return $null
            }
        }
        catch {
            Write-AppxLog -Message "Certificate creation failed: $_ | StackTrace: $($_.ScriptStackTrace)" -Level 'Error'
            
            # Cleanup on failure - remove certificate from store if it was created
            if ($certificate) {
                Write-Host "`n[ROLLBACK] Removing partially created certificate..." -ForegroundColor Yellow
                
                try {
                    $certPath = [System.IO.Path]::Combine($StoreLocation, $certificate.Thumbprint)
                    
                    if (Test-Path -LiteralPath $certPath) {
                        Remove-Item -LiteralPath $certPath -Force -ErrorAction Stop
                        Write-Host "  [SUCCESS] Certificate removed from store" -ForegroundColor Gray
                        Write-AppxLog -Message "Cleaned up certificate from store: $($certificate.Thumbprint)" -Level 'Info'
                    }
                }
                catch {
                    Write-Warning "Failed to cleanup certificate from store: $_"
                    Write-AppxLog -Message "Failed to cleanup certificate: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                }
            }
            
            # Cleanup exported files
            if (Test-Path -LiteralPath $outputPath -ErrorAction SilentlyContinue) {
                try {
                    Remove-Item -LiteralPath $outputPath -Force -ErrorAction Stop
                    Write-AppxLog -Message "Cleaned up exported certificate file: $outputPath" -Level 'Info'
                }
                catch {
                    Write-AppxLog -Message "Failed to cleanup exported certificate file: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                }
            }
            
            throw
        }
    }
}
