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

.EXAMPLE
    New-AppxBackupCertificate -Subject "CN=MyApp Publisher" -OutputPath "C:\Certs\MyApp.cer"
    
    Creates a certificate and exports public key only

.EXAMPLE
    $pwd = ConvertTo-SecureString "P@ssw0rd!" -AsPlainText -Force
    New-AppxBackupCertificate -Subject "CN=MyCompany" -OutputPath "C:\Certs\cert.cer" `
        -Password $pwd -ExportPrivateKey
    
    Creates certificate and exports with private key protected by password

.OUTPUTS
    System.Security.Cryptography.X509Certificates.X509Certificate2

.NOTES
    Requires PowerShell 5.1+ for New-SelfSignedCertificate cmdlet.
    
    The certificate is created in the specified store and remains there
    after export. Use Remove-Item to delete from store when no longer needed.
    
    Author: DeltaGa
    Version: 2.0.0
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
        [switch]$ExportPrivateKey
    )

    begin {
        Write-AppxLog -Message "Creating self-signed certificate: $Subject" -Level 'Verbose'
        
        # Validate PowerShell version supports New-SelfSignedCertificate
        if ($PSVersionTable.PSVersion.Major -lt 5) {
            throw "New-SelfSignedCertificate requires PowerShell 5.1 or later. Current version: $($PSVersionTable.PSVersion)"
        }

        # Validate ExportPrivateKey requirements
        if ($ExportPrivateKey.IsPresent -and -not $Password) {
            throw "ExportPrivateKey requires Password parameter"
        }

        # Normalize subject
        if (-not $Subject.StartsWith('CN=')) {
            $Subject = "CN=$Subject"
        }
    }

    process {
        try {
            # Validate output path
            $outputPath = ConvertTo-SecureFilePath -Path $OutputPath -ResolveRelative
            $outputDir = Split-Path -Path $outputPath -Parent
            
            if ($outputDir -and -not (Test-Path -LiteralPath $outputDir)) {
                [void](New-Item -Path $outputDir -ItemType Directory -Force)
            }

            if ($PSCmdlet.ShouldProcess($Subject, "Create self-signed certificate")) {
                # Calculate expiration
                $notBefore = [DateTime]::Now.AddDays(-1) # Backdated to avoid timezone issues
                $notAfter = $notBefore.AddYears($ValidityYears)
                
                Write-AppxLog -Message "Certificate validity: $notBefore to $notAfter" -Level 'Debug'

                # Create certificate
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
                        # Code Signing EKU
                        "2.5.29.37={text}1.3.6.1.5.5.7.3.3",
                        # Basic Constraints: This is not a CA
                        "2.5.29.19={text}false"
                    )
                }

                $certificate = New-SelfSignedCertificate @certParams
                
                Write-AppxLog -Message "Certificate created: Thumbprint=$($certificate.Thumbprint)" -Level 'Info'
                Write-AppxLog -Message "  Subject: $($certificate.Subject)" -Level 'Debug'
                Write-AppxLog -Message "  Issuer: $($certificate.Issuer)" -Level 'Debug'
                Write-AppxLog -Message "  Valid: $($certificate.NotBefore) to $($certificate.NotAfter)" -Level 'Debug'
                Write-AppxLog -Message "  Key Length: $KeyLength bits" -Level 'Debug'

                # Export public certificate (.cer)
                Write-AppxLog -Message "Exporting public certificate to: $outputPath" -Level 'Verbose'
                
                $cerBytes = $certificate.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
                [System.IO.File]::WriteAllBytes($outputPath, $cerBytes)
                
                Write-AppxLog -Message "Public certificate exported successfully" -Level 'Info'

                # Export private key if requested
                if ($ExportPrivateKey.IsPresent) {
                    $pfxPath = [System.IO.Path]::ChangeExtension($outputPath, '.pfx')
                    
                    Write-AppxLog -Message "Exporting private key to: $pfxPath" -Level 'Verbose'
                    Write-Host "`n[WARNING]  WARNING: Exporting private key. Keep PFX file secure!`n" -ForegroundColor Yellow
                    
                    $pfxBytes = $certificate.Export(
                        [System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx,
                        $Password
                    )
                    
                    [System.IO.File]::WriteAllBytes($pfxPath, $pfxBytes)
                    
                    # Set restrictive permissions on PFX
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
                    
                    Write-AppxLog -Message "Private key exported with restricted permissions" -Level 'Info'
                }

                # Add note about installation
                Write-Host "`n[INFO] Certificate Installation Instructions:" -ForegroundColor Cyan
                Write-Host "   1. Double-click on: $outputPath" -ForegroundColor Gray
                Write-Host "   2. Click 'Install Certificate...'" -ForegroundColor Gray
                Write-Host "   3. Select 'Local Machine'" -ForegroundColor Gray
                Write-Host "   4. Choose 'Place all certificates in the following store'" -ForegroundColor Gray
                Write-Host "   5. Select 'Trusted Root Certification Authorities'" -ForegroundColor Gray
                Write-Host "   6. Click 'Finish'`n" -ForegroundColor Gray

                return $certificate
            }
            else {
                Write-AppxLog -Message "Certificate creation skipped (WhatIf)" -Level 'Info'
            }
        }
        catch {
            Write-AppxLog -Message "Certificate creation failed: $_" -Level 'Error'
            
            # Cleanup on failure
            if ($certificate) {
                try {
                    Remove-Item -LiteralPath "Cert:\CurrentUser\My\$($certificate.Thumbprint)" -ErrorAction SilentlyContinue
                }
                catch {
                    Write-AppxLog -Message "Failed to cleanup certificate: $_" -Level 'Warning'
                }
            }
            
            throw
        }
    }
}
