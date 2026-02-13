<#
.SYNOPSIS
    Installs a certificate to the Trusted Root store with privilege fallback.

.DESCRIPTION
    Attempts to install a certificate file to the LocalMachine\Root store
    (system-wide trust, requires Administrator). If that fails due to
    insufficient privileges, falls back to CurrentUser\Root (user-level trust).

    This centralizes the certificate installation logic that was previously
    duplicated across Backup-AppxPackage (3 locations) and Install-AppxBackup.

.PARAMETER CertificatePath
    Full path to the .cer certificate file to install.

.PARAMETER PackageIdentifier
    Optional descriptive identifier for log messages (e.g., package name).

.OUTPUTS
    PSCustomObject with:
    - Installed: [bool] Whether the certificate was successfully installed
    - StoreLocation: [string] The store where it was installed ('LocalMachine' or 'CurrentUser')
    - Error: [string] Error message on failure, $null on success

.NOTES
    Author: DeltaGa
    Version: 2.0.1
#>

function Install-AppxCertificateToStore {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$CertificatePath,

        [Parameter(Position = 1)]
        [string]$PackageIdentifier = ''
    )

    process {
        $logPrefix = if ($PackageIdentifier) { "[$PackageIdentifier] " } else { '' }

        try {
            if (-not (Test-Path -LiteralPath $CertificatePath)) {
                throw "Certificate file not found: $CertificatePath"
            }

            # Strategy 1: LocalMachine\Root (system-wide trust, requires Administrator)
            try {
                Import-Certificate -FilePath $CertificatePath `
                    -CertStoreLocation 'Cert:\LocalMachine\Root' `
                    -ErrorAction Stop | Out-Null

                Write-AppxLog -Message "${logPrefix}Certificate installed to LocalMachine\Root (system-wide trust)" -Level 'Info'

                return [PSCustomObject]@{
                    Installed     = $true
                    StoreLocation = 'LocalMachine'
                    Error         = $null
                }
            }
            catch {
                Write-AppxLog -Message "${logPrefix}Cannot install to LocalMachine\Root (not Administrator), trying CurrentUser..." -Level 'Debug'
            }

            # Strategy 2: CurrentUser\Root (user-level trust, no admin required)
            Import-Certificate -FilePath $CertificatePath `
                -CertStoreLocation 'Cert:\CurrentUser\Root' `
                -ErrorAction Stop | Out-Null

            Write-AppxLog -Message "${logPrefix}Certificate installed to CurrentUser\Root (user trust only)" -Level 'Warning'
            Write-AppxLog -Message "${logPrefix}For system-wide trust, run as Administrator" -Level 'Warning'

            return [PSCustomObject]@{
                Installed     = $true
                StoreLocation = 'CurrentUser'
                Error         = $null
            }
        }
        catch {
            Write-AppxLog -Message "${logPrefix}Failed to install certificate: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'

            return [PSCustomObject]@{
                Installed     = $false
                StoreLocation = $null
                Error         = $_.Exception.Message
            }
        }
    }
}
