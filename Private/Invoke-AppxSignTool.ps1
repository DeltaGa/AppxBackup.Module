<#
.SYNOPSIS
    Signs an APPX/MSIX package using SignTool.exe with automatic tool discovery.

.DESCRIPTION
    Encapsulates the complete SignTool signing workflow:
    1. Discovers SignTool.exe via Find-AppxSdkTool
    2. Builds the appropriate argument list for APPX signing
    3. Executes SignTool via Invoke-ProcessSafely
    4. Validates the result and provides structured output

    Replaces three duplicate SignTool invocation blocks in Backup-AppxPackage.

.PARAMETER PackagePath
    Full path to the package file to sign.

.PARAMETER CertificateThumbprint
    Thumbprint of the signing certificate (must be in CurrentUser\My store).

.PARAMETER HashAlgorithm
    File digest algorithm for signing. Default: SHA256.

.PARAMETER TimeoutSeconds
    Maximum time to wait for SignTool. Default: loaded from ToolConfiguration.

.OUTPUTS
    PSCustomObject with:
    - Success: [bool] Whether signing succeeded
    - SignToolPath: [string] Path to the SignTool used
    - Output: [string] SignTool output
    - Error: [string] Error details on failure

.NOTES
    Author: DeltaGa
    Version: 2.0.2
#>

function Invoke-AppxSignTool {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$PackagePath,

        [Parameter(Mandatory, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$CertificateThumbprint,

        [Parameter()]
        [ValidateSet('SHA256', 'SHA384', 'SHA512')]
        [string]$HashAlgorithm = 'SHA256',

        [Parameter()]
        [ValidateRange(1, 3600)]
        [int]$TimeoutSeconds
    )

    process {
        try {
            # Discover SignTool
            $signToolPath = Find-AppxSdkTool -ToolName 'signtool.exe'

            if ($null -eq $signToolPath) {
                throw "SignTool.exe not found. Install Windows SDK to sign APPX packages."
            }

            Write-AppxLog -Message "Using SignTool: $signToolPath" -Level 'Debug'

            # Load timeout from configuration if not specified
            if (-not $PSBoundParameters.ContainsKey('TimeoutSeconds')) {
                try {
                    $toolConfig = Get-AppxConfiguration -ConfigName 'ToolConfiguration'
                    $TimeoutSeconds = $toolConfig.toolConfigurations.SignTool.timeoutSeconds
                }
                catch {
                    $TimeoutSeconds = 300
                }
            }

            # Build SignTool arguments for APPX signing
            # sign   = sign command
            # /fd    = file digest algorithm (required)
            # /sha1  = certificate thumbprint
            # /v     = verbose output
            $signArgs = @(
                'sign',
                '/fd', $HashAlgorithm,
                '/sha1', $CertificateThumbprint,
                '/v',
                $PackagePath
            )

            Write-AppxLog -Message "SignTool command: $signToolPath $($signArgs -join ' ')" -Level 'Debug'

            # Execute SignTool
            $signResult = Invoke-ProcessSafely -FilePath $signToolPath `
                -ArgumentList $signArgs `
                -TimeoutSeconds $TimeoutSeconds `
                -NoWindow

            if (-not $signResult.Success) {
                $errorDetails = "SignTool failed with exit code $($signResult.ExitCode)"
                if ($signResult.StandardError) {
                    $errorDetails += "`nSTDERR: $($signResult.StandardError)"
                }
                if ($signResult.StandardOutput) {
                    $errorDetails += "`nSTDOUT: $($signResult.StandardOutput)"
                }
                throw $errorDetails
            }

            Write-AppxLog -Message "Package signed successfully" -Level 'Info'

            if ($signResult.StandardOutput) {
                Write-AppxLog -Message "SignTool output: $($signResult.StandardOutput)" -Level 'Debug'
            }

            return [PSCustomObject]@{
                Success      = $true
                SignToolPath = $signToolPath
                Output       = $signResult.StandardOutput
                Error        = $null
            }
        }
        catch {
            Write-AppxLog -Message "Failed to sign package: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'

            return [PSCustomObject]@{
                Success      = $false
                SignToolPath = $signToolPath
                Output       = $null
                Error        = $_.Exception.Message
            }
        }
    }
}
