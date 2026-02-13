<#
.SYNOPSIS
    Removes a file or directory with configurable retry logic.

.DESCRIPTION
    Attempts to remove an item (file or directory) with configurable retry
    count and exponential backoff delays. Designed for cleanup of temp
    directories that may be briefly locked by antivirus or other processes.

    Never throws - cleanup failures are logged as warnings but do not
    fail the calling operation.

.PARAMETER Path
    Full path to the file or directory to remove.

.PARAMETER Recurse
    If specified, removes directories recursively. Required for directories.

.OUTPUTS
    [bool] Whether cleanup succeeded.

.NOTES
    Author: DeltaGa
    Version: 2.0.2
#>

function Remove-AppxItemWithRetry {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter()]
        [switch]$Recurse
    )

    process {
        if (-not (Test-Path -LiteralPath $Path)) {
            Write-AppxLog -Message "Nothing to clean up (path does not exist): $Path" -Level 'Debug'
            return $true
        }

        Write-AppxLog -Message "Cleaning up: $Path" -Level 'Debug'

        # Load cleanup configuration
        $maxAttempts = Get-AppxDefault 'sleepDelays.maxCleanupAttempts' -Fallback 3
        $delays = Get-AppxDefault 'sleepDelays.cleanupRetryDelaysMilliseconds' -Fallback @(300, 2000, 5000)

        for ($attempt = 0; $attempt -lt $maxAttempts; $attempt++) {
            try {
                # Wait before retry (skip on first attempt)
                if ($attempt -gt 0) {
                    $delayIndex = [Math]::Min($attempt - 1, $delays.Count - 1)
                    $delayMs = $delays[$delayIndex]
                    Write-AppxLog -Message "Waiting $($delayMs)ms before cleanup attempt $($attempt + 1)/$maxAttempts..." -Level 'Debug'
                    Start-Sleep -Milliseconds $delayMs
                }

                # Attempt cleanup
                $removeParams = @{
                    LiteralPath = $Path
                    Force       = $true
                    ErrorAction = 'Stop'
                }
                if ($Recurse.IsPresent) {
                    $removeParams['Recurse'] = $true
                }

                Remove-Item @removeParams
                Write-AppxLog -Message "Cleanup successful on attempt $($attempt + 1)" -Level 'Debug'
                return $true
            }
            catch {
                Write-AppxLog -Message "Cleanup attempt $($attempt + 1) failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'

                if ($attempt -eq ($maxAttempts - 1)) {
                    # Final attempt failed - log warning but don't fail the operation
                    Write-AppxLog -Message "Failed to cleanup after $maxAttempts attempts: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                    Write-AppxLog -Message "Files remain at: $Path" -Level 'Warning'
                    Write-AppxLog -Message "These files can be manually deleted later or will be cleaned up on next reboot" -Level 'Info'
                }
            }
        }

        return $false
    }
}
