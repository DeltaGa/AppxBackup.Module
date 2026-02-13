<#
.SYNOPSIS
    Validates disk space and package size before APPX packaging operations.

.DESCRIPTION
    Calculates source directory size, checks available disk space against
    configurable thresholds, and warns users about large packages or low
    disk space conditions. Uses configuration from PackageConfiguration.json.

    This is a pre-flight validation step - it throws only for insufficient
    disk space, and issues warnings for other concerning conditions.

.PARAMETER SourcePath
    Full path to the source directory to be packaged.

.PARAMETER OutputPath
    Full path where the output package will be created.

.OUTPUTS
    PSCustomObject with:
    - SourceSizeMB: [double] Size of source directory in MB
    - FileCount: [int] Number of files in source directory
    - AvailableSpaceMB: [double] Available disk space on output drive
    - RequiredSpaceMB: [double] Estimated required space
    - IsLargePackage: [bool] Whether package exceeds size threshold
    - IsHighFileCount: [bool] Whether file count exceeds threshold

.NOTES
    Author: DeltaGa
    Version: 2.0.1
#>

function Test-AppxDiskSpace {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$SourcePath,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$OutputPath
    )

    process {
        Write-AppxLog -Message "Validating package size and disk space..." -Level 'Verbose'

        $result = [PSCustomObject]@{
            SourceSizeMB     = 0
            FileCount        = 0
            AvailableSpaceMB = 0
            RequiredSpaceMB  = 0
            IsLargePackage   = $false
            IsHighFileCount  = $false
        }

        # Calculate source directory size
        $sourceFiles = Get-ChildItem -LiteralPath $SourcePath -Recurse -File -ErrorAction Stop
        $sourceSize = ($sourceFiles | Measure-Object -Property Length -Sum).Sum
        $result.SourceSizeMB = [Math]::Round($sourceSize / 1MB, 2)
        $result.FileCount = $sourceFiles.Count

        Write-AppxLog -Message "Source directory: $($result.SourceSizeMB) MB ($($result.FileCount) files)" -Level 'Verbose'

        # Load thresholds from configuration
        $largeSizeThreshold = Get-AppxDefault 'packageSizeThresholds.largeSizeMB' -Fallback 1024
        $largeFileCountThreshold = Get-AppxDefault 'packageSizeThresholds.largeFileCount' -Fallback 10000

        # Warn if package is large
        if ($result.SourceSizeMB -gt $largeSizeThreshold) {
            $result.IsLargePackage = $true
            Write-Host "[WARNING] Large package detected ($($result.SourceSizeMB) MB, $($result.FileCount) files)" -ForegroundColor Yellow
            Write-Host "  Processing may take several minutes. Please be patient..." -ForegroundColor Yellow
            Write-AppxLog -Message "Large package warning displayed to user" -Level 'Info'
        }

        # Warn if too many files
        if ($result.FileCount -gt $largeFileCountThreshold) {
            $result.IsHighFileCount = $true
            Write-Host "[WARNING] Package contains many files ($($result.FileCount))" -ForegroundColor Yellow
            Write-Host "  Packaging performance may be slower than usual..." -ForegroundColor Yellow
            Write-AppxLog -Message "High file count warning displayed to user" -Level 'Info'
        }

        # Check available disk space
        $outputDrive = [System.IO.Path]::GetPathRoot($OutputPath)
        if ($outputDrive) {
            $driveInfo = [System.IO.DriveInfo]::new($outputDrive)
            $result.AvailableSpaceMB = [Math]::Round($driveInfo.AvailableFreeSpace / 1MB, 2)

            # Load disk space calculation parameters from configuration
            $multiplier = Get-AppxDefault 'diskSpaceRequirements.sourceMultiplier' -Fallback 2
            $safetyMargin = Get-AppxDefault 'diskSpaceRequirements.safetyMarginMB' -Fallback 100
            $minimalRemaining = Get-AppxDefault 'diskSpaceRequirements.minimalAvailableSpaceMB' -Fallback 500

            $result.RequiredSpaceMB = ($result.SourceSizeMB * $multiplier) + $safetyMargin

            Write-AppxLog -Message "Disk space - Available: $($result.AvailableSpaceMB) MB, Required: $($result.RequiredSpaceMB) MB" -Level 'Debug'

            if ($result.AvailableSpaceMB -lt $result.RequiredSpaceMB) {
                $shortfall = [Math]::Round($result.RequiredSpaceMB - $result.AvailableSpaceMB, 2)
                throw "Insufficient disk space on $outputDrive. Need $shortfall MB more space. Available: $($result.AvailableSpaceMB) MB, Required: $($result.RequiredSpaceMB) MB"
            }

            # Warn if available space is marginal
            if (($result.AvailableSpaceMB - $result.RequiredSpaceMB) -lt $minimalRemaining) {
                Write-Warning "Disk space will be low after packaging (less than $minimalRemaining MB remaining)"
                Write-AppxLog -Message "Low disk space warning displayed" -Level 'Warning'
            }
        }

        return $result
    }
}
