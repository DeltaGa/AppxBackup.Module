<#
.SYNOPSIS
    Copies an APPX source directory using a multi-tier fallback strategy.

.DESCRIPTION
    Performs directory copy with three fallback tiers:
    1. Robocopy (fastest, best for WindowsApps access)
    2. PowerShell Copy-Item (handles most cases)
    3. .NET System.IO APIs (most granular control, error-tolerant)

    After successful copy, removes existing signature files to prepare
    for repackaging and validates critical files are present.

.PARAMETER SourcePath
    Full path to the source directory to copy.

.PARAMETER DestinationPath
    Full path to the destination directory.

.OUTPUTS
    PSCustomObject with:
    - Success: [bool] Whether the copy succeeded
    - CopyMethod: [string] The method that succeeded ('Robocopy', 'Copy-Item', '.NET APIs')
    - Error: [string] Error message on failure

.NOTES
    Author: DeltaGa
    Version: 2.0.1
#>

function Copy-AppxSourceDirectory {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$SourcePath,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$DestinationPath
    )

    process {
        $copySucceeded = $false
        $copyMethod = 'Unknown'

        try {
            # STRATEGY 1: Robocopy with simpler flags (fastest, most reliable for WindowsApps)
            try {
                Write-AppxLog -Message "Attempting copy with Robocopy..." -Level 'Debug'

                # Simplified Robocopy - NO QUOTES in array, PowerShell handles this
                $robocopyArgs = @(
                    $SourcePath,
                    $DestinationPath,
                    '*.*',
                    '/E',
                    '/COPY:DAT',
                    '/DCOPY:DA',
                    '/R:1',
                    '/W:1',
                    '/NP',
                    '/NJH',
                    '/NJS'
                )

                Write-AppxLog -Message "Robocopy command: robocopy.exe $($robocopyArgs -join ' ')" -Level 'Debug'

                $robocopyOutput = & robocopy.exe @robocopyArgs 2>&1
                $robocopyExitCode = $LASTEXITCODE

                Write-AppxLog -Message "Robocopy exit code: $robocopyExitCode" -Level 'Debug'

                # Robocopy exit codes: 0-7 are success, 8+ are errors
                if ($robocopyExitCode -lt 8) {
                    $copySucceeded = $true
                    $copyMethod = 'Robocopy'
                    Write-AppxLog -Message "Robocopy succeeded" -Level 'Debug'
                }
                else {
                    $robocopyOutputStr = $robocopyOutput -join "`n"
                    Write-AppxLog -Message "Robocopy failed (exit $robocopyExitCode): $robocopyOutputStr" -Level 'Warning'
                    throw "Robocopy exit code $robocopyExitCode"
                }
            }
            catch {
                Write-AppxLog -Message "Robocopy failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                Write-AppxLog -Message "Falling back to PowerShell native copy..." -Level 'Verbose'

                # STRATEGY 2: PowerShell Copy-Item (handles most cases)
                try {
                    Write-AppxLog -Message "Attempting copy with Copy-Item..." -Level 'Debug'

                    $itemsToCopy = Get-ChildItem -Path $SourcePath -Recurse -Force -ErrorAction Stop

                    foreach ($item in @($itemsToCopy)) {
                        $relativePath = $item.FullName.Substring($SourcePath.Length).TrimStart('\')
                        $destPath = [System.IO.Path]::Combine($DestinationPath, $relativePath)

                        if ($item.PSIsContainer) {
                            if (-not (Test-Path -LiteralPath $destPath)) {
                                [void](New-Item -Path $destPath -ItemType Directory -Force -ErrorAction SilentlyContinue)
                            }
                        }
                        else {
                            $destDir = Split-Path -Path $destPath -Parent
                            if ($destDir -and -not (Test-Path -LiteralPath $destDir)) {
                                [void](New-Item -Path $destDir -ItemType Directory -Force -ErrorAction SilentlyContinue)
                            }
                            Copy-Item -LiteralPath $item.FullName -Destination $destPath -Force -ErrorAction Stop
                        }
                    }

                    $copySucceeded = $true
                    $copyMethod = 'Copy-Item'
                    Write-AppxLog -Message "Copy-Item succeeded" -Level 'Debug'
                }
                catch {
                    Write-AppxLog -Message "Copy-Item failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                    Write-AppxLog -Message "Falling back to .NET file copy..." -Level 'Verbose'

                    # STRATEGY 3: .NET Directory copy (most granular control)
                    try {
                        Write-AppxLog -Message "Attempting copy with .NET APIs..." -Level 'Debug'

                        Copy-DirectoryNET -Source $SourcePath -Destination $DestinationPath

                        $copySucceeded = $true
                        $copyMethod = '.NET APIs'
                        Write-AppxLog -Message ".NET copy succeeded" -Level 'Debug'
                    }
                    catch {
                        Write-AppxLog -Message ".NET copy failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'
                        throw "All copy strategies failed. Last error: $_"
                    }
                }
            }

            if (-not $copySucceeded) {
                throw "Failed to copy source files"
            }

            Write-AppxLog -Message "Copy succeeded using: $copyMethod" -Level 'Verbose'

            # Remove existing signature files (invalid for repackaging)
            Remove-AppxSignatureFiles -DirectoryPath $DestinationPath

            # Verify critical files exist after copy
            $criticalFiles = @('AppxManifest.xml')
            foreach ($critical in @($criticalFiles)) {
                $criticalPath = [System.IO.Path]::Combine($DestinationPath, $critical)
                if (-not (Test-Path -LiteralPath $criticalPath)) {
                    throw "Critical file missing after copy: $critical (MakeAppx requires this file)"
                }
            }

            return [PSCustomObject]@{
                Success    = $true
                CopyMethod = $copyMethod
                Error      = $null
            }
        }
        catch {
            return [PSCustomObject]@{
                Success    = $false
                CopyMethod = $copyMethod
                Error      = $_.Exception.Message
            }
        }
    }
}

<#
.SYNOPSIS
    Recursively copies a directory using .NET System.IO APIs with error tolerance.

.DESCRIPTION
    Internal helper for Copy-AppxSourceDirectory. Uses .NET APIs for the most
    granular control over file copy operations, skipping inaccessible files/directories
    rather than failing the entire operation.

.PARAMETER Source
    Source directory path.

.PARAMETER Destination
    Destination directory path.

.NOTES
    Author: DeltaGa
    Version: 2.0.1
#>

function Copy-DirectoryNET {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Source,

        [Parameter(Mandatory)]
        [string]$Destination
    )

    process {
        $sourceDir = [System.IO.DirectoryInfo]::new($Source)

        # Create destination directory
        if (-not [System.IO.Directory]::Exists($Destination)) {
            [void][System.IO.Directory]::CreateDirectory($Destination)
        }

        # Copy files with error tolerance
        foreach ($file in $sourceDir.GetFiles()) {
            try {
                $destFile = [System.IO.Path]::Combine($Destination, $file.Name)
                [System.IO.File]::Copy($file.FullName, $destFile, $true)
            }
            catch {
                Write-AppxLog -Message "  Skipping file (access denied): $($file.Name)" -Level 'Debug'
            }
        }

        # Copy subdirectories recursively
        foreach ($dir in $sourceDir.GetDirectories()) {
            try {
                $destDir = [System.IO.Path]::Combine($Destination, $dir.Name)
                Copy-DirectoryNET -Source $dir.FullName -Destination $destDir
            }
            catch {
                Write-AppxLog -Message "  Skipping directory (access denied): $($dir.Name)" -Level 'Debug'
            }
        }
    }
}

<#
.SYNOPSIS
    Removes signature files from a package directory for repackaging.

.DESCRIPTION
    Removes AppxSignature.p7x and AppxBlockMap.xml (loaded from PackageConfiguration)
    to prepare a directory for clean repackaging with MakeAppx.

.PARAMETER DirectoryPath
    Path to the directory containing signature files to remove.

.NOTES
    Author: DeltaGa
    Version: 2.0.1
#>

function Remove-AppxSignatureFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$DirectoryPath
    )

    process {
        # Load signature file list from configuration
        try {
            $pkgConfig = Get-AppxConfiguration -ConfigName 'PackageConfiguration'
            $signatureFiles = $pkgConfig.signatureFiles.files
        }
        catch {
            Write-AppxLog -Message "Failed to load signature files config, using fallback: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
            $signatureFiles = @('AppxSignature.p7x', 'AppxBlockMap.xml')
        }

        foreach ($sigFile in @($signatureFiles)) {
            $sigPath = [System.IO.Path]::Combine($DirectoryPath, $sigFile)
            if (Test-Path -LiteralPath $sigPath) {
                Remove-Item -LiteralPath $sigPath -Force -ErrorAction SilentlyContinue
                Write-AppxLog -Message "Removed existing $sigFile for repackaging" -Level 'Debug'
            }
        }
    }
}
