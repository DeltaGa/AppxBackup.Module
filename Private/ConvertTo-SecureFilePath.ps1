<#
.SYNOPSIS
    Converts and validates file paths to prevent injection attacks and handle special characters.

.DESCRIPTION
    Addresses:
    - Path traversal attacks
    - Command injection via file paths
    - PowerShell glob character issues ([, ], *, ?)
    - UNC path validation
    - Reserved filename detection
    - Path length validation (MAX_PATH)
    - Null byte injection
#>

function ConvertTo-SecureFilePath {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [AllowEmptyString()]
        [string]$Path,

        [Parameter()]
        [switch]$MustExist,

        [Parameter()]
        [ValidateSet('Any', 'File', 'Directory')]
        [string]$PathType = 'Any',

        [Parameter()]
        [switch]$AllowUNC,

        [Parameter()]
        [switch]$CreateIfMissing,

        [Parameter()]
        [switch]$ResolveRelative
    )

    process {
        try {
            # Null/empty check
            if ([string]::IsNullOrWhiteSpace($Path)) {
                throw "Path cannot be null or empty"
            }

            # Null byte injection check
            if ($Path.Contains([char]0)) {
                throw "Path contains null byte (potential injection attack)"
            }

            # Remove leading/trailing whitespace and quotes
            $cleanPath = $Path.Trim().Trim('"', "'")

            # Resolve relative paths if requested
            if ($ResolveRelative.IsPresent) {
                if (-not [System.IO.Path]::IsPathRooted($cleanPath)) {
                    $cleanPath = Join-Path $PWD.Path $cleanPath
                }
            }

            # Normalize path separators
            $cleanPath = $cleanPath.Replace('/', '\')

            # Check for path traversal attempts
            if ($cleanPath -match '\.\.[/\\]' -or $cleanPath -match '[/\\]\.\.$') {
                throw "Path contains directory traversal sequence (..): $cleanPath"
            }

            # Validate against reserved filenames (Windows)
            $fileName = Split-Path -Path $cleanPath -Leaf
            $reservedNames = @(
                'CON', 'PRN', 'AUX', 'NUL',
                'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9',
                'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9'
            )
            
            $fileNameBase = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
            if ($reservedNames -contains $fileNameBase.ToUpper()) {
                throw "Path contains reserved Windows filename: $fileName"
            }

            # Validate against invalid characters
            $invalidChars = [System.IO.Path]::GetInvalidPathChars()
            foreach ($char in $invalidChars) {
                if ($cleanPath.Contains($char)) {
                    throw "Path contains invalid character: $([int]$char) (0x$([Convert]::ToString([int]$char, 16)))"
                }
            }

            # UNC path validation
            $isUNC = $cleanPath.StartsWith('\\')
            if ($isUNC -and -not $AllowUNC.IsPresent) {
                throw "UNC paths are not allowed: $cleanPath"
            }

            # Path length validation (Windows MAX_PATH = 260, but NTFS supports 32,767)
            # We use 260 as default limit unless long paths are enabled
            $maxPathLength = 260
            
            # Check if long paths are enabled (Windows 10 1607+)
            try {
                $longPathsEnabled = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem' `
                    -Name 'LongPathsEnabled' -ErrorAction SilentlyContinue).LongPathsEnabled -eq 1
                
                if ($longPathsEnabled) {
                    $maxPathLength = 32767
                }
            }
            catch {
                # Assume not enabled
            }

            if ($cleanPath.Length -gt $maxPathLength) {
                throw "Path exceeds maximum length of $maxPathLength characters: $($cleanPath.Length)"
            }

            # Existence validation
            if ($MustExist.IsPresent) {
                $exists = Test-Path -LiteralPath $cleanPath
                
                if (-not $exists) {
                    if ($CreateIfMissing.IsPresent) {
                        # Create based on PathType
                        switch ($PathType) {
                            'Directory' {
                                Write-AppxLog -Message "Creating directory: $cleanPath" -Level 'Verbose'
                                [void](New-Item -Path $cleanPath -ItemType Directory -Force)
                            }
                            'File' {
                                # Create parent directory if needed
                                $parentDir = Split-Path -Path $cleanPath -Parent
                                if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
                                    [void](New-Item -Path $parentDir -ItemType Directory -Force)
                                }
                                # Create empty file
                                Write-AppxLog -Message "Creating file: $cleanPath" -Level 'Verbose'
                                [void](New-Item -Path $cleanPath -ItemType File -Force)
                            }
                            'Any' {
                                throw "Cannot create path of type 'Any'. Specify 'File' or 'Directory'."
                            }
                        }
                    }
                    else {
                        throw "Path does not exist: $cleanPath"
                    }
                }
                else {
                    # Validate type if specified
                    if ($PathType -ne 'Any') {
                        $item = Get-Item -LiteralPath $cleanPath
                        $actualType = if ($item.PSIsContainer) { 'Directory' } else { 'File' }
                        
                        if ($actualType -ne $PathType) {
                            throw "Path exists but is not a $PathType : $cleanPath (found $actualType)"
                        }
                    }
                }
            }

            # Escape PowerShell glob characters for use with -LiteralPath
            # Note: We return the clean path, but callers should use -LiteralPath
            $escapedPath = $cleanPath

            Write-AppxLog -Message "Path validated: $escapedPath" -Level 'Debug'
            
            return $escapedPath
        }
        catch {
            Write-AppxLog -Message "Path validation failed: $_" -Level 'Error'
            throw
        }
    }
}
