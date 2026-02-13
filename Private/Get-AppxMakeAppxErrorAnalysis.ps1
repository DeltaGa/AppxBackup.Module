<#
.SYNOPSIS
    Analyzes MakeAppx.exe error output and provides actionable diagnostics.

.DESCRIPTION
    Matches MakeAppx error output against 15+ known error patterns to provide
    specific, actionable guidance for resolving packaging failures. Builds
    comprehensive error messages with error analysis, command details,
    and captured output.

.PARAMETER ExitCode
    The MakeAppx.exe exit code.

.PARAMETER StandardError
    Captured stderr output from MakeAppx.

.PARAMETER StandardOutput
    Captured stdout output from MakeAppx.

.PARAMETER MakeAppxPath
    Path to the MakeAppx.exe that was executed.

.PARAMETER Arguments
    The argument list passed to MakeAppx.

.PARAMETER SourcePath
    The source directory path used for packaging.

.PARAMETER OutputPath
    The output package path.

.OUTPUTS
    [string] Comprehensive error message with analysis and diagnostics.

.NOTES
    Author: DeltaGa
    Version: 2.0.2
#>

function Get-AppxMakeAppxErrorAnalysis {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [int]$ExitCode,

        [Parameter()]
        [string]$StandardError = '',

        [Parameter()]
        [string]$StandardOutput = '',

        [Parameter()]
        [string]$MakeAppxPath = '',

        [Parameter()]
        [string[]]$Arguments = @(),

        [Parameter()]
        [string]$SourcePath = '',

        [Parameter()]
        [string]$OutputPath = ''
    )

    process {
        $stderrContent = if ($StandardError) { $StandardError.Trim() } else { "" }
        $stdoutContent = if ($StandardOutput) { $StandardOutput.Trim() } else { "" }
        $fullError = "$stderrContent`n$stdoutContent"

        $errorAnalysis = @()

        # Common MakeAppx error patterns with actionable guidance
        if ($fullError -match 'Access is denied|ERROR_ACCESS_DENIED|0x80070005') {
            $errorAnalysis += "PERMISSION ISSUE: MakeAppx cannot access one or more files"
            $errorAnalysis += "  - Try running PowerShell as Administrator"
            $errorAnalysis += "  - Check if files in '$SourcePath' are locked by another process"
            $errorAnalysis += "  - Verify NTFS permissions on source and output directories"
        }

        if ($fullError -match 'The system cannot find the file specified|ERROR_FILE_NOT_FOUND|0x80070002') {
            $errorAnalysis += "FILE NOT FOUND: MakeAppx cannot locate required files"
            $errorAnalysis += "  - Verify AppxManifest.xml exists in source directory"
            $errorAnalysis += "  - Check if files referenced in manifest are present"
            $errorAnalysis += "  - Ensure logo/icon files specified in manifest exist"
        }

        if ($fullError -match 'The parameter is incorrect|ERROR_INVALID_PARAMETER|0x80070057') {
            $errorAnalysis += "INVALID PARAMETER: Command-line arguments may be malformed"
            $errorAnalysis += "  - Check for special characters in paths"
            $errorAnalysis += "  - Verify path lengths are under 260 characters"
        }

        if ($fullError -match 'invalid|malformed|corrupt|error in manifest|0x8007000d') {
            $errorAnalysis += "MANIFEST ERROR: The AppxManifest.xml file contains errors"
            $errorAnalysis += "  - Validate XML syntax in AppxManifest.xml"
            $errorAnalysis += "  - Check for missing required elements"
            $errorAnalysis += "  - Verify all file references are correct"
            $errorAnalysis += "  - Ensure namespace declarations are present"
        }

        if ($fullError -match '0x80080204|E_APPX_INVALID_MANIFEST') {
            $errorAnalysis += "INVALID MANIFEST SCHEMA: The manifest doesn't conform to the APPX schema"
            $errorAnalysis += "  - Check that all required manifest elements are present"
            $errorAnalysis += "  - Verify Identity, Properties, and Applications elements exist"
            $errorAnalysis += "  - Ensure Publisher matches certificate subject name"
        }

        if ($fullError -match '0x80080206|E_APPX_INVALID_BLOCKMAP') {
            $errorAnalysis += "BLOCKMAP ERROR: Issue with package block mapping"
            $errorAnalysis += "  - This usually means AppxBlockMap.xml is corrupt or inconsistent"
            $errorAnalysis += "  - Try deleting any existing AppxBlockMap.xml and let MakeAppx regenerate it"
        }

        if ($fullError -match 'ERROR_PATH_NOT_FOUND|0x80070003') {
            $errorAnalysis += "PATH NOT FOUND: One or more directory paths are invalid"
            $errorAnalysis += "  - Verify source directory exists: $SourcePath"
            $errorAnalysis += "  - Check output directory is valid: $OutputPath"
        }

        if ($fullError -match 'invalid manifest|manifest.*error') {
            $errorAnalysis += "MANIFEST ERROR: AppxManifest.xml contains errors"
            $errorAnalysis += "  - Manifest path: $([System.IO.Path]::Combine($SourcePath, 'AppxManifest.xml'))"
            $errorAnalysis += "  - Use Get-AppxManifestData to validate manifest structure"
        }

        if ($fullError -match '0x80080204|invalid signature') {
            $errorAnalysis += "SIGNATURE ERROR: Existing signature is invalid or corrupted"
            $errorAnalysis += "  - Remove AppxSignature.p7x from source if present"
            $errorAnalysis += "  - Package will be re-signed after creation"
        }

        if ($fullError -match 'already exists|file.*in use') {
            $errorAnalysis += "OUTPUT FILE CONFLICT: Target file exists or is locked"
            $errorAnalysis += "  - Output: $OutputPath"
            $errorAnalysis += "  - Try deleting the file manually or closing programs using it"
        }

        # If no specific pattern matched, provide generic help
        if ($errorAnalysis.Count -eq 0) {
            $errorAnalysis += "UNKNOWN ERROR: MakeAppx failed for an unrecognized reason"
            $errorAnalysis += "  - Check the log file for complete output"
            $errorAnalysis += "  - Try the .NET fallback by temporarily renaming MakeAppx.exe"
        }

        # Build comprehensive error message
        $errorMsg = "MakeAppx.exe failed with exit code $ExitCode`n`n"
        $errorMsg += "=== ERROR ANALYSIS ===`n"
        $errorMsg += ($errorAnalysis -join "`n") + "`n`n"
        $errorMsg += "=== COMMAND DETAILS ===`n"
        $errorMsg += "Executable: $MakeAppxPath`n"
        $errorMsg += "Arguments: $($Arguments -join ' ')`n"
        $errorMsg += "Source: $SourcePath`n"
        $errorMsg += "Output: $OutputPath`n`n"

        if ($stderrContent) {
            $errorMsg += "=== STDERR (Error Output) ===`n$stderrContent`n`n"
        }
        if ($stdoutContent) {
            $errorMsg += "=== STDOUT (Standard Output) ===`n$stdoutContent`n`n"
        }

        if (-not $stderrContent -and -not $stdoutContent) {
            $errorMsg += "=== NO ERROR OUTPUT CAPTURED ===`n"
            $errorMsg += "MakeAppx provided no error details. Possible causes:`n"
            $errorMsg += "  - Process terminated before writing output`n"
            $errorMsg += "  - Permissions issue preventing output capture`n"
            $errorMsg += "  - Try running PowerShell as Administrator`n`n"
        }

        $errorMsg += "Full details logged to: $env:TEMP\AppxBackup_$(Get-Date -Format 'yyyyMMdd').log"

        return $errorMsg
    }
}
