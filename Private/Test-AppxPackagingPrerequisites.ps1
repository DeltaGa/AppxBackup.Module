<#
.SYNOPSIS
    Runs pre-flight diagnostics before MakeAppx.exe execution.

.DESCRIPTION
    Validates common failure conditions before invoking MakeAppx:
    1. Source path character validation (problematic chars)
    2. Critical file readability (AppxManifest.xml, etc.)
    3. Output directory writability
    4. File count threshold warnings

    Throws on fatal conditions (unwritable output directory).
    Logs warnings for non-fatal concerns.

.PARAMETER SourcePath
    Full path to the source directory for packaging.

.PARAMETER OutputPath
    Full path where the output package will be created.

.OUTPUTS
    PSCustomObject with:
    - FileCount: [int] Number of files in source directory
    - Passed: [bool] Whether all critical checks passed
    - Warnings: [string[]] Non-fatal warning messages

.NOTES
    Author: DeltaGa
    Version: 2.0.1
#>

function Test-AppxPackagingPrerequisites {
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
        Write-AppxLog -Message "Running pre-flight diagnostics..." -Level 'Debug'

        $warnings = [System.Collections.ArrayList]::new()

        # Check 1: Verify source path doesn't contain problematic characters
        if ($SourcePath -match '[<>"|?*]') {
            Write-AppxLog -Message "WARNING: Source path contains characters that may cause issues: $SourcePath" -Level 'Warning'
            [void]$warnings.Add("Source path contains potentially problematic characters")
        }

        # Check 2: Verify we can read critical files
        $criticalFiles = @('AppxManifest.xml', 'AppxBlockMap.xml', 'AppxSignature.p7x')
        foreach ($file in @($criticalFiles)) {
            $testPath = [System.IO.Path]::Combine($SourcePath, $file)
            if (Test-Path -LiteralPath $testPath) {
                try {
                    $null = Get-Content -LiteralPath $testPath -TotalCount 1 -ErrorAction Stop
                    Write-AppxLog -Message "  [OK] Can read: $file" -Level 'Debug'
                }
                catch {
                    Write-AppxLog -Message "  [FAIL] Cannot read file: $file - $_" -Level 'Warning'
                    Write-AppxLog -Message "This may indicate permission or locking issues" -Level 'Warning'
                    [void]$warnings.Add("Cannot read $file - possible permission or lock issue")
                }
            }
        }

        # Check 3: Verify output directory is writable
        $outputDir = Split-Path -Path $OutputPath -Parent
        $testFile = [System.IO.Path]::Combine($outputDir, "AppxBackup_WriteTest_$(New-Guid).tmp")
        try {
            [void][System.IO.File]::WriteAllText($testFile, "test")
            Remove-Item -LiteralPath $testFile -Force -ErrorAction SilentlyContinue
            Write-AppxLog -Message "  [OK] Output directory is writable" -Level 'Debug'
        }
        catch {
            Write-AppxLog -Message "  [FAIL] Cannot write to output directory: $outputDir" -Level 'Warning'
            throw "Output directory is not writable: $_"
        }

        # Check 4: Warn if source has too many files (can cause timeout)
        $fileCount = 0
        try {
            $fileCount = (Get-ChildItem -LiteralPath $SourcePath -Recurse -File -ErrorAction SilentlyContinue | Measure-Object).Count
            Write-AppxLog -Message "  [INFO] Package contains $fileCount files" -Level 'Debug'
            if ($fileCount -gt 10000) {
                Write-AppxLog -Message "  [WARN] Large file count may slow packaging" -Level 'Warning'
                [void]$warnings.Add("Large file count ($fileCount) may slow packaging")
            }
        }
        catch {
            Write-AppxLog -Message "  [WARN] Could not count files: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
        }

        Write-AppxLog -Message "Pre-flight checks complete" -Level 'Debug'

        return [PSCustomObject]@{
            FileCount = $fileCount
            Passed    = $true
            Warnings  = $warnings.ToArray()
        }
    }
}
