<#
.SYNOPSIS
    Discovers a Windows SDK tool by checking MakeAppx colocated directory, then PATH.

.DESCRIPTION
    Searches for a specified SDK tool executable using a two-tier strategy:
    1. Locates MakeAppx.exe and checks its parent directory for the requested tool
    2. Falls back to searching the system PATH

    This centralizes the SDK tool discovery logic that was previously duplicated
    across multiple functions in Backup-AppxPackage.

.PARAMETER ToolName
    Name of the tool executable to find (e.g., 'signtool.exe').

.OUTPUTS
    [string] Full path to the tool, or $null if not found.

.NOTES
    Author: DeltaGa
    Version: 2.0.2
#>

function Find-AppxSdkTool {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$ToolName
    )

    process {
        $toolPath = $null

        # Strategy 1: Find tool in same directory as MakeAppx.exe (SDK bin folder)
        try {
            $makeAppxCmd = Get-Command 'makeappx.exe' -ErrorAction SilentlyContinue |
                Select-Object -First 1 -ExpandProperty Source

            if ($makeAppxCmd) {
                $sdkDir = Split-Path -Path $makeAppxCmd -Parent
                $candidatePath = [System.IO.Path]::Combine($sdkDir, $ToolName)

                if (Test-Path -LiteralPath $candidatePath) {
                    $toolPath = $candidatePath
                    Write-AppxLog -Message "Found $ToolName in SDK directory: $toolPath" -Level 'Debug'
                    return $toolPath
                }
            }
        }
        catch {
            Write-AppxLog -Message "SDK directory search for $ToolName failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
        }

        # Strategy 2: Search PATH
        try {
            $pathCmd = Get-Command $ToolName -ErrorAction SilentlyContinue |
                Select-Object -First 1

            if ($pathCmd) {
                $toolPath = $pathCmd.Source
                Write-AppxLog -Message "Found $ToolName on PATH: $toolPath" -Level 'Debug'
                return $toolPath
            }
        }
        catch {
            Write-AppxLog -Message "PATH search for $ToolName failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
        }

        Write-AppxLog -Message "$ToolName not found via SDK directory or PATH" -Level 'Debug'
        return $null
    }
}
