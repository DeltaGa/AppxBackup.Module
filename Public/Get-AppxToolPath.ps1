<#
.SYNOPSIS
    Gets the file path to Windows SDK tools used for APPX operations.

.DESCRIPTION
    Utility function to locate SDK tools (MakeAppx, SignTool, etc.).
    Useful for troubleshooting and advanced scenarios.

.PARAMETER ToolName
    Name of the tool to locate.

.PARAMETER Refresh
    Forces a refresh of the tool cache.

.EXAMPLE
    Get-AppxToolPath -ToolName 'MakeAppx'
    
.OUTPUTS
    String (tool path) or $null if not found
#>

function Get-AppxToolPath {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory, Position = 0)]
        [ValidateSet('MakeAppx', 'SignTool', 'Certutil')]
        [string]$ToolName,

        [Parameter()]
        [switch]$Refresh
    )

    process {
        try {
            $available = Test-AppxToolAvailability -ToolName $ToolName -UpdateCache:$Refresh
            
            if ($available -and $script:ToolCache.ContainsKey($ToolName)) {
                return $script:ToolCache[$ToolName]
            }
            else {
                Write-Warning "Tool not found: $ToolName"
                return $null
            }
        }
        catch {
            Write-Error "Failed to get tool path: $_"
            return $null
        }
    }
}
