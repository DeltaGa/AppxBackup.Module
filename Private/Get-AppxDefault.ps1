<#
.SYNOPSIS
    Retrieves default values from module configuration files.

.DESCRIPTION
    Helper function to easily access nested default values from configuration JSON files.
    Provides dot-notation path access to configuration values with optional fallback.
    
    Features:
    - Dot-notation path navigation (e.g., 'pathDefaults.maxPathLength')
    - Support for any configuration file (ModuleDefaults, ZipPackagingConfiguration, etc.)
    - Automatic type preservation (int, string, array, object)
    - Optional fallback values if path not found
    - Caching for performance
    - Null-safe navigation

.PARAMETER Path
    Dot-separated path to the configuration value.
    Example: 'timeoutDefaults.processExecutionDefaultSeconds'

.PARAMETER ConfigName
    Name of the configuration file to load (without .json extension).
    Default: 'ModuleDefaults'
    Valid values: 'ModuleDefaults', 'ZipPackagingConfiguration', 'PackageConfiguration', etc.

.PARAMETER Fallback
    Optional fallback value if path is not found in configuration.
    If not specified and path not found, returns $null.

.EXAMPLE
    $timeout = Get-AppxDefault 'timeoutDefaults.processExecutionDefaultSeconds'
    # Returns: 3600 from ModuleDefaults.json

.EXAMPLE
    $zipDir = Get-AppxDefault 'archiveStructure.packagesDirectory' 'ZipPackagingConfiguration' 'Packages'
    # Returns value from ZipPackagingConfiguration.json or 'Packages' as fallback

.EXAMPLE
    $maxSize = Get-AppxDefault 'logConfiguration.maxLogSizeMB' -Fallback 10
    # Returns: 10 (from config) or fallback if config fails

.OUTPUTS
    Object - The configuration value (preserves type: int, string, array, PSCustomObject)

.NOTES
    This is a convenience wrapper around Get-AppxConfiguration.
    All configuration values should be accessed through this function.
    
    Author: DeltaGa
    Version: 2.0.1
#>

function Get-AppxDefault {
    [CmdletBinding()]
    [OutputType([Object])]
    param(
        [Parameter(Mandatory, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter(Position = 1)]
        [string]$ConfigName = 'ModuleDefaults',

        [Parameter(Position = 2)]
        [Object]$Fallback = $null
    )

    process {
        try {
            # Load specified configuration (cached)
            $config = Get-AppxConfiguration -ConfigName $ConfigName
            
            # Navigate path using dot notation
            $pathParts = $Path.Split('.')
            $currentValue = $config
            
            foreach ($part in @($pathParts)) {
                if ($null -eq $currentValue) {
                    return $Fallback
                }
                
                # Check if it's a PSCustomObject with properties
                if ($currentValue.PSObject.Properties.Name -contains $part) {
                    $currentValue = $currentValue.$part
                }
                # Check if it's a hashtable
                elseif ($currentValue -is [System.Collections.IDictionary] -and $currentValue.ContainsKey($part)) {
                    $currentValue = $currentValue[$part]
                }
                else {
                    return $Fallback
                }
            }
            
            # Return the value (type preserved)
            return $currentValue
        }
        catch {
            return $Fallback
        }
    }
}