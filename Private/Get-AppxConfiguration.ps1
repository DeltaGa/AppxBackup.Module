<#
.SYNOPSIS
    Loads and caches external configuration files (JSON databases).

.DESCRIPTION
    Centralized configuration loader that reads JSON files from the Config directory
    and caches them in module scope for performance.
    
    Features:
    - Automatic path resolution relative to module root
    - In-memory caching for performance
    - Error handling with graceful fallbacks
    - Validation of loaded data
    - Support for all module configuration files

.PARAMETER ConfigName
     Name of the configuration file to load (without .json extension).
     Valid values: 'ToolConfiguration', 'WindowsReservedNames', 
                   'PackageConfiguration', 'ModuleDefaults', 'ZipPackagingConfiguration'

.PARAMETER Reload
    If specified, forces reload from disk even if cached.

.OUTPUTS
    PSCustomObject - Parsed JSON configuration

.NOTES
    This is a critical infrastructure function used by multiple modules.
    Configuration files are located in: <ModuleRoot>/Config/*.json
    
    Author: DeltaGa
    Version: 2.0.1
#>

function Get-AppxConfiguration {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory, Position = 0)]
         [ValidateSet('ToolConfiguration', 'WindowsReservedNames', 'PackageConfiguration', 'ModuleDefaults', 'ZipPackagingConfiguration')]
         [string]$ConfigName,

        [Parameter()]
        [switch]$Reload
    )

    begin {
        # Initialize configuration cache if not exists
        if (-not $script:ConfigCache) {
            $script:ConfigCache = @{}
        }
    }

    process {
        try {
            # Check cache first (unless Reload is specified)
            if (-not $Reload.IsPresent -and $script:ConfigCache.ContainsKey($ConfigName)) {
                return $script:ConfigCache[$ConfigName]
            }

            # Determine module root directory
            # This function is in Private/, so module root is parent of parent
            $moduleRoot = Split-Path -Path $PSScriptRoot -Parent
            $configPath = [System.IO.Path]::Combine($moduleRoot, 'Config', "$ConfigName.json")

            # Validate file exists
            if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) {
                throw "Configuration file not found: $configPath"
            }

            # Read and parse JSON
            $jsonContent = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 -ErrorAction Stop
            
            if ([string]::IsNullOrWhiteSpace($jsonContent)) {
                throw "Configuration file is empty: $configPath"
            }

            # Parse JSON with error handling
            try {
                $config = $jsonContent | ConvertFrom-Json -ErrorAction Stop
            }
            catch {
                throw "Failed to parse JSON in ${configPath}: $_"
            }

            # Validate basic structure
            if ($null -eq $config) {
                throw "Configuration parsing returned null: $configPath"
            }

            # Configuration-specific validation
            switch ($ConfigName) {
                'ToolConfiguration' {
                    if (-not $config.toolConfigurations) {
                        throw "ToolConfiguration missing 'toolConfigurations' property"
                    }
                    if (-not $config.defaultConfiguration) {
                        throw "ToolConfiguration missing 'defaultConfiguration' property"
                    }
                }
                
                'WindowsReservedNames' {
                    if (-not $config.reservedNames) {
                        throw "WindowsReservedNames missing 'reservedNames' property"
                    }
                    if ($config.reservedNames.Count -eq 0) {
                        throw "WindowsReservedNames has empty 'reservedNames' array"
                    }
                }
            }

            # Cache the configuration
             $script:ConfigCache[$ConfigName] = $config
            
            return $config
        }
        catch {
            $errorMsg = "Failed to load configuration '$ConfigName': $_"
            throw $errorMsg
        }
    }
}