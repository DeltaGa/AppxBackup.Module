<#
.SYNOPSIS
    Tests if an APPX package architecture is compatible with the current system.

.DESCRIPTION
    Evaluates architecture compatibility using the Windows compatibility matrix:
    - neutral: Compatible with all architectures
    - x86: Compatible with all architectures (WoW64 emulation)
    - x64: Compatible with x64 and ARM64 (x64 emulation on ARM)
    - arm: Compatible with ARM and ARM64
    - arm64: Compatible only with ARM64

.PARAMETER PackageArchitecture
    The architecture specified in the package manifest (neutral, x86, x64, arm, arm64).

.PARAMETER SystemArchitecture
    The current system architecture. If not specified, auto-detected.

.OUTPUTS
    PSCustomObject with:
    - Compatible: [bool] Whether the architecture is compatible
    - PackageArchitecture: [string] The package architecture
    - SystemArchitecture: [string] The system architecture
    - EmulationNote: [string] Note about emulation if applicable

.NOTES
    Author: DeltaGa
    Version: 2.0.2
#>

function Test-AppxArchitectureCompatibility {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$PackageArchitecture,

        [Parameter()]
        [string]$SystemArchitecture
    )

    process {
        # Auto-detect system architecture if not provided
        if (-not $SystemArchitecture) {
            if ([Environment]::Is64BitOperatingSystem) {
                $SystemArchitecture = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') { 'ARM64' } else { 'x64' }
            }
            else {
                $SystemArchitecture = 'x86'
            }
        }

        $emulationNote = $null

        $compatible = switch ($PackageArchitecture.ToLower()) {
            'neutral' {
                $emulationNote = 'Architecture-neutral package runs on all platforms'
                $true
            }
            'x86' {
                if ($SystemArchitecture -ne 'x86') {
                    $emulationNote = 'x86 package will run under WoW64 emulation'
                }
                $true  # x86 runs on all architectures
            }
            'x64' {
                if ($SystemArchitecture -eq 'ARM64') {
                    $emulationNote = 'x64 package will run under ARM64 x64 emulation'
                }
                $SystemArchitecture -in @('x64', 'ARM64')
            }
            'arm' {
                $SystemArchitecture -in @('ARM', 'ARM64')
            }
            'arm64' {
                $SystemArchitecture -eq 'ARM64'
            }
            default { $false }
        }

        return [PSCustomObject]@{
            Compatible          = $compatible
            PackageArchitecture = $PackageArchitecture
            SystemArchitecture  = $SystemArchitecture
            EmulationNote       = $emulationNote
        }
    }
}
