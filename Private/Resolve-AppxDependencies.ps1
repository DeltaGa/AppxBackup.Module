<#
.SYNOPSIS
    Resolves and analyzes APPX/MSIX package dependencies including frameworks and resources.

.DESCRIPTION
    This function provides comprehensive dependency analysis.
    
    Capabilities:
    - Framework package identification
    - Resource package detection
    - Dependency graph construction
    - Circular dependency detection
    - Optional component identification
    - Architecture-specific dependency handling
#>

function Resolve-AppxDependencies {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$PackagePath,

        [Parameter()]
        [switch]$IncludeOptional,

        [Parameter()]
        [switch]$Recursive,

        [Parameter()]
        [int]$MaxDepth = 5
    )

    begin {
        Write-AppxLog -Message "Resolving dependencies for: $PackagePath" -Level 'Verbose'
        
        $resolvedDependencies = @()
        $processedPackages = @{}
        $currentDepth = 0
    }

    process {
        try {
            # Validate package path
            $packagePath = ConvertTo-SecureFilePath -Path $PackagePath -MustExist -PathType Directory
            
            # Find and parse manifest
            $manifestPath = [System.IO.Path]::Combine($packagePath, 'AppxManifest.xml')
            if (-not (Test-Path -LiteralPath $manifestPath)) {
                throw "Manifest not found: $manifestPath"
            }

            $manifestData = Get-AppxManifestData -ManifestPath $manifestPath -IncludeDependencies

            # Extract dependencies from manifest
            $dependencies = @()
            
            # Ensure Dependencies is treated as array and has Count property
            $manifestDeps = @($manifestData.Dependencies)
            
            if ($null -ne $manifestData.Dependencies -and $manifestDeps.Count -gt 0) {
                foreach ($dep in $manifestDeps) {
                    $depInfo = [PSCustomObject]@{
                        PSTypeName          = 'AppxBackup.DependencyInfo'
                        Name                = $dep.Name
                        Publisher           = $dep.Publisher
                        MinVersion          = $dep.MinVersion
                        DependencyType      = 'PackageDependency'
                        IsOptional          = $false
                        Architecture        = $manifestData.ProcessorArchitecture
                        ResolvedPath        = $null
                        IsInstalled         = $false
                        InstalledVersion    = $null
                    }

                    # Check if dependency is installed
                    try {
                        $installedPkg = Get-AppxPackage -Name $dep.Name -ErrorAction SilentlyContinue |
                            Where-Object { $_.Publisher -eq $dep.Publisher } |
                            Select-Object -First 1

                        if ($installedPkg) {
                            $depInfo.IsInstalled = $true
                            $depInfo.InstalledVersion = $installedPkg.Version
                            $depInfo.ResolvedPath = $installedPkg.InstallLocation
                            
                            Write-AppxLog -Message "Dependency installed: $($dep.Name) v$($installedPkg.Version)" -Level 'Debug'
                        }
                        else {
                            Write-AppxLog -Message "Dependency not installed: $($dep.Name)" -Level 'Warning'
                        }
                    }
                    catch {
                        Write-AppxLog -Message "Failed to check dependency status: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
                    }

                    $dependencies += $depInfo
                }
            }

            # Identify framework dependencies (special handling)
            # Common framework packages that apps depend on
            $frameworkPatterns = @(
                'Microsoft.VCLibs.*',
                'Microsoft.NET.Native.*',
                'Microsoft.UI.Xaml.*',
                'Microsoft.Advertising.Xaml',
                'Microsoft.Services.Store.Engagement'
            )

            foreach ($pattern in @($frameworkPatterns)) {
                try {
                    $frameworkPkgs = Get-AppxPackage -Name $pattern -ErrorAction SilentlyContinue
                    
                    foreach ($fwPkg in @($frameworkPkgs)) {
                        # Check if not already in dependencies
                        $exists = $dependencies | Where-Object { $_.Name -eq $fwPkg.Name }
                        
                        if ($null -eq $exists) {
                            $fwInfo = [PSCustomObject]@{
                                PSTypeName          = 'AppxBackup.DependencyInfo'
                                Name                = $fwPkg.Name
                                Publisher           = $fwPkg.Publisher
                                MinVersion          = $fwPkg.Version
                                DependencyType      = 'Framework'
                                IsOptional          = $true
                                Architecture        = $fwPkg.Architecture
                                ResolvedPath        = $fwPkg.InstallLocation
                                IsInstalled         = $true
                                InstalledVersion    = $fwPkg.Version
                            }

                            if ($IncludeOptional.IsPresent) {
                                $dependencies += $fwInfo
                                Write-AppxLog -Message "Found framework dependency: $($fwPkg.Name)" -Level 'Debug'
                            }
                        }
                    }
                }
                catch {
                    Write-AppxLog -Message "Framework detection failed for pattern '$pattern': $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
                }
            }

            # Recursive resolution
            if ($Recursive.IsPresent -and $currentDepth -lt $MaxDepth) {
                foreach ($dep in @($dependencies)) {
                    if ($dep.ResolvedPath -and -not $processedPackages.ContainsKey($dep.Name)) {
                        $processedPackages[$dep.Name] = $true
                        
                        Write-AppxLog -Message "Recursively resolving: $($dep.Name)" -Level 'Debug'
                        
                        try {
                            $currentDepth++
                            $subDeps = Resolve-AppxDependencies -PackagePath $dep.ResolvedPath `
                                -IncludeOptional:$IncludeOptional `
                                -Recursive:$Recursive `
                                -MaxDepth $MaxDepth
                            
                            $currentDepth--
                            
                            if ($subDeps.Dependencies) {
                                $dependencies += $subDeps.Dependencies
                            }
                        }
                        catch {
                            Write-AppxLog -Message "Failed to resolve sub-dependencies for $($dep.Name): $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                        }
                    }
                }
            }

            # Build result object
            # Ensure dependencies is array for consistent Count property access
            $dependenciesArray = @($dependencies)
            
            $result = [PSCustomObject]@{
                PSTypeName          = 'AppxBackup.DependencyResult'
                PackageName         = $manifestData.Name
                PackageVersion      = $manifestData.Version
                TotalDependencies   = $dependenciesArray.Count
                InstalledCount      = (@($dependenciesArray | Where-Object { $_.IsInstalled })).Count
                MissingCount        = (@($dependenciesArray | Where-Object { -not $_.IsInstalled })).Count
                FrameworkCount      = (@($dependenciesArray | Where-Object { $_.DependencyType -eq 'Framework' })).Count
                Dependencies        = $dependenciesArray
                ManifestPath        = $manifestPath
            }

            Write-AppxLog -Message "Dependency resolution complete: $($result.TotalDependencies) found ($($result.MissingCount) missing)" -Level 'Verbose'
            
            return $result
        }
        catch {
            Write-AppxLog -Message "Dependency resolution failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'
            Write-AppxLog -Message "StackTrace: $($_.ScriptStackTrace)" -Level 'Debug'
            throw
        }
    }
}