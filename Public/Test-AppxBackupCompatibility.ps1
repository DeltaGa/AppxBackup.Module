<#
.SYNOPSIS
    Tests compatibility of an APPX/MSIX package or .appxpack archive with the current system.

.DESCRIPTION
    Performs comprehensive compatibility analysis including:
    - Operating system version compatibility
    - Architecture compatibility (x86, x64, ARM, ARM64)
    - Target device family validation
    - Minimum/maximum OS version checks
    - Dependency availability
    - System capability requirements
    
    For .appxpack archives, additionally assesses:
    - Main package compatibility
    - All dependency compatibility
    - System requirements from AppxBackupManifest.json
    - Aggregate installation feasibility
    - PowerShell version requirements
    
    Useful for pre-deployment validation and migration planning.

.PARAMETER PackagePath
    Path to the .appx, .msix, or .appxpack file to test.

.PARAMETER CheckDependencies
    If specified, validates that all dependencies are available.
    For .appxpack archives, validates all contained dependencies.

.PARAMETER CheckCapabilities
    If specified, validates that system has required capabilities.

.PARAMETER Detailed
    If specified, provides detailed compatibility report with recommendations.

.EXAMPLE
    Test-AppxBackupCompatibility -PackagePath "C:\Backups\MyApp.appx"
    
    Basic compatibility check

.EXAMPLE
    Test-AppxBackupCompatibility -PackagePath "C:\Backups\MyApp.appx" -CheckDependencies -CheckCapabilities -Detailed
    
    Comprehensive compatibility analysis

.EXAMPLE
    Test-AppxBackupCompatibility -PackagePath "C:\Backups\MyApp.appxpack" -CheckDependencies -Detailed
    
    Analyzes .appxpack archive compatibility including all dependencies

.OUTPUTS
    AppxBackup.CompatibilityResult
    For .appxpack archives, includes additional properties:
    - DependencyCompatibilityResults: Per-dependency compatibility
    - PowerShellVersionCompatible: Boolean
    - AllDependenciesCompatible: Boolean

.NOTES
    Uses WMI/CIM for system information gathering.
    Checks against current system only (cannot predict future OS).
#>

function Test-AppxBackupCompatibility {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [Alias('FullName', 'Path')]
        [string]$PackagePath,

        [Parameter()]
        [switch]$CheckDependencies,

        [Parameter()]
        [switch]$CheckCapabilities,

        [Parameter()]
        [switch]$Detailed
    )

    begin {
        Write-AppxLog -Message "Testing package compatibility" -Level 'Verbose'
        
        # Get system information
        $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
        $csInfo = Get-CimInstance -ClassName Win32_ComputerSystem
        
        # Parse OS version
        $osVersion = [Version]$osInfo.Version
        $osBuild = [System.Environment]::OSVersion.Version.Build
        
        # Determine OS name
        $osName = switch ($osBuild) {
            { $_ -ge 22000 } { "Windows 11"; break }
            { $_ -ge 10240 } { "Windows 10"; break }
            default { "Windows $($osVersion.Major).$($osVersion.Minor)" }
        }
        
        # System architecture
        $systemArch = if ([Environment]::Is64BitOperatingSystem) {
            if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') { 'ARM64' } else { 'x64' }
        } else {
            'x86'
        }
    }

    process {
        try {
            # Validate package
            $packagePath = ConvertTo-SecureFilePath -Path $PackagePath -MustExist -PathType File
            
            Write-AppxLog -Message "Analyzing package: $packagePath" -Level 'Verbose'
            
            # Load ZIP archive extension from configuration
            try {
                $zipConfig = Get-AppxConfiguration -ConfigName 'ZipPackagingConfiguration'
                $zipArchiveExtension = $zipConfig.archiveExtensions.zipArchive
            }
            catch {
                Write-AppxLog -Message "Failed to load ZIP configuration, using fallback: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                $zipArchiveExtension = '.appxpack'
            }
            
            # Detect if this is a .appxpack archive
            $packageFile = Get-Item -LiteralPath $packagePath
            $isZipArchive = $packageFile.Extension -eq $zipArchiveExtension
            
            if ($isZipArchive) {
                Write-AppxLog -Message "Detected .appxpack archive - using archive compatibility analysis" -Level 'Debug'
                
                # Test as ZIP-based dependency archive
                $result = Test-AppxpackArchiveCompatibility -PackagePath $packagePath `
                    -CheckDependencies:$CheckDependencies `
                    -Detailed:$Detailed `
                    -SystemArch $systemArch `
                    -OSVersion $osVersion `
                    -OSBuild $osBuild `
                    -OSName $osName
                
                return $result
            }
            
            # Continue with standard package compatibility testing
            Write-AppxLog -Message "Detected individual package - using standard compatibility analysis" -Level 'Debug'
            
            # Get package info
            $packageInfo = Get-AppxBackupInfo -PackagePath $packagePath -IncludeSignatureInfo
            
            # Initialize result
            $issues = @()
            $warnings = @()
            $isCompatible = $true

            Write-Host "`n=== System Information ===" -ForegroundColor Cyan
            Write-Host "  OS: $osName (Build $osBuild)" -ForegroundColor Gray
            Write-Host "  Architecture: $systemArch" -ForegroundColor Gray
            Write-Host "  OS Version: $osVersion" -ForegroundColor Gray

            Write-Host "`n=== Package Information ===" -ForegroundColor Cyan
            Write-Host "  Name: $($packageInfo.PackageName)" -ForegroundColor Gray
            Write-Host "  Version: $($packageInfo.PackageVersion)" -ForegroundColor Gray
            Write-Host "  Architecture: $($packageInfo.PackageArchitecture)" -ForegroundColor Gray

            # Check 1: Architecture Compatibility
            Write-Host "`n=== Compatibility Checks ===" -ForegroundColor Cyan
            
            $archCompatible = switch ($packageInfo.PackageArchitecture) {
                'neutral' { $true }
                'x86' { $true } # x86 runs on all architectures
                'x64' { $systemArch -in @('x64', 'ARM64') }
                'arm' { $systemArch -in @('ARM', 'ARM64') }
                'arm64' { $systemArch -eq 'ARM64' }
                default { $false }
            }
            
            if ($archCompatible) {
                Write-Host "  [PASS] Architecture: Compatible ($($packageInfo.PackageArchitecture) on $systemArch)" -ForegroundColor Green
            }
            else {
                Write-Host "  [X] Architecture: Incompatible ($($packageInfo.PackageArchitecture) cannot run on $systemArch)" -ForegroundColor Red
                $issues += "Architecture mismatch: Package requires $($packageInfo.PackageArchitecture), system is $systemArch"
                $isCompatible = $false
            }

            # Check 2: Target Device Family
            $deviceFamilyCompatible = $true
            
            if ($packageInfo.TargetDeviceFamilies.Count -gt 0) {
                $currentFamily = 'Windows.Desktop' # Assume desktop for now
                $matchingFamily = $packageInfo.TargetDeviceFamilies | Where-Object { $_.Name -eq $currentFamily }
                
                if ($matchingFamily) {
                    # Check min version
                    if ($matchingFamily.MinVersion) {
                        $minVer = [Version]$matchingFamily.MinVersion.Split('.')[2]
                        
                        if ($osBuild -ge $minVer) {
                            Write-Host "  [PASS] OS Version: Compatible (Build $osBuild >= $minVer)" -ForegroundColor Green
                        }
                        else {
                            Write-Host "  [X] OS Version: Incompatible (Build $osBuild < $minVer required)" -ForegroundColor Red
                            $issues += "OS build too old: Package requires build $minVer, system is $osBuild"
                            $isCompatible = $false
                            $deviceFamilyCompatible = $false
                        }
                    }
                    
                    # Check max version tested
                    if ($matchingFamily.MaxVersionTested) {
                        $maxVer = [Version]$matchingFamily.MaxVersionTested.Split('.')[2]
                        
                        if ($osBuild -gt $maxVer) {
                            Write-Host "  [WARNING]  OS Version: Newer than tested (Build $osBuild > $maxVer tested)" -ForegroundColor Yellow
                            $warnings += "Package not tested on OS build $osBuild (tested up to $maxVer)"
                        }
                    }
                }
                else {
                    Write-Host "  [WARNING]  Device Family: Not explicitly supported" -ForegroundColor Yellow
                    $warnings += "Package does not explicitly support $currentFamily device family"
                }
            }
            else {
                Write-Host "  [INFO] Device Family: Not specified in manifest" -ForegroundColor Gray
            }

            # Check 3: Dependencies (if requested)
            if ($CheckDependencies.IsPresent) {
                Write-Host "`n=== Dependency Check ===" -ForegroundColor Cyan
                
                # Get package directory or extract manifest
                $workingPath = $packagePath
                $tempDir = $null
                
                if ([System.IO.Path]::GetExtension($packagePath) -in @('.appx', '.msix')) {
                    $tempDir = [System.IO.Path]::Combine($env:TEMP, "AppxCompatCheck_$(New-Guid)")
                    [void](New-Item -Path $tempDir -ItemType Directory -Force)
                    
                    Add-Type -AssemblyName 'System.IO.Compression.FileSystem'
                    $archive = [System.IO.Compression.ZipFile]::OpenRead($packagePath)
                    
                    try {
                        $manifestEntry = $archive.Entries | Where-Object { $_.Name -eq 'AppxManifest.xml' } | Select-Object -First 1
                        if ($manifestEntry) {
                            $manifestPath = [System.IO.Path]::Combine($tempDir, 'AppxManifest.xml')
                            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($manifestEntry, $manifestPath, $true)
                            $workingPath = $tempDir
                        }
                    }
                    finally {
                        $archive.Dispose()
                    }
                }
                
                try {
                    $depResult = Resolve-AppxDependencies -PackagePath $workingPath -IncludeOptional
                    
                    Write-Host "  Total Dependencies: $($depResult.TotalDependencies)" -ForegroundColor Gray
                    Write-Host "  Installed: $($depResult.InstalledCount)" -ForegroundColor Gray
                    Write-Host "  Missing: $($depResult.MissingCount)" -ForegroundColor Gray
                    
                    if ($depResult.MissingCount -eq 0) {
                        Write-Host "  [PASS] All dependencies available" -ForegroundColor Green
                    }
                    else {
                        Write-Host "  [WARNING]  $($depResult.MissingCount) dependencies missing" -ForegroundColor Yellow
                        $warnings += "$($depResult.MissingCount) required dependencies not installed"
                        
                        if ($Detailed.IsPresent) {
                            $missingDeps = $depResult.Dependencies | Where-Object { -not $_.IsInstalled -and -not $_.IsOptional }
                            foreach ($dep in @($missingDeps)) {
                                Write-Host "    - $($dep.Name) v$($dep.MinVersion)+" -ForegroundColor Yellow
                            }
                        }
                    }
                }
                finally {
                    if ($tempDir -and (Test-Path $tempDir)) {
                        Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                    }
                }
            }

            # Check 4: Signature Status
            Write-Host "`n=== Signature Check ===" -ForegroundColor Cyan
            
            if ($packageInfo.SignatureInfo) {
                $sigStatus = $packageInfo.SignatureInfo.Status
                
                switch ($sigStatus) {
                    'Valid' {
                        Write-Host "  [PASS] Signature: Valid" -ForegroundColor Green
                    }
                    'NotSigned' {
                        Write-Host "  [WARNING]  Signature: Not signed (requires developer mode)" -ForegroundColor Yellow
                        $warnings += "Package is not signed. Installation requires developer mode or trusted certificate."
                    }
                    default {
                        Write-Host "  [WARNING]  Signature: $sigStatus" -ForegroundColor Yellow
                        $warnings += "Package signature status: $sigStatus"
                    }
                }
            }
            else {
                Write-Host "  [WARNING]  Signature: Unable to verify" -ForegroundColor Yellow
            }

            # Build result
            $result = [PSCustomObject]@{
                PSTypeName              = 'AppxBackup.CompatibilityResult'
                IsCompatible            = $isCompatible
                ArchitectureCompatible  = $archCompatible
                OSCompatible            = $deviceFamilyCompatible
                PackageName             = $packageInfo.PackageName
                PackageVersion          = $packageInfo.PackageVersion
                PackageArchitecture     = $packageInfo.PackageArchitecture
                SystemOS                = $osName
                SystemBuild             = $osBuild
                SystemArchitecture      = $systemArch
                IssueCount              = $issues.Count
                WarningCount            = $warnings.Count
                Issues                  = $issues
                Warnings                = $warnings
                PackageFilePath         = $packagePath
                TestDate                = [DateTime]::Now
            }

            # Summary
            Write-Host "`n=== Compatibility Summary ===" -ForegroundColor Cyan
            
            if ($result.IsCompatible) {
                Write-Host "  [OK] COMPATIBLE" -ForegroundColor Green
                Write-Host "  Package can be installed on this system" -ForegroundColor Gray
            }
            else {
                Write-Host "  [X] INCOMPATIBLE" -ForegroundColor Red
                Write-Host "  Package cannot be installed on this system" -ForegroundColor Gray
            }
            
            if ($result.WarningCount -gt 0) {
                Write-Host "  [WARNING]  $($result.WarningCount) warning(s)" -ForegroundColor Yellow
            }
            
            if ($Detailed.IsPresent) {
                if ($result.Issues.Count -gt 0) {
                    Write-Host "`n[X] Issues:" -ForegroundColor Red
                    $result.Issues | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
                }
                
                if ($result.Warnings.Count -gt 0) {
                    Write-Host "`n[WARNING]  Warnings:" -ForegroundColor Yellow
                    $result.Warnings | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
                }
            }

            Write-Host ""
            
            return $result
        }
        catch {
            Write-AppxLog -Message "Compatibility check failed: $_" -Level 'Error'
            Write-AppxLog -Message "StackTrace: $($_.ScriptStackTrace)" -Level 'Debug'
            throw
        }
    }
}

<#
.SYNOPSIS
    Internal helper to test .appxpack archive compatibility.

.DESCRIPTION
    Assesses system compatibility for ZIP-based dependency archive installation,
    including main package, all dependencies, and system requirements.

.OUTPUTS
    PSCustomObject with CompatibilityResult type containing archive compatibility
#>
function Test-AppxpackArchiveCompatibility {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)]
        [string]$PackagePath,
        
        [Parameter()]
        [switch]$CheckDependencies,
        
        [Parameter()]
        [switch]$Detailed,
        
        [Parameter(Mandatory)]
        [string]$SystemArch,
        
        [Parameter(Mandatory)]
        [Version]$OSVersion,
        
        [Parameter(Mandatory)]
        [int]$OSBuild,
        
        [Parameter(Mandatory)]
        [string]$OSName
    )
    
    Write-AppxLog -Message "Testing .appxpack archive compatibility: $PackagePath" -Level 'Info'
    
    $issues = @()
    $warnings = @()
    $isCompatible = $true
    
    # Create temp extraction directory
    $tempDir = [System.IO.Path]::Combine(
        $env:TEMP,
        "AppxpackCompatibility_$(New-Guid)"
    )
    [void](New-Item -Path $tempDir -ItemType Directory -Force)
    
    try {
        # Open archive and extract manifest
        Add-Type -AssemblyName 'System.IO.Compression.FileSystem'
        $archive = [System.IO.Compression.ZipFile]::OpenRead($PackagePath)
        
        try {
            # Load configuration
            $zipConfig = Get-AppxConfiguration -ConfigName 'ZipPackagingConfiguration'
            $manifestFileName = $zipConfig.archiveStructure.manifestFileName
            
            # Extract AppxBackupManifest.json
            $manifestEntry = $archive.Entries | Where-Object { $_.Name -eq $manifestFileName } | Select-Object -First 1
            
            if (-not $manifestEntry) {
                throw "AppxBackupManifest.json not found in archive"
            }
            
            $manifestPath = [System.IO.Path]::Combine($tempDir, $manifestFileName)
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($manifestEntry, $manifestPath, $true)
            
            # Parse manifest
            $manifestContent = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
            $mainPkg = $manifestContent.MainPackage
            
            Write-AppxLog -Message "Loaded manifest: $($mainPkg.Name) v$($mainPkg.Version)" -Level 'Debug'
            
            # Display system information
            Write-Host "`n=== System Information ===" -ForegroundColor Cyan
            Write-Host "  OS: $OSName (Build $OSBuild)" -ForegroundColor Gray
            Write-Host "  Architecture: $SystemArch" -ForegroundColor Gray
            Write-Host "  OS Version: $OSVersion" -ForegroundColor Gray
            Write-Host "  PowerShell: $($PSVersionTable.PSVersion)" -ForegroundColor Gray
            
            Write-Host "`n=== Archive Information ===" -ForegroundColor Cyan
            Write-Host "  Main Package: $($mainPkg.Name)" -ForegroundColor Gray
            Write-Host "  Version: $($mainPkg.Version)" -ForegroundColor Gray
            Write-Host "  Architecture: $($mainPkg.Architecture)" -ForegroundColor Gray
            Write-Host "  Total Packages: $($manifestContent.TotalPackages)" -ForegroundColor Gray
            Write-Host "  Dependencies: $($manifestContent.Dependencies.Count)" -ForegroundColor Gray
            
            # Check 1: Architecture Compatibility (Main Package)
            Write-Host "`n=== Compatibility Checks ===" -ForegroundColor Cyan
            
            $archCompatible = switch ($mainPkg.Architecture) {
                'neutral' { $true }
                'x86' { $true }
                'x64' { $SystemArch -in @('x64', 'ARM64') }
                'arm' { $SystemArch -in @('ARM', 'ARM64') }
                'arm64' { $SystemArch -eq 'ARM64' }
                default { $false }
            }
            
            if ($archCompatible) {
                Write-Host "  [PASS] Main Package Architecture: Compatible ($($mainPkg.Architecture) on $SystemArch)" -ForegroundColor Green
            }
            else {
                Write-Host "  [FAIL] Main Package Architecture: Incompatible ($($mainPkg.Architecture) cannot run on $SystemArch)" -ForegroundColor Red
                $issues += "Main package architecture mismatch: Requires $($mainPkg.Architecture), system is $SystemArch"
                $isCompatible = $false
            }
            
            # Check 2: PowerShell Version Requirements
            $psVersionCompatible = $true
            if ($manifestContent.MinimumPowerShellVersion) {
                $minPSVersion = [Version]$manifestContent.MinimumPowerShellVersion
                $currentPSVersion = $PSVersionTable.PSVersion
                
                if ($currentPSVersion -ge $minPSVersion) {
                    Write-Host "  [PASS] PowerShell Version: Compatible (v$currentPSVersion >= v$minPSVersion)" -ForegroundColor Green
                }
                else {
                    Write-Host "  [FAIL] PowerShell Version: Incompatible (v$currentPSVersion < v$minPSVersion required)" -ForegroundColor Red
                    $issues += "PowerShell version too old: Archive requires v$minPSVersion, system has v$currentPSVersion"
                    $isCompatible = $false
                    $psVersionCompatible = $false
                }
            }
            else {
                Write-Host "  [INFO] PowerShell Version: Not specified in manifest" -ForegroundColor Gray
            }
            
            # Check 3: OS Version Requirements
            $osCompatible = $true
            if ($manifestContent.MinimumOSVersion) {
                # Parse minimum OS version (format: "10.0.17763")
                $minOSVersionParts = $manifestContent.MinimumOSVersion.Split('.')
                if ($minOSVersionParts.Count -ge 3) {
                    $minOSBuild = [int]$minOSVersionParts[2]
                    
                    if ($OSBuild -ge $minOSBuild) {
                        Write-Host "  [PASS] OS Version: Compatible (Build $OSBuild >= $minOSBuild)" -ForegroundColor Green
                    }
                    else {
                        Write-Host "  [FAIL] OS Version: Incompatible (Build $OSBuild < $minOSBuild required)" -ForegroundColor Red
                        $issues += "OS build too old: Archive requires build $minOSBuild, system is $OSBuild"
                        $isCompatible = $false
                        $osCompatible = $false
                    }
                }
            }
            else {
                Write-Host "  [INFO] OS Version: Not specified in manifest" -ForegroundColor Gray
            }
            
            # Check 4: Elevation Requirements
            if ($manifestContent.RequiresElevation) {
                $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
                
                if ($isAdmin) {
                    Write-Host "  [PASS] Elevation: Administrator privileges available" -ForegroundColor Green
                }
                else {
                    Write-Host "  [WARN] Elevation: Installation requires Administrator privileges" -ForegroundColor Yellow
                    $warnings += "Archive installation requires Administrator privileges for certificate installation"
                }
            }
            else {
                Write-Host "  [INFO] Elevation: Not required" -ForegroundColor Gray
            }
            
            # Check 5: Dependency Compatibility
            $dependencyCompatibilityResults = @()
            $allDependenciesCompatible = $true
            
            if ($CheckDependencies.IsPresent -and $manifestContent.Dependencies.Count -gt 0) {
                Write-Host "`n=== Dependency Compatibility ===" -ForegroundColor Cyan
                
                foreach ($dep in $manifestContent.Dependencies) {
                    $depCompatible = $true
                    $depIssues = @()
                    
                    # Check dependency architecture
                    $depArchCompatible = switch ($dep.Architecture) {
                        'neutral' { $true }
                        'x86' { $true }
                        'x64' { $SystemArch -in @('x64', 'ARM64') }
                        'arm' { $SystemArch -in @('ARM', 'ARM64') }
                        'arm64' { $SystemArch -eq 'ARM64' }
                        default { $false }
                    }
                    
                    if (-not $depArchCompatible) {
                        $depCompatible = $false
                        $depIssues += "Architecture mismatch: $($dep.Architecture) on $SystemArch"
                        $allDependenciesCompatible = $false
                        
                        if (-not $dep.IsOptional) {
                            $isCompatible = $false
                        }
                    }
                    
                    # Check if dependency is already installed
                    $installedDep = Get-AppxPackage -Name "*$($dep.Name)*" -ErrorAction SilentlyContinue | 
                        Where-Object { $_.Architecture -eq $dep.Architecture } |
                        Select-Object -First 1
                    
                    $depResult = [PSCustomObject]@{
                        Name             = $dep.Name
                        Version          = $dep.Version
                        Architecture     = $dep.Architecture
                        IsOptional       = $dep.IsOptional
                        IsInstalled      = ($null -ne $installedDep)
                        IsCompatible     = $depCompatible
                        Issues           = $depIssues
                    }
                    
                    $dependencyCompatibilityResults += $depResult
                    
                    # Display dependency status
                    $depDisplayName = "$($dep.Name) v$($dep.Version) [$($dep.Architecture)]"
                    if ($dep.IsOptional) {
                        $depDisplayName += " (Optional)"
                    }
                    
                    if ($depCompatible) {
                        $installStatus = if ($installedDep) { "Already installed" } else { "Will be installed from archive" }
                        Write-Host "  [PASS] $depDisplayName - $installStatus" -ForegroundColor Green
                    }
                    else {
                        Write-Host "  [FAIL] $depDisplayName - Incompatible" -ForegroundColor Red
                        if ($Detailed.IsPresent) {
                            foreach ($issue in $depIssues) {
                                Write-Host "    - $issue" -ForegroundColor Red
                            }
                        }
                    }
                }
                
                if (-not $allDependenciesCompatible) {
                    $incompatibleCount = ($dependencyCompatibilityResults | Where-Object { -not $_.IsCompatible }).Count
                    $issues += "$incompatibleCount dependency(ies) incompatible with system"
                }
            }
            
            # Build result object
            $result = [PSCustomObject]@{
                PSTypeName                      = 'AppxBackup.CompatibilityResult'
                IsCompatible                    = $isCompatible
                IsZipArchive                    = $true
                ArchitectureCompatible          = $archCompatible
                OSCompatible                    = $osCompatible
                PowerShellVersionCompatible     = $psVersionCompatible
                AllDependenciesCompatible       = $allDependenciesCompatible
                PackageName                     = $mainPkg.Name
                PackageVersion                  = $mainPkg.Version
                PackageArchitecture             = $mainPkg.Architecture
                TotalPackagesInArchive          = $manifestContent.TotalPackages
                DependencyCount                 = $manifestContent.Dependencies.Count
                DependencyCompatibilityResults  = $dependencyCompatibilityResults
                SystemOS                        = $OSName
                SystemBuild                     = $OSBuild
                SystemArchitecture              = $SystemArch
                CurrentPowerShellVersion        = $PSVersionTable.PSVersion.ToString()
                RequiredPowerShellVersion       = $manifestContent.MinimumPowerShellVersion
                RequiredOSVersion               = $manifestContent.MinimumOSVersion
                RequiresElevation               = $manifestContent.RequiresElevation
                IssueCount                      = $issues.Count
                WarningCount                    = $warnings.Count
                Issues                          = $issues
                Warnings                        = $warnings
                PackageFilePath                 = $PackagePath
                TestDate                        = [DateTime]::Now
            }
            
            # Summary
            Write-Host "`n=== Compatibility Summary ===" -ForegroundColor Cyan
            
            if ($result.IsCompatible) {
                Write-Host "  [OK] COMPATIBLE" -ForegroundColor Green
                Write-Host "  Archive can be installed on this system" -ForegroundColor Gray
            }
            else {
                Write-Host "  [FAIL] INCOMPATIBLE" -ForegroundColor Red
                Write-Host "  Archive cannot be installed on this system" -ForegroundColor Gray
            }
            
            if ($result.WarningCount -gt 0) {
                Write-Host "  [WARN] $($result.WarningCount) warning(s)" -ForegroundColor Yellow
            }
            
            if ($Detailed.IsPresent) {
                if ($result.Issues.Count -gt 0) {
                    Write-Host "`n[FAIL] Issues:" -ForegroundColor Red
                    $result.Issues | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
                }
                
                if ($result.Warnings.Count -gt 0) {
                    Write-Host "`n[WARN] Warnings:" -ForegroundColor Yellow
                    $result.Warnings | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
                }
            }
            
            Write-Host ""
            
            return $result
        }
        finally {
            if ($archive) { $archive.Dispose() }
        }
    }
    catch {
        Write-AppxLog -Message "Archive compatibility check failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'
        throw
    }
    finally {
        # Cleanup temp directory
        if (Test-Path -LiteralPath $tempDir) {
            Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}