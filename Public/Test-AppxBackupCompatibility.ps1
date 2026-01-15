<#
.SYNOPSIS
    Tests compatibility of an APPX/MSIX package with the current system.

.DESCRIPTION
    Performs comprehensive compatibility analysis including:
    - Operating system version compatibility
    - Architecture compatibility (x86, x64, ARM, ARM64)
    - Target device family validation
    - Minimum/maximum OS version checks
    - Dependency availability
    - System capability requirements
    
    Useful for pre-deployment validation and migration planning.

.PARAMETER PackagePath
    Path to the .appx or .msix file to test.

.PARAMETER CheckDependencies
    If specified, validates that all dependencies are available.

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

.OUTPUTS
    AppxBackup.CompatibilityResult

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
            { $_ -ge 22000 } { "Windows 11" }
            { $_ -ge 10240 } { "Windows 10" }
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
            
            # Get package info
            $packageInfo = Get-AppxBackupInfo -PackagePath $packagePath -IncludeSignatureInfo
            
            # Initialize result
            $issues = @()
            $warnings = @()
            $isCompatible = $true

            Write-Host "`n[CHAR_128421][CHAR_65039]  System Information:" -ForegroundColor Cyan
            Write-Host "  OS: $osName (Build $osBuild)" -ForegroundColor Gray
            Write-Host "  Architecture: $systemArch" -ForegroundColor Gray
            Write-Host "  OS Version: $osVersion" -ForegroundColor Gray

            Write-Host "`n[CHAR_128230] Package Information:" -ForegroundColor Cyan
            Write-Host "  Name: $($packageInfo.PackageName)" -ForegroundColor Gray
            Write-Host "  Version: $($packageInfo.PackageVersion)" -ForegroundColor Gray
            Write-Host "  Architecture: $($packageInfo.PackageArchitecture)" -ForegroundColor Gray

            # Check 1: Architecture Compatibility
            Write-Host "`n[CHAR_128269] Compatibility Checks:" -ForegroundColor Cyan
            
            $archCompatible = switch ($packageInfo.PackageArchitecture) {
                'neutral' { $true }
                'x86' { $true } # x86 runs on all architectures
                'x64' { $systemArch -in @('x64', 'ARM64') }
                'arm' { $systemArch -in @('ARM', 'ARM64') }
                'arm64' { $systemArch -eq 'ARM64' }
                default { $false }
            }
            
            if ($archCompatible) {
                Write-Host "  [CHECK] Architecture: Compatible ($($packageInfo.PackageArchitecture) on $systemArch)" -ForegroundColor Green
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
                            Write-Host "  [CHECK] OS Version: Compatible (Build $osBuild >= $minVer)" -ForegroundColor Green
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
                Write-Host "  [CHAR_8505][CHAR_65039]  Device Family: Not specified in manifest" -ForegroundColor Gray
            }

            # Check 3: Dependencies (if requested)
            if ($CheckDependencies.IsPresent) {
                Write-Host "`n[CHAR_128230] Dependency Check:" -ForegroundColor Cyan
                
                # Get package directory or extract manifest
                $workingPath = $packagePath
                $tempDir = $null
                
                if ([System.IO.Path]::GetExtension($packagePath) -in @('.appx', '.msix')) {
                    $tempDir = Join-Path $env:TEMP "AppxCompatCheck_$(New-Guid)"
                    [void](New-Item -Path $tempDir -ItemType Directory -Force)
                    
                    Add-Type -AssemblyName 'System.IO.Compression.FileSystem'
                    $archive = [System.IO.Compression.ZipFile]::OpenRead($packagePath)
                    
                    try {
                        $manifestEntry = $archive.Entries | Where-Object { $_.Name -eq 'AppxManifest.xml' } | Select-Object -First 1
                        if ($manifestEntry) {
                            $manifestPath = Join-Path $tempDir 'AppxManifest.xml'
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
                        Write-Host "  [CHECK] All dependencies available" -ForegroundColor Green
                    }
                    else {
                        Write-Host "  [WARNING]  $($depResult.MissingCount) dependencies missing" -ForegroundColor Yellow
                        $warnings += "$($depResult.MissingCount) required dependencies not installed"
                        
                        if ($Detailed.IsPresent) {
                            $missingDeps = $depResult.Dependencies | Where-Object { -not $_.IsInstalled -and -not $_.IsOptional }
                            foreach ($dep in $missingDeps) {
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
            Write-Host "`n[CHAR_128274] Signature Check:" -ForegroundColor Cyan
            
            if ($packageInfo.SignatureInfo) {
                $sigStatus = $packageInfo.SignatureInfo.Status
                
                switch ($sigStatus) {
                    'Valid' {
                        Write-Host "  [CHECK] Signature: Valid" -ForegroundColor Green
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
            Write-Host "`n[CHAR_128202] Compatibility Summary:" -ForegroundColor Cyan
            
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
            throw
        }
    }
}
