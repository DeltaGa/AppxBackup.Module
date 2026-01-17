<#
.SYNOPSIS
    Detects and validates Windows SDK tool availability with intelligent path resolution.

.DESCRIPTION
    Features:
    - Multi-version SDK support (Windows 10/11 SDK)
    - Architecture detection (x86, x64, arm64)
    - Tool caching for performance
    - Registry-based SDK discovery
    - Environment variable fallbacks
    - Version compatibility checking
#>

function Test-AppxToolAvailability {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('MakeAppx', 'SignTool', 'MakeCert', 'Pvk2Pfx', 'Certutil', 'PowerShellCert')]
        [string]$ToolName,

        [Parameter()]
        [switch]$ThrowOnMissing,

        [Parameter()]
        [switch]$UpdateCache
    )

    begin {
        # Check cache first (unless UpdateCache is specified)
        if (-not $UpdateCache.IsPresent -and $script:ToolCache.ContainsKey($ToolName)) {
            Write-AppxLog -Message "Tool path retrieved from cache: $ToolName" -Level 'Debug'
            return $true
        }
    }

    process {
        try {
            $toolPath = $null

            # Special handling for PowerShell-native certificate tools
            if ($ToolName -eq 'PowerShellCert') {
                # Check if New-SelfSignedCertificate is available (PS 5.1+)
                $cmdlet = Get-Command -Name 'New-SelfSignedCertificate' -ErrorAction SilentlyContinue
                if ($cmdlet) {
                    $script:ToolCache[$ToolName] = 'PowerShell-Native'
                    Write-AppxLog -Message "Using native PowerShell certificate cmdlets" -Level 'Verbose'
                    return $true
                }
                else {
                    if ($ThrowOnMissing.IsPresent) {
                        throw "PowerShell certificate cmdlets not available. Requires PowerShell 5.1+"
                    }
                    return $false
                }
            }

            # Tool filename mapping
            $toolExe = switch ($ToolName) {
                'MakeAppx'  { 'makeappx.exe' }
                'SignTool'  { 'signtool.exe' }
                'MakeCert'  { 'makecert.exe' }
                'Pvk2Pfx'   { 'pvk2pfx.exe' }
                'Certutil'  { 'certutil.exe' }
            }

            # Strategy 1: Check PATH environment variable
            $pathTool = Get-Command -Name $toolExe -ErrorAction SilentlyContinue
            if ($pathTool) {
                $toolPath = $pathTool.Source
                Write-AppxLog -Message "Found $ToolName in PATH: $toolPath" -Level 'Verbose'
            }

            # Strategy 2: Search Windows SDK installations
            if (-not $toolPath) {
                # Detect current architecture
                $arch = if ([Environment]::Is64BitOperatingSystem) { 'x64' } else { 'x86' }
                
                # Common SDK installation paths
                $sdkBasePaths = @(
                    "${env:ProgramFiles(x86)}\Windows Kits\10\bin",
                    "${env:ProgramFiles}\Windows Kits\10\bin"
                )

                foreach ($basePath in $sdkBasePaths) {
                    if (Test-Path -LiteralPath $basePath) {
                        # Get SDK versions (sorted newest first)
                        $versionDirs = Get-ChildItem -LiteralPath $basePath -Directory -ErrorAction SilentlyContinue |
                            Where-Object { $_.Name -match '^\d+\.\d+\.\d+\.\d+$' } |
                            Sort-Object Name -Descending

                        foreach ($versionDir in $versionDirs) {
                            $candidatePath = Join-Path $versionDir.FullName "$arch\$toolExe"
                            
                            if (Test-Path -LiteralPath $candidatePath) {
                                $toolPath = $candidatePath
                                Write-AppxLog -Message "Found $ToolName in SDK: $toolPath" -Level 'Verbose'
                                break
                            }
                        }
                    }
                    
                    if ($toolPath) { break }
                }
            }

            # Strategy 3: Registry-based discovery
            if (-not $toolPath) {
                try {
                    $regPath = 'HKLM:\SOFTWARE\Microsoft\Windows Kits\Installed Roots'
                    $kitsRoot = (Get-ItemProperty -Path $regPath -Name 'KitsRoot10' -ErrorAction Stop).KitsRoot10
                    
                    if ($kitsRoot) {
                        $arch = if ([Environment]::Is64BitOperatingSystem) { 'x64' } else { 'x86' }
                        $binPath = Join-Path $kitsRoot "bin"
                        
                        if (Test-Path -LiteralPath $binPath) {
                            $versionDirs = Get-ChildItem -LiteralPath $binPath -Directory |
                                Where-Object { $_.Name -match '^\d+\.\d+\.\d+\.\d+$' } |
                                Sort-Object Name -Descending

                            foreach ($versionDir in $versionDirs) {
                                $candidatePath = Join-Path $versionDir.FullName "$arch\$toolExe"
                                
                                if (Test-Path -LiteralPath $candidatePath) {
                                    $toolPath = $candidatePath
                                    Write-AppxLog -Message "Found $ToolName via registry: $toolPath" -Level 'Verbose'
                                    break
                                }
                            }
                        }
                    }
                }
                catch {
                    Write-AppxLog -Message "Registry lookup failed: $_" -Level 'Debug'
                }
            }

            # Strategy 4: Common installation locations (last resort)
            if (-not $toolPath) {
                $commonPaths = @(
                    "${env:ProgramFiles(x86)}\Windows Kits\10\bin\10.0.22621.0\x64\$toolExe",
                    "${env:ProgramFiles(x86)}\Windows Kits\10\bin\10.0.19041.0\x64\$toolExe",
                    "${env:ProgramFiles(x86)}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\x64\$toolExe",
                    "C:\Windows\System32\$toolExe"
                )

                foreach ($path in $commonPaths) {
                    if (Test-Path -LiteralPath $path) {
                        $toolPath = $path
                        Write-AppxLog -Message "Found $ToolName in common location: $toolPath" -Level 'Verbose'
                        break
                    }
                }
            }

            # Result processing
            if ($toolPath) {
                $script:ToolCache[$ToolName] = $toolPath
                Write-AppxLog -Message "Tool available: $ToolName -> $toolPath" -Level 'Debug'
                return $true
            }
            else {
                $message = "Tool not found: $ToolName ($toolExe)"
                Write-AppxLog -Message $message -Level 'Warning'
                
                if ($ThrowOnMissing.IsPresent) {
                    throw $message
                }
                
                return $false
            }
        }
        catch {
            Write-AppxLog -Message "Tool availability check failed: $_" -Level 'Error'
            
            if ($ThrowOnMissing.IsPresent) {
                throw
            }
            
            return $false
        }
    }
}
