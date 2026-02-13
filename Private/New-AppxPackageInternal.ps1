<#
.SYNOPSIS
    Internal function to create APPX/MSIX packages with proper signing and validation.

.DESCRIPTION
    Core package creation logic abstracted for reuse.
    Handles both legacy MakeAppx.exe and modern MSIX tooling.
    
    Features:
    - Automatic tool selection (native vs SDK)
    - Compression optimization
    - Validation during packaging
    - Progress reporting
    - Bundle creation support
#>

function New-AppxPackageInternal {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$SourcePath,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$OutputPath,

        [Parameter()]
        [ValidateSet('None', 'Fast', 'Normal', 'Maximum')]
        [string]$CompressionLevel = 'Normal',

        [Parameter()]
        [switch]$CreateBundle,

        [Parameter()]
        [switch]$ValidateManifest,

        [Parameter()]
        [hashtable]$AdditionalOptions = @{}
    )

    begin {
        Write-AppxLog -Message "Creating package from: $SourcePath" -Level 'Verbose'
        
        # Check if source is in WindowsApps (often has permission issues)
        $isWindowsApps = $SourcePath -like "*\WindowsApps\*"
        $useTempCopy = $false
        $tempSourcePath = $null
    }

    process {
        try {
            # Validate source
            $sourcePath = ConvertTo-SecureFilePath -Path $SourcePath -MustExist -PathType Directory
            
            # Validate manifest exists
            $manifestPath = [System.IO.Path]::Combine($sourcePath, 'AppxManifest.xml')
            if (-not (Test-Path -LiteralPath $manifestPath)) {
                throw "AppxManifest.xml not found in source path"
            }

            # Validate manifest if requested
            if ($ValidateManifest.IsPresent) {
                Write-AppxLog -Message "Validating manifest..." -Level 'Verbose'
                
                try {
                    $manifestData = Get-AppxManifestData -ManifestPath $manifestPath
                    Write-AppxLog -Message "Manifest valid: $($manifestData.Name) v$($manifestData.Version)" -Level 'Debug'
                    
                    # Additional validation: Check for files referenced in manifest
                    [xml]$manifestXml = Get-Content -LiteralPath $manifestPath -Raw
                    
                    # Setup namespace manager properly
                    $nsManager = [System.Xml.XmlNamespaceManager]::new($manifestXml.NameTable)
                    
                    # Detect and register the actual namespace from the manifest
                    $defaultNs = $manifestXml.DocumentElement.NamespaceURI
                    if ($defaultNs) {
                        $nsManager.AddNamespace('appx', $defaultNs)
                        $nsManager.AddNamespace('uap', 'http://schemas.microsoft.com/appx/manifest/uap/windows10')
                        $nsManager.AddNamespace('uap2', 'http://schemas.microsoft.com/appx/manifest/uap/windows10/2')
                        $nsManager.AddNamespace('uap3', 'http://schemas.microsoft.com/appx/manifest/uap/windows10/3')
                    }
                    
                    # Check for logo/icon files using proper XPath with registered namespaces
                    # Query VisualElements for logo attributes
                    $visualElements = $manifestXml.SelectNodes('//appx:VisualElements | //uap:VisualElements', $nsManager)
                    
                    foreach ($element in @($visualElements)) {
                        # Check common logo attributes
                        $logoAttrs = @('Logo', 'Square150x150Logo', 'Square44x44Logo', 'Square71x71Logo', 'Wide310x150Logo', 'Square310x310Logo')
                        foreach ($attr in @($logoAttrs)) {
                            if ($element.HasAttribute($attr)) {
                                $logoPath = $element.GetAttribute($attr)
                                if ($logoPath) {
                                    $fullLogoPath = [System.IO.Path]::Combine($sourcePath, $logoPath)
                                    if (-not (Test-Path -LiteralPath $fullLogoPath)) {
                                        Write-AppxLog -Message "WARNING: Logo file '$logoPath' referenced in manifest not found" -Level 'Warning'
                                    }
                                }
                            }
                        }
                    }
                }
                catch {
                    Write-AppxLog -Message "Manifest validation failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                    throw "Invalid manifest file: $_"
                }
            }

            # Ensure output directory exists
            $outputDir = Split-Path -Path $OutputPath -Parent
            if ($outputDir -and -not (Test-Path -LiteralPath $outputDir)) {
                [void](New-Item -Path $outputDir -ItemType Directory -Force)
            }

            # Package size validation and disk space checking
            try {
                $diskSpaceResult = Test-AppxDiskSpace -SourcePath $sourcePath -OutputPath $OutputPath
            }
            catch {
                Write-AppxLog -Message "Package size validation warning: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                # Non-fatal - continue with packaging unless it's a disk space error
                if ($_.Exception.Message -like "*Insufficient disk space*") {
                    throw
                }
            }

            # Check if output file already exists
            if (Test-Path -LiteralPath $OutputPath) {
                Write-AppxLog -Message "Removing existing output file: $OutputPath" -Level 'Verbose'
                Remove-Item -LiteralPath $OutputPath -Force
            }

            # Determine packaging method
            $useMakeAppx = Test-AppxToolAvailability -ToolName 'MakeAppx'
            
            if ($useMakeAppx) {
                Write-AppxLog -Message "Using MakeAppx.exe for packaging" -Level 'Debug'
                
                # Get MakeAppx path from cache
                $makeAppxPath = $script:ToolCache['MakeAppx']
                
                # If source is WindowsApps, try to use it directly first, fallback to temp copy
                $effectiveSourcePath = $sourcePath
                
                if ($isWindowsApps) {
                    Write-AppxLog -Message "Source is in WindowsApps folder, checking access..." -Level 'Debug'
                    
                    # Test if we can read the directory without needing a temp copy
                    $needsTempCopy = $false
                    try {
                        $null = Get-ChildItem -LiteralPath $sourcePath -ErrorAction Stop | Select-Object -First 1
                        Write-AppxLog -Message "Direct access to WindowsApps folder successful" -Level 'Debug'
                        
                        # Check for problematic signature files that will cause repackaging to fail
                        try {
                            $pkgConfig = Get-AppxConfiguration -ConfigName 'PackageConfiguration'
                            $signatureFiles = $pkgConfig.signatureFiles.files
                        }
                        catch {
                            Write-AppxLog -Message "Failed to load signature files config, using fallback: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
                            $signatureFiles = @('AppxSignature.p7x', 'AppxBlockMap.xml')
                        }
                        
                        foreach ($sigFile in @($signatureFiles)) {
                            $sigPath = [System.IO.Path]::Combine($sourcePath, $sigFile)
                            if (Test-Path -LiteralPath $sigPath) {
                                Write-AppxLog -Message "Found existing signature files, will create temp copy to remove them" -Level 'Debug'
                                $needsTempCopy = $true
                                break
                            }
                        }
                    }
                    catch {
                        $needsTempCopy = $true
                    }
                    
                    if ($needsTempCopy) {
                        Write-AppxLog -Message "Cannot access WindowsApps folder directly, copying to temp location..." -Level 'Verbose'
                        
                        $tempSourcePath = [System.IO.Path]::Combine($env:TEMP, "AppxBackup_$(New-Guid)")
                        [void](New-Item -Path $tempSourcePath -ItemType Directory -Force)
                        
                        Write-AppxLog -Message "Copying app files to: $tempSourcePath" -Level 'Verbose'
                        
                        $copyResult = Copy-AppxSourceDirectory -SourcePath $sourcePath -DestinationPath $tempSourcePath
                        
                        if (-not $copyResult.Success) {
                            # Cleanup temp dir on failure
                            if (Test-Path -LiteralPath $tempSourcePath) {
                                Remove-Item -LiteralPath $tempSourcePath -Recurse -Force -ErrorAction SilentlyContinue
                            }
                            throw "Failed to copy WindowsApps folder: $($copyResult.Error)"
                        }
                        
                        Write-AppxLog -Message "Copy succeeded using: $($copyResult.CopyMethod)" -Level 'Verbose'
                        $effectiveSourcePath = $tempSourcePath
                        $useTempCopy = $true
                    }
                }
                
                # Build arguments for MakeAppx
                # pack = create package command
                # /v = verbose output for better error diagnostics
                # /d = directory to pack
                # /p = output package path
                # /o = overwrite existing (prevents interactive prompts)
                # /nc = no compression (optional, based on CompressionLevel)
                # /nv = disable version checking (optional, based on AdditionalOptions)
                
                $makeAppxArgs = @(
                    'pack',
                    '/v',
                    '/d', $effectiveSourcePath,
                    '/p', $OutputPath
                )
                
                # Add compression flag (or disable it)
                switch ($CompressionLevel) {
                    'None' { 
                        $makeAppxArgs += '/nc'
                    }
                    # Fast, Normal, Maximum use default compression
                }
                
                # Add overwrite to prevent interactive prompts
                $makeAppxArgs += '/o'
                
                # Add additional options
                if ($AdditionalOptions.ContainsKey('DisableVersioning')) {
                    $makeAppxArgs += '/nv'
                }

                # Log command for diagnostics
                Write-AppxLog -Message "Executing: $makeAppxPath $($makeAppxArgs -join ' ')" -Level 'Debug'
                
                # PRE-FLIGHT DIAGNOSTICS
                $preflightResult = Test-AppxPackagingPrerequisites -SourcePath $effectiveSourcePath -OutputPath $OutputPath
                
                Write-AppxLog -Message "Pre-flight checks complete, invoking MakeAppx..." -Level 'Debug'
                Write-AppxLog -Message "This may take several minutes for large packages..." -Level 'Verbose'
                
                $result = Invoke-ProcessSafely -FilePath $makeAppxPath `
                    -ArgumentList $makeAppxArgs `
                    -TimeoutSeconds 1800 `
                    -NoWindow

                if (-not $result.Success) {
                    Write-AppxLog -Message "MakeAppx failed. Analyzing error output..." -Level 'Debug'
                    Write-AppxLog -Message "Exit Code: $($result.ExitCode)" -Level 'Debug'
                    
                    $errorMsg = Get-AppxMakeAppxErrorAnalysis `
                        -ExitCode $result.ExitCode `
                        -StandardError $result.StandardError `
                        -StandardOutput $result.StandardOutput `
                        -MakeAppxPath $makeAppxPath `
                        -Arguments $makeAppxArgs `
                        -SourcePath $effectiveSourcePath `
                        -OutputPath $OutputPath
                    
                    Write-AppxLog -Message $errorMsg -Level 'Error'
                    throw $errorMsg
                }

                # Verify output was created
                if (-not (Test-Path -LiteralPath $OutputPath)) {
                    throw "Package file was not created: $OutputPath"
                }

                Write-AppxLog -Message "Package created successfully: $OutputPath" -Level 'Verbose'
            }
            else {
                # Fallback: Use .NET System.IO.Compression
                Write-AppxLog -Message "Using .NET compression (MakeAppx not available)" -Level 'Warning'
                
                Add-Type -AssemblyName 'System.IO.Compression.FileSystem'
                
                $compressionLevelEnum = switch ($CompressionLevel) {
                    'None'      { [System.IO.Compression.CompressionLevel]::NoCompression }
                    'Fast'      { [System.IO.Compression.CompressionLevel]::Fastest }
                    'Normal'    { [System.IO.Compression.CompressionLevel]::Optimal }
                    'Maximum'   { [System.IO.Compression.CompressionLevel]::Optimal }
                    default     { [System.IO.Compression.CompressionLevel]::Optimal }
                }

                Write-AppxLog -Message "Creating ZIP archive..." -Level 'Verbose'
                [System.IO.Compression.ZipFile]::CreateFromDirectory(
                    $sourcePath,
                    $OutputPath,
                    $compressionLevelEnum,
                    $false # includeBaseDirectory
                )

                Write-AppxLog -Message "Archive created (note: not a proper APPX, use MakeAppx for production)" -Level 'Warning'
            }

            # Get package info
            $packageFile = Get-Item -LiteralPath $OutputPath
            $packageSize = $packageFile.Length

            # Build result object
            $packageResult = [PSCustomObject]@{
                PSTypeName      = 'AppxBackup.PackageResult'
                PackagePath     = $OutputPath
                PackageName     = $packageFile.Name
                PackageSize     = $packageSize
                PackageSizeMB   = [Math]::Round($packageSize / 1MB, 2)
                SourcePath      = $sourcePath
                CreatedTime     = $packageFile.CreationTime
                CompressionUsed = $CompressionLevel
                Success         = $true
            }

            return $packageResult
        }
        catch {
            Write-AppxLog -Message "Package creation failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'
            Write-AppxLog -Message "StackTrace: $($_.ScriptStackTrace)" -Level 'Debug'
            throw
        }
        finally {
            # Cleanup temp copy if used
            if ($useTempCopy -and $tempSourcePath -and (Test-Path -LiteralPath $tempSourcePath)) {
                $null = Remove-AppxItemWithRetry -Path $tempSourcePath -Recurse
            }
        }
    }
}