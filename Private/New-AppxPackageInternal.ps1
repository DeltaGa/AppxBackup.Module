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
            Write-AppxLog -Message "Validating package size and disk space..." -Level 'Verbose'
            
            try {
                # Calculate source directory size
                $sourceFiles = Get-ChildItem -LiteralPath $sourcePath -Recurse -File -ErrorAction Stop
                $sourceSize = ($sourceFiles | Measure-Object -Property Length -Sum).Sum
                $sourceSizeMB = [Math]::Round($sourceSize / 1MB, 2)
                $fileCount = $sourceFiles.Count
                
                Write-AppxLog -Message "Source directory: $sourceSizeMB MB ($fileCount files)" -Level 'Verbose'
                
                # Load thresholds from configuration
                $largeSizeThreshold = Get-AppxDefault 'packageSizeThresholds.largeSizeMB' -Fallback 1024
                $largeFileCountThreshold = Get-AppxDefault 'packageSizeThresholds.largeFileCount' -Fallback 10000
                
                # Warn if package is large (>configured threshold)
                if ($sourceSizeMB -gt $largeSizeThreshold) {
                    Write-Host "[WARNING] Large package detected ($sourceSizeMB MB, $fileCount files)" -ForegroundColor Yellow
                    Write-Host "  Processing may take several minutes. Please be patient..." -ForegroundColor Yellow
                    Write-AppxLog -Message "Large package warning displayed to user" -Level 'Info'
                }
                
                # Warn if too many files (>configured threshold)
                if ($fileCount -gt $largeFileCountThreshold) {
                    Write-Host "[WARNING] Package contains many files ($fileCount)" -ForegroundColor Yellow
                    Write-Host "  Packaging performance may be slower than usual..." -ForegroundColor Yellow
                    Write-AppxLog -Message "High file count warning displayed to user" -Level 'Info'
                }
                
                # Check available disk space (require configured multiplier * source size + safety margin)
                $outputDrive = [System.IO.Path]::GetPathRoot($OutputPath)
                if ($outputDrive) {
                    $driveInfo = [System.IO.DriveInfo]::new($outputDrive)
                    $availableSpaceMB = [Math]::Round($driveInfo.AvailableFreeSpace / 1MB, 2)
                    
                    # Load disk space calculation parameters from configuration
                    $multiplier = Get-AppxDefault 'diskSpaceRequirements.sourceMultiplier' -Fallback 2
                    $safetyMargin = Get-AppxDefault 'diskSpaceRequirements.safetyMarginMB' -Fallback 100
                    $minimalRemaining = Get-AppxDefault 'diskSpaceRequirements.minimalAvailableSpaceMB' -Fallback 500
                    
                    $requiredSpaceMB = ($sourceSizeMB * $multiplier) + $safetyMargin
                    
                    Write-AppxLog -Message "Disk space - Available: $availableSpaceMB MB, Required: $requiredSpaceMB MB" -Level 'Debug'
                    
                    if ($availableSpaceMB -lt $requiredSpaceMB) {
                        $shortfall = [Math]::Round($requiredSpaceMB - $availableSpaceMB, 2)
                        throw "Insufficient disk space on $outputDrive. Need $shortfall MB more space. Available: $availableSpaceMB MB, Required: $requiredSpaceMB MB"
                    }
                    
                    # Warn if available space is marginal (<configured minimum after operation)
                    if (($availableSpaceMB - $requiredSpaceMB) -lt $minimalRemaining) {
                        Write-Warning "Disk space will be low after packaging (less than $minimalRemaining MB remaining)"
                        Write-AppxLog -Message "Low disk space warning displayed" -Level 'Warning'
                    }
                }
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
                    
                    # Test if we can read the directory
                    try {
                        $null = Get-ChildItem -LiteralPath $sourcePath -ErrorAction Stop | Select-Object -First 1
                        Write-AppxLog -Message "Direct access to WindowsApps folder successful" -Level 'Debug'
                        
                        # Check for problematic signature files that will cause repackaging to fail
                        # Load from configuration
                        try {
                            $pkgConfig = Get-AppxConfiguration -ConfigName 'PackageConfiguration'
                            $signatureFiles = $pkgConfig.signatureFiles.files
                        }
                        catch {
                            Write-AppxLog -Message "Failed to load signature files config, using fallback: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
                            $signatureFiles = @('AppxSignature.p7x', 'AppxBlockMap.xml')
                        }
                        
                        $hasSignatureFiles = $false
                        foreach ($sigFile in @($signatureFiles)) {
                            $sigPath = [System.IO.Path]::Combine($sourcePath, $sigFile)
                            if (Test-Path -LiteralPath $sigPath) {
                                $hasSignatureFiles = $true
                                break
                            }
                        }
                        
                        # If signature files exist, we need to use temp copy to remove them
                        if ($hasSignatureFiles) {
                            Write-AppxLog -Message "Found existing signature files, will create temp copy to remove them" -Level 'Debug'
                            throw "Need temp copy to remove signature files"
                        }
                    }
                    catch {
                        Write-AppxLog -Message "Cannot access WindowsApps folder directly, copying to temp location..." -Level 'Verbose'
                        
                        # Create temp directory
                        $tempSourcePath = [System.IO.Path]::Combine($env:TEMP, "AppxBackup_$(New-Guid)")
                        [void](New-Item -Path $tempSourcePath -ItemType Directory -Force)
                        
                        # Copy with permissions override (using robocopy for better handling)
                        Write-AppxLog -Message "Copying app files to: $tempSourcePath" -Level 'Verbose'
                        
                        try {
                            $copySucceeded = $false
                            $copyMethod = 'Unknown'
                            
                            # STRATEGY 1: Robocopy with simpler flags (fastest, most reliable for WindowsApps)
                            try {
                                Write-AppxLog -Message "Attempting copy with Robocopy..." -Level 'Debug'
                                
                                # Simplified Robocopy - NO QUOTES in array, PowerShell handles this
                                # Using direct paths without embedded quotes
                                $robocopyArgs = @(
                                    $sourcePath,          # No quotes - PowerShell adds them
                                    $tempSourcePath,      # No quotes - PowerShell adds them
                                    '*.*',                # All files
                                    '/E',                 # Copy subdirectories including empty
                                    '/COPY:DAT',          # Copy Data, Attributes, Timestamps
                                    '/DCOPY:DA',          # Copy directory Attributes
                                    '/R:1',               # Retry once
                                    '/W:1',               # Wait 1 second
                                    '/NP',                # No progress
                                    '/NJH',               # No job header
                                    '/NJS'                # No job summary
                                )
                                
                                Write-AppxLog -Message "Robocopy command: robocopy.exe $($robocopyArgs -join ' ')" -Level 'Debug'
                                
                                # Capture output for diagnostics
                                $robocopyOutput = & robocopy.exe @robocopyArgs 2>&1
                                $robocopyExitCode = $LASTEXITCODE
                                
                                Write-AppxLog -Message "Robocopy exit code: $robocopyExitCode" -Level 'Debug'
                                
                                # Robocopy exit codes: 0-7 are success, 8+ are errors
                                if ($robocopyExitCode -lt 8) {
                                    $copySucceeded = $true
                                    $copyMethod = 'Robocopy'
                                    Write-AppxLog -Message "Robocopy succeeded" -Level 'Debug'
                                }
                                else {
                                    $robocopyOutputStr = $robocopyOutput -join "`n"
                                    Write-AppxLog -Message "Robocopy failed (exit $robocopyExitCode): $robocopyOutputStr" -Level 'Warning'
                                    throw "Robocopy exit code $robocopyExitCode"
                                }
                            }
                            catch {
                                Write-AppxLog -Message "Robocopy failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                                Write-AppxLog -Message "Falling back to PowerShell native copy..." -Level 'Verbose'
                                
                                # STRATEGY 2: PowerShell Copy-Item (handles most cases)
                                try {
                                    Write-AppxLog -Message "Attempting copy with Copy-Item..." -Level 'Debug'
                                    
                                    # Get all items first, then copy (works around permission issues)
                                    $itemsToCopy = Get-ChildItem -Path $sourcePath -Recurse -Force -ErrorAction Stop
                                    
                                    foreach ($item in @($itemsToCopy)) {
                                        $relativePath = $item.FullName.Substring($sourcePath.Length).TrimStart('\')
                                        $destPath = [System.IO.Path]::Combine($tempSourcePath, $relativePath)
                                        
                                        if ($item.PSIsContainer) {
                                            if (-not (Test-Path -LiteralPath $destPath)) {
                                                [void](New-Item -Path $destPath -ItemType Directory -Force -ErrorAction SilentlyContinue)
                                            }
                                        }
                                        else {
                                            $destDir = Split-Path -Path $destPath -Parent
                                            if ($destDir -and -not (Test-Path -LiteralPath $destDir)) {
                                                [void](New-Item -Path $destDir -ItemType Directory -Force -ErrorAction SilentlyContinue)
                                            }
                                            Copy-Item -LiteralPath $item.FullName -Destination $destPath -Force -ErrorAction Stop
                                        }
                                    }
                                    
                                    $copySucceeded = $true
                                    $copyMethod = 'Copy-Item'
                                    Write-AppxLog -Message "Copy-Item succeeded" -Level 'Debug'
                                }
                                catch {
                                    Write-AppxLog -Message "Copy-Item failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                                    Write-AppxLog -Message "Falling back to .NET file copy..." -Level 'Verbose'
                                    
                                    # STRATEGY 3: .NET Directory copy (most granular control)
                                    try {
                                        Write-AppxLog -Message "Attempting copy with .NET APIs..." -Level 'Debug'
                                        
                                        # Recursive .NET copy function with error tolerance
                                        function Copy-DirectoryNET {
                                            param($Source, $Destination)
                                            
                                            $sourceDir = [System.IO.DirectoryInfo]::new($Source)
                                            
                                            # Create destination directory
                                            if (-not [System.IO.Directory]::Exists($Destination)) {
                                                [void][System.IO.Directory]::CreateDirectory($Destination)
                                            }
                                            
                                            # Copy files with error tolerance
                                            foreach ($file in $sourceDir.GetFiles()) {
                                                try {
                                                    $destFile = [System.IO.Path]::Combine($Destination, $file.Name)
                                                    [System.IO.File]::Copy($file.FullName, $destFile, $true)
                                                }
                                                catch {
                                                    Write-AppxLog -Message "  Skipping file (access denied): $($file.Name)" -Level 'Debug'
                                                }
                                            }
                                            
                                            # Copy subdirectories recursively
                                            foreach ($dir in $sourceDir.GetDirectories()) {
                                                try {
                                                    $destDir = [System.IO.Path]::Combine($Destination, $dir.Name)
                                                    Copy-DirectoryNET -Source $dir.FullName -Destination $destDir
                                                }
                                                catch {
                                                    Write-AppxLog -Message "  Skipping directory (access denied): $($dir.Name)" -Level 'Debug'
                                                }
                                            }
                                        }
                                        
                                        Copy-DirectoryNET -Source $sourcePath -Destination $tempSourcePath
                                        
                                        $copySucceeded = $true
                                        $copyMethod = '.NET APIs'
                                        Write-AppxLog -Message ".NET copy succeeded" -Level 'Debug'
                                    }
                                    catch {
                                        Write-AppxLog -Message ".NET copy failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'
                                        Write-AppxLog -Message "StackTrace: $($_.ScriptStackTrace)" -Level 'Debug'
                                        throw "All copy strategies failed. Last error: $_"
                                    }
                                }
                            }
                            
                            if ($null -eq $copySucceeded) {
                                throw "Failed to copy source files"
                            }
                            
                            Write-AppxLog -Message "Copy succeeded using: $copyMethod" -Level 'Verbose'
                            
                            $effectiveSourcePath = $tempSourcePath
                            $useTempCopy = $true
                            Write-AppxLog -Message "Copy complete, using temp location" -Level 'Debug'
                            
                            # Remove existing signature files as they're invalid for repackaging
                            # Load from configuration (should already be loaded, but fallback included)
                            try {
                                if ($null -eq $signatureFiles) {
                                    $pkgConfig = Get-AppxConfiguration -ConfigName 'PackageConfiguration'
                                    $signatureFiles = $pkgConfig.signatureFiles.files
                                }
                            }
                            catch {
                                $signatureFiles = @('AppxSignature.p7x', 'AppxBlockMap.xml')
                            }
                            foreach ($sigFile in @($signatureFiles)) {
                                $sigPath = [System.IO.Path]::Combine($tempSourcePath, $sigFile)
                                if (Test-Path -LiteralPath $sigPath) {
                                    Remove-Item -LiteralPath $sigPath -Force -ErrorAction SilentlyContinue
                                    Write-AppxLog -Message "Removed existing $sigFile for repackaging" -Level 'Debug'
                                }
                            }
                            
                            # Verify critical files exist after copy
                            $criticalFiles = @('AppxManifest.xml', '[Content_Types].xml')
                            foreach ($critical in @($criticalFiles)) {
                                $criticalPath = [System.IO.Path]::Combine($tempSourcePath, $critical)
                                if (-not (Test-Path -LiteralPath $criticalPath)) {
                                    
                                    # Special handling for [Content_Types].xml - generate if missing
                                    if ($critical -eq '[Content_Types].xml') {
                                        Write-AppxLog -Message "WARNING: [Content_Types].xml missing from source, generating comprehensive version" -Level 'Warning'
                                        
                                        # Scan directory for all file extensions to include
                                        $extensions = @{}
                                        try {
                                            Get-ChildItem -LiteralPath $tempSourcePath -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object {
                                                $ext = $_.Extension.TrimStart('.')
                                                if ($ext -and -not $extensions.ContainsKey($ext)) {
                                                    $extensions[$ext] = $true
                                                }
                                            }
                                            Write-AppxLog -Message "Found $($extensions.Count) unique file extensions in package" -Level 'Debug'
                                        }
                                        catch {
                                            Write-AppxLog -Message "Could not scan extensions: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
                                        }
                                        
                                        # Build comprehensive Content_Types.xml with MIME types from external database
                                        $sb = [System.Text.StringBuilder]::new()
                                        [void]$sb.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
                                        [void]$sb.AppendLine('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">')
                                        
                                        # Load MIME type database from external configuration
                                        try {
                                            $mimeConfig = Get-AppxConfiguration -ConfigName 'MimeTypes'
                                            $mimeTypes = @{}
                                            
                                            # Convert PSCustomObject properties to hashtable for easier lookup
                                            foreach ($property in @($mimeConfig.mimeTypes.PSObject.Properties)) {
                                                $mimeTypes[$property.Name] = $property.Value
                                            }
                                            
                                            $defaultContentType = $mimeConfig.defaultContentType
                                            
                                            Write-AppxLog -Message "Loaded $($mimeTypes.Count) MIME types from configuration" -Level 'Debug'
                                        }
                                        catch {
                                            # Fallback to minimal MIME types if configuration fails
                                            Write-AppxLog -Message "Failed to load MIME configuration, using fallback: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                                            
                                            $mimeTypes = @{
                                                'xml'   = 'application/xml'
                                                'dll'   = 'application/x-msdownload'
                                                'exe'   = 'application/x-msdownload'
                                                'png'   = 'image/png'
                                                'jpg'   = 'image/jpeg'
                                                'txt'   = 'text/plain'
                                            }
                                            $defaultContentType = 'application/octet-stream'
                                        }
                                        
                                        # Add default entries for all found extensions
                                        foreach ($ext in @($extensions.Keys)) {
                                            $contentType = if ($mimeTypes.ContainsKey($ext)) {
                                                $mimeTypes[$ext]
                                            } else {
                                                $defaultContentType
                                            }
                                            [void]$sb.AppendLine("  <Default Extension=`"$ext`" ContentType=`"$contentType`"/>")
                                        }
                                        
                                        # Add standard entries that might not exist yet but are needed
                                        foreach ($kvp in $mimeTypes.GetEnumerator()) {
                                            if (-not $extensions.ContainsKey($kvp.Key)) {
                                                [void]$sb.AppendLine("  <Default Extension=`"$($kvp.Key)`" ContentType=`"$($kvp.Value)`"/>")
                                            }
                                        }
                                        
                                        # Add APPX-specific overrides
                                        [void]$sb.AppendLine('  <Override PartName="/AppxManifest.xml" ContentType="application/vnd.ms-appx.manifest+xml"/>')
                                        [void]$sb.AppendLine('  <Override PartName="/AppxBlockMap.xml" ContentType="application/vnd.ms-appx.blockmap+xml"/>')
                                        [void]$sb.AppendLine('  <Override PartName="/AppxSignature.p7x" ContentType="application/vnd.ms-appx.signature"/>')
                                        [void]$sb.AppendLine('</Types>')
                                        
                                        $contentTypesXml = $sb.ToString()
                                        
                                        try {
                                            [System.IO.File]::WriteAllText($criticalPath, $contentTypesXml, [System.Text.UTF8Encoding]::new($false))
                                            Write-AppxLog -Message "Generated [Content_Types].xml with $($extensions.Count + $mimeTypes.Count) entries" -Level 'Debug'
                                        }
                                        catch {
                                            Write-AppxLog -Message "Failed to generate [Content_Types].xml: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'
                throw "Failed to generate [Content_Types].xml: $_"
                                        }
                                    }
                                    else {
                                        throw "Critical file missing after copy: $critical (MakeAppx requires this file)"
                                    }
                                }
                            }
                        }
                        catch {
                            # Cleanup temp dir on failure
                            if (Test-Path -LiteralPath $tempSourcePath) {
                                Remove-Item -LiteralPath $tempSourcePath -Recurse -Force -ErrorAction SilentlyContinue
                            }
                            throw "Failed to copy WindowsApps folder: $_"
                        }
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
                
                # PRE-FLIGHT DIAGNOSTICS: Check for common MakeAppx failure conditions
                Write-AppxLog -Message "Running pre-flight diagnostics..." -Level 'Debug'
                
                # Check 1: Verify source path doesn't contain problematic characters
                if ($effectiveSourcePath -match '[<>"|?*]') {
                    Write-AppxLog -Message "WARNING: Source path contains characters that may cause issues: $effectiveSourcePath" -Level 'Warning'
                }
                
                # Check 2: Verify we can actually read critical files
                $criticalFiles = @('AppxManifest.xml', 'AppxBlockMap.xml', 'AppxSignature.p7x')
                foreach ($file in @($criticalFiles)) {
                    $testPath = [System.IO.Path]::Combine($effectiveSourcePath, $file)
                    if (Test-Path -LiteralPath $testPath) {
                        try {
                            $null = Get-Content -LiteralPath $testPath -TotalCount 1 -ErrorAction Stop
                            Write-AppxLog -Message "  [OK] Can read: $file" -Level 'Debug'
                        }
                        catch {
                            Write-AppxLog -Message "  [FAIL] Cannot read file: $file - $_" -Level 'Warning'
                            Write-AppxLog -Message "This may indicate permission or locking issues" -Level 'Warning'
                        }
                    }
                }
                
                # Check 3: Verify output directory is writable
                $outputDir = Split-Path -Path $OutputPath -Parent
                $testFile = [System.IO.Path]::Combine($outputDir, "AppxBackup_WriteTest_$(New-Guid).tmp")
                try {
                    [void][System.IO.File]::WriteAllText($testFile, "test")
                    Remove-Item -LiteralPath $testFile -Force -ErrorAction SilentlyContinue
                    Write-AppxLog -Message "  [OK] Output directory is writable" -Level 'Debug'
                }
                catch {
                    Write-AppxLog -Message "  [FAIL] Cannot write to output directory: $outputDir" -Level 'Warning'
                    throw "Output directory is not writable: $_"
                }
                
                # Check 4: Warn if source has too many files (can cause timeout)
                try {
                    $fileCount = (Get-ChildItem -LiteralPath $effectiveSourcePath -Recurse -File -ErrorAction SilentlyContinue | Measure-Object).Count
                    Write-AppxLog -Message "  [INFO] Package contains $fileCount files" -Level 'Debug'
                    if ($fileCount -gt 10000) {
                        Write-AppxLog -Message "  [WARN] Large file count may slow packaging" -Level 'Warning'
                    }
                }
                catch {
                    Write-AppxLog -Message "  [WARN] Could not count files: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
                }
                
                Write-AppxLog -Message "Pre-flight checks complete, invoking MakeAppx..." -Level 'Debug'
                Write-AppxLog -Message "This may take several minutes for large packages..." -Level 'Verbose'
                
                $result = Invoke-ProcessSafely -FilePath $makeAppxPath `
                    -ArgumentList $makeAppxArgs `
                    -TimeoutSeconds 1800 `
                    -NoWindow

                if (-not $result.Success) {
                    # MakeAppx writes errors to STDOUT and STDERR (no log file - /l flag doesn't exist)
                    # Verbose mode (/v) should give us detailed error information
                    
                    Write-AppxLog -Message "MakeAppx failed. Analyzing error output..." -Level 'Debug'
                    Write-AppxLog -Message "Exit Code: $($result.ExitCode)" -Level 'Debug'
                    
                    # Collect all available output
                    $stderrContent = if ($result.StandardError) { $result.StandardError.Trim() } else { "" }
                    $stdoutContent = if ($result.StandardOutput) { $result.StandardOutput.Trim() } else { "" }
                    
                    Write-AppxLog -Message "STDERR Length: $($stderrContent.Length) bytes" -Level 'Debug'
                    Write-AppxLog -Message "STDOUT Length: $($stdoutContent.Length) bytes" -Level 'Debug'
                    
                    if ($stderrContent) {
                        Write-AppxLog -Message "STDERR Output:`n$stderrContent" -Level 'Error'
                    }
                    
                    if ($stdoutContent) {
                        Write-AppxLog -Message "STDOUT Output:`n$stdoutContent" -Level 'Error'
                    }
                    
                    # Analyze the error to provide helpful diagnostics
                    $errorAnalysis = @()
                    $fullError = "$stderrContent`n$stdoutContent"
                    
                    # Common MakeAppx error patterns with actionable guidance
                    if ($fullError -match 'Access is denied|ERROR_ACCESS_DENIED|0x80070005') {
                        $errorAnalysis += "PERMISSION ISSUE: MakeAppx cannot access one or more files"
                        $errorAnalysis += "  - Try running PowerShell as Administrator"
                        $errorAnalysis += "  - Check if files in '$effectiveSourcePath' are locked by another process"
                        $errorAnalysis += "  - Verify NTFS permissions on source and output directories"
                    }
                    
                    if ($fullError -match 'The system cannot find the file specified|ERROR_FILE_NOT_FOUND|0x80070002') {
                        $errorAnalysis += "FILE NOT FOUND: MakeAppx cannot locate required files"
                        $errorAnalysis += "  - Verify AppxManifest.xml exists in source directory"
                        $errorAnalysis += "  - Check if files referenced in manifest are present"
                        $errorAnalysis += "  - Ensure logo/icon files specified in manifest exist"
                    }
                    
                    if ($fullError -match 'The parameter is incorrect|ERROR_INVALID_PARAMETER|0x80070057') {
                        $errorAnalysis += "INVALID PARAMETER: Command-line arguments may be malformed"
                        $errorAnalysis += "  - Check for special characters in paths"
                        $errorAnalysis += "  - Verify path lengths are under 260 characters"
                    }
                    
                    if ($fullError -match 'invalid|malformed|corrupt|error in manifest|0x8007000d') {
                        $errorAnalysis += "MANIFEST ERROR: The AppxManifest.xml file contains errors"
                        $errorAnalysis += "  - Validate XML syntax in AppxManifest.xml"
                        $errorAnalysis += "  - Check for missing required elements"
                        $errorAnalysis += "  - Verify all file references are correct"
                        $errorAnalysis += "  - Ensure namespace declarations are present"
                    }
                    
                    if ($fullError -match '0x80080204|E_APPX_INVALID_MANIFEST') {
                        $errorAnalysis += "INVALID MANIFEST SCHEMA: The manifest doesn't conform to the APPX schema"
                        $errorAnalysis += "  - Check that all required manifest elements are present"
                        $errorAnalysis += "  - Verify Identity, Properties, and Applications elements exist"
                        $errorAnalysis += "  - Ensure Publisher matches certificate subject name"
                    }
                    
                    if ($fullError -match '0x80080206|E_APPX_INVALID_BLOCKMAP') {
                        $errorAnalysis += "BLOCKMAP ERROR: Issue with package block mapping"
                        $errorAnalysis += "  - This usually means AppxBlockMap.xml is corrupt or inconsistent"
                        $errorAnalysis += "  - Try deleting any existing AppxBlockMap.xml and let MakeAppx regenerate it"
                    }
                    
                    if ($fullError -match 'ERROR_PATH_NOT_FOUND|0x80070003') {
                        $errorAnalysis += "PATH NOT FOUND: One or more directory paths are invalid"
                        $errorAnalysis += "  - Verify source directory exists: $effectiveSourcePath"
                        $errorAnalysis += "  - Check output directory is valid: $OutputPath"
                    }
                    
                    if ($fullError -match 'invalid manifest|manifest.*error') {
                        $errorAnalysis += "MANIFEST ERROR: AppxManifest.xml contains errors"
                        $errorAnalysis += "  - Manifest path: $([System.IO.Path]::Combine($effectiveSourcePath, 'AppxManifest.xml'))"
                        $errorAnalysis += "  - Use Get-AppxManifestData to validate manifest structure"
                    }
                    
                    if ($fullError -match '0x80080204|invalid signature') {
                        $errorAnalysis += "SIGNATURE ERROR: Existing signature is invalid or corrupted"
                        $errorAnalysis += "  - Remove AppxSignature.p7x from source if present"
                        $errorAnalysis += "  - Package will be re-signed after creation"
                    }
                    
                    if ($fullError -match 'already exists|file.*in use') {
                        $errorAnalysis += "OUTPUT FILE CONFLICT: Target file exists or is locked"
                        $errorAnalysis += "  - Output: $OutputPath"
                        $errorAnalysis += "  - Try deleting the file manually or closing programs using it"
                    }
                    
                    # If no specific pattern matched, provide generic help
                    if ($errorAnalysis.Count -eq 0) {
                        $errorAnalysis += "UNKNOWN ERROR: MakeAppx failed for an unrecognized reason"
                        $errorAnalysis += "  - Check the log file for complete output"
                        $errorAnalysis += "  - Try the .NET fallback by temporarily renaming MakeAppx.exe"
                    }
                    
                    # Build comprehensive error message
                    $errorMsg = "MakeAppx.exe failed with exit code $($result.ExitCode)`n`n"
                    $errorMsg += "=== ERROR ANALYSIS ===`n"
                    $errorMsg += ($errorAnalysis -join "`n") + "`n`n"
                    $errorMsg += "=== COMMAND DETAILS ===`n"
                    $errorMsg += "Executable: $makeAppxPath`n"
                    $errorMsg += "Arguments: $($makeAppxArgs -join ' ')`n"
                    $errorMsg += "Source: $effectiveSourcePath`n"
                    $errorMsg += "Output: $OutputPath`n`n"
                    
                    if ($stderrContent) {
                        $errorMsg += "=== STDERR (Error Output) ===`n$stderrContent`n`n"
                    }
                    if ($stdoutContent) {
                        $errorMsg += "=== STDOUT (Standard Output) ===`n$stdoutContent`n`n"
                    }
                    
                    if (-not $stderrContent -and -not $stdoutContent) {
                        $errorMsg += "=== NO ERROR OUTPUT CAPTURED ===`n"
                        $errorMsg += "MakeAppx provided no error details. Possible causes:`n"
                        $errorMsg += "  - Process terminated before writing output`n"
                        $errorMsg += "  - Permissions issue preventing output capture`n"
                        $errorMsg += "  - Try running PowerShell as Administrator`n`n"
                    }
                    
                    $errorMsg += "Full details logged to: $env:TEMP\AppxBackup_$(Get-Date -Format 'yyyyMMdd').log"
                    
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
                Write-AppxLog -Message "Cleaning up temporary copy: $tempSourcePath" -Level 'Debug'
                
                # Load cleanup configuration from config
                $maxCleanupAttempts = Get-AppxDefault 'sleepDelays.maxCleanupAttempts' -Fallback 3
                $cleanupDelays = Get-AppxDefault 'sleepDelays.cleanupRetryDelaysMilliseconds' -Fallback @(300, 2000, 5000)
                $cleanupSuccess = $false
                
                for ($attempt = 0; $attempt -lt $maxCleanupAttempts; $attempt++) {
                    try {
                        # Wait before retry (skip on first attempt)
                        if ($attempt -gt 0) {
                            $delayMs = $cleanupDelays[$attempt - 1]
                            Write-AppxLog -Message "Waiting $($delayMs)ms before cleanup attempt $($attempt + 1)/$maxCleanupAttempts..." -Level 'Debug'
                            Start-Sleep -Milliseconds $delayMs
                        }
                        
                        # Attempt cleanup
                        Remove-Item -LiteralPath $tempSourcePath -Recurse -Force -ErrorAction Stop
                        $cleanupSuccess = $true
                        Write-AppxLog -Message "Cleanup successful on attempt $($attempt + 1)" -Level 'Debug'
                        break
                    }
                    catch {
                        Write-AppxLog -Message "Cleanup attempt $($attempt + 1) failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
                        
                        if ($attempt -eq ($maxCleanupAttempts - 1)) {
                            # Final attempt failed - log warning but don't fail the operation
                            Write-AppxLog -Message "Failed to cleanup temp directory after $maxCleanupAttempts attempts: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                            Write-AppxLog -Message "Temp files remain at: $tempSourcePath" -Level 'Warning'
                            Write-AppxLog -Message "These files can be manually deleted later or will be cleaned up on next reboot" -Level 'Info'
                        }
                    }
                }
                
                if ($cleanupSuccess) {
                    Write-AppxLog -Message "Temporary files cleaned up successfully" -Level 'Debug'
                }
            }
        }
    }
}