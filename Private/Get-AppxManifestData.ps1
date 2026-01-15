<#
.SYNOPSIS
    Parses APPX/MSIX manifest files with full namespace support and validation.

.DESCRIPTION
    Replaces the naive XML parsing from 2016 that ignored namespaces.
    
    Handles:
    - Multiple manifest schema versions (Windows 8.1 through Windows 11)
    - Namespace resolution for all standard APPX namespaces
    - MSIX-specific elements
    - Bundle manifests
    - Dependency extraction
    - Capability enumeration
    - Target device family detection
#>

function Get-AppxManifestData {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory, Position = 0)]
        [ValidateScript({
            if (-not (Test-Path -LiteralPath $_)) {
                throw "Manifest file not found: $_"
            }
            if ((Get-Item -LiteralPath $_).Extension -ne '.xml') {
                throw "File must be an XML manifest"
            }
            $true
        })]
        [string]$ManifestPath,

        [Parameter()]
        [switch]$IncludeDependencies,

        [Parameter()]
        [switch]$IncludeCapabilities,

        [Parameter()]
        [switch]$ValidateSchema
    )

    begin {
        Write-AppxLog -Message "Parsing manifest: $ManifestPath" -Level 'Verbose'
    }

    process {
        try {
            # Load XML with proper error handling
            [xml]$manifest = Get-Content -LiteralPath $ManifestPath -ErrorAction Stop
            
            # Validate root element
            if ($manifest.DocumentElement.LocalName -ne 'Package') {
                throw "Invalid manifest: Root element must be 'Package', found '$($manifest.DocumentElement.LocalName)'"
            }

            # Setup namespace manager for XPath queries
            $nsManager = [System.Xml.XmlNamespaceManager]::new($manifest.NameTable)
            
            # Register common APPX namespaces
            $commonNamespaces = @{
                'appx'         = 'http://schemas.microsoft.com/appx/manifest/foundation/windows10'
                'appx2015'     = 'http://schemas.microsoft.com/appx/manifest/foundation/windows10/2'
                'uap'          = 'http://schemas.microsoft.com/appx/manifest/uap/windows10'
                'uap2'         = 'http://schemas.microsoft.com/appx/manifest/uap/windows10/2'
                'uap3'         = 'http://schemas.microsoft.com/appx/manifest/uap/windows10/3'
                'uap4'         = 'http://schemas.microsoft.com/appx/manifest/uap/windows10/4'
                'uap5'         = 'http://schemas.microsoft.com/appx/manifest/uap/windows10/5'
                'uap6'         = 'http://schemas.microsoft.com/appx/manifest/uap/windows10/6'
                'uap10'        = 'http://schemas.microsoft.com/appx/manifest/uap/windows10/10'
                'mp'           = 'http://schemas.microsoft.com/appx/2014/phone/manifest'
                'build'        = 'http://schemas.microsoft.com/developer/appx/2015/build'
                'rescap'       = 'http://schemas.microsoft.com/appx/manifest/foundation/windows10/restrictedcapabilities'
                'desktop'      = 'http://schemas.microsoft.com/appx/manifest/desktop/windows10'
                'iot'          = 'http://schemas.microsoft.com/appx/manifest/iot/windows10'
                'mobile'       = 'http://schemas.microsoft.com/appx/manifest/mobile/windows10'
                'serverpreview'= 'http://schemas.microsoft.com/appx/manifest/serverpreview/windows10'
            }
            
            # Add fallback for legacy Windows 8.x manifests
            if ($manifest.Package.NamespaceURI -like '*windows8*') {
                $commonNamespaces['appx'] = $manifest.Package.NamespaceURI
            }
            
            foreach ($ns in $commonNamespaces.GetEnumerator()) {
                $nsManager.AddNamespace($ns.Key, $ns.Value)
            }

            # Extract Identity (works with or without namespace prefix)
            $identityNode = $manifest.SelectSingleNode('//appx:Package/appx:Identity', $nsManager)
            if (-not $identityNode) {
                # Try without namespace (legacy)
                $identityNode = $manifest.Package.Identity
            }
            
            if (-not $identityNode) {
                throw "Invalid manifest: Identity element not found"
            }

            # Build base result object
            $result = [PSCustomObject]@{
                PSTypeName              = 'AppxBackup.ManifestData'
                Name                    = $identityNode.Name
                Publisher               = $identityNode.Publisher
                Version                 = $identityNode.Version
                ProcessorArchitecture   = $identityNode.ProcessorArchitecture
                ResourceId              = $identityNode.GetAttribute('ResourceId')
                PublisherDisplayName    = $null
                DisplayName             = $null
                Description             = $null
                Logo                    = $null
                PackageFamilyName       = $null
                PackageFullName         = $null
                ManifestVersion         = $manifest.Package.GetAttribute('xmlns')
                IsBundle                = $false
                IsMSIX                  = $false
                TargetDeviceFamilies    = @()
                Dependencies            = @()
                Capabilities            = @()
                Applications            = @()
                ManifestPath            = $ManifestPath
            }

            # Extract Properties
            $propertiesNode = $manifest.SelectSingleNode('//appx:Package/appx:Properties', $nsManager)
            if ($propertiesNode) {
                $result.DisplayName = $propertiesNode.DisplayName
                $result.PublisherDisplayName = $propertiesNode.PublisherDisplayName
                $result.Description = $propertiesNode.Description
                $result.Logo = $propertiesNode.Logo
            }

            # Detect if this is a bundle manifest
            # Check if Bundle element exists in XML (not Package.Bundle property)
            try {
                $bundleNode = $manifest.SelectSingleNode('/Bundle', $nsManager)
                if ($bundleNode) {
                    $result.IsBundle = $true
                    Write-AppxLog -Message "Detected bundle manifest" -Level 'Debug'
                }
            }
            catch {
                # Not a bundle, continue
                $result.IsBundle = $false
            }

            # Detect MSIX (presence of specific namespaces or features)
            # Safely check if IgnorableNamespaces property exists
            try {
                $ignorableNS = $null
                if ($manifest.Package.PSObject.Properties.Name -contains 'IgnorableNamespaces') {
                    $ignorableNS = $manifest.Package.IgnorableNamespaces
                }
                
                if ($ignorableNS -and $ignorableNS -match 'build|uap10') {
                    $result.IsMSIX = $true
                    Write-AppxLog -Message "Detected MSIX package" -Level 'Debug'
                }
            }
            catch {
                # Not MSIX or property doesn't exist, continue
                $result.IsMSIX = $false
            }

            # Extract Target Device Families
            $prereqNode = $manifest.SelectSingleNode('//appx:Package/appx:Prerequisites', $nsManager)
            if ($prereqNode) {
                $targetFamilies = $prereqNode.SelectNodes('appx:TargetDeviceFamily', $nsManager)
                foreach ($family in $targetFamilies) {
                    $result.TargetDeviceFamilies += [PSCustomObject]@{
                        Name = $family.Name
                        MinVersion = $family.MinVersion
                        MaxVersionTested = $family.MaxVersionTested
                    }
                }
            }

            # Extract Dependencies (if requested)
            if ($IncludeDependencies.IsPresent) {
                $depsNode = $manifest.SelectSingleNode('//appx:Package/appx:Dependencies', $nsManager)
                if ($depsNode) {
                    $packages = $depsNode.SelectNodes('appx:PackageDependency', $nsManager)
                    foreach ($pkg in $packages) {
                        $result.Dependencies += [PSCustomObject]@{
                            Name = $pkg.Name
                            Publisher = $pkg.Publisher
                            MinVersion = $pkg.MinVersion
                        }
                    }
                }
            }

            # Extract Capabilities (if requested)
            if ($IncludeCapabilities.IsPresent) {
                $capsNode = $manifest.SelectSingleNode('//appx:Package/appx:Capabilities', $nsManager)
                if ($capsNode) {
                    $caps = $capsNode.SelectNodes('*')
                    foreach ($cap in $caps) {
                        $capName = $cap.Name
                        if ($cap.LocalName -eq 'Capability') {
                            $capName = $cap.Name
                        }
                        elseif ($cap.LocalName -eq 'DeviceCapability') {
                            $capName = "Device: $($cap.Name)"
                        }
                        $result.Capabilities += $capName
                    }
                }
            }

            # Extract Applications
            $appsNode = $manifest.SelectSingleNode('//appx:Package/appx:Applications', $nsManager)
            if ($appsNode) {
                $apps = $appsNode.SelectNodes('appx:Application', $nsManager)
                foreach ($app in $apps) {
                    $result.Applications += [PSCustomObject]@{
                        Id = $app.Id
                        Executable = $app.Executable
                        EntryPoint = $app.EntryPoint
                    }
                }
            }

            # Calculate Package Family Name (for reference)
            try {
                $publisherHash = [System.Security.Cryptography.SHA256]::Create().ComputeHash(
                    [System.Text.Encoding]::Unicode.GetBytes($result.Publisher)
                )
                $hashString = [System.BitConverter]::ToString($publisherHash).Replace('-', '').Substring(0, 13)
                $result.PackageFamilyName = "$($result.Name)_$hashString"
            }
            catch {
                Write-AppxLog -Message "Failed to calculate PackageFamilyName: $_" -Level 'Warning'
            }

            Write-AppxLog -Message "Manifest parsed successfully: $($result.Name) v$($result.Version)" -Level 'Verbose'
            
            return $result
        }
        catch {
            Write-AppxLog -Message "Failed to parse manifest: $_" -Level 'Error'
            throw
        }
    }
}