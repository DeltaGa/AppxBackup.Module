<#
.SYNOPSIS
    Parses APPX/MSIX manifest files with comprehensive edge-case handling.

.DESCRIPTION
    Enterprise-grade manifest parser that handles:
    - Multiple manifest schema versions (Windows 8.1 through Windows 11)
    - Missing or non-standard namespaces
    - Optional elements and properties
    - Malformed or incomplete manifests
    - Bundle manifests
    - MSIX-specific elements
    
    Uses multi-tier fallback strategy for maximum compatibility.
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
        
        # Helper function for safe XML node access
        function Get-SafeXmlValue {
            param(
                [Parameter(Mandatory)]
                [System.Xml.XmlNode]$Node,
                
                [Parameter(Mandatory)]
                [string]$PropertyName,
                
                [Parameter()]
                $DefaultValue = $null
            )
            
            try {
                if ($Node -and $Node.$PropertyName) {
                    return $Node.$PropertyName
                }
            }
            catch {
                # Property doesn't exist
            }
            
            return $DefaultValue
        }
        
        # Helper function for safe attribute access
        function Get-SafeAttribute {
            param(
                [Parameter(Mandatory)]
                [System.Xml.XmlNode]$Node,
                
                [Parameter(Mandatory)]
                [string]$AttributeName,
                
                [Parameter()]
                $DefaultValue = $null
            )
            
            try {
                if ($Node -and $Node.HasAttribute($AttributeName)) {
                    return $Node.GetAttribute($AttributeName)
                }
            }
            catch {
                # Attribute doesn't exist
            }
            
            return $DefaultValue
        }
    }

    process {
        try {
            # Load XML with proper error handling
            [xml]$manifest = Get-Content -LiteralPath $ManifestPath -ErrorAction Stop
            
            # Validate root element
            if (-not $manifest.DocumentElement) {
                throw "Invalid manifest: No document element found"
            }
            
            $rootName = $manifest.DocumentElement.LocalName
            if ($rootName -ne 'Package' -and $rootName -ne 'Bundle') {
                throw "Invalid manifest: Root element must be 'Package' or 'Bundle', found '$rootName'"
            }

            # Detect actual namespace used in manifest
            $defaultNamespace = $manifest.DocumentElement.NamespaceURI
            Write-AppxLog -Message "Detected namespace: $defaultNamespace" -Level 'Debug'

            # Setup namespace manager for XPath queries
            $nsManager = [System.Xml.XmlNamespaceManager]::new($manifest.NameTable)
            
            # Register the actual default namespace
            if ($defaultNamespace) {
                $nsManager.AddNamespace('appx', $defaultNamespace)
            }
            else {
                # No namespace declaration - use empty namespace
                $nsManager.AddNamespace('appx', '')
            }
            
            # Register all common APPX namespaces (for completeness)
            $commonNamespaces = @{
                'appx2015'      = 'http://schemas.microsoft.com/appx/manifest/foundation/windows10/2'
                'uap'           = 'http://schemas.microsoft.com/appx/manifest/uap/windows10'
                'uap2'          = 'http://schemas.microsoft.com/appx/manifest/uap/windows10/2'
                'uap3'          = 'http://schemas.microsoft.com/appx/manifest/uap/windows10/3'
                'uap4'          = 'http://schemas.microsoft.com/appx/manifest/uap/windows10/4'
                'uap5'          = 'http://schemas.microsoft.com/appx/manifest/uap/windows10/5'
                'uap6'          = 'http://schemas.microsoft.com/appx/manifest/uap/windows10/6'
                'uap10'         = 'http://schemas.microsoft.com/appx/manifest/uap/windows10/10'
                'mp'            = 'http://schemas.microsoft.com/appx/2014/phone/manifest'
                'build'         = 'http://schemas.microsoft.com/developer/appx/2015/build'
                'rescap'        = 'http://schemas.microsoft.com/appx/manifest/foundation/windows10/restrictedcapabilities'
                'desktop'       = 'http://schemas.microsoft.com/appx/manifest/desktop/windows10'
                'iot'           = 'http://schemas.microsoft.com/appx/manifest/iot/windows10'
                'mobile'        = 'http://schemas.microsoft.com/appx/manifest/mobile/windows10'
                'serverpreview' = 'http://schemas.microsoft.com/appx/manifest/serverpreview/windows10'
            }
            
            foreach ($ns in $commonNamespaces.GetEnumerator()) {
                try {
                    $nsManager.AddNamespace($ns.Key, $ns.Value)
                }
                catch {
                    # Namespace already exists or error - continue
                }
            }

            # Multi-tier strategy for finding Identity node
            $identityNode = $null
            
            # Strategy 1: XPath with namespace
            try {
                $identityNode = $manifest.SelectSingleNode('//appx:Package/appx:Identity', $nsManager)
            }
            catch {
                Write-AppxLog -Message "XPath with namespace failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
            }
            
            # Strategy 2: Direct property access (no namespace)
            if ($null -eq $identityNode) {
                try {
                    $identityNode = $manifest.Package.Identity
                }
                catch {
                    Write-AppxLog -Message "Direct property access failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
                }
            }
            
            # Strategy 3: Search by element name only (namespace-agnostic)
            if ($null -eq $identityNode) {
                try {
                    $identityNode = $manifest.GetElementsByTagName('Identity') | Select-Object -First 1
                }
                catch {
                    Write-AppxLog -Message "Element search failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
                }
            }
            
            if ($null -eq $identityNode) {
                throw "Invalid manifest: Identity element not found using any strategy"
            }

            # Build base result object with safe defaults
            $result = [PSCustomObject]@{
                PSTypeName              = 'AppxBackup.ManifestData'
                Name                    = Get-SafeAttribute -Node $identityNode -AttributeName 'Name' -DefaultValue 'Unknown'
                Publisher               = Get-SafeAttribute -Node $identityNode -AttributeName 'Publisher' -DefaultValue 'Unknown'
                Version                 = Get-SafeAttribute -Node $identityNode -AttributeName 'Version' -DefaultValue '0.0.0.0'
                ProcessorArchitecture   = Get-SafeAttribute -Node $identityNode -AttributeName 'ProcessorArchitecture' -DefaultValue 'neutral'
                ResourceId              = Get-SafeAttribute -Node $identityNode -AttributeName 'ResourceId'
                PublisherDisplayName    = $null
                DisplayName             = $null
                Description             = $null
                Logo                    = $null
                PackageFamilyName       = $null
                PackageFullName         = $null
                ManifestVersion         = $defaultNamespace
                IsBundle                = ($rootName -eq 'Bundle')
                IsMSIX                  = $false
                TargetDeviceFamilies    = @()
                Dependencies            = @()
                Capabilities            = @()
                Applications            = @()
                ManifestPath            = $ManifestPath
            }

            # Extract Properties with multi-tier strategy
            $propertiesNode = $null
            
            # Strategy 1: XPath with namespace
            try {
                $propertiesNode = $manifest.SelectSingleNode('//appx:Package/appx:Properties', $nsManager)
            }
            catch {
                Write-AppxLog -Message "XPath for Properties failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
            }
            
            # Strategy 2: Direct access
            if ($null -eq $propertiesNode) {
                try {
                    $propertiesNode = $manifest.Package.Properties
                }
                catch {
                    Write-AppxLog -Message "Direct Properties access failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
                }
            }
            
            # Strategy 3: Element search
            if ($null -eq $propertiesNode) {
                try {
                    $propertiesNode = $manifest.GetElementsByTagName('Properties') | Select-Object -First 1
                }
                catch {
                    Write-AppxLog -Message "Element search for Properties failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
                }
            }
            
            # Extract properties if found
            if ($propertiesNode) {
                $result.DisplayName = Get-SafeXmlValue -Node $propertiesNode -PropertyName 'DisplayName'
                $result.PublisherDisplayName = Get-SafeXmlValue -Node $propertiesNode -PropertyName 'PublisherDisplayName'
                $result.Description = Get-SafeXmlValue -Node $propertiesNode -PropertyName 'Description'
                $result.Logo = Get-SafeXmlValue -Node $propertiesNode -PropertyName 'Logo'
            }
            else {
                Write-AppxLog -Message "Properties node not found - using defaults" -Level 'Debug'
            }

            # Detect MSIX (presence of specific namespaces or features)
            try {
                $ignorableNS = Get-SafeAttribute -Node $manifest.Package -AttributeName 'IgnorableNamespaces'
                if ($ignorableNS -and ($ignorableNS -match 'build|uap10')) {
                    $result.IsMSIX = $true
                    Write-AppxLog -Message "Detected MSIX package" -Level 'Debug'
                }
                
                # Also check for MSIX-specific version format
                if ($result.Version -match '^\d+\.\d+\.\d+\.\d+$') {
                    $versionParts = $result.Version -split '\.'
                    if ([int]$versionParts[0] -ge 1) {
                        $result.IsMSIX = $true
                    }
                }
            }
            catch {
                Write-AppxLog -Message "MSIX detection failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
            }

            # Extract Dependencies (if requested)
            if ($IncludeDependencies.IsPresent) {
                $dependencies = @()
                
                # Multi-tier strategy for Dependencies
                $dependenciesNode = $null
                try {
                    $dependenciesNode = $manifest.SelectSingleNode('//appx:Package/appx:Dependencies', $nsManager)
                }
                catch {}
                
                if ($null -eq $dependenciesNode) {
                    try {
                        $dependenciesNode = $manifest.Package.Dependencies
                    }
                    catch {}
                }
                
                if ($null -eq $dependenciesNode) {
                    try {
                        $dependenciesNode = $manifest.GetElementsByTagName('Dependencies') | Select-Object -First 1
                    }
                    catch {}
                }
                
                if ($dependenciesNode) {
                    # Try to get PackageDependency elements
                    try {
                        $depElements = $dependenciesNode.SelectNodes('.//appx:PackageDependency', $nsManager)
                        if ($null -eq $depElements) {
                            $depElements = $dependenciesNode.PackageDependency
                        }
                        if ($null -eq $depElements) {
                            $depElements = $dependenciesNode.GetElementsByTagName('PackageDependency')
                        }
                        
                        foreach ($dep in @($depElements)) {
                            $dependencies += [PSCustomObject]@{
                                Name = Get-SafeAttribute -Node $dep -AttributeName 'Name' -DefaultValue 'Unknown'
                                Publisher = Get-SafeAttribute -Node $dep -AttributeName 'Publisher'
                                MinVersion = Get-SafeAttribute -Node $dep -AttributeName 'MinVersion' -DefaultValue '0.0.0.0'
                            }
                        }
                    }
                    catch {
                        Write-AppxLog -Message "Failed to extract dependencies: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
                    }
                }
                
                $result.Dependencies = $dependencies
                Write-AppxLog -Message "Extracted $($dependencies.Count) dependencies" -Level 'Debug'
            }

            # Extract Capabilities (if requested)
            if ($IncludeCapabilities.IsPresent) {
                $capabilities = @()
                
                # Multi-tier strategy for Capabilities
                $capabilitiesNode = $null
                try {
                    $capabilitiesNode = $manifest.SelectSingleNode('//appx:Package/appx:Capabilities', $nsManager)
                }
                catch {}
                
                if ($null -eq $capabilitiesNode) {
                    try {
                        $capabilitiesNode = $manifest.Package.Capabilities
                    }
                    catch {}
                }
                
                if ($null -eq $capabilitiesNode) {
                    try {
                        $capabilitiesNode = $manifest.GetElementsByTagName('Capabilities') | Select-Object -First 1
                    }
                    catch {}
                }
                
                if ($capabilitiesNode) {
                    # Get all capability elements
                    try {
                        $capElements = $capabilitiesNode.ChildNodes | Where-Object { $_.LocalName -match 'Capability$' }
                        
                        foreach ($cap in @($capElements)) {
                            $capName = Get-SafeAttribute -Node $cap -AttributeName 'Name'
                            if ($capName) {
                                $capabilities += $capName
                            }
                        }
                    }
                    catch {
                        Write-AppxLog -Message "Failed to extract capabilities: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
                    }
                }
                
                $result.Capabilities = $capabilities
                Write-AppxLog -Message "Extracted $($capabilities.Count) capabilities" -Level 'Debug'
            }

            # Extract Applications
            try {
                $appsNode = $null
                
                # Multi-tier strategy
                try {
                    $appsNode = $manifest.SelectSingleNode('//appx:Package/appx:Applications', $nsManager)
                }
                catch {}
                
                if ($null -eq $appsNode) {
                    try {
                        $appsNode = $manifest.Package.Applications
                    }
                    catch {}
                }
                
                if ($null -eq $appsNode) {
                    try {
                        $appsNode = $manifest.GetElementsByTagName('Applications') | Select-Object -First 1
                    }
                    catch {}
                }
                
                if ($appsNode) {
                    $applications = @()
                    
                    $appElements = $appsNode.ChildNodes | Where-Object { $_.LocalName -eq 'Application' }
                    
                    foreach ($app in @($appElements)) {
                        $applications += [PSCustomObject]@{
                            Id = Get-SafeAttribute -Node $app -AttributeName 'Id' -DefaultValue 'App'
                            Executable = Get-SafeAttribute -Node $app -AttributeName 'Executable'
                            EntryPoint = Get-SafeAttribute -Node $app -AttributeName 'EntryPoint'
                        }
                    }
                    
                    $result.Applications = $applications
                    Write-AppxLog -Message "Extracted $($applications.Count) applications" -Level 'Debug'
                }
            }
            catch {
                Write-AppxLog -Message "Failed to extract applications: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
            }

            # Calculate Package Family Name
            try {
                if ($result.Name -and $result.Publisher) {
                    # Simple hash-based calculation (approximation)
                    $publisherHash = [System.Security.Cryptography.SHA256]::Create().ComputeHash(
                        [System.Text.Encoding]::Unicode.GetBytes($result.Publisher)
                    )
                    $hashString = [System.BitConverter]::ToString($publisherHash).Replace('-', '').Substring(0, 13).ToLower()
                    $result.PackageFamilyName = "$($result.Name)_$hashString"
                }
            }
            catch {
                Write-AppxLog -Message "Failed to calculate PackageFamilyName: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
            }

            # Calculate Package Full Name
            try {
                if ($result.Name -and $result.Version -and $result.ProcessorArchitecture -and $result.PackageFamilyName) {
                    $familyId = $result.PackageFamilyName.Split('_')[1]
                    $resourceSuffix = if ($result.ResourceId) { "_$($result.ResourceId)" } else { "" }
                    $result.PackageFullName = "$($result.Name)_$($result.Version)_$($result.ProcessorArchitecture)$resourceSuffix`_$familyId"
                }
            }
            catch {
                Write-AppxLog -Message "Failed to calculate PackageFullName: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
            }

            Write-AppxLog -Message "Manifest parsed successfully: $($result.Name) v$($result.Version)" -Level 'Verbose'
            
            return $result
        }
        catch {
            Write-AppxLog -Message "Failed to parse manifest: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'
            Write-AppxLog -Message "Stack trace: $($_.ScriptStackTrace)" -Level 'Debug'
            throw
        }
    }
}