<#
.SYNOPSIS
    Resolves an XML manifest node using a multi-tier fallback strategy.

.DESCRIPTION
    Attempts to locate an XML node in an APPX manifest using three strategies:
    1. XPath with namespace manager
    2. Direct property access (no namespace)
    3. Element tag name search (namespace-agnostic)

    Designed for maximum compatibility across APPX manifest schema versions
    (Windows 8.1 through Windows 11).

.PARAMETER Manifest
    The loaded XML document (AppxManifest.xml).

.PARAMETER NodeName
    The name of the node to find (e.g., 'Identity', 'Properties', 'Dependencies').

.PARAMETER NamespaceManager
    The XmlNamespaceManager configured for the manifest.

.PARAMETER XPathExpression
    Optional explicit XPath expression. If not specified, defaults to '//appx:Package/appx:<NodeName>'.

.PARAMETER ParentPropertyPath
    Optional dot-separated path for direct property access.
    If not specified, defaults to 'Package.<NodeName>'.

.OUTPUTS
    [System.Xml.XmlNode] The resolved node, or $null if not found via any strategy.

.NOTES
    Author: DeltaGa
    Version: 2.0.1
#>

function Resolve-AppxManifestNode {
    [CmdletBinding()]
    [OutputType([System.Xml.XmlNode])]
    param(
        [Parameter(Mandatory)]
        [System.Xml.XmlDocument]$Manifest,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$NodeName,

        [Parameter(Mandatory)]
        [System.Xml.XmlNamespaceManager]$NamespaceManager,

        [Parameter()]
        [string]$XPathExpression,

        [Parameter()]
        [string]$ParentPropertyPath
    )

    process {
        $resolvedNode = $null

        # Derive defaults if not explicitly specified
        if (-not $XPathExpression) {
            $XPathExpression = "//appx:Package/appx:$NodeName"
        }
        if (-not $ParentPropertyPath) {
            $ParentPropertyPath = "Package.$NodeName"
        }

        # Strategy 1: XPath with namespace
        try {
            $resolvedNode = $Manifest.SelectSingleNode($XPathExpression, $NamespaceManager)
        }
        catch {
            Write-AppxLog -Message "XPath with namespace failed for ${NodeName}: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
        }

        # Strategy 2: Direct property access (no namespace)
        if ($null -eq $resolvedNode) {
            try {
                $pathParts = $ParentPropertyPath.Split('.')
                $currentNode = $Manifest

                foreach ($part in @($pathParts)) {
                    $currentNode = $currentNode.$part
                    if ($null -eq $currentNode) { break }
                }

                $resolvedNode = $currentNode
            }
            catch {
                Write-AppxLog -Message "Direct property access failed for ${NodeName}: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
            }
        }

        # Strategy 3: Search by element name only (namespace-agnostic)
        if ($null -eq $resolvedNode) {
            try {
                $resolvedNode = $Manifest.GetElementsByTagName($NodeName) | Select-Object -First 1
            }
            catch {
                Write-AppxLog -Message "Element search failed for ${NodeName}: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Debug'
            }
        }

        return $resolvedNode
    }
}
