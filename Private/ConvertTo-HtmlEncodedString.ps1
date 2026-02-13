<#
.SYNOPSIS
    HTML-encodes a string for safe inclusion in HTML output.

.DESCRIPTION
    Prevents XSS by encoding special HTML characters. Uses System.Web.HttpUtility
    when available, falls back to manual replacement for core entities.

.PARAMETER Text
    The string to HTML-encode.

.OUTPUTS
    [string] The HTML-encoded string.

.NOTES
    Author: DeltaGa
    Version: 2.0.1
#>

function ConvertTo-HtmlEncodedString {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(ValueFromPipeline)]
        [string]$Text
    )

    process {
        if ([string]::IsNullOrEmpty($Text)) { return '' }

        # Try System.Web first (loaded on demand)
        try {
            if (-not ([System.Management.Automation.PSTypeName]'System.Web.HttpUtility').Type) {
                Add-Type -AssemblyName System.Web -ErrorAction Stop
            }
            return [System.Web.HttpUtility]::HtmlEncode($Text)
        }
        catch {
            # Fallback: manual encoding of core HTML entities
            return $Text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;').Replace("'", '&#39;')
        }
    }
}
