<#
.SYNOPSIS
    Exports comprehensive dependency information for an APPX/MSIX package.

.DESCRIPTION
    Analyzes and exports detailed dependency information to various formats.
    Useful for documentation, compliance, and deployment planning.
    
    Features:
    - Multiple output formats (JSON, XML, CSV, HTML)
    - Recursive dependency resolution
    - Framework package identification
    - Missing dependency detection
    - Compatibility analysis

.PARAMETER PackagePath
    Path to the package directory or .appx file to analyze.

.PARAMETER OutputPath
    Path where the dependency report will be saved.
    File extension determines format if Format not specified.

.PARAMETER Format
    Output format: JSON, XML, CSV, or HTML.
    If not specified, inferred from OutputPath extension.

.PARAMETER Recursive
    If specified, recursively analyzes dependencies of dependencies.

.PARAMETER IncludeOptional
    If specified, includes optional framework dependencies.

.PARAMETER MaxDepth
    Maximum recursion depth for dependency analysis.
    Default: 3

.EXAMPLE
    Export-AppxDependencies -PackagePath "C:\Program Files\WindowsApps\MyApp" -OutputPath "C:\Reports\deps.json"
    
    Exports dependencies to JSON format

.EXAMPLE
    Export-AppxDependencies -PackagePath "C:\Backups\MyApp.appx" -OutputPath "C:\Reports\deps.html" -Recursive -IncludeOptional
    
    Creates comprehensive HTML report with all dependencies

.OUTPUTS
    File path to the generated report

.NOTES
    For .appx files, temporarily extracts manifest for analysis.
    Supports both installed packages and backup files.
#>

function Export-AppxDependencies {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [Alias('InstallLocation', 'FullName', 'Path')]
        [string]$PackagePath,

        [Parameter(Mandatory, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$OutputPath,

        [Parameter()]
        [ValidateSet('JSON', 'XML', 'CSV', 'HTML')]
        [string]$Format,

        [Parameter()]
        [switch]$Recursive,

        [Parameter()]
        [switch]$IncludeOptional,

        [Parameter()]
        [ValidateRange(1, 10)]
        [int]$MaxDepth = 3
    )

    begin {
        Write-AppxLog -Message "Exporting dependency information" -Level 'Verbose'
        
        # Auto-detect format from extension if not specified
        if ($null -eq $Format) {
            $extension = [System.IO.Path]::GetExtension($OutputPath).TrimStart('.')
            $Format = switch ($extension.ToUpper()) {
                'JSON' { 'JSON' }
                'XML'  { 'XML' }
                'CSV'  { 'CSV' }
                'HTML' { 'HTML' }
                'HTM'  { 'HTML' }
                default { 'JSON' }
            }
            
            Write-AppxLog -Message "Auto-detected format: $Format" -Level 'Debug'
        }
    }

    process {
        try {
            # Check if OutputPath is a directory and auto-generate filename
            $isDirectory = $false
            if (Test-Path -LiteralPath $OutputPath -PathType Container) {
                $isDirectory = $true
            }
            elseif ($OutputPath.EndsWith('\') -or $OutputPath.EndsWith('/')) {
                $isDirectory = $true
            }
            
            # Generate output file path
            if ($isDirectory) {
                # Extract package name from path for filename
                $pkgName = if (Test-Path -LiteralPath $PackagePath -PathType Container) {
                    Split-Path -Path $PackagePath -Leaf
                } else {
                    [System.IO.Path]::GetFileNameWithoutExtension($PackagePath)
                }
                
                # Set default format if not specified (can't auto-detect from directory)
                if ([string]::IsNullOrEmpty($Format)) {
                    $Format = 'JSON'
                    Write-AppxLog -Message "No format specified, defaulting to JSON" -Level 'Debug'
                }
                
                # Use Format to determine extension
                $extension = $Format.ToLower()
                
                $fileName = "${pkgName}_Dependencies.$extension"
                $outputPath = [System.IO.Path]::Combine($OutputPath, $fileName)
                
                Write-AppxLog -Message "Auto-generated filename: $fileName" -Level 'Debug'
            }
            else {
                $outputPath = $OutputPath
            }
            
            # Validate output path
            $outputPath = ConvertTo-SecureFilePath -Path $outputPath -ResolveRelative
            $outputDir = Split-Path -Path $outputPath -Parent
            
            if ($outputDir -and -not (Test-Path -LiteralPath $outputDir)) {
                [void](New-Item -Path $outputDir -ItemType Directory -Force)
            }

            # Determine if PackagePath is a file or directory
            $isFile = (Test-Path -LiteralPath $PackagePath -PathType Leaf)
            $tempDir = $null
            $workingPath = $null

            if ($isFile) {
                # Extract manifest from package file
                Write-AppxLog -Message "Extracting manifest from package file" -Level 'Verbose'
                
                $tempDir = [System.IO.Path]::Combine($env:TEMP, "AppxDepsExport_$(New-Guid)")
                [void](New-Item -Path $tempDir -ItemType Directory -Force)
                
                Add-Type -AssemblyName 'System.IO.Compression.FileSystem'
                $archive = [System.IO.Compression.ZipFile]::OpenRead($PackagePath)
                
                try {
                    $manifestEntry = $archive.Entries | Where-Object { $_.Name -eq 'AppxManifest.xml' } | Select-Object -First 1
                    
                    if ($null -eq $manifestEntry) {
                        throw "AppxManifest.xml not found in package"
                    }
                    
                    $manifestPath = [System.IO.Path]::Combine($tempDir, 'AppxManifest.xml')
                    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($manifestEntry, $manifestPath, $true)
                    
                    $workingPath = $tempDir
                }
                finally {
                    $archive.Dispose()
                }
            }
            else {
                # Use directory directly
                $workingPath = ConvertTo-SecureFilePath -Path $PackagePath -MustExist -PathType Directory
            }

            # Resolve dependencies
            Write-AppxLog -Message "Resolving dependencies..." -Level 'Verbose'
            
            $depResult = Resolve-AppxDependencies -PackagePath $workingPath `
                -Recursive:$Recursive `
                -IncludeOptional:$IncludeOptional `
                -MaxDepth $MaxDepth

            Write-AppxLog -Message "Found $($depResult.TotalDependencies) dependencies ($($depResult.MissingCount) missing)" -Level 'Info'

            # Export to specified format
            Write-AppxLog -Message "Exporting to $Format format: $outputPath" -Level 'Verbose'
            
            switch ($Format) {
                'JSON' {
                    # JSON export
                    $jsonData = [PSCustomObject]@{
                        ExportDate = [DateTime]::Now.ToString('o')
                        PackageName = $depResult.PackageName
                        PackageVersion = $depResult.PackageVersion
                        SourcePath = $PackagePath
                        Summary = [PSCustomObject]@{
                            TotalDependencies = $depResult.TotalDependencies
                            InstalledCount = $depResult.InstalledCount
                            MissingCount = $depResult.MissingCount
                            FrameworkCount = $depResult.FrameworkCount
                        }
                        Dependencies = $depResult.Dependencies
                    }
                    
                    $jsonData | ConvertTo-Json -Depth 10 | Out-File -FilePath $outputPath -Encoding UTF8
                }
                
                'XML' {
                    # XML export
                    $xmlDoc = [System.Xml.XmlDocument]::new()
                    $root = $xmlDoc.CreateElement('DependencyReport')
                    $xmlDoc.AppendChild($root) | Out-Null
                    
                    # Metadata
                    $metadata = $xmlDoc.CreateElement('Metadata')
                    $metadata.SetAttribute('ExportDate', [DateTime]::Now.ToString('o'))
                    $metadata.SetAttribute('PackageName', $depResult.PackageName)
                    $metadata.SetAttribute('PackageVersion', $depResult.PackageVersion)
                    $root.AppendChild($metadata) | Out-Null
                    
                    # Summary
                    $summary = $xmlDoc.CreateElement('Summary')
                    $summary.SetAttribute('Total', $depResult.TotalDependencies)
                    $summary.SetAttribute('Installed', $depResult.InstalledCount)
                    $summary.SetAttribute('Missing', $depResult.MissingCount)
                    $summary.SetAttribute('Framework', $depResult.FrameworkCount)
                    $root.AppendChild($summary) | Out-Null
                    
                    # Dependencies
                    $dependencies = $xmlDoc.CreateElement('Dependencies')
                    
                    foreach ($dep in @($depResult.Dependencies)) {
                        $depNode = $xmlDoc.CreateElement('Dependency')
                        $depNode.SetAttribute('Name', $dep.Name)
                        $depNode.SetAttribute('Publisher', $dep.Publisher)
                        $depNode.SetAttribute('MinVersion', $dep.MinVersion)
                        $depNode.SetAttribute('Type', $dep.DependencyType)
                        $depNode.SetAttribute('IsInstalled', $dep.IsInstalled)
                        $depNode.SetAttribute('IsOptional', $dep.IsOptional)
                        
                        if ($dep.InstalledVersion) {
                            $depNode.SetAttribute('InstalledVersion', $dep.InstalledVersion)
                        }
                        
                        $dependencies.AppendChild($depNode) | Out-Null
                    }
                    
                    $root.AppendChild($dependencies) | Out-Null
                    
                    $xmlDoc.Save($outputPath)
                }
                
                'CSV' {
                    # CSV export
                    $csvData = $depResult.Dependencies | ForEach-Object {
                        [PSCustomObject]@{
                            Name = $_.Name
                            Publisher = $_.Publisher
                            MinVersion = $_.MinVersion
                            DependencyType = $_.DependencyType
                            IsInstalled = $_.IsInstalled
                            InstalledVersion = $_.InstalledVersion
                            IsOptional = $_.IsOptional
                            Architecture = $_.Architecture
                            ResolvedPath = $_.ResolvedPath
                        }
                    }
                    
                    $csvData | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
                }
                
                'HTML' {
                    # HTML export with styling
                    $html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Dependency Report - $($depResult.PackageName)</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 20px; background: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #0078D4; margin-bottom: 10px; }
        .metadata { background: #f8f9fa; padding: 15px; border-radius: 5px; margin: 20px 0; }
        .metadata p { margin: 5px 0; color: #666; }
        .summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin: 20px 0; }
        .summary-card { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; border-radius: 8px; text-align: center; }
        .summary-card.warning { background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); }
        .summary-card h3 { font-size: 2em; margin-bottom: 5px; }
        .summary-card p { opacity: 0.9; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        thead { background: #0078D4; color: white; }
        th { padding: 12px; text-align: left; font-weight: 600; }
        td { padding: 10px; border-bottom: 1px solid #ddd; }
        tr:hover { background: #f8f9fa; }
        .status-installed { color: #28a745; font-weight: bold; }
        .status-missing { color: #dc3545; font-weight: bold; }
        .badge { display: inline-block; padding: 4px 8px; border-radius: 3px; font-size: 0.85em; font-weight: 600; }
        .badge-framework { background: #6f42c1; color: white; }
        .badge-package { background: #007bff; color: white; }
        .badge-optional { background: #ffc107; color: #333; }
        .footer { margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd; text-align: center; color: #999; font-size: 0.9em; }
    </style>
</head>
<body>
    <div class="container">
        <h1>[CHAR_128230] Dependency Report</h1>
        
        <div class="metadata">
            <p><strong>Package:</strong> $($depResult.PackageName) v$($depResult.PackageVersion)</p>
            <p><strong>Source:</strong> $PackagePath</p>
            <p><strong>Export Date:</strong> $([DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))</p>
        </div>
        
        <div class="summary">
            <div class="summary-card">
                <h3>$($depResult.TotalDependencies)</h3>
                <p>Total Dependencies</p>
            </div>
            <div class="summary-card">
                <h3>$($depResult.InstalledCount)</h3>
                <p>Installed</p>
            </div>
            <div class="summary-card warning">
                <h3>$($depResult.MissingCount)</h3>
                <p>Missing</p>
            </div>
            <div class="summary-card">
                <h3>$($depResult.FrameworkCount)</h3>
                <p>Frameworks</p>
            </div>
        </div>
        
        <h2 style="margin-top: 30px; color: #333;">Dependency Details</h2>
        <table>
            <thead>
                <tr>
                    <th>Name</th>
                    <th>Type</th>
                    <th>Min Version</th>
                    <th>Status</th>
                    <th>Installed Version</th>
                    <th>Architecture</th>
                </tr>
            </thead>
            <tbody>
$(
    $depResult.Dependencies | ForEach-Object {
        $statusClass = if ($_.IsInstalled) { 'status-installed' } else { 'status-missing' }
        $statusText = if ($_.IsInstalled) { '[CHECK] Installed' } else { '[X] Missing' }
        $typeClass = if ($_.DependencyType -eq 'Framework') { 'badge-framework' } else { 'badge-package' }
        $optionalBadge = if ($_.IsOptional) { '<span class="badge badge-optional">Optional</span>' } else { '' }
        
        @"
                <tr>
                    <td><strong>$($_.Name)</strong></td>
                    <td><span class="badge $typeClass">$($_.DependencyType)</span> $optionalBadge</td>
                    <td>$($_.MinVersion)</td>
                    <td class="$statusClass">$statusText</td>
                    <td>$($_.InstalledVersion)</td>
                    <td>$($_.Architecture)</td>
                </tr>
"@
    }
)
            </tbody>
        </table>
        
        <div class="footer">
            <p>Generated by AppxBackup v2.0.0 | DeltaGa</p>
        </div>
    </div>
</body>
</html>
"@
                    
                    $html | Out-File -FilePath $outputPath -Encoding UTF8
                }
                
                default {
                    # Fallback to JSON if Format is unrecognized or null
                    Write-AppxLog -Message "Unrecognized format '$Format', defaulting to JSON" -Level 'Warning'
                    
                    $jsonData = [PSCustomObject]@{
                        ExportDate = [DateTime]::Now.ToString('o')
                        PackageName = $depResult.PackageName
                        PackageVersion = $depResult.PackageVersion
                        SourcePath = $PackagePath
                        Summary = [PSCustomObject]@{
                            TotalDependencies = $depResult.TotalDependencies
                            InstalledCount = $depResult.InstalledCount
                            MissingCount = $depResult.MissingCount
                            FrameworkCount = $depResult.FrameworkCount
                        }
                        Dependencies = $depResult.Dependencies
                    }
                    
                    $jsonData | ConvertTo-Json -Depth 10 | Out-File -FilePath $outputPath -Encoding UTF8
                }
            }

            Write-AppxLog -Message "Export complete: $outputPath" -Level 'Info'
            Write-Host "`n[OK] Dependency report exported successfully" -ForegroundColor Green
            Write-Host "[CHAR_128196] Location: $outputPath" -ForegroundColor Gray
            
            return $outputPath
        }
        catch {
            Write-AppxLog -Message "Failed to export dependencies: $_" -Level 'Error'
            Write-AppxLog -Message "StackTrace: $($_.ScriptStackTrace)" -Level 'Debug'
            throw
        }
        finally {
            # Cleanup temp directory
            if ($tempDir -and (Test-Path -LiteralPath $tempDir)) {
                Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
}