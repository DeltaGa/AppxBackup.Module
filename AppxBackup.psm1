#Requires -Version 5.1

<#
.SYNOPSIS
    AppxBackup Module - Enterprise-grade Windows Application Package backup toolkit

.DESCRIPTION
    Modern replacement for 2016-era APPX backup solutions.
    Provides comprehensive, secure, and performant package backup/restore capabilities.

.NOTES
    Name: AppxBackup
    Author: DeltaGa
    Version: 2.0.1
    LastModified: 2026-01-13
    
    Requires:
    - PowerShell 5.1+ (7.4+ recommended)
    - Windows 10 1809+ or Windows 11
    - Administrator privileges for certificate operations
    
    Architecture:
    - Public: User-facing cmdlets
    - Private: Internal helper functions
    - Zero external dependencies (SDK tools optional)
#>

#region Module Initialization

# Strict mode for catching errors early
Set-StrictMode -Version Latest

# Module-scoped variables
$script:ModuleRoot = $PSScriptRoot
$script:LogPath = Join-Path $env:TEMP "AppxBackup_$(Get-Date -Format 'yyyyMMdd').log"
$script:ToolCache = @{}
$script:PackageCache = @{}
$script:ConfigCache = @{}

# Module configuration - will be populated from ModuleDefaults.json after functions load
$script:AppxBackupConfig = @{
    MaxLogSizeMB = 10  # Temporary default, will be overridden
    DefaultCertificateValidityYears = 3
    DefaultHashAlgorithm = 'SHA256'
    DefaultKeyLength = 4096
    EnableProgressIndicators = $true
    EnableVerboseLogging = $false
    MaxParallelOperations = 4
    PackageValidationLevel = 'Standard'
    CertificateStorageLocation = 'Cert:\CurrentUser\My'
    TempDirectoryPath = $null
    RetryAttempts = 3
    RetryDelaySeconds = 2
    TimeoutSeconds = 3600
}

#endregion

#region Load Private Functions

Write-Verbose "Loading private functions..."

$privateFiles = @(
    'Get-AppxConfiguration.ps1',
    'Get-AppxDefault.ps1',
    'Invoke-ProcessSafely.ps1',
    'Write-AppxLog.ps1',
    'Get-AppxManifestData.ps1',
    'New-AppxPackageInternal.ps1',
    'Test-AppxToolAvailability.ps1',
    'Resolve-AppxDependencies.ps1',
    'ConvertTo-SecureFilePath.ps1',
    'New-AppxBackupZipArchive.ps1',
    'New-AppxBackupManifest.ps1',
    'New-AppxDependencyCertificate.ps1'
)

foreach ($file in $privateFiles) {
    $filePath = Join-Path (Join-Path $script:ModuleRoot 'Private') $file
    
    if (Test-Path $filePath) {
        try {
            . $filePath
            Write-Verbose "Loaded: $file"
        }
        catch {
            Write-Error "Failed to load private function '$file': $_"
            throw
        }
    }
    else {
        Write-Warning "Private function file not found: $file"
    }
}

#endregion

#region Load Public Functions

Write-Verbose "Loading public functions..."

$publicFiles = @(
    'Backup-AppxPackage.ps1',
    'Install-AppxBackup.ps1',
    'New-AppxBackupCertificate.ps1',
    'Test-AppxPackageIntegrity.ps1',
    'Get-AppxBackupInfo.ps1',
    'Export-AppxDependencies.ps1',
    'Get-AppxToolPath.ps1',
    'Test-AppxBackupCompatibility.ps1'
)

foreach ($file in $publicFiles) {
    $filePath = Join-Path (Join-Path $script:ModuleRoot 'Public') $file
    
    if (Test-Path $filePath) {
        try {
            . $filePath
            Write-Verbose "Loaded: $file"
        }
        catch {
            Write-Error "Failed to load public function '$file': $_"
            throw
        }
    }
    else {
        Write-Warning "Public function file not found: $file"
    }
}

#endregion

#region Initialize Module Configuration from JSON

Write-Verbose "Initializing module configuration from ModuleDefaults.json..."

try {
    # Load configuration (Get-AppxConfiguration and Get-AppxDefault are now available)
    $defaults = Get-AppxConfiguration -ConfigName 'ModuleDefaults'
    
    # Update AppxBackupConfig with values from JSON
    $script:AppxBackupConfig['MaxLogSizeMB'] = $defaults.logConfiguration.maxLogSizeMB
    $script:AppxBackupConfig['DefaultCertificateValidityYears'] = $defaults.certificateDefaults.defaultValidityYears
    $script:AppxBackupConfig['DefaultHashAlgorithm'] = $defaults.certificateDefaults.hashAlgorithm
    $script:AppxBackupConfig['DefaultKeyLength'] = $defaults.certificateDefaults.defaultKeyLength
    $script:AppxBackupConfig['RetryAttempts'] = $defaults.sleepDelays.maxCleanupAttempts
    $script:AppxBackupConfig['RetryDelaySeconds'] = $defaults.sleepDelays.verificationDelaySeconds
    $script:AppxBackupConfig['TimeoutSeconds'] = $defaults.timeoutDefaults.processExecutionDefaultSeconds
    
    Write-Verbose "Module configuration loaded from ModuleDefaults.json"
}
catch {
    Write-Warning "Failed to load ModuleDefaults.json, using hardcoded fallbacks: $_"
    # Fallback values already set in initial $script:AppxBackupConfig declaration
}

#endregion

#region Export Module Members

# Export functions (defined in manifest, but explicit export for clarity)
$functionsToExport = @(
    'Backup-AppxPackage',
    'Install-AppxBackup',
    'New-AppxBackupCertificate',
    'Test-AppxPackageIntegrity',
    'Get-AppxBackupInfo',
    'Export-AppxDependencies',
    'Get-AppxToolPath',
    'Test-AppxBackupCompatibility'
)

Export-ModuleMember -Function $functionsToExport

# Export aliases
Set-Alias -Name 'Backup-AppX' -Value 'Backup-AppxPackage' -Scope Global
Set-Alias -Name 'Export-AppX' -Value 'Backup-AppxPackage' -Scope Global
Set-Alias -Name 'Save-AppxPackage' -Value 'Backup-AppxPackage' -Scope Global
Set-Alias -Name 'Restore-AppxPackage' -Value 'Install-AppxBackup' -Scope Global

Export-ModuleMember -Alias @('Backup-AppX', 'Export-AppX', 'Save-AppxPackage', 'Restore-AppxPackage')

# Export configuration for advanced users
Export-ModuleMember -Variable 'AppxBackupConfig'

#endregion

#region Module Cleanup

$MyInvocation.MyCommand.ScriptBlock.Module.OnRemove = {
    Write-Verbose "Cleaning up AppxBackup module..."
    
    # Clear caches
    if ($script:ToolCache) { $script:ToolCache.Clear() }
    if ($script:PackageCache) { $script:PackageCache.Clear() }
    
    # Remove aliases
    Remove-Item Alias:Backup-AppX -ErrorAction SilentlyContinue
    Remove-Item Alias:Export-AppX -ErrorAction SilentlyContinue
    Remove-Item Alias:Save-AppxPackage -ErrorAction SilentlyContinue
    Remove-Item Alias:Restore-AppxPackage -ErrorAction SilentlyContinue
    
    Write-Verbose "Module cleanup complete"
}

#endregion

# Module initialization message
Write-Verbose "AppxBackup v2.0.1 loaded successfully"
Write-Verbose "Private functions: $($privateFiles.Count)"
Write-Verbose "Public functions: $($publicFiles.Count)"
Write-Verbose "Log path: $script:LogPath"