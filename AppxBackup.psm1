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
    Version: 2.0.2
    LastModified: 2026-02-13
    
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

# Module configuration - Emergency fallback values (matches ModuleDefaults.json)
# These values are only used if JSON configuration loading fails during module initialization
# All values below match their counterparts in Config/ModuleDefaults.json for consistency
$script:AppxBackupConfig = @{
    MaxLogSizeMB = 10  # Emergency fallback - matches ModuleDefaults.json:logConfiguration.maxLogSizeMB
    DefaultCertificateValidityYears = 3  # Emergency fallback - matches ModuleDefaults.json:certificateDefaults.defaultValidityYears
    DefaultHashAlgorithm = 'SHA256'  # Emergency fallback - matches ModuleDefaults.json:certificateDefaults.hashAlgorithm
    DefaultKeyLength = 4096  # Emergency fallback - matches ModuleDefaults.json:certificateDefaults.keyLength
    EnableProgressIndicators = $true  # Emergency fallback - matches ModuleDefaults.json:displayDefaults.enableProgressIndicators
    EnableVerboseLogging = $false  # Emergency fallback - matches ModuleDefaults.json:displayDefaults.enableVerboseLogging
    MaxParallelOperations = 4  # Emergency fallback - matches ModuleDefaults.json:performanceDefaults.maxParallelOperations
    PackageValidationLevel = 'Standard'  # Emergency fallback - matches ModuleDefaults.json:validationDefaults.level
    CertificateStorageLocation = 'Cert:\CurrentUser\My'  # Emergency fallback - matches ModuleDefaults.json:certificateDefaults.storageLocation
    TempDirectoryPath = $null  # Emergency fallback - computed from ModuleDefaults.json:pathDefaults.tempDirectoryBase
    RetryAttempts = 3  # Emergency fallback - matches ModuleDefaults.json:sleepDelays.maxCleanupAttempts
    RetryDelaySeconds = 2  # Emergency fallback - matches ModuleDefaults.json:sleepDelays.verificationDelaySeconds
    TimeoutSeconds = 3600  # Emergency fallback - matches ModuleDefaults.json:timeoutDefaults.processExecutionDefaultSeconds
}

#endregion

#region Load Private Functions

Write-Verbose "Loading private functions..."

$privateFiles = @(
    # Core infrastructure (no dependencies on other private functions)
    'Get-AppxConfiguration.ps1',
    'Get-AppxDefault.ps1',
    'Write-AppxLog.ps1',
    'ConvertTo-SecureFilePath.ps1',
    'ConvertTo-HtmlEncodedString.ps1',
    
    # Atomic helpers (depend on core infrastructure only)
    'Find-AppxSdkTool.ps1',
    'Resolve-AppxManifestNode.ps1',
    'Test-AppxDiskSpace.ps1',
    'Test-AppxArchitectureCompatibility.ps1',
    'Test-AppxPackagingPrerequisites.ps1',
    'Get-AppxMakeAppxErrorAnalysis.ps1',
    'Remove-AppxItemWithRetry.ps1',
    'Install-AppxCertificateToStore.ps1',
    'Copy-AppxSourceDirectory.ps1',
    
    # Mid-level helpers (may depend on atomic helpers)
    'Invoke-ProcessSafely.ps1',
    'Invoke-AppxSignTool.ps1',
    'Test-AppxToolAvailability.ps1',
    'Get-AppxManifestData.ps1',
    'Resolve-AppxDependencies.ps1',
    
    # High-level helpers (depend on mid-level helpers)
    'New-AppxPackageInternal.ps1',
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

# Note: AppxBackupConfig is intentionally NOT exported to prevent runtime modifications
# Advanced users can access it via $script:AppxBackupConfig if needed, but direct modification is discouraged

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
Write-Verbose "AppxBackup v2.0.2 loaded successfully"
Write-Verbose "Private functions: $($privateFiles.Count)"
Write-Verbose "Public functions: $($publicFiles.Count)"
Write-Verbose "Log path: $script:LogPath"