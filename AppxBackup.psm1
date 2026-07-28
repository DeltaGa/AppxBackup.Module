#Requires -Version 7.4

<#
.SYNOPSIS
    AppxBackup Module - Enterprise-grade Windows Application Package backup toolkit

.DESCRIPTION
    Modern replacement for 2016-era APPX backup solutions.
    Provides comprehensive, secure, and performant package backup/restore capabilities.

.NOTES
    Name: AppxBackup
    Author: DeltaGa
    Version: 2.1.0
    LastModified: 2026-07-27
    
    Requires:
    - PowerShell 7.4+
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

#region Load Function Files

# Every file under Private/ and Public/ contains only function definitions, so the
# dot-source order carries no meaning: PowerShell resolves a call target when the
# function runs, not when it is defined. Files are therefore discovered from disk
# rather than listed by hand, which removes a class of bug where a new file was
# added but never registered here and silently failed to load.

$privateFiles = @(Get-ChildItem -LiteralPath (Join-Path $script:ModuleRoot 'Private') -Filter '*.ps1' -File -ErrorAction SilentlyContinue | Sort-Object Name)
$publicFiles = @(Get-ChildItem -LiteralPath (Join-Path $script:ModuleRoot 'Public') -Filter '*.ps1' -File -ErrorAction SilentlyContinue | Sort-Object Name)

if ($publicFiles.Count -eq 0) {
    throw "AppxBackup: no public function files found under '$(Join-Path $script:ModuleRoot 'Public')'. The module installation is incomplete."
}

foreach ($file in ($privateFiles + $publicFiles)) {
    try {
        . $file.FullName
        Write-Verbose "Loaded: $($file.Name)"
    }
    catch {
        # A function file that fails to parse leaves the module half-built, so fail
        # the import loudly rather than exporting a partially working surface.
        throw "Failed to load function file '$($file.Name)': $_"
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

# One public function per file under Public/, so the file names are the export list.
# AppxBackup.psd1 still names each function explicitly and remains the authoritative
# public contract - the manifest intersects with whatever is exported here, and the
# test suite asserts the two agree.
$functionsToExport = @($publicFiles | ForEach-Object { $_.BaseName })

Export-ModuleMember -Function $functionsToExport

# Aliases are created in module scope and published via Export-ModuleMember. They were
# previously created with -Scope Global, which wrote into the caller's session directly
# and left them behind for any consumer that did not trigger OnRemove. Exported aliases
# are tied to the module and are withdrawn automatically by Remove-Module.
Set-Alias -Name 'Backup-AppX' -Value 'Backup-AppxPackage'
Set-Alias -Name 'Export-AppX' -Value 'Backup-AppxPackage'
Set-Alias -Name 'Save-AppxPackage' -Value 'Backup-AppxPackage'
Set-Alias -Name 'Restore-AppxPackage' -Value 'Install-AppxBackup'

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
    if ($script:ConfigCache) { $script:ConfigCache.Clear() }

    # Aliases are exported module members and are withdrawn by Remove-Module itself,
    # so there is nothing to unregister here.

    Write-Verbose "Module cleanup complete"
}

#endregion

# Module initialization message
Write-Verbose "AppxBackup loaded successfully"
Write-Verbose "Private function files: $($privateFiles.Count)"
Write-Verbose "Public function files: $($publicFiles.Count)"
Write-Verbose "Log path: $script:LogPath"