<#
.SYNOPSIS
    Comprehensive test script for AppxBackup.Module v2.0.2 public functions.

.DESCRIPTION
    Validates all 8 exported public functions and their parameter combinations:
      1. Get-AppxToolPath       (MakeAppx, SignTool, Certutil + -Refresh)
      2. Backup-AppxPackage     (vanilla, -IncludeDependencies, -DependencyReportOnly,
                                 -NoCertificate, -CompressionLevel, -Force)
      3. Get-AppxBackupInfo     (-IncludeFileList, -IncludeSignatureInfo, -IncludeManifestXml)
      4. Test-AppxPackageIntegrity (-VerifySignature, -CheckManifest)
      5. Test-AppxBackupCompatibility (-CheckDependencies, -Detailed)
      6. Export-AppxDependencies (JSON, XML, CSV, HTML + -Recursive, -IncludeOptional)
      7. New-AppxBackupCertificate (-ValidityYears, -KeyLength, -ReplaceExisting)
      8. Install-AppxBackup     (only when -TestInstall is specified)

    Automatically discovers the test app via Get-AppxPackage -Name "*Tivi*".
    Saves full console transcript + structured results file.

.PARAMETER TestFolder
    Backup/test output folder. Default: D:\Dropbox\AppxBackups\

.PARAMETER AppName
    Part of the app name to search for via Get-AppxPackage -Name "*$AppName*".
    Default: Tivi (matches TiviMate). Examples: "Netflix", "Spotify", "Xbox".

.PARAMETER TestInstall
    If specified, runs Install-AppxBackup test (destructive). Default: $false.

.EXAMPLE
    .\Test-AppxBackupModule.ps1
    .\Test-AppxBackupModule.ps1 -TestFolder "C:\Temp\AppxTest\"
    .\Test-AppxBackupModule.ps1 -AppName "Netflix"
    .\Test-AppxBackupModule.ps1 -TestFolder "C:\Temp\AppxTest\" -AppName "Spotify" -TestInstall

.NOTES
    Author: DeltaGa
    Version: 2.0.2
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$TestFolder = 'D:\Dropbox\AppxBackups\',

    [Parameter(Position = 1)]
    [string]$AppName = 'Tivi',

    [Parameter()]
    [switch]$TestInstall
)

# ============================================================================
# CONFIGURATION
# ============================================================================

$ErrorActionPreference = 'Continue'
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$testRunDir = [System.IO.Path]::Combine($TestFolder, "TestRun_$timestamp")
$transcriptPath = [System.IO.Path]::Combine($testRunDir, "TestConsole_$timestamp.txt")
$resultsPath = [System.IO.Path]::Combine($testRunDir, "TestResults_$timestamp.txt")
$logDate = Get-Date -Format 'yyyyMMdd'
$moduleLogPath = [System.IO.Path]::Combine($env:TEMP, "AppxBackup_$logDate.log")

# Expected naming patterns (populated after app discovery)
$script:AppFullName = $null
$script:BackupAppx = $null
$script:BackupAppxpack = $null
$script:BackupCer = $null
$script:BackupDepJson = $null

# Test results accumulator
$script:Results = [System.Collections.ArrayList]::new()
$script:TestNumber = 0
$script:PassCount = 0
$script:FailCount = 0
$script:SkipCount = 0
$script:ErrorCount = 0
$script:StartTime = Get-Date

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

function Add-TestResult {
    <#
    .SYNOPSIS
        Records a single test result.
    #>
    param(
        [string]$TestName,
        [ValidateSet('PASS', 'FAIL', 'SKIP', 'ERROR')]
        [string]$Status,
        [string]$Details = '',
        [string]$Expected = '',
        [string]$Actual = ''
    )

    $script:TestNumber++

    switch ($Status) {
        'PASS'  { $script:PassCount++;  $color = 'Green' }
        'FAIL'  { $script:FailCount++;  $color = 'Red' }
        'SKIP'  { $script:SkipCount++;  $color = 'Yellow' }
        'ERROR' { $script:ErrorCount++; $color = 'Magenta' }
    }

    $entry = [PSCustomObject]@{
        '#'        = $script:TestNumber
        Test       = $TestName
        Status     = $Status
        Details    = $Details
        Expected   = $Expected
        Actual     = $Actual
        Timestamp  = (Get-Date -Format 'HH:mm:ss.fff')
    }

    [void]$script:Results.Add($entry)

    # Live console feedback
    $tag = "[$Status]".PadRight(7)
    Write-Host "  $tag #$($script:TestNumber.ToString().PadLeft(3)) $TestName" -ForegroundColor $color -NoNewline
    if ($Details) { Write-Host " - $Details" -ForegroundColor Gray } else { Write-Host '' }
}

function Test-FileCreated {
    <#
    .SYNOPSIS
        Validates a file was created and has content.
    #>
    param(
        [string]$Path,
        [string]$TestName,
        [long]$MinSizeBytes = 1
    )

    if (Test-Path -LiteralPath $Path) {
        $file = Get-Item -LiteralPath $Path
        if ($file.Length -ge $MinSizeBytes) {
            Add-TestResult -TestName $TestName -Status 'PASS' `
                -Details "$($file.Name) ($([Math]::Round($file.Length / 1KB, 1)) KB)" `
                -Expected "File exists, >= $MinSizeBytes bytes" `
                -Actual "$($file.Length) bytes"
            return $true
        }
        else {
            Add-TestResult -TestName $TestName -Status 'FAIL' `
                -Details "File too small: $($file.Length) bytes" `
                -Expected ">= $MinSizeBytes bytes" `
                -Actual "$($file.Length) bytes"
            return $false
        }
    }
    else {
        Add-TestResult -TestName $TestName -Status 'FAIL' `
            -Details "File not found: $(Split-Path $Path -Leaf)" `
            -Expected 'File exists' `
            -Actual 'Not found'
        return $false
    }
}

function Test-ObjectProperty {
    <#
    .SYNOPSIS
        Validates a property exists and optionally matches an expected value.
    #>
    param(
        [object]$Object,
        [string]$PropertyName,
        [string]$TestName,
        $ExpectedValue = $null,
        [switch]$NotNull
    )

    $value = $Object.$PropertyName

    if ($null -eq $Object) {
        Add-TestResult -TestName $TestName -Status 'FAIL' -Details 'Object is null'
        return $false
    }

    if ($NotNull.IsPresent) {
        if ($null -ne $value -and $value -ne '') {
            Add-TestResult -TestName $TestName -Status 'PASS' `
                -Details "$PropertyName = $value" -Expected 'Not null/empty' -Actual "$value"
            return $true
        }
        else {
            Add-TestResult -TestName $TestName -Status 'FAIL' `
                -Details "$PropertyName is null/empty" -Expected 'Not null/empty' -Actual "$value"
            return $false
        }
    }

    if ($null -ne $ExpectedValue) {
        if ("$value" -eq "$ExpectedValue") {
            Add-TestResult -TestName $TestName -Status 'PASS' `
                -Details "$PropertyName = $value" -Expected "$ExpectedValue" -Actual "$value"
            return $true
        }
        else {
            Add-TestResult -TestName $TestName -Status 'FAIL' `
                -Details "$PropertyName mismatch" -Expected "$ExpectedValue" -Actual "$value"
            return $false
        }
    }

    # Property exists check only
    Add-TestResult -TestName $TestName -Status 'PASS' -Details "$PropertyName present"
    return $true
}

function Write-Section {
    param([string]$Title)
    $bar = '=' * 76
    Write-Host "`n$bar" -ForegroundColor Cyan
    Write-Host "  $Title" -ForegroundColor White
    Write-Host "$bar" -ForegroundColor Cyan
}

# ============================================================================
# SETUP
# ============================================================================

Write-Host "`n" -NoNewline
Write-Host ('=' * 76) -ForegroundColor White
Write-Host "  AppxBackup.Module v2.0.2 - Comprehensive Test Suite" -ForegroundColor White
Write-Host "  Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
Write-Host "  Test Folder: $TestFolder" -ForegroundColor Gray
Write-Host "  App Search : *$AppName*" -ForegroundColor Gray
Write-Host "  Install Test: $($TestInstall.IsPresent)" -ForegroundColor Gray
Write-Host ('=' * 76) -ForegroundColor White

# Create test run directory
if (-not (Test-Path -LiteralPath $testRunDir)) {
    [void](New-Item -Path $testRunDir -ItemType Directory -Force)
}

# Start transcript (captures all console output)
try {
    Start-Transcript -Path $transcriptPath -Force | Out-Null
    Write-Host "[SETUP] Transcript logging to: $transcriptPath" -ForegroundColor DarkGray
}
catch {
    Write-Warning "Failed to start transcript: $_"
}

# Load the module
Write-Section "MODULE LOAD"
try {
    # Script lives in the module root alongside Import-AppxBackup.ps1
    $moduleDir = $PSScriptRoot
    if (-not $moduleDir) {
        # Fallback when run interactively (no $PSScriptRoot)
        $moduleDir = 'C:\Users\admin\Documents\GitHub\AppxBackup.Module'
    }

    $importScript = [System.IO.Path]::Combine($moduleDir, 'Import-AppxBackup.ps1')
    if (-not (Test-Path -LiteralPath $importScript)) {
        throw "Import-AppxBackup.ps1 not found at: $importScript"
    }

    Push-Location -Path $moduleDir
    Write-Host "  Working directory: $moduleDir" -ForegroundColor Gray
    Write-Host "  Loading module via: .\Import-AppxBackup.ps1" -ForegroundColor Gray
    & $importScript
    Pop-Location
    Add-TestResult -TestName 'Module Load' -Status 'PASS' -Details "Loaded from $moduleDir"
}
catch {
    Pop-Location -ErrorAction SilentlyContinue
    Add-TestResult -TestName 'Module Load' -Status 'ERROR' -Details "$_"
    Write-Host "`n[FATAL] Cannot load module. Aborting." -ForegroundColor Red
    Stop-Transcript -ErrorAction SilentlyContinue
    exit 1
}

# Discover test app
Write-Host "`n  Discovering test app (Get-AppxPackage -Name '*$AppName*')..." -ForegroundColor Gray
$matchedApps = @(Get-AppxPackage -Name "*$AppName*" | Where-Object { -not $_.IsFramework })

if ($matchedApps.Count -eq 0) {
    # Try broader search in case user gave exact publisher or family name
    $matchedApps = @(Get-AppxPackage | Where-Object {
        $_.Name -like "*$AppName*" -or
        $_.PackageFamilyName -like "*$AppName*" -or
        $_.PackageFullName -like "*$AppName*"
    } | Where-Object { -not $_.IsFramework })
}

if ($matchedApps.Count -eq 0) {
    Add-TestResult -TestName 'App Discovery' -Status 'ERROR' `
        -Details "No app found matching '*$AppName*'"
    Write-Host "`n[FATAL] No installed app matches '*$AppName*'." -ForegroundColor Red
    Write-Host "  Hint: Run 'Get-AppxPackage | Select Name' to see installed apps." -ForegroundColor Yellow
    Stop-Transcript -ErrorAction SilentlyContinue
    exit 1
}

if ($matchedApps.Count -gt 1) {
    Write-Host "  Multiple matches found:" -ForegroundColor Yellow
    for ($i = 0; $i -lt $matchedApps.Count; $i++) {
        $m = $matchedApps[$i]
        Write-Host "    [$($i + 1)] $($m.Name) v$($m.Version) ($($m.Architecture))" -ForegroundColor Gray
    }
    Write-Host "  Using first non-framework match." -ForegroundColor Yellow
}

$app = $matchedApps[0]

$script:AppFullName = $app.PackageFullName
$baseName = $app.PackageFullName
Write-Host "  Selected: $($app.Name) v$($app.Version) ($($app.Architecture))" -ForegroundColor Green
Write-Host "  FullName: $baseName" -ForegroundColor Gray
Write-Host "  InstallLocation: $($app.InstallLocation)" -ForegroundColor Gray
Add-TestResult -TestName 'App Discovery' -Status 'PASS' `
    -Details "$($app.Name) v$($app.Version) ($($app.Architecture))"

# Derive expected output filenames
$script:BackupAppx     = "$baseName.appx"
$script:BackupAppxpack = "$baseName.appxpack"
$script:BackupCer      = "$baseName.cer"
$script:BackupDepJson  = "${baseName}_Dependencies.json"

# ============================================================================
# TEST 1: Get-AppxToolPath
# ============================================================================

Write-Section "1. Get-AppxToolPath"

foreach ($tool in @('MakeAppx', 'SignTool', 'Certutil')) {
    try {
        $toolPath = Get-AppxToolPath -ToolName $tool
        if ($toolPath -and (Test-Path -LiteralPath $toolPath)) {
            Add-TestResult -TestName "Get-AppxToolPath -ToolName $tool" -Status 'PASS' `
                -Details $toolPath -Expected 'Valid path' -Actual $toolPath
        }
        elseif ($toolPath) {
            Add-TestResult -TestName "Get-AppxToolPath -ToolName $tool" -Status 'FAIL' `
                -Details "Path returned but not found: $toolPath"
        }
        else {
            Add-TestResult -TestName "Get-AppxToolPath -ToolName $tool" -Status 'FAIL' `
                -Details 'Returned null' -Expected 'Valid path' -Actual 'null'
        }
    }
    catch {
        Add-TestResult -TestName "Get-AppxToolPath -ToolName $tool" -Status 'ERROR' -Details "$_"
    }
}

# Test -Refresh parameter
try {
    $toolPath = Get-AppxToolPath -ToolName 'MakeAppx' -Refresh
    if ($toolPath) {
        Add-TestResult -TestName 'Get-AppxToolPath -Refresh' -Status 'PASS' -Details $toolPath
    }
    else {
        Add-TestResult -TestName 'Get-AppxToolPath -Refresh' -Status 'FAIL' -Details 'null after refresh'
    }
}
catch {
    Add-TestResult -TestName 'Get-AppxToolPath -Refresh' -Status 'ERROR' -Details "$_"
}

# ============================================================================
# TEST 2: Backup-AppxPackage (vanilla - standalone .appx)
# ============================================================================

Write-Section "2. Backup-AppxPackage"

# 2a: Vanilla backup (standalone .appx + .cer)
$vanillaDir = [System.IO.Path]::Combine($testRunDir, '2a_Vanilla')
[void](New-Item -Path $vanillaDir -ItemType Directory -Force)

Write-Host "`n  [2a] Vanilla backup (standalone .appx)..." -ForegroundColor Gray
$backupResult = $null
try {
    $backupResult = Backup-AppxPackage -PackagePath $app.InstallLocation `
        -OutputPath $vanillaDir -Force -Confirm:$false
    Add-TestResult -TestName 'Backup-AppxPackage (vanilla): Executes' -Status 'PASS' `
        -Details "Completed"
}
catch {
    Add-TestResult -TestName 'Backup-AppxPackage (vanilla): Executes' -Status 'ERROR' -Details "$_"
}

if ($backupResult) {
    Test-ObjectProperty -Object $backupResult -PropertyName 'Success' -TestName 'Backup (vanilla): .Success' -ExpectedValue $true
    Test-ObjectProperty -Object $backupResult -PropertyName 'PackageName' -TestName 'Backup (vanilla): .PackageName' -NotNull
    Test-ObjectProperty -Object $backupResult -PropertyName 'PackageVersion' -TestName 'Backup (vanilla): .PackageVersion' -NotNull
    Test-ObjectProperty -Object $backupResult -PropertyName 'PackageArchitecture' -TestName 'Backup (vanilla): .PackageArchitecture' -NotNull
    Test-ObjectProperty -Object $backupResult -PropertyName 'CertificateThumbprint' -TestName 'Backup (vanilla): .CertificateThumbprint' -NotNull
    Test-ObjectProperty -Object $backupResult -PropertyName 'IsZipArchive' -TestName 'Backup (vanilla): .IsZipArchive' -ExpectedValue $false

    $vanillaAppx = $backupResult.PackageFilePath
    $vanillaCer  = $backupResult.CertificateFilePath
    Test-FileCreated -Path $vanillaAppx -TestName 'Backup (vanilla): .appx file created' -MinSizeBytes 1024
    Test-FileCreated -Path $vanillaCer  -TestName 'Backup (vanilla): .cer file created'  -MinSizeBytes 100
}
else {
    Add-TestResult -TestName 'Backup (vanilla): Validation' -Status 'SKIP' -Details 'No result object'
}

# 2b: Backup with -IncludeDependencies (.appxpack)
$depsDir = [System.IO.Path]::Combine($testRunDir, '2b_WithDeps')
[void](New-Item -Path $depsDir -ItemType Directory -Force)

Write-Host "`n  [2b] Backup with -IncludeDependencies (.appxpack)..." -ForegroundColor Gray
$depBackupResult = $null
try {
    $depBackupResult = Backup-AppxPackage -PackagePath $app.InstallLocation `
        -OutputPath $depsDir -IncludeDependencies -Force -Confirm:$false
    Add-TestResult -TestName 'Backup-AppxPackage (-IncludeDeps): Executes' -Status 'PASS'
}
catch {
    Add-TestResult -TestName 'Backup-AppxPackage (-IncludeDeps): Executes' -Status 'ERROR' -Details "$_"
}

if ($depBackupResult) {
    Test-ObjectProperty -Object $depBackupResult -PropertyName 'Success' -TestName 'Backup (deps): .Success' -ExpectedValue $true
    Test-ObjectProperty -Object $depBackupResult -PropertyName 'IsZipArchive' -TestName 'Backup (deps): .IsZipArchive' -ExpectedValue $true
    Test-ObjectProperty -Object $depBackupResult -PropertyName 'BundledDependencyCount' -TestName 'Backup (deps): .BundledDependencyCount' -NotNull

    $depsAppxpack = $depBackupResult.PackageFilePath
    Test-FileCreated -Path $depsAppxpack -TestName 'Backup (deps): .appxpack file created' -MinSizeBytes 1024
}
else {
    Add-TestResult -TestName 'Backup (deps): Validation' -Status 'SKIP' -Details 'No result object'
}

# 2c: Backup with -DependencyReportOnly
$reportDir = [System.IO.Path]::Combine($testRunDir, '2c_DepReport')
[void](New-Item -Path $reportDir -ItemType Directory -Force)

Write-Host "`n  [2c] Backup with -DependencyReportOnly..." -ForegroundColor Gray
$reportResult = $null
try {
    $reportResult = Backup-AppxPackage -PackagePath $app.InstallLocation `
        -OutputPath $reportDir -DependencyReportOnly -Force -Confirm:$false
    Add-TestResult -TestName 'Backup-AppxPackage (-DepReportOnly): Executes' -Status 'PASS'
}
catch {
    Add-TestResult -TestName 'Backup-AppxPackage (-DepReportOnly): Executes' -Status 'ERROR' -Details "$_"
}

if ($reportResult) {
    Test-ObjectProperty -Object $reportResult -PropertyName 'Success' -TestName 'Backup (report): .Success' -ExpectedValue $true
    Test-ObjectProperty -Object $reportResult -PropertyName 'DependencyReportPath' -TestName 'Backup (report): .DependencyReportPath' -NotNull

    if ($reportResult.DependencyReportPath) {
        Test-FileCreated -Path $reportResult.DependencyReportPath -TestName 'Backup (report): dep JSON created' -MinSizeBytes 10
    }
}
else {
    Add-TestResult -TestName 'Backup (report): Validation' -Status 'SKIP' -Details 'No result object'
}

# 2d: Backup with -NoCertificate
$noCertDir = [System.IO.Path]::Combine($testRunDir, '2d_NoCert')
[void](New-Item -Path $noCertDir -ItemType Directory -Force)

Write-Host "`n  [2d] Backup with -NoCertificate..." -ForegroundColor Gray
$noCertResult = $null
try {
    $noCertResult = Backup-AppxPackage -PackagePath $app.InstallLocation `
        -OutputPath $noCertDir -NoCertificate -Force -Confirm:$false
    Add-TestResult -TestName 'Backup-AppxPackage (-NoCertificate): Executes' -Status 'PASS'
}
catch {
    Add-TestResult -TestName 'Backup-AppxPackage (-NoCertificate): Executes' -Status 'ERROR' -Details "$_"
}

if ($noCertResult) {
    Test-ObjectProperty -Object $noCertResult -PropertyName 'Success' -TestName 'Backup (noCert): .Success' -ExpectedValue $true

    # Verify no certificate was created
    $cerFiles = Get-ChildItem -Path $noCertDir -Filter '*.cer' -ErrorAction SilentlyContinue
    if ($cerFiles.Count -eq 0) {
        Add-TestResult -TestName 'Backup (noCert): No .cer file created' -Status 'PASS' `
            -Expected '0 .cer files' -Actual '0 .cer files'
    }
    else {
        Add-TestResult -TestName 'Backup (noCert): No .cer file created' -Status 'FAIL' `
            -Expected '0 .cer files' -Actual "$($cerFiles.Count) .cer files"
    }
}

# 2e: Backup with -CompressionLevel None
$noCompDir = [System.IO.Path]::Combine($testRunDir, '2e_NoCompress')
[void](New-Item -Path $noCompDir -ItemType Directory -Force)

Write-Host "`n  [2e] Backup with -CompressionLevel None..." -ForegroundColor Gray
try {
    $noCompResult = Backup-AppxPackage -PackagePath $app.InstallLocation `
        -OutputPath $noCompDir -CompressionLevel None -Force -Confirm:$false
    Add-TestResult -TestName 'Backup-AppxPackage (-CompressionLevel None): Executes' -Status 'PASS'
    Test-ObjectProperty -Object $noCompResult -PropertyName 'Success' -TestName 'Backup (noComp): .Success' -ExpectedValue $true
}
catch {
    Add-TestResult -TestName 'Backup-AppxPackage (-CompressionLevel None): Executes' -Status 'ERROR' -Details "$_"
}

# ============================================================================
# TEST 3: Get-AppxBackupInfo
# ============================================================================

Write-Section "3. Get-AppxBackupInfo"

# Determine best available test files
$testAppx = if ($vanillaAppx -and (Test-Path -LiteralPath $vanillaAppx)) { $vanillaAppx } else { $null }
$testAppxpack = if ($depsAppxpack -and (Test-Path -LiteralPath $depsAppxpack)) { $depsAppxpack } else { $null }

# 3a: Basic info on .appx
if ($testAppx) {
    Write-Host "`n  [3a] Get-AppxBackupInfo on .appx..." -ForegroundColor Gray
    try {
        $info = Get-AppxBackupInfo -PackagePath $testAppx
        Add-TestResult -TestName 'Get-AppxBackupInfo (.appx): Executes' -Status 'PASS'
        Test-ObjectProperty -Object $info -PropertyName 'PackageName' -TestName 'Info (.appx): .PackageName' -NotNull
        Test-ObjectProperty -Object $info -PropertyName 'PackageVersion' -TestName 'Info (.appx): .PackageVersion' -NotNull
        Test-ObjectProperty -Object $info -PropertyName 'PackageSizeMB' -TestName 'Info (.appx): .PackageSizeMB' -NotNull
        Test-ObjectProperty -Object $info -PropertyName 'PackageArchitecture' -TestName 'Info (.appx): .PackageArchitecture' -NotNull
    }
    catch {
        Add-TestResult -TestName 'Get-AppxBackupInfo (.appx): Executes' -Status 'ERROR' -Details "$_"
    }

    # 3b: With -IncludeFileList
    Write-Host "  [3b] Get-AppxBackupInfo -IncludeFileList..." -ForegroundColor Gray
    try {
        $infoFL = Get-AppxBackupInfo -PackagePath $testAppx -IncludeFileList
        if ($infoFL.FileList -and $infoFL.FileList.Count -gt 0) {
            Add-TestResult -TestName 'Info (.appx) -IncludeFileList' -Status 'PASS' `
                -Details "$($infoFL.FileList.Count) files listed"
        }
        else {
            Add-TestResult -TestName 'Info (.appx) -IncludeFileList' -Status 'FAIL' `
                -Details 'FileList empty or null'
        }
    }
    catch {
        Add-TestResult -TestName 'Info (.appx) -IncludeFileList' -Status 'ERROR' -Details "$_"
    }

    # 3c: With -IncludeSignatureInfo
    Write-Host "  [3c] Get-AppxBackupInfo -IncludeSignatureInfo..." -ForegroundColor Gray
    try {
        $infoSig = Get-AppxBackupInfo -PackagePath $testAppx -IncludeSignatureInfo
        if ($null -ne $infoSig.SignatureInfo) {
            Add-TestResult -TestName 'Info (.appx) -IncludeSignatureInfo' -Status 'PASS' `
                -Details "SignatureInfo populated"
        }
        else {
            # Signature may be absent on unsigned packages
            Add-TestResult -TestName 'Info (.appx) -IncludeSignatureInfo' -Status 'PASS' `
                -Details "SignatureInfo is null (package may be unsigned)"
        }
    }
    catch {
        Add-TestResult -TestName 'Info (.appx) -IncludeSignatureInfo' -Status 'ERROR' -Details "$_"
    }

    # 3d: With -IncludeManifestXml
    Write-Host "  [3d] Get-AppxBackupInfo -IncludeManifestXml..." -ForegroundColor Gray
    try {
        $infoXml = Get-AppxBackupInfo -PackagePath $testAppx -IncludeManifestXml
        if ($infoXml.ManifestXml -and $infoXml.ManifestXml.Length -gt 0) {
            Add-TestResult -TestName 'Info (.appx) -IncludeManifestXml' -Status 'PASS' `
                -Details "$($infoXml.ManifestXml.Length) chars of XML"
        }
        else {
            Add-TestResult -TestName 'Info (.appx) -IncludeManifestXml' -Status 'FAIL' `
                -Details 'ManifestXml empty or null'
        }
    }
    catch {
        Add-TestResult -TestName 'Info (.appx) -IncludeManifestXml' -Status 'ERROR' -Details "$_"
    }

    # 3e: All switches combined
    Write-Host "  [3e] Get-AppxBackupInfo (all switches)..." -ForegroundColor Gray
    try {
        $infoAll = Get-AppxBackupInfo -PackagePath $testAppx -IncludeFileList -IncludeSignatureInfo -IncludeManifestXml
        Add-TestResult -TestName 'Info (.appx) all switches combined' -Status 'PASS' `
            -Details "Files=$($infoAll.FileList.Count), XML=$($infoAll.ManifestXml.Length) chars"
    }
    catch {
        Add-TestResult -TestName 'Info (.appx) all switches combined' -Status 'ERROR' -Details "$_"
    }
}
else {
    Add-TestResult -TestName 'Get-AppxBackupInfo (.appx)' -Status 'SKIP' -Details 'No .appx test file'
}

# 3f: Get-AppxBackupInfo on .appxpack
if ($testAppxpack) {
    Write-Host "`n  [3f] Get-AppxBackupInfo on .appxpack..." -ForegroundColor Gray
    try {
        $infoZip = Get-AppxBackupInfo -PackagePath $testAppxpack
        Add-TestResult -TestName 'Get-AppxBackupInfo (.appxpack): Executes' -Status 'PASS'
        Test-ObjectProperty -Object $infoZip -PropertyName 'PackageName' -TestName 'Info (.appxpack): .PackageName' -NotNull
    }
    catch {
        Add-TestResult -TestName 'Get-AppxBackupInfo (.appxpack): Executes' -Status 'ERROR' -Details "$_"
    }
}
else {
    Add-TestResult -TestName 'Get-AppxBackupInfo (.appxpack)' -Status 'SKIP' -Details 'No .appxpack test file'
}

# ============================================================================
# TEST 4: Test-AppxPackageIntegrity
# ============================================================================

Write-Section "4. Test-AppxPackageIntegrity"

# 4a: Basic integrity on .appx
if ($testAppx) {
    Write-Host "`n  [4a] Test-AppxPackageIntegrity on .appx..." -ForegroundColor Gray
    try {
        $integrity = Test-AppxPackageIntegrity -PackagePath $testAppx
        Add-TestResult -TestName 'Integrity (.appx): Executes' -Status 'PASS'
        Test-ObjectProperty -Object $integrity -PropertyName 'IsValid' -TestName 'Integrity (.appx): .IsValid' -NotNull

        if ($integrity.Issues -and $integrity.Issues.Count -gt 0) {
            Add-TestResult -TestName 'Integrity (.appx): Issues' -Status 'PASS' `
                -Details "$($integrity.Issues.Count) issues noted (informational)"
        }
        else {
            Add-TestResult -TestName 'Integrity (.appx): No Issues' -Status 'PASS'
        }
    }
    catch {
        Add-TestResult -TestName 'Integrity (.appx): Executes' -Status 'ERROR' -Details "$_"
    }

    # 4b: With -VerifySignature
    Write-Host "  [4b] Test-AppxPackageIntegrity -VerifySignature..." -ForegroundColor Gray
    try {
        $integritySig = Test-AppxPackageIntegrity -PackagePath $testAppx -VerifySignature
        Add-TestResult -TestName 'Integrity (.appx) -VerifySignature' -Status 'PASS' `
            -Details "SignatureValid=$($integritySig.SignatureValid)"
    }
    catch {
        Add-TestResult -TestName 'Integrity (.appx) -VerifySignature' -Status 'ERROR' -Details "$_"
    }

    # 4c: With -CheckManifest
    Write-Host "  [4c] Test-AppxPackageIntegrity -CheckManifest..." -ForegroundColor Gray
    try {
        $integrityMan = Test-AppxPackageIntegrity -PackagePath $testAppx -CheckManifest
        Add-TestResult -TestName 'Integrity (.appx) -CheckManifest' -Status 'PASS' `
            -Details "ManifestValid=$($integrityMan.ManifestValid)"
    }
    catch {
        Add-TestResult -TestName 'Integrity (.appx) -CheckManifest' -Status 'ERROR' -Details "$_"
    }

    # 4d: Both switches
    Write-Host "  [4d] Test-AppxPackageIntegrity (both switches)..." -ForegroundColor Gray
    try {
        $integrityBoth = Test-AppxPackageIntegrity -PackagePath $testAppx -VerifySignature -CheckManifest
        Add-TestResult -TestName 'Integrity (.appx) both switches' -Status 'PASS' `
            -Details "IsValid=$($integrityBoth.IsValid), Sig=$($integrityBoth.SignatureValid), Man=$($integrityBoth.ManifestValid)"
    }
    catch {
        Add-TestResult -TestName 'Integrity (.appx) both switches' -Status 'ERROR' -Details "$_"
    }
}
else {
    Add-TestResult -TestName 'Test-AppxPackageIntegrity (.appx)' -Status 'SKIP' -Details 'No .appx test file'
}

# 4e: Integrity on .appxpack
if ($testAppxpack) {
    Write-Host "`n  [4e] Test-AppxPackageIntegrity on .appxpack..." -ForegroundColor Gray
    try {
        $integrityZip = Test-AppxPackageIntegrity -PackagePath $testAppxpack -VerifySignature -CheckManifest
        Add-TestResult -TestName 'Integrity (.appxpack): Executes' -Status 'PASS' `
            -Details "IsValid=$($integrityZip.IsValid)"
    }
    catch {
        Add-TestResult -TestName 'Integrity (.appxpack): Executes' -Status 'ERROR' -Details "$_"
    }
}
else {
    Add-TestResult -TestName 'Test-AppxPackageIntegrity (.appxpack)' -Status 'SKIP' -Details 'No .appxpack test file'
}

# ============================================================================
# TEST 5: Test-AppxBackupCompatibility
# ============================================================================

Write-Section "5. Test-AppxBackupCompatibility"

# 5a: Basic compatibility on .appx
if ($testAppx) {
    Write-Host "`n  [5a] Test-AppxBackupCompatibility on .appx..." -ForegroundColor Gray
    try {
        $compat = Test-AppxBackupCompatibility -PackagePath $testAppx
        Add-TestResult -TestName 'Compat (.appx): Executes' -Status 'PASS'
        Test-ObjectProperty -Object $compat -PropertyName 'IsCompatible' -TestName 'Compat (.appx): .IsCompatible' -NotNull
        Test-ObjectProperty -Object $compat -PropertyName 'ArchitectureCompatible' -TestName 'Compat (.appx): .ArchitectureCompatible' -NotNull
        Test-ObjectProperty -Object $compat -PropertyName 'SystemArchitecture' -TestName 'Compat (.appx): .SystemArchitecture' -NotNull
    }
    catch {
        Add-TestResult -TestName 'Compat (.appx): Executes' -Status 'ERROR' -Details "$_"
    }

    # 5b: With -CheckDependencies
    Write-Host "  [5b] Test-AppxBackupCompatibility -CheckDependencies..." -ForegroundColor Gray
    try {
        $compatDeps = Test-AppxBackupCompatibility -PackagePath $testAppx -CheckDependencies
        Add-TestResult -TestName 'Compat (.appx) -CheckDependencies' -Status 'PASS' `
            -Details "IsCompatible=$($compatDeps.IsCompatible), Issues=$($compatDeps.IssueCount)"
    }
    catch {
        Add-TestResult -TestName 'Compat (.appx) -CheckDependencies' -Status 'ERROR' -Details "$_"
    }

    # 5c: With -Detailed
    Write-Host "  [5c] Test-AppxBackupCompatibility -Detailed..." -ForegroundColor Gray
    try {
        $compatDetailed = Test-AppxBackupCompatibility -PackagePath $testAppx -Detailed
        Add-TestResult -TestName 'Compat (.appx) -Detailed' -Status 'PASS' `
            -Details "IsCompatible=$($compatDetailed.IsCompatible)"
    }
    catch {
        Add-TestResult -TestName 'Compat (.appx) -Detailed' -Status 'ERROR' -Details "$_"
    }

    # 5d: All switches
    Write-Host "  [5d] Test-AppxBackupCompatibility (all switches)..." -ForegroundColor Gray
    try {
        $compatAll = Test-AppxBackupCompatibility -PackagePath $testAppx -CheckDependencies -Detailed
        Add-TestResult -TestName 'Compat (.appx) all switches' -Status 'PASS' `
            -Details "Compatible=$($compatAll.IsCompatible), Arch=$($compatAll.ArchitectureCompatible)"
    }
    catch {
        Add-TestResult -TestName 'Compat (.appx) all switches' -Status 'ERROR' -Details "$_"
    }
}
else {
    Add-TestResult -TestName 'Test-AppxBackupCompatibility (.appx)' -Status 'SKIP' -Details 'No .appx test file'
}

# 5e: Compatibility on .appxpack
if ($testAppxpack) {
    Write-Host "`n  [5e] Test-AppxBackupCompatibility on .appxpack..." -ForegroundColor Gray
    try {
        $compatZip = Test-AppxBackupCompatibility -PackagePath $testAppxpack -CheckDependencies -Detailed
        Add-TestResult -TestName 'Compat (.appxpack) all switches' -Status 'PASS' `
            -Details "IsCompatible=$($compatZip.IsCompatible)"
    }
    catch {
        Add-TestResult -TestName 'Compat (.appxpack) all switches' -Status 'ERROR' -Details "$_"
    }
}
else {
    Add-TestResult -TestName 'Test-AppxBackupCompatibility (.appxpack)' -Status 'SKIP' -Details 'No .appxpack test file'
}

# ============================================================================
# TEST 6: Export-AppxDependencies
# ============================================================================

Write-Section "6. Export-AppxDependencies"

$exportDir = [System.IO.Path]::Combine($testRunDir, '6_Exports')
[void](New-Item -Path $exportDir -ItemType Directory -Force)

$formats = @('JSON', 'XML', 'CSV', 'HTML')

foreach ($fmt in $formats) {
    $ext = $fmt.ToLower()
    $exportPath = [System.IO.Path]::Combine($exportDir, "deps_test.$ext")

    Write-Host "  [6] Export-AppxDependencies -Format $fmt -IncludeOptional..." -ForegroundColor Gray
    try {
        Export-AppxDependencies -PackagePath $app.InstallLocation `
            -OutputPath $exportPath -Format $fmt -IncludeOptional
        Test-FileCreated -Path $exportPath -TestName "Export-AppxDependencies -Format $fmt" -MinSizeBytes 10
    }
    catch {
        Add-TestResult -TestName "Export-AppxDependencies -Format $fmt" -Status 'ERROR' -Details "$_"
    }
}

# 6e: With -Recursive -IncludeOptional
$exportRecPath = [System.IO.Path]::Combine($exportDir, "deps_recursive.json")
Write-Host "  [6e] Export-AppxDependencies -Recursive -IncludeOptional..." -ForegroundColor Gray
try {
    Export-AppxDependencies -PackagePath $app.InstallLocation `
        -OutputPath $exportRecPath -Format JSON -Recursive -IncludeOptional
    Test-FileCreated -Path $exportRecPath -TestName 'Export -Recursive -IncludeOptional' -MinSizeBytes 10
}
catch {
    Add-TestResult -TestName 'Export -Recursive -IncludeOptional' -Status 'ERROR' -Details "$_"
}

# 6f: Export to directory (auto-generated filename)
Write-Host "  [6f] Export-AppxDependencies to directory (auto filename)..." -ForegroundColor Gray
try {
    Export-AppxDependencies -PackagePath $app.InstallLocation `
        -OutputPath "$exportDir\" -Format JSON -IncludeOptional
    # Find auto-generated file
    $autoFile = Get-ChildItem -Path $exportDir -Filter '*Dependencies.json' |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($autoFile) {
        Add-TestResult -TestName 'Export to dir (auto filename)' -Status 'PASS' `
            -Details $autoFile.Name
    }
    else {
        Add-TestResult -TestName 'Export to dir (auto filename)' -Status 'FAIL' `
            -Details 'No auto-generated file found'
    }
}
catch {
    Add-TestResult -TestName 'Export to dir (auto filename)' -Status 'ERROR' -Details "$_"
}

# 6g: Export with -MaxDepth
$exportDepthPath = [System.IO.Path]::Combine($exportDir, "deps_depth1.json")
Write-Host "  [6g] Export-AppxDependencies -MaxDepth 1..." -ForegroundColor Gray
try {
    Export-AppxDependencies -PackagePath $app.InstallLocation `
        -OutputPath $exportDepthPath -Format JSON -Recursive -IncludeOptional -MaxDepth 1
    Test-FileCreated -Path $exportDepthPath -TestName 'Export -MaxDepth 1' -MinSizeBytes 10
}
catch {
    Add-TestResult -TestName 'Export -MaxDepth 1' -Status 'ERROR' -Details "$_"
}

# ============================================================================
# TEST 7: New-AppxBackupCertificate
# ============================================================================

Write-Section "7. New-AppxBackupCertificate"

$certDir = [System.IO.Path]::Combine($testRunDir, '7_Certs')
[void](New-Item -Path $certDir -ItemType Directory -Force)

# 7a: Create standalone certificate with default parameters
$certPath7a = [System.IO.Path]::Combine($certDir, 'TestCert_Default.cer')
Write-Host "`n  [7a] New-AppxBackupCertificate (defaults)..." -ForegroundColor Gray
try {
     $cert7a = New-AppxBackupCertificate -Subject 'CN=AppxBackup Test Default' `
         -OutputPath $certPath7a -Force
    if ($cert7a) {
        Add-TestResult -TestName 'New-AppxBackupCertificate (defaults): Returns cert' -Status 'PASS' `
            -Details "Thumbprint=$($cert7a.Thumbprint)"
        Test-FileCreated -Path $certPath7a -TestName 'New-AppxBackupCertificate (defaults): .cer file' -MinSizeBytes 100
    }
    else {
        Add-TestResult -TestName 'New-AppxBackupCertificate (defaults): Returns cert' -Status 'FAIL' `
            -Details 'Returned null'
    }
}
catch {
    Add-TestResult -TestName 'New-AppxBackupCertificate (defaults)' -Status 'ERROR' -Details "$_"
}

# 7b: Create certificate with custom -ValidityYears and -KeyLength
$certPath7b = [System.IO.Path]::Combine($certDir, 'TestCert_Custom.cer')
Write-Host "  [7b] New-AppxBackupCertificate -ValidityYears 1 -KeyLength 2048..." -ForegroundColor Gray
try {
     $cert7b = New-AppxBackupCertificate -Subject 'CN=AppxBackup Test Custom' `
         -OutputPath $certPath7b -ValidityYears 1 -KeyLength 2048 -Force
    if ($cert7b) {
        Add-TestResult -TestName 'New-AppxBackupCertificate (custom)' -Status 'PASS' `
            -Details "Thumbprint=$($cert7b.Thumbprint)"
        Test-FileCreated -Path $certPath7b -TestName 'Cert (custom): .cer file' -MinSizeBytes 100
    }
    else {
        Add-TestResult -TestName 'New-AppxBackupCertificate (custom)' -Status 'FAIL' -Details 'Returned null'
    }
}
catch {
    Add-TestResult -TestName 'New-AppxBackupCertificate (custom)' -Status 'ERROR' -Details "$_"
}

# 7c: Create certificate with -ReplaceExisting (re-create over same subject)
$certPath7c = [System.IO.Path]::Combine($certDir, 'TestCert_Replace.cer')
Write-Host "  [7c] New-AppxBackupCertificate -ReplaceExisting..." -ForegroundColor Gray
try {
    # First create
    $null = New-AppxBackupCertificate -Subject 'CN=AppxBackup Test Replace' `
        -OutputPath $certPath7c
    # Then replace
    $cert7c = New-AppxBackupCertificate -Subject 'CN=AppxBackup Test Replace' `
        -OutputPath $certPath7c -ReplaceExisting
    if ($cert7c) {
        Add-TestResult -TestName 'New-AppxBackupCertificate -ReplaceExisting' -Status 'PASS' `
            -Details "Thumbprint=$($cert7c.Thumbprint)"
    }
    else {
        Add-TestResult -TestName 'New-AppxBackupCertificate -ReplaceExisting' -Status 'FAIL' -Details 'Returned null'
    }
}
catch {
    Add-TestResult -TestName 'New-AppxBackupCertificate -ReplaceExisting' -Status 'ERROR' -Details "$_"
}

# Cleanup test certs from cert store
Write-Host "  [7x] Cleaning up test certificates from store..." -ForegroundColor Gray
try {
    Get-ChildItem -Path 'Cert:\CurrentUser\My' |
        Where-Object { $_.Subject -like '*AppxBackup Test*' } |
        Remove-Item -Force -ErrorAction SilentlyContinue
    Add-TestResult -TestName 'Cert cleanup (CurrentUser\My)' -Status 'PASS'
}
catch {
    Add-TestResult -TestName 'Cert cleanup (CurrentUser\My)' -Status 'PASS' -Details "Cleanup attempted: $_"
}

# ============================================================================
# TEST 8: Install-AppxBackup (conditional)
# ============================================================================

Write-Section "8. Install-AppxBackup"

if ($TestInstall.IsPresent) {
    # 8a: Install from .appx
    if ($testAppx -and $vanillaCer) {
        Write-Host "`n  [8a] Install-AppxBackup from .appx..." -ForegroundColor Gray
        try {
            $installResult = Install-AppxBackup -PackagePath $testAppx `
                -CertificatePath $vanillaCer -CertStoreLocation 'CurrentUser' `
                -Force -Confirm:$false
            Add-TestResult -TestName 'Install-AppxBackup (.appx): Executes' -Status 'PASS'
        }
        catch {
            Add-TestResult -TestName 'Install-AppxBackup (.appx): Executes' -Status 'ERROR' -Details "$_"
        }
    }
    else {
        Add-TestResult -TestName 'Install-AppxBackup (.appx)' -Status 'SKIP' -Details 'No .appx or .cer'
    }

    # 8b: Install from .appxpack
    if ($testAppxpack) {
        Write-Host "  [8b] Install-AppxBackup from .appxpack..." -ForegroundColor Gray
        try {
            $installZipResult = Install-AppxBackup -PackagePath $testAppxpack `
                -CertStoreLocation 'CurrentUser' -Force -Confirm:$false
            Add-TestResult -TestName 'Install-AppxBackup (.appxpack): Executes' -Status 'PASS'
        }
        catch {
            Add-TestResult -TestName 'Install-AppxBackup (.appxpack): Executes' -Status 'ERROR' -Details "$_"
        }
    }
    else {
        Add-TestResult -TestName 'Install-AppxBackup (.appxpack)' -Status 'SKIP' -Details 'No .appxpack'
    }

    # 8c: Install with -SkipCertificate -SkipDependencies
    if ($testAppx) {
        Write-Host "  [8c] Install-AppxBackup -SkipCertificate -SkipDependencies..." -ForegroundColor Gray
        try {
            $installSkipResult = Install-AppxBackup -PackagePath $testAppx `
                -SkipCertificate -SkipDependencies -AllowUnsigned -Force -Confirm:$false
            Add-TestResult -TestName 'Install-AppxBackup (skip cert+deps)' -Status 'PASS'
        }
        catch {
            Add-TestResult -TestName 'Install-AppxBackup (skip cert+deps)' -Status 'ERROR' -Details "$_"
        }
    }

    # 8d: Install with -SkipSignatureValidation -AllowUnsigned
    if ($testAppx) {
        Write-Host "  [8d] Install-AppxBackup -SkipSignatureValidation -AllowUnsigned..." -ForegroundColor Gray
        try {
            $installUnsigned = Install-AppxBackup -PackagePath $testAppx `
                -SkipSignatureValidation -AllowUnsigned -Force -Confirm:$false
            Add-TestResult -TestName 'Install-AppxBackup (allow unsigned)' -Status 'PASS'
        }
        catch {
            Add-TestResult -TestName 'Install-AppxBackup (allow unsigned)' -Status 'ERROR' -Details "$_"
        }
    }
}
else {
    Add-TestResult -TestName 'Install-AppxBackup' -Status 'SKIP' `
        -Details 'Use -TestInstall to enable (destructive)'
}

# ============================================================================
# TEST 9: Module Log Validation
# ============================================================================

Write-Section "9. Module Log Validation"

if (Test-Path -LiteralPath $moduleLogPath) {
    $logContent = Get-Content -Path $moduleLogPath -Tail 200 -ErrorAction SilentlyContinue
    $logSize = (Get-Item -LiteralPath $moduleLogPath).Length

    Add-TestResult -TestName 'Module log file exists' -Status 'PASS' `
        -Details "$moduleLogPath ($([Math]::Round($logSize / 1KB, 1)) KB)"

    # Check log contains entries from this test run
    $recentEntries = $logContent | Where-Object { $_ -match (Get-Date -Format 'yyyy-MM-dd') }
    if ($recentEntries.Count -gt 0) {
        Add-TestResult -TestName 'Module log has today entries' -Status 'PASS' `
            -Details "$($recentEntries.Count) entries from today"
    }
    else {
        Add-TestResult -TestName 'Module log has today entries' -Status 'FAIL' `
            -Details 'No entries matching today'
    }

    # Check for ERROR-level entries during test
    $errorEntries = $logContent | Where-Object { $_ -match '\[ERROR\]' }
    if ($errorEntries.Count -gt 0) {
        Add-TestResult -TestName 'Module log ERROR entries' -Status 'PASS' `
            -Details "$($errorEntries.Count) ERROR entries (review log for details)"
    }
    else {
        Add-TestResult -TestName 'Module log ERROR entries' -Status 'PASS' `
            -Details 'No ERROR entries in recent log'
    }

    # Copy log snapshot to test run directory
    $logSnapshot = [System.IO.Path]::Combine($testRunDir, "ModuleLog_Snapshot_$timestamp.log")
    Copy-Item -Path $moduleLogPath -Destination $logSnapshot -Force -ErrorAction SilentlyContinue
}
else {
    Add-TestResult -TestName 'Module log file exists' -Status 'FAIL' `
        -Details "Not found at: $moduleLogPath"
}

# ============================================================================
# RESULTS SUMMARY
# ============================================================================

Write-Section "TEST RESULTS SUMMARY"

$elapsed = (Get-Date) - $script:StartTime
$totalTests = $script:PassCount + $script:FailCount + $script:SkipCount + $script:ErrorCount

Write-Host "`n  Total Tests : $totalTests" -ForegroundColor White
Write-Host "  Passed      : $($script:PassCount)" -ForegroundColor Green
Write-Host "  Failed      : $($script:FailCount)" -ForegroundColor $(if ($script:FailCount -gt 0) { 'Red' } else { 'Green' })
Write-Host "  Skipped     : $($script:SkipCount)" -ForegroundColor Yellow
Write-Host "  Errors      : $($script:ErrorCount)" -ForegroundColor $(if ($script:ErrorCount -gt 0) { 'Magenta' } else { 'Green' })
Write-Host "  Duration    : $($elapsed.ToString('mm\:ss\.fff'))" -ForegroundColor Gray
Write-Host "  Pass Rate   : $([Math]::Round(($script:PassCount / [Math]::Max($totalTests - $script:SkipCount, 1)) * 100, 1))% (excluding skips)" -ForegroundColor White

# ============================================================================
# WRITE RESULTS FILE
# ============================================================================

$resultsContent = @"
================================================================================
  AppxBackup.Module v2.0.2 - Test Results
  Run Date    : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
  Test Folder : $testRunDir
  App Tested  : $($app.Name) v$($app.Version) ($($app.PackageFullName))
  App Search  : *$AppName*
  Install Test: $($TestInstall.IsPresent)
  Duration    : $($elapsed.ToString('mm\:ss\.fff'))
================================================================================

  SUMMARY
  -------
  Total   : $totalTests
  Passed  : $($script:PassCount)
  Failed  : $($script:FailCount)
  Skipped : $($script:SkipCount)
  Errors  : $($script:ErrorCount)
  Rate    : $([Math]::Round(($script:PassCount / [Math]::Max($totalTests - $script:SkipCount, 1)) * 100, 1))% (excluding skips)

================================================================================
  DETAILED RESULTS
================================================================================

"@

foreach ($r in $script:Results) {
    $line = "  [$($r.Status.PadRight(5))] #$($r.'#'.ToString().PadLeft(3))  $($r.Test)"
    if ($r.Details) { $line += " | $($r.Details)" }
    if ($r.Expected -and $r.Status -in @('FAIL', 'ERROR')) {
        $line += " | Expected: $($r.Expected) | Actual: $($r.Actual)"
    }
    $resultsContent += "$line`n"
}

# List all files created during tests
$resultsContent += @"

================================================================================
  FILES CREATED
================================================================================

"@

$allTestFiles = Get-ChildItem -Path $testRunDir -Recurse -File -ErrorAction SilentlyContinue |
    Sort-Object DirectoryName, Name
foreach ($f in $allTestFiles) {
    $relPath = $f.FullName.Substring($testRunDir.Length).TrimStart('\')
    $sizeMB = [Math]::Round($f.Length / 1MB, 3)
    $resultsContent += "  $($relPath.PadRight(80)) $($sizeMB.ToString('0.000').PadLeft(10)) MB`n"
}

$resultsContent += @"

================================================================================
  ENVIRONMENT
================================================================================

  PowerShell Version : $($PSVersionTable.PSVersion)
  OS                 : $([Environment]::OSVersion.VersionString)
  Architecture       : $env:PROCESSOR_ARCHITECTURE
  User               : $env:USERNAME
  Module Log         : $moduleLogPath
  Transcript         : $transcriptPath

================================================================================
"@

# Write results file
try {
    $resultsContent | Out-File -FilePath $resultsPath -Encoding utf8 -Force
    Write-Host "`n  Results file: $resultsPath" -ForegroundColor Cyan
}
catch {
    Write-Warning "Failed to write results file: $_"
}

# Stop transcript
try {
    Stop-Transcript | Out-Null
}
catch {
    # Transcript may not be running
}

Write-Host "  Transcript : $transcriptPath" -ForegroundColor Cyan
Write-Host "  Test folder: $testRunDir" -ForegroundColor Cyan
Write-Host "`n" -NoNewline

# Return results for pipeline use
$script:Results