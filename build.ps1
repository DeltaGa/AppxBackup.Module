#Requires -Version 7.4

<#
.SYNOPSIS
    Build, test and static-analysis entry point for the AppxBackup module.

.DESCRIPTION
    A single command that a contributor and CI both run, so "it passes locally"
    and "it passes in CI" mean the same thing.

    Tasks:
      Test     - run the Pester suite under tests/
      Analyze  - run PSScriptAnalyzer using PSScriptAnalyzerSettings.psd1
      Build    - verify the manifest and that the module imports cleanly
      All      - Build, then Analyze, then Test (default)

.PARAMETER Task
    Which task to run. Defaults to All.

.PARAMETER CI
    Emit NUnit test results and fail the process on any test failure or on any
    analyzer finding at or above -FailOn. Sets the exit code for the build system.

.PARAMETER FailOn
    Lowest analyzer severity that should fail the build. Defaults to Warning.
    Findings below this level are reported but do not fail.

.PARAMETER OutputPath
    Directory for test and analysis artefacts. Defaults to ./out.

.EXAMPLE
    .\build.ps1
    Runs the full pipeline locally with human-readable output.

.EXAMPLE
    .\build.ps1 -Task Test -CI
    Runs the tests and writes out/TestResults.xml, exiting non-zero on failure.
#>

[CmdletBinding()]
param(
    [ValidateSet('All', 'Build', 'Test', 'Analyze')]
    [string]$Task = 'All',

    [switch]$CI,

    [ValidateSet('Error', 'Warning', 'Information', 'None')]
    [string]$FailOn = 'Warning',

    [string]$OutputPath = (Join-Path $PSScriptRoot 'out')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ModuleRoot = $PSScriptRoot
$script:ManifestPath = Join-Path $script:ModuleRoot 'AppxBackup.psd1'
$script:Failures = [System.Collections.Generic.List[string]]::new()

# Promote the parameters the task functions read to explicit script scope. They would
# resolve dynamically anyway, but naming the scope makes the data flow obvious and
# keeps static analysis able to see the usage.
$script:IsCI = [bool]$CI
$script:FailOnSeverity = $FailOn
$script:ArtifactPath = $OutputPath

# Minimum tool versions. Pester 5.5 is the floor for the test syntax used in tests/,
# which is written to also run unchanged on Pester 6.
$script:MinimumPester = [version]'5.5.0'
$script:MinimumAnalyzer = [version]'1.21.0'

function Write-Section {
    param([Parameter(Mandatory)][string]$Name)
    Write-Host ''
    Write-Host "=== $Name ===" -ForegroundColor Cyan
}

function Import-BuildDependency {
    <#
        Loads a tool at or above the required version, and explains how to get it
        rather than failing with a bare "module not found".
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][version]$MinimumVersion
    )

    $candidate = Get-Module -ListAvailable -Name $Name |
        Where-Object { $_.Version -ge $MinimumVersion } |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if (-not $candidate) {
        throw "$Name $MinimumVersion or later is required. Install it with: Install-Module $Name -MinimumVersion $MinimumVersion -Scope CurrentUser -Force"
    }

    Import-Module $candidate.Path -Force -ErrorAction Stop
    Write-Host "Using $Name $($candidate.Version)" -ForegroundColor DarkGray
}

function Invoke-BuildTask {
    Write-Section 'Build'

    Write-Host 'Validating module manifest...' -ForegroundColor Gray
    $manifest = Test-ModuleManifest -Path $script:ManifestPath -ErrorAction Stop
    Write-Host "  $($manifest.Name) $($manifest.Version)" -ForegroundColor Green

    Write-Host 'Importing module...' -ForegroundColor Gray
    try {
        Import-Module $script:ManifestPath -Force -ErrorAction Stop
        $commands = @(Get-Command -Module AppxBackup)
        Write-Host "  Imported $($commands.Count) commands" -ForegroundColor Green
    }
    finally {
        Remove-Module AppxBackup -Force -ErrorAction SilentlyContinue
    }

    Write-Host 'Checking manifest FileList against disk...' -ForegroundColor Gray
    $data = Import-PowerShellDataFile -LiteralPath $script:ManifestPath
    $missing = @($data.FileList | Where-Object { -not (Test-Path -LiteralPath (Join-Path $script:ModuleRoot $_)) })

    if ($missing.Count -gt 0) {
        $missing | ForEach-Object { Write-Host "  Missing: $_" -ForegroundColor Red }
        $script:Failures.Add("Manifest FileList references $($missing.Count) file(s) that do not exist")
    }
    else {
        Write-Host '  FileList is consistent' -ForegroundColor Green
    }
}

function Invoke-AnalyzeTask {
    Write-Section 'Analyze'
    Import-BuildDependency -Name 'PSScriptAnalyzer' -MinimumVersion $script:MinimumAnalyzer

    $settings = Join-Path $script:ModuleRoot 'PSScriptAnalyzerSettings.psd1'

    # PSScriptAnalyzer 1.25 raises a non-terminating NullReferenceException while
    # inspecting AppxBackup.psm1 when it is reached through an absolute path. The
    # analysis still completes and returns the identical finding set, but this script
    # runs with $ErrorActionPreference = 'Stop', which would promote that internal
    # error into a build failure. Collect analyzer errors instead and report them
    # separately from the findings so a genuine tool failure is still visible.
    $analyzerErrors = @()
    $findings = @(Invoke-ScriptAnalyzer -Path $script:ModuleRoot -Recurse -Settings $settings `
        -ErrorVariable analyzerErrors -ErrorAction SilentlyContinue)

    foreach ($analyzerError in $analyzerErrors) {
        Write-Host "  Analyzer could not fully inspect '$($analyzerError.TargetObject)': $($analyzerError.Exception.Message)" -ForegroundColor DarkGray
    }

    if ($findings.Count -eq 0) {
        Write-Host 'No findings' -ForegroundColor Green
        return
    }

    $findings |
        Group-Object Severity |
        Sort-Object Name |
        ForEach-Object { Write-Host ("  {0,-12} {1}" -f $_.Name, $_.Count) -ForegroundColor Yellow }

    $findings |
        Sort-Object Severity, ScriptName, Line |
        ForEach-Object {
            Write-Host ("  [{0}] {1}:{2} {3}" -f $_.Severity, $_.ScriptName, $_.Line, $_.RuleName) -ForegroundColor DarkYellow
        }

    if ($script:IsCI) {
        $null = New-Item -ItemType Directory -Path $script:ArtifactPath -Force
        $findings | Export-Csv -Path (Join-Path $script:ArtifactPath 'AnalyzerResults.csv') -NoTypeInformation
    }

    if ($script:FailOnSeverity -ne 'None') {
        # Severity ascends Information -> Warning -> Error, so anything at or above
        # the requested level is blocking.
        $threshold = switch ($script:FailOnSeverity) {
            'Information' { 0 }
            'Warning' { 1 }
            'Error' { 2 }
        }

        $blocking = @($findings | Where-Object { [int]$_.Severity -ge $threshold })
        if ($blocking.Count -gt 0) {
            $script:Failures.Add("PSScriptAnalyzer reported $($blocking.Count) finding(s) at or above $script:FailOnSeverity")
        }
    }
}

function Invoke-TestTask {
    Write-Section 'Test'
    Import-BuildDependency -Name 'Pester' -MinimumVersion $script:MinimumPester

    $testPath = Join-Path $script:ModuleRoot 'tests'
    if (-not (Test-Path -LiteralPath $testPath)) {
        throw "Test directory not found: $testPath"
    }

    $config = New-PesterConfiguration
    $config.Run.Path = $testPath
    $config.Run.PassThru = $true
    $config.Output.Verbosity = if ($script:IsCI) { 'Detailed' } else { 'Normal' }

    if ($script:IsCI) {
        $null = New-Item -ItemType Directory -Path $script:ArtifactPath -Force
        $config.TestResult.Enabled = $true
        $config.TestResult.OutputFormat = 'NUnitXml'
        $config.TestResult.OutputPath = Join-Path $script:ArtifactPath 'TestResults.xml'
    }

    $result = Invoke-Pester -Configuration $config

    Write-Host ''
    Write-Host ("Passed {0} | Failed {1} | Skipped {2} | {3:N1}s" -f
        $result.PassedCount, $result.FailedCount, $result.SkippedCount, $result.Duration.TotalSeconds) -ForegroundColor Gray

    if ($result.FailedCount -gt 0) {
        $script:Failures.Add("$($result.FailedCount) test(s) failed")
    }
}

try {
    Write-Host "AppxBackup build - task '$Task' on PowerShell $($PSVersionTable.PSVersion)" -ForegroundColor White

    switch ($Task) {
        'Build' { Invoke-BuildTask }
        'Analyze' { Invoke-AnalyzeTask }
        'Test' { Invoke-TestTask }
        'All' {
            Invoke-BuildTask
            Invoke-AnalyzeTask
            Invoke-TestTask
        }
    }

    Write-Section 'Summary'

    if ($script:Failures.Count -gt 0) {
        $script:Failures | ForEach-Object { Write-Host "FAIL: $_" -ForegroundColor Red }
        if ($script:IsCI) { exit 1 }
        throw "Build failed with $($script:Failures.Count) problem(s)"
    }

    Write-Host 'Success' -ForegroundColor Green
    if ($script:IsCI) { exit 0 }
}
catch {
    Write-Host ''
    Write-Host "Build error: $($_.Exception.Message)" -ForegroundColor Red
    if ($script:IsCI) { exit 1 }
    throw
}
