#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.5.0' }

<#
.SYNOPSIS
    End-to-end packaging tests against the real Windows SDK MakeAppx.

.DESCRIPTION
    These exercise the actual packaging pipeline rather than mocks. They need the
    Windows SDK on the machine, so they carry the 'E2E' tag and are excluded from
    the default suite:

      Invoke-Pester -Path ./tests                              # skips this file's tests
      Invoke-Pester -Path ./tests -TagFilter 'E2E'             # runs only these

    The fixture is a minimal but valid package directory synthesised in a temp
    location. No installed application or administrator rights are required, only
    makeappx.exe. On a machine without the SDK every test is skipped.
#>

BeforeDiscovery {
    # SDK detection must happen here, not in BeforeAll: -Skip on It blocks is
    # evaluated during discovery, before any BeforeAll code has run.
    $script:ModuleRoot = Split-Path -Path $PSScriptRoot -Parent
    Import-Module (Join-Path $script:ModuleRoot 'AppxBackup.psd1') -Force -ErrorAction Stop

    $script:MakeAppx = InModuleScope AppxBackup {
        Test-AppxToolAvailability -ToolName 'MakeAppx' -WarningAction SilentlyContinue
    }
    $script:HasMakeAppx = -not [string]::IsNullOrEmpty($script:MakeAppx)
}

BeforeAll {
    # Another file's AfterAll may have removed the module between discovery and
    # this file's run; make sure it is loaded before InModuleScope is used.
    # $PSScriptRoot is used here rather than the discovery-time variable: in
    # Pester, $script: values set in BeforeDiscovery do not persist into the
    # run phase.
    if (-not (Get-Module AppxBackup)) {
        Import-Module (Join-Path (Split-Path -Path $PSScriptRoot -Parent) 'AppxBackup.psd1') -Force -ErrorAction Stop
    }

    # Build a minimal, valid package source tree.
    function New-PackageFixture {
        param([Parameter(Mandatory)][string]$Root)

        $src = Join-Path $Root 'src'
        New-Item -ItemType Directory -Path (Join-Path $src 'Assets') -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $src 'app.exe') -Value 'stub' -Encoding ascii
        [System.IO.File]::WriteAllBytes((Join-Path $src 'Assets\logo.png'),
            [byte[]](0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A))

        $ns = 'http://schemas.microsoft.com/appx/manifest/foundation/windows10'
        $uap = 'http://schemas.microsoft.com/appx/manifest/uap/windows10'
        $rescap = 'http://schemas.microsoft.com/appx/manifest/foundation/windows10/restrictedcapabilities'

        Set-Content -LiteralPath (Join-Path $src 'AppxManifest.xml') -Encoding utf8 -Value @(
            '<?xml version="1.0" encoding="utf-8"?>'
            ('<Package xmlns="{0}" xmlns:uap="{1}" xmlns:rescap="{2}" IgnorableNamespaces="uap rescap">' -f $ns, $uap, $rescap)
            '  <Identity Name="AppxBackup.E2ETest" Publisher="CN=AppxBackupLocalTest" Version="1.0.0.0" ProcessorArchitecture="x64" />'
            '  <Properties><DisplayName>E2ETest</DisplayName><PublisherDisplayName>Local Test</PublisherDisplayName><Logo>Assets\logo.png</Logo></Properties>'
            '  <Dependencies><TargetDeviceFamily Name="Windows.Universal" MinVersion="10.0.17763.0" MaxVersionTested="10.0.17763.0" /></Dependencies>'
            '  <Resources><Resource Language="en-us" /></Resources>'
            '  <Capabilities><rescap:Capability Name="runFullTrust" /></Capabilities>'
            '  <Applications><Application Id="App" Executable="app.exe" EntryPoint="Windows.FullTrustApplication">'
            '    <uap:VisualElements DisplayName="E2ETest" Description="E2ETest" BackgroundColor="transparent" Square150x150Logo="Assets\logo.png" Square44x44Logo="Assets\logo.png" />'
            '  </Application></Applications>'
            '</Package>'
        )

        return $src
    }
}

AfterAll {
    Remove-Module AppxBackup -Force -ErrorAction SilentlyContinue
}

Describe 'New-AppxPackageInternal against the real MakeAppx' -Tag 'E2E' {

    BeforeEach {
        $script:Sandbox = Join-Path ([System.IO.Path]::GetTempPath()) "AppxE2ETests_$([guid]::NewGuid())"
        New-Item -ItemType Directory -Path $script:Sandbox -Force | Out-Null
    }

    AfterEach {
        if ($script:Sandbox -and (Test-Path -LiteralPath $script:Sandbox)) {
            Remove-Item -LiteralPath $script:Sandbox -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'is skipped when the Windows SDK is not installed' -Skip:$script:HasMakeAppx {
        # Landing here means MakeAppx was not found. This test exists so the run
        # reports a visible skip reason instead of silence.
        Set-ItResult -Skipped -Because 'makeappx.exe was not found on this machine'
    }

    It 'packages a valid source directory into a real .appx' -Skip:(-not $script:HasMakeAppx) {
        $src = New-PackageFixture -Root $script:Sandbox
        $target = Join-Path $script:Sandbox 'E2ETest.appx'

        $result = InModuleScope AppxBackup -Parameters @{ Source = $src; Target = $target } {
            param($Source, $Target)
            New-AppxPackageInternal -SourcePath $Source -OutputPath $Target -CompressionLevel Normal
        }

        $result.Success | Should -BeTrue
        Test-Path -LiteralPath $target | Should -BeTrue
        $result.PackageSize | Should -BeGreaterThan 0

        # A real APPX is a ZIP whose first entry set includes the block map
        Add-Type -AssemblyName 'System.IO.Compression.FileSystem'
        $archive = [System.IO.Compression.ZipFile]::OpenRead($target)
        try {
            @($archive.Entries | Where-Object { $_.FullName -eq 'AppxManifest.xml' }).Count | Should -Be 1
            @($archive.Entries | Where-Object { $_.FullName -eq 'AppxBlockMap.xml' }).Count | Should -Be 1
        }
        finally {
            $archive.Dispose()
        }
    }

    It 'surfaces the MakeAppx error text when the manifest is invalid' -Skip:(-not $script:HasMakeAppx) {
        # Regression guard for the output-capture fix: before it, a failed packaging
        # run surfaced nothing but the MakeAppx banner, because event-based output
        # collection dropped nearly every line.
        $src = New-PackageFixture -Root $script:Sandbox
        Set-Content -LiteralPath (Join-Path $src 'AppxManifest.xml') -Value 'this is not xml' -Encoding utf8
        $target = Join-Path $script:Sandbox 'Broken.appx'

        $thrown = InModuleScope AppxBackup -Parameters @{ Source = $src; Target = $target } {
            param($Source, $Target)
            try {
                New-AppxPackageInternal -SourcePath $Source -OutputPath $Target -ErrorAction Stop
                $null
            }
            catch {
                $_.Exception.Message
            }
        }

        $thrown | Should -Not -BeNullOrEmpty
        # The analysis must contain the tool's own diagnosis, not just an exit code.
        $thrown | Should -Match '(?i)manifest|MakeAppx|0x[0-9A-F]{8}'
        Test-Path -LiteralPath $target | Should -BeFalse
    }

    It 'reports the correct size in the result object' -Skip:(-not $script:HasMakeAppx) {
        $src = New-PackageFixture -Root $script:Sandbox
        $target = Join-Path $script:Sandbox 'E2ETest.appx'

        $result = InModuleScope AppxBackup -Parameters @{ Source = $src; Target = $target } {
            param($Source, $Target)
            New-AppxPackageInternal -SourcePath $Source -OutputPath $Target -CompressionLevel Normal
        }

        $actual = (Get-Item -LiteralPath $target).Length
        $result.PackageSize | Should -Be $actual
        $result.PackageSizeMB | Should -Be ([Math]::Round($actual / 1MB, 2))
    }
}
