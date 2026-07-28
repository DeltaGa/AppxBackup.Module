#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.5.0' }

<#
.SYNOPSIS
    Tests for the JSON configuration loader and the default-value accessor.

.DESCRIPTION
    Every tunable in the module is read through these two functions, and several of
    the values are fed to .NET APIs. Types matter as much as values here: JSON numbers
    arrive as Int64, which has already caused one overload-binding defect.
#>

BeforeAll {
    $script:ModuleRoot = Split-Path -Path $PSScriptRoot -Parent
    Import-Module (Join-Path $script:ModuleRoot 'AppxBackup.psd1') -Force -ErrorAction Stop

    $script:ConfigNames = @(
        'ToolConfiguration'
        'WindowsReservedNames'
        'PackageConfiguration'
        'ModuleDefaults'
        'ZipPackagingConfiguration'
    )
}

AfterAll {
    Remove-Module AppxBackup -Force -ErrorAction SilentlyContinue
}

Describe 'Get-AppxConfiguration' {

    It 'loads and validates <_>' -ForEach @(
        'ToolConfiguration'
        'WindowsReservedNames'
        'PackageConfiguration'
        'ModuleDefaults'
        'ZipPackagingConfiguration'
    ) {
        $name = $_
        InModuleScope AppxBackup -Parameters @{ Name = $name } {
            param($Name)
            $config = Get-AppxConfiguration -ConfigName $Name
            $config | Should -Not -BeNullOrEmpty
        }
    }

    It 'returns the identical cached instance on a second call' {
        InModuleScope AppxBackup {
            $first = Get-AppxConfiguration -ConfigName 'ModuleDefaults'
            $second = Get-AppxConfiguration -ConfigName 'ModuleDefaults'
            [object]::ReferenceEquals($first, $second) | Should -BeTrue
        }
    }

    It 'returns a fresh instance when -Reload is supplied' {
        InModuleScope AppxBackup {
            $first = Get-AppxConfiguration -ConfigName 'ModuleDefaults'
            $reloaded = Get-AppxConfiguration -ConfigName 'ModuleDefaults' -Reload
            [object]::ReferenceEquals($first, $reloaded) | Should -BeFalse
            $reloaded | Should -Not -BeNullOrEmpty
        }
    }

    It 'rejects a configuration name outside the supported set' {
        InModuleScope AppxBackup {
            { Get-AppxConfiguration -ConfigName 'NoSuchConfig' -ErrorAction Stop } | Should -Throw
        }
    }

    It 'exposes a Config file for every supported configuration name' {
        foreach ($name in $script:ConfigNames) {
            Test-Path -LiteralPath (Join-Path $script:ModuleRoot "Config\$name.json") |
                Should -BeTrue -Because "$name is in the ValidateSet"
        }
    }
}

Describe 'Get-AppxDefault' {

    It 'resolves a nested value using dot notation' {
        InModuleScope AppxBackup {
            Get-AppxDefault 'bufferSizes.stdoutBuilderInitialCapacity' | Should -Be 16384
        }
    }

    It 'returns the fallback when the path does not exist' {
        InModuleScope AppxBackup {
            Get-AppxDefault 'no.such.path' -Fallback 'sentinel' | Should -Be 'sentinel'
        }
    }

    It 'returns the fallback when the configuration name is unknown' {
        InModuleScope AppxBackup {
            Get-AppxDefault 'any.path' 'NotAConfig' 'sentinel' | Should -Be 'sentinel'
        }
    }

    It 'reads from a non-default configuration file when one is named' {
        InModuleScope AppxBackup {
            Get-AppxDefault 'archiveStructure.packagesDirectory' 'ZipPackagingConfiguration' 'fallback' |
                Should -Not -Be 'fallback'
        }
    }

    It 'preserves the underlying JSON type rather than stringifying' {
        InModuleScope AppxBackup {
            $value = Get-AppxDefault 'bufferSizes.stdoutBuilderInitialCapacity'
            $value | Should -BeOfType ([long])
        }
    }

    It 'returns numbers as Int64, which callers must cast before binding .NET overloads' {
        # Documents the trap that produced the StringBuilder(String) defect: an Int64
        # will not bind an Int32-only overload when a String overload is also present.
        InModuleScope AppxBackup {
            $capacity = Get-AppxDefault 'bufferSizes.stdoutBuilderInitialCapacity' -Fallback 16384
            [System.Text.StringBuilder]::new([int]$capacity).Length | Should -Be 0
        }
    }
}

Describe 'Module configuration bootstrap' {

    It 'populates the runtime configuration table from ModuleDefaults.json' {
        InModuleScope AppxBackup {
            $script:AppxBackupConfig | Should -Not -BeNullOrEmpty
            $script:AppxBackupConfig.DefaultHashAlgorithm | Should -Be 'SHA256'
            $script:AppxBackupConfig.DefaultCertificateValidityYears | Should -BeGreaterThan 0
            $script:AppxBackupConfig.DefaultKeyLength | Should -BeGreaterOrEqual 2048
        }
    }

    It 'does not export the configuration table to the caller' {
        Get-Variable -Name 'AppxBackupConfig' -Scope Global -ErrorAction SilentlyContinue |
            Should -BeNullOrEmpty
    }
}
