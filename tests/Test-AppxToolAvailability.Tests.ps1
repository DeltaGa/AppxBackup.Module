#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.5.0' }

<#
.SYNOPSIS
    Tests for SDK tool discovery and its caching behaviour.

.DESCRIPTION
    Callers treat the return value as a single path string and feed it straight into
    [string] parameters, so the cardinality of the result is part of the contract.
#>

BeforeAll {
    $script:ModuleRoot = Split-Path -Path $PSScriptRoot -Parent
    Import-Module (Join-Path $script:ModuleRoot 'AppxBackup.psd1') -Force -ErrorAction Stop
}

AfterAll {
    Remove-Module AppxBackup -Force -ErrorAction SilentlyContinue
}

Describe 'Test-AppxToolAvailability' {

    Context 'caching' {

        It 'returns a single string on a cache hit, not a collection' {
            # Regression: the cache check lived in begin{} and used return. return in a
            # begin block exits only that block, so the cached value was emitted and the
            # full search then ran and emitted a second one. Every cache hit produced a
            # two-element array from a function declared [OutputType([string])].
            InModuleScope AppxBackup {
                $script:ToolCache.Clear()

                $cold = Test-AppxToolAvailability -ToolName 'PowerShellCert'
                $warm = Test-AppxToolAvailability -ToolName 'PowerShellCert'

                @($cold).Count | Should -Be 1
                @($warm).Count | Should -Be 1
                $warm | Should -BeOfType ([string])
                $warm | Should -Be $cold
            }
        }

        It 'populates the module cache after a successful lookup' {
            InModuleScope AppxBackup {
                $script:ToolCache.Clear()
                $null = Test-AppxToolAvailability -ToolName 'PowerShellCert'
                $script:ToolCache.ContainsKey('PowerShellCert') | Should -BeTrue
            }
        }

        It 'stays single-valued across repeated calls' {
            InModuleScope AppxBackup {
                $script:ToolCache.Clear()
                foreach ($i in 1..5) {
                    @(Test-AppxToolAvailability -ToolName 'PowerShellCert').Count | Should -Be 1
                }
            }
        }
    }

    Context 'missing tools' {

        It 'returns nothing when a tool cannot be found and no throw is requested' {
            InModuleScope AppxBackup {
                $script:ToolCache.Clear()
                Mock Get-Command -MockWith { $null } -ParameterFilter { $Name -eq 'makeappx.exe' }
                Mock Test-Path -MockWith { $false }

                $result = Test-AppxToolAvailability -ToolName 'MakeAppx' -WarningAction SilentlyContinue
                $result | Should -BeNullOrEmpty
            }
        }

        It 'throws when -ThrowOnError is supplied and the tool is absent' {
            InModuleScope AppxBackup {
                $script:ToolCache.Clear()
                Mock Get-Command -MockWith { $null } -ParameterFilter { $Name -eq 'makeappx.exe' }
                Mock Test-Path -MockWith { $false }

                { Test-AppxToolAvailability -ToolName 'MakeAppx' -ThrowOnError -WarningAction SilentlyContinue -ErrorAction Stop } |
                    Should -Throw -ExpectedMessage '*Tool not found*'
            }
        }
    }

    Context 'parameter contract' {

        It 'rejects a tool name outside the supported set' {
            InModuleScope AppxBackup {
                { Test-AppxToolAvailability -ToolName 'NotARealTool' -ErrorAction Stop } | Should -Throw
            }
        }
    }
}

Describe 'Get-AppxToolPath' {

    It 'accepts every tool name it advertises' {
        $validValues = (Get-Command Get-AppxToolPath).Parameters['ToolName'].Attributes |
            Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] } |
            Select-Object -ExpandProperty ValidValues

        $validValues | Should -Not -BeNullOrEmpty
        foreach ($name in $validValues) {
            { Get-AppxToolPath -ToolName $name -WarningAction SilentlyContinue } | Should -Not -Throw
        }
    }

    It 'returns a single value or nothing, never a collection' {
        $result = Get-AppxToolPath -ToolName 'MakeAppx' -WarningAction SilentlyContinue
        @($result).Count | Should -BeLessOrEqual 1
    }
}
