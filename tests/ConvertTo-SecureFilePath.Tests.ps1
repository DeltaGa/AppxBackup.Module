#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.5.0' }

<#
.SYNOPSIS
    Tests for the ConvertTo-SecureFilePath path validation helper.

.DESCRIPTION
    ConvertTo-SecureFilePath is the module's single gate for untrusted paths, so its
    rejection rules are load-bearing for security and its creation behaviour is relied
    on by Backup-AppxPackage to materialise the output directory.
#>

BeforeAll {
    $script:ModuleRoot = Split-Path -Path $PSScriptRoot -Parent
    Import-Module (Join-Path $script:ModuleRoot 'AppxBackup.psd1') -Force -ErrorAction Stop
}

AfterAll {
    Remove-Module AppxBackup -Force -ErrorAction SilentlyContinue
}

Describe 'ConvertTo-SecureFilePath' {

    BeforeEach {
        $script:Sandbox = Join-Path ([System.IO.Path]::GetTempPath()) "AppxBackupTests_$([guid]::NewGuid())"
        New-Item -ItemType Directory -Path $script:Sandbox -Force | Out-Null
    }

    AfterEach {
        if ($script:Sandbox -and (Test-Path -LiteralPath $script:Sandbox)) {
            Remove-Item -LiteralPath $script:Sandbox -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    Context 'existence handling' {

        It 'throws a descriptive error when -MustExist is given a missing path' {
            # Regression: the guard compared Test-Path's [bool] result against $null, so
            # it never fired. Callers fell through to Get-Item and surfaced
            # "The property 'PSIsContainer' cannot be found" instead of the real problem.
            $missing = Join-Path $script:Sandbox 'no-such-directory'

            InModuleScope AppxBackup -Parameters @{ Target = $missing } {
                param($Target)
                { ConvertTo-SecureFilePath -Path $Target -MustExist -PathType Directory -ErrorAction Stop } |
                    Should -Throw -ExpectedMessage '*does not exist*' -Because "'$Target' is absent"
            }
        }

        It 'returns the path when -MustExist is given an existing directory' {
            InModuleScope AppxBackup -Parameters @{ Target = $script:Sandbox } {
                param($Target)
                ConvertTo-SecureFilePath -Path $Target -MustExist -PathType Directory | Should -Be $Target
            }
        }

        It 'throws when the path exists but is the wrong PathType' {
            $file = Join-Path $script:Sandbox 'a-file.txt'
            Set-Content -LiteralPath $file -Value 'x'

            InModuleScope AppxBackup -Parameters @{ Target = $file } {
                param($Target)
                { ConvertTo-SecureFilePath -Path $Target -MustExist -PathType Directory -ErrorAction Stop } |
                    Should -Throw -ExpectedMessage '*is not a Directory*' -Because "'$Target' is a file"
            }
        }
    }

    Context 'creation handling' {

        It 'creates a missing directory when -CreateIfMissing is used on its own' {
            # Regression: creation was nested inside the -MustExist branch, and that
            # branch was unreachable. Backup-AppxPackage passes -CreateIfMissing without
            # -MustExist, so -OutputPath was never actually created.
            $target = Join-Path $script:Sandbox 'created-dir'

            InModuleScope AppxBackup -Parameters @{ Target = $target } {
                param($Target)
                $null = ConvertTo-SecureFilePath -Path $Target -PathType Directory -CreateIfMissing
            }

            Test-Path -LiteralPath $target -PathType Container | Should -BeTrue
        }

        It 'creates a missing file and its parent directory' {
            $target = Join-Path $script:Sandbox 'nested\deeper\created.txt'

            InModuleScope AppxBackup -Parameters @{ Target = $target } {
                param($Target)
                $null = ConvertTo-SecureFilePath -Path $Target -PathType File -CreateIfMissing
            }

            Test-Path -LiteralPath $target -PathType Leaf | Should -BeTrue
        }

        It 'refuses to create when the PathType is ambiguous' {
            $target = Join-Path $script:Sandbox 'ambiguous'

            InModuleScope AppxBackup -Parameters @{ Target = $target } {
                param($Target)
                { ConvertTo-SecureFilePath -Path $Target -PathType Any -CreateIfMissing -ErrorAction Stop } |
                    Should -Throw -ExpectedMessage "*Cannot create path of type 'Any'*" -Because "'$Target' has no determinable type"
            }
        }

        It 'leaves an existing directory untouched' {
            $marker = Join-Path $script:Sandbox 'marker.txt'
            Set-Content -LiteralPath $marker -Value 'keep me'

            InModuleScope AppxBackup -Parameters @{ Target = $script:Sandbox } {
                param($Target)
                $null = ConvertTo-SecureFilePath -Path $Target -PathType Directory -CreateIfMissing
            }

            Get-Content -LiteralPath $marker | Should -Be 'keep me'
        }
    }

    Context 'rejection rules' {

        It 'rejects <Description>' -ForEach @(
            @{ Description = 'an empty path'; Path = '   '; Expected = '*null or empty*' }
            @{ Description = 'a relative traversal sequence'; Path = 'C:\data\..\Windows'; Expected = '*traversal*' }
            @{ Description = 'a trailing traversal sequence'; Path = 'C:\data\..'; Expected = '*traversal*' }
            @{ Description = 'a reserved Windows device name'; Path = 'C:\data\CON.txt'; Expected = '*reserved*' }
            @{ Description = 'a UNC path by default'; Path = '\\server\share\file'; Expected = '*UNC*' }
        ) {
            InModuleScope AppxBackup -Parameters @{ Target = $Path; Pattern = $Expected } {
                param($Target, $Pattern)
                { ConvertTo-SecureFilePath -Path $Target -ErrorAction Stop } |
                    Should -Throw -ExpectedMessage $Pattern -Because "'$Target' must be rejected"
            }
        }

        It 'rejects a path containing an embedded null byte' {
            InModuleScope AppxBackup {
                $poisoned = "C:\data$([char]0)\file.txt"
                { ConvertTo-SecureFilePath -Path $poisoned -ErrorAction Stop } | Should -Throw
            }
        }

        It 'accepts a UNC path when -AllowUNC is supplied' {
            InModuleScope AppxBackup {
                ConvertTo-SecureFilePath -Path '\\server\share\file' -AllowUNC | Should -Be '\\server\share\file'
            }
        }
    }

    Context 'normalisation' {

        It 'converts forward slashes to backslashes' {
            InModuleScope AppxBackup {
                ConvertTo-SecureFilePath -Path 'C:/data/sub' | Should -Be 'C:\data\sub'
            }
        }

        It 'strips surrounding quotes and whitespace' {
            InModuleScope AppxBackup {
                ConvertTo-SecureFilePath -Path '  "C:\data\sub"  ' | Should -Be 'C:\data\sub'
            }
        }

        It 'returns a single string rather than a collection' {
            InModuleScope AppxBackup {
                $result = ConvertTo-SecureFilePath -Path 'C:\data\sub'
                @($result).Count | Should -Be 1
                $result | Should -BeOfType ([string])
            }
        }
    }
}
