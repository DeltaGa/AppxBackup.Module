#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.5.0' }

<#
.SYNOPSIS
    Contract tests for the AppxBackup module manifest and public surface.

.DESCRIPTION
    These tests protect the published contract. The manifest, the loader and the
    files on disk each describe the module's public surface, and they must agree:
    a mismatch means consumers get a different module than the manifest advertises.
#>

BeforeAll {
    $script:ModuleRoot = Split-Path -Path $PSScriptRoot -Parent
    $script:ManifestPath = Join-Path $script:ModuleRoot 'AppxBackup.psd1'

    Import-Module $script:ManifestPath -Force -ErrorAction Stop
    $script:Manifest = Import-PowerShellDataFile -LiteralPath $script:ManifestPath
}

AfterAll {
    Remove-Module AppxBackup -Force -ErrorAction SilentlyContinue
}

Describe 'Module manifest' {

    It 'is a valid module manifest' {
        { Test-ModuleManifest -Path $script:ManifestPath -ErrorAction Stop } | Should -Not -Throw
    }

    It 'declares a parseable semantic version' {
        [version]$script:Manifest.ModuleVersion | Should -BeOfType ([version])
    }

    It 'declares a GUID' {
        [guid]::Parse($script:Manifest.GUID) | Should -Not -BeNullOrEmpty
    }

    It 'requires a PowerShell version the module actually targets' {
        [version]$script:Manifest.PowerShellVersion | Should -BeGreaterOrEqual ([version]'7.4')
    }

    It 'lists only files that exist on disk' {
        $missing = @($script:Manifest.FileList | Where-Object {
            -not (Test-Path -LiteralPath (Join-Path $script:ModuleRoot $_))
        })
        $missing | Should -BeNullOrEmpty -Because 'FileList drives what ships in the package'
    }

    It 'ships every function file that the loader will import' {
        $onDisk = @(
            Get-ChildItem -LiteralPath (Join-Path $script:ModuleRoot 'Public') -Filter '*.ps1' -File
            Get-ChildItem -LiteralPath (Join-Path $script:ModuleRoot 'Private') -Filter '*.ps1' -File
        ) | ForEach-Object { $_.FullName.Substring($script:ModuleRoot.Length + 1) }

        $notListed = @($onDisk | Where-Object { $script:Manifest.FileList -notcontains $_ })
        $notListed | Should -BeNullOrEmpty -Because 'a function file missing from FileList would not ship'
    }
}

Describe 'Public surface' {

    It 'exports exactly the functions named in the manifest' {
        $exported = @(Get-Command -Module AppxBackup -CommandType Function).Name | Sort-Object
        $declared = @($script:Manifest.FunctionsToExport) | Sort-Object
        Compare-Object -ReferenceObject $declared -DifferenceObject $exported | Should -BeNullOrEmpty
    }

    It 'exports one function per file under Public/' {
        $fromFiles = @(Get-ChildItem -LiteralPath (Join-Path $script:ModuleRoot 'Public') -Filter '*.ps1' -File).BaseName | Sort-Object
        $exported = @(Get-Command -Module AppxBackup -CommandType Function).Name | Sort-Object
        Compare-Object -ReferenceObject $fromFiles -DifferenceObject $exported | Should -BeNullOrEmpty
    }

    It 'exports exactly the aliases named in the manifest' {
        $exported = @(Get-Command -Module AppxBackup -CommandType Alias).Name | Sort-Object
        $declared = @($script:Manifest.AliasesToExport) | Sort-Object
        Compare-Object -ReferenceObject $declared -DifferenceObject $exported | Should -BeNullOrEmpty
    }

    It 'keeps every private function out of the public surface' {
        $privateNames = @(Get-ChildItem -LiteralPath (Join-Path $script:ModuleRoot 'Private') -Filter '*.ps1' -File).BaseName
        $exported = @(Get-Command -Module AppxBackup -CommandType Function).Name
        @($privateNames | Where-Object { $exported -contains $_ }) | Should -BeNullOrEmpty
    }

    It 'gives every exported function comment-based help with a synopsis' {
        foreach ($name in @(Get-Command -Module AppxBackup -CommandType Function).Name) {
            (Get-Help $name -ErrorAction SilentlyContinue).Synopsis |
                Should -Not -BeNullOrEmpty -Because "$name is public"
        }
    }
}

Describe 'Source integrity' {

    It 'parses every script file without syntax errors' {
        $files = Get-ChildItem -LiteralPath $script:ModuleRoot -Recurse -File -Include '*.ps1', '*.psm1' |
            Where-Object { $_.FullName -notmatch '\\\.git\\' }

        foreach ($file in $files) {
            $errors = $null
            [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$errors) | Out-Null
            $errors | Should -BeNullOrEmpty -Because "$($file.Name) must parse"
        }
    }

    It 'contains only function definitions in Public/ and Private/ files' {
        # The loader dot-sources these files in name order and relies on them having
        # no side effects. Top-level executable code would reintroduce order coupling.
        $files = @(
            Get-ChildItem -LiteralPath (Join-Path $script:ModuleRoot 'Public') -Filter '*.ps1' -File
            Get-ChildItem -LiteralPath (Join-Path $script:ModuleRoot 'Private') -Filter '*.ps1' -File
        )

        foreach ($file in $files) {
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$null)
            $nonFunction = @($ast.EndBlock.Statements | Where-Object {
                $_ -isnot [System.Management.Automation.Language.FunctionDefinitionAst]
            })
            $nonFunction | Should -BeNullOrEmpty -Because "$($file.Name) must only define functions"
        }
    }

    It 'parses every JSON configuration file' {
        foreach ($config in Get-ChildItem -LiteralPath (Join-Path $script:ModuleRoot 'Config') -Filter '*.json' -File) {
            { Get-Content -LiteralPath $config.FullName -Raw | ConvertFrom-Json -ErrorAction Stop } |
                Should -Not -Throw -Because "$($config.Name) is loaded at runtime"
        }
    }
}

Describe 'Import and removal lifecycle' {

    It 'withdraws its aliases when the module is removed' {
        # Aliases used to be created with -Scope Global, which left them resolvable in
        # the caller's session after Remove-Module.
        Import-Module $script:ManifestPath -Force
        Get-Command 'Backup-AppX' -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty

        Remove-Module AppxBackup -Force
        Get-Command 'Backup-AppX' -ErrorAction SilentlyContinue | Should -BeNullOrEmpty

        Import-Module $script:ManifestPath -Force
    }

    It 'imports repeatedly without error' {
        { 1..3 | ForEach-Object { Import-Module $script:ManifestPath -Force -ErrorAction Stop } } | Should -Not -Throw
    }
}
