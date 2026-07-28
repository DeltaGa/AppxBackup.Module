<#
    PSScriptAnalyzer configuration for AppxBackup.

    The intent is a gate that is worth keeping green. Rules are excluded only where
    they are inapplicable to this project by design, and each exclusion carries the
    reason. Anything not listed here is expected to pass, so a new finding means a
    real problem rather than known noise.

    Run via: .\build.ps1 -Task Analyze
#>
@{
    IncludeDefaultRules = $true

    ExcludeRules = @(
        # This module is an interactive operator tool: it reports staged progress,
        # colour-coded status and post-failure guidance straight to the console.
        # Since PowerShell 5 Write-Host writes to the information stream, so the
        # output is still redirectable and capturable via 6>. Rewriting ~350 call
        # sites would change the user-facing presentation without making the module
        # more correct. Values that callers consume are returned as objects; only
        # presentation goes through Write-Host.
        'PSAvoidUsingWriteHost'

        # Export-AppxDependencies is an exported command and part of the published
        # contract, and Resolve-AppxDependencies / Test-AppxPackagingPrerequisites
        # are its internal counterparts. Renaming the public command to satisfy a
        # naming rule would break every existing caller for no functional gain; the
        # private names are kept aligned with it deliberately.
        'PSUseSingularNouns'

        # Flagged only on private helpers (New-AppxPackageInternal,
        # New-AppxBackupZipArchive, New-AppxBackupManifest, Remove-AppxItemWithRetry
        # and friends). Confirmation is handled once, at the public entry point that
        # the user actually invokes; adding SupportsShouldProcess to internal helpers
        # would prompt repeatedly for a single logical operation.
        'PSUseShouldProcessForStateChangingFunctions'

        # Formatting only, and present on roughly 830 lines of the existing sources.
        # See the note below on why this gate stays scoped to correctness.
        'PSAvoidTrailingWhitespace'
    )

    Rules = @{
        # The module targets PowerShell 7.4+ on Windows, as declared in the manifest.
        PSUseCompatibleSyntax = @{
            Enable = $true
            TargetVersions = @('7.0')
        }
    }

    # Deliberately NOT enabled: PSPlaceOpenBrace, PSPlaceCloseBrace,
    # PSUseConsistentIndentation, PSUseConsistentWhitespace and
    # PSAvoidTrailingWhitespace. Turning them on reports over 1600 findings against
    # the existing sources. Satisfying them means reformatting the whole codebase,
    # which would produce a diff large enough to hide real changes in history and in
    # review for no correctness benefit. This gate is scoped to correctness rules so
    # that a finding is always worth acting on. Formatting is better addressed as a
    # deliberate one-off pass with Invoke-Formatter if the project wants it.
}
