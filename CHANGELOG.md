# Changelog

All notable changes to AppxBackup are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

Releases before 2.1.0 were recorded only in the `ReleaseNotes` field of
`AppxBackup.psd1`; they are summarised here for continuity.

## [Unreleased]

## [2.1.0] - 2026-07-27

A correctness release. Six defects shared one root cause: a guard compared the
result of a boolean-returning call against `$null`. Because `[bool]` is never
`$null`, each condition was always false and the code it guarded never ran. One
of them could hang a backup indefinitely; another silently corrupted every
captured tool output. This release also gives the project its first automated
tests, a static-analysis configuration, and CI.

### Fixed

- **Process timeouts were never enforced.** `Invoke-ProcessSafely` called
  `WaitForExit([int])`, which returns `$false` on timeout, and tested the result
  against `$null`. The kill branch was unreachable, so control fell through to the
  unbounded `WaitForExit()` and blocked the caller forever. `MakeAppx` is invoked
  with a 1800 second timeout, so a wedged tool hung the session with no recourse.
  Measured against the previous commit: a 3 second timeout ran a 30 second process
  to completion and reported success. It now aborts at 3 seconds and kills the
  process tree.
- **Failed process starts went undetected.** `Process.Start()` returns `[bool]`;
  the same `$null` comparison meant a failure to start was never caught.
- **Captured tool output was corrupted.** The stdout and stderr buffers were built
  with `[StringBuilder]::new($capacity)` where `$capacity` came from
  `ConvertFrom-Json` and is therefore `Int64`. `StringBuilder` offers
  `(Int32 capacity)` and `(String value)` but no `Int64` overload, so PowerShell
  bound the string one and seeded each buffer with its own capacity as text. Every
  captured stream began with `16384` or `4096`, which fed directly into MakeAppx
  error analysis and the diagnostics printed on failure.
- **`-MustExist` was not enforced.** `ConvertTo-SecureFilePath` fell through to
  `Get-Item` on a missing path, so callers saw
  `The property 'PSIsContainer' cannot be found on this object` instead of a
  usable message.
- **`-CreateIfMissing` created nothing.** In addition to the dead guard, creation
  was nested inside the `-MustExist` branch. `Backup-AppxPackage` passes
  `-CreateIfMissing` on its own, so `-OutputPath` was never created.
- **Tool discovery returned an array on cache hits.** The cache check used `return`
  inside a `begin` block. `return` there exits only that block, so the cached value
  was emitted and the full search then emitted a second one — a two-element array
  from a function declared `[OutputType([string])]`.
- **Real errors were replaced by strict-mode noise.** `Backup-AppxPackage` read
  `$packageOutputPath` in its `catch` block. When a failure occurred before that
  variable was assigned, `Set-StrictMode` raised
  `variable cannot be retrieved` and the genuine cause was lost.
- **Logger guards.** Removed a dead preference check in `Write-AppxLog` and
  corrected a `$null` comparison against `Split-Path -Parent`, which returns an
  empty string rather than `$null`.

### Changed

- Aliases are now exported module members rather than being created with
  `-Scope Global`. `Remove-Module` withdraws them instead of leaving them behind in
  the caller's session. The four alias names are unchanged.
- The module loader discovers `Public/` and `Private/` from disk instead of using
  two hand-maintained ordered lists. Every file in those folders defines only
  functions, so load order carries no meaning — PowerShell resolves a call target
  when a function runs, not when it is defined. Adding a file no longer requires
  registering it separately, which removes a class of silent load failure.
- `CompatiblePSEditions` is now `Core` only. The Desktop edition tops out at
  Windows PowerShell 5.1 and could never satisfy `PowerShellVersion = '7.4'`, so
  advertising it made the module look installable where it is not.
- `AppxBackup.psd1` and `Private/Write-AppxLog.ps1` contain non-ASCII characters
  and now carry a UTF-8 BOM, so ANSI-defaulting hosts read them correctly.

### Added

- `tests/` — 71 Pester tests covering path validation, external process execution,
  SDK discovery, configuration loading, and the manifest/loader/disk agreement that
  defines the public surface. Written in the subset shared by Pester 5 and 6. Every
  defect listed above has a regression test that fails against the old behaviour.
- `build.ps1` — one entry point for `Build`, `Analyze` and `Test`, used by both
  contributors and CI so that local and CI results mean the same thing. `-CI`
  produces NUnit test results and a CSV of analyzer findings, and sets the exit code.
- `PSScriptAnalyzerSettings.psd1` — a correctness-scoped rule set. Each exclusion
  records why the rule does not apply here, so any finding is worth acting on.
- `.github/workflows/ci.yml` — Windows CI across PowerShell 7.4 and 7.5.
- `.gitignore` — notably excludes certificates and key material, and the package
  artefacts produced by running the module in place.
- This changelog.

### Known issues

- PSScriptAnalyzer reports 35 warnings that predate this release: unused variables
  and parameters, and deliberately silent catch blocks in the logger. CI blocks on
  errors and reports warnings without failing. See "Recommended next steps" in
  `README.md`.
- Four declared parameters are accepted but ignored: `-CheckCapabilities`
  (`Test-AppxBackupCompatibility`), `-ValidateSchema` (`Get-AppxManifestData`),
  `-CreateBundle` (`New-AppxPackageInternal`) and `-OutputDirectory`
  (`New-AppxBackupManifest`). Only the first is on a public command.
- PSScriptAnalyzer 1.25 raises an internal `NullReferenceException` while inspecting
  `AppxBackup.psm1` through an absolute path. Analysis still completes and returns
  the full finding set; `build.ps1` reports the condition instead of failing on it.

## [2.0.2] - 2026-02-13

### Added

- Eleven private helper functions extracted from core logic, removing roughly 600
  lines of duplicated code from `Backup-AppxPackage` and `New-AppxPackageInternal`.
- Three-tier fallback directory copying (Robocopy, `Copy-Item`, .NET APIs).
- Disk space validation, prerequisite checking and architecture compatibility checks.
- Certificate store installation and a SignTool execution wrapper.
- `Test-AppxBackupModule.ps1` interactive harness covering the 8 public functions.

## [2.0.1] - 2026-01

### Fixed

- Recursive dependency resolution.

## [2.0.0] - 2025

### Added

- Initial 2.x line: ZIP-based dependency packaging (`.appxpack`) with installation
  orchestration, native certificate management, and MSIX support.

[Unreleased]: https://github.com/DeltaGa/AppxBackup.Module/compare/v2.1.0...HEAD
[2.1.0]: https://github.com/DeltaGa/AppxBackup.Module/compare/v2.0.2...v2.1.0
[2.0.2]: https://github.com/DeltaGa/AppxBackup.Module/compare/v2.0.1...v2.0.2
[2.0.1]: https://github.com/DeltaGa/AppxBackup.Module/compare/v2.0.0...v2.0.1
[2.0.0]: https://github.com/DeltaGa/AppxBackup.Module/releases/tag/v2.0.0
