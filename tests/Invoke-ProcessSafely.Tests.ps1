#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.5.0' }

<#
.SYNOPSIS
    Tests for the Invoke-ProcessSafely external process wrapper.

.DESCRIPTION
    Every external tool the module runs (MakeAppx, SignTool, Robocopy) goes through
    this wrapper, and MakeAppx is invoked with a 1800 second timeout. If the timeout
    is not enforced a wedged tool blocks the caller with no way out, so the timeout
    path is covered explicitly.
#>

BeforeAll {
    $script:ModuleRoot = Split-Path -Path $PSScriptRoot -Parent
    Import-Module (Join-Path $script:ModuleRoot 'AppxBackup.psd1') -Force -ErrorAction Stop

    # Long-running and trivially available on every supported Windows build.
    $script:Ping = Join-Path $env:SystemRoot 'System32\PING.EXE'
    $script:Cmd = Join-Path $env:SystemRoot 'System32\cmd.exe'
}

AfterAll {
    Remove-Module AppxBackup -Force -ErrorAction SilentlyContinue
}

Describe 'Invoke-ProcessSafely' {

    Context 'timeout enforcement' {

        It 'terminates a process that outruns its timeout, and does so promptly' {
            # Regression: WaitForExit([int]) returns $false on timeout, but the guard
            # tested it against $null. The kill branch was unreachable and control fell
            # through to the unbounded WaitForExit(), hanging the caller. Before the fix
            # this call returned success only after the process ended on its own.
            $elapsed = InModuleScope AppxBackup -Parameters @{ Exe = $script:Ping } {
                param($Exe)
                $sw = [System.Diagnostics.Stopwatch]::StartNew()
                try {
                    Invoke-ProcessSafely -FilePath $Exe -ArgumentList @('-n', '30', '127.0.0.1') `
                        -TimeoutSeconds 2 -NoWindow -ErrorAction Stop | Out-Null
                    throw 'PESTER-NO-TIMEOUT'
                }
                catch {
                    if ($_.Exception.Message -eq 'PESTER-NO-TIMEOUT') { throw }
                    $_.Exception.Message | Should -BeLike '*timed out after 2 seconds*'
                }
                $sw.Elapsed.TotalSeconds
            }

            # The process runs for roughly 29s; anything near that means it was awaited.
            $elapsed | Should -BeLessThan 15
        }

        It 'leaves no orphaned child process behind after a timeout' {
            $before = @(Get-Process -Name 'PING' -ErrorAction SilentlyContinue).Count

            InModuleScope AppxBackup -Parameters @{ Exe = $script:Ping } {
                param($Exe)
                try {
                    Invoke-ProcessSafely -FilePath $Exe -ArgumentList @('-n', '30', '127.0.0.1') `
                        -TimeoutSeconds 2 -NoWindow -ErrorAction Stop | Out-Null
                }
                catch {
                    # The timeout is expected here; this test asserts on the process
                    # table afterwards rather than on the exception.
                    Write-Verbose "Expected timeout: $($_.Exception.Message)"
                }
            }

            Start-Sleep -Milliseconds 500
            @(Get-Process -Name 'PING' -ErrorAction SilentlyContinue).Count | Should -BeLessOrEqual $before
        }
    }

    Context 'successful execution' {

        It 'reports empty output for a process that writes nothing' {
            # Regression: the capture buffers were built with
            # [StringBuilder]::new($capacity) where $capacity came from ConvertFrom-Json
            # as Int64. With no Int64 overload PowerShell bound StringBuilder(String),
            # so every capture started with its own capacity as text - stdout began
            # "16384" and stderr "4096", corrupting all downstream output parsing.
            InModuleScope AppxBackup -Parameters @{ Exe = $script:Cmd } {
                param($Exe)
                $result = Invoke-ProcessSafely -FilePath $Exe -ArgumentList @('/c', 'exit 0') -NoWindow
                $result.StandardOutput | Should -BeNullOrEmpty
                $result.StandardError | Should -BeNullOrEmpty
            }
        }

        It 'does not prefix captured output with the buffer capacity' {
            InModuleScope AppxBackup -Parameters @{ Exe = $script:Cmd } {
                param($Exe)
                $result = Invoke-ProcessSafely -FilePath $Exe -ArgumentList @('/c', 'echo Marker') -NoWindow
                $result.StandardOutput.TrimStart() | Should -BeLike 'Marker*'
            }
        }

        It 'reports success and a zero exit code' {
            InModuleScope AppxBackup -Parameters @{ Exe = $script:Cmd } {
                param($Exe)
                $result = Invoke-ProcessSafely -FilePath $Exe -ArgumentList @('/c', 'exit 0') -NoWindow
                $result.Success | Should -BeTrue
                $result.ExitCode | Should -Be 0
            }
        }

        It 'captures standard output' {
            InModuleScope AppxBackup -Parameters @{ Exe = $script:Cmd } {
                param($Exe)
                $result = Invoke-ProcessSafely -FilePath $Exe -ArgumentList @('/c', 'echo AppxBackupMarker') -NoWindow
                $result.StandardOutput | Should -BeLike '*AppxBackupMarker*'
            }
        }

        It 'captures every line of a multi-line stream' {
            # Regression: output was collected through Register-ObjectEvent -Action.
            # PowerShell dispatches those handlers on its own event loop, which is not
            # pumped while the pipeline blocks in WaitForExit, so most output was lost.
            # Against MakeAppx, which emits 36 lines, exactly one line survived and
            # MakeAppx error analysis was left parsing a banner instead of the error.
            InModuleScope AppxBackup -Parameters @{ Exe = $script:Cmd } {
                param($Exe)
                $result = Invoke-ProcessSafely -FilePath $Exe `
                    -ArgumentList @('/c', 'for /L %i in (1,1,200) do @echo line%i') -NoWindow

                $lines = @($result.StandardOutput -split "`r?`n" | Where-Object { $_.Trim() })
                $lines.Count | Should -Be 200
                $lines[0].Trim() | Should -Be 'line1'
                $lines[-1].Trim() | Should -Be 'line200'
            }
        }

        It 'captures output that arrives in a burst before the process exits quickly' {
            InModuleScope AppxBackup -Parameters @{ Exe = $script:Cmd } {
                param($Exe)
                $result = Invoke-ProcessSafely -FilePath $Exe `
                    -ArgumentList @('/c', 'for /L %i in (1,1,50) do @echo burst%i') -NoWindow

                @($result.StandardOutput -split "`r?`n" | Where-Object { $_ -match 'burst' }).Count |
                    Should -Be 50
            }
        }

        It 'captures standard error separately from standard output' {
            InModuleScope AppxBackup -Parameters @{ Exe = $script:Cmd } {
                param($Exe)
                $result = Invoke-ProcessSafely -FilePath $Exe `
                    -ArgumentList @('/c', 'echo ErrMarker 1>&2 & exit 0') -NoWindow
                $result.StandardError | Should -BeLike '*ErrMarker*'
                $result.StandardOutput | Should -Not -BeLike '*ErrMarker*'
            }
        }

        It 'returns a result object carrying the documented contract' {
            InModuleScope AppxBackup -Parameters @{ Exe = $script:Cmd } {
                param($Exe)
                $result = Invoke-ProcessSafely -FilePath $Exe -ArgumentList @('/c', 'exit 0') -NoWindow
                $result.PSObject.TypeNames | Should -Contain 'AppxBackup.ProcessResult'
                foreach ($property in 'ExitCode', 'StandardOutput', 'StandardError', 'Success', 'Duration', 'ToolName') {
                    $result.PSObject.Properties.Name | Should -Contain $property
                }
            }
        }
    }

    Context 'failure handling' {

        It 'throws on a non-zero exit code by default' {
            InModuleScope AppxBackup -Parameters @{ Exe = $script:Cmd } {
                param($Exe)
                { Invoke-ProcessSafely -FilePath $Exe -ArgumentList @('/c', 'exit 3') -NoWindow -ErrorAction Stop } |
                    Should -Throw -ExpectedMessage '*exit code 3*' -Because "$Exe exits 3"
            }
        }

        It 'returns the failing result instead of throwing when -PassThru is used' {
            InModuleScope AppxBackup -Parameters @{ Exe = $script:Cmd } {
                param($Exe)
                $result = Invoke-ProcessSafely -FilePath $Exe -ArgumentList @('/c', 'exit 3') -NoWindow -PassThru
                $result.Success | Should -BeFalse
                $result.ExitCode | Should -Be 3
            }
        }

        It 'throws when the executable does not exist' {
            InModuleScope AppxBackup {
                { Invoke-ProcessSafely -FilePath 'C:\definitely\not\here.exe' -ErrorAction Stop } |
                    Should -Throw -ExpectedMessage '*Executable not found*'
            }
        }

        It 'throws when the working directory does not exist' {
            InModuleScope AppxBackup -Parameters @{ Exe = $script:Cmd } {
                param($Exe)
                { Invoke-ProcessSafely -FilePath $Exe -ArgumentList @('/c', 'exit 0') `
                    -WorkingDirectory 'C:\no\such\dir' -ErrorAction Stop } |
                    Should -Throw -ExpectedMessage '*Working directory not found*' -Because "$Exe cannot start there"
            }
        }
    }

    Context 'argument handling' {

        It 'passes an argument containing spaces without splitting it' {
            # ArgumentList quotes the value, so cmd echoes it with the quotes intact.
            # What matters is that the words arrive contiguously rather than as
            # separate arguments.
            InModuleScope AppxBackup -Parameters @{ Exe = $script:Cmd } {
                param($Exe)
                $result = Invoke-ProcessSafely -FilePath $Exe -ArgumentList @('/c', 'echo', 'one two three') -NoWindow
                $result.StandardOutput | Should -BeLike '*one two three*'
            }
        }
    }
}
