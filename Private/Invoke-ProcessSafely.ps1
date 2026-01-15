<#
.SYNOPSIS
    Safely invokes an external process with comprehensive error handling and timeout protection.

.DESCRIPTION
    Replaces the catastrophically broken Run-Process from the 2016 implementation.
    
    Key improvements:
    - Asynchronous I/O to prevent deadlocks
    - Configurable timeout with forceful termination
    - Proper exit code checking
    - Separated stdout/stderr streams
    - Structured result object
    - Resource cleanup guarantees
    - Progress indication support

.NOTES
    This is the foundation upon which all external tool invocations are built.
    The 2016 version would hang indefinitely on large output. This version is bulletproof.
#>

function Invoke-ProcessSafely {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$FilePath,

        [Parameter()]
        [AllowEmptyString()]
        [string]$Arguments = '',

        [Parameter()]
        [ValidateRange(1, 86400)]
        [int]$TimeoutSeconds = 3600,

        [Parameter()]
        [string]$WorkingDirectory = $PWD.Path,

        [Parameter()]
        [switch]$NoWindow,

        [Parameter()]
        [hashtable]$EnvironmentVariables = @{},

        [Parameter()]
        [switch]$PassThru
    )

    begin {
        Write-AppxLog -Message "Invoking process: $FilePath $Arguments" -Level 'Verbose'
        
        # Validate executable exists
        if (-not (Test-Path -LiteralPath $FilePath -PathType Leaf)) {
            throw "Executable not found: $FilePath"
        }

        # Validate working directory
        if (-not (Test-Path -LiteralPath $WorkingDirectory -PathType Container)) {
            throw "Working directory not found: $WorkingDirectory"
        }
    }

    process {
        $process = $null
        $stdoutEvent = $null
        $stderrEvent = $null
        
        try {
            # Configure process start info
            $psi = [System.Diagnostics.ProcessStartInfo]::new()
            $psi.FileName = $FilePath
            $psi.Arguments = $Arguments
            $psi.WorkingDirectory = $WorkingDirectory
            $psi.UseShellExecute = $false
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError = $true
            $psi.CreateNoWindow = $NoWindow.IsPresent
            
            # Apply environment variables
            foreach ($key in $EnvironmentVariables.Keys) {
                $psi.EnvironmentVariables[$key] = $EnvironmentVariables[$key]
            }

            # Create process
            $process = [System.Diagnostics.Process]::new()
            $process.StartInfo = $psi
            
            # StringBuilder for async output collection (thread-safe)
            $stdoutBuilder = [System.Text.StringBuilder]::new(16384)
            $stderrBuilder = [System.Text.StringBuilder]::new(4096)
            
            # Register async output handlers to prevent deadlock
            $stdoutEvent = Register-ObjectEvent -InputObject $process `
                -EventName OutputDataReceived `
                -Action {
                    if ($EventArgs.Data) {
                        [void]$Event.MessageData.AppendLine($EventArgs.Data)
                    }
                } `
                -MessageData $stdoutBuilder
                
            $stderrEvent = Register-ObjectEvent -InputObject $process `
                -EventName ErrorDataReceived `
                -Action {
                    if ($EventArgs.Data) {
                        [void]$Event.MessageData.AppendLine($EventArgs.Data)
                    }
                } `
                -MessageData $stderrBuilder
            
            # Start process
            $startTime = [DateTime]::Now
            $started = $process.Start()
            
            if (-not $started) {
                throw "Failed to start process: $FilePath"
            }

            # Begin async read operations
            $process.BeginOutputReadLine()
            $process.BeginErrorReadLine()
            
            Write-AppxLog -Message "Process started: PID $($process.Id)" -Level 'Debug'
            
            # Wait with timeout
            $timeoutMs = $TimeoutSeconds * 1000
            $completed = $process.WaitForExit($timeoutMs)
            
            if (-not $completed) {
                Write-AppxLog -Message "Process exceeded timeout of $TimeoutSeconds seconds" -Level 'Warning'
                
                # Forceful termination
                try {
                    $process.Kill($true) # Kill entire process tree
                    Write-AppxLog -Message "Process terminated forcefully" -Level 'Warning'
                }
                catch {
                    Write-AppxLog -Message "Failed to kill process: $_" -Level 'Error'
                }
                
                throw "Process timed out after $TimeoutSeconds seconds: $FilePath"
            }
            
            # Wait for async events to complete (critical!)
            $process.WaitForExit()
            Start-Sleep -Milliseconds 1000 # Ensure event handlers finish - increased for tools like MakeAppx with verbose output
            
            # Collect results
            $exitCode = $process.ExitCode
            $standardOutput = $stdoutBuilder.ToString().TrimEnd()
            $standardError = $stderrBuilder.ToString().TrimEnd()
            $duration = [DateTime]::Now - $startTime
            
            Write-AppxLog -Message "Process completed: Exit code $exitCode in $($duration.TotalSeconds.ToString('F2'))s" -Level 'Debug'
            Write-AppxLog -Message "Captured STDOUT: $($standardOutput.Length) bytes, STDERR: $($standardError.Length) bytes" -Level 'Debug'
            
            # Build result object
            $result = [PSCustomObject]@{
                PSTypeName = 'AppxBackup.ProcessResult'
                FilePath = $FilePath
                Arguments = $Arguments
                ExitCode = $exitCode
                StandardOutput = $standardOutput
                StandardError = $standardError
                Success = ($exitCode -eq 0)
                Duration = $duration
                ProcessId = $process.Id
                StartTime = $startTime
                ExitTime = $process.ExitTime
            }
            
            # Add formatting
            $result.PSObject.TypeNames.Insert(0, 'AppxBackup.ProcessResult')
            
            # Throw on failure unless PassThru specified
            if (-not $result.Success -and -not $PassThru.IsPresent) {
                # Build comprehensive error message
                $errorMsg = "Process failed with exit code $exitCode"
                
                # Create detailed error log (truncate each to 2000 chars for readability)
                $stderrTrimmed = if ($standardError -and $standardError.Length -gt 0) {
                    if ($standardError.Length -gt 2000) {
                        $standardError.Substring(0, 2000) + "`n... (truncated)"
                    } else {
                        $standardError
                    }
                } else {
                    $null
                }
                
                $stdoutTrimmed = if ($standardOutput -and $standardOutput.Length -gt 0) {
                    if ($standardOutput.Length -gt 2000) {
                        $standardOutput.Substring(0, 2000) + "`n... (truncated)"
                    } else {
                        $standardOutput
                    }
                } else {
                    $null
                }
                
                # Log full details to log file
                Write-AppxLog -Message "=== PROCESS FAILURE DETAILS ===" -Level 'Error'
                Write-AppxLog -Message "Exit Code: $exitCode" -Level 'Error'
                Write-AppxLog -Message "Command: $FilePath $Arguments" -Level 'Error'
                if ($standardError) {
                    Write-AppxLog -Message "STDERR Output:`n$standardError" -Level 'Error'
                }
                if ($standardOutput) {
                    Write-AppxLog -Message "STDOUT Output:`n$standardOutput" -Level 'Error'
                }
                Write-AppxLog -Message "=== END FAILURE DETAILS ===" -Level 'Error'
                
                # Build exception message - ALWAYS include stderr if present (highest priority)
                if ($standardError -and $standardError.Trim().Length -gt 0) {
                    $errorMsg += "`n`n--- Error Output (stderr) ---`n$stderrTrimmed"
                }
                
                # Include stdout (many tools write errors to stdout)
                if ($standardOutput -and $standardOutput.Trim().Length -gt 0) {
                    $errorMsg += "`n`n--- Standard Output (stdout) ---`n$stdoutTrimmed"
                }
                
                # Add helpful hint
                $errorMsg += "`n`nCheck log file for complete output: $env:TEMP\AppxBackup_$(Get-Date -Format 'yyyyMMdd').log"
                
                throw $errorMsg
            }
            
            return $result
        }
        catch {
            Write-AppxLog -Message "Process invocation failed: $_" -Level 'Error'
            throw
        }
        finally {
            # Guaranteed cleanup (even on exceptions)
            if ($stdoutEvent) {
                Unregister-Event -SourceIdentifier $stdoutEvent.Name -ErrorAction SilentlyContinue
            }
            if ($stderrEvent) {
                Unregister-Event -SourceIdentifier $stderrEvent.Name -ErrorAction SilentlyContinue
            }
            if ($process) {
                $process.Dispose()
            }
        }
    }
}