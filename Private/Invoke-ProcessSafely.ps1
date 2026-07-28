<#
.SYNOPSIS
    Safely invokes an external process with comprehensive error handling and timeout protection.

.DESCRIPTION
    Enterprise-grade external process execution with advanced features.
    
    Key improvements:
    - Asynchronous I/O to prevent deadlocks
    - Configurable timeout with forceful termination
    - Tool-specific timeout defaults (MakeAppx, SignTool, Robocopy)
    - Proper exit code checking with tool-specific interpretation
    - Separated stdout/stderr streams
    - Structured result object
    - Resource cleanup guarantees
    - Progress indication support

.PARAMETER FilePath
    Path to the executable to invoke.

.PARAMETER ArgumentList
    Array of arguments to pass to the executable.
    Proper handling of quotes and special characters.

.PARAMETER TimeoutSeconds
    Maximum execution time in seconds.
    If not specified, uses tool-specific defaults:
    - MakeAppx: 600 seconds (10 minutes)
    - SignTool: 60 seconds (1 minute)
    - Robocopy: 300 seconds (5 minutes)
    - Default: 3600 seconds (1 hour)

.PARAMETER WorkingDirectory
    Working directory for the process. Defaults to current directory.

.PARAMETER NoWindow
    If specified, creates the process without a visible window.

.PARAMETER EnvironmentVariables
    Hashtable of environment variables to set for the process.

.PARAMETER PassThru
    If specified, returns the result object even on failure (does not throw).

.OUTPUTS
    AppxBackup.ProcessResult - Contains ExitCode, StandardOutput, StandardError, Success, Duration, etc.

.NOTES
    This is the foundation upon which all external tool invocations are built.
    Tool-specific exit code interpretation ensures robust error handling.
    
    Author: DeltaGa
    Version: 2.0.2
#>

function Invoke-ProcessSafely {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$FilePath,

        [Parameter()]
        [string[]]$ArgumentList = @(),

        [Parameter()]
        [ValidateRange(1, 86400)]
        [int]$TimeoutSeconds,

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
        # Load tool-specific configuration from external database
        try {
            $toolConfig = Get-AppxConfiguration -ConfigName 'ToolConfiguration'
            Write-AppxLog -Message "Loaded tool configuration database" -Level 'Debug'
        }
        catch {
            Write-AppxLog -Message "Failed to load tool configuration, using hardcoded defaults: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
            
            # Fallback to minimal hardcoded configuration
            $toolConfig = [PSCustomObject]@{
                defaultConfiguration = [PSCustomObject]@{
                    timeoutSeconds = 3600
                    asyncWaitMilliseconds = 1000
                }
                toolConfigurations = [PSCustomObject]@{
                    MakeAppx = [PSCustomObject]@{
                        timeoutSeconds = 600
                        asyncWaitMilliseconds = 2000
                    }
                    SignTool = [PSCustomObject]@{
                        timeoutSeconds = 60
                        asyncWaitMilliseconds = 1000
                    }
                    Robocopy = [PSCustomObject]@{
                        timeoutSeconds = 300
                        asyncWaitMilliseconds = 1500
                    }
                }
                exitCodeInterpretation = [PSCustomObject]@{
                    Robocopy = [PSCustomObject]@{
                        successCodes = @(0,1,2,3,4,5,6,7)
                        errorCodesStart = 8
                    }
                }
            }
        }
        
        # Detect tool from FilePath
        $toolName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
        
        # Apply tool-specific defaults if not explicitly specified
        if (-not $PSBoundParameters.ContainsKey('TimeoutSeconds')) {
            $toolSpecificConfig = $toolConfig.toolConfigurations.PSObject.Properties |
                Where-Object { $_.Name -eq $toolName } |
                Select-Object -First 1
            
            if ($toolSpecificConfig) {
                $TimeoutSeconds = $toolSpecificConfig.Value.timeoutSeconds
            }
            else {
                $TimeoutSeconds = $toolConfig.defaultConfiguration.timeoutSeconds
            }
        }
        
        Write-AppxLog -Message "Invoking process: $FilePath" -Level 'Verbose'
        Write-AppxLog -Message "Timeout: $TimeoutSeconds seconds" -Level 'Debug'
        
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

        try {
            # Configure process start info
            $psi = [System.Diagnostics.ProcessStartInfo]::new()
            $psi.FileName = $FilePath
            
            # Convert ArgumentList array to proper arguments
            if ($null -ne $ArgumentList -and $ArgumentList.Count -gt 0) {
                # Join arguments - proper handling of quotes and spaces
                foreach ($arg in @($ArgumentList)) {
                    if ($null -ne $arg) {
                        $psi.ArgumentList.Add($arg.ToString())
                    }
                }
            }
            
            $psi.WorkingDirectory = $WorkingDirectory
            $psi.UseShellExecute = $false
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError = $true
            $psi.CreateNoWindow = $NoWindow.IsPresent
            
            # Apply environment variables
            foreach ($key in @($EnvironmentVariables.Keys)) {
                $psi.EnvironmentVariables[$key] = $EnvironmentVariables[$key]
            }

            # Create process
            $process = [System.Diagnostics.Process]::new()
            $process.StartInfo = $psi
            
            # StringBuilder for async output collection (thread-safe)
            # Load initial capacities from configuration for optimal memory allocation
            # Start process
            $startTime = [DateTime]::Now
            $started = $process.Start()

            # Process.Start() returns [bool]; it is never $null, so this must be a
            # falsiness test. Comparing against $null here silently disabled the check.
            if (-not $started) {
                throw "Failed to start process: $FilePath"
            }

            # Read both streams with the .NET async readers rather than
            # Register-ObjectEvent. PowerShell dispatches -Action handlers on its own
            # event loop, which only pumps while the pipeline is idle; during
            # WaitForExit the queue is not drained and output was lost. Measured
            # against MakeAppx, which emits 36 lines: exactly one line survived.
            # Get-AppxMakeAppxErrorAnalysis therefore parsed a banner instead of the
            # actual error, so packaging failures produced no usable diagnosis.
            #
            # Starting both reads before waiting is also what keeps this deadlock-free:
            # neither pipe can fill up while we block on the other.
            $stdoutTask = $process.StandardOutput.ReadToEndAsync()
            $stderrTask = $process.StandardError.ReadToEndAsync()

            Write-AppxLog -Message "Process started: PID $($process.Id)" -Level 'Debug'
            
            # Wait with timeout
            $timeoutMs = $TimeoutSeconds * 1000
            $completed = $process.WaitForExit($timeoutMs)

            # WaitForExit([int]) returns [bool] - $false on timeout, never $null.
            # Comparing against $null made this branch unreachable, so a hung tool was
            # never killed and execution fell through to the unbounded WaitForExit()
            # below, blocking the caller forever. This must be a falsiness test.
            if (-not $completed) {
                Write-AppxLog -Message "Process exceeded timeout of $TimeoutSeconds seconds" -Level 'Warning'
                
                # Forceful termination
                try {
                    $process.Kill($true) # Kill entire process tree
                    Write-AppxLog -Message "Process terminated forcefully" -Level 'Warning'
                }
                catch {
                    Write-AppxLog -Message "Failed to kill process: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'
                    Write-AppxLog -Message "StackTrace: $($_.ScriptStackTrace)" -Level 'Debug'
                }
                
                throw "Process timed out after $TimeoutSeconds seconds: $FilePath"
            }
            
            # Only reached when the process exited within the timeout. The parameterless
            # overload is what flushes the redirected stdout/stderr readers, so it must
            # run after the timed wait succeeds - and never on the timeout path.
            $process.WaitForExit()

            $exitCode = $process.ExitCode
            $duration = [DateTime]::Now - $startTime

            # The reads complete once the child closes its stream handles, which has
            # already happened by the time WaitForExit() returns. GetAwaiter() is used
            # rather than .Result so a fault surfaces as the original exception instead
            # of an AggregateException.
            try {
                $standardOutput = $stdoutTask.GetAwaiter().GetResult()
            }
            catch {
                Write-AppxLog -Message "Failed to read standard output: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                $standardOutput = ''
            }

            try {
                $standardError = $stderrTask.GetAwaiter().GetResult()
            }
            catch {
                Write-AppxLog -Message "Failed to read standard error: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                $standardError = ''
            }

            if ($null -eq $standardOutput) { $standardOutput = '' }
            if ($null -eq $standardError) { $standardError = '' }

            Write-AppxLog -Message "Process completed: Exit code $exitCode in $($duration.TotalSeconds.ToString('F2'))s" -Level 'Debug'
            Write-AppxLog -Message "Captured STDOUT: $($standardOutput.Length) chars, STDERR: $($standardError.Length) chars" -Level 'Debug'
            
            # Tool-specific exit code interpretation from configuration
            $isSuccess = $false
            
            # Check if tool has custom exit code interpretation
            $exitCodeConfig = $toolConfig.exitCodeInterpretation.PSObject.Properties |
                Where-Object { $_.Name -eq $toolName } |
                Select-Object -First 1
            
            if ($exitCodeConfig) {
                # Tool has custom interpretation
                $successCodes = $exitCodeConfig.Value.successCodes
                
                if ($successCodes -contains $exitCode) {
                    $isSuccess = $true
                }
                elseif ($exitCodeConfig.Value.errorCodesStart) {
                    # Range-based check (e.g., Robocopy: 0-7 success, 8+ error)
                    $isSuccess = ($exitCode -lt $exitCodeConfig.Value.errorCodesStart)
                }
                
                if (-not $isSuccess -and $exitCode -ge $exitCodeConfig.Value.errorCodesStart) {
                    Write-AppxLog -Message "$toolName failed with exit code $exitCode ($($exitCodeConfig.Value.errorCodesStart)+ indicates error)" -Level 'Warning'
                }
                elseif ($isSuccess -and $exitCode -gt 0) {
                    Write-AppxLog -Message "$toolName completed with warnings (exit code $exitCode)" -Level 'Debug'
                }
            }
            else {
                # Standard interpretation: 0 = success, non-zero = error
                $isSuccess = ($exitCode -eq 0)
            }
            
            # Build result object
            $result = [PSCustomObject]@{
                PSTypeName = 'AppxBackup.ProcessResult'
                FilePath = $FilePath
                Arguments = if ($ArgumentList) { $ArgumentList -join ' ' } else { '' }
                ArgumentList = $ArgumentList
                ExitCode = $exitCode
                StandardOutput = $standardOutput
                StandardError = $standardError
                Success = $isSuccess
                Duration = $duration
                ProcessId = $process.Id
                StartTime = $startTime
                ExitTime = $process.ExitTime
                ToolName = $toolName
            }
            
            # Add formatting
            $result.PSObject.TypeNames.Insert(0, 'AppxBackup.ProcessResult')
            
            # Throw on failure unless PassThru specified
            if (-not $result.Success -and -not $PassThru.IsPresent) {
                # Build comprehensive error message
                $errorMsg = "Process failed with exit code $exitCode"
                
                # Load truncation length from configuration
                $truncateLength = Get-AppxDefault 'bufferSizes.errorMessageTruncateLength' -Fallback 2000
                
                # Create detailed error log (truncate each to configured length for readability)
                $stderrTrimmed = if ($standardError -and $standardError.Length -gt 0) {
                    if ($standardError.Length -gt $truncateLength) {
                        $standardError.Substring(0, $truncateLength) + "`n... (truncated)"
                    } else {
                        $standardError
                    }
                } else {
                    $null
                }
                
                $stdoutTrimmed = if ($standardOutput -and $standardOutput.Length -gt 0) {
                    if ($standardOutput.Length -gt $truncateLength) {
                        $standardOutput.Substring(0, $truncateLength) + "`n... (truncated)"
                    } else {
                        $standardOutput
                    }
                } else {
                    $null
                }
                
                # Log full details to log file
                Write-AppxLog -Message "=== PROCESS FAILURE DETAILS ===" -Level 'Error'
                Write-AppxLog -Message "Exit Code: $exitCode" -Level 'Error'
                Write-AppxLog -Message "Command: $FilePath $($ArgumentList -join ' ')" -Level 'Error'
                
                # CRITICAL: Output full error details to console for debugging
                Write-Host "`n============================================" -ForegroundColor Red
                Write-Host "=== PROCESS FAILED ===" -ForegroundColor Red
                Write-Host "============================================" -ForegroundColor Red
                Write-Host "Exit Code: $exitCode" -ForegroundColor Red
                Write-Host "Tool: $toolName" -ForegroundColor Yellow
                Write-Host "Command: $FilePath" -ForegroundColor Yellow
                Write-Host "Arguments: $($ArgumentList -join ' ')" -ForegroundColor Yellow
                Write-Host "Duration: $($duration.TotalSeconds.ToString('F2'))s" -ForegroundColor Yellow
                
                if ($standardError -and $standardError.Trim().Length -gt 0) {
                    Write-AppxLog -Message "STDERR Output ($($standardError.Length) chars):`n$standardError" -Level 'Error'
                    Write-Host "`n--- STDERR ($($standardError.Length) chars) ---" -ForegroundColor Red
                    Write-Host $standardError -ForegroundColor Gray
                }
                else {
                    Write-Host "`n--- STDERR: (empty) ---" -ForegroundColor Red
                }
                
                if ($standardOutput -and $standardOutput.Trim().Length -gt 0) {
                    Write-AppxLog -Message "STDOUT Output ($($standardOutput.Length) chars):`n$standardOutput" -Level 'Error'
                    Write-Host "`n--- STDOUT ($($standardOutput.Length) chars) ---" -ForegroundColor Yellow
                    Write-Host $standardOutput -ForegroundColor Gray
                }
                else {
                    Write-Host "`n--- STDOUT: (empty) ---" -ForegroundColor Yellow
                }
                
                Write-Host "============================================" -ForegroundColor Red
                Write-Host "=== END FAILURE DETAILS ===" -ForegroundColor Red
                Write-Host "============================================`n" -ForegroundColor Red
                
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
            Write-AppxLog -Message "Process invocation failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'
            Write-AppxLog -Message "StackTrace: $($_.ScriptStackTrace)" -Level 'Debug'
            throw
        }
        finally {
            # Guaranteed cleanup (even on exceptions)
            if ($process) {
                $process.Dispose()
            }
        }
    }
}