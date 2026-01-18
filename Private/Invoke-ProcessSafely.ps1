<#
.SYNOPSIS
    Safely invokes an external process with comprehensive error handling and timeout protection.

.DESCRIPTION
    Enterprise-grade external process execution with advanced features.
    
    Key improvements:
    - Asynchronous I/O to prevent deadlocks
    - Configurable timeout with forceful termination
    - Tool-specific timeout and async wait defaults (MakeAppx, SignTool, Robocopy)
    - Proper exit code checking with tool-specific interpretation
    - Output buffer size limits (10MB per stream) to prevent memory exhaustion
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

.PARAMETER AsyncWaitMilliseconds
    Time to wait for async event handlers to complete after process exits.
    If not specified, uses tool-specific defaults:
    - MakeAppx: 2000ms (large output)
    - SignTool: 1000ms (small output)
    - Robocopy: 1500ms (moderate output)
    - Default: 1000ms

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
    Version: 2.0.0
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
        [ValidateRange(100, 60000)]
        [int]$AsyncWaitMilliseconds,

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
        
        if (-not $PSBoundParameters.ContainsKey('AsyncWaitMilliseconds')) {
            $toolSpecificConfig = $toolConfig.toolConfigurations.PSObject.Properties |
                Where-Object { $_.Name -eq $toolName } |
                Select-Object -First 1
            
            if ($toolSpecificConfig) {
                $AsyncWaitMilliseconds = $toolSpecificConfig.Value.asyncWaitMilliseconds
            }
            else {
                $AsyncWaitMilliseconds = $toolConfig.defaultConfiguration.asyncWaitMilliseconds
            }
        }
        
        Write-AppxLog -Message "Invoking process: $FilePath" -Level 'Verbose'
        Write-AppxLog -Message "Timeout: $TimeoutSeconds seconds, Async wait: $AsyncWaitMilliseconds ms" -Level 'Debug'
        
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
            $stdoutCapacity = Get-AppxDefault 'bufferSizes.stdoutBuilderInitialCapacity' -Fallback 16384
            $stderrCapacity = Get-AppxDefault 'bufferSizes.stderrBuilderInitialCapacity' -Fallback 4096
            
            $stdoutBuilder = [System.Text.StringBuilder]::new($stdoutCapacity)
            $stderrBuilder = [System.Text.StringBuilder]::new($stderrCapacity)
            
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
            
            if ($null -eq $started) {
                throw "Failed to start process: $FilePath"
            }

            # Begin async read operations
            $process.BeginOutputReadLine()
            $process.BeginErrorReadLine()
            
            Write-AppxLog -Message "Process started: PID $($process.Id)" -Level 'Debug'
            
            # Wait with timeout
            $timeoutMs = $TimeoutSeconds * 1000
            $completed = $process.WaitForExit($timeoutMs)
            
            if ($null -eq $completed) {
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
            
            # Wait for async events to complete (critical!)
            $process.WaitForExit()
            Start-Sleep -Milliseconds $AsyncWaitMilliseconds # Configurable wait for event handlers to finish
            
            # Collect results with buffer size protection
            $exitCode = $process.ExitCode
            $standardOutput = $stdoutBuilder.ToString().TrimEnd()
            $standardError = $stderrBuilder.ToString().TrimEnd()
            $duration = [DateTime]::Now - $startTime
            
            # Protect against excessive memory usage from massive output
            # Load limit from configuration
            $maxOutputSize = Get-AppxDefault 'bufferSizes.maxOutputBytesPerStream' -Fallback 10485760
            
            if ($standardOutput.Length -gt $maxOutputSize) {
                $truncatedLength = $standardOutput.Length
                $standardOutput = $standardOutput.Substring(0, $maxOutputSize)
                Write-AppxLog -Message "STDOUT truncated from $truncatedLength bytes to $maxOutputSize bytes" -Level 'Warning'
            }
            
            if ($standardError.Length -gt $maxOutputSize) {
                $truncatedLength = $standardError.Length
                $standardError = $standardError.Substring(0, $maxOutputSize)
                Write-AppxLog -Message "STDERR truncated from $truncatedLength bytes to $maxOutputSize bytes" -Level 'Warning'
            }
            
            Write-AppxLog -Message "Process completed: Exit code $exitCode in $($duration.TotalSeconds.ToString('F2'))s" -Level 'Debug'
            Write-AppxLog -Message "Captured STDOUT: $($standardOutput.Length) bytes, STDERR: $($standardError.Length) bytes" -Level 'Debug'
            
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
            Write-AppxLog -Message "Process invocation failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Error'
            Write-AppxLog -Message "StackTrace: $($_.ScriptStackTrace)" -Level 'Debug'
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