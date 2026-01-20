<#
.SYNOPSIS
    Structured logging function with multiple output targets and severity levels.

.DESCRIPTION
    Enterprise-grade logging system with comprehensive safety features.
    
    Features:
    - Multiple severity levels (Debug, Verbose, Info, Warning, Error, Critical)
    - Simultaneous file and stream output
    - Automatic log rotation with configurable size limits
    - Structured format with timestamps and context
    - Thread-safe file operations with mutex protection
    - Concurrent session handling
    - Disk space validation
    - PowerShell preference variable integration
    - Graceful fallback mechanisms
    - UTF-8 encoding with BOM for universal compatibility

.NOTES
    This is the foundational logging system used by all module functions.
    Designed to never fail catastrophically—always has fallback mechanisms.
    
    Author: DeltaGa
    Version: 2.0.1
#>

function Write-AppxLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [AllowEmptyString()]
        [string]$Message,

        [Parameter()]
        [ValidateSet('Debug', 'Verbose', 'Info', 'Warning', 'Error', 'Critical')]
        [string]$Level = 'Info',

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$Component = 'AppxBackup',

        [Parameter()]
        [hashtable]$Context = @{},

        [Parameter()]
        [switch]$NoConsole,

        [Parameter()]
        [switch]$NoFile
    )

    begin {
        # Determine if we should write based on preference variables
        $shouldWrite = switch ($Level) {
            'Debug'     { $DebugPreference -ne 'SilentlyContinue' }
            'Verbose'   { $VerbosePreference -ne 'SilentlyContinue' }
            'Info'      { $true }
            'Warning'   { $WarningPreference -ne 'SilentlyContinue' }
            'Error'     { $ErrorActionPreference -ne 'SilentlyContinue' }
            'Critical'  { $true }
            default     { $true }
        }

        if ($null -eq $shouldWrite) {
            return
        }
    }

    process {
        try {
            # Build structured log entry
            $timestamp = [DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss.fff')
            $threadId = [System.Threading.Thread]::CurrentThread.ManagedThreadId
            
            # Get caller information safely
            $caller = 'Unknown'
            try {
                $callStack = Get-PSCallStack
                if ($callStack -and $callStack.Count -gt 1) {
                    $caller = $callStack[1].Command
                }
            }
            catch {
                # Fallback if callstack unavailable
                $caller = 'Unknown'
            }
            
            # Format context if provided
            $contextString = ''
            if ($null -ne $Context -and $Context.Count -gt 0) {
                try {
                    $contextPairs = $Context.GetEnumerator() | ForEach-Object { 
                        "$($_.Key)=$($_.Value)" 
                    }
                    $contextString = " [$($contextPairs -join ', ')]"
                }
                catch {
                    $contextString = " [Context serialization failed]"
                }
            }
            
            # Build log line
            $logLine = "[$timestamp] [$Level] [Thread-$threadId] [$Component::$caller]$contextString $Message"
            
            # Console output (respects NoConsole)
            if (-not $NoConsole.IsPresent) {
                try {
                    switch ($Level) {
                        'Debug'     { Write-Debug $Message }
                        'Verbose'   { Write-Verbose $Message }
                        'Info'      { 
                            # Silent for Info level - only log to file
                            # Actual user messages use Write-Host directly in functions
                        }
                        'Warning'   { Write-Warning $Message }
                        'Error'     { Write-Error $Message -ErrorAction Continue }
                        'Critical'  { Write-Error $Message -ErrorAction Continue }
                    }
                }
                catch {
                    # If console write fails, continue—file logging is more important
                }
            }
            
            # File output (respects NoFile)
            if (-not $NoFile.IsPresent -and $script:LogPath) {
                # Ensure log directory exists
                $logDir = Split-Path -Path $script:LogPath -Parent
                
                if ($null -eq $logDir) {
                    # No parent directory means script:LogPath is just a filename
                    # This shouldn't happen but handle gracefully
                    return
                }
                
                if (-not (Test-Path -LiteralPath $logDir -ErrorAction SilentlyContinue)) {
                    try {
                        [void](New-Item -Path $logDir -ItemType Directory -Force -ErrorAction Stop)
                    }
                    catch {
                        # Cannot create log directory—fall back to console only
                        return
                    }
                }
                
                # Check available disk space before writing
                try {
                    $drive = [System.IO.Path]::GetPathRoot($script:LogPath)
                    if ($drive) {
                        $driveInfo = [System.IO.DriveInfo]::new($drive)
                        $availableSpaceMB = $driveInfo.AvailableFreeSpace / 1MB
                        
                        # Load minimum required space from configuration
                        $minSpaceMB = Get-AppxDefault 'diskSpaceRequirements.minimalLogSpaceMB' -Fallback 10
                        
                        # Require at least configured minimum free space
                        if ($availableSpaceMB -lt $minSpaceMB) {
                            # Disk space critically low—skip file logging
                            if ($Level -eq 'Critical' -or $Level -eq 'Error') {
                                Write-Warning "Log file skipped: Disk space critically low ($([Math]::Round($availableSpaceMB, 2)) MB available)"
                            }
                            return
                        }
                    }
                }
                catch {
                    # If disk space check fails, attempt to write anyway
                    # (might be network drive or other special case)
                }
                
                # Check log size and rotate if needed
                $shouldRotate = $false
                
                # Load max size and archive count from configuration
                $maxLogSizeMB = Get-AppxDefault 'logConfiguration.maxLogSizeMB' -Fallback 10
                $maxArchives = Get-AppxDefault 'logConfiguration.maxLogArchives' -Fallback 5
                
                # Override with module config if available
                if ($script:AppxBackupConfig -and $script:AppxBackupConfig.MaxLogSizeMB) {
                    $maxLogSizeMB = $script:AppxBackupConfig.MaxLogSizeMB
                }
                
                if (Test-Path -LiteralPath $script:LogPath -ErrorAction SilentlyContinue) {
                    try {
                        $logFile = Get-Item -LiteralPath $script:LogPath -ErrorAction Stop
                        $maxSizeBytes = $maxLogSizeMB * 1MB
                        
                        if ($logFile.Length -gt $maxSizeBytes) {
                            $shouldRotate = $true
                        }
                    }
                    catch {
                        # Cannot check file size—attempt write anyway
                    }
                }
                
                # Perform log rotation if needed
                if ($shouldRotate) {
                    try {
                        # Archive with timestamp
                        $archivePath = "$($script:LogPath).$([DateTime]::Now.ToString('yyyyMMddHHmmss')).archive"
                        
                        # Move existing log to archive
                        Move-Item -LiteralPath $script:LogPath -Destination $archivePath -Force -ErrorAction Stop
                        
                        # Cleanup old archives (keep last N configured)
                        $archivePattern = "$($script:LogPath).*.archive"
                        $archives = Get-ChildItem -Path (Split-Path $script:LogPath) -Filter ([System.IO.Path]::GetFileName($archivePattern)) -ErrorAction SilentlyContinue |
                            Sort-Object -Property CreationTime -Descending
                        
                        if ($archives -and $archives.Count -gt $maxArchives) {
                            $archives | Select-Object -Skip $maxArchives | ForEach-Object {
                                try {
                                    Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
                                }
                                catch {
                                    # Ignore cleanup failures
                                }
                            }
                        }
                    }
                    catch {
                        # Log rotation failed—continue with current log file
                        # It will grow beyond limit but that's acceptable rather than losing logs
                    }
                }
                
                # Thread-safe file write using mutex
                $mutexName = "Global\AppxBackup_LogMutex_$([System.IO.Path]::GetFileName($script:LogPath))"
                $mutex = $null
                $mutexAcquired = $false
                
                try {
                    # Create or open mutex
                    $mutex = [System.Threading.Mutex]::new($false, $mutexName)
                    
                    # Wait for mutex - load timeout from configuration
                    $mutexTimeout = Get-AppxDefault 'logConfiguration.mutexTimeoutMilliseconds' -Fallback 5000
                    $mutexAcquired = $mutex.WaitOne($mutexTimeout)
                    
                    if ($mutexAcquired) {
                        # We have exclusive access—write to file
                        try {
                            # Use UTF-8 with BOM for maximum compatibility
                            $encoding = [System.Text.UTF8Encoding]::new($true) # true = include BOM
                            
                            # Append to file using StreamWriter for better performance
                            $stream = [System.IO.FileStream]::new(
                                $script:LogPath,
                                [System.IO.FileMode]::Append,
                                [System.IO.FileAccess]::Write,
                                [System.IO.FileShare]::Read
                            )
                            
                            $writer = [System.IO.StreamWriter]::new($stream, $encoding)
                            $writer.WriteLine($logLine)
                            $writer.Flush()
                            $writer.Close()
                            $stream.Close()
                        }
                        catch {
                            # File write failed even with mutex—try fallback
                            throw
                        }
                    }
                    else {
                        # Could not acquire mutex in time—fall back to direct write
                        # This may cause occasional corruption but prevents log loss
                        throw "Mutex timeout"
                    }
                }
                catch {
                    # Fallback: Try direct Out-File (may have race conditions but better than nothing)
                    try {
                        $logLine | Out-File -FilePath $script:LogPath -Append -Encoding utf8 -ErrorAction Stop
                    }
                    catch {
                        # Last resort fallback: Write to temp directory with unique name
                        try {
                            $fallbackLog = [System.IO.Path]::Combine(
                                $env:TEMP,
                                "AppxBackup_fallback_$(Get-Date -Format 'yyyyMMdd').log"
                            )
                            
                            "[$timestamp] [WARNING] Failed to write to main log, using fallback" | 
                                Out-File -FilePath $fallbackLog -Append -Encoding utf8 -ErrorAction SilentlyContinue
                            
                            $logLine | Out-File -FilePath $fallbackLog -Append -Encoding utf8 -ErrorAction SilentlyContinue
                        }
                        catch {
                            # Complete logging failure—silently fail rather than crash
                            # Console output (if not suppressed) will still show message
                        }
                    }
                }
                finally {
                    # Always release mutex if acquired
                    if ($mutexAcquired -and $null -ne $mutex) {
                        try {
                            $mutex.ReleaseMutex()
                        }
                        catch {
                            # Mutex release failed—ignore
                        }
                    }
                    
                    # Dispose mutex
                    if ($null -ne $mutex) {
                        try {
                            $mutex.Dispose()
                        }
                        catch {
                            # Disposal failed—ignore
                        }
                    }
                }
            }
        }
        catch {
            # Catastrophic logging failure—last resort warning
            # Use Write-Warning directly to avoid recursion
            try {
                Write-Warning "AppxBackup logging system failure: $_"
            }
            catch {
                # Even Write-Warning failed—complete silence
                # This is the absolute last resort
            }
        }
    }
}
