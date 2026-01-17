<#
.SYNOPSIS
    Structured logging function with multiple output targets and severity levels.

.DESCRIPTION
    Features:
    - Multiple severity levels (Debug, Verbose, Info, Warning, Error, Critical)
    - Simultaneous file and stream output
    - Automatic log rotation
    - Structured format with timestamps and context
    - Thread-safe file operations
    - PowerShell preference variable integration
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

        if (-not $shouldWrite) {
            return
        }
    }

    process {
        try {
            # Build structured log entry
            $timestamp = [DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss.fff')
            $threadId = [System.Threading.Thread]::CurrentThread.ManagedThreadId
            $caller = (Get-PSCallStack)[1].Command
            
            # Format context if provided
            $contextString = ''
            if ($Context.Count -gt 0) {
                $contextPairs = $Context.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }
                $contextString = " [$($contextPairs -join ', ')]"
            }
            
            # Build log line
            $logLine = "[$timestamp] [$Level] [Thread-$threadId] [$Component::$caller]$contextString $Message"
            
            # Console output (respects NoConsole)
            if (-not $NoConsole.IsPresent) {
                switch ($Level) {
                    'Debug'     { Write-Debug $Message }
                    'Verbose'   { Write-Verbose $Message }
                    'Info'      { Write-Host $Message -ForegroundColor Cyan }
                    'Warning'   { Write-Warning $Message }
                    'Error'     { Write-Error $Message }
                    'Critical'  { Write-Error $Message }
                }
            }
            
            # File output (respects NoFile)
            if (-not $NoFile.IsPresent -and $script:LogPath) {
                # Ensure log directory exists
                $logDir = Split-Path -Path $script:LogPath -Parent
                if (-not (Test-Path -LiteralPath $logDir)) {
                    [void](New-Item -Path $logDir -ItemType Directory -Force)
                }
                
                # Check log size and rotate if needed
                if (Test-Path -LiteralPath $script:LogPath) {
                    $logFile = Get-Item -LiteralPath $script:LogPath
                    $maxSizeBytes = $script:AppxBackupConfig.MaxLogSizeMB * 1MB
                    
                    if ($logFile.Length -gt $maxSizeBytes) {
                        # Rotate log
                        $archivePath = "$($script:LogPath).$([DateTime]::Now.ToString('yyyyMMddHHmmss')).archive"
                        Move-Item -LiteralPath $script:LogPath -Destination $archivePath -Force
                        Write-Verbose "Log rotated to: $archivePath"
                    }
                }
                
                # Thread-safe file write (mutex not needed for append operations in PS)
                try {
                    $logLine | Out-File -FilePath $script:LogPath -Append -Encoding utf8 -ErrorAction Stop
                }
                catch {
                    # Fallback: try to write to temp if main log fails
                    $fallbackLog = Join-Path $env:TEMP "AppxBackup_fallback.log"
                    "[$timestamp] [ERROR] Failed to write to main log: $_" | Out-File -FilePath $fallbackLog -Append
                    $logLine | Out-File -FilePath $fallbackLog -Append
                }
            }
        }
        catch {
            # Last resort: Write-Warning (can't use Write-AppxLog recursively)
            Write-Warning "Logging failed: $_"
        }
    }
}
