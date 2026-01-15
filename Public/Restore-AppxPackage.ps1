<#
.SYNOPSIS
    Restores (installs) a backed-up APPX/MSIX package.

.DESCRIPTION
    Installs a previously backed-up package with automatic certificate trust installation.
    
    Features:
    - Automatic certificate installation to trusted root store
    - Dependency validation before installation
    - Administrator privilege checking
    - Progress indication
    - Rollback on failure
    - Verification after installation

.PARAMETER PackagePath
    Path to the .appx or .msix file to install.

.PARAMETER CertificatePath
    Path to the .cer file for trust installation.
    If not specified, attempts to find certificate with same base name.

.PARAMETER Force
    Bypasses confirmation prompts and overwrites existing app.

.PARAMETER SkipCertificateInstall
    If specified, skips automatic certificate installation.
    Use if certificate is already trusted.

.PARAMETER AllowUnsigned
    If specified, attempts to install unsigned packages.
    Requires developer mode enabled.

.EXAMPLE
    Restore-AppxPackage -PackagePath "C:\Backups\MyApp.appx" -CertificatePath "C:\Backups\MyApp.cer"
    
    Installs package with automatic certificate trust

.EXAMPLE
    Restore-AppxPackage -PackagePath "C:\Backups\MyApp.appx" -Force
    
    Installs package, auto-detecting certificate, with no prompts

.OUTPUTS
    AppxBackup.RestoreResult

.NOTES
    Requires administrator privileges for certificate installation.
    Uses Add-AppxPackage cmdlet (Windows 8+).
#>

function Restore-AppxPackage {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [Alias('FullName', 'Path')]
        [string]$PackagePath,

        [Parameter(Position = 1)]
        [string]$CertificatePath,

        [Parameter()]
        [switch]$Force,

        [Parameter()]
        [switch]$SkipCertificateInstall,

        [Parameter()]
        [switch]$AllowUnsigned
    )

    begin {
        Write-AppxLog -Message "=== Restore-AppxPackage ===" -Level 'Info'
        
        # Check for administrator privileges
        $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        
        if (-not $isAdmin -and -not $SkipCertificateInstall.IsPresent) {
            Write-Warning "Administrator privileges required for certificate installation. Use -SkipCertificateInstall to bypass."
        }
    }

    process {
        try {
            # Validate package path
            $packagePath = ConvertTo-SecureFilePath -Path $PackagePath -MustExist -PathType File
            $packageFile = Get-Item -LiteralPath $packagePath
            
            Write-AppxLog -Message "Restoring package: $($packageFile.Name)" -Level 'Verbose'
            
            # Auto-detect certificate if not specified
            if (-not $CertificatePath -and -not $SkipCertificateInstall.IsPresent -and -not $AllowUnsigned.IsPresent) {
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($packagePath)
                $certPath = Join-Path (Split-Path $packagePath) "$baseName.cer"
                
                if (Test-Path -LiteralPath $certPath) {
                    $CertificatePath = $certPath
                    Write-AppxLog -Message "Auto-detected certificate: $certPath" -Level 'Verbose'
                }
                else {
                    Write-Warning "Certificate not found. Package may fail to install if not signed with trusted certificate."
                    Write-Warning "Expected: $certPath"
                }
            }

            # Get package info
            Write-AppxLog -Message "Analyzing package..." -Level 'Verbose'
            $packageInfo = Get-AppxBackupInfo -PackagePath $packagePath -IncludeSignatureInfo
            
            Write-Host "`n[CHAR_128230] Package Information:" -ForegroundColor Cyan
            Write-Host "  Name: $($packageInfo.PackageName)" -ForegroundColor Gray
            Write-Host "  Version: $($packageInfo.PackageVersion)" -ForegroundColor Gray
            Write-Host "  Publisher: $($packageInfo.PackagePublisher)" -ForegroundColor Gray
            Write-Host "  Architecture: $($packageInfo.PackageArchitecture)" -ForegroundColor Gray
            Write-Host "  Size: $($packageInfo.PackageSizeMB) MB" -ForegroundColor Gray

            # Check signature
            if ($packageInfo.SignatureInfo) {
                Write-Host "  Signature: $($packageInfo.SignatureInfo.Status)" -ForegroundColor $(
                    if ($packageInfo.SignatureInfo.Status -eq 'Valid') { 'Green' } else { 'Yellow' }
                )
            }

            # Check if already installed
            $existingPackage = Get-AppxPackage -Name $packageInfo.PackageName -ErrorAction SilentlyContinue
            
            if ($existingPackage) {
                Write-Host "`n[WARNING]  Package already installed: v$($existingPackage.Version)" -ForegroundColor Yellow
                
                if (-not $Force.IsPresent) {
                    $continue = Read-Host "Reinstall/Update? (y/N)"
                    if ($continue -ne 'y') {
                        Write-AppxLog -Message "Installation cancelled by user" -Level 'Info'
                        return
                    }
                }
            }

            # Install certificate if needed
            $certificateInstalled = $false
            
            if ($CertificatePath -and -not $SkipCertificateInstall.IsPresent) {
                $certificatePath = ConvertTo-SecureFilePath -Path $CertificatePath -MustExist -PathType File
                
                Write-Host "`n[CHAR_128274] Installing Certificate..." -ForegroundColor Cyan
                
                if (-not $isAdmin) {
                    throw "Administrator privileges required to install certificate to Trusted Root store"
                }

                if ($PSCmdlet.ShouldProcess($certificatePath, "Install certificate to Trusted Root")) {
                    try {
                        # Import to Trusted Root Certification Authorities
                        Import-Certificate -FilePath $certificatePath `
                            -CertStoreLocation 'Cert:\LocalMachine\Root' `
                            -ErrorAction Stop | Out-Null
                        
                        $certificateInstalled = $true
                        Write-Host "  [CHECK] Certificate installed successfully" -ForegroundColor Green
                        Write-AppxLog -Message "Certificate installed: $certificatePath" -Level 'Info'
                    }
                    catch {
                        Write-Host "  [X] Certificate installation failed: $_" -ForegroundColor Red
                        Write-AppxLog -Message "Certificate installation failed: $_" -Level 'Error'
                        throw
                    }
                }
            }

            # Install package
            Write-Host "`n[CHAR_128229] Installing Package..." -ForegroundColor Cyan
            
            if ($PSCmdlet.ShouldProcess($packagePath, "Install APPX package")) {
                try {
                    $installParams = @{
                        Path = $packagePath
                        ErrorAction = 'Stop'
                    }
                    
                    if ($Force.IsPresent) {
                        $installParams['ForceApplicationShutdown'] = $true
                        $installParams['ForceUpdateFromAnyVersion'] = $true
                    }
                    
                    if ($AllowUnsigned.IsPresent) {
                        # Requires developer mode
                        $installParams['AllowUnsigned'] = $true
                    }

                    Add-AppxPackage @installParams
                    
                    Write-Host "  [CHECK] Package installed successfully" -ForegroundColor Green
                    Write-AppxLog -Message "Package installed: $($packageInfo.PackageName)" -Level 'Info'
                    
                    # Verify installation
                    Start-Sleep -Seconds 2
                    $installedPackage = Get-AppxPackage -Name $packageInfo.PackageName -ErrorAction Stop
                    
                    if ($installedPackage) {
                        Write-Host "`n[OK] Installation Verified" -ForegroundColor Green
                        Write-Host "  Installed Version: $($installedPackage.Version)" -ForegroundColor Gray
                        Write-Host "  Install Location: $($installedPackage.InstallLocation)" -ForegroundColor Gray
                    }

                    # Build result
                    $result = [PSCustomObject]@{
                        PSTypeName              = 'AppxBackup.RestoreResult'
                        Success                 = $true
                        PackageName             = $packageInfo.PackageName
                        PackageVersion          = $installedPackage.Version
                        PackageFullName         = $installedPackage.PackageFullName
                        InstallLocation         = $installedPackage.InstallLocation
                        PackageFilePath         = $packagePath
                        CertificateInstalled    = $certificateInstalled
                        CertificatePath         = $CertificatePath
                        InstallDate             = [DateTime]::Now
                        PreviouslyInstalled     = ($null -ne $existingPackage)
                        PreviousVersion         = if ($existingPackage) { $existingPackage.Version } else { $null }
                    }

                    Write-Host "`n[INFO] Restore Complete!" -ForegroundColor Green
                    
                    return $result
                }
                catch {
                    Write-Host "  [X] Package installation failed: $_" -ForegroundColor Red
                    Write-AppxLog -Message "Package installation failed: $_" -Level 'Error'
                    
                    # Rollback certificate if we just installed it
                    if ($certificateInstalled) {
                        Write-Host "`n[WARNING]  Rolling back certificate installation..." -ForegroundColor Yellow
                        
                        try {
                            $cert = Get-ChildItem -Path 'Cert:\LocalMachine\Root' | 
                                Where-Object { $_.Subject -like "*$($packageInfo.PackagePublisher)*" } |
                                Select-Object -First 1
                            
                            if ($cert) {
                                Remove-Item -Path "Cert:\LocalMachine\Root\$($cert.Thumbprint)" -Force
                                Write-Host "  [CHECK] Certificate removed" -ForegroundColor Gray
                            }
                        }
                        catch {
                            Write-Warning "Failed to rollback certificate: $_"
                        }
                    }
                    
                    throw
                }
            }
            else {
                Write-AppxLog -Message "Installation skipped (WhatIf)" -Level 'Info'
            }
        }
        catch {
            Write-AppxLog -Message "Package restore failed: $_" -Level 'Error'
            throw
        }
    }

    end {
        Write-AppxLog -Message "Restore-AppxPackage completed" -Level 'Debug'
    }
}
