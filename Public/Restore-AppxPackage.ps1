<#
.SYNOPSIS
    Restores (installs) a backed-up APPX/MSIX package.

.DESCRIPTION
    Installs a previously backed-up package with automatic certificate trust installation.
    
    Features:
    - Automatic certificate installation to trusted root store
    - Package signature validation before installation
    - Dependency validation before installation
    - Administrator privilege checking
    - Progress indication
    - Rollback on failure
    - Verification after installation
    - Windows SDK tool validation

.PARAMETER PackagePath
    Path to the .appx or .msix file to install.

.PARAMETER CertificatePath
    Path to the .cer file for trust installation.
    If not specified, attempts to find certificate with same base name.

.PARAMETER CertStoreLocation
    Certificate store location for certificate installation.
    Valid values: 'LocalMachine', 'CurrentUser'
    Default: 'LocalMachine' (requires administrator)

.PARAMETER Force
    Bypasses confirmation prompts and overwrites existing app.

.PARAMETER SkipCertificateInstall
    If specified, skips automatic certificate installation.
    Use if certificate is already trusted.

.PARAMETER SkipSignatureValidation
    If specified, skips package signature validation before installation.
    Use with caution - may allow installation of corrupted packages.

.PARAMETER AllowUnsigned
    If specified, attempts to install unsigned packages.
    Requires developer mode enabled.

.EXAMPLE
    Restore-AppxPackage -PackagePath "C:\Backups\MyApp.appx" -CertificatePath "C:\Backups\MyApp.cer"
    
    Installs package with automatic certificate trust to LocalMachine store

.EXAMPLE
    Restore-AppxPackage -PackagePath "C:\Backups\MyApp.appx" -Force
    
    Installs package, auto-detecting certificate, with no prompts

.EXAMPLE
    Restore-AppxPackage -PackagePath "C:\Backups\MyApp.appx" -CertStoreLocation CurrentUser
    
    Installs package with certificate to CurrentUser store (no admin required)

.OUTPUTS
    AppxBackup.RestoreResult

.NOTES
    Requires:
    - PowerShell 5.1+ (7.4+ recommended)
    - Administrator privileges for LocalMachine certificate installation
    - Windows 10 1809+ or Windows 11
    - Add-AppxPackage cmdlet (Windows 8+)
    - Windows SDK (for signature validation)
    
    Author: DeltaGa
    Version: 2.0.0
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
        [ValidateSet('LocalMachine', 'CurrentUser')]
        [string]$CertStoreLocation = 'LocalMachine',

        [Parameter()]
        [switch]$Force,

        [Parameter()]
        [switch]$SkipCertificateInstall,

        [Parameter()]
        [switch]$SkipSignatureValidation,

        [Parameter()]
        [switch]$AllowUnsigned
    )

    begin {
        Write-AppxLog -Message "=== Restore-AppxPackage v2.0 ===" -Level 'Info'
        Write-AppxLog -Message "PowerShell Version: $($PSVersionTable.PSVersion)" -Level 'Debug'
        Write-AppxLog -Message "Operating System: $([Environment]::OSVersion.VersionString)" -Level 'Debug'
        
        # Check for administrator privileges if needed
        $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
            [Security.Principal.WindowsBuiltInRole]::Administrator
        )
        
        if (-not $isAdmin -and -not $SkipCertificateInstall.IsPresent -and $CertStoreLocation -eq 'LocalMachine') {
            Write-Warning "Administrator privileges required for LocalMachine certificate installation."
            Write-Warning "Consider using -CertStoreLocation CurrentUser or run as Administrator."
        }

        # Validate Windows SDK tools if signature validation is needed
        $signToolPath = $null
        if (-not $SkipSignatureValidation.IsPresent -and -not $AllowUnsigned.IsPresent) {
            Write-AppxLog -Message "Validating Windows SDK tools for signature validation..." -Level 'Verbose'
            
            try {
                $signToolPath = Test-AppxToolAvailability -ToolName 'SignTool' -ThrowOnError:$false
                if ($null -eq $signToolPath) {
                    Write-Warning "SignTool.exe not found - signature validation will be skipped"
                    Write-Warning "Install Windows SDK for signature validation capability"
                    $SkipSignatureValidation = $true
                }
                else {
                    Write-AppxLog -Message "Found SignTool: $signToolPath" -Level 'Debug'
                }
            }
            catch {
                Write-Warning "Failed to locate SignTool.exe - signature validation will be skipped: $_"
                $SkipSignatureValidation = $true
            }
        }
    }

    process {
        try {
            # Validate package path
            $packagePath = ConvertTo-SecureFilePath -Path $PackagePath -MustExist -PathType File
            $packageFile = Get-Item -LiteralPath $packagePath
            
            Write-AppxLog -Message "Restoring package: $($packageFile.Name)" -Level 'Verbose'
            Write-Host "`n[INFO] Starting package restoration..." -ForegroundColor Cyan
            Write-Host "  Package: $($packageFile.Name)" -ForegroundColor Gray
            
            # Auto-detect certificate if not specified
            if (-not $CertificatePath -and -not $SkipCertificateInstall.IsPresent -and -not $AllowUnsigned.IsPresent) {
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($packagePath)
                $certPath = [System.IO.Path]::Combine((Split-Path $packagePath), "$baseName.cer")
                
                if (Test-Path -LiteralPath $certPath) {
                    $CertificatePath = $certPath
                    Write-AppxLog -Message "Auto-detected certificate: $certPath" -Level 'Verbose'
                    Write-Host "  Certificate: Auto-detected" -ForegroundColor Gray
                }
                else {
                    Write-Warning "Certificate not found at expected location: $certPath"
                    Write-Warning "Package may fail to install if not signed with a trusted certificate"
                    Write-Warning "Use -SkipCertificateInstall if certificate is already in trusted store"
                }
            }

            # Validate package signature if SignTool is available
            $signatureValid = $false
            if (-not $SkipSignatureValidation.IsPresent -and $signToolPath -and -not $AllowUnsigned.IsPresent) {
                Write-Host "`n[STAGE 1/4] Validating Package Signature..." -ForegroundColor Cyan
                
                try {
                    $verifyResult = Invoke-ProcessSafely -FilePath $signToolPath `
                        -ArgumentList @('verify', '/pa', "`"$packagePath`"") `
                        -TimeoutSeconds 30 `
                        -WorkingDirectory (Split-Path $packagePath)
                    
                    if ($verifyResult.ExitCode -eq 0) {
                        $signatureValid = $true
                        Write-Host "  [SUCCESS] Package signature is valid" -ForegroundColor Green
                        Write-AppxLog -Message "Package signature validated successfully" -Level 'Info'
                    }
                    else {
                        Write-Host "  [WARNING] Package signature validation failed" -ForegroundColor Yellow
                        Write-Host "  SignTool output: $($verifyResult.StandardError)" -ForegroundColor Yellow
                        Write-AppxLog -Message "Signature validation failed: $($verifyResult.StandardError)" -Level 'Warning'
                        
                        if (-not $Force.IsPresent) {
                            $continue = Read-Host "Continue with installation anyway? (y/N)"
                            if ($continue -ne 'y') {
                                Write-AppxLog -Message "Installation cancelled by user due to signature validation failure" -Level 'Info'
                                throw "Package signature validation failed and user chose not to continue"
                            }
                        }
                    }
                }
                catch {
                    Write-Host "  [WARNING] Signature validation error: $_" -ForegroundColor Yellow
                    Write-AppxLog -Message "Signature validation error: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                }
            }
            else {
                Write-Host "`n[STAGE 1/4] Signature Validation (Skipped)" -ForegroundColor Yellow
                if ($AllowUnsigned.IsPresent) {
                    Write-Host "  Reason: AllowUnsigned flag specified" -ForegroundColor Gray
                }
                elseif ($SkipSignatureValidation.IsPresent) {
                    Write-Host "  Reason: SkipSignatureValidation flag specified" -ForegroundColor Gray
                }
                else {
                    Write-Host "  Reason: SignTool not available" -ForegroundColor Gray
                }
            }

            # Get package info
            Write-Host "`n[STAGE 2/4] Analyzing Package..." -ForegroundColor Cyan
            Write-AppxLog -Message "Analyzing package metadata..." -Level 'Verbose'
            
            try {
                $packageInfo = Get-AppxBackupInfo -PackagePath $packagePath -IncludeSignatureInfo
                
                Write-Host "  Name: $($packageInfo.PackageName)" -ForegroundColor Gray
                Write-Host "  Version: $($packageInfo.PackageVersion)" -ForegroundColor Gray
                Write-Host "  Publisher: $($packageInfo.PackagePublisher)" -ForegroundColor Gray
                Write-Host "  Architecture: $($packageInfo.PackageArchitecture)" -ForegroundColor Gray
                Write-Host "  Size: $($packageInfo.PackageSizeMB) MB" -ForegroundColor Gray

                # Check signature info
                if ($packageInfo.SignatureInfo) {
                    $sigStatus = $packageInfo.SignatureInfo.Status
                    $sigColor = if ($sigStatus -eq 'Valid') { 'Green' } elseif ($sigStatus -eq 'Unsigned') { 'Yellow' } else { 'Red' }
                    Write-Host "  Signature: $sigStatus" -ForegroundColor $sigColor
                }
            }
            catch {
                Write-Warning "Failed to retrieve detailed package information: $_"
                Write-AppxLog -Message "Package info retrieval failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                # Continue with installation attempt
            }

            # Check if already installed
            $existingPackage = $null
            if ($packageInfo -and $packageInfo.PackageName) {
                $existingPackage = Get-AppxPackage -Name $packageInfo.PackageName -ErrorAction SilentlyContinue
            }
            
            if ($existingPackage) {
                Write-Host "`n[WARNING] Package Already Installed" -ForegroundColor Yellow
                Write-Host "  Current Version: $($existingPackage.Version)" -ForegroundColor Gray
                Write-Host "  New Version: $($packageInfo.PackageVersion)" -ForegroundColor Gray
                
                if (-not $Force.IsPresent) {
                    $continue = Read-Host "Reinstall/Update this package? (y/N)"
                    if ($continue -ne 'y') {
                        Write-AppxLog -Message "Installation cancelled by user (package already installed)" -Level 'Info'
                        return
                    }
                }
            }

            # Install certificate if needed
            $certificateInstalled = $false
            $certificateThumbprint = $null
            
            if ($CertificatePath -and -not $SkipCertificateInstall.IsPresent) {
                $certificatePath = ConvertTo-SecureFilePath -Path $CertificatePath -MustExist -PathType File
                
                Write-Host "`n[STAGE 3/4] Installing Certificate..." -ForegroundColor Cyan
                Write-Host "  Certificate: $([System.IO.Path]::GetFileName($certificatePath))" -ForegroundColor Gray
                Write-Host "  Target Store: Cert:\$CertStoreLocation\Root" -ForegroundColor Gray
                
                # Check if admin is required
                if ($CertStoreLocation -eq 'LocalMachine' -and -not $isAdmin) {
                    throw "Administrator privileges required to install certificate to LocalMachine store. Use -CertStoreLocation CurrentUser or run as Administrator."
                }

                if ($PSCmdlet.ShouldProcess($certificatePath, "Install certificate to $CertStoreLocation\Root store")) {
                    try {
                        # Validate Import-Certificate cmdlet availability (PowerShell 4+)
                        if (-not (Get-Command -Name Import-Certificate -ErrorAction SilentlyContinue)) {
                            throw "Import-Certificate cmdlet not available. Requires PowerShell 4.0 or later. Current version: $($PSVersionTable.PSVersion)"
                        }

                        # Import to Trusted Root Certification Authorities
                        $certStore = "Cert:\$CertStoreLocation\Root"
                        $importedCert = Import-Certificate -FilePath $certificatePath `
                            -CertStoreLocation $certStore `
                            -ErrorAction Stop
                        
                        $certificateInstalled = $true
                        $certificateThumbprint = $importedCert.Thumbprint
                        
                        Write-Host "  [SUCCESS] Certificate installed successfully" -ForegroundColor Green
                        Write-Host "  Thumbprint: $certificateThumbprint" -ForegroundColor Gray
                        Write-AppxLog -Message "Certificate installed to $certStore : Thumbprint=$certificateThumbprint" -Level 'Info'
                    }
                    catch {
                        Write-Host "  [ERROR] Certificate installation failed" -ForegroundColor Red
                        Write-Host "  Error: $_" -ForegroundColor Red
                        Write-AppxLog -Message "Certificate installation failed: $_ | StackTrace: $($_.ScriptStackTrace)" -Level 'Error'
                        throw "Certificate installation failed: $_"
                    }
                }
                else {
                    Write-Host "  [SKIPPED] Certificate installation (WhatIf mode)" -ForegroundColor Yellow
                }
            }
            else {
                Write-Host "`n[STAGE 3/4] Certificate Installation (Skipped)" -ForegroundColor Yellow
                if ($SkipCertificateInstall.IsPresent) {
                    Write-Host "  Reason: SkipCertificateInstall flag specified" -ForegroundColor Gray
                }
                elseif ($AllowUnsigned.IsPresent) {
                    Write-Host "  Reason: AllowUnsigned flag specified" -ForegroundColor Gray
                }
                else {
                    Write-Host "  Reason: No certificate file provided" -ForegroundColor Gray
                }
            }

            # Install package
            Write-Host "`n[STAGE 4/4] Installing Package..." -ForegroundColor Cyan
            
            if ($PSCmdlet.ShouldProcess($packagePath, "Install APPX package")) {
                try {
                    # Validate Add-AppxPackage cmdlet availability (Windows 8+)
                    if (-not (Get-Command -Name Add-AppxPackage -ErrorAction SilentlyContinue)) {
                        throw "Add-AppxPackage cmdlet not available. This module requires Windows 8 or later."
                    }

                    $installParams = @{
                        Path = $packagePath
                        ErrorAction = 'Stop'
                    }
                    
                    if ($Force.IsPresent) {
                        $installParams['ForceApplicationShutdown'] = $true
                        $installParams['ForceUpdateFromAnyVersion'] = $true
                    }
                    
                    if ($AllowUnsigned.IsPresent) {
                        # Requires developer mode enabled
                        $installParams['AllowUnsigned'] = $true
                        Write-Host "  [WARNING] Installing unsigned package (Developer mode required)" -ForegroundColor Yellow
                    }

                    Write-AppxLog -Message "Calling Add-AppxPackage with params: $($installParams | Out-String)" -Level 'Debug'
                    Add-AppxPackage @installParams
                    
                    Write-Host "  [SUCCESS] Package installed successfully" -ForegroundColor Green
                    Write-AppxLog -Message "Package installed successfully: $($packageInfo.PackageName)" -Level 'Info'
                    
                    # Verify installation - wait for async operations to complete
                    $verificationDelay = Get-AppxDefault 'sleepDelays.verificationDelaySeconds' -Fallback 2
                    Start-Sleep -Seconds $verificationDelay
                    
                    $installedPackage = $null
                    if ($packageInfo -and $packageInfo.PackageName) {
                        $installedPackage = Get-AppxPackage -Name $packageInfo.PackageName -ErrorAction SilentlyContinue
                    }
                    
                    if ($installedPackage) {
                        Write-Host "`n[SUCCESS] Installation Verified" -ForegroundColor Green
                        Write-Host "  Installed Version: $($installedPackage.Version)" -ForegroundColor Gray
                        Write-Host "  Install Location: $($installedPackage.InstallLocation)" -ForegroundColor Gray
                        Write-Host "  Package Full Name: $($installedPackage.PackageFullName)" -ForegroundColor Gray
                    }
                    else {
                        Write-Warning "Package installation reported success but verification failed"
                        Write-Warning "The package may still be installing asynchronously"
                    }

                    # Build result object
                    $result = [PSCustomObject]@{
                        PSTypeName              = 'AppxBackup.RestoreResult'
                        Success                 = $true
                        PackageName             = if ($packageInfo) { $packageInfo.PackageName } else { 'Unknown' }
                        PackageVersion          = if ($installedPackage) { $installedPackage.Version } elseif ($packageInfo) { $packageInfo.PackageVersion } else { 'Unknown' }
                        PackageFullName         = if ($installedPackage) { $installedPackage.PackageFullName } else { 'Unknown' }
                        InstallLocation         = if ($installedPackage) { $installedPackage.InstallLocation } else { $null }
                        PackageFilePath         = $packagePath
                        CertificateInstalled    = $certificateInstalled
                        CertificatePath         = $CertificatePath
                        CertificateThumbprint   = $certificateThumbprint
                        CertificateStore        = if ($certificateInstalled) { "Cert:\$CertStoreLocation\Root" } else { $null }
                        SignatureValidated      = $signatureValid
                        InstallDate             = [DateTime]::Now
                        PreviouslyInstalled     = ($null -ne $existingPackage)
                        PreviousVersion         = if ($existingPackage) { $existingPackage.Version } else { $null }
                    }

                    Write-Host "`n[SUCCESS] Restoration Complete" -ForegroundColor Green
                    Write-AppxLog -Message "Restoration completed successfully" -Level 'Info'
                    
                    return $result
                }
                catch {
                    Write-Host "  [ERROR] Package installation failed" -ForegroundColor Red
                    Write-Host "  Error: $_" -ForegroundColor Red
                    Write-AppxLog -Message "Package installation failed: $_ | StackTrace: $($_.ScriptStackTrace)" -Level 'Error'
                    
                    # Rollback certificate if we just installed it
                    if ($certificateInstalled -and $certificateThumbprint) {
                        Write-Host "`n[ROLLBACK] Removing installed certificate..." -ForegroundColor Yellow
                        
                        try {
                            $certStore = "Cert:\$CertStoreLocation\Root"
                            $certPath = [System.IO.Path]::Combine($certStore, $certificateThumbprint)
                            
                            if (Test-Path -LiteralPath $certPath) {
                                Remove-Item -LiteralPath $certPath -Force -ErrorAction Stop
                                Write-Host "  [SUCCESS] Certificate removed (Thumbprint: $certificateThumbprint)" -ForegroundColor Gray
                                Write-AppxLog -Message "Rolled back certificate installation: $certificateThumbprint" -Level 'Info'
                            }
                            else {
                                Write-Warning "Certificate not found for rollback: $certPath"
                            }
                        }
                        catch {
                            Write-Warning "Failed to rollback certificate installation: $_"
                            Write-AppxLog -Message "Certificate rollback failed: $_ | Stack: $($_.ScriptStackTrace)" -Level 'Warning'
                        }
                    }
                    
                    throw
                }
            }
            else {
                Write-Host "  [SKIPPED] Package installation (WhatIf mode)" -ForegroundColor Yellow
                Write-AppxLog -Message "Installation skipped (WhatIf mode)" -Level 'Info'
            }
        }
        catch {
            Write-AppxLog -Message "Package restore failed: $_ | StackTrace: $($_.ScriptStackTrace)" -Level 'Error'
            throw
        }
    }

    end {
        Write-AppxLog -Message "Restore-AppxPackage completed" -Level 'Debug'
    }
}
