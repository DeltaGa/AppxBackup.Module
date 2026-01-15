<#
.SYNOPSIS
    Imports the AppxBackup module for immediate use.

.DESCRIPTION
    Helper script to import the AppxBackup module from the current directory.
    Run this before using any Backup-AppxPackage commands.

.EXAMPLE
    .\Import-AppxBackup.ps1
    
    Then use: Backup-AppxPackage -PackagePath ...
#>

$ModulePath = $PSScriptRoot

Write-Host "`n=== Importing AppxBackup Module ===" -ForegroundColor Cyan
Write-Host "Location: $ModulePath`n" -ForegroundColor Gray

try {
    # Remove if already loaded
    if (Get-Module AppxBackup) {
        Remove-Module AppxBackup -Force
    }

    # Import module
    Import-Module "$ModulePath\AppxBackup.psd1" -Force -Verbose:$false
    
    Write-Host "[SUCCESS] Module imported successfully!`n" -ForegroundColor Green
    
    # Show available commands
    Write-Host "Available Commands:" -ForegroundColor Cyan
    Get-Command -Module AppxBackup | Format-Table Name, CommandType -AutoSize
    
    Write-Host "`nExample Usage:" -ForegroundColor Yellow
    Write-Host '  $app = Get-AppxPackage -Name "*YourApp*"' -ForegroundColor Gray
    Write-Host '  Backup-AppxPackage -PackagePath $app.InstallLocation -OutputPath "C:\Backups"' -ForegroundColor Gray
    Write-Host ""
}
catch {
    Write-Host "[ERROR] Failed to import module: $_" -ForegroundColor Red
    Write-Host "`nTroubleshooting:" -ForegroundColor Yellow
    Write-Host "  1. Ensure you're running from the AppxBackup.Module directory" -ForegroundColor Gray
    Write-Host "  2. Check that AppxBackup.psd1 exists in current directory" -ForegroundColor Gray
    Write-Host "  3. Run PowerShell as Administrator if needed" -ForegroundColor Gray
}