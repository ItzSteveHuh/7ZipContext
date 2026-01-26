# 7-Zip Context Menu - Unregister Script for MSI
# Called by Windows Installer during uninstallation

$ErrorActionPreference = "SilentlyContinue"

# Setup logging
$LogFile = Join-Path $env:TEMP "7ZipContext_unregister.log"
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] $Message"
    Add-Content -Path $LogFile -Value $logMessage -ErrorAction SilentlyContinue
}

Write-Log "=== 7ZipContext Unregistration Started ==="

# Remove package
$pkg = Get-AppxPackage -Name "7ZipContext"
if ($pkg) {
    Write-Log "Removing package: $($pkg.PackageFullName)"
    Remove-AppxPackage -Package $pkg.PackageFullName
    Write-Log "Package removed"
} else {
    Write-Log "No package found to remove"
}

# Remove certificates
$certUser = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -eq "CN=7ZipContext" }
if ($certUser) {
    Write-Log "Removing user certificate: $($certUser.Thumbprint)"
    Remove-Item $certUser.PSPath
    Write-Log "User certificate removed"
} else {
    Write-Log "No user certificate found"
}

$certRoot = Get-ChildItem Cert:\LocalMachine\Root | Where-Object { $_.Subject -eq "CN=7ZipContext" }
if ($certRoot) {
    Write-Log "Removing root certificate: $($certRoot.Thumbprint)"
    Remove-Item $certRoot.PSPath
    Write-Log "Root certificate removed"
} else {
    Write-Log "No root certificate found"
}

Write-Log "=== 7ZipContext Unregistration Completed ==="
