# 7-Zip Context Menu - Unregister Script for MSI
# Called by Windows Installer during uninstallation

$ErrorActionPreference = "SilentlyContinue"

# Remove package
$pkg = Get-AppxPackage -Name "7ZipContext"
if ($pkg) {
    Remove-AppxPackage -Package $pkg.PackageFullName
}

# Remove certificates
$certUser = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -eq "CN=7ZipContext" }
if ($certUser) {
    Remove-Item $certUser.PSPath
}

$certRoot = Get-ChildItem Cert:\LocalMachine\Root | Where-Object { $_.Subject -eq "CN=7ZipContext" }
if ($certRoot) {
    Remove-Item $certRoot.PSPath
}
