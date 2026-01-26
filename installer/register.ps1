# 7-Zip Context Menu - Register Script for MSI
# Called by Windows Installer during installation

param(
    [string]$InstallDir = $PSScriptRoot
)

$ErrorActionPreference = "Stop"

# Normalize path
$InstallDir = $InstallDir.TrimEnd('\', '.')
if (-not $InstallDir) {
    $InstallDir = $PSScriptRoot
}

$CertName = "7ZipContext"
$CertFile = Join-Path $InstallDir "$CertName.pfx"
$CertPassword = "7zip"
$ManifestFile = Join-Path $InstallDir "AppxManifest.xml"

# Create certificate if not exists
$cert = Get-ChildItem Cert:\CurrentUser\My -ErrorAction SilentlyContinue | Where-Object { $_.Subject -eq "CN=$CertName" } | Select-Object -First 1

if (-not $cert) {
    $cert = New-SelfSignedCertificate `
        -Type Custom `
        -Subject "CN=$CertName" `
        -KeyUsage DigitalSignature `
        -FriendlyName $CertName `
        -CertStoreLocation "Cert:\CurrentUser\My" `
        -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.3", "2.5.29.19={text}")

    $pwd = ConvertTo-SecureString -String $CertPassword -Force -AsPlainText
    Export-PfxCertificate -Cert $cert -FilePath $CertFile -Password $pwd | Out-Null
    Import-PfxCertificate -FilePath $CertFile -CertStoreLocation Cert:\LocalMachine\Root -Password $pwd | Out-Null
}

# Remove existing package
$pkg = Get-AppxPackage -Name "7ZipContext" -ErrorAction SilentlyContinue
if ($pkg) {
    Remove-AppxPackage -Package $pkg.PackageFullName -ErrorAction SilentlyContinue
}

# Register sparse package
Add-AppxPackage -Register $ManifestFile -ExternalLocation $InstallDir
