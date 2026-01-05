# 7-Zip Context  - Uninstall Script
# Run as Administrator: powershell -ExecutionPolicy Bypass -File uninstall.ps1

$ErrorActionPreference = "SilentlyContinue"

# Check admin rights
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "[ERROR] Please run this script as Administrator." -ForegroundColor Red
    exit 1
}

Write-Host "========================================" -ForegroundColor Green
Write-Host "7-Zip Context - Uninstall" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

Write-Host "[1/2] Removing package..." -ForegroundColor Cyan
$pkg = Get-AppxPackage -Name "7ZipContextMenu"
if ($pkg) {
    Remove-AppxPackage -Package $pkg.PackageFullName
    Write-Host "  Package removed." -ForegroundColor Green
} else {
    Write-Host "  Package not found."
}

Write-Host "[2/2] Removing certificates..." -ForegroundColor Cyan
$certUser = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -eq "CN=7ZipContextMenu" }
if ($certUser) {
    Remove-Item $certUser.PSPath
    Write-Host "  Certificate removed from CurrentUser."
}

$certRoot = Get-ChildItem Cert:\LocalMachine\Root | Where-Object { $_.Subject -eq "CN=7ZipContextMenu" }
if ($certRoot) {
    Remove-Item $certRoot.PSPath
    Write-Host "  Certificate removed from Root."
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "Uninstall completed!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Please restart Windows Explorer:" -ForegroundColor Yellow
Write-Host "  Stop-Process -Name explorer -Force; Start-Process explorer" -ForegroundColor Cyan
Write-Host ""
