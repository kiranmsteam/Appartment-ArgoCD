# PowerShell script to stop local Kubernetes cluster using Minikube
$env:Path = "$PSScriptRoot\bin;" + $env:Path

# Ensure script is running with Administrator privileges required for Hyper-V
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Hyper-V driver requires Administrator privileges to manage the VM." -ForegroundColor Yellow
    Write-Host "Attempting to launch elevated PowerShell window..." -ForegroundColor Cyan
    try {
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        Write-Host "Launched elevated PowerShell window. Please accept the UAC prompt." -ForegroundColor Green
        exit 0
    } catch {
        Write-Host "ERROR: Failed to elevate. Please right-click PowerShell and select 'Run as Administrator', then execute .\stop-cluster.ps1" -ForegroundColor Red
        exit 1
    }
}

Write-Host "Stopping Minikube Cluster..." -ForegroundColor Yellow
& "$PSScriptRoot\bin\minikube.exe" stop

if ($LASTEXITCODE -eq 0) {
    Write-Host "`nMinikube cluster has been stopped successfully." -ForegroundColor Green
} else {
    Write-Host "`nMinikube stop command exited with code $LASTEXITCODE." -ForegroundColor Red
}
