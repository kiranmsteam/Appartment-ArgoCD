# PowerShell script to start local Kubernetes cluster using Minikube
$env:Path = "$PSScriptRoot\bin;" + $env:Path

Write-Host "1. Checking Hyper-V Administrators Group Membership..." -ForegroundColor Cyan

$username = $env:USERNAME
try {
    $isMember = Get-LocalGroupMember -Group "Hyper-V Administrators" | Where-Object { $_.Name -like "*\$username" -or $_.Name -eq $username }
    if (-not $isMember) {
        Write-Host "Adding user '$username' to Hyper-V Administrators group..." -ForegroundColor Yellow
        Add-LocalGroupMember -Group "Hyper-V Administrators" -Member $username -ErrorAction Stop
        Write-Host "User added to Hyper-V Administrators group!" -ForegroundColor Green
    } else {
        Write-Host "User '$username' is already in Hyper-V Administrators group." -ForegroundColor Green
    }
} catch {
    Write-Host "Note: Group check requires Administrator PowerShell window." -ForegroundColor Yellow
}

Write-Host "`n2. Cleaning up any previous Minikube profile..." -ForegroundColor Cyan
& "$PSScriptRoot\bin\minikube.exe" delete

Write-Host "`n3. Starting Minikube Cluster..." -ForegroundColor Green
& "$PSScriptRoot\bin\minikube.exe" start --driver=hyperv

Write-Host "`n4. Checking Cluster Status..." -ForegroundColor Yellow
& "$PSScriptRoot\bin\kubectl.exe" get nodes
