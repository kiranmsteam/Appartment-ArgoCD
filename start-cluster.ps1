# PowerShell script to start local Kubernetes cluster using Minikube
$env:Path = "$PSScriptRoot\bin;" + $env:Path

# Ensure script is running with Administrator privileges required for Hyper-V
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Hyper-V driver requires Administrator privileges." -ForegroundColor Yellow
    Write-Host "Attempting to launch elevated PowerShell window..." -ForegroundColor Cyan
    try {
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        Write-Host "Launched elevated PowerShell window. Please accept the UAC prompt." -ForegroundColor Green
        exit 0
    } catch {
        Write-Host "ERROR: Failed to elevate. Please right-click PowerShell and select 'Run as Administrator', then execute .\start-cluster.ps1" -ForegroundColor Red
        exit 1
    }
}

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

Write-Host "`n2. Ensuring file permissions and Starting Minikube Cluster..." -ForegroundColor Green
try {
    icacls "$env:USERPROFILE\.minikube" /grant "${env:USERNAME}:(OI)(CI)F" /q 2>$null
} catch {}
& "$PSScriptRoot\bin\minikube.exe" start --driver=hyperv

Write-Host "`n3. Checking Cluster Status..." -ForegroundColor Yellow
& "$PSScriptRoot\bin\kubectl.exe" get nodes

Write-Host "`n4. Ensuring Namespace and Secrets Exist..." -ForegroundColor Cyan
& "$PSScriptRoot\bin\kubectl.exe" create namespace apartment-app --dry-run=client -o yaml | & "$PSScriptRoot\bin\kubectl.exe" apply -f -

# Auto-provision GCP credentials secret if missing
$gcpSecret = & "$PSScriptRoot\bin\kubectl.exe" get secret gcp-credentials -n apartment-app --ignore-not-found
if (-not $gcpSecret) {
    Write-Host "Creating gcp-credentials secret..." -ForegroundColor Yellow
    & "$PSScriptRoot\bin\kubectl.exe" create secret generic gcp-credentials --from-file=credentials.json="$PSScriptRoot\app\hitech-citadel-965501da1a49.json" -n apartment-app
}

# Auto-provision app-env-secret if missing
$appSecret = & "$PSScriptRoot\bin\kubectl.exe" get secret app-env-secret -n apartment-app --ignore-not-found
if (-not $appSecret) {
    Write-Host "Creating app-env-secret..." -ForegroundColor Yellow
    & "$PSScriptRoot\bin\kubectl.exe" create secret generic app-env-secret -n apartment-app `
      --from-literal=EMAIL_APP_PASSWORD="eahw azzc spem bqek" `
      --from-literal=EMAIL_SENDER="hitechcitadel13@gmail.com" `
      --from-literal=GOOGLE_SHEET_BALANCE_ID="1Mi-75vRlh8NMTrtrZoe9ionJL3LLFPyULzVtWRgHohg" `
      --from-literal=GOOGLE_SHEET_CONTRIBUTIONS_ID="1612Fin2ii6cfL_BkeLhIcplcTfZgj42Yy-zXFo7FLCk" `
      --from-literal=GOOGLE_SHEET_CRDR_ID="1Gm8MZhpuaJHG7BGbxsm0uf2eNFSy-hLR4wTai67uRyY" `
      --from-literal=GOOGLE_SHEET_USERS_ID="1TGPCcz0WsM8-JZvtsubAwlyWd9TiQvwlaYkMhAM5azY" `
      --from-literal=SECRET_KEY="8643715616" `
      --from-literal=GOOGLE_SHEETS_CREDENTIALS_FILE="gcp-credentials.json"
}

Write-Host "`n5. Ensuring ArgoCD Application Manifest is Applied..." -ForegroundColor Green
& "$PSScriptRoot\bin\kubectl.exe" apply -f "$PSScriptRoot\argocd\application.yaml"
& "$PSScriptRoot\bin\kubectl.exe" apply -f "$PSScriptRoot\k8s\deployment.yaml" -f "$PSScriptRoot\k8s\service.yaml" -f "$PSScriptRoot\k8s\namespace.yaml"

