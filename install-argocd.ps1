# PowerShell helper script to install ArgoCD and retrieve initial login password
$env:Path = "$PSScriptRoot\bin;" + $env:Path

Write-Host "Creating 'argocd' namespace..." -ForegroundColor Cyan
& "$PSScriptRoot\bin\kubectl.exe" create namespace argocd --dry-run=client -o yaml | & "$PSScriptRoot\bin\kubectl.exe" apply -f -

Write-Host "`nDeploying ArgoCD control plane..." -ForegroundColor Green
& "$PSScriptRoot\bin\kubectl.exe" apply --server-side -n argocd -f https://raw.githubusercontent.com/argoproj/argo-cd/stable/manifests/install.yaml

Write-Host "`nWaiting for ArgoCD pods to start..." -ForegroundColor Yellow
Start-Sleep -Seconds 10
& "$PSScriptRoot\bin\kubectl.exe" wait --for=condition=Ready pod -l app.kubernetes.io/name=argocd-server -n argocd --timeout=300s

Write-Host "`nArgoCD Installed Successfully!" -ForegroundColor Green
Write-Host "Fetching initial admin password..." -ForegroundColor Cyan

$encoded = & "$PSScriptRoot\bin\kubectl.exe" -n argocd get secret argocd-initial-admin-secret -o jsonpath="{.data.password}"
$password = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($encoded))

Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host " ArgoCD Login Credentials" -ForegroundColor Green
Write-Host " Username: admin" -ForegroundColor White
Write-Host " Password: $password" -ForegroundColor White
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "`nTo access ArgoCD UI, run:" -ForegroundColor Yellow
Write-Host "  .\bin\kubectl.exe port-forward svc/argocd-server -n argocd 8080:443" -ForegroundColor White
Write-Host "Then open https://localhost:8080 in your browser." -ForegroundColor White
