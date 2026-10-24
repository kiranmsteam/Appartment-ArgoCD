# Local Kubernetes & ArgoCD GitOps Workflow for Apartment Application

This project sets up a complete **GitOps CI/CD pipeline** locally using **Kubernetes**, **ArgoCD**, **GitHub Actions**, and your **Apartment Application**.

---

## 📁 Repository Structure

```
c:\K8-ArgoCD\
├── .github/
│   └── workflows/
│       └── ci-cd.yml         # GitHub Actions CI pipeline (Build, Push, GitOps tag update)
├── app/                      # Apartment Application Source Code & Dockerfile
├── bin/                      # Downloaded CLI binaries (minikube.exe, kubectl.exe, helm.exe)
├── k8s/                      # Kubernetes Manifests (Declarative Infrastructure)
│   ├── namespace.yaml
│   ├── deployment.yaml
│   ├── service.yaml
│   └── kustomization.yaml
├── argocd/                   # ArgoCD Configuration
│   └── application.yaml      # ArgoCD Custom Resource defining app & Git sync
├── start-cluster.ps1         # Script to launch local Minikube cluster
└── install-argocd.ps1       # Script to install ArgoCD control plane & fetch credentials
```

---

## 🚀 Quickstart Guide

### 1. Start the Local Kubernetes Cluster
Open **PowerShell as Administrator** and run:
```powershell
.\start-cluster.ps1
```
> *Note: Minikube uses Hyper-V by default on Windows Pro. Alternatively, start with Docker Desktop using `.\bin\minikube.exe start --driver=docker`.*

### 2. Install ArgoCD on the Cluster
Run the installation script:
```powershell
.\install-argocd.ps1
```
This script will:
- Create the `argocd` namespace.
- Deploy ArgoCD controller, server, repo-server, and redis.
- Print out the initial `admin` password.

### 3. Expose ArgoCD Web UI
Run port-forwarding to access ArgoCD locally:
```powershell
.\bin\kubectl.exe port-forward svc/argocd-server -n argocd 8080:443
```
Open **https://localhost:8080** in your browser and log in with username `admin` and the password printed from step 2.

### 4. Push Repository to GitHub
Initialize git and push to your GitHub repository:
```bash
git init
git add .
git commit -m "feat: initial commit for K8s GitOps setup"
git remote add origin https://github.com/YOUR_USERNAME/K8-ArgoCD.git
git branch -M main
git push -u origin main
```

### 5. Connect ArgoCD to GitHub Repository
Update `argocd/application.yaml` with your repository URL, then apply it:
```powershell
.\bin\kubectl.exe apply -f argocd/application.yaml
```

ArgoCD will automatically watch the `k8s/` directory in your GitHub repo and deploy the **Apartment Application** into your Kubernetes cluster!

---

## 🔄 How the GitOps Pipeline Works

1. **Code Commit**: You push code changes in `app/` to GitHub (`main` branch).
2. **GitHub Actions (CI)**:
   - Builds the Docker image.
   - Pushes image to GitHub Container Registry (`ghcr.io`).
   - Updates `newTag` in `k8s/kustomization.yaml` and commits back to GitHub.
3. **ArgoCD (CD / GitOps)**:
   - ArgoCD detects the change in `k8s/kustomization.yaml`.
   - Performs automated sync and rolls out the new app version in the cluster zero-downtime!

---

## 🌐 Accessing Your Apartment Application
Once deployed by ArgoCD, access the app via Minikube service or port forward:
```powershell
.\bin\kubectl.exe port-forward svc/apartment-app-service -n apartment-app 8081:8080
```
Open **http://localhost:8081** in your browser.

---

## 📊 Monthly Balance Sheet Generator (`generate_monthly_balance_sheet.py`)

Whenever a new month ends and transactions are recorded in the `Cr Dr` Google Sheet, you can automatically generate and format the new monthly balance sheet tab in Google Sheets using the automated generator.

### What it does:
1. **Reads Previous Month Balances**: Automatically looks up the prior month's tab in the `Current Balance Sheet` and carries forward the bank opening balance and cash in hand.
2. **Aggregates Current Month Transactions**: Reads maintenance contributions, bank interest, cheque/online payments (consolidating multi-meter BESCOM entries), and cash expenses from the corresponding tab in the `Cr Dr` Google Sheet.
3. **Builds & Styles New Tab**: Creates and formats the monthly sheet (e.g. `Sep 2026`) with all formulas (`Total of A`, `Total of B`, `Total of C`, `Closing Balance`), styling, colors, and borders.
4. **Instant Web App Sync**: The web app detects the new tab immediately under the relevant Financial Year without requiring any app restart.

---

### How to Run:

#### Option 1: Inside Kubernetes (Recommended - No local Python or dependency setup needed)
Run this command directly from your `C:\K8-ArgoCD` terminal:
```powershell
.\bin\kubectl.exe exec -n apartment-app deploy/apartment-app-deployment -- python /app/generate_monthly_balance_sheet.py --month <MonthName> --year <Year>
```

**Examples:**
```powershell
# For September 2026:
.\bin\kubectl.exe exec -n apartment-app deploy/apartment-app-deployment -- python /app/generate_monthly_balance_sheet.py --month September --year 2026

# For October 2026:
.\bin\kubectl.exe exec -n apartment-app deploy/apartment-app-deployment -- python /app/generate_monthly_balance_sheet.py --month October --year 2026
```

---

#### Option 2: Running Locally via Python
If running directly on your Windows host machine:
```powershell
cd C:\K8-ArgoCD\app
python generate_monthly_balance_sheet.py --month September --year 2026
```
*(Requires `Apartment.env` loaded and dependencies installed via `pip install gspread google-auth`)*

