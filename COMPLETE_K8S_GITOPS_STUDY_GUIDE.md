# Complete Local Kubernetes, ArgoCD, Helm & GitOps Study Guide

> **Project**: Apartment Application  
> **Target Environment**: Local Windows 11 (Minikube + Hyper-V)  
> **GitOps & CI/CD**: ArgoCD, GitHub Actions, Helm Charts, GHCR / Docker Registries  
> **Repositories**:  
> - **Repo 1 (App Source Code)**: `https://github.com/TechFamily-21/Apartment`  
> - **Repo 2 (GitOps Infrastructure)**: `https://github.com/kiranmsteam/Appartment-ArgoCD`  

---

## 📋 Table of Contents
1. [Architecture Overview](#1-architecture-overview)
2. [Local Environment & Prerequisites Setup](#2-local-environment--prerequisites-setup)
3. [Kubernetes Infrastructure & ArgoCD Setup](#3-kubernetes-infrastructure--argocd-setup)
4. [Helm Chart Architecture & Customization](#4-helm-chart-architecture--customization)
5. [CI/CD Pipelines & Cross-Repository Workflows](#5-cicd-pipelines--cross-repository-workflows)
6. [Google Sheets Integration & Secret Management](#6-google-sheets-integration--secret-management)
7. [Essential Command Reference Cheat Sheet](#7-essential-command-reference-cheat-sheet)
8. [Complete Q&A Compendium](#8-complete-qa-compendium)

---

## 1. Architecture Overview

### Dual-Repository GitOps Pattern

```
 +--------------------------------------------------------------------------------------+
 | REPOSITORY 1: APPLICATION CODE                                                        |
 | (TechFamily-21/Apartment)                                                            |
 |                                                                                      |
 | Contains: app.py, templates/, Dockerfile, requirements.txt                           |
 | Role: Developers write application logic here.                                        |
 +------------------------------------------+-------------------------------------------+
                                            |
                                            | 1. Push Code to 'main'
                                            v
 +--------------------------------------------------------------------------------------+
 | CI PIPELINE (GitHub Actions in Repo 1)                                                |
 | - Runs linting & syntax tests                                                        |
 | - Builds Docker Image: ghcr.io/kiranmsteam/apartment-app:sha-<hash>                  |
 | - Pushes Docker Image to GitHub Container Registry (GHCR)                            |
 | - 🚀 CROSS-REPO COMMIT: Updates image tag in Repo 2!                                  |
 +------------------------------------------+-------------------------------------------+
                                            |
                                            | 2. Updates Image Tag in values.yaml
                                            v
 +--------------------------------------------------------------------------------------+
 | REPOSITORY 2: GITOPS / INFRASTRUCTURE CONFIG                                         |
 | (kiranmsteam/Appartment-ArgoCD)                                                     |
 |                                                                                      |
 | Contains: k8s/deployment.yaml, helm-chart/values.yaml, argocd/application.yaml      |
 | Role: Single Source of Truth for desired Kubernetes cluster state.                   |
 +------------------------------------------+-------------------------------------------+
                                            |
                                            | 3. ArgoCD Polls / Receives Webhook
                                            v
 +--------------------------------------------------------------------------------------+
 | LOCAL KUBERNETES CLUSTER (ArgoCD Operator in Minikube)                               |
 | - Detects diff between Git and Cluster state                                         |
 | - Pulls new container image (sha-<hash>) from GHCR                                   |
 | - Performs zero-downtime rolling update of pods                                      |
 +--------------------------------------------------------------------------------------+
```

---

## 2. Local Environment & Prerequisites Setup

### System Configuration (Windows 11 + Hyper-V)

1. **Enable Hyper-V PowerShell Module** (Run as Administrator once):
   ```powershell
   Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V-Tools-All -All
   ```
2. **Add User to Hyper-V Administrators Local Group**:
   ```powershell
   Add-LocalGroupMember -Group "Hyper-V Administrators" -Member $env:USERNAME
   ```
3. **Local CLI Binaries Directory (`C:\K8-ArgoCD\bin`)**:
   - `minikube.exe` (v1.39.0)
   - `kubectl.exe` (v1.37.0)
   - `helm.exe` (v3.15.4)

### Cluster Startup Script (`start-cluster.ps1`)
```powershell
# C:\K8-ArgoCD\start-cluster.ps1
$env:Path = "$PSScriptRoot\bin;" + $env:Path

Write-Host "1. Checking Hyper-V Group Membership..." -ForegroundColor Cyan
$username = $env:USERNAME
try {
    $isMember = Get-LocalGroupMember -Group "Hyper-V Administrators" | Where-Object { $_.Name -like "*\$username" -or $_.Name -eq $username }
    if (-not $isMember) {
        Add-LocalGroupMember -Group "Hyper-V Administrators" -Member $username -ErrorAction Stop
    }
} catch {
    Write-Host "Note: Group check requires Administrator PowerShell window." -ForegroundColor Yellow
}

Write-Host "`n2. Starting Minikube Cluster..." -ForegroundColor Green
& "$PSScriptRoot\bin\minikube.exe" start --driver=hyperv

Write-Host "`n3. Checking Cluster Status..." -ForegroundColor Yellow
& "$PSScriptRoot\bin\kubectl.exe" get nodes
```

---

## 3. Kubernetes Infrastructure & ArgoCD Setup

### A. Namespace (`k8s/namespace.yaml`)
```yaml
apiVersion: v1
kind: Namespace
metadata:
  name: apartment-app
```

### B. Deployment (`k8s/deployment.yaml`)
```yaml
apiVersion: apps/v1
kind: Deployment
metadata:
  name: apartment-app-deployment
  namespace: apartment-app
  labels:
    app: apartment-app
spec:
  replicas: 2
  selector:
    matchLabels:
      app: apartment-app
  template:
    metadata:
      labels:
        app: apartment-app
    spec:
      containers:
      - name: apartment-app
        image: ghcr.io/kiranmsteam/apartment-app:latest
        imagePullPolicy: IfNotPresent
        ports:
        - containerPort: 8080
        resources:
          limits:
            cpu: "500m"
            memory: "512Mi"
          requests:
            cpu: "100m"
            memory: "256Mi"
        envFrom:
        - secretRef:
            name: app-env-secret
        volumeMounts:
        - name: gcp-creds
          mountPath: /app/gcp-credentials.json
          subPath: credentials.json
          readOnly: true
      volumes:
      - name: gcp-creds
        secret:
          secretName: gcp-credentials
          optional: true
```

### C. Service (`k8s/service.yaml`)
```yaml
apiVersion: v1
kind: Service
metadata:
  name: apartment-app-service
  namespace: apartment-app
spec:
  type: NodePort
  ports:
  - port: 8080
    targetPort: 8080
    nodePort: 30080
  selector:
    app: apartment-app
```

### D. ArgoCD Application Manifest (`argocd/application.yaml`)
```yaml
apiVersion: argoproj.io/v1alpha1
kind: Application
metadata:
  name: apartment-app
  namespace: argocd
  finalizers:
    - resources-finalizer.argocd.argoproj.io
spec:
  project: default
  source:
    repoURL: 'https://github.com/kiranmsteam/Appartment-ArgoCD.git'
    targetRevision: HEAD
    path: k8s
  destination:
    server: 'https://kubernetes.default.svc'
    namespace: apartment-app
  syncPolicy:
    automated:
      prune: true
      selfHeal: true
    syncOptions:
      - CreateNamespace=true
```

---

## 4. Helm Chart Architecture & Customization

Directory location: `C:\K8-ArgoCD\helm-chart\apartment-app`

### `values.yaml`
```yaml
replicaCount: 2

image:
  repository: ghcr.io/kiranmsteam/apartment-app
  pullPolicy: IfNotPresent
  tag: "latest"

service:
  type: NodePort
  port: 8080
  nodePort: 30080

resources:
  limits:
    cpu: 500m
    memory: 512Mi
  requests:
    cpu: 100m
    memory: 256Mi

env:
  port: "8080"
  envName: "development"
  googleSheetsCredentialsFile: "gcp-credentials.json"
  googleSheetBalanceId: "YOUR_BALANCE_SHEET_ID"
  googleSheetContributionsId: "YOUR_CONTRIBUTIONS_SHEET_ID"
  googleSheetCrdrId: "YOUR_CRDR_SHEET_ID"
  googleSheetUsersId: "YOUR_USERS_SHEET_ID"
```

---

## 5. CI/CD Pipelines & Cross-Repository Workflows

### Repo 1 Workflow (`TechFamily-21/Apartment/.github/workflows/ci.yml`)

```yaml
name: CI/CD Pipeline

on:
  push:
    branches: ["**"]

permissions:
  contents: read
  packages: write

jobs:
  test:
    name: Lint & smoke-test
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - name: Set up Python 3.11
        uses: actions/setup-python@v5
        with:
          python-version: "3.11"
      - name: Install dependencies
        run: pip install -r requirements.txt
      - name: Syntax & Import Check
        run: python -c "import app; print('App module loaded successfully')"

  build-and-deploy-gitops:
    name: Build Docker Image & Update GitOps Repo
    needs: test
    if: github.event_name == 'push' && github.ref == 'refs/heads/main'
    runs-on: ubuntu-latest

    steps:
      - name: Checkout Code
        uses: actions/checkout@v4

      - name: Log in to GHCR
        uses: docker/login-action@v3
        with:
          registry: ghcr.io
          username: ${{ github.actor }}
          password: ${{ secrets.GITHUB_TOKEN }}

      - name: Build and Push Docker Image
        uses: docker/build-push-action@v5
        with:
          context: .
          push: true
          tags: |
            ghcr.io/kiranmsteam/apartment-app:sha-${{ github.sha }}
            ghcr.io/kiranmsteam/apartment-app:latest

      - name: Update GitOps Repository
        run: |
          IMAGE_TAG="sha-${{ github.sha }}"
          git clone https://x-access-token:${{ secrets.GITOPS_PAT }}@github.com/kiranmsteam/Appartment-ArgoCD.git gitops-repo
          
          cd gitops-repo/k8s
          sed -i "s/newTag:.*/newTag: ${IMAGE_TAG}/g" kustomization.yaml
          
          cd ../helm-chart/apartment-app
          sed -i "s/tag:.*/tag: \"${IMAGE_TAG}\"/g" values.yaml
          
          git config user.name "GitHub Action"
          git config user.email "action@github.com"
          git commit -am "ci: automated image tag update to ${IMAGE_TAG}" || echo "No changes"
          git push
```

---

## 6. Google Sheets Integration & Secret Management

### Safe Local Secret Creation (Never Committed to Git)

Create local Kubernetes secrets on your laptop out-of-band:

```powershell
# 1. Create Secret for GCP Service Account JSON key:
.\bin\kubectl.exe create secret generic gcp-credentials --from-file=credentials.json=C:\K8-ArgoCD\app\gcp-key.json -n apartment-app

# 2. Create Secret for Google Sheet IDs and Passwords:
.\bin\kubectl.exe create secret generic app-env-secret -n apartment-app `
  --from-literal=GOOGLE_SHEET_BALANCE_ID='<YOUR_BALANCE_SHEET_ID>' `
  --from-literal=GOOGLE_SHEET_CONTRIBUTIONS_ID='<YOUR_CONTRIBUTIONS_SHEET_ID>' `
  --from-literal=GOOGLE_SHEET_CRDR_ID='<YOUR_CRDR_SHEET_ID>' `
  --from-literal=GOOGLE_SHEET_USERS_ID='<YOUR_USERS_SHEET_ID>' `
  --from-literal=EMAIL_SENDER='<YOUR_EMAIL>' `
  --from-literal=EMAIL_APP_PASSWORD='<YOUR_APP_PASSWORD>' `
  --from-literal=SECRET_KEY='<YOUR_SECRET_KEY>'
```

---

## 7. Essential Command Reference Cheat Sheet

```powershell
# Check nodes and cluster
.\bin\kubectl.exe get nodes -o wide

# Check pods in apartment-app namespace
.\bin\kubectl.exe get pods -n apartment-app

# Port-Forward Apartment Application -> http://localhost:8081
.\bin\kubectl.exe port-forward svc/apartment-app-service -n apartment-app 8081:8080

# Port-Forward ArgoCD UI -> https://localhost:8080
.\bin\kubectl.exe port-forward svc/argocd-server -n argocd 8080:443
```
