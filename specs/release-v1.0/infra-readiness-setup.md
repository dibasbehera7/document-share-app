# Infrastructure Readiness Guide — Document Share App (release-v1.0)

> **Audience:** DevOps / Platform Engineering team  
> **Purpose:** Step-by-step setup and verification guide to make all required infrastructure components ready before development and testing can begin.  
> **Last updated:** September 2026

---

## Overview

This guide covers the three tracks that must be completed before the application can be developed, run locally, or tested end-to-end:

| Track | Description | Who |
|---|---|---|
| **Track A** | Microsoft 365 Developer Tenant + Entra ID App Registration | DevOps / Developer |
| **Track B** | Local Development Environment (Docker Compose) | Developer |
| **Track C** | AWS Cloud Infrastructure (for cloud-global topology) | DevOps / Cloud Engineer |

Each track is independent. For initial development and local testing, only **Track A + Track B** are required. Track C is required before staging/production deployment.

---

## Track A — Microsoft 365 Developer Tenant + Entra ID

> **Why this is needed:** SharePoint Online requires an organisational Microsoft 365 tenant. Personal Microsoft accounts (`@outlook.com`, `@hotmail.com`) cannot access SharePoint APIs or register app credentials that work with Graph API scopes. The error `AADSTS500200` confirms this limitation.

> **Recommended path:** Microsoft 365 Developer Program — a **free** E5 sandbox tenant with SharePoint Online, full Graph API access, and 25 user licences. It renews automatically while actively used for development.

### Step A1 — Sign up for the M365 Developer Program

1. Go to: https://developer.microsoft.com/en-us/microsoft-365/dev-program
2. Click **Join now**
3. Sign in with your personal Microsoft account (`@outlook.com` is fine for signup)
4. Fill in the form:
   - Country: your location
   - Company: your organisation name (can be anything for dev)
   - Primary focus: **Enterprise application development**
5. Choose **Instant sandbox** (recommended — comes pre-seeded with SharePoint, users, and sample data)
6. Set a password for the admin account
7. Note down your new tenant domain: `yourname.onmicrosoft.com`

**Verification:** You should receive an admin account like `admin@yourname.onmicrosoft.com`. Log in at https://admin.microsoft.com — you should see Microsoft 365 E5 licences.

**Expected time:** ~10 minutes

---

### Step A2 — Verify SharePoint Online is Provisioned

1. Log in to https://yourname.sharepoint.com with your admin account
2. You should see the SharePoint home page with sample document libraries
3. If prompted to set up SharePoint, click through the setup wizard

**Verification:** Navigate to https://yourname.sharepoint.com/sites — you should see at least one sample site (e.g., `Mark 8 Project Team` or similar from the sandbox).

---

### Step A3 — Register an App in Entra ID (Microsoft Entra Admin Center)

This creates the application identity that the Document Share App backend uses to call Graph API.

1. Go to: https://entra.microsoft.com (sign in with your `admin@yourname.onmicrosoft.com` account)
2. Navigate to: **Identity → Applications → App registrations → New registration**
3. Fill in:
   - **Name:** `document-share-app-dev`
   - **Supported account types:** `Accounts in this organizational directory only (Single tenant)`
   - **Redirect URI:** `Web` → `http://localhost:3000/auth/callback`
4. Click **Register**
5. Note down:
   - **Application (client) ID** → save as `ENTRA_CLIENT_ID`
   - **Directory (tenant) ID** → save as `ENTRA_TENANT_ID`

---

### Step A4 — Create a Client Secret

1. In your app registration, go to **Certificates & secrets → Client secrets → New client secret**
2. Description: `dev-secret`
3. Expiry: `24 months`
4. Click **Add**
5. **Copy the secret value immediately** — it is only shown once
6. Save as `ENTRA_CLIENT_SECRET`

---

### Step A5 — Configure API Permissions

1. In your app registration, go to **API permissions → Add a permission**
2. Select **Microsoft Graph**
3. Select **Application permissions** and add:
   - `Sites.Read.All`
   - `Sites.ReadWrite.All`
   - `Files.ReadWrite.All`
4. Select **Delegated permissions** and add:
   - `Files.Read.All`
   - `offline_access`
   - `openid`
   - `profile`
   - `email`
5. Click **Add permissions**
6. Click **Grant admin consent for [your tenant]** → Confirm

**Verification:** All permissions should show a green ✅ checkmark in the **Status** column.

---

### Step A6 — Get Your SharePoint Site and Drive IDs

These IDs are needed in environment variables.

1. Open a browser and go to:
   ```
   https://graph.microsoft.com/v1.0/sites/yourname.sharepoint.com:/sites/Mark8ProjectTeam
   ```
   (Replace `Mark8ProjectTeam` with any existing sample site name)

   Or list all sites:
   ```
   https://graph.microsoft.com/v1.0/sites?search=*
   ```

   Use the [Microsoft Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) — sign in with your admin account and run the query.

2. Note down:
   - `id` field → save as `SPO_SITE_ID`
   - Run `GET /sites/{siteId}/drives` to get the drive ID → save as `SPO_DRIVE_ID`

---

### Step A7 — Create a Document Library for Testing

1. Go to your SharePoint site: https://yourname.sharepoint.com/sites/Mark8ProjectTeam
2. Click **New → Document library**
3. Name it: `DocumentShareAppDev`
4. Note the library's list ID from the URL or via Graph API:
   ```
   GET https://graph.microsoft.com/v1.0/sites/{siteId}/lists?$filter=displayName eq 'DocumentShareAppDev'
   ```
5. Save the `id` field as `SPO_LIBRARY_LIST_ID`

---

### Step A8 — Environment Variable Checklist (from Track A)

At the end of Track A, you should have all of the following:

```env
# Entra ID / Microsoft Identity
ENTRA_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
ENTRA_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
ENTRA_CLIENT_SECRET=your-client-secret-value

# SharePoint Online
SPO_TENANT=yourname.onmicrosoft.com
SPO_SITE_URL=https://yourname.sharepoint.com/sites/Mark8ProjectTeam
SPO_SITE_ID=yourname.sharepoint.com,xxxxxxxx-xxxx,xxxxxxxx-xxxx
SPO_DRIVE_ID=b!xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
SPO_LIBRARY_LIST_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
```

---

## Track B — Local Development Environment (Docker Compose)

> **Purpose:** Run the full application stack on a developer machine without needing a live cloud account for every component. SharePoint Online (from Track A) is still needed for actual SP operations; everything else runs locally.

### Prerequisites

| Tool | Minimum version | Install link |
|---|---|---|
| Docker Desktop | 4.x | https://www.docker.com/products/docker-desktop |
| Docker Compose | v2 (bundled with Docker Desktop) | — |
| Node.js | 22.x | https://nodejs.org |
| Git | any | https://git-scm.com |

---

### Step B1 — Clone the Repository

```bash
git clone https://github.com/your-org/document-share-app.git
cd document-share-app
```

---

### Step B2 — Create the Local Environment File

```bash
cp .env.example .env
```

Edit `.env` and fill in the values from Track A plus the local service defaults:

```env
# ── Auth ──────────────────────────────────────────────────────
AUTH_DRIVER=entra-global
ENTRA_TENANT_ID=<from Track A>
ENTRA_CLIENT_ID=<from Track A>
ENTRA_CLIENT_SECRET=<from Track A>

# ── SharePoint ────────────────────────────────────────────────
SHAREPOINT_DRIVER=graph-global
SPO_TENANT=<yourname>.onmicrosoft.com
SPO_SITE_ID=<from Track A>
SPO_DRIVE_ID=<from Track A>
SPO_LIBRARY_LIST_ID=<from Track A>

# ── Storage (local MinIO) ─────────────────────────────────────
STORAGE_DRIVER=s3-minio
MINIO_ENDPOINT=http://localhost:9000
MINIO_ACCESS_KEY=minioadmin
MINIO_SECRET_KEY=minioadmin
MINIO_BUCKET=document-share-dev

# ── Metadata Store (local PostgreSQL) ─────────────────────────
METADATA_DRIVER=postgres
DATABASE_URL=postgresql://docshare:docshare@localhost:5432/docshare

# ── Session ───────────────────────────────────────────────────
SESSION_SECRET=change-me-for-production-use-32-chars-min
SESSION_TTL_HOURS=8

# ── WOPI (use SharePoint as WOPI host) ────────────────────────
WOPI_MOCK=false

# ── Webhook ───────────────────────────────────────────────────
# For local dev, use ngrok or similar to expose /webhook endpoint
WEBHOOK_NOTIFICATION_URL=https://your-ngrok-url.ngrok.io/webhook/graph

# ── Debounce (local Redis via BullMQ) ─────────────────────────
REDIS_URL=redis://localhost:6379

# ── Audit Log (stdout in local dev) ──────────────────────────
AUDIT_DRIVER=stdout

# ── App ───────────────────────────────────────────────────────
PORT=3000
NODE_ENV=development
LOG_LEVEL=debug
```

---

### Step B3 — Start All Local Services

```bash
docker compose up -d
```

This starts:
- `postgres` — metadata store on port 5432
- `minio` — S3-compatible object storage on port 9000 (console on 9001)
- `redis` — debounce timer queue on port 6379
- `keycloak` — OIDC stub for local auth testing on port 8080 (only used when `AUTH_DRIVER=keycloak`)

**Verification:**
```bash
docker compose ps
# All services should show status: Up
```

---

### Step B4 — Run Database Migrations

```bash
npm install
npm run db:migrate
```

**Verification:**
```bash
npm run db:status
# Should show all migrations applied
```

---

### Step B5 — Seed Sample Data

```bash
npm run db:seed
```

This inserts two sample document records and one owner-permission record so the UI is immediately usable after startup.

---

### Step B6 — Start the API Service

```bash
npm run dev
```

The API service starts with hot-reload. Source file changes reload within 10 seconds.

**Verification — Health Check:**
```bash
curl http://localhost:3000/health
```

Expected response:
```json
{
  "status": "ok",
  "adapters": {
    "sharepoint": "ok",
    "storage": "ok",
    "metadata": "ok",
    "auth": "ok"
  }
}
```

If `sharepoint` shows `unavailable`, check your `.env` values from Track A.

---

### Step B7 — Expose Webhook Endpoint (for SharePoint notifications)

SharePoint needs a publicly reachable HTTPS URL to send webhook notifications. Use [ngrok](https://ngrok.com) for local development:

```bash
# Install ngrok: https://ngrok.com/download
ngrok http 3000
```

Copy the `https://` forwarding URL (e.g., `https://abc123.ngrok.io`) and update `.env`:

```env
WEBHOOK_NOTIFICATION_URL=https://abc123.ngrok.io/webhook/graph
```

Restart the API service after updating `.env`.

---

### Step B8 — Register Graph Subscription (Webhook)

Call the API to register the webhook subscription against your SharePoint library:

```bash
curl -X POST http://localhost:3000/admin/libraries/register \
  -H "Content-Type: application/json" \
  -H "Cookie: <your-session-cookie-after-login>" \
  -d '{
    "site_id": "<SPO_SITE_ID>",
    "drive_id": "<SPO_DRIVE_ID>",
    "library_list_id": "<SPO_LIBRARY_LIST_ID>"
  }'
```

**Verification:** The response should contain a `subscription_id`. In SharePoint, go to your document library, upload a file, and check the API logs — you should see an incoming webhook notification within 1–5 minutes.

---

### Step B9 — Local Environment Verification Checklist

Run through each item before declaring the local environment ready:

- [ ] `docker compose ps` — all containers `Up`
- [ ] `GET /health` returns `{"status":"ok"}` for all adapters
- [ ] Can sign in via Entra ID (browser opens `http://localhost:3000/auth/login`, completes OIDC flow)
- [ ] `GET /documents` returns HTTP 200 (with seeded sample documents)
- [ ] Upload a test `.docx` file via `POST /documents/upload` → returns HTTP 201
- [ ] Download the uploaded file via `GET /documents/{docId}/download` → file downloads correctly
- [ ] Webhook notification arrives in API logs after editing the uploaded file in SharePoint Online
- [ ] MinIO console (http://localhost:9001) shows the uploaded file object in `document-share-dev` bucket
- [ ] Audit log entries appear in stdout after upload + download

---

## Track C — AWS Cloud Infrastructure (Cloud-Global Topology)

> **Required for:** Staging and production deployments. Not needed for local development.

### Prerequisites

| Tool | Version | Notes |
|---|---|---|
| AWS CLI | v2 | Configure with IAM credentials |
| AWS CDK CLI | v2 | `npm install -g aws-cdk` |
| Node.js | 22.x | For CDK TypeScript |
| An AWS account | — | IAM user with AdministratorAccess for initial bootstrap |

---

### Step C1 — Configure AWS CLI

```bash
aws configure
# AWS Access Key ID: <your-key>
# AWS Secret Access Key: <your-secret>
# Default region: ap-southeast-1  (or your preferred region)
# Default output format: json
```

**Verification:**
```bash
aws sts get-caller-identity
# Should return your account ID and IAM user ARN
```

---

### Step C2 — Bootstrap CDK

```bash
cd infra/cloud
npm install
cdk bootstrap aws://YOUR_ACCOUNT_ID/YOUR_REGION
```

This creates the CDK bootstrap stack (S3 bucket + ECR repo for CDK assets).

---

### Step C3 — Deploy Dev Stack

```bash
cdk deploy DocumentShareApp-dev --require-approval never
```

This provisions:
- API Gateway + Lambda / ECS task definition
- DynamoDB table (no Object Lock, 30-day log retention in dev)
- S3 bucket for documents (versioning enabled, Block Public Access, SSE-KMS)
- S3 bucket for audit logs
- CloudFront distribution
- IAM roles with least-privilege policies
- EventBridge rules for debounce timer

**Verification:**
```bash
aws cloudformation describe-stacks --stack-name DocumentShareApp-dev \
  --query 'Stacks[0].StackStatus'
# Expected: "CREATE_COMPLETE" or "UPDATE_COMPLETE"
```

---

### Step C4 — Set Secrets in AWS Secrets Manager

```bash
aws secretsmanager create-secret \
  --name "document-share-app/dev/entra" \
  --secret-string '{
    "tenant_id": "<ENTRA_TENANT_ID>",
    "client_id": "<ENTRA_CLIENT_ID>",
    "client_secret": "<ENTRA_CLIENT_SECRET>"
  }'

aws secretsmanager create-secret \
  --name "document-share-app/dev/sharepoint" \
  --secret-string '{
    "site_id": "<SPO_SITE_ID>",
    "drive_id": "<SPO_DRIVE_ID>",
    "library_list_id": "<SPO_LIBRARY_LIST_ID>"
  }'

aws secretsmanager create-secret \
  --name "document-share-app/dev/session" \
  --secret-string '{"secret": "<32-char-random-string>"}'
```

---

### Step C5 — Deploy Application

```bash
# Build and push Docker image (if using ECS)
docker build -t document-share-app .
aws ecr get-login-password | docker login --username AWS --password-stdin \
  YOUR_ACCOUNT_ID.dkr.ecr.YOUR_REGION.amazonaws.com
docker tag document-share-app:latest \
  YOUR_ACCOUNT_ID.dkr.ecr.YOUR_REGION.amazonaws.com/document-share-app:latest
docker push YOUR_ACCOUNT_ID.dkr.ecr.YOUR_REGION.amazonaws.com/document-share-app:latest

# Trigger ECS service update
aws ecs update-service \
  --cluster document-share-app-dev \
  --service api-service \
  --force-new-deployment
```

---

### Step C6 — Cloud Infrastructure Verification Checklist

- [ ] CloudFormation stack status: `CREATE_COMPLETE`
- [ ] ECS service or Lambda function is healthy (0 failed tasks)
- [ ] `GET https://<cloudfront-url>/health` returns `{"status":"ok"}` for all adapters
- [ ] S3 documents bucket exists with Block Public Access enabled
- [ ] S3 audit bucket exists with Object Lock enabled (check: `aws s3api get-object-lock-configuration --bucket <audit-bucket-name>`)
- [ ] DynamoDB table exists and point-in-time recovery is enabled
- [ ] CloudFront distribution is deployed and returning non-403 responses
- [ ] Entra ID app registration redirect URI updated to include the CloudFront HTTPS URL

---

## Webhook Public URL Requirements

For SharePoint to deliver webhook notifications, the API service must be reachable from the internet over HTTPS.

| Environment | Recommended approach |
|---|---|
| Local development | ngrok (`ngrok http 3000`) — free tier sufficient |
| Staging / Production | CloudFront URL (Track C) or your custom domain with SSL |
| CI/CD pipeline testing | Use a dedicated ngrok authtoken or a tunnel service |

> **Important:** The webhook validation handshake must respond within **5 seconds**. Ensure your tunnel or deployment has low enough latency for SharePoint to receive the validation token response before it times out.

---

## Quick Reference: Minimum Variables by Environment

### Local Development (`.env`)

```
AUTH_DRIVER, ENTRA_TENANT_ID, ENTRA_CLIENT_ID, ENTRA_CLIENT_SECRET
SHAREPOINT_DRIVER, SPO_SITE_ID, SPO_DRIVE_ID, SPO_LIBRARY_LIST_ID
STORAGE_DRIVER=s3-minio, MINIO_ENDPOINT, MINIO_ACCESS_KEY, MINIO_SECRET_KEY
METADATA_DRIVER=postgres, DATABASE_URL
SESSION_SECRET, REDIS_URL
WEBHOOK_NOTIFICATION_URL (ngrok HTTPS URL)
```

### AWS Dev / Staging (from Secrets Manager + CDK outputs)

```
AUTH_DRIVER=entra-global
SHAREPOINT_DRIVER=graph-global
STORAGE_DRIVER=s3-aws
METADATA_DRIVER=dynamodb
AUDIT_DRIVER=s3
AWS_REGION, S3_DOCUMENTS_BUCKET, S3_AUDIT_BUCKET
DYNAMODB_TABLE_NAME
SECRETS_MANAGER_PREFIX=document-share-app/dev/
```

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `AADSTS500200` on sign-in | Personal Microsoft account used for app registration | Complete Track A — use M365 Developer tenant |
| `AADSTS700016` (app not found) | Wrong tenant ID in `.env` | Verify `ENTRA_TENANT_ID` matches the tenant where the app was registered |
| `403 Forbidden` from Graph API | Admin consent not granted | Repeat Step A5 — grant admin consent |
| SharePoint health adapter `unavailable` | `SPO_SITE_ID` or `SPO_DRIVE_ID` incorrect | Re-run Step A6 to get correct IDs |
| Webhook validation timeout | ngrok tunnel not running or wrong port | Ensure `ngrok http 3000` is running and URL in `.env` is current |
| MinIO `NoSuchBucket` error | Bucket not created | Run `npm run storage:init` or create bucket manually in MinIO console (port 9001) |
| DynamoDB `ResourceNotFoundException` | Migrations not run for cloud | Run `npm run db:migrate -- --env cloud` |
| `401 Unauthorized` on API calls | Session cookie expired or missing | Sign in again at `/auth/login` |

---

## Summary: What "Infra Ready" Means for v1.0

Before any developer can run or test the application, the following must all be true:

1. ✅ M365 Developer tenant exists at `*.onmicrosoft.com`
2. ✅ Entra ID app registration exists with admin consent granted on all required Graph API permissions
3. ✅ SharePoint document library created and its IDs recorded
4. ✅ `.env` file populated with all required values from Track A
5. ✅ `docker compose up` starts all local services healthy
6. ✅ `GET /health` returns `ok` for all adapters
7. ✅ ngrok tunnel active (for webhook testing)
8. ✅ Graph subscription registered via the API (`/admin/libraries/register`)
