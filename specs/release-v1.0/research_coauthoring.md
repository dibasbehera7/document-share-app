# SharePoint Co-authoring, Document Management & API Research

> **Research compiled:** September 2026  
> **Scope:** SharePoint Online (SPO), SharePoint Server On-Premises, Microsoft Graph API, Webhooks, Co-authoring protocols, Sovereign/Sensitive country deployments, AWS + S3 integration architecture

---

## Table of Contents

1. [How SharePoint Stores Documents](#1-how-sharepoint-stores-documents)
2. [Co-authoring — How It Works Technically](#2-co-authoring--how-it-works-technically)
3. [Version History — Storage Model](#3-version-history--storage-model)
4. [Webhooks — Full Technical Detail](#4-webhooks--full-technical-detail)
5. [Microsoft Graph API — Complete Reference](#5-microsoft-graph-api--complete-reference)
6. [SharePoint REST API — Complete Reference](#6-sharepoint-rest-api--complete-reference)
7. [WOPI Protocol — Document Editing Integration](#7-wopi-protocol--document-editing-integration)
8. [Compliance — DLP, Sensitivity Labels, IRM](#8-compliance--dlp-sensitivity-labels-irm)
9. [Sovereign Countries — China, KSA, Malaysia](#9-sovereign-countries--china-ksa-malaysia)
10. [Architecture — Enterprise Cloud (AWS Global + SPO)](#10-architecture--enterprise-cloud-aws-global--spo)
11. [Architecture — Sovereign/On-Premises (China, KSA)](#11-architecture--sovereignon-premises-china-ksa)
12. [S3 Sync Pattern — Version Bumping Flow](#12-s3-sync-pattern--version-bumping-flow)
13. [API Limits & Throttling](#13-api-limits--throttling)
14. [Official Documentation Links](#14-official-documentation-links)

---

## 1. How SharePoint Stores Documents

### Storage Architecture

SharePoint stores documents across two layers:

| Layer | SharePoint Online (SPO) | SharePoint Server (On-Prem) |
|---|---|---|
| **Binary (file blob)** | Azure Blob Storage | SQL Server FILESTREAM or Remote Blob Store (RBS) |
| **Metadata** | Azure SQL Database | SQL Server Content Database |
| **Search Index** | Azure Cognitive Search | SharePoint Fast Search |
| **Versioning** | Delta-compressed blobs in Azure | SQL FILESTREAM (full or delta) |

### Upload Flow (REST API)

```
Client → POST /_api/web/GetFolderByServerRelativeUrl('{path}')/Files/add(url='{filename}',overwrite=true)
           ↓
     SPO receives binary stream
           ↓
     Blob stored → Azure Blob Storage
     Metadata stored → Content Database (SQL)
     UniqueId (GUID) assigned
     ETag assigned (changes on every version)
     List item record created/updated
           ↓
     Response: file metadata JSON
```

### Upload via Microsoft Graph

```http
# Small files < 4 MB
PUT https://graph.microsoft.com/v1.0/sites/{siteId}/drives/{driveId}/items/{parentId}:/{filename}:/content
Authorization: Bearer {token}
Content-Type: application/octet-stream

[file binary]

# Large files > 4 MB — requires upload session
POST https://graph.microsoft.com/v1.0/sites/{siteId}/drives/{driveId}/items/{parentId}:/{filename}:/createUploadSession
Content-Type: application/json
{
  "item": {
    "@microsoft.graph.conflictBehavior": "replace",
    "name": "filename.docx"
  }
}
# Returns: uploadUrl (valid 24 hours)

# Then upload chunks:
PUT {uploadUrl}
Content-Range: bytes 0-4999999/10000000
Content-Length: 5000000
[chunk binary]
```

---

## 2. Co-authoring — How It Works Technically

### Protocol Stack

SharePoint uses two layered protocols for real-time co-authoring:

```
┌─────────────────────────────────────────────────────────┐
│              Office Web Apps (Browser)                   │
├─────────────────────────────────────────────────────────┤
│  WOPI Protocol (Web App Open Platform Interface)         │
│  → File access: CheckFileInfo, GetFile, PutFile         │
│  → Session management: Lock, RefreshLock, Unlock        │
├─────────────────────────────────────────────────────────┤
│  MS-FSSHTTP Protocol (File Sync via SOAP over HTTP)      │
│  → Co-authoring sessions: Join, Refresh, Leave          │
│  → Shared locks: CoauthSubrequest                       │
│  → Cell-level sync: CellSubrequest (partial file sync)  │
├─────────────────────────────────────────────────────────┤
│  SharePoint CellStorage Web Service                      │
│  Endpoint: /_vti_bin/cell.svc                           │
└─────────────────────────────────────────────────────────┘
```

**Protocol specs:**
- MS-FSSHTTP: https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fsshttp/6d078cbe-2651-43a0-b460-685ac3f14c45
- WOPI: https://learn.microsoft.com/en-us/microsoft-365/cloud-storage-partner-program/online/

### Co-authoring Session Lifecycle

```
USER A opens document
    │
    ▼
WOPI: CheckFileInfo  →  Returns: file metadata, permissions, version, supports coauth=true
    │
    ▼
MS-FSSHTTP: CoauthSubrequest (JoinCoauthoringSession)
    │  Server checks: any other editors?
    ├── No other editors  → CoauthStatus = "Alone"   (exclusive-like behavior)
    └── Other editors    → CoauthStatus = "Active"   (shared lock issued)
    │
    ▼
USER B opens SAME document
    │
    ▼
MS-FSSHTTP: CoauthSubrequest (JoinCoauthoringSession)
    │  Server: existing session found
    └── CoauthStatus = "Active"  (User B joins shared lock)
    │
    ▼
BOTH USERS EDIT
    │
    ├── Each client sends CellSubrequest (partial content sync, binary delta)
    ├── Server merges changes using OfficeMath conflict resolution
    ├── Each client receives updates from other authors
    ├── Presence/cursor info exchanged
    │
    ▼
MS-FSSHTTP: CoauthSubrequest (RefreshCoauthoringSession) every ~60s
    │  Keeps shared lock alive (prevents timeout)
    │
    ▼
USER A closes document
    │
    ▼
MS-FSSHTTP: CoauthSubrequest (LeaveCoauthoringSession)
    │  Server removes User A from session
    ├── Other editors still present → session continues
    └── No editors left → session ends, major version saved, lock released
```

### Lock Types in FSSHTTP

| Lock Type | When Used | Behavior |
|---|---|---|
| **Exclusive Lock** | Single editor only | Only one user can edit; blocks others |
| **Shared Lock** | Co-authoring active | Multiple users; changes merged |
| **Schema Lock** | Transition state | Converting exclusive → shared when 2nd user joins |
| **No Lock** | View-only / read | No modifications allowed |

### CoauthVersionPeriod (On-Premises)

SharePoint Server saves a version snapshot during co-authoring at a configurable interval:

```powershell
# Default: 30 minutes
# Check current setting:
$web = Get-SPWeb "https://intranet.contoso.com"
$list = $web.Lists["Documents"]
$list.CoauthoringVersionPeriod   # returns minutes

# Change to 15 minutes:
$list.CoauthoringVersionPeriod = 15
$list.Update()
```

Docs: https://learn.microsoft.com/en-us/sharepoint/governance/configure-the-co-authoring-versioning-period

---

## 3. Version History — Storage Model

### Version Numbering

| Type | Format | When Created |
|---|---|---|
| Major version | `1.0`, `2.0`, `3.0` | Manual publish, Check-In (major), or when co-auth session ends |
| Minor version | `1.1`, `1.2`, `2.1` | Auto-save during editing, draft saves |

### Storage Structure

```
Document Library
└── proposal.docx  (current = v3.0)
    │
    ├── v1.0 → Full binary blob (first upload)
    ├── v2.0 → Binary delta from v1.0 (compressed diff)
    ├── v3.0 → Binary delta from v2.0 (CURRENT — live file)
    │
    └── Each version metadata:
        ├── VersionLabel: "3.0"
        ├── Created: 2026-09-06T10:30:00Z
        ├── CreatedBy: john@contoso.com
        ├── IsCurrentVersion: true
        └── Url: (internal blob URL)
```

**Storage efficiency:** SPO uses binary delta compression after v1 — only the diff is stored, not the full file per version.

### Version History APIs

```http
# Graph API — list versions
GET https://graph.microsoft.com/v1.0/sites/{siteId}/drives/{driveId}/items/{itemId}/versions

# Graph API — download specific version
GET https://graph.microsoft.com/v1.0/sites/{siteId}/drives/{driveId}/items/{itemId}/versions/{versionId}/content

# Graph API — restore a previous version
POST https://graph.microsoft.com/v1.0/sites/{siteId}/drives/{driveId}/items/{itemId}/versions/{versionId}/restoreVersion

# SharePoint REST API — list versions
GET https://{tenant}.sharepoint.com/sites/{site}/_api/web/lists/getbytitle('{library}')/items({itemId})/versions

# SharePoint REST API — version response fields:
# VersionLabel, Created, CreatedBy, IsCurrentVersion, Url, CheckInComment
```

### Version Limits (Configurable)

```
Library Settings → Versioning Settings:
  ├── Keep major versions: 50 (default), max unlimited
  ├── Keep minor versions: 50 (default)
  └── Require Check Out: Yes/No
```

---

## 4. Webhooks — Full Technical Detail

### Critical Clarifications

> **SPO webhooks do NOT fire on "user closes document."**  
> They fire on **list-level item changes** — any save, upload, update, delete, or permission change.  
> There is **no "edit session ended" event** in SPO webhooks.  
> There is **no UI to configure webhooks** — 100% API/code-driven.

### What Events Trigger SPO Webhooks

```
TRIGGERS (fires webhook):          DOES NOT TRIGGER:
├── File uploaded                  ├── User opens document
├── File saved / auto-saved        ├── User starts editing
├── File version created           ├── User closes browser tab
├── File deleted                   ├── Co-authoring session joined
├── Metadata updated               └── Co-authoring session left
├── File moved/renamed
└── Permissions changed
```

### Webhook Delivery Behavior

```
Event occurs in SPO
    │
    ├── SPO batches notifications (does NOT send immediately)
    ├── Delivery delay: 0 to ~5 minutes (variable, not guaranteed)
    ├── Payload: minimal — only subscription ID + site info
    ├── Your endpoint must respond HTTP 200 within 5 seconds
    └── Retry on failure: up to 5 attempts, then dropped permanently
```

### Webhook Payload (What You Receive)

```json
{
  "value": [
    {
      "subscriptionId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
      "clientState": "your-secret-validation-token",
      "expirationDateTime": "2026-12-31T00:00:00.000Z",
      "resource": "lists/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
      "tenantId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
      "siteUrl": "https://tenant.sharepoint.com/sites/mysite",
      "webId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    }
  ]
}
```

Note: No file name, no version, no user, no change type in the payload. You must call the Change Log API afterward.

### Step 1 — Deploy Your Webhook Receiver Endpoint

Requirements:
- Publicly reachable HTTPS endpoint
- Must handle the validation handshake on registration
- Must return HTTP 200 within 5 seconds (offload processing async)

```
Validation handshake (one-time, on registration):
GET https://your-app.com/webhook/sharepoint?validationtoken=abc123xyz
→ Respond: HTTP 200, Content-Type: text/plain, Body: abc123xyz
```

### Step 2 — Register Webhook via API

```http
POST https://{tenant}.sharepoint.com/sites/{site}/_api/web/lists('{listId}')/subscriptions
Authorization: Bearer {token}
Content-Type: application/json

{
  "resource": "https://{tenant}.sharepoint.com/sites/{site}/_api/web/lists('{listId}')",
  "notificationUrl": "https://your-app.com/webhook/sharepoint",
  "expirationDateTime": "2026-12-31T00:00:00.000Z",
  "clientState": "your-secret-validation-token"
}
```

**Max expiry:** 180 days. Must renew before expiry or subscription is deleted.

### Step 3 — Query Change Log After Webhook Fires

```http
POST https://{tenant}.sharepoint.com/sites/{site}/_api/web/lists('{listId}')/getchanges
Authorization: Bearer {token}
Content-Type: application/json

{
  "query": {
    "__metadata": { "type": "SP.ChangeQuery" },
    "ChangeTokenStart": { "StringValue": "{last-stored-change-token}" },
    "Item": true,
    "Update": true,
    "Add": true,
    "DeleteObject": true,
    "File": true
  }
}
```

Store the `ChangeToken` from the last processed change to use as the start token next time.

### Manage Webhook Subscriptions

```http
# List all subscriptions on a library
GET https://{tenant}.sharepoint.com/sites/{site}/_api/web/lists('{listId}')/subscriptions

# Get specific subscription
GET https://{tenant}.sharepoint.com/sites/{site}/_api/web/lists('{listId}')/subscriptions('{subscriptionId}')

# Renew (PATCH to extend expiry)
PATCH https://{tenant}.sharepoint.com/sites/{site}/_api/web/lists('{listId}')/subscriptions('{subscriptionId}')
Content-Type: application/json
{
  "notificationUrl": "https://your-app.com/webhook/sharepoint",
  "expirationDateTime": "2027-06-01T00:00:00.000Z"
}

# Delete subscription
DELETE https://{tenant}.sharepoint.com/sites/{site}/_api/web/lists('{listId}')/subscriptions('{subscriptionId}')
```

### Microsoft Graph Change Notifications (Preferred — Richer)

Graph subscriptions provide richer payloads than SPO REST webhooks and unify OneDrive + SPO.

```http
# Register subscription
POST https://graph.microsoft.com/v1.0/subscriptions
Authorization: Bearer {token}
Content-Type: application/json

{
  "changeType": "updated",
  "notificationUrl": "https://your-app.com/webhook/graph",
  "resource": "/sites/{siteId}/drives/{driveId}/root",
  "expirationDateTime": "2026-09-30T23:59:59.000Z",
  "clientState": "your-secret-token",
  "lifecycleNotificationUrl": "https://your-app.com/webhook/lifecycle"
}
```

Graph notification payload (richer than SPO REST):

```json
{
  "value": [
    {
      "subscriptionId": "xxxxxxxx",
      "changeType": "updated",
      "resource": "sites/{siteId}/drives/{driveId}/root",
      "resourceData": {
        "@odata.type": "#Microsoft.Graph.DriveItem",
        "@odata.id": "sites/{siteId}/drives/{driveId}/items/{itemId}",
        "id": "{itemId}"
      },
      "clientState": "your-secret-token",
      "tenantId": "xxxxxxxx",
      "subscriptionExpirationDateTime": "2026-09-30T23:59:59Z"
    }
  ]
}
```

**Max expiry:** 30 days for driveItem/list subscriptions. Must renew proactively.

```http
# Renew Graph subscription
PATCH https://graph.microsoft.com/v1.0/subscriptions/{subscriptionId}
Content-Type: application/json
{ "expirationDateTime": "2026-10-30T23:59:59.000Z" }

# Delete Graph subscription
DELETE https://graph.microsoft.com/v1.0/subscriptions/{subscriptionId}
```

### "Editing Stopped" Signal — Options Compared

| Option | Accuracy | Complexity | Latency |
|---|---|---|---|
| Webhook + debounce (3 min no activity) | Medium | Low | ~3 min after last save |
| Poll `CheckOut` lock status | High | Medium | Near real-time (poll interval) |
| Check-In/Check-Out enforcement + RER | Highest | Medium | Immediate on check-in |
| Graph socket.IO subscription | Medium | Medium | Near real-time |

**Recommended for SPO:** Debounce pattern on Graph subscriptions:

```
onChange notification received
    │
    ▼
Reset 3-minute inactivity timer
    │
    ▼  (if no new notification within 3 minutes)
Timer fires → "editing stopped" assumed
    │
    ▼
Fetch latest file from SPO → compare ETag/version
    │
    ▼
If newer than stored → sync to S3
```

**Recommended for On-Premises (China/KSA):** Enforce Check-In/Check-Out:

```powershell
# Enforce checkout on library (on-prem PowerShell)
$web = Get-SPWeb "https://sp.contoso.local/sites/mysite"
$list = $web.Lists["Documents"]
$list.ForceCheckout = $true
$list.Update()
```

ItemCheckedIn Remote Event Receiver (RER) then fires a deterministic, reliable "editing done" event.

### On-Premises: Remote Event Receivers (RER) vs SPO Webhooks

| Feature | SPO Webhooks | On-Prem Remote Event Receivers |
|---|---|---|
| Registration | REST API call | Visual Studio / PowerShell on farm |
| Transport | HTTPS POST (push) | WCF SOAP service |
| Events | List-level change only | ItemAdded, ItemUpdated, ItemDeleted, ItemCheckedIn, etc. |
| "Edit done" event | No native support | `ItemCheckedIn` (when Check-Out enforced) |
| Payload richness | Minimal | Full event properties |
| Availability | SPO only | SP Server 2013/2016/2019/SE |

Reference: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins-modernize/from-remote-event-receivers-to-webhooks

---

## 5. Microsoft Graph API — Complete Reference

Base URL: `https://graph.microsoft.com/v1.0/`  
China (21Vianet): `https://microsoftgraph.chinacloudapi.cn/v1.0/`

### Authentication

```http
# Global
POST https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token
grant_type=client_credentials&client_id={id}&client_secret={secret}
&scope=https://graph.microsoft.com/.default

# China (21Vianet)
POST https://login.partner.microsoftonline.cn/{tenantId}/oauth2/v2.0/token
scope=https://microsoftgraph.chinacloudapi.cn/.default
```

### Required Permissions

| Permission | Type | Capability |
|---|---|---|
| `Sites.Read.All` | Application | Read sites and lists |
| `Sites.ReadWrite.All` | Application | Read and write files |
| `Files.ReadWrite.All` | Application | Full file operations |
| `Files.Read.All` | Delegated | Read on behalf of user |

### Sites

```http
# Get site by URL
GET /sites/{hostname}:/{server-relative-path}
# Example: /sites/contoso.sharepoint.com:/sites/mysite

# Get site by ID
GET /sites/{siteId}

# Search all sites
GET /sites?search=*

# Get site drives (document libraries)
GET /sites/{siteId}/drives
GET /sites/{siteId}/drive          # default drive
```

### Files & Folders

```http
# List root contents
GET /sites/{siteId}/drives/{driveId}/root/children

# List folder contents
GET /sites/{siteId}/drives/{driveId}/items/{folderId}/children

# Get item metadata
GET /sites/{siteId}/drives/{driveId}/items/{itemId}
GET /sites/{siteId}/drives/{driveId}/items/{itemId}?$select=id,name,eTag,lastModifiedDateTime,size,webUrl

# Get item by path
GET /sites/{siteId}/drives/{driveId}/root:/{path-to-file}

# Download file content
GET /sites/{siteId}/drives/{driveId}/items/{itemId}/content

# Upload small file (< 4 MB)
PUT /sites/{siteId}/drives/{driveId}/items/{parentId}:/{filename}:/content

# Create upload session (large files)
POST /sites/{siteId}/drives/{driveId}/items/{parentId}:/{filename}:/createUploadSession

# Copy file
POST /sites/{siteId}/drives/{driveId}/items/{itemId}/copy
{ "parentReference": { "driveId": "{id}", "id": "{folderId}" }, "name": "copy.docx" }

# Move / rename
PATCH /sites/{siteId}/drives/{driveId}/items/{itemId}
{ "parentReference": { "id": "{newParentId}" }, "name": "new-name.docx" }

# Delete file
DELETE /sites/{siteId}/drives/{driveId}/items/{itemId}

# Search
GET /sites/{siteId}/drives/{driveId}/root/search(q='keyword')

# Get sharing URL
GET /sites/{siteId}/drives/{driveId}/items/{itemId}?$select=webUrl,webDavUrl
```

### Versions

```http
# List all versions
GET /sites/{siteId}/drives/{driveId}/items/{itemId}/versions

# Get specific version
GET /sites/{siteId}/drives/{driveId}/items/{itemId}/versions/{versionId}

# Download specific version
GET /sites/{siteId}/drives/{driveId}/items/{itemId}/versions/{versionId}/content

# Restore a version
POST /sites/{siteId}/drives/{driveId}/items/{itemId}/versions/{versionId}/restoreVersion
```

### Permissions & Sharing

```http
# Create sharing link
POST /sites/{siteId}/drives/{driveId}/items/{itemId}/createLink
{ "type": "edit", "scope": "organization" }   # type: view|edit|embed, scope: anonymous|organization

# List permissions on item
GET /sites/{siteId}/drives/{driveId}/items/{itemId}/permissions

# Grant access to user
POST /sites/{siteId}/drives/{driveId}/items/{itemId}/invite
{
  "requireSignIn": true,
  "sendInvitation": false,
  "roles": ["write"],
  "recipients": [{ "email": "user@contoso.com" }]
}

# Remove permission
DELETE /sites/{siteId}/drives/{driveId}/items/{itemId}/permissions/{permId}
```

### Lists & List Items

```http
# Get all lists in a site
GET /sites/{siteId}/lists

# Get list items
GET /sites/{siteId}/lists/{listId}/items?expand=fields

# Create list item
POST /sites/{siteId}/lists/{listId}/items
{ "fields": { "Title": "Document Title", "Status": "Draft" } }

# Update list item fields
PATCH /sites/{siteId}/lists/{listId}/items/{itemId}/fields
{ "Status": "Published" }

# Delete list item
DELETE /sites/{siteId}/lists/{listId}/items/{itemId}
```

### Subscriptions (Webhooks)

```http
# Create subscription
POST /subscriptions
{
  "changeType": "updated",
  "notificationUrl": "https://your-app.com/webhook/graph",
  "resource": "/sites/{siteId}/drives/{driveId}/root",
  "expirationDateTime": "2026-09-30T23:59:59.000Z",
  "clientState": "your-secret-token"
}

# List subscriptions
GET /subscriptions

# Renew subscription
PATCH /subscriptions/{subscriptionId}
{ "expirationDateTime": "2026-10-30T23:59:59.000Z" }

# Delete subscription
DELETE /subscriptions/{subscriptionId}
```

### Socket.IO (Near Real-Time)

```http
# Get socket.IO endpoint for near real-time notifications
GET /sites/{siteId}/lists/{listId}/subscriptions/socketIo
# Returns: notificationUrl for socket.io client connection
```

Docs: https://learn.microsoft.com/en-us/graph/api/subscriptions-socketio

---

## 6. SharePoint REST API — Complete Reference

Base URL: `https://{tenant}.sharepoint.com/sites/{site}/_api/`

### Authentication Headers Required

```http
Authorization: Bearer {token}
Accept: application/json;odata=verbose
Content-Type: application/json;odata=verbose
X-RequestDigest: {digest}     # required for write operations
```

Get Request Digest:
```http
POST https://{tenant}.sharepoint.com/sites/{site}/_api/contextinfo
Authorization: Bearer {token}
# Returns: FormDigestValue (use as X-RequestDigest header)
```

### File Operations

```http
# Upload file
POST /_api/web/GetFolderByServerRelativeUrl('/sites/{site}/Shared Documents')/Files/add(url='{filename}',overwrite=true)
Content-Type: application/octet-stream
[binary]

# Get file by path
GET /_api/web/GetFileByServerRelativeUrl('/sites/{site}/Shared Documents/{filename}')

# Get file metadata + ETag
GET /_api/web/GetFileByServerRelativeUrl('/sites/{site}/Shared Documents/{filename}')?$select=UniqueId,Name,TimeCreated,TimeLastModified,Length,UIVersion,UIVersionLabel,ETag

# Delete file
POST /_api/web/GetFileByServerRelativeUrl('/sites/{site}/Shared Documents/{filename}')
X-HTTP-Method: DELETE

# Move file
POST /_api/web/GetFileByServerRelativeUrl('{source}')/moveto(newurl='{dest}',flags=1)

# Copy file
POST /_api/web/GetFileByServerRelativeUrl('{source}')/copyto(strnewurl='{dest}',boverwrite=true)
```

### Check-Out / Check-In (Locking)

```http
# Check Out (exclusive edit lock)
POST /_api/web/GetFileByServerRelativeUrl('/sites/{site}/Shared Documents/{filename}')/CheckOut

# Check In (release lock — triggers ItemCheckedIn event)
POST /_api/web/GetFileByServerRelativeUrl('/sites/{site}/Shared Documents/{filename}')/CheckIn(comment='Saved via API',checkintype=1)
# checkintype: 0 = minor version, 1 = major version, 2 = overwrite

# Discard checkout
POST /_api/web/GetFileByServerRelativeUrl('/sites/{site}/Shared Documents/{filename}')/UndoCheckOut
```

### Library & List Operations

```http
# List items in library
GET /_api/web/lists/getbytitle('Documents')/items?$select=FileLeafRef,FileRef,Modified,Editor/Title&$expand=Editor

# Get item versions
GET /_api/web/lists/getbytitle('Documents')/items({itemId})/versions

# Get change log (call after webhook)
POST /_api/web/lists('{listId}')/getchanges
Content-Type: application/json
{
  "query": {
    "__metadata": { "type": "SP.ChangeQuery" },
    "ChangeTokenStart": { "StringValue": "{last-change-token}" },
    "Item": true,
    "Update": true,
    "Add": true,
    "DeleteObject": true
  }
}
```

### Folder Operations

```http
# Create folder
POST /_api/web/folders
Content-Type: application/json
{ "__metadata": { "type": "SP.Folder" }, "ServerRelativeUrl": "/sites/{site}/Shared Documents/NewFolder" }

# List folder contents
GET /_api/web/GetFolderByServerRelativeUrl('/sites/{site}/Shared Documents')/Files

# Get folder
GET /_api/web/GetFolderByServerRelativeUrl('/sites/{site}/Shared Documents/MyFolder')
```

### WOPI URL Generation

```http
# Get edit URL for Office Web Apps
POST /_api/web/GetFileByServerRelativeUrl('{file-path}')/ListItemAllFields/GetWopiFrameUrl(@action)?@action='edit'

# Get view URL
POST /_api/web/GetFileByServerRelativeUrl('{file-path}')/ListItemAllFields/GetWopiFrameUrl(@action)?@action='view'

# Open in desktop Office client (protocol handler)
# Returns: ms-word:ofe|u|https://...  (open in Word desktop)
POST /_api/web/GetFileByServerRelativeUrl('{file-path}')/ListItemAllFields/GetWopiFrameUrl(@action)?@action='mobileView'
```

---

## 7. WOPI Protocol — Document Editing Integration

### What is WOPI

WOPI (Web Application Open Platform Interface) is a REST-based protocol that allows Office Web Apps / Office Online Server to open, edit, and save documents stored in your application.

Reference: https://learn.microsoft.com/en-us/microsoft-365/cloud-storage-partner-program/online/

### WOPI Host Requirements (Your Application)

If you build your own WOPI host (for on-prem or custom storage), you must implement these endpoints:

```
GET  /wopi/files/{fileId}                   → CheckFileInfo
GET  /wopi/files/{fileId}/contents          → GetFile
POST /wopi/files/{fileId}/contents          → PutFile  (save from Office)
POST /wopi/files/{fileId}                   → Lock / RefreshLock / Unlock / GetLock
                                              (action in X-WOPI-Override header)
```

### CheckFileInfo Response (Required Fields)

```json
{
  "BaseFileName": "document.docx",
  "OwnerId": "user@contoso.com",
  "Size": 102400,
  "UserId": "user@contoso.com",
  "Version": "3",
  "UserCanWrite": true,
  "SupportsCoauth": true,
  "SupportedShareUrlTypes": ["ReadWrite"],
  "UserFriendlyName": "John Smith",
  "IsAnonymousUser": false,
  "ReadOnly": false,
  "SupportsUpdate": true,
  "SupportsLocks": true
}
```

### WOPI Lock Operations

```http
# Lock file (Office acquires lock before editing)
POST /wopi/files/{fileId}
X-WOPI-Override: LOCK
X-WOPI-Lock: {lockId}

# Refresh lock (keep alive)
POST /wopi/files/{fileId}
X-WOPI-Override: REFRESH_LOCK
X-WOPI-Lock: {lockId}

# Unlock file (editing ended)
POST /wopi/files/{fileId}
X-WOPI-Override: UNLOCK
X-WOPI-Lock: {lockId}

# Get current lock
POST /wopi/files/{fileId}
X-WOPI-Override: GET_LOCK
# Response header X-WOPI-Lock: {lockId} or empty if unlocked
```

### Using SPO as WOPI Host (No Custom Implementation Needed)

When using SharePoint Online or SharePoint Server as the document store, SPO is the WOPI host. You only call `GetWopiFrameUrl` to get the edit URL and redirect the user's browser to it.

```
Your App → GET GetWopiFrameUrl(action='edit')
              ↓
         Returns: https://word-edit.officeapps.live.com/we/wordeditorframe.aspx?...
              ↓
         Browser opens → Office Web App loads → User edits
              ↓
         Office saves back to SPO automatically via WOPI PutFile
              ↓
         SPO fires webhook/change notification on save
```

### WOPI for On-Premises (Office Online Server)

On-premises deployments require Office Online Server (OOS) running separately:

```powershell
# Configure SharePoint Server to use OOS (run on SP server)
New-SPWOPIBinding -ServerName "oos.contoso.local" -AllowHTTP

# Verify binding
Get-SPWOPIBinding

# Set zone
Set-SPWOPIZone -zone "internal-http"   # or "internal-https"
```

OOS documentation: https://learn.microsoft.com/en-us/officeonlineserver/configure-office-online-server-for-sharepoint-server-2016/configure-office-online-server-for-sharepoint-server-2016

### WOPI Alternatives for Sovereign Deployments (China)

| Solution | WOPI Compatible | Self-Hosted | Office Format Support |
|---|---|---|---|
| **OnlyOffice Docs** | Yes | Yes | docx, xlsx, pptx (high fidelity) |
| **Collabora Online** | Yes | Yes | odt + docx/xlsx via conversion |
| **ThinkFree Office** | Yes | Yes | docx, xlsx, pptx |
| **Office Online Server (OOS)** | Yes (is the server) | Yes | docx, xlsx, pptx (native) |

OnlyOffice + SharePoint integration: https://helpcenter.onlyoffice.com/integration/sharepoint.aspx

---

## 8. Compliance — DLP, Sensitivity Labels, IRM

### Microsoft Purview DLP on SharePoint

When a document is uploaded or shared, Purview DLP engine scans content:

```
Document uploaded to SPO
    │
    ▼
DLP Policy Engine scans (async, within minutes)
    │
    ├── Pattern matched (PII, credit card, classified keywords)?
    │       ├── Block external sharing
    │       ├── Apply sensitivity label automatically
    │       ├── Notify compliance officer
    │       └── Create audit log entry (immutable)
    └── Clean → normal flow, no restriction
```

DLP Docs: https://learn.microsoft.com/en-us/purview/dlp-sensitivity-label-as-condition

### Sensitivity Labels (Microsoft Purview Information Protection)

Labels travel with the document — even when downloaded from SPO:

```
Label tiers (example):
  Public      → No restrictions
  Internal    → Block anonymous sharing
  Confidential → Encrypt + restrict to org users
  Highly Confidential → Encrypt + no download + no print
```

Key behaviors:
- Labels applied in SPO extend to downloads: https://learn.microsoft.com/en-us/purview/sensitivity-labels-sharepoint-extend-permissions
- Auto-labeling based on content: https://learn.microsoft.com/en-us/purview/apply-sensitivity-label-automatically
- Enable labels for SPO/OneDrive: https://learn.microsoft.com/en-us/purview/sensitivity-labels-sharepoint-onedrive-files

### IRM (Information Rights Management) — On-Premises

On-prem equivalent using AD RMS (Active Directory Rights Management Services):

```
Document downloaded from SP Server
    │
    ▼
AD RMS encrypts document with policy
    │
    ├── No copy/paste outside org
    ├── No print
    ├── No screenshot
    ├── Expiry date enforced
    └── Access revocable centrally
```

### Cross-Border Data Policy — What Happens When Document Leaves

| Scenario | SPO Global | SPO with Purview | SP On-Prem |
|---|---|---|---|
| User emails Confidential doc externally | Allowed (warning only, by default) | Blocked by DLP policy | Blocked by AD RMS encryption |
| User downloads and uploads to personal drive | Allowed | Label + encryption travels with doc | RMS encryption travels |
| External user opens shared link | Allowed if sharing enabled | Blocked if label = Confidential | Cannot open without AD RMS license |
| Data stored outside country | Yes (Azure region) | Yes, but encrypted | No — stays on-prem |

---

## 9. Sovereign Countries — China, KSA, Malaysia

### SharePoint Availability Matrix

| Country | SPO Global | 21Vianet SPO | SP Server On-Prem | Notes |
|---|---|---|---|---|
| **China (PRC)** | ❌ Blocked (GFW) | ✅ Available (`.sharepoint.cn`) | ✅ Recommended for regulated | 21Vianet operated; ICP license required for public sites |
| **KSA (Saudi Arabia)** | ✅ Available | N/A | ✅ For strict regulated orgs | MS KSA region (Jeddah/Riyadh) launched 2023; data residency available |
| **Malaysia** | ✅ Available (APAC) | N/A | ✅ Optional | No sovereign restriction; APAC region |
| **UAE** | ✅ Available | N/A | ✅ Optional | Local DC since 2019 |

### China (21Vianet) — Technical Details

```
Endpoints (different from global):
  SharePoint:   https://{tenant}.sharepoint.cn
  Auth:         https://login.partner.microsoftonline.cn/{tenantId}/oauth2/v2.0/token
  Graph API:    https://microsoftgraph.chinacloudapi.cn/v1.0/
  Admin center: https://admin.microsoft.cn

Operated by: 21Vianet (third party, not Microsoft directly)
License required: ICP (Internet Content Provider) for public-facing domains
Compliance: Meets MLPS (Multi-Level Protection Scheme) requirements
Missing features: Some Power Platform connectors, some Graph API endpoints, certain compliance features
```

21Vianet reference: https://learn.microsoft.com/en-us/microsoft-365/admin/services-in-china/services-in-china

### China On-Premises SharePoint Farm (Recommended for Regulated)

**Why on-prem eliminates cross-border risk:**

| Risk | 21Vianet SPO | SP Server On-Prem |
|---|---|---|
| Data stored outside your DC | ✅ (stored in 21Vianet DC) | ❌ (stored on your hardware) |
| Auth calls leaving your network | ✅ (Azure China AD) | ❌ (ADFS on-prem) |
| WOPI/editing calls external | ✅ (21Vianet OOS) | ❌ (OOS on your servers) |
| Microsoft telemetry | ✅ (some sent) | ❌ (can be blocked) |
| Audit logs leaving org | ✅ (stored in 21Vianet) | ❌ (stored locally) |

### KSA Data Residency

Microsoft's KSA (Saudi Arabia) region provides:
- Data stored in Jeddah and Riyadh
- Meets NCA (National Cybersecurity Authority) requirements
- SAMA (Saudi Arabian Monetary Authority) compliance for banking
- Full SPO feature set available

For government / classified workloads: SP Server on-prem in KSA government DC required.

### AWS Regions for Sensitive Countries

| Country | AWS Region | S3 Available | Lambda | Notes |
|---|---|---|---|---|
| China | `cn-north-1` (Beijing) | ✅ | ✅ | ICP required; Cognito unavailable |
| China | `cn-northwest-1` (Ningxia) | ✅ | ✅ | ICP required |
| KSA | `me-south-1` (Bahrain) | ✅ | ✅ | Closest AWS region to KSA |
| Malaysia | `ap-southeast-1` (Singapore) | ✅ | ✅ | |

Note: AWS does not have a native KSA or Malaysia region as of 2026. AWS `me-south-1` (Bahrain) or Azure KSA is used.

---

## 10. Architecture — Enterprise Cloud (AWS Global + SPO)

```
┌──────────────────────────────────────────────────────────────────────────┐
│                    ENTERPRISE CLOUD DESIGN                                │
│               AWS Global + SharePoint Online (Global)                     │
└──────────────────────────────────────────────────────────────────────────┘

  ┌──────────┐    MS OIDC/OAuth2     ┌───────────────────────────┐
  │  Browser │◄─────────────────────►│  Microsoft Entra ID       │
  │  (User)  │    MSAL redirect      │  (Azure AD / Global)      │
  └────┬─────┘                       │  login.microsoftonline.com│
       │                             └──────────────┬────────────┘
       │                                            │ access_token
       │ HTTPS                                      │ (Files.ReadWrite.All)
       ▼                                            ▼
  ┌──────────────────────────────────────────────────────────────┐
  │                  AWS Application Layer                        │
  │                                                              │
  │  ┌───────────────┐   ┌─────────────────┐   ┌─────────────┐ │
  │  │  CloudFront   │──►│  API Gateway    │──►│  Lambda /   │ │
  │  │  CDN          │   │  (Auth Middleware│   │  ECS App    │ │
  │  │               │   │   JWT verify)   │   │             │ │
  │  └───────────────┘   └─────────────────┘   └──────┬──────┘ │
  │                                                    │         │
  │                            ┌───────────────────────┘         │
  │                            │                                  │
  │          ┌─────────────────▼───────────────────────┐         │
  │          │          Document Logic Service          │         │
  │          │                                         │         │
  │          │  1. User requests to open doc           │         │
  │          │  2. Call SPO GetWopiFrameUrl(edit)      │         │
  │          │  3. Return edit URL to browser          │         │
  │          │  4. Browser opens Office Web App        │         │
  │          │  5. Graph webhook fires on each save    │         │
  │          │  6. Debounce (3-min inactivity timer)   │         │
  │          │  7. On timer fire → fetch latest        │         │
  │          │  8. Compare ETag vs stored ETag         │         │
  │          │  9. If newer → upload to S3 (bump ver)  │         │
  │          │  10. Update DynamoDB metadata           │         │
  │          └─────────────────┬───────────────────────┘         │
  └───────────────────────────-┼──────────────────────────────────┘
                               │
          ┌────────────────────┼─────────────────────┐
          │                    │                     │
          ▼                    ▼                     ▼
  ┌──────────────┐   ┌─────────────────────┐  ┌────────────────┐
  │ SharePoint   │   │   AWS S3 Bucket      │  │  DynamoDB      │
  │ Online       │   │   (Document Store)   │  │  (Version Meta)│
  │              │   │                     │  │                │
  │ - Co-author  │   │  docs/{id}/v1/f.docx│  │  doc_id        │
  │ - Versioning │   │  docs/{id}/v2/f.docx│  │  version       │
  │ - WOPI host  │   │  docs/{id}/v3/f.docx│  │  s3_key        │
  │ - Webhooks   │   │  (current = latest) │  │  spo_etag      │
  │ - Editing    │   │                     │  │  modified_at   │
  └──────────────┘   └─────────────────────┘  └────────────────┘

  ┌─────────────────────────────────────────────────────────────┐
  │            Webhook → Debounce → S3 Sync Flow                │
  │                                                             │
  │  SPO/Graph webhook → POST /webhook/graph                    │
  │       │                                                     │
  │       ▼                                                     │
  │  Lambda: reset 3-min timer per docId (DynamoDB TTL)        │
  │       │                                                     │
  │  Timer fires → EventBridge rule triggers Lambda            │
  │       │                                                     │
  │       ▼                                                     │
  │  GET /drives/{id}/items/{id}?$select=eTag,lastModified      │
  │       │                                                     │
  │  Compare: stored_etag == current_etag?                      │
  │       ├── SAME  → skip (no change)                         │
  │       └── DIFF  → download content stream from SPO         │
  │                     │                                       │
  │                     ▼                                       │
  │               PUT s3://bucket/docs/{id}/v{n+1}/file.docx   │
  │                     │                                       │
  │                     ▼                                       │
  │               UPDATE DynamoDB (version++, etag, s3_key)    │
  │                     │                                       │
  │                     ▼                                       │
  │               NOTIFY application callback (optional)       │
  └─────────────────────────────────────────────────────────────┘
```

---

## 11. Architecture — Sovereign/On-Premises (China, KSA)

```
┌──────────────────────────────────────────────────────────────────────────┐
│              SOVEREIGN / ON-PREMISES DESIGN                               │
│          China (21Vianet or On-Prem) | KSA | Malaysia                    │
│              ── All data stays within sovereign boundary ──               │
└──────────────────────────────────────────────────────────────────────────┘

  ┌──────────┐    OIDC/OAuth2       ┌──────────────────────────────────┐
  │  Browser │◄───────────────────►│  Identity Provider (On-Prem)     │
  │  (User)  │                     │                                  │
  └────┬─────┘                     │  China Option A: 21Vianet AD     │
       │                           │    login.partner.microsoftonline.cn
       │                           │  China Option B: ADFS (on-prem)  │
       │                           │  China Option C: Keycloak        │
       │                           │  KSA: Entra ID / ADFS            │
       │                           └──────────────────────────────────┘
       │
       │  ─────────────── SOVEREIGN DATA BOUNDARY ─────────────────────
       ▼
  ┌──────────────────────────────────────────────────────────────────┐
  │                    Application Layer                              │
  │                                                                   │
  │  ┌─────────────┐   ┌──────────────────┐   ┌─────────────────┐   │
  │  │  Nginx /    │──►│  App Servers     │──►│  Document       │   │
  │  │  Load       │   │  (K8s / VMs)     │   │  Logic Service  │   │
  │  │  Balancer   │   │                  │   │                 │   │
  │  └─────────────┘   └──────────────────┘   └────────┬────────┘   │
  └────────────────────────────────────────────────────-┼────────────┘
                                                         │
       ┌─────────────────────────────────────────────────┘
       │
  ┌────▼──────────────────────────────────────────────────────────────┐
  │                   Storage & Collaboration Layer                    │
  │                                                                    │
  │  ┌──────────────────────┐       ┌──────────────────────────────┐  │
  │  │  SharePoint Server   │       │  Object Storage              │  │
  │  │  2019 / SE Farm      │       │                              │  │
  │  │                      │       │  China:  Alibaba OSS         │  │
  │  │  WFE Servers (x2 HA) │       │          Azure China Blob    │  │
  │  │  App Server          │       │          MinIO (self-hosted) │  │
  │  │  SQL Server (Always- │       │  KSA:    AWS me-south-1 S3   │  │
  │  │    On, Primary +     │       │          Azure KSA Blob      │  │
  │  │    Secondary)        │       │  MYS:    AWS ap-southeast-1  │  │
  │  │                      │       │  OnPrem: MinIO / Ceph        │  │
  │  │  + Office Online     │       │                              │  │
  │  │    Server (OOS) for  │       │  Versioned prefix structure: │  │
  │  │    browser editing   │       │  /docs/{id}/v1/file.docx     │  │
  │  │                      │       │  /docs/{id}/v2/file.docx     │  │
  │  │  All within your DC  │       │  /docs/{id}/v{n}/file.docx   │  │
  │  └──────────────────────┘       └──────────────────────────────┘  │
  │                                                                    │
  │  ┌──────────────────────┐       ┌──────────────────────────────┐  │
  │  │  SQL Server /        │       │  DLP / Compliance            │  │
  │  │  PostgreSQL          │       │                              │  │
  │  │  (Version metadata,  │       │  On-prem:  AD RMS            │  │
  │  │   audit logs)        │       │  Custom:   Label scanner     │  │
  │  │                      │       │  Audit:    WORM log storage  │  │
  │  └──────────────────────┘       └──────────────────────────────┘  │
  └────────────────────────────────────────────────────────────────────┘

  ┌─────────────────────────────────────────────────────────────────┐
  │           China-Specific: SPO Replacement Stack                 │
  │           (when 21Vianet is not acceptable)                     │
  │                                                                 │
  │  Office Web Apps (OOS)    →  OnlyOffice Docs (self-hosted)     │
  │  SharePoint Document Lib  →  MinIO + custom metadata DB        │
  │  SharePoint Versioning    →  Custom version table (SQL)        │
  │  SP Webhooks              →  Application-level callbacks        │
  │  Azure AD                 →  Keycloak + OIDC (self-hosted)     │
  │  Microsoft Graph          →  Custom REST API (your own)        │
  └─────────────────────────────────────────────────────────────────┘
```

### On-Prem Edit Session Flow with Remote Event Receiver

```
ENFORCE CHECK-OUT LIBRARY SETTING = YES

User opens doc in browser (via OOS WOPI)
    │
    ▼
SP Server: auto Check-Out issued → exclusive lock
    │
    ▼
User edits (Office Online / OOS renders doc)
    │
    ▼
User saves and closes
    │
    ▼
SP Server: Check-In dialog → user confirms → major version created
    │
    ▼
Remote Event Receiver fires: ItemCheckedIn event
    │
    ▼
Your application handler:
  1. Fetch latest file from SP Server
  2. Compare stored version vs SP version
  3. If newer → upload to MinIO / S3-compatible storage
  4. Update version metadata table
  5. Notify application (webhook callback or message queue)
```

---

## 12. S3 Sync Pattern — Version Bumping Flow

### Core Principle

```
S3 is the system of record for long-term storage.
SPO / SP Server is the ephemeral editing surface.
Application syncs from SPO → S3 after editing completes.
```

### DynamoDB Version Metadata Schema

```json
{
  "doc_id": "uuid-of-document",
  "current_version": 5,
  "s3_key": "docs/uuid/v5/proposal.docx",
  "spo_etag": "\"{GUID},{version-number}\"",
  "spo_item_id": "42",
  "spo_drive_id": "driveId",
  "spo_site_id": "siteId",
  "filename": "proposal.docx",
  "last_modified": "2026-09-06T10:30:00Z",
  "last_synced_at": "2026-09-06T10:33:00Z",
  "versions": [
    { "version": 1, "s3_key": "docs/uuid/v1/proposal.docx", "modified": "...", "modified_by": "user@..." },
    { "version": 2, "s3_key": "docs/uuid/v2/proposal.docx", "modified": "...", "modified_by": "user@..." }
  ]
}
```

### Version Bump Logic (Pseudocode)

```python
def sync_document_to_s3(doc_id: str):
    # 1. Get stored metadata
    stored = dynamodb.get_item(doc_id)

    # 2. Get current SPO metadata
    spo_item = graph.get(
        f"/sites/{siteId}/drives/{driveId}/items/{stored['spo_item_id']}"
        "?$select=eTag,lastModifiedDateTime,name"
    )

    # 3. Compare ETags
    if spo_item['eTag'] == stored['spo_etag']:
        logger.info("No change — skip sync")
        return

    # 4. Download from SPO
    content = graph.get(
        f"/sites/{siteId}/drives/{driveId}/items/{stored['spo_item_id']}/content"
    )

    # 5. Bump version
    new_version = stored['current_version'] + 1
    s3_key = f"docs/{doc_id}/v{new_version}/{stored['filename']}"

    # 6. Upload to S3
    s3.put_object(
        Bucket="documents-bucket",
        Key=s3_key,
        Body=content,
        Metadata={
            "doc_id": doc_id,
            "version": str(new_version),
            "spo_etag": spo_item['eTag'],
            "modified": spo_item['lastModifiedDateTime']
        }
    )

    # 7. Update metadata store
    dynamodb.update_item(doc_id, {
        "current_version": new_version,
        "s3_key": s3_key,
        "spo_etag": spo_item['eTag'],
        "last_synced_at": datetime.utcnow().isoformat()
    })

    # 8. Notify application
    notify_application(doc_id=doc_id, new_version=new_version, s3_key=s3_key)
```

---

## 13. API Limits & Throttling

| Limit | Value | Notes |
|---|---|---|
| SPO Webhook max expiry | 180 days | Must renew before expiry |
| Graph subscription max expiry (driveItem) | 30 days | Auto-renew required |
| Graph subscription max expiry (list) | 30 days | |
| Webhook delivery delay | 0–5 minutes | Not guaranteed real-time |
| Webhook retry on failure | 5 attempts | Dropped after 5 failures |
| Webhook validation response time | 5 seconds | Must echo validationtoken within 5s |
| Small file upload (Graph) | 4 MB max | Use upload session above 4 MB |
| Large file upload session | 250 GB max | |
| Upload session validity | 24 hours | Must complete upload within 24h |
| Max concurrent upload sessions | 5 per app per drive | |
| Graph API throttle | 10,000 req / 10 min / app | Per-application throttle |
| SP REST API throttle | 200 req/sec per site | Per-site throttle |
| Co-authoring version period | 30 min (default) | Configurable via PowerShell on-prem |
| File lock detection | No native API | Attempt write → 423 Locked response = file in use |
| Max file size (SP library) | 250 GB (SPO) | Configurable on-prem |
| Max list items per library | 30 million items | Indexed columns recommended at scale |

---

## 14. Official Documentation Links

### Core API References

| Resource | URL |
|---|---|
| Graph API — Files & SharePoint Overview | https://learn.microsoft.com/en-us/graph/api/resources/onedrive |
| Graph API — driveItem Resource | https://learn.microsoft.com/en-us/graph/api/resources/driveitem |
| Graph API — site Resource | https://learn.microsoft.com/en-us/graph/api/resources/site |
| Graph API — Subscription Resource | https://learn.microsoft.com/en-us/graph/api/resources/subscription |
| Graph API — Change Notifications Overview | https://learn.microsoft.com/en-us/graph/api/resources/change-notifications-api-overview |
| Graph API — Webhooks Delivery | https://learn.microsoft.com/en-us/graph/change-notifications-delivery-webhooks |
| Graph API — Rich Notifications with Resource Data | https://learn.microsoft.com/en-us/graph/change-notifications-with-resource-data |
| Graph API — Socket.IO | https://learn.microsoft.com/en-us/graph/api/subscriptions-socketio |
| SP REST API — Getting Started | https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service |
| SP REST API — Files & Folders | https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest |
| SP REST API — CRUD Operations | https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/complete-basic-operations-using-sharepoint-rest-endpoints |
| SP REST API — via Graph v2 endpoints | https://learn.microsoft.com/en-us/sharepoint/dev/apis/sharepoint-rest-graph |

### Webhooks

| Resource | URL |
|---|---|
| SP Webhooks — Overview | https://learn.microsoft.com/en-us/sharepoint/dev/apis/webhooks/overview-sharepoint-webhooks |
| SP Webhooks — List Webhooks | https://learn.microsoft.com/en-us/sharepoint/dev/apis/webhooks/lists/overview-sharepoint-list-webhooks |
| SP Webhooks — Getting Started | https://learn.microsoft.com/en-us/sharepoint/dev/apis/webhooks/get-started-webhooks |
| SP Webhooks — Reference Implementation | https://learn.microsoft.com/en-us/sharepoint/dev/apis/webhooks/webhooks-reference-implementation |
| RER to Webhooks Migration | https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins-modernize/from-remote-event-receivers-to-webhooks |
| SP Embedded + Webhooks | https://learn.microsoft.com/en-us/sharepoint/dev/embedded/build/respond-to-changes-webhooks |

### Co-authoring & WOPI

| Resource | URL |
|---|---|
| Co-authoring Overview (SP Server) | https://learn.microsoft.com/en-us/sharepoint/governance/co-authoring-overview |
| Co-authoring Version Period Config | https://learn.microsoft.com/en-us/sharepoint/governance/configure-the-co-authoring-versioning-period |
| WOPI Protocol Overview | https://learn.microsoft.com/en-us/microsoft-365/cloud-storage-partner-program/online/ |
| WOPI Co-authoring Extensions | https://learn.microsoft.com/en-us/microsoft-365/cloud-storage-partner-program/plus/coauthoring/coauthoring-protocol-about |
| Co-author in M365 Web | https://learn.microsoft.com/en-us/microsoft-365/cloud-storage-partner-program/online/scenarios/coauth |
| MS-FSSHTTP Protocol Spec | https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fsshttp/6d078cbe-2651-43a0-b460-685ac3f14c45 |
| MS-FSSHTTP — Join Coauthoring Session | https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fsshttp/87ca4a76-f156-4cc7-b2f2-f7624d5a9887 |
| MS-FSSHTTP — Refresh Coauthoring Session | https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fsshttp/2eca5d82-69b7-4b86-9a59-4809a2c5e342 |

### Office Online Server (On-Premises WOPI Host)

| Resource | URL |
|---|---|
| OOS Overview | https://learn.microsoft.com/en-us/officeonlineserver/office-online-server-overview |
| OOS Configure for SharePoint Server 2016+ | https://learn.microsoft.com/en-us/officeonlineserver/configure-office-online-server-for-sharepoint-server-2016/configure-office-online-server-for-sharepoint-server-2016 |
| OnlyOffice + SharePoint Integration | https://helpcenter.onlyoffice.com/integration/sharepoint.aspx |

### Compliance & DLP

| Resource | URL |
|---|---|
| Sensitivity Labels Overview | https://learn.microsoft.com/en-us/purview/sensitivity-labels |
| Enable Labels for SPO/OneDrive | https://learn.microsoft.com/en-us/purview/sensitivity-labels-sharepoint-onedrive-files |
| Extend Labels to Downloads | https://learn.microsoft.com/en-us/purview/sensitivity-labels-sharepoint-extend-permissions |
| DLP with Sensitivity Labels | https://learn.microsoft.com/en-us/purview/dlp-sensitivity-label-as-condition |
| Auto-apply Sensitivity Labels | https://learn.microsoft.com/en-us/purview/apply-sensitivity-label-automatically |
| Block External Sharing via DLP | https://learn.microsoft.com/en-us/purview/dlp-create-policy-spo-odb-external |

### China (21Vianet) & Sovereign

| Resource | URL |
|---|---|
| Office 365 operated by 21Vianet | https://learn.microsoft.com/en-us/microsoft-365/admin/services-in-china/services-in-china |
| 21Vianet Setup Guide | https://support.microsoft.com/en-us/office/set-up-your-organization-for-office-365-operated-by-21vianet-5c2d5b59-314f-4fdd-a4ea-2ee01a574029 |

---

*Content was compiled and paraphrased for compliance with licensing restrictions. All referenced documentation is from official Microsoft Learn (learn.microsoft.com) sources.*
