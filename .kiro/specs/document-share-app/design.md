# Design Document: Document Share App

## Overview

Document Share App is a full-stack document sharing and co-authoring platform that uses SharePoint (Online or On-Premises) as the collaborative editing surface and S3-compatible object storage as the versioned, durable system of record. The platform must operate across three deployment topologies without codebase forks: Cloud Global (AWS + SharePoint Online + Entra ID), Sovereign/On-Premises (SharePoint Server 2019/SE + MinIO or Alibaba OSS + Keycloak or ADFS), and Local Development (Docker Compose with stubs).

The core document lifecycle is: upload → browse/search → co-author via WOPI → webhook notification → debounce → ETag-based sync to object storage → version history available for download.

### Key Design Decisions

**SharePoint as WOPI host, not a custom WOPI implementation.** Because SharePoint already implements the WOPI host interface and manages MS-FSSHTTP co-authoring internally, the application only needs to call `GetWopiFrameUrl` to redirect the browser to the Office Web App editor. This eliminates the complexity of implementing WOPI lock management, CellSubrequest merging, and co-authoring state in application code.

**Driver pattern for all external adapters.** Every external integration (Auth, SharePoint, Storage, MetadataStore) is expressed as an interface with multiple driver implementations. Environment variables select the driver at startup. This enables a single deployment pipeline to target all topologies.

**Debounce + ETag idempotence for sync.** Rather than trying to detect "editor closed" (which SharePoint does not expose natively), the platform uses a per-document debounce timer that resets on every webhook notification. When the timer fires without reset, the Sync_Service compares the current ETag against the stored ETag and only writes a new version if they differ.

**Append-only audit log as a first-class concern.** Audit log writes are blocking — if the write fails after retries, the triggering operation is aborted. This ensures the audit log is always consistent with the operations that were committed.

---

## Architecture

### Component Overview

```
┌─────────────────────────────────────────────────────────────────────┐
│                         Browser (SPA)                                │
│  - MSAL.js / OIDC client for token acquisition (PKCE flow)          │
│  - Document browser, version history, permission management UI      │
└──────────────────────────────┬──────────────────────────────────────┘
                               │ HTTPS / Bearer token
                               ▼
┌─────────────────────────────────────────────────────────────────────┐
│                        API_Service (REST)                            │
│                                                                      │
│  ┌──────────────┐  ┌──────────────┐  ┌─────────────────────────┐   │
│  │ Auth_Service │  │ Webhook_     │  │ Sync_Service            │   │
│  │ (OIDC/OAuth2)│  │ Handler      │  │ (ETag-based sync)       │   │
│  └──────────────┘  └──────────────┘  └─────────────────────────┘   │
│                                                                      │
│  ┌──────────────────────────────────────────────────────────────┐   │
│  │               Adapter Layer                                  │   │
│  │  SharePoint_Adapter  |  Storage_Adapter  |  Metadata_Store  │   │
│  └──────────────────────────────────────────────────────────────┘   │
└──────────────────────────────────────────────────────────────────────┘
         │                     │                     │
         ▼                     ▼                     ▼
┌─────────────────┐  ┌──────────────────┐  ┌─────────────────────┐
│  SharePoint     │  │  Object Storage  │  │  Metadata Store     │
│  (SPO or        │  │  (S3 / MinIO /   │  │  (DynamoDB /        │
│   SP Server)    │  │   Alibaba OSS)   │  │   PostgreSQL /      │
│                 │  │                  │  │   SQL Server)       │
│  WOPI host      │  │  Versioned docs  │  │  Documents,         │
│  Co-authoring   │  │  Audit logs      │  │  versions,          │
│  Webhooks       │  │  (WORM)          │  │  subscriptions,     │
└─────────────────┘  └──────────────────┘  │  permissions,       │
                                           │  audit events       │
                                           └─────────────────────┘
```

### Component Responsibilities

| Component | Responsibility |
|---|---|
| **API_Service** | REST API endpoint routing, request validation, session middleware, WOPI URL generation, permission enforcement, confirmation token issuance |
| **Auth_Service** | OIDC token validation, session lifecycle, silent refresh, multi-provider driver |
| **Webhook_Handler** | Validation handshake, async notification processing, Change Log API calls, Debounce_Timer reset, clientState verification |
| **Sync_Service** | ETag comparison, SharePoint content streaming, Storage_Adapter write, Metadata_Store conditional update, version record creation |
| **SharePoint_Adapter** | Unified interface over Graph API (global/China) and SharePoint REST API (on-prem); file CRUD, upload sessions, webhooks, WOPI URL, Change Log, permissions |
| **Storage_Adapter** | Unified interface over AWS S3 / MinIO / Alibaba OSS; versioned put, get, list, presigned URL |
| **Metadata_Store** | Unified interface over DynamoDB / PostgreSQL / SQL Server; document CRUD, version records, subscription records, permission records, audit log |

### Deployment Topology: Cloud Global

```
Internet → CloudFront → API Gateway → Lambda / ECS (API_Service)
                                             │
                    ┌────────────────────────┼───────────────────┐
                    ▼                        ▼                   ▼
             SharePoint Online         AWS S3 (docs)       DynamoDB
             (Entra ID auth)           AWS S3 (audit)      (metadata)
             graph.microsoft.com       Object Lock WORM
                    │
                    └── Graph Subscriptions → API Gateway /webhook endpoint
                         EventBridge → Lambda (debounce timer)
```

### Deployment Topology: Sovereign / On-Premises

```
Clients → Nginx (LB) → App Servers (K8s / VMs) (API_Service)
                               │
          ┌────────────────────┼─────────────────────┐
          ▼                    ▼                     ▼
   SharePoint Server     MinIO / Alibaba OSS    PostgreSQL / SQL Server
   2019 / SE Farm        (docs + audit WORM)    (metadata + audit log)
   OOS / OnlyOffice
   SP REST + RER
```

### Deployment Topology: Local Development

```
Docker Compose:
  api_service (hot-reload)
  keycloak (dev mode, OIDC stub)
  minio (S3-compatible)
  postgres (metadata store)
  wopi_stub (mock WOPI server)
  sharepoint_mock (optional MSGraph mock for offline dev)
```

### Webhook → Debounce → Sync Flow

```
SPO / Graph webhook POST /webhook/graph
    │
    ▼
Webhook_Handler: validate clientState, enqueue async task, return HTTP 200 < 5s
    │
    ▼ (async)
Webhook_Handler: call Change Log API with stored Change_Token
    ├── On success: update Change_Token, reset Debounce_Timer (3 min) for each changed doc
    └── On failure (3 retries): log change-log-fetch-failed, preserve Change_Token, discard
    │
Debounce_Timer fires (no notification within 3 minutes)
    │
    ▼
Sync_Service: GET SharePoint item ETag + lastModified
    ├── ETag == stored spo_etag → record no-change, done
    └── ETag differs
          │
          ▼
        Stream content from SharePoint
        PUT to Storage_Adapter at docs/{doc_id}/v{n+1}/{filename}
          │
          ▼
        Conditional write to Metadata_Store (condition: spo_etag == stored_etag)
          ├── Win (first writer): increment version, update s3_key, spo_etag, last_synced_at
          └── Lose (concurrent write): silently discard
```

---

## Components and Interfaces

### Auth_Service Interface

```typescript
interface AuthDriver {
  // Validate an inbound Bearer token; returns parsed claims
  validateToken(token: string): Promise<Claims>;

  // Build OIDC authorisation redirect URL (PKCE)
  buildAuthorizationUrl(state: string, codeVerifier: string): string;

  // Exchange authorisation code for tokens
  exchangeCode(code: string, codeVerifier: string): Promise<TokenSet>;

  // Silent token refresh
  refreshTokens(refreshToken: string): Promise<TokenSet>;

  // Revoke refresh token with IdP
  revokeToken(refreshToken: string): Promise<void>;
}

// Driver implementations: EntraGlobalDriver, EntraChinaDriver, AdfsDriver, KeycloakDriver
```

### SharePoint_Adapter Interface

```typescript
interface SharePointDriver {
  // File operations
  uploadSmall(libraryRef: LibraryRef, filename: string, content: Buffer): Promise<SPItem>;
  createUploadSession(libraryRef: LibraryRef, filename: string): Promise<UploadSession>;
  uploadChunk(session: UploadSession, chunk: Buffer, range: ByteRange): Promise<UploadProgress>;
  getItem(itemRef: ItemRef): Promise<SPItem>;
  downloadItem(itemRef: ItemRef): Promise<NodeJS.ReadableStream>;
  deleteItem(itemRef: ItemRef): Promise<void>;

  // Metadata
  getETag(itemRef: ItemRef): Promise<string>;
  listChildren(folderRef: FolderRef, pageToken?: string): Promise<PagedResult<SPItem>>;
  search(libraryRef: LibraryRef, query: string): Promise<SPItem[]>;

  // WOPI
  getWopiFrameUrl(itemRef: ItemRef, action: 'edit' | 'view'): Promise<string>;

  // Webhooks / subscriptions
  registerSubscription(libraryRef: LibraryRef, notificationUrl: string, clientState: string): Promise<Subscription>;
  renewSubscription(subId: string, newExpiry: Date): Promise<Subscription>;
  deleteSubscription(subId: string): Promise<void>;
  getChanges(libraryRef: LibraryRef, changeToken: string): Promise<ChangeResult>;

  // Permissions
  grantPermission(itemRef: ItemRef, email: string, role: 'read' | 'write'): Promise<void>;
  revokePermission(itemRef: ItemRef, permissionId: string): Promise<void>;
  listPermissions(itemRef: ItemRef): Promise<SPPermission[]>;
}

// Driver implementations:
//   GraphGlobalDriver  — graph.microsoft.com
//   GraphChinaDriver   — microsoftgraph.chinacloudapi.cn
//   SharePointRestOnPremDriver — SP Server REST API
```

### Storage_Adapter Interface

```typescript
interface StorageDriver {
  // Write an immutable version object
  putObject(key: string, stream: NodeJS.ReadableStream, metadata: ObjectMetadata): Promise<void>;

  // Stream an object to the caller
  getObject(key: string): Promise<NodeJS.ReadableStream>;

  // Check object exists
  headObject(key: string): Promise<ObjectHead | null>;

  // Paginated list under a prefix
  listObjects(prefix: string, pageToken?: string): Promise<PagedResult<ObjectInfo>>;
}

// Key format: docs/{doc_id}/v{version}/{filename}
//             audit/{year}/{month}/{day}/{entry_id}.json
// Driver implementations: S3AwsDriver, S3MinioDriver, S3OssDriver
```

### Metadata_Store Interface

```typescript
interface MetadataDriver {
  // Document records
  createDocument(doc: DocumentRecord): Promise<void>;
  getDocument(docId: string): Promise<DocumentRecord | null>;
  updateDocument(docId: string, patch: Partial<DocumentRecord>, condition?: ConditionalWrite): Promise<void>;
  listDocuments(filter: DocumentFilter, page: PageCursor): Promise<PagedResult<DocumentRecord>>;

  // Version records
  createVersionRecord(record: VersionRecord): Promise<void>;
  getVersionRecord(docId: string, version: number): Promise<VersionRecord | null>;
  listVersionRecords(docId: string, page: PageCursor): Promise<PagedResult<VersionRecord>>;

  // Permission records
  upsertPermission(perm: PermissionRecord): Promise<void>;
  deletePermission(docId: string, userId: string): Promise<void>;
  getPermission(docId: string, userId: string): Promise<PermissionRecord | null>;
  listPermissions(docId: string): Promise<PermissionRecord[]>;

  // Subscription records
  upsertSubscription(sub: SubscriptionRecord): Promise<void>;
  deleteSubscription(subId: string): Promise<void>;
  getSubscription(subId: string): Promise<SubscriptionRecord | null>;
  listSubscriptionsExpiringBefore(cutoff: Date): Promise<SubscriptionRecord[]>;

  // Debounce timers (DynamoDB TTL items; EventBridge targets on cloud)
  setDebounceTimer(docId: string, firesAt: Date): Promise<void>;
  clearDebounceTimer(docId: string): Promise<void>;

  // Audit log (append-only)
  appendAuditEntry(entry: AuditEntry): Promise<void>;
}

// Driver implementations: DynamoDbDriver, PostgresDriver, SqlServerDriver
```

---

## Data Models

### DocumentRecord

```typescript
interface DocumentRecord {
  doc_id: string;                 // UUID, primary key
  filename: string;               // Original filename
  file_size_bytes: number;
  spo_item_id: string;            // SharePoint driveItem ID
  spo_drive_id: string;
  spo_site_id: string;
  current_version: number;        // Monotonically increasing integer, starts at 1
  s3_key: string;                 // Key for current version in object storage
  spo_etag: string;               // ETag from SharePoint, used for sync idempotence
  upload_timestamp: string;       // ISO 8601 UTC millisecond precision
  uploader_user_id: string;
  last_modified_at: string;       // ISO 8601 UTC
  last_modified_by: string;
  last_synced_at: string;         // ISO 8601 UTC
  sensitivity_label: SensitivityLabel; // "Public" | "Internal" | "Confidential" | "Highly_Confidential"
  highly_confidential_edit_override: boolean;
  status: 'active' | 'deleted';
  change_token: string;           // SharePoint Change Log cursor for this document's library
  library_id: string;             // The Document_Library this document belongs to
  created_at: string;             // ISO 8601 UTC
}

type SensitivityLabel = 'Public' | 'Internal' | 'Confidential' | 'Highly_Confidential';
```

### VersionRecord

```typescript
interface VersionRecord {
  doc_id: string;                 // Partition key
  version_number: number;         // Sort key
  s3_key: string;                 // Immutable object key in Storage_Adapter
  filename: string;               // Filename at sync time (may differ from current if renamed)
  file_size_bytes: number;
  spo_etag: string;               // SharePoint ETag that triggered this version — UNIQUE per doc
  sync_timestamp: string;         // ISO 8601 UTC when sync completed
  synced_by_user_id: string;      // User who last modified in SharePoint (from Change Log)
  sharepoint_version_label: string; // e.g. "3.0"
}
```

### PermissionRecord

```typescript
interface PermissionRecord {
  doc_id: string;
  user_id: string;
  user_email: string;
  display_name: string;
  permission_level: 'view' | 'edit' | 'owner';
  spo_permission_id: string;      // Graph API permission ID for revocation
  granted_at: string;             // ISO 8601 UTC
  granted_by: string;
}
```

### SubscriptionRecord

```typescript
interface SubscriptionRecord {
  subscription_id: string;        // Primary key (Graph or SPO webhook ID)
  library_id: string;             // Document_Library this subscription covers
  subscription_type: 'graph' | 'spo-rest';
  notification_url: string;
  client_state_hash: string;      // HMAC-SHA256 of the clientState secret (not the secret itself)
  client_state_secret: string;    // Stored encrypted; used for validation
  expiry: string;                 // ISO 8601 UTC
  registered_at: string;
  last_renewed_at: string;
  site_id: string;
  drive_id: string;
}
```

### AuditEntry

```typescript
interface AuditEntry {
  entry_id: string;               // UUID
  event_type: AuditEventType;
  doc_id: string | null;          // null for system-level events
  user_id: string;                // user identity or "system"
  user_email: string | null;
  timestamp: string;              // ISO 8601 UTC millisecond precision
  client_ip: string | null;
  outcome: 'success' | 'failure' | 'denied';
  details: Record<string, unknown>; // Event-specific fields (old_label, new_label, version, etc.)
}

type AuditEventType =
  | 'document.upload'
  | 'document.download'
  | 'document.edit_session.open'
  | 'document.edit_session.close'
  | 'document.sync.success'
  | 'document.sync.failure'
  | 'document.sync.no_change'
  | 'document.delete'
  | 'document.search'
  | 'permission.grant'
  | 'permission.revoke'
  | 'permission.denied'
  | 'webhook.validation_mismatch'
  | 'webhook.change_log_fetch_failed'
  | 'subscription.registered'
  | 'subscription.renewed'
  | 'subscription.re_registered'
  | 'subscription.registration_failed'
  | 'sensitivity_label.changed'
  | 'version_history.list';
```

### Debounce_Timer (DynamoDB TTL / Redis key)

```typescript
interface DebounceTimer {
  doc_id: string;
  fires_at_epoch_seconds: number; // DynamoDB TTL attribute; or Redis EXPIREAT key
  library_id: string;
  subscription_id: string;
}
```

---

## API Design

### Authentication Endpoints

```
GET  /auth/login          → Redirect to OIDC provider (PKCE flow)
GET  /auth/callback       → Exchange code, establish session
POST /auth/logout         → Revoke refresh token, clear session
GET  /auth/session        → Return current session claims
```

### Document Endpoints

```
GET    /documents
  Query: page_cursor, page_size (1–200), folder_path, extension, sensitivity_label,
         modified_after, modified_before, include_deleted
  Response: { items: DocumentListItem[], next_cursor, total_count }

POST   /documents/upload
  Body: multipart/form-data (file, sensitivity_label?, target_folder?)
  Response 201: { doc_id, filename, size, upload_timestamp, spo_item_url, version }
  Errors: 400 (empty file), 403 (no edit permission), 500 (partial failure)

GET    /documents/search
  Query: q (1–1000 chars), page_cursor, page_size
  Response: { items: DocumentListItem[], next_cursor }

GET    /documents/:docId
  Response: DocumentDetail (includes confirmation_token for delete, valid 15 min)

DELETE /documents/:docId
  Body: { confirmation_token }
  Response 204
  Errors: 403, 409 (token absent/expired/invalid)

GET    /documents/:docId/versions
  Query: page_cursor, page_size (1–200)
  Response: { items: VersionRecord[], next_cursor, total_count, earliest_version_at }

GET    /documents/:docId/versions/:versionNumber/download
  Response: file stream with Content-Disposition: attachment

GET    /documents/:docId/download
  Response: file stream (current version)

GET    /documents/:docId/wopi-url
  Query: action (edit | view)
  Response: { wopi_frame_url }
  Errors: 403, 502 (WOPI_URL_FETCH_FAILED)
```

### Permission Endpoints

```
GET    /documents/:docId/permissions          → 403 if not owner
POST   /documents/:docId/permissions          → Grant permission
  Body: { user_email, permission_level }
DELETE /documents/:docId/permissions/:userId  → Revoke permission
```

### Sensitivity Label Endpoint

```
PATCH  /documents/:docId/sensitivity-label
  Body: { sensitivity_label, highly_confidential_edit_override? }
  Response 200
  Errors: 403 (not owner), 500 (rollback on partial failure)
```

### Webhook Endpoints

```
GET  /webhook/graph     → Validation handshake (validationtoken echo)
POST /webhook/graph     → Graph subscription notifications
GET  /webhook/sharepoint → Validation handshake
POST /webhook/sharepoint → SPO REST webhook notifications
```

### Operational Endpoints

```
GET  /health                    → { status, adapters: { sharepoint, storage, metadata, auth } }
GET  /metrics                   → Prometheus text format
GET  /compliance/data-residency → { adapters: [{ name, endpoint, data_residency_zone }] }
```

### Error Response Schema

```json
{
  "error_code": "DOCUMENT_NOT_FOUND",
  "message": "Document with ID abc123 does not exist",
  "request_id": "550e8400-e29b-41d4-a716-446655440000",
  "timestamp": "2026-09-06T10:30:00.000Z"
}
```

### DocumentListItem Shape

```typescript
interface DocumentListItem {
  doc_id: string;
  filename: string;
  file_size_bytes: number;
  last_modified_at: string;
  last_modified_by_display_name: string;
  current_version: number;
  effective_permission: 'view' | 'edit' | 'owner';
  sensitivity_label: SensitivityLabel;
  wopi_view_url?: string;          // present for Office-compatible files
}
```

---

## Technology Stack

| Layer | Cloud Global | Sovereign / On-Prem | Local Dev | Rationale |
|---|---|---|---|---|
| **Runtime** | Node.js 22 on Lambda / ECS | Node.js 22 on K8s | Node.js 22 (Docker) | Single codebase; async I/O suits streaming large files |
| **API framework** | Fastify | Fastify | Fastify | Low overhead; native streaming; schema validation |
| **Auth (IdP)** | Microsoft Entra ID | Keycloak / ADFS | Keycloak (dev mode) | OIDC-standard; PKCE required by requirements |
| **SharePoint** | Graph API (global) | Graph (21Vianet) / SP REST | SP REST mock | Adapter pattern isolates API differences |
| **Object Storage** | AWS S3 | MinIO / Alibaba OSS | MinIO | S3-compatible API across all drivers |
| **Metadata Store** | DynamoDB | PostgreSQL / SQL Server | PostgreSQL | Adapter pattern; DynamoDB for scale, PG for sovereign |
| **Debounce Timer** | EventBridge Scheduler | BullMQ (Redis) on K8s | BullMQ (Redis) | DynamoDB TTL + EventBridge for cloud; Redis for on-prem |
| **IaC (cloud)** | AWS CDK (TypeScript) | — | — | Type-safe; co-located with app code |
| **IaC (sovereign)** | — | Helm + Docker Compose | Docker Compose | Kubernetes-native for sovereign; Compose for dev |
| **Observability** | CloudWatch + X-Ray | Prometheus + Grafana | stdout JSON | Prometheus /metrics endpoint on all topologies |
| **PBT framework** | fast-check (TypeScript) | fast-check | fast-check | Mature; integrates with Jest/Vitest; arbitrary generators |

---

## Security Design

### Token Management

- Access tokens are never stored on the client beyond the session lifetime; they are held server-side in the session store.
- The server-side session is referenced by a secure, HttpOnly, SameSite=Strict cookie.
- Refresh tokens are stored encrypted at rest in the session store (AES-256-GCM with a KMS-managed key on cloud; Keycloak-encrypted storage on sovereign).
- Silent refresh happens 5 minutes before expiry using a background timer per session (Requirement 1.3).
- On sign-out, the refresh token is revoked at the IdP before the session is cleared.

### Request Identity Propagation

Every inbound request is annotated with a UUID `X-Request-ID` at the API gateway / Nginx ingress. All internal service calls and log entries carry this ID, satisfying Requirement 23.5.

### Webhook clientState Secret

The `clientState` field sent in subscription registration is a 32-byte cryptographically random secret, unique per subscription. It is stored in the Metadata_Store as both its raw value (for outbound renewal) and an HMAC-SHA256 hash (for quick inbound validation). Inbound webhook notifications are accepted only when `HMAC-SHA256(incoming_clientState) == stored_hash` — never raw string comparison to avoid timing attacks.

### Sensitivity Label Enforcement

| Label | Anonymous link | WOPI URL | Object storage |
|---|---|---|---|
| Public | ✅ Allowed | edit + view | Unrestricted |
| Internal | ❌ Blocked | edit + view | Unrestricted |
| Confidential | ❌ Blocked | edit + view | Unrestricted |
| Highly_Confidential | ❌ Blocked | view-only (unless `highly_confidential_edit_override`) | Unrestricted |

### Audit Log WORM

**Cloud Global:** Objects written to the audit S3 bucket are subject to Object Lock in Compliance mode (7-year retention). The bucket has versioning enabled, Block Public Access, SSE-KMS, and no `s3:DeleteObject` in any IAM policy attached to application roles.

**Sovereign / On-Prem:** The audit log table is written via an insert-only database user (no UPDATE or DELETE grants). PostgreSQL row-level security prevents modification by any role except the service account that writes new rows. The Helm chart `audit.wormStorageClass` controls the storage backend for WORM compliance.

**Local Dev:** Audit entries are written to stdout as structured JSON (Requirement 15.6). No WORM enforcement.

### Data Residency Guard (Sovereign)

The `GraphChinaDriver` and all sovereign auth drivers include a URL guard that checks every outbound request URL against a blocklist of global Microsoft endpoints (`graph.microsoft.com`, `login.microsoftonline.com`, `*.sharepoint.com`). If matched, the call is rejected with an error and logged — no request is made. This implements Requirement 17.4 and 20.4.

### Confirmation Token for Deletion

On `GET /documents/:docId`, the API_Service generates a HMAC-SHA256 token embedding `doc_id + expires_at` (15 minutes from generation), signed with a secret key. The `DELETE /documents/:docId` handler validates the token signature and expiry before proceeding. This prevents accidental deletions and CSRF-style delete attacks.

---

## Sync Service — Detailed Flow

### ETag-Based Idempotent Sync

```
Sync_Service.syncDocument(docId):
  1. stored = MetadataStore.getDocument(docId)
  2. spo_item = SharePointAdapter.getItem({ driveId: stored.spo_drive_id, itemId: stored.spo_item_id })
     ├── On 3× failure → audit etag-fetch-failed, abort
  3. if spo_item.eTag == stored.spo_etag:
       audit no-change → return
  4. content_stream = SharePointAdapter.downloadItem(...)
     ├── On 3× failure → audit download-failed, abort (no partial S3 write)
  5. new_version = stored.current_version + 1
     s3_key = `docs/${docId}/v${new_version}/${stored.filename}`
  6. StorageAdapter.putObject(s3_key, content_stream, { spo_etag, version, doc_id })
     ├── Retry 3× with 30s/60s/120s back-off
     ├── On exhaustion → audit sync-failed, emit alert, abort
  7. MetadataStore.updateDocument(docId, {
       current_version: new_version,
       s3_key,
       spo_etag: spo_item.eTag,
       last_synced_at: now()
     }, condition: { spo_etag: stored.spo_etag })     // Conditional write
     ├── Condition met (won the race) → also createVersionRecord(...)
     └── Condition not met (lost the race) → silently discard (S3 object is orphaned but harmless;
         a cleanup job can reconcile orphaned objects using version records)
  8. AuditLog.append(document.sync.success, ...)
```

### Debounce Timer Implementation

**Cloud Global:** Each webhook reset writes a DynamoDB item with TTL = now + 180 seconds. An EventBridge Scheduler rule fires for every DynamoDB TTL expiry event targeting the `Sync_Service` Lambda. Because DynamoDB TTL fires within minutes of the configured epoch time, the actual debounce window is "3 minutes or slightly more."

**Sovereign / On-Prem:** BullMQ (backed by Redis) is used. Each `setDebounceTimer` call does `queue.add('sync', { docId }, { delay: 180_000, jobId: docId, removeOnComplete: true })`. A subsequent call with the same `jobId` replaces the existing job (BullMQ `repeat: false` + `jobId` uniqueness), implementing the reset.

---

## Webhook Registration and Lifecycle

### Registration on Library Enrollment

When a Document_Library is registered:
1. Generate a 32-byte random `clientState` secret.
2. For SPO (Graph): `POST /subscriptions` with `expirationDateTime = now + 29 days`.
3. For on-prem: `POST /_api/web/lists('{listId}')/subscriptions` with expiry = now + 179 days.
4. Store `SubscriptionRecord` in MetadataStore.
5. Schedule first renewal check for `expiry - 48h`.

### Renewal Job (Daily Cron)

```
for each subscription expiring within 48 hours:
  try:
    PATCH subscription with new expiry (now + 29 days)
    update SubscriptionRecord.last_renewed_at
  catch (up to 3 retries, 60s/120s/240s back-off):
    register new subscription
    delete old SubscriptionRecord
    insert new SubscriptionRecord
    audit subscription.re_registered
  if all retries exhausted:
    audit subscription.registration_failed
    emit operational alert
```

The renewal job is idempotent: running it twice within 10 minutes is safe because the second run will find no subscriptions expiring within 48 hours (they were just renewed to `now + 29 days`). If the first run partially failed, the second run retries the failed ones only.

---

## Error Handling

### Retry Policies Summary

| Operation | Max Retries | Back-off | On Exhaustion |
|---|---|---|---|
| Change Log API call | 3 | 30s fixed | Audit `change-log-fetch-failed`, discard notification |
| ETag fetch from SharePoint | 3 | 30s fixed | Audit `etag-fetch-failed`, abort sync |
| SharePoint content download | 3 | 30s fixed | Audit `download-failed`, abort sync (no partial S3 write) |
| S3 upload in Sync_Service | 3 | 30/60/120s exp | Audit `sync-failed`, emit alert |
| Subscription renewal | 3 | 60/120/240s exp | Re-register, audit `re_registered` |
| Graph API 429 | 5 | `Retry-After` header (default 60s) | Surface error, audit failed operation |
| Metadata_Store write (delete) | 3 | immediate | Return HTTP 500 |
| Audit log write | 3 | within 5s | Abort operation, return HTTP 500 |

### Transactional Guarantees

**Permission grant/revoke:** Both the Graph API call and the Metadata_Store write must succeed. If one fails, the other is rolled back. The rollback is best-effort (a compensation call). If compensation also fails, the inconsistency is logged at ERROR level and an alert is emitted. A reconciliation job can detect and repair mismatches.

**Upload failure cleanup:** If the SharePoint upload succeeds but the Metadata_Store write or S3 upload fails, the API_Service attempts to delete the orphaned SharePoint item via the SharePoint_Adapter. If that deletion also fails, an orphan-cleanup job (reconciling SharePoint items without Metadata_Store records) will clean it up.

### Partial Upload (Chunked)

The `createUploadSession` returns an `uploadUrl` valid for 24 hours. If the browser or client disconnects mid-upload, the client can resume by querying the upload session URL for the next expected byte range. If the session expires (24h), the API_Service returns HTTP 408 and instructs the user to retry the full upload.

---

## Testing Strategy

The testing approach is dual: unit/example-based tests for specific scenarios and error conditions, and property-based tests for universal correctness properties across arbitrary inputs.

### Property-Based Testing

The platform uses **fast-check** (TypeScript) for property-based testing. Each property test is configured to run a minimum of **100 iterations**. Tests are tagged with a comment referencing the design property they validate, using the format:

```
// Feature: document-share-app, Property N: <property text>
```

Property-based tests are appropriate for this feature because it contains:
- A round-trip serialisation requirement (Req 22) with a wide, typed input space
- An idempotence requirement (Req 19) over arbitrary ETag values
- A subscription renewal idempotence property (Req 18.3) over arbitrary iteration counts N

### Unit / Example-Based Tests

Unit tests cover:
- Specific HTTP status code responses (400, 403, 404, 409, 500, 502) for concrete inputs
- Error handling branches (retry exhaustion, partial failure rollback)
- Sensitivity label enforcement rules
- Confirmation token issuance and validation
- Rate limiting sliding window logic
- clientState HMAC validation

### Integration Tests

Integration tests (2–3 representative examples) cover:
- WOPI URL generation and redirect behaviour (Req 5 — 100 iterations add no value)
- Download streaming for specific version numbers (Req 4)
- Data residency endpoint blocking (Req 20) — 1 unit test per blocked domain + smoke test

### Smoke Tests

- Local Docker Compose startup and health check passes within 60 seconds (Req 15.2)
- All adapter health probes return `ok` in a known-good environment



---

## Correctness Properties

*A property is a characteristic or behavior that should hold true across all valid executions of a system — essentially, a formal statement about what the system should do. Properties serve as the bridge between human-readable specifications and machine-verifiable correctness guarantees.*

PBT is applicable to this feature for three specific requirements: Requirement 18.3 (subscription renewal idempotence), Requirement 19 (ETag-based sync idempotence), and Requirement 22 (round-trip serialisation). These were identified in the requirements document's testability notes and are confirmed by the prework analysis above.

The following properties were derived from the acceptance criteria after prework analysis and redundancy reflection:
- Requirements 19.1, 19.2, and 19.3 are related (19.1 and 19.3 are logically equivalent; 19.2 is the concurrent variant of 19.1). They are combined into a single comprehensive property.
- Requirement 22.2 is an edge-case specification (not a new property) — it constrains the generator for Property 3 rather than forming its own property.

---

### Property 1: Subscription Renewal Idempotence

*For any* set of Document_Library registrations and any integer N ≥ 2, running the subscription-renewal job N times in succession within a window shorter than the renewed subscription's validity period SHALL produce the same set of subscriptions as running it once — specifically, exactly one active subscription per library with no duplicates.

**Validates: Requirements 18.3**

---

### Property 2: ETag-Based Sync Idempotence and Uniqueness

*For any* document and any sequence of Sync_Service trigger events (including repeated triggers with the same ETag value), the number of version records created in the Metadata_Store SHALL equal the number of distinct ETag values in the trigger sequence — never more. Equivalently: no two version records for the same document SHALL share the same `spo_etag` value, regardless of how many times sync is triggered (sequentially or concurrently) for the same SharePoint content snapshot.

**Validates: Requirements 19.1, 19.2, 19.3**

---

### Property 3: Metadata Serialisation Round-Trip

*For any* valid `DocumentRecord` — including records with UTC timestamps at millisecond precision, filenames drawn from the Unicode CJK Unified Ideographs block (U+4E00–U+9FFF), large version integers, and S3 key strings containing forward-slash path separators — serialising the record to the Metadata_Store storage format and deserialising it back SHALL produce a record that is value-wise equal to the original across all fields.

**Validates: Requirements 22.1, 22.2**

---

### Property 4: API-Level Metadata Round-Trip

*For any* valid document metadata submitted via a write operation (upload or metadata update), the JSON representation returned by the corresponding read operation (GET `/documents/:docId`) SHALL contain field values that are equal to the submitted values — preserving string content (including Unicode), numeric precision, and timestamp precision to the millisecond.

**Validates: Requirements 22.3**

---

### PBT Configuration

```typescript
// Feature: document-share-app, Property 1: subscription renewal idempotence
test.prop([fc.array(libraryArbitrary(), { minLength: 1, maxLength: 20 }), fc.integer({ min: 2, max: 10 })])(
  'running renewal job N times produces same subscriptions as once',
  async (libraries, n) => { /* ... */ }
);

// Feature: document-share-app, Property 2: ETag-based sync idempotence and uniqueness
test.prop([documentRecordArbitrary(), fc.array(etagArbitrary(), { minLength: 1, maxLength: 10 })])(
  'sync trigger sequence produces version count == distinct ETag count',
  async (doc, etags) => { /* ... */ }
);

// Feature: document-share-app, Property 3: metadata serialisation round-trip
test.prop([documentRecordArbitrary()])(
  'deserialise(serialise(record)) equals record',
  async (record) => { /* ... */ }
);

// Feature: document-share-app, Property 4: API-level metadata round-trip
test.prop([uploadPayloadArbitrary()])(
  'GET response fields match POST payload fields',
  async (payload) => { /* ... */ }
);
```

Each test is configured with `numRuns: 100` (minimum). The `documentRecordArbitrary()` generator MUST produce records that include:
- Timestamps with millisecond-level precision (e.g., `"2026-04-15T08:30:00.123Z"`)
- Filenames with characters from U+4E00–U+9FFF (e.g., `"报告_2026.docx"`)
- Version numbers up to `Number.MAX_SAFE_INTEGER`
- S3 keys with multiple `/` separators (e.g., `"docs/uuid/v42/folder/report.docx"`)
- All four `SensitivityLabel` values

