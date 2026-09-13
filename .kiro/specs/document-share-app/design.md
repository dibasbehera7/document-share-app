# Design Document: Document Share App

## Overview

Document Share App is a full-stack document sharing and co-authoring platform that uses SharePoint (Online or On-Premises) as the collaborative editing surface and S3-compatible object storage as the versioned, durable system of record. The platform must operate across three deployment topologies without codebase forks: Cloud Global (AWS + SharePoint Online + Entra ID), Sovereign/On-Premises (SharePoint Server 2019/SE + MinIO or Alibaba OSS + Keycloak or ADFS), and Local Development (Docker Compose with stubs).

The core document lifecycle is: upload → browse/search → co-author via WOPI → webhook notification (ItemCheckedIn or debounce fallback) → ETag-based sync to object storage → version history available for download.

### Key Design Decisions

**SharePoint as WOPI host, not a custom WOPI implementation.** Because SharePoint already implements the WOPI host interface and manages MS-FSSHTTP co-authoring internally, the application only needs to call `GetWopiFrameUrl` to redirect the browser to the Office Web App editor. This eliminates the complexity of implementing WOPI lock management, CellSubrequest merging, and co-authoring state in application code.

**Driver pattern for all external adapters.** Every external integration (Auth, SharePoint, Storage, MetadataStore) is expressed as an interface with multiple driver implementations. Environment variables select the driver at startup. This enables a single deployment pipeline to target all topologies.

**ItemCheckedIn webhook + ETag idempotence for sync (debounce as fallback).** When the Document_Library has `ForceCheckout = true` enforced, the platform subscribes to the `ItemCheckedIn` SPO webhook event, which fires a deterministic "editing done" signal immediately when a user checks in. This is the preferred path for SPO cloud and SharePoint Server On-Premises. When Check-Out enforcement is not enabled (e.g., a library that allows direct save without checkout), the platform falls back to a per-document debounce timer that resets on every webhook notification. In both paths, the Sync_Service compares the current ETag against the stored ETag before writing — ensuring idempotent, content-driven version creation.

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
| **Webhook_Handler** | Validation handshake, async notification processing, Change Log API calls, ItemCheckedIn event processing, direct sync trigger on check-in events, Debounce_Timer reset (fallback path), clientState verification |
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
  dex (static password connector — OIDC provider, no DB, no Entra ID required)
  minio (S3-compatible)
  postgres (metadata store)
  wopi_stub (mock WOPI server)
  sharepoint_mock (optional MSGraph mock for offline dev)
```

### Webhook → Sync Flow

Two paths exist depending on whether the Document_Library has Check-Out enforcement enabled.

#### Path A: Check-Out Enforced — ItemCheckedIn (Preferred)

```
SPO webhook POST /webhook/sharepoint (ItemCheckedIn event)
    │
    ▼
Webhook_Handler: validate clientState, enqueue async task, return HTTP 200 < 5s
    │
    ▼ (async)
Webhook_Handler: call Change Log API with stored Change_Token
    ├── Identify ItemCheckedIn change record for the document
    ├── Update Change_Token
    └── Trigger Sync_Service immediately (no Debounce_Timer wait)
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

#### Path B: Check-Out Not Enforced — Debounce Fallback

```
SPO / Graph webhook POST /webhook/graph (ItemUpdated event)
    │
    ▼
Webhook_Handler: validate clientState, enqueue async task,
                 return HTTP 200 < 3s (Graph) / 5s (SPO REST)
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

`java
public interface AuthDriver {

    /** Validate an inbound Bearer token; returns parsed claims. */
    Claims validateToken(String token);

    /** Build OIDC authorisation redirect URL (PKCE). */
    String buildAuthorizationUrl(String state, String codeVerifier);

    /** Exchange authorisation code for tokens. */
    TokenSet exchangeCode(String code, String codeVerifier);

    /** Silent token refresh. */
    TokenSet refreshTokens(String refreshToken);

    /** Revoke refresh token with IdP. */
    void revokeToken(String refreshToken);
}

// Driver implementations (selected via @ConditionalOnProperty("auth.driver")):
//   EntraGlobalDriver  — cloud global
//   EntraChinaDriver   — 21Vianet SPO
//   AdfsDriver         — on-prem sovereign
//   KeycloakDriver     — sovereign / on-prem
//   DexDriver          — local development (static password connector, no DB)
`

### SharePoint_Adapter Interface

`java
public interface SharePointDriver {

    // File operations
    SpItem uploadSmall(LibraryRef libraryRef, String filename, byte[] content);
    UploadSession createUploadSession(LibraryRef libraryRef, String filename);
    UploadProgress uploadChunk(UploadSession session, byte[] chunk, ByteRange range);
    SpItem getItem(ItemRef itemRef);
    InputStream downloadItem(ItemRef itemRef);
    void deleteItem(ItemRef itemRef);

    // Metadata
    String getETag(ItemRef itemRef);
    PagedResult<SpItem> listChildren(FolderRef folderRef, String pageToken);
    List<SpItem> search(LibraryRef libraryRef, String query);

    // WOPI
    String getWopiFrameUrl(ItemRef itemRef, WopiAction action);  // WopiAction: EDIT | VIEW

    // Webhooks / subscriptions
    Subscription registerSubscription(LibraryRef libraryRef, String notificationUrl, String clientState);
    Subscription renewSubscription(String subId, Instant newExpiry);
    void deleteSubscription(String subId);
    ChangeResult getChanges(LibraryRef libraryRef, String changeToken);

    // Permissions
    void grantPermission(ItemRef itemRef, String email, SpRole role);  // SpRole: READ | WRITE
    void revokePermission(ItemRef itemRef, String permissionId);
    List<SpPermission> listPermissions(ItemRef itemRef);
}

// Driver implementations (via @ConditionalOnProperty("sharepoint.driver")):
//   GraphGlobalDriver   — graph.microsoft.com
//   GraphChinaDriver    — microsoftgraph.chinacloudapi.cn
//   SharePointRestDriver — SP Server REST API (on-prem)
`

### Storage_Adapter Interface

`java
public interface StorageDriver {

    /** Write an immutable version object. */
    void putObject(String key, InputStream stream, ObjectMetadata metadata);

    /** Stream an object to the caller. */
    InputStream getObject(String key);

    /** Check object exists; returns null if not found. */
    ObjectHead headObject(String key);

    /** Paginated list under a prefix. */
    PagedResult<ObjectInfo> listObjects(String prefix, String pageToken);
}

// Key format: docs/{doc_id}/v{version}/{filename}
//             audit/{year}/{month}/{day}/{entry_id}.json
// Driver implementations (via @ConditionalOnProperty("storage.driver")):
//   S3AwsDriver, S3MinioDriver, S3OssDriver
`

### Metadata_Store Interface

`java
public interface MetadataDriver {

    // Document records
    void createDocument(DocumentRecord doc);
    Optional<DocumentRecord> getDocument(String docId);
    void updateDocument(String docId, DocumentPatch patch, ConditionalWrite condition);
    PagedResult<DocumentRecord> listDocuments(DocumentFilter filter, PageCursor page);

    // Version records
    void createVersionRecord(VersionRecord record);
    Optional<VersionRecord> getVersionRecord(String docId, int version);
    PagedResult<VersionRecord> listVersionRecords(String docId, PageCursor page);

    // Permission records
    void upsertPermission(PermissionRecord perm);
    void deletePermission(String docId, String userId);
    Optional<PermissionRecord> getPermission(String docId, String userId);
    List<PermissionRecord> listPermissions(String docId);

    // Subscription records
    void upsertSubscription(SubscriptionRecord sub);
    void deleteSubscription(String subId);
    Optional<SubscriptionRecord> getSubscription(String subId);
    List<SubscriptionRecord> listSubscriptionsExpiringBefore(Instant cutoff);

    // Debounce timers (DynamoDB TTL items on cloud; Redisson delayed queue on sovereign)
    void setDebounceTimer(String docId, Instant firesAt);
    void clearDebounceTimer(String docId);

    // Audit log (append-only)
    void appendAuditEntry(AuditEntry entry);
}

// Driver implementations (via @ConditionalOnProperty("metadata.driver")):
//   DynamoDbDriver, JpaPostgresDriver, JpaSqlServerDriver
`

---
## Data Models

### DocumentRecord

`java
public record DocumentRecord(
    String docId,                   // UUID, primary key
    String filename,
    long fileSizeBytes,
    String spoItemId,               // SharePoint driveItem ID
    String spoDriveId,
    String spoSiteId,
    int currentVersion,             // Monotonically increasing integer, starts at 1
    String s3Key,                   // Key for current version in object storage
    String spoEtag,                 // ETag from SharePoint, used for sync idempotence
    Instant uploadTimestamp,        // UTC, millisecond precision
    String uploaderUserId,
    Instant lastModifiedAt,
    String lastModifiedBy,
    Instant lastSyncedAt,
    SensitivityLabel sensitivityLabel,
    boolean highlyConfidentialEditOverride,
    DocumentStatus status,          // ACTIVE | DELETED
    String changeToken,             // SharePoint Change Log cursor
    String libraryId,
    Instant createdAt
) {}

public enum SensitivityLabel { PUBLIC, INTERNAL, CONFIDENTIAL, HIGHLY_CONFIDENTIAL }
public enum DocumentStatus    { ACTIVE, DELETED }
`

### VersionRecord

`java
public record VersionRecord(
    String docId,                   // Partition key
    int versionNumber,              // Sort key
    String s3Key,                   // Immutable object key in Storage_Adapter
    String filename,                // Filename at sync time
    long fileSizeBytes,
    String spoEtag,                 // SharePoint ETag that triggered this version — UNIQUE per doc
    Instant syncTimestamp,
    String syncedByUserId,
    String sharepointVersionLabel   // e.g. "3.0"
) {}
`

### PermissionRecord

`java
public record PermissionRecord(
    String docId,
    String userId,
    String userEmail,
    String displayName,
    PermissionLevel permissionLevel, // VIEW | EDIT | OWNER
    String spoPermissionId,          // Graph API permission ID for revocation
    Instant grantedAt,
    String grantedBy
) {}

public enum PermissionLevel { VIEW, EDIT, OWNER }
`

### SubscriptionRecord

`java
public record SubscriptionRecord(
    String subscriptionId,          // Primary key (Graph or SPO webhook ID)
    String libraryId,
    SubscriptionType subscriptionType, // GRAPH | SPO_REST
    String notificationUrl,
    String clientStateHash,         // HMAC-SHA256 of clientState secret (never stored raw)
    String clientStateSecret,       // Stored encrypted (AES-256-GCM)
    Instant expiry,
    Instant registeredAt,
    Instant lastRenewedAt,
    String siteId,
    String driveId
) {}

public enum SubscriptionType { GRAPH, SPO_REST }
`

### AuditEntry

`java
public record AuditEntry(
    String entryId,                 // UUID
    AuditEventType eventType,
    String docId,                   // null for system-level events
    String userId,                  // user identity or "system"
    String userEmail,
    Instant timestamp,              // UTC, millisecond precision
    String clientIp,
    AuditOutcome outcome,           // SUCCESS | FAILURE | DENIED
    Map<String, Object> details     // Event-specific fields
) {}

public enum AuditEventType {
    DOCUMENT_UPLOAD,
    DOCUMENT_DOWNLOAD,
    DOCUMENT_EDIT_SESSION_OPEN,
    DOCUMENT_EDIT_SESSION_CLOSE,
    DOCUMENT_EDIT_SESSION_CHECKIN,
    DOCUMENT_SYNC_SUCCESS,
    DOCUMENT_SYNC_FAILURE,
    DOCUMENT_SYNC_NO_CHANGE,
    DOCUMENT_DELETE,
    DOCUMENT_SEARCH,
    PERMISSION_GRANT,
    PERMISSION_REVOKE,
    PERMISSION_DENIED,
    WEBHOOK_VALIDATION_MISMATCH,
    WEBHOOK_CHANGE_LOG_FETCH_FAILED,
    SUBSCRIPTION_REGISTERED,
    SUBSCRIPTION_RENEWED,
    SUBSCRIPTION_RE_REGISTERED,
    SUBSCRIPTION_REGISTRATION_FAILED,
    SENSITIVITY_LABEL_CHANGED,
    VERSION_HISTORY_LIST
}

public enum AuditOutcome { SUCCESS, FAILURE, DENIED }
`

### DebounceTimer

`java
public record DebounceTimer(
    String docId,
    long firesAtEpochSeconds,  // DynamoDB TTL attribute; or Redisson delayed queue expiry
    String libraryId,
    String subscriptionId
) {}
`
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

```java
public record DocumentListItem(
    String docId,
    String filename,
    long fileSizeBytes,
    Instant lastModifiedAt,
    String lastModifiedByDisplayName,
    int currentVersion,
    PermissionLevel effectivePermission,  // VIEW | EDIT | OWNER
    SensitivityLabel sensitivityLabel,
    String wopiViewUrl                    // null for non-Office files
) {}```
---

## Technology Stack

| Layer | Cloud Global | Sovereign / On-Prem | Local Dev | Rationale |
|---|---|---|---|---|
| **Language / Runtime** | Java 21 (LTS) on ECS Fargate | Java 21 on K8s | Java 21 (Docker) | Virtual threads (Project Loom); non-blocking I/O without reactive complexity; single codebase across all topologies |
| **Build tool** | Gradle 8 (Kotlin DSL) | Gradle 8 | Gradle 8 | Faster incremental builds than Maven; type-safe DSL |
| **API framework** | Spring Boot 3.3 (Spring MVC + virtual threads) | Spring Boot 3.3 | Spring Boot 3.3 | Industry standard; streaming via StreamingResponseBody; @ConditionalOnProperty driver selection |
| **Auth (IdP)** | Microsoft Entra ID | Keycloak / ADFS | **Dex** (static password connector) | OIDC-standard; PKCE required; Dex is a single-binary OIDC provider — no DB, <2s startup, no Entra ID setup needed for local dev |
| **Auth library** | Spring Security 6 + spring-security-oauth2-resource-server | Spring Security 6 | Spring Security 6 | JWT validation, PKCE flow, multi-provider via Spring @Profile |
| **SharePoint** | Microsoft Graph SDK for Java v6 (global) | Graph SDK (21Vianet) / SP REST via RestClient | SP REST mock | Official MS SDK; adapter pattern isolates API differences |
| **Object Storage** | AWS SDK for Java v2 (S3AsyncClient) | MinIO Java SDK / Alibaba OSS Java SDK | MinIO Java SDK | S3-compatible streaming; async client for non-blocking uploads |
| **Metadata Store** | AWS SDK for Java v2 (DynamoDbAsyncClient) | Spring Data JPA + PostgreSQL / SQL Server | Spring Data JPA + PostgreSQL | Adapter pattern; DynamoDB for cloud scale, JPA for sovereign |
| **Debounce Timer** | EventBridge Scheduler + DynamoDB TTL | Redisson (Redis delayed queue) on K8s | Redisson (Redis) | EventBridge for cloud; Redisson replaces BullMQ; Java-native Redis client |
| **IaC (cloud)** | AWS CDK for Java | — | — | Type-safe Java CDK constructs; co-located with app module |
| **IaC (sovereign)** | — | Helm + Docker Compose | Docker Compose | Kubernetes-native for sovereign; Compose for dev |
| **Observability** | Micrometer + CloudWatch / X-Ray | Micrometer + Prometheus + Grafana | Micrometer to stdout | Native Spring Boot ActuatorMetrics; same /metrics endpoint across all topologies |
| **Logging** | Logback + logstash-logback-encoder (JSON) | Logback JSON | Logback JSON to stdout | Structured JSON logs; MDC for X-Request-ID propagation |
| **PBT framework** | jqwik 1.8 (JUnit 5) | jqwik | jqwik | Java-native PBT; @Property + @ForAll + Arbitraries; integrates with JUnit 5 |
| **Test framework** | JUnit 5 + Mockito + Testcontainers | JUnit 5 + Mockito + Testcontainers | JUnit 5 + Mockito + Testcontainers | Testcontainers spins up PostgreSQL, MinIO, Keycloak for integration tests |

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

### Dex Configuration for Local Development

Dex runs as a Docker Compose service and is configured via a committed `dex-config.yaml` file. Three named test users are pre-defined to cover all permission levels and enable immediate multi-user co-authoring testing without any manual setup:

```yaml
# dex-config.yaml (committed to repo — local dev only, no real credentials)
issuer: http://dex:5556/dex

storage:
  type: memory

web:
  http: 0.0.0.0:5556

staticClients:
  - id: document-share-app
    redirectURIs:
      - 'http://localhost:8080/auth/callback'
    name: 'Document Share App (local)'
    secret: local-dev-secret

enablePasswordDB: true
staticPasswords:
  - email: 'alice@example.com'
    hash: '$2a$10$2b2cU8CPhOTaGrs1HRQuAueS7JTT5ZHsHSzYRoutBpxIa6grF6.Ra'  # password: password
    username: 'alice'
    userID: 'user-alice'
    # Role in sample data: owner — can edit, share, and delete
  - email: 'bob@example.com'
    hash: '$2a$10$2b2cU8CPhOTaGrs1HRQuAueS7JTT5ZHsHSzYRoutBpxIa6grF6.Ra'  # password: password
    username: 'bob'
    userID: 'user-bob'
    # Role in sample data: edit — can co-author; primary co-authoring partner for alice
  - email: 'charlie@example.com'
    hash: '$2a$10$2b2cU8CPhOTaGrs1HRQuAueS7JTT5ZHsHSzYRoutBpxIa6grF6.Ra'  # password: password
    username: 'charlie'
    userID: 'user-charlie'
    # Role in sample data: view — read-only; useful for testing permission enforcement
```

Spring Security points to `http://dex:5556/dex` as the OIDC issuer URI via the `OIDC_ISSUER_URI` environment variable. No Entra ID tenant, no Keycloak realm import, no external account required.
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
     s3_key = "docs/" + docId + "/v" + new_version + "/" + stored.filename
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

> **Note:** The debounce timer is used only when Check-Out enforcement is not enabled on the Document_Library. When `ItemCheckedIn` webhooks are the trigger (Path A), the timer is not involved.

**Cloud Global:** Each webhook reset writes a DynamoDB item with TTL = now + 180 seconds. An EventBridge Scheduler rule fires for every DynamoDB TTL expiry event targeting the `Sync_Service` Lambda. Because DynamoDB TTL fires within minutes of the configured epoch time, the actual debounce window is "3 minutes or slightly more."

**Sovereign / On-Prem:** Redisson (Java Redis client) is used. Each `setDebounceTimer` call schedules a delayed entry in a Redisson RDelayedQueue<String>: `delayedQueue.offer(docId, 180, TimeUnit.SECONDS)`. A reset cancels any pending entry for the same `docId` and re-schedules with a fresh 180-second delay. A dedicated consumer thread drains the queue and invokes the Sync_Service.

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

The platform uses **jqwik 1.8** (JUnit 5) for property-based testing. Each property test is annotated with @Property(tries = 100) (minimum 100 iterations). Tests are tagged with a comment referencing the design property they validate, using the format:

`java
// Feature: document-share-app, Property N: <property text>
`

Property-based tests are appropriate for this feature because it contains:
- A round-trip serialisation requirement (Req 22) with a wide, typed input space
- An idempotence requirement (Req 19) over arbitrary ETag values
- A subscription renewal idempotence property (Req 18.3) over arbitrary iteration counts N

### Unit / Example-Based Tests

Unit tests (JUnit 5 + Mockito) cover:
- Specific HTTP status code responses (400, 403, 404, 409, 500, 502) for concrete inputs
- Error handling branches (retry exhaustion, partial failure rollback)
- Sensitivity label enforcement rules
- Confirmation token issuance and validation
- Rate limiting sliding window logic
- clientState HMAC validation (MessageDigest constant-time comparison)

### Integration Tests

Integration tests use **Testcontainers** to spin up real PostgreSQL and MinIO containers (Dex runs as a plain Docker image in the compose stack; it does not need Testcontainers):
- WOPI URL generation and redirect behaviour (Req 5 — 100 iterations add no value over 2–3 representative scenarios)
- Download streaming for specific version numbers (Req 4)
- Data residency endpoint blocking (Req 20) — 1 unit test per blocked domain + smoke test

#### Multi-User Co-Authoring Integration Tests (Req 5)

Three patterns are available depending on what layer is under test:

**Pattern A — Manual (developer workflow, not automated):**
Open two browser sessions as different identities using browser profiles or an incognito window. Log in as `alice` in one and `bob` in the other via the Dex login screen (password: `password`). Both open the same document and exercise the WOPI co-authoring session. No code required; immediate after `docker compose up`.

**Pattern B — Playwright multi-context (automated, full WOPI session):**
Uses Playwright's `BrowserContext` isolation — each context has independent cookies and session state, so two contexts can be logged in as different users simultaneously.

```java
// JUnit 5 + Playwright Java — tests the full WOPI co-authoring session end-to-end
@Test
void aliceAndBobCoAuthorSameDocument() {
    try (Playwright playwright = Playwright.create()) {
        Browser browser = playwright.chromium().launch();

        // Alice opens doc for editing
        BrowserContext aliceCtx = browser.newContext();
        Page alicePage = aliceCtx.newPage();
        loginAsDexUser(alicePage, "alice@example.com", "password");
        alicePage.navigate("http://localhost:8080/documents/" + DOC_ID + "/edit");

        // Bob opens the same doc — WOPI shared lock activates
        BrowserContext bobCtx = browser.newContext();
        Page bobPage = bobCtx.newPage();
        loginAsDexUser(bobPage, "bob@example.com", "password");
        bobPage.navigate("http://localhost:8080/documents/" + DOC_ID + "/edit");

        // Assert co-authoring indicators visible in both sessions
        alicePage.waitForSelector(".coauth-indicator");
        bobPage.waitForSelector(".coauth-indicator");
    }
}
```

**Pattern C — Token injection (automated, backend sync logic only):**
Mints JWTs programmatically and submits concurrent requests against the API. No browser or Dex OIDC flow involved. Best for testing the ETag idempotence property (Property 2) under concurrent load.

```java
// Spring Boot MockMvc test — concurrent sync from two identities
@Test
void concurrentSyncProducesExactlyOneVersionBump() throws Exception {
    String aliceToken = mintTestJwt("user-alice", "alice@example.com");
    String bobToken   = mintTestJwt("user-bob",   "bob@example.com");

    CountDownLatch start = new CountDownLatch(1);
    CompletableFuture<Void> aliceSync = CompletableFuture.runAsync(() -> {
        start.await();
        mockMvc.perform(post("/internal/sync/" + DOC_ID)
            .header("Authorization", "Bearer " + aliceToken));
    });
    CompletableFuture<Void> bobSync = CompletableFuture.runAsync(() -> {
        start.await();
        mockMvc.perform(post("/internal/sync/" + DOC_ID)
            .header("Authorization", "Bearer " + bobToken));
    });

    start.countDown();  // release both at the same time
    CompletableFuture.allOf(aliceSync, bobSync).join();

    assertThat(metadataStore.getDocument(DOC_ID).currentVersion()).isEqualTo(INITIAL_VERSION + 1);
}
```

### Smoke Tests

- Local Docker Compose startup and health check passes within 60 seconds (Req 15.2)
- All adapter health probes return ok in a known-good environment


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

`java
// Feature: document-share-app, Property 1: subscription renewal idempotence
@Property(tries = 100)
void runningRenewalJobNTimesProducesSameSubscriptionsAsOnce(
        @ForAll @Size(min = 1, max = 20) List<@From("libraries") LibraryRef> libraries,
        @ForAll @IntRange(min = 2, max = 10) int n) {
    // ... assert exactly one active subscription per library after N runs
}

// Feature: document-share-app, Property 2: ETag-based sync idempotence and uniqueness
@Property(tries = 100)
void syncTriggerSequenceProducesVersionCountEqualsDistinctEtagCount(
        @ForAll @From("documentRecords") DocumentRecord doc,
        @ForAll @Size(min = 1, max = 10) List<@From("etags") String> etags) {
    // ... assert version count == number of distinct ETags in trigger sequence
}

// Feature: document-share-app, Property 3: metadata serialisation round-trip
@Property(tries = 100)
void deserialiseSerialiseRoundTrip(
        @ForAll @From("documentRecords") DocumentRecord record) {
    DocumentRecord roundTripped = deserialise(serialise(record));
    assertThat(roundTripped).isEqualTo(record);
}

// Feature: document-share-app, Property 4: API-level metadata round-trip
@Property(tries = 100)
void getResponseFieldsMatchPostPayloadFields(
        @ForAll @From("uploadPayloads") UploadPayload payload) {
    // ... POST upload, GET document, assert all fields equal
}
`

Each test runs a minimum of 100 iterations (@Property(tries = 100)). The @Provide("documentRecords") arbitrary MUST generate records that include:
- Instant timestamps with millisecond precision (e.g., Instant.parse("2026-04-15T08:30:00.123Z"))
- Filenames with characters from the Unicode CJK Unified Ideographs block (U+4E00–U+9FFF) (e.g., "报告_2026.docx")
- Version numbers up to Integer.MAX_VALUE
- S3 keys with multiple / separators (e.g., "docs/uuid/v42/folder/report.docx")
- All four SensitivityLabel enum values