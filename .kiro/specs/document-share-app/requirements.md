# Requirements Document

## Introduction

This document specifies the requirements for **Document Share App** — a full-stack document sharing and co-authoring platform that enables users to upload, browse, manage, and collaboratively edit documents stored in SharePoint (Online or On-Premises), with S3-compatible object storage as the long-term, versioned system of record.

The platform supports two primary deployment topologies:

1. **Cloud (Global):** AWS infrastructure (Lambda/ECS, API Gateway, DynamoDB, S3) integrated with SharePoint Online (Microsoft 365 Global) and Microsoft Entra ID for identity.
2. **Sovereign/On-Premises:** SharePoint Server 2019/SE with MinIO, Alibaba OSS, or AWS regional S3 as object storage; Keycloak or ADFS for identity — targeting China (PRC), KSA, and Malaysia deployments where data must not leave the sovereign boundary.

A third topology for **local development** using Docker Compose is required to allow engineers to run the full stack without a live SharePoint tenant or cloud account.

The application must manage the full document lifecycle: upload → browse → co-author in real time via WOPI → detect editing-session completion via webhooks and debounce → sync versioned copy to object storage → provide downloadable version history.

---

## Glossary

- **System**: The Document Share App application (frontend + backend services collectively).
- **API_Service**: The backend REST API service that handles document metadata, auth token management, WOPI URL generation, webhook processing, and storage synchronisation.
- **Auth_Service**: The authentication and authorisation component responsible for validating OIDC/OAuth2 tokens and enforcing per-document permissions.
- **Webhook_Handler**: The backend component that receives and processes change notifications from SharePoint (Graph subscriptions or SPO REST webhooks) and triggers the debounce timer.
- **Sync_Service**: The backend component responsible for detecting editing-session completion and synchronising the latest document version from SharePoint to object storage.
- **Storage_Adapter**: The abstraction layer over object storage (S3 / MinIO / Alibaba OSS) that provides versioned put, get, list, and delete operations.
- **Metadata_Store**: The persistent store for document metadata and version records (DynamoDB in cloud; PostgreSQL or SQL Server in sovereign/on-prem).
- **SharePoint_Adapter**: The abstraction layer over SharePoint APIs (Microsoft Graph for SPO; SharePoint REST API for on-prem), exposing a unified interface for file, version, and webhook operations.
- **WOPI_Host**: SharePoint (Online or Server) acting as the WOPI host, serving document content to Office Web Apps / Office Online Server for browser-based editing.
- **OWA / OOS**: Office Web Apps (cloud) or Office Online Server (on-premises) — the WOPI client that renders and saves Office documents in the browser.
- **Graph_Subscription**: A Microsoft Graph change-notification subscription registered against a SharePoint drive or list, delivering webhook payloads to the Webhook_Handler.
- **SPO_Webhook**: A SharePoint REST API list subscription delivering change notifications to the Webhook_Handler (used when Graph subscriptions are unavailable, e.g., on-prem).
- **RER**: Remote Event Receiver — a WCF SOAP endpoint registered on SharePoint Server On-Premises that receives synchronous or asynchronous events such as `ItemCheckedIn`.
- **Debounce_Timer**: A per-document inactivity timer (default 3 minutes) reset on every incoming webhook notification; when it fires without reset, editing is considered complete.
- **Version**: An immutable, numbered snapshot of a document stored in object storage. Version numbers are monotonically increasing integers starting at 1.
- **ETag**: The SharePoint-assigned entity tag for a driveItem or file, used to detect whether content has changed since the last sync.
- **Sensitivity_Label**: A classification marker (e.g., Public, Internal, Confidential, Highly Confidential) attached to a document, sourced from Microsoft Purview or a custom label schema.
- **Sovereign_Boundary**: The physical and logical data perimeter within which all document data, metadata, auth tokens, and audit logs must remain for regulated deployments (China, KSA, Malaysia).
- **IaC**: Infrastructure as Code — AWS CDK or Terraform templates for cloud; Docker Compose / Helm charts for sovereign/on-prem.
- **Document_Library**: A SharePoint document library (list of type DocumentLibrary) used as the collaborative editing surface.
- **WOPI_Frame_URL**: A URL returned by SharePoint's `GetWopiFrameUrl` API that the browser opens to launch the Office Web App editor.
- **Change_Token**: An opaque cursor returned by the SharePoint Change Log API, used to retrieve only changes that occurred after the last processed change.
- **Audit_Log**: An append-only, tamper-evident log of all document access, edit, share, and delete events.
- **DLP**: Data Loss Prevention — a policy engine (Microsoft Purview DLP or custom) that scans documents for sensitive content and applies restrictions.
- **ADFS**: Active Directory Federation Services — an on-premises identity federation server used in sovereign deployments as the OIDC/WS-Federation provider.
- **Keycloak**: An open-source identity and access management solution used as the OIDC provider in sovereign/on-prem deployments where Microsoft identity services are unavailable.

---

## Requirements

---

### Requirement 1: User Authentication and Session Management

**User Story:** As a user, I want to sign in using my organisational identity, so that I can access documents I am authorised to see without managing separate credentials.

#### Acceptance Criteria

1. WHEN a user navigates to the application without an active session — including when a session cookie is present but is invalid or expired — THE Auth_Service SHALL redirect the user to the configured OIDC provider (Microsoft Entra ID for global; ADFS or Keycloak for sovereign deployments) using the authorisation code flow with PKCE.
2. WHEN the OIDC provider returns a successful authorisation response, THE Auth_Service SHALL validate the ID token signature, issuer, audience, and expiry, establish an application session with a lifetime between 1 and 24 hours, and store the session server-side.
3. WHEN an authenticated user's access token is within 5 minutes of expiry, THE Auth_Service SHALL silently refresh the token using the refresh token — with no redirect or blocking prompt shown to the user — before the token expires.
4. IF token refresh fails due to an expired or revoked refresh token, THEN THE Auth_Service SHALL terminate the session and redirect the user to the sign-in page.
5. WHEN a user signs out, THE Auth_Service SHALL revoke the refresh token with the identity provider and clear all session state from the browser (cookies, sessionStorage, and localStorage) and from the server-side session store.
6. WHERE the deployment target is cloud-global, THE Auth_Service SHALL authenticate against Microsoft Entra ID at `login.microsoftonline.com` and acquire tokens scoped to `https://graph.microsoft.com/.default`.
7. WHERE the deployment target is China (21Vianet), THE Auth_Service SHALL authenticate against the 21Vianet endpoint at `login.partner.microsoftonline.cn` and acquire tokens scoped to `https://microsoftgraph.chinacloudapi.cn/.default`.
8. WHERE the deployment target is sovereign on-premises, THE Auth_Service SHALL authenticate against the configured Keycloak or ADFS OIDC endpoint and acquire tokens usable by the SharePoint_Adapter.
9. THE Auth_Service SHALL propagate identity claims (user ID, email, display name, group memberships) to all downstream services within the application boundary within 2 seconds of session establishment.

---

### Requirement 2: Document Upload

**User Story:** As a user, I want to upload documents to the platform, so that they become available for browsing, sharing, and co-authoring.

#### Acceptance Criteria

1. WHEN a user submits a file for upload, THE API_Service SHALL accept files between 1 byte and 250 GB in size; IF the submitted file is zero bytes, THEN THE API_Service SHALL return HTTP 400 with an error message indicating that empty files are not accepted.
2. WHEN an uploaded file is 4 MB or smaller, THE SharePoint_Adapter SHALL upload the file to the target Document_Library using a single PUT request to the Microsoft Graph `items/{parentId}:/{filename}:/content` endpoint (or equivalent SharePoint REST endpoint for on-prem).
3. WHEN an uploaded file exceeds 4 MB, THE SharePoint_Adapter SHALL create an upload session via `items/{parentId}:/{filename}:/createUploadSession` and upload the file in chunks not exceeding 10 MB each, within the 24-hour session validity window.
4. WHEN an upload completes successfully, THE API_Service SHALL create a Version 1 record in the Metadata_Store containing the document ID, filename, SharePoint item ID, drive ID, site ID, initial ETag, upload timestamp, and uploader identity.
5. WHEN an upload completes successfully, THE Storage_Adapter SHALL store an initial copy of the document in object storage at the key `docs/{doc_id}/v1/{filename}` and record the S3 key in the Metadata_Store.
6. IF an upload fails for any reason — including session expiry, network interruption, or SharePoint-side rejection — THEN THE API_Service SHALL attempt to clean up any partially-created upload session and return an appropriate HTTP error (408 for session expiry, 502 for upstream SharePoint failure) with a message describing the failure and instructing the user to retry.
7. IF the Document_Library already contains a file with the same name, THEN THE SharePoint_Adapter SHALL apply the `replace` conflict behaviour, incrementing the SharePoint version rather than creating a duplicate filename, and the existing Metadata_Store record for that document SHALL be updated in-place rather than a new record created.
8. WHEN an upload succeeds, THE API_Service SHALL return HTTP 201 with a response body containing the document ID, filename, size, upload timestamp, SharePoint item URL, and initial version number.
9. IF a user attempts to upload a document without holding at least the `editor` permission on the target Document_Library, THEN THE API_Service SHALL return HTTP 403 without uploading the file.
10. IF the SharePoint upload completes successfully but the subsequent Metadata_Store write or object storage write fails, THEN THE API_Service SHALL return HTTP 500, attempt to delete the orphaned SharePoint item, and ensure no partial Version 1 record remains in the Metadata_Store.

---

### Requirement 3: Document Browsing and Search

**User Story:** As a user, I want to browse and search documents in the platform, so that I can quickly find the files I need.

#### Acceptance Criteria

1. WHEN a user requests the document list, THE API_Service SHALL return a cursor-paginated list of documents from the Metadata_Store, ordered by last-modified date descending, with a default page size of 50 items and a maximum page size of 200 items; IF the `page_size` parameter exceeds 200, THE API_Service SHALL return HTTP 400.
2. WHEN a user provides a search query string of 1 to 1000 characters, THE SharePoint_Adapter SHALL execute a search against the Document_Library and return at most 200 matching items with their metadata; IF the query string is empty or exceeds 1000 characters, THE API_Service SHALL return HTTP 400.
3. THE API_Service SHALL include the following fields in each document list item: document ID, filename, file size, last-modified timestamp, last-modified-by display name, current version number, and the authenticated user's effective permission (view / edit / owner); for Office-compatible files (.docx, .xlsx, .pptx, .odt, .ods, .odp), a WOPI frame URL for viewing SHALL also be included.
4. WHEN a user requests folder contents for an existing accessible folder, THE SharePoint_Adapter SHALL list the immediate children of the specified folder; IF the folder does not exist, THE API_Service SHALL return HTTP 404; IF the folder exists but the user lacks access, THE API_Service SHALL return HTTP 403.
5. THE API_Service SHALL support filtering the document list by folder path, file extension, sensitivity label, and last-modified date range; multiple filters SHALL be combined with AND logic; IF a filter value is invalid (e.g., unrecognised sensitivity label), THE API_Service SHALL return HTTP 400.
6. WHEN a user requests the document list or a specific document's metadata, IF the document's record exists in the Metadata_Store but the corresponding SharePoint item no longer exists, THEN THE API_Service SHALL mark the document as `deleted` in the Metadata_Store, exclude it from default list results, and return it only when the `include_deleted=true` query parameter is provided.

---

### Requirement 4: Document Download

**User Story:** As a user, I want to download the current or a specific historical version of a document, so that I can work with the file offline.

#### Acceptance Criteria

1. WHEN an authenticated user with at least `view` permission requests the current version of a document, THE API_Service SHALL stream the file content from the Storage_Adapter and set the `Content-Disposition` response header to `attachment; filename="{filename}"` where `{filename}` is sourced from the document metadata in the Metadata_Store.
2. WHEN a user requests a specific version number, THE API_Service SHALL retrieve the corresponding S3 key from the version history in the Metadata_Store and stream the file from the Storage_Adapter.
3. IF the requested document ID does not exist in the Metadata_Store, THEN THE API_Service SHALL return HTTP 404 with a message identifying the missing document; IF the document exists but the requested version number does not exist in its version history, THEN THE API_Service SHALL return HTTP 404 with a message identifying the missing version.
4. IF a download request arrives without a valid authenticated session, or from an authenticated user who holds none of the `view`, `edit`, or `owner` permissions on the document, THEN THE API_Service SHALL return HTTP 403.
5. WHEN a download is served successfully, THE Audit_Log SHALL record the document ID, version number, requesting user identity, timestamp, and client IP address.
6. IF the Storage_Adapter returns an error while streaming the file, THEN THE API_Service SHALL abort the response, return HTTP 502 with an error body advising the user to retry, and record the failure in the Audit_Log.

---

### Requirement 5: Real-Time Co-authoring via WOPI

**User Story:** As a user, I want to open and edit documents in my browser with other users simultaneously, so that we can collaborate without needing desktop applications or manual merge steps.

#### Acceptance Criteria

1. WHEN a user with `edit` or `owner` permission requests to edit a document, THE API_Service SHALL call the SharePoint `GetWopiFrameUrl` API (action `edit`) and return the WOPI_Frame_URL to the browser with HTTP 200; IF the `GetWopiFrameUrl` call fails, THE API_Service SHALL return HTTP 502 with error code `WOPI_URL_FETCH_FAILED`.
2. WHEN the browser opens the WOPI_Frame_URL, THE WOPI_Host (SharePoint) SHALL serve the document to the OWA / OOS editor, allowing the user to edit the document.
3. WHEN a second user opens the same document for editing while a first user holds an active shared or exclusive lock on the document, THE WOPI_Host SHALL issue a shared lock (MS-FSSHTTP `CoauthStatus = Active`), enabling both users to edit concurrently with server-side change merging.
4. WHILE a co-authoring session is active, THE WOPI_Host SHALL accept partial cell sync requests from each editor every 60 seconds or less to maintain session liveness.
5. WHEN a user's browser closes the WOPI editor, THE WOPI_Host SHALL remove that user from the co-authoring session; if no other editors remain, THE WOPI_Host SHALL release the shared lock and create a new SharePoint major version.
6. WHERE the deployment target is sovereign on-premises and Office Online Server is configured, THE API_Service SHALL use the OOS `GetWopiFrameUrl` binding configured on the SharePoint Server farm for WOPI URL generation.
7. WHERE the deployment target is sovereign on-premises and OOS is not available, THE API_Service SHALL use a configured OnlyOffice Docs instance as the WOPI-compatible editor, generating WOPI frame URLs that point to the OnlyOffice server.
8. WHEN a user with `edit` or `owner` permission requests a WOPI frame URL, THE API_Service SHALL return an edit-mode WOPI_Frame_URL with HTTP 200.
9. WHEN a user with only `view` permission requests a WOPI frame URL, THE API_Service SHALL return a view-only WOPI_Frame_URL (action `view`) with HTTP 200; IF the user holds none of `view`, `edit`, or `owner` permission, THE API_Service SHALL return HTTP 403.

---

### Requirement 6: Webhook Registration and Lifecycle Management

**User Story:** As a platform operator, I want the system to register and maintain webhook subscriptions automatically, so that document change events are received without manual intervention.

#### Acceptance Criteria

1. WHEN a new Document_Library is registered with the system, THE API_Service SHALL register a Graph_Subscription against the library's drive root with `changeType: updated`, setting the `expirationDateTime` to 29 days from the registration date, and store the subscription ID and expiry in the Metadata_Store.
2. WHILE a Graph_Subscription's expiry is within 48 hours of the current time, THE API_Service SHALL renew the subscription by sending a PATCH to `https://graph.microsoft.com/v1.0/subscriptions/{subscriptionId}` with a new `expirationDateTime` of 29 days from the renewal date.
3. IF a Graph_Subscription renewal attempt fails, THEN THE API_Service SHALL retry the renewal up to 3 times with exponential back-off starting at 60 seconds (60s, 120s, 240s); IF all 3 retries are exhausted, THE API_Service SHALL register a new subscription, replace the old subscription record in the Metadata_Store with the new subscription ID and expiry, and record the re-registration event in the Audit_Log.
4. WHERE the deployment target is SharePoint Server On-Premises, THE API_Service SHALL register an SPO_Webhook subscription on the Document_Library's list ID using the SharePoint REST API, setting the `expirationDateTime` to 179 days from registration, and renew it using the same 3-retry exponential back-off logic described in criterion 3.
5. WHEN a validation handshake request arrives at the webhook endpoint (GET with `validationtoken` query parameter), THE Webhook_Handler SHALL respond within 5 seconds with HTTP 200, `Content-Type: text/plain`, and the body set to the exact value of the `validationtoken` parameter.
6. WHEN an incoming webhook payload's `clientState` field matches the stored secret token for the subscription ID, THE Webhook_Handler SHALL process the notification and return HTTP 200.
7. IF the `clientState` field in an incoming webhook payload does not match the stored secret token for the subscription ID, THEN THE Webhook_Handler SHALL return HTTP 200 without processing the payload and record the mismatch — including the subscription ID and received clientState hash — in the Audit_Log.
8. WHEN a Graph_Subscription lifecycle notification indicates the subscription has been deleted externally, THE API_Service SHALL remove the subscription record from the Metadata_Store and register a new subscription for the affected library within 5 minutes; IF the re-registration also fails after 3 retries, THE API_Service SHALL record a `subscription-registration-failed` alert in the Audit_Log and emit an operational alert.

---

### Requirement 7: Change Detection and Debounce

**User Story:** As the platform, I want to detect when users have stopped editing a document, so that the latest version can be synchronised to object storage without creating spurious intermediate versions.

#### Acceptance Criteria

1. WHEN a webhook notification is received for a document, THE Webhook_Handler SHALL enqueue an async processing task and return HTTP 200 to the caller within 5 seconds.
2. WHEN the async processing task runs, THE Webhook_Handler SHALL call the SharePoint Change Log API (`getchanges`) using the stored Change_Token as the start cursor, retrieve all change records since the last processed change, update the stored Change_Token with the latest token from the response, and reset the Debounce_Timer for each changed document ID to 3 minutes.
3. WHEN the Debounce_Timer for a document fires without having been reset, THE Sync_Service SHALL be triggered to evaluate whether the document requires synchronisation to object storage.
4. WHILE a webhook notification for a given document was received within the preceding 30 seconds and another notification for the same document arrives, THE Webhook_Handler SHALL reset the Debounce_Timer without performing a Change Log API call, batching Change Log queries to at most once every 30 seconds per document.
5. WHERE the deployment target is SharePoint Server On-Premises with Check-Out enforcement enabled, THE API_Service SHALL register an ItemCheckedIn Remote Event Receiver (RER) on the Document_Library.
6. WHEN the RER fires for a document, THE Sync_Service SHALL be triggered within 60 seconds without waiting for a Debounce_Timer.
7. IF the Change Log API call in criterion 2 fails, THEN THE Webhook_Handler SHALL retry the call up to 3 times with 30-second intervals; IF all retries are exhausted, THE Webhook_Handler SHALL record a `change-log-fetch-failed` event in the Audit_Log, preserve the existing Change_Token unchanged, and discard the current notification without triggering the Sync_Service.
8. IF the Debounce_Timer fires and the Sync_Service determines that no item-update records exist for the document since the last Change_Token, THEN THE Sync_Service SHALL skip the synchronisation and record a `no-change` event in the Audit_Log.

---

### Requirement 8: Version Synchronisation to Object Storage

**User Story:** As a platform operator, I want edited documents to be automatically synchronised to object storage after each editing session, so that a versioned, durable copy is always available independently of SharePoint.

#### Acceptance Criteria

1. WHEN the Sync_Service is triggered for a document, THE Sync_Service SHALL retrieve the document's current ETag and last-modified timestamp from SharePoint and compare the ETag against the `spo_etag` stored in the Metadata_Store.
2. IF the retrieved ETag is identical to the stored `spo_etag`, THEN THE Sync_Service SHALL skip the upload and record a `no-change` outcome in the Audit_Log.
3. WHEN the ETag differs from the stored value, THE Sync_Service SHALL stream the document content from SharePoint and write it to the Storage_Adapter at the key `docs/{doc_id}/v{new_version}/{filename}`, where `new_version` is `current_version + 1`.
4. WHEN the upload to the Storage_Adapter completes successfully, THE Sync_Service SHALL atomically update the Metadata_Store with: incremented `current_version`, new `s3_key`, updated `spo_etag`, updated `last_synced_at`, and a new version history record containing version number, S3 key, modified timestamp, and modifying user identity.
5. IF the upload to the Storage_Adapter fails, THEN THE Sync_Service SHALL retry the upload up to 3 times with exponential back-off (30s, 60s, 120s, max 240s); IF all retries are exhausted, THE Sync_Service SHALL record a `sync-failed` event in the Audit_Log and emit an alert to the configured alerting channel.
6. IF the Sync_Service is triggered concurrently for the same document by two independent triggers, THEN THE Sync_Service SHALL use a conditional-write on the Metadata_Store (e.g., condition: `spo_etag == stored_etag`) to ensure that at most one version bump occurs per unique ETag value; the losing concurrent write SHALL be silently discarded.
7. THE Storage_Adapter SHALL store each version as an immutable object; once written, a version object SHALL NOT be overwritten or deleted by the Sync_Service.
8. WHEN synchronisation completes, the next GET `/documents/{docId}` response issued within 5 seconds SHALL reflect the updated `current_version` and `last_synced_at` values.
9. IF the ETag fetch from SharePoint fails, THEN THE Sync_Service SHALL retry the fetch up to 3 times with 30-second intervals; IF all retries are exhausted, THE Sync_Service SHALL record an `etag-fetch-failed` event in the Audit_Log and abort the sync without modifying the Metadata_Store.
10. IF the document content download from SharePoint fails during streaming, THEN THE Sync_Service SHALL retry the download up to 3 times with 30-second intervals; IF all retries are exhausted, THE Sync_Service SHALL record a `download-failed` event in the Audit_Log and abort the sync without writing a partial object to the Storage_Adapter.

---

### Requirement 9: Version History

**User Story:** As a user, I want to view and download previous versions of a document, so that I can audit changes and recover earlier content.

#### Acceptance Criteria

1. WHEN a user requests the version history for a document, THE API_Service SHALL return a cursor-paginated list of all versions from the Metadata_Store, ordered by version number descending, with a default page size of 50 and a maximum of 200 per page; each version record SHALL contain: version number, S3 key, filename, file size, creation timestamp, and the identity of the user who triggered the sync.
2. WHEN a user requests to download a specific version, THE API_Service SHALL retrieve the object at the corresponding S3 key from the Storage_Adapter and stream it to the user with the `Content-Disposition: attachment; filename="{filename}"` header.
3. IF a download request for a version arrives from an unauthenticated request or from a user who holds none of `view`, `edit`, or `owner` permission on the document, THEN THE API_Service SHALL return HTTP 403.
4. WHEN a user requests a document's metadata, THE API_Service SHALL include the total version count and the creation timestamp of the earliest stored version in the response body.
5. IF a version object is missing from the Storage_Adapter but the record exists in the Metadata_Store, THEN THE API_Service SHALL return HTTP 410 (Gone) for that version download request and record the inconsistency — including document ID and version number — in the Audit_Log.
6. WHEN a user successfully retrieves the version history list for a document, THE Audit_Log SHALL record the document ID, requesting user identity, timestamp, and outcome.

---

### Requirement 10: Document Permissions

**User Story:** As a document owner, I want to control who can view or edit my document, so that sensitive content is accessible only to authorised users.

#### Acceptance Criteria

1. THE API_Service SHALL enforce three permission levels per document and per user: `view` (read-only access, no editing), `edit` (read-write access, co-authoring allowed), and `owner` (full control including permission management and deletion).
2. WHEN a document is uploaded, THE API_Service SHALL assign `owner` permission to the uploading user and record this in the Metadata_Store.
3. WHEN an owner grants a permission to another user, THE API_Service SHALL call the Graph API invite endpoint with the corresponding role (`read` for `view`, `write` for `edit`) and simultaneously record the permission in the Metadata_Store; IF either the Graph API call or the Metadata_Store write fails, THE API_Service SHALL roll back the successful operation and return HTTP 500.
4. WHEN an owner revokes a permission from a user, THE API_Service SHALL call the Graph API delete-permission endpoint and remove the corresponding record from the Metadata_Store; IF either operation fails, THE API_Service SHALL roll back the successful operation and return HTTP 500.
5. IF a user attempts to perform an action requiring a permission they do not hold — `view` for download or version history, `edit` for editing, `owner` for granting permission or deletion — THEN THE API_Service SHALL return HTTP 403 and record the access denial in the Audit_Log.
6. WHEN a user's permission is modified (granted or revoked), THE Audit_Log SHALL record the document ID, target user identity, old permission level, new permission level, actor identity, and timestamp.
7. IF a user requests the list of permissions on a document and that user does not hold `owner` permission, THEN THE API_Service SHALL return HTTP 403.

---

### Requirement 11: Document Deletion

**User Story:** As a document owner, I want to delete documents I no longer need, so that the platform does not retain obsolete files.

#### Acceptance Criteria

1. WHEN an owner submits a valid deletion request, THE API_Service SHALL delete the document from SharePoint; if the SharePoint deletion succeeds but the Metadata_Store status update to `deleted` fails, THE API_Service SHALL retry the Metadata_Store update up to 3 times before returning HTTP 500.
2. THE Storage_Adapter SHALL retain all versioned objects in object storage after a document deletion; the physical S3 objects SHALL NOT be deleted by the delete operation.
3. WHEN a document is successfully deleted, THE Audit_Log SHALL record the document ID, filename, deleting user identity, and deletion timestamp.
4. IF a deletion request is made by a user who does not hold `owner` permission, THEN THE API_Service SHALL return HTTP 403 and record the denied attempt in the Audit_Log.
5. THE API_Service SHALL require a `confirmation_token` — a short-lived token (valid for 15 minutes) returned in the document's GET metadata response — to be included in the delete request body; IF the token is absent, expired, or invalid, THE API_Service SHALL return HTTP 409 without deleting the document.

---

### Requirement 12: Audit Logging

**User Story:** As a compliance officer, I want a tamper-evident log of all document operations, so that I can investigate incidents and satisfy regulatory audit requirements.

#### Acceptance Criteria

1. THE Audit_Log SHALL record an entry for each of the following operations: document upload, document download (any version), document edit session opened, document edit session closed, version sync to object storage, permission granted, permission revoked, document deleted, document searched, webhook validation mismatch, and sync failure.
2. EACH Audit_Log entry SHALL contain at minimum: event type, document ID, user identity (or `system` for automated operations), timestamp (UTC, millisecond precision), client IP address, and outcome (success / failure / denied).
3. THE Audit_Log storage SHALL be append-only; existing entries SHALL NOT be modifiable or deletable by any application component.
4. WHERE the deployment target is cloud-global, THE Audit_Log SHALL be persisted to a dedicated, versioning-enabled S3 bucket with Object Lock in Compliance mode for a retention period of 7 years.
5. WHERE the deployment target is sovereign on-premises, THE Audit_Log SHALL be persisted to a WORM-capable on-premises storage system (e.g., immutable SQL insert-only table with write access limited to the API_Service service account).
6. IF an Audit_Log write fails after 3 retry attempts within 5 seconds, THEN THE API_Service SHALL NOT proceed with the requested operation, SHALL return HTTP 500 to the caller, and SHALL ensure that no partial state from the intended operation is committed.

---

### Requirement 13: Sensitivity Labels

**User Story:** As a compliance officer, I want documents to carry sensitivity classification labels, so that data handling policies are enforced automatically and consistently.

#### Acceptance Criteria

1. THE API_Service SHALL support the following sensitivity label values: `Public`, `Internal`, `Confidential`, and `Highly_Confidential`.
2. WHEN a document is uploaded without an explicit sensitivity label, THE API_Service SHALL assign the `Internal` label as the default.
3. IF a user requests the generation of an anonymous-scope sharing link for a document tagged with `Confidential` or `Highly_Confidential`, THEN THE API_Service SHALL return HTTP 403 with an error body identifying the label as the reason for rejection.
4. WHEN a document carries the `Highly_Confidential` label, THE API_Service SHALL only return view-only WOPI_Frame_URLs (action `view`) for all users, including `edit` and `owner` permission holders; the exception is when the document owner has set the per-document boolean flag `highly_confidential_edit_override` to `true`, in which case `owner`-permission holders receive an edit-mode URL.
5. WHERE the deployment target is cloud-global and Microsoft Purview Information Protection is enabled, THE API_Service SHALL read the sensitivity label from the SharePoint `sensitivityLabel` field on the driveItem on each document metadata retrieval and on each webhook notification; IF the Purview label differs from the locally stored label, the Purview label SHALL take precedence and the Metadata_Store SHALL be updated accordingly.
6. WHEN the sensitivity label on a document changes — either via user action or Purview sync — THE Audit_Log SHALL record the document ID, old label, new label, actor identity (or `system` for Purview sync), and timestamp.
7. THE API_Service SHALL expose the current sensitivity label in all document metadata API responses.
8. WHEN a document owner changes the sensitivity label on an existing document, THE API_Service SHALL update both the SharePoint `sensitivityLabel` field and the Metadata_Store atomically; IF either update fails, THE API_Service SHALL roll back the successful update and return HTTP 500.

---

### Requirement 14: Multi-Deployment Configuration

**User Story:** As a platform engineer, I want the application to support cloud-global, sovereign, and local-development deployment modes from a single codebase, so that I can manage one implementation across all target environments.

#### Acceptance Criteria

1. THE System SHALL determine all environment-specific configuration (identity provider URL, SharePoint tenant URL, object storage endpoint, metadata store connection, audit log destination) from environment variables; if an environment variable is absent, THE System SHALL fall back to a configuration file injected at deployment time; environment variables SHALL take precedence over configuration file values when both are present; no environment-specific values SHALL be hard-coded in application source.
2. THE SharePoint_Adapter SHALL expose an abstraction that accepts driver configuration for: `graph-global` (Microsoft Graph global endpoint), `graph-china` (Microsoft Graph 21Vianet endpoint), and `sharepoint-rest-onprem` (SharePoint Server REST API with NTLM or form digest auth).
3. THE Storage_Adapter SHALL expose an abstraction that accepts driver configuration for: `s3-aws` (AWS S3), `s3-minio` (MinIO self-hosted), and `s3-oss` (Alibaba Cloud OSS S3-compatible endpoint).
4. THE Auth_Service SHALL accept driver configuration for: `entra-global` (Microsoft Entra ID global), `entra-china` (Microsoft Entra ID 21Vianet), `adfs` (on-prem ADFS OIDC), and `keycloak` (Keycloak OIDC).
5. WHERE the deployment target is cloud-global, THE System SHALL use AWS DynamoDB as the Metadata_Store, AWS S3 as the Storage_Adapter, and Microsoft Entra ID as the Auth_Service.
6. WHERE the deployment target is sovereign on-premises, THE System SHALL use PostgreSQL or SQL Server as the Metadata_Store, MinIO or Alibaba OSS as the Storage_Adapter, and Keycloak or ADFS as the Auth_Service.
7. THE System SHALL expose a `/health` endpoint that returns HTTP 200 with a JSON body containing a `status` field (`ok` / `degraded` / `unavailable`) for each adapter (SharePoint_Adapter, Storage_Adapter, Metadata_Store, Auth_Service) and a top-level aggregate `status`; IF any adapter reports `unavailable`, the endpoint SHALL return HTTP 503.
8. IF the configured driver key for any adapter is not one of the supported values, THEN THE System SHALL fail to start and emit a descriptive error message identifying the invalid driver key and the list of accepted values.

---

### Requirement 15: Local Development Environment

**User Story:** As a developer, I want to run the complete application locally using Docker Compose, so that I can develop and test features without requiring access to a live SharePoint tenant or cloud account.

#### Acceptance Criteria

1. THE System SHALL provide a `docker-compose.yml` file in the repository root that starts all required services: API_Service, a MinIO instance (acting as the Storage_Adapter), a PostgreSQL instance (acting as the Metadata_Store), a Keycloak instance in development mode (acting as the OIDC provider), and a bundled WOPI stub server for local document editing.
2. WHEN `docker compose up` is executed from the repository root, THE System SHALL start all services and the API_Service SHALL pass all health checks within 60 seconds.
3. IF the `WOPI_MOCK` environment variable is set to `true`, THEN THE API_Service SHALL replace calls to SharePoint's `GetWopiFrameUrl` with a locally generated URL pointing to the bundled WOPI stub server.
4. THE local development environment SHALL pre-seed the Metadata_Store with at least two sample documents and one sample user with `owner` permission so that developers can exercise the UI immediately after startup.
5. THE local development environment SHALL mount the API_Service source directory into the container; WHEN a source file changes, THE API_Service SHALL reload within 10 seconds without requiring a container restart.
6. WHEN running in local development mode, THE Audit_Log SHALL write structured JSON entries to stdout rather than to a WORM storage backend.
7. THE System SHALL provide a `.env.example` file documenting all required and optional environment variables with descriptions, accepted values, and default values for local development.

---

### Requirement 16: Infrastructure as Code — Cloud (AWS)

**User Story:** As a platform engineer, I want the cloud infrastructure defined as code, so that I can provision and update the production environment reproducibly.

#### Acceptance Criteria

1. THE System SHALL provide AWS CDK (TypeScript) or Terraform HCL stacks in an `infra/cloud` directory that define all required AWS resources: API Gateway, Lambda functions or ECS task definition, DynamoDB table with version history composite key schema, S3 bucket for documents with versioning and Object Lock enabled, S3 bucket for audit logs with Object Lock in Compliance mode, CloudFront distribution, IAM roles with least-privilege policies, and EventBridge rules for debounce timer invocation.
2. THE IaC stacks SHALL accept a `deploymentEnvironment` parameter with values `dev`, `staging`, and `production`, with the following concrete per-environment configurations: `dev` — no Object Lock, 30-day log retention, minimal DynamoDB capacity; `staging` — Object Lock with 30-day Governance mode retention, 90-day log retention; `production` — Object Lock with 7-year Compliance mode retention, full DynamoDB capacity with auto-scaling.
3. WHEN the IaC stacks are applied with no pre-existing state, THE System SHALL provision all resources within a single `deploy` or `apply` command without manual intervention.
4. THE IaC stacks SHALL enforce that the documents S3 bucket has Block Public Access enabled, server-side encryption with AWS KMS enabled, and no public-read bucket policies.
5. THE IaC stacks SHALL configure the DynamoDB table with point-in-time recovery enabled and a TTL attribute on Debounce_Timer records.

---

### Requirement 17: Infrastructure as Code — Sovereign / On-Premises

**User Story:** As a platform engineer, I want the sovereign and on-premises infrastructure defined as code, so that I can deploy to regulated environments reproducibly.

#### Acceptance Criteria

1. THE System SHALL provide Docker Compose files and/or Helm charts in an `infra/sovereign` directory that define all required components: API_Service, MinIO (or Alibaba OSS SDK adapter), PostgreSQL, Keycloak, Nginx reverse proxy, and optional OnlyOffice Docs for WOPI.
2. THE Helm charts SHALL accept a `global.storageBackend` value with options `minio`, `alibaba-oss`, and `aws-s3-regional` to configure the Storage_Adapter driver.
3. THE Helm charts SHALL accept a `global.identityProvider` value with options `keycloak` and `adfs` to configure the Auth_Service driver.
4. WHEN the sovereign deployment is configured for China (PRC), THE System SHALL route all SharePoint API calls to the 21Vianet Graph endpoint `https://microsoftgraph.chinacloudapi.cn/v1.0/` and all auth calls to `https://login.partner.microsoftonline.cn`; the SharePoint_Adapter SHALL return an error without making the call if a configured endpoint URL contains the domains `graph.microsoft.com`, `login.microsoftonline.com`, or `*.sharepoint.com` in a China sovereign deployment.
5. THE sovereign IaC definitions SHALL not require outbound internet access during deployment; all container images SHALL be sourced from a configured private registry.
6. THE System SHALL provide a Helm chart value `audit.wormStorageClass` that specifies the Kubernetes StorageClass used for WORM audit log storage; a StorageClass is considered compliant if it is backed by a PersistentVolume that rejects delete and overwrite operations at the storage layer; this value defaults to `standard` for local development.

---

### Requirement 18: Webhook Subscription Renewal — Correctness Property

**User Story:** As a platform operator, I want webhook subscriptions to be renewed automatically and without duplicates, so that the system never misses document change events.

#### Acceptance Criteria

1. THE API_Service SHALL maintain at most one active Graph_Subscription per Document_Library at any time; WHEN a renewal creates a new subscription due to a failed renewal of the existing one, THE API_Service SHALL delete the old subscription record from the Metadata_Store before registering the new one.
2. THE API_Service subscription-audit job SHALL run daily and emit a metric `subscription_coverage_ratio` equal to `(libraries_with_valid_subscription / total_registered_libraries)`; IF this ratio drops below 1.0, THE API_Service SHALL emit an operational alert identifying the libraries with missing or expired subscriptions.
3. WHEN the subscription-renewal job runs, THE API_Service SHALL process all subscriptions expiring within 48 hours in a single pass; the set of subscriptions processed in consecutive renewal runs SHALL be idempotent — running the renewal job twice within 10 minutes SHALL NOT produce duplicate subscriptions.

---

### Requirement 19: ETag-Based Sync Idempotence

**User Story:** As the platform, I want the version sync process to be idempotent, so that repeated triggers for the same document state do not create unnecessary version objects.

#### Acceptance Criteria

1. FOR ALL documents tracked by the system, applying the sync operation twice with the same SharePoint ETag value SHALL result in exactly one version record in the Metadata_Store and exactly one corresponding object in the Storage_Adapter — not two.
2. WHEN the Sync_Service is invoked concurrently by two independent triggers for the same document, the Metadata_Store conditional-write mechanism SHALL ensure that the final `current_version` is incremented exactly once per unique ETag value.
3. FOR ALL version records in the Metadata_Store, the `spo_etag` field SHALL uniquely identify the SharePoint content snapshot that produced that version; no two version records for the same document SHALL share the same `spo_etag` value.

---

### Requirement 20: Data Residency Enforcement

**User Story:** As a compliance officer deploying to a sovereign region, I want assurance that no document content, metadata, or authentication tokens leave the sovereign boundary, so that regulatory data-residency requirements are met.

#### Acceptance Criteria

1. WHERE the deployment target is sovereign on-premises, THE System SHALL store all document binaries in the on-premises Storage_Adapter (MinIO or Alibaba OSS) and NOT transmit document content to any endpoint outside the configured sovereign boundary.
2. WHERE the deployment target is sovereign on-premises, THE Metadata_Store SHALL reside on on-premises infrastructure; document metadata, version records, and audit logs SHALL NOT be stored in cloud-hosted services outside the sovereign boundary.
3. WHERE the deployment target is sovereign on-premises, THE Auth_Service SHALL authenticate users against the on-premises Keycloak or ADFS instance; no authentication request or token validation SHALL be directed to `login.microsoftonline.com` or any Microsoft global endpoint.
4. WHERE the deployment target is China (21Vianet SPO), THE SharePoint_Adapter SHALL use only the 21Vianet Graph endpoint (`microsoftgraph.chinacloudapi.cn`) and the 21Vianet auth endpoint (`login.partner.microsoftonline.cn`); IF the SharePoint_Adapter is configured with an endpoint URL containing `graph.microsoft.com` or `login.microsoftonline.com`, it SHALL return an error and refuse to make the call.
5. THE System SHALL expose a `/compliance/data-residency` endpoint that returns a machine-readable JSON report listing each adapter (SharePoint, Storage, Auth, MetadataStore, AuditLog), its configured endpoint URL, and its classified data-residency zone; this endpoint SHALL be accessible to `owner`-level users and platform operators only.

---

### Requirement 21: API Rate Limit and Throttle Handling

**User Story:** As a platform engineer, I want the application to handle SharePoint and Graph API throttling gracefully, so that brief quota exhaustion events do not cause data loss or user-visible errors.

#### Acceptance Criteria

1. WHEN the SharePoint_Adapter receives an HTTP 429 (Too Many Requests) response from any Graph or SharePoint REST API call, THE SharePoint_Adapter SHALL honour the `Retry-After` response header value (an integer number of seconds) and retry the request after waiting that duration.
2. IF no `Retry-After` header is present in a 429 response, THEN THE SharePoint_Adapter SHALL wait 60 seconds before retrying.
3. THE SharePoint_Adapter SHALL retry throttled requests up to 5 times; IF all 5 retries are exhausted, THEN THE SharePoint_Adapter SHALL surface an error to the calling service and record the failed operation in the Audit_Log.
4. THE API_Service SHALL implement a per-user request rate limit of 100 requests per minute using a sliding-window algorithm; WHEN a user exceeds this limit, THE API_Service SHALL return HTTP 429 with a `Retry-After` header (integer seconds until the window resets) and an `X-RateLimit-Reset` header (UTC epoch seconds of the reset time).
5. WHILE the Graph API application-level quota (10,000 requests per 10 minutes) is approaching saturation (above 80% utilisation as estimated by a request counter), THE SharePoint_Adapter SHALL throttle outbound calls using a token-bucket algorithm before sending, to prevent triggering server-side throttling.

---

### Requirement 22: Round-Trip Serialisation of Document Metadata

**User Story:** As a developer, I want the document metadata serialisation to be lossless, so that no metadata is corrupted when stored and retrieved across serialisation boundaries.

#### Acceptance Criteria

1. THE Metadata_Store serialiser SHALL serialise document metadata objects to the storage format (JSON or SQL row) and deserialise them back to in-memory objects; FOR ALL valid document metadata objects, `deserialise(serialise(metadata))` SHALL produce an object that is value-wise equal to the original — all fields match in value, type, and precision.
2. THE Metadata_Store serialiser SHALL correctly round-trip the following field types without loss: UTC timestamps with millisecond precision, Unicode filenames including characters from the CJK Unified Ideographs block (U+4E00–U+9FFF), version numbers as integers, and S3 key strings containing forward-slash path separators.
3. THE API_Service SHALL serialise all JSON API response bodies in UTF-8; FOR ALL valid document metadata inputs, the JSON representation parsed from the API response SHALL contain values equal to the values submitted in the corresponding write operation.

---

### Requirement 23: Error Handling and Observability

**User Story:** As a platform engineer, I want structured error responses and operational metrics, so that I can diagnose issues quickly and maintain SLA targets.

#### Acceptance Criteria

1. THE API_Service SHALL return error responses in a consistent JSON schema containing at minimum: `error_code` (machine-readable string), `message` (human-readable description), `request_id` (UUID traceable in logs), and `timestamp`.
2. THE API_Service SHALL emit structured JSON logs for every inbound request, including: request method, path, authenticated user ID, response status code, response time in milliseconds, and request ID.
3. THE API_Service SHALL expose a Prometheus-compatible `/metrics` endpoint that responds within 200 ms and reports: request count by method and status code, request latency histogram by path, Sync_Service trigger count, sync success count, sync failure count, active Graph_Subscription count, and Debounce_Timer active count.
4. WHEN an unhandled exception occurs in any component, THE System SHALL log the full stack trace, request context, and affected document ID (if applicable) at `ERROR` level without including document content or user credentials in the log entry.
5. THE System SHALL propagate a `X-Request-ID` header through all internal service calls, ensuring that a single user request generates log entries bearing the same request ID across all components.

---

## Notes on Testability

The following requirements are identified as candidates for property-based testing:

- **Requirement 19** (ETag-based idempotence): The idempotence and uniqueness properties of the sync operation across arbitrary ETag values are natural property tests.
- **Requirement 22** (Round-trip serialisation): The round-trip property `deserialise(serialise(x)) == x` (value-wise equality) over a wide variety of metadata inputs — including Unicode filenames, edge-case timestamps, and large version numbers — is a classic property-based test.
- **Requirement 18.3** (Idempotent renewal): The property that running the renewal job N times within a short window produces exactly one subscription per library can be verified with a property test over N.

The following requirements are best tested with integration tests using representative examples:

- **Requirement 5** (Co-authoring via WOPI): Co-authoring behaviour depends on SharePoint Server or a real WOPI host; 100 iterations add no value over 2–3 representative scenarios.
- **Requirement 4** (Download): Download correctness is verifiable with 2–3 representative version numbers; streaming behaviour does not benefit from property-based testing.
- **Requirement 20** (Data residency): Endpoint blocking is verified by unit tests on the SharePoint_Adapter guard logic plus 1–2 integration smoke tests confirming no calls reach global Microsoft endpoints.
