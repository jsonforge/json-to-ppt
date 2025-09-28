# Hosted API Specification (Draft)

Base URL: `https://api.example.com`
Authentication: API Key via header `X-API-Key: <token>`
Versioning: Prefix all paths with `/v1/`. Breaking changes bump major.
Format: JSON request/response. PPTX binary delivered via download URL.

## 1. Create Render Job
`POST /v1/render`

### Purpose
Submit a JSON payload (schema compatible) to generate a PPTX. Small jobs may complete synchronously.

### Request Body
```
{
  "payload": { ... original ppt JSON ... },
  "mode": "auto" | "sync" | "async",
  "callbackUrl": "https://.../webhook" (optional),
  "options": {
    "priority": "normal" | "high",
    "returnSlidesHash": true
  }
}
```
Constraints:
- `payload.ppt.slides.length` <= plan limit
- Raw body size <= 1 MB (sync) else force async

### Response (Sync Success)
```
{
  "jobId": "job_123",
  "status": "completed",
  "slides": 12,
  "fileSize": 842311,
  "downloadUrl": "https://api.example.com/v1/jobs/job_123/download",
  "hash": "sha256:..." (optional)
}
```

### Response (Accepted Async)
```
{
  "jobId": "job_124",
  "status": "queued",
  "etaSeconds": 8
}
```

## 2. Get Job Status
`GET /v1/jobs/{jobId}`
```
Response:
{
  "jobId": "job_124",
  "status": "queued" | "processing" | "completed" | "failed" | "expired",
  "slides": 12,
  "error": null,
  "downloadUrl": "..." (when completed),
  "createdAt": "2025-09-29T12:00:00Z",
  "updatedAt": "2025-09-29T12:00:05Z"
}
```
Status TTL: jobs auto-expire after 24h (file deleted) unless plan upgrade.

## 3. Download File
`GET /v1/jobs/{jobId}/download`
- Returns PPTX (Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation)
- 302 redirect to signed storage URL (S3 / GCS) recommended

## 4. Usage Metrics
`GET /v1/usage` – current billing period usage.
```
{
  "periodStart": "2025-09-01",
  "periodEnd": "2025-09-30",
  "slidesRendered": 18420,
  "planLimit": 20000,
  "overageSlides": 0,
  "jobs": 612
}
```

## 5. API Key Management
`POST /v1/keys` -> create (role-limited)
`DELETE /v1/keys/{keyId}` -> revoke
`GET /v1/keys` -> list

Key object:
```
{
  "keyId": "k_abc123",
  "prefix": "ppt_live_abc",
  "createdAt": "2025-09-29T12:00:00Z",
  "lastUsedAt": "2025-09-29T12:05:00Z",
  "scopes": ["render:write", "usage:read"],
  "revoked": false
}
```

## 6. Webhook Delivery
On completion (if `callbackUrl` provided) send POST:
```
{
  "event": "job.completed",
  "jobId": "job_124",
  "slides": 12,
  "downloadUrl": "...",
  "hash": "sha256:..."
}
```
Headers include `X-Signature: sha256=...` (HMAC with user secret) for authenticity.

## 7. Errors
```
{
  "error": {
    "code": "INVALID_SCHEMA" | "RATE_LIMIT" | "UNAUTHORIZED" | "INTERNAL_ERROR",
    "message": "...",
    "details": { }
  }
}
```

## 8. Rate Limiting
- 429 with `Retry-After` seconds header
- Per-key sliding window (e.g. 120 requests / minute free tier)

## 9. Hashing & Integrity
- Generation returns deterministic SHA256 of PPTX for regression tracking.
- Optional `slidesHash` per individual slide (future extension).

## 10. Future Extensions
- `PATCH /v1/render/{jobId}` for JSON patch re-renders
- Template ID referencing: `{ "templateId": "tpl_finance_qtr_v1", "data": {...}}`
- Bulk batch endpoint for scheduled nightly jobs
