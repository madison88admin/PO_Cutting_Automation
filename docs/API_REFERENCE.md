# PO Cutting Automation — Complete API Reference

Base URLs:
- **Local**: `http://localhost:3000` (or the port shown by `npm run dev`)
- **VPS Production**: `https://po-cutting.5-223-78-194.sslip.io`
- **Netlify Frontend**: `https://m88-po-cutting.netlify.app` (proxies to VPS backend)

Auth notes:
- Production upload endpoints validate `Origin` / `x-forwarded-host` against `ALLOWED_UPLOAD_ORIGINS`.
- `processing-jobs` / `header-preview-jobs` require header `x-processing-key: <ADMIN_PANEL_PASSWORD>` in production.
- Admin endpoints require the `admin_session` cookie (login via `/api/admin/session`).

---

## 1. Health

### `GET /api/health`
Simple health check.

```json
{ "status": "ok", "service": "po-line", "timestamp": "2026-09-17T08:00:00.000Z" }
```

---

## 2. Upload & Processing (Core Workflow)

### `POST /api/upload` — Full upload pipeline (main endpoint)
The complete workflow: parse Excel buy file(s) → validate → match styles in NextGen → generate ORDERS/LINES/ORDER_SIZES output workbooks.

**Request**: `multipart/form-data`
| Field | Type | Description |
|---|---|---|
| `file` | File (repeatable, max 5) | Excel buy files (.xlsx/.xls, ≤30 MB each) |
| `manualPo` | string | Optional manual PO number override |
| `manualDestination` | string | Optional transport destination |
| `manualProductRange` / `manualSeason` | string | Product range / season override |
| `manualTemplate` / `manualLinesTemplate` | string | ORDERS / LINES template override |
| `manualComments` | string | Comments for output |
| `manualKeyDate` | string | Key date |
| `manualKeyUser1..5` | string | Key user overrides |
| `manualCustomer` | string | Customer override |
| `manualBrand` | string | Brand key override (used for NextGen search + learning cache) |
| `headerMappings` | JSON string | `{ [filename]: { headers: string[], mapping: {...} } }` confirmed header mapping from preview step |
| `nextgenOverrides` | JSON string | `{ "style|color": { product, colorName, colorCode, productExternalRef, productCustomerRef } }` user corrections for ambiguous matches |
| `x-user-id` (header) | UUID | Optional user id for audit/run history |

**Response** (200):
```json
{
  "success": true,
  "runId": "uuid-or-null",
  "dataCount": 4,
  "errors": [ { "field": "...", "row": 3, "message": "...", "severity": "CRITICAL|WARNING" } ],
  "canProceed": true,
  "files": { "orders": "<base64 xlsx>", "lines": "<base64 xlsx>", "sizes": "<base64 xlsx>" },
  "fileSummary": [ { "filename": "...", "orders": 4, "lines": 152, "sizes": 152, "errors": 0, "warnings": 2, "brands": ["TNF"] } ],
  "output": [ /* ProcessedPO[] merged across files */ ],
  "mergedSummary": { "orders": 4, "lines": 152, "sizes": 152, "errors": 0, "warnings": 2 },
  "formatDetection": { "<filename>": { /* FormatDetection */ } },
  "needsAttention": [ { "code": "MISSING_STYLE|MISSING_COLOUR|AMBIGUOUS_NEXGEN_MATCH|NEXGEN_MATCH_NOT_FOUND", "purchaseOrder": "...", "lineItem": 1, "style": "...", "colour": "...", "message": "...", "candidates": [ /* NextGenVariantCandidate[] */ ] } ],
  "nexgenVariantSummary": { "requested": 12, "resolved": 11, "blocked": 1 }
}
```

Errors: `400` no files/too many/too big/wrong type · `403` invalid origin (prod) · `429` rate limit (10/min) · `500` server error (includes `stage`).

### `POST /api/extract-buy-file` — Newer deterministic extraction pipeline
Extracts items + rich product data and generates the 3 output workbooks in one call. Does **not** create run history or audit events.

**Request**: `multipart/form-data` with `file` (single buy file) and optional `files` (product/reference sheets).

**Response** (200):
```json
{
  "success": true,
  "filename": "Copy of FW2627 BULK BUY 4.xlsx",
  "model": "exceljs-local",
  "result": {
    "items": [ /* BuyFileItem[] with style/color/size/qty + NextGen match info */ ],
    "productData": [ /* ProductData[] single source of truth for generators */ ],
    "headerRow": 3,
    "headers": [ "...raw headers..." ],
    "mapping": { /* canonical → column name */ },
    "templateUsed": true,
    "unmappedColumns": [ "..." ],
    "files": { "orders": "<base64>", "lines": "<base64>", "sizes": "<base64>" },
    "latestPO": null,
    "colorNames": {},
    "matchSummary": { "matched": 150, "ambiguous": 1, "unmatched": 1, "not_checked": 0, "totalQuantity": 500, "matchedQuantity": 490 },
    "matchIssues": [ { "field": "Nexgen Product Match", "row": 12, "severity": "CRITICAL", "status": "ambiguous", "message": "...", "candidates": [...] } ]
  }
}
```

### `POST /api/header-preview` — Header detection & mapping preview
For each uploaded file, picks the best sheet: detects the header row (AI), maps headers to canonical fields (learned template → AI + aliases), and runs semantic validation.

**Request**: `multipart/form-data` with one or more `file` fields.

**Response** (200):
```json
{
  "success": true,
  "previews": [
    {
      "filename": "Buy.xlsx",
      "worksheet": "Buy Sheet",
      "headerRow": 4,
      "headers": ["PO #", "Style#", "Qty"],
      "mapping": { "po_number": "PO #", "style": "Style#", "quantity": "Qty" },
      "confidence": 90,
      "fieldConfidence": { "po_number": 95, "style": 95, "quantity": 90 },
      "mappingWarnings": ["Header 'Qty' looks like a date column but is mapped to quantity"],
      "unmappedColumns": ["Notes"],
      "source": "learned template | Qwen + header aliases"
    }
  ]
}
```

---

## 3. Async Job API (Netlify-safe wrappers)

The Netlify frontend uses these to avoid function timeouts. VPS runs the same routes internally.

### `POST /api/processing/jobs` (frontend proxy) / `POST /api/processing-jobs` (VPS)
Starts an async full upload job. Body = same `multipart/form-data` as `/api/upload`.
**Response** `202`: `{ "jobId": "uuid", "status": "processing" }`

### `GET /api/processing/jobs?id=<uuid>` (proxy) / `GET /api/processing-jobs/<id>` (VPS)
Poll job status.
```json
{ "id": "uuid", "status": "processing|completed|failed", "result": { /* same shape as /api/upload response */ }, "error": null }
```

### `POST /api/processing/header-preview` (proxy) / `POST /api/header-preview-jobs` (VPS)
Async header preview job. Body = same as `/api/header-preview`. Returns `202 { jobId, status }`.

### `GET /api/processing/header-preview?id=<uuid>` (proxy) / `GET /api/processing-jobs/<id>` (VPS)
Poll preview job; `result` has the same shape as `/api/header-preview` response.

---

## 4. NextGen API Endpoints

All NextGen endpoints use credentials from `.env.local`:
- `NEXTGEN_BASE_URL` (https://nextgen.madison88.com) + `NEXTGEN_USERNAME` / `NEXTGEN_PASSWORD` — **read** access
- `NEXTGEN_WRITE_BASE_URL` (:8443) + `NEXTGEN_WRITE_USERNAME` / `NEXTGEN_WRITE_PASSWORD` — **write** access (not used by any endpoint yet)

Under the hood:
- Login via `/Account/Login` (anti-forgery token + session cookie, auto re-login on 302/401)
- Data via `POST /PurchaseOrder/Read` (Kendo-style: `sort`, `filter`, `page`, `pageSize`)
- PO number field: `PrimaryUserDefinedFieldValuesTextUdf3`
- Global search via `GET /Search/GetSearchResults?criteria=...&searchEntityTypes=...`
- Product options via `POST /ProductOption/ProductOptionsGridRead`

### `GET /api/nextgen-po-lines` — **NEW: fetch PO lines** (also supports POST)
Fetch real PO lines from NextGen `PurchaseOrder/Read`.

**Query params** (GET) or JSON body (POST):
| Param | Description |
|---|---|
| `poNumber` | Exact PO number filter (e.g. `VUOUS0925B`) |
| `style` | Style/code substring — matched across CommodityName, Style, ProductCode, Material, SKU, refs |
| `pageSize` | Max lines returned (default 200, max 1000) |

**Response** (200):
```json
{
  "poNumber": "VUOUS0925B",
  "style": null,
  "count": 3,
  "lines": [
    {
      "id": "261710",
      "poNumber": "VUOUS0925B",
      "style": "M88130210",
      "color": "VUO - BLACK",
      "size": "One Size",
      "quantity": 1,
      "factory": "PT. UWU JUMP INDONESIA",
      "customer": "Vuori",
      "season": "",
      "unitCost": 12.735,
      "subtotal": null,
      "...": "plus all raw NextGen row fields (AR dates, UDFs, etc.)"
    }
  ]
}
```

Examples:
```
GET /api/nextgen-po-lines?poNumber=VUOUS0925B
GET /api/nextgen-po-lines?style=M88130210
GET /api/nextgen-po-lines?pageSize=100
POST { "style": "M88130210", "pageSize": 50 }
```

### `POST /api/validate-nextgen` — Validate uploaded lines against NextGen
Compares uploaded lines with real NextGen lines (matched / missing / extra / cost mismatches).

**Request**:
```json
{
  "poNumber": "M88PO123",
  "lines": [ { "style": "...", "color": "...", "size": "...", "quantity": 10, "subtotal": 123.45, "unitCost": 12.34 } ]
}
```

**Response**:
```json
{
  "poNumber": "M88PO123",
  "exists": true,
  "lines": [ /* NextGenPOLine[] */ ],
  "matched": [ /* lines found in NextGen with cost validation */ ],
  "missing": [ /* uploaded lines not found in NextGen */ ],
  "extra": [ /* NextGen lines not in upload */ ],
  "costMismatches": [ { "style": "...", "color": "...", "size": "...", "field": "subtotal|unitCost", "uploadValue": 1, "nextgenValue": 2, "difference": 1 } ]
}
```

### `GET /api/nextgen-latest-po` — Latest PO number
```json
{ "poNumber": "VUOUS0925B" }
```
(or `{ "poNumber": null, "message": "No PO found in NextGen" }`)

### `POST /api/nextgen-color-lookup` — Color name lookup by SKU/value
**Request**: `{ "sku": "NF0A8CGZ" }` or `{ "skus": ["NF0A8CGZ", "..."] }`
**Response**: `{ "results": { "NF0A8CGZ": "TNF Black", "...": null } }`

### Internal style-search pipeline (used by upload/extract, not exposed as its own route)
`NextGenCachedClient.searchVariant(style, color, brand)`:
1. Learning layer — user correction cache → color mapping cache → match cache
2. NextGen global search (`Search/GetSearchResults`) with buyer-style expansion
3. Product options fetch + color scoring (code match=100, exact=95, substring=85, token≥75% =55+)
4. Returns `NextGenStyleInfo` with product, colorName, factory, currency, unitCost, costingReference, sellPrice, candidates, matchScore
5. Successful matches are saved back to the learning cache

### `POST /api/diff-report` — **NEW: Buy file vs NextGen PO Line Dump diff**
Runs the same comparison engine as `scripts/diff-report.ts` (see §11) but over HTTP: processes an uploaded buy file through the ExcelEngine and diffs every line against the PO Line Data Dump ground truth.

**Request**: `multipart/form-data`
| Field | Type | Description |
|---|---|---|
| `file` | File | Excel buy file (.xlsx/.xls) |
| `dump` | File | Optional — PO Line Data Dump xlsx. If omitted, the server uses the dump path in `PO_LINE_DUMP` env var |
| `po` | string | Optional — manual PO number for buy files with no PO column (e.g. `M88PO2488`) |
| `costTolerance` | number | Optional — cost delta tolerance (default 0.01) |

**Response** (200):
```json
{
  "success": true,
  "buyFile": "Copy of FW2627 BULK BUY 4 - MADISON88 Rev1.xlsx",
  "dumpLines": 28464,
  "uploadLines": 152,
  "summary": { "MATCHED": 0, "QTY_MISMATCH": 135, "COST_MISMATCH": 17, "MISSING": 0, "EXTRA": 28446 },
  "rows": [
    {
      "matchType": "QTY_MISMATCH",
      "poNumber": "M88PO2488",
      "lineRef": "L1",
      "style": "M88129888",
      "color": "ROS-A01 OLIVE SHADOW",
      "size": "One Size",
      "uploadQty": 7,
      "dumpQty": 12,
      "uploadCost": null,
      "dumpCost": 41.82,
      "delta": null,
      "note": "qty Δ -5"
    }
  ]
}
```

Notes:
- Matching keys: style (M-code or buyer style) + color + size; sizes normalized (M ↔ Medium).
- Dump columns are auto-detected (handles `Product Buyer Style Number New` variants).
- `EXTRA` counts every dump line not referenced by the upload — with a 28K-row dump this dominates the counts; filter by the other match types for actionable rows.

CLI equivalent (writes a color-coded xlsx to `reports/`):
```bash
npx tsx scripts/diff-report.ts "<buy-file.xlsx>" [--po M88PO123] [dump.xlsx]
```

---

## 5. OCR (Gemini)

### `POST /api/ocr-gemini` — PDF/image PO extraction
**Request**: `multipart/form-data` with `file` (PDF/image) and optional `fillFromNextgen=true`.

**Response**:
```json
{
  "ocrResults": [ { "poNumber": "...", "style": "...", "color": "...", "size": "...", "quantity": 1, "factory": "...", "customer": "...", "season": "...", "exFtyDate": "...", "transportMethod": "...", "plant": "..." } ],
  "mergedResults": [ /* OCR results enriched with NextGen data when fillFromNextgen=true */ ],
  "filename": "po.pdf",
  "nextgenUsed": true,
  "nextgenError": null
}
```

---

## 6. Brand Config

### `GET /api/brand-config?brand=tnf` — MLO config for output generation
**Response**:
```json
{
  "brand": "TNF",
  "orders_template": "ORDERS_TNF",
  "lines_template": "LINES_TNF",
  "valid_statuses": ["Confirmed", "Cancelled"],
  "keyusers": { "KeyUser1": "...", "KeyUser2": "...", "KeyUser3": "", "KeyUser4": "...", "KeyUser5": "...", "KeyUser6": "", "KeyUser7": "", "KeyUser8": "" }
}
```

---

## 7. Admin API (requires `admin_session` cookie + role permission)

### `POST /api/admin/session` — Login
**Request**: `{ "password": "<ADMIN_PANEL_PASSWORD>" }` → sets httpOnly `admin_session` cookie (8h).
`GET /api/admin/session` → `{ "authenticated": true|false }`
`DELETE /api/admin/session` → logout.

### `GET /api/admin/runs` — Run history
### `GET /api/admin/audit` — Audit logs
### `GET /api/admin/factory` · `POST /api/admin/factory` · `DELETE /api/admin/factory`
Factory mapping CRUD. POST body: `{ "brand", "category", "product_supplier" }`. DELETE body: `{ "id" }`.
### `GET /api/admin/columns?customer=X` · `POST /api/admin/columns`
Column mapping (buy file column → internal field) list/upsert. POST body: `{ "customer", "buy_file_column", "internal_field", "notes" }`.
### `GET /api/admin/mlo` · `POST /api/admin/mlo`
MLO per-brand config. POST body: `{ "brand", "keyuser1", "keyuser2", "keyuser4", "keyuser5", "orders_template", "lines_template", "valid_statuses" }`.

---

## 8. Database Layer (Supabase, schema: `po_cutting`)

All app data lives in the **`po_cutting`** schema (21 tables) on the VPS Supabase at `http://5.223.78.194:8000`:

| Table | Purpose |
|---|---|
| `users` | App users & roles |
| `run_history` | One row per upload run (status, row counts, errors) |
| `audit_logs` | Event log (BUY_FILE_UPLOADED, WORKFLOW_STARTED, DATA_EXTRACTION_COMPLETE…) |
| `factory_mapping` | Brand+category → product supplier |
| `column_mapping` | Customer buy-file column → internal field |
| `mlo_mapping` | Per-brand templates, key users, valid statuses |
| `buy_file_templates` | Learned header layouts (headers → mapping) |
| `nextgen_match_cache` | Learning layer: style+color+brand → NextGen product |
| `nextgen_user_corrections` | Learning layer: user override memory |
| `header_mapping_cache` | Learning layer: file signature → header mapping |
| `color_mapping_cache` | Learning layer: raw color → NextGen color name/code |

Migration script: `npx tsx apply-db-migrations.ts` (idempotent, creates everything with RLS).

---

## 9. Environment Variables

| Var | Purpose |
|---|---|
| `NEXTGEN_BASE_URL` / `NEXTGEN_USERNAME` / `NEXTGEN_PASSWORD` | NextGen read access |
| `NEXTGEN_WRITE_BASE_URL` / `NEXTGEN_WRITE_USERNAME` / `NEXTGEN_WRITE_PASSWORD` | NextGen write access (unused yet) |
| `NEXTGEN_ENABLED=false` | Disable NextGen lookups |
| `NEXTGEN_SEARCH_DELAY_MS` | Delay between style searches (default 500) |
| `NEXTGEN_SEARCH_ENTITY_TYPES` | Entity types for global search |
| `NEXTGEN_REQUEST_TIMEOUT_MS` | Per-request timeout (default 20000) |
| `SUPABASE_URL` / `SUPABASE_ANON_KEY` / `SUPABASE_SERVICE_ROLE_KEY` | Supabase connection |
| `SUPABASE_DB_SCHEMA` | DB schema (defaults to `po_cutting`) |
| `ADMIN_PANEL_PASSWORD` | Admin panel + processing job auth key |
| `ALLOWED_UPLOAD_ORIGINS` | Comma-separated origin allowlist |
| `PROCESSING_API_URL` | Backend base used by processing proxies |
| `PROCESSING_INTERNAL_BASE_URL` | Internal base for VPS async jobs (default http://127.0.0.1:3003) |
| `GROQ_API_KEY` / `GROQ_MODEL` | Groq AI for header mapping |
| `OLLAMA_BASE_URL` | Local AI fallback for header mapping |
| `GEMINI_API_KEY` | Gemini OCR |

---

## 10. Quick cURL Examples

```bash
# Health
curl https://po-cutting.5-223-78-194.sslip.io/api/health

# Fetch PO lines for one PO
curl "http://localhost:3000/api/nextgen-po-lines?poNumber=VUOUS0925B"

# Fetch PO lines by style
curl "http://localhost:3000/api/nextgen-po-lines?style=M88130210&pageSize=50"

# Full upload
curl -X POST http://localhost:3000/api/upload \
  -H "Origin: http://localhost:3000" \
  -F "file=@BuyFile.xlsx" \
  -F "manualBrand=TNF"

# Header preview
curl -X POST http://localhost:3000/api/header-preview -F "file=@BuyFile.xlsx"

# Validate against NextGen
curl -X POST http://localhost:3000/api/validate-nextgen \
  -H "Content-Type: application/json" \
  -d '{"poNumber":"M88PO123","lines":[{"style":"NF0A8CGZ","color":"JK3","size":"M","quantity":10}]}'

# Latest PO
curl http://localhost:3000/api/nextgen-latest-po
```
