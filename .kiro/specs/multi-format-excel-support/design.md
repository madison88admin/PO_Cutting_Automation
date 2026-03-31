# Design Document: multi-format-excel-support

## Overview

This feature extends the PO cutting automation system to reliably ingest buy files from multiple brands
(TNF, Columbia, Arcteryx, 511 Tactical, Burton, EVO/Oyuki, Fox Racing, Haglofs) and future customers
whose Excel column naming conventions differ from current defaults. It adds UI controls for manual brand and destination overrides, improves date parsing to handle
Excel serial numbers and additional string formats, expands the column alias map in both the TypeScript
engine and the Python generator, and surfaces format-detection feedback in the Audit step.

All changes span four files:
- `src/lib/excel-engine.ts` — TypeScript engine (primary implementation)
- `src/app/api/upload/route.ts` — Next.js API route
- `src/components/Workflow.tsx` — React UI
- `ng_excel_generator.py` — Python CLI generator (must stay in parity)

---

## Architecture

The system follows a linear pipeline:

```mermaid
flowchart LR
    UI[Workflow.tsx\nACQUISITION step] -->|FormData POST| Route[route.ts\n/api/upload]
    Route -->|processBuyFile options| Engine[ExcelEngine\nexcel-engine.ts]
    Engine -->|ProcessedPO[] + FormatDetection| Route
    Route -->|JSON response\nformatDetection[]| UI
    UI -->|VALIDATE step| Audit[Audit display\ndetectedFormat + unmappedColumns]
```

The Python generator mirrors the engine's logic for CLI/batch use and must be kept in parity.

---

## Components and Interfaces

### 1. ExcelEngine — `getFallbackColumnAliases()` expansion

**Location:** `src/lib/excel-engine.ts`, method `getFallbackColumnAliases()` (~line 270).

Add the following entries to the returned object. Each entry maps a lowercase alias string to an
internal field name. Insert them in the appropriate semantic group:

```typescript
// product group
'style #': 'product',
'style no.': 'product',

// colour group
'colour code': 'colour',
'colour desc': 'colour',
'colour description': 'colour',
'color code': 'colour',
'color desc': 'colour',

// exFtyDate group
'ship date': 'exFtyDate',
'ship window': 'exFtyDate',
'planned ship date': 'exFtyDate',
'in-dc date': 'exFtyDate',
'in dc date': 'exFtyDate',
'dc arrival date': 'exFtyDate',

// poIssuanceDate group
'po date': 'poIssuanceDate',
'issue date': 'poIssuanceDate',

// cancelDate group
'cancellation date': 'cancelDate',

// quantity group
'qty ordered': 'quantity',
'total qty': 'quantity',
'units': 'quantity',

// productSupplier group
'supplier code': 'productSupplier',
'mfr code': 'productSupplier',

// transportLocation group
'ship to': 'transportLocation',
'ship to country': 'transportLocation',

// transportMethod group
'mode': 'transportMethod',
'freight mode': 'transportMethod',
'shipping method': 'transportMethod',

// category group
'division': 'category',
'product division': 'category',
'gender': 'category',
'gender code': 'category',

// season group
'season code': 'season',

// buyerPoNumber group
'buyer po': 'buyerPoNumber',
'buyer po #': 'buyerPoNumber',
'customer po': 'buyerPoNumber',
```

The existing `'vendor': 'productSupplier'` entry already covers "Vendor"; the new entries add
"Supplier Code" and "Mfr Code". The existing `'destination': 'transportLocation'` covers "Destination";
"Ship To" and "Ship To Country" are new.

**Python parity — `HEADER_ALIASES` in `ng_excel_generator.py`:**

The Python dict `HEADER_ALIASES` uses the same structure (field → list of aliases). Add the same
strings to the appropriate lists:

```python
"product":            [..., "style #", "style no."],
"colour":             [..., "colour code", "colour desc", "colour description",
                           "color code", "color desc"],
"orig_ex_fac":        [..., "ship date", "ship window", "planned ship date",
                           "in-dc date", "in dc date", "dc arrival date"],
"buy_date":           [..., "po date", "issue date"],   # maps to poIssuanceDate
"cancel_date":        [..., "cancellation date"],
"qty":                [..., "qty ordered", "total qty", "units"],
"vendor_code":        [..., "supplier code", "mfr code"],
"transport_location": [..., "ship to", "ship to country"],
"trans_cond":         [..., "mode", "freight mode", "shipping method"],
"category":           [..., "division", "product division", "gender", "gender code"],
"season":             [..., "season code"],
# new top-level key for buyerPoNumber aliases:
"buyer_po_number":    ["buyer po", "buyer po #", "customer po"],
```

> **Important:** `buyer_po_number` is a new top-level key in `HEADER_ALIASES`. This means the Python row-mapping logic must also handle it — wherever the Python code reads `buyer_po_number` from a row (or wherever `buyerPoNumber` is assigned), ensure the field is read using this alias key. Check `process_buy_file()` for the row-field extraction loop and add `buyer_po_number` to the list of fields that get extracted into the row dict.

---

### 2. Manual Brand Override

#### 2a. `Workflow.tsx` — new state + input field

Add a new state variable alongside the existing manual override states (~line 30):

```typescript
const [manualBrand, setManualBrand] = useState("");
```

In `handleStartUpload()`, append to `formData` after the existing manual fields (~line 65):

```typescript
if (manualBrand.trim()) formData.append("manualBrand", manualBrand.trim());
```

In the UPLOAD step JSX, add a new input block after the existing "Orders KeyDate" field:

```tsx
<div className="space-y-2">
  <label className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">
    Manual Brand
  </label>
  <input
    value={manualBrand}
    onChange={(e) => setManualBrand(e.target.value)}
    placeholder="tnf / columbia / arcteryx"
    className="w-full rounded-xl bg-white/5 border border-white/10 px-4 py-3 text-sm
               text-white placeholder:text-slate-600 focus:outline-none
               focus:ring-2 focus:ring-blue-500/40"
  />
</div>
```

#### 2b. `route.ts` — read and forward `manualBrand`

In the `POST` handler, after the existing `manualKeyDate` extraction (~line 75):

```typescript
const manualBrand = (formData.get("manualBrand")?.toString() || "").trim();
```

Pass it to `processBuyFile`:

```typescript
await engine.processBuyFile(buffer, {
    // ...existing options...
    manualBrand: manualBrand || undefined,
});
```

#### 2c. `ExcelEngine.processBuyFile()` — options type + per-row logic

Extend the `options` parameter type:

```typescript
options?: {
    // ...existing fields...
    manualBrand?: string;
}
```

Inside the row-processing loop, where `brand` is currently read (~line 1050):

```typescript
// Before:
const brand = this.stripBrackets(getVal('brand') || '');

// After:
const fileBrand = this.stripBrackets(getVal('brand') || '');
const brand = fileBrand || (options?.manualBrand?.trim() || '');
```

Row-level brand takes priority; `manualBrand` fills in only when the file has no brand value.

> **Note on brand detection conflict:** The engine detects `detectedCustomer` from column headers (not cell values) before row processing begins. `manualBrand` does NOT affect `detectedCustomer` — it only affects the per-row `brand` variable used for supplier/customer/template resolution. This means the column mapping loaded is still based on header-detected customer, while the brand used for output resolution uses `manualBrand` as fallback. This is intentional: column structure detection and brand-level output resolution are separate concerns.

#### 2d. Python parity — `ng_excel_generator.py`

Add `--manual-brand` to the `argparse` setup and thread it through to the per-row brand resolution
in `process_buy_file()` (or equivalent function). Apply the same precedence: file row brand wins,
`manual_brand` fills in when empty.

---

### 3. Manual Destination Override UI

The route already reads `manualDestination` and passes it to the engine. The engine already applies
it in the row loop. The only missing piece is the UI field.

#### 3a. `Workflow.tsx` — new state + input field

Add state (~line 30):

```typescript
const [manualDestination, setManualDestination] = useState("");
```

Append to `formData` in `handleStartUpload()`:

```typescript
if (manualDestination.trim()) formData.append("manualDestination", manualDestination.trim());
```

Add input in the UPLOAD step JSX (same pattern as Manual Brand):

```tsx
<div className="space-y-2">
  <label className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">
    Manual Destination
  </label>
  <input
    value={manualDestination}
    onChange={(e) => setManualDestination(e.target.value)}
    placeholder="VN / Vietnam / Indonesia"
    className="w-full rounded-xl bg-white/5 border border-white/10 px-4 py-3 text-sm
               text-white placeholder:text-slate-600 focus:outline-none
               focus:ring-2 focus:ring-blue-500/40"
  />
</div>
```

The engine's existing `normalizeTransportLocation()` already expands two-letter ISO codes via
`COUNTRY_NAME_MAP`, so no engine changes are needed for requirement 3.4.

#### 3b. Python parity

`ng_excel_generator.py` already has `--manual-destination` wired through the CLI. Verify the
per-row logic mirrors the TS engine: `manualDestination || getVal('transportLocation')`, then
`COUNTRY_NAME_MAP` expansion.

---

### 4. Robust Date Parsing — `parseDate()` expansion

**Location:** `src/lib/excel-engine.ts`, private method `parseDate()` (~line 730).

Current implementation handles: ISO `YYYY-MM-DD`, US slash/dash `M/D/YYYY` and `M-D-YYYY`,
and `DD-Mon-YYYY`.

Replace the method body with the expanded version:

```typescript
private parseDate(raw: string | Date | number | undefined): Date | null {
    if (raw == null || raw === '') return null;
    if (raw instanceof Date) return isNaN(raw.getTime()) ? null : raw;

    // Excel serial number: numeric >= 1 and <= 2958465 (1900-01-01 to 9999-12-31)
    if (typeof raw === 'number') {
        if (raw >= 1 && raw <= 2958465) {
            // Excel epoch is 1899-12-30; JS Date uses 1970-01-01
            const EXCEL_EPOCH_MS = new Date(1899, 11, 30).getTime();
            const date = new Date(EXCEL_EPOCH_MS + raw * 86400000);
            return isNaN(date.getTime()) ? null : date;
        }
        return null;
    }

    const text = String(raw).trim();
    if (!text) return null;

    // ISO: YYYY-MM-DD
    const isoMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (isoMatch) {
        const d = new Date(+isoMatch[1], +isoMatch[2] - 1, +isoMatch[3]);
        return isNaN(d.getTime()) ? null : d;
    }

    // ISO slash: YYYY/MM/DD
    const isoSlash = text.match(/^(\d{4})\/(\d{2})\/(\d{2})$/);
    if (isoSlash) {
        const d = new Date(+isoSlash[1], +isoSlash[2] - 1, +isoSlash[3]);
        return isNaN(d.getTime()) ? null : d;
    }

    // US slash or dash: M/D/YYYY or M-D-YYYY (month-first, US default)
    // Disambiguation rule: if first segment > 12, it MUST be a day (European DD/MM/YYYY).
    // If both segments ≤ 12 (ambiguous, e.g. 01/04/2026), default to US (M/D/YYYY) and
    // emit a WARNING so the operator is aware. A future `manualDateFormat` option can
    // override this default if a brand uses European dates exclusively.
    const usMatch = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (usMatch) {
        const a = +usMatch[1], b = +usMatch[2], yr = +usMatch[3];
        if (a > 12) {
            // Unambiguously European: DD/MM/YYYY
            const d = new Date(yr, b - 1, a);
            return isNaN(d.getTime()) ? null : d;
        }
        // Ambiguous or unambiguously US — default to US M/D/YYYY
        // Caller should emit a WARNING if both a ≤ 12 and b ≤ 12 (truly ambiguous)
        const d = new Date(yr, a - 1, b);
        return isNaN(d.getTime()) ? null : d;
    }

    // DD-Mon-YYYY: e.g. 17-Jun-2026
    const monDash = text.match(/^(\d{1,2})-([A-Za-z]+)-(\d{4})$/);
    if (monDash) {
        const months = ['jan','feb','mar','apr','may','jun',
                        'jul','aug','sep','oct','nov','dec'];
        const mi = months.findIndex(m => monDash[2].toLowerCase().startsWith(m));
        if (mi >= 0) {
            const d = new Date(+monDash[3], mi, +monDash[1]);
            return isNaN(d.getTime()) ? null : d;
        }
    }

    // Mon DD YYYY: e.g. Jun 17 2026
    const monSpace = text.match(/^([A-Za-z]+)\s+(\d{1,2})\s+(\d{4})$/);
    if (monSpace) {
        const months = ['jan','feb','mar','apr','may','jun',
                        'jul','aug','sep','oct','nov','dec'];
        const mi = months.findIndex(m => monSpace[1].toLowerCase().startsWith(m));
        if (mi >= 0) {
            const d = new Date(+monSpace[3], mi, +monSpace[2]);
            return isNaN(d.getTime()) ? null : d;
        }
    }

    return null;
}
```

The `getCellValue()` method must also pass numeric cell values through to `parseDate()` rather than
converting them to strings first. Check the cell-reading path and ensure numeric values are passed
as `number` type when the cell format is not already a Date.

Also emit a WARNING when a date string is ambiguous (both segments ≤ 12 in a slash/dash pattern).
**Important:** `parseDate()` returns only `Date | null` — it has no side-channel for warnings.
The ambiguous-date WARNING must be emitted by the **caller** (the row loop in `processBuyFile()`),
not from inside `parseDate()` itself. The pattern is:

```typescript
// In the row loop, after calling parseDate() on a slash/dash date string:
if (usMatch) {
    const a = +usMatch[1], b = +usMatch[2];
    if (a <= 12 && b <= 12) {
        // Truly ambiguous — defaulted to US M/D/YYYY, warn operator
        this.errors.push({
            field: fieldName, row: rowNumber,
            message: `Row ${rowNumber}: date "${text}" is ambiguous (could be M/D or D/M). Defaulted to US format (M/D/YYYY).`,
            severity: 'WARNING',
        });
    }
}
```

To support this, `parseDate()` should expose whether it took the ambiguous-US path. One clean approach:
return a tagged result `{ date: Date | null; ambiguous?: boolean }` from a private helper, and keep
the public `parseDate()` signature as `Date | null`. Alternatively, the row loop can detect ambiguity
independently by checking the raw string before calling `parseDate()`.

**Unparseable date warning:** In `processBuyFile()`, wherever date fields are read (exFtyDate,
cancelDate, poIssuanceDate, buyDate), add a warning when `parseDate()` returns null for a non-empty
raw value:

```typescript
const exFtyRaw = getRawVal('exFtyDate') || getRawVal('confirmedExFac');
if (exFtyRaw && !this.parseDate(exFtyRaw as any)) {
    this.errors.push({
        field: 'exFtyDate', row: rowNumber,
        message: `Row ${rowNumber} PO ${poNumber}: unrecognised date format "${exFtyRaw}".`,
        severity: 'WARNING',
    });
}
```

**Python parity — `_parse_date()` in `ng_excel_generator.py`:**

Replace the format list in `_parse_date()` with:

```python
def _parse_date(value: Any) -> datetime | None:
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    # Excel serial number
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        n = float(value)
        if 1 <= n <= 2958465:
            from datetime import timedelta
            epoch = datetime(1899, 12, 30)
            return epoch + timedelta(days=n)
        return None
    raw = str(value).strip()
    if not raw:
        return None
    for fmt in (
        "%Y-%m-%d",       # ISO
        "%Y/%m/%d",       # ISO slash
        # NOTE: %m/%d/%Y (US) intentionally appears before %d/%m/%Y (European).
        # This matches the TS engine's US-first disambiguation rule: when both
        # segments are ≤ 12 (e.g. "01/04/2026"), we default to US M/D/YYYY.
        # Do NOT reorder these two entries — it would silently break parity.
        "%m/%d/%Y",       # US slash  ← US-first (intentional, matches TS engine)
        "%m-%d-%Y",       # US dash
        "%d/%m/%Y",       # European slash (only reached when first segment > 12)
        "%d-%b-%Y",       # DD-Mon-YYYY
        "%d-%B-%Y",       # DD-Month-YYYY
        "%b %d %Y",       # Mon DD YYYY
        "%B %d %Y",       # Month DD YYYY
    ):
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    return None
```

> **Important — format order:** `%m/%d/%Y` appears before `%d/%m/%Y` intentionally. Python's `strptime` tries formats in order and returns on the first match, so for an ambiguous date like `"01/04/2026"` (both segments ≤ 12), the US interpretation (Jan 4) wins — consistent with the TS engine's disambiguation rule. A code comment is required in the implementation to prevent this from being "fixed" accidentally.

---

### 5. Format Detection Feedback

#### 5a. Engine — new return field

Extend the `processBuyFile` return type to include format detection:

```typescript
export interface FormatDetection {
    detectedCustomer: string;
    detectedFormat: string;
    unmappedColumns: string[];
}

// processBuyFile return type:
Promise<{
    data: ProcessedPO[];
    errors: ValidationError[];
    formatDetection: FormatDetection;
}>
```

Inside `processBuyFile()`, collect unmapped columns during header scanning. The existing code already
pushes a WARNING for each unmapped column; capture those column names separately:

```typescript
const unmappedColumnNames: string[] = [];

// In the headerRow.eachCell loop, where unmapped columns are detected:
if (!this.shouldSilentlyIgnoreHeader(headerText)) {
    unmappedColumnNames.push(headerText);   // ← add this
    this.errors.push({ ... });
}
```

Build the `detectedFormat` label at the end of `processBuyFile()` using a **private method** `buildFormatLabel()` on the `ExcelEngine` class (not an inline function, to keep it testable and consistent):

```typescript
private buildFormatLabel(customer: string): string {
    if (!customer || customer === 'DEFAULT') return 'Unknown format';
    const labels: Record<string, string> = {
        columbia: 'COL buy file',
        col: 'COL buy file',
        tnf: 'TNF buy file',
        'the north face': 'TNF buy file',
        arcteryx: 'Arcteryx buy file',
        "arc'teryx": 'Arcteryx buy file',
    };
    return labels[customer.toLowerCase()] || `${customer} buy file`;
}

const formatDetection: FormatDetection = {
    detectedCustomer: detectedCustomer || 'Unknown',
    detectedFormat: buildFormatLabel(detectedCustomer),
    unmappedColumns: unmappedColumnNames,
};

return { data: processedData, errors: this.errors, formatDetection };
```

#### 5b. `route.ts` — include `formatDetection` in response

Collect per-file format detection results:

```typescript
const perFileFormatDetection: Record<string, FormatDetection> = {};

// Inside the buy-file loop, after processBuyFile:
const { data, errors, formatDetection } = await engine.processBuyFile(buffer, { ... });
perFileFormatDetection[file.name] = formatDetection;
```

Include in the JSON response:

```typescript
return NextResponse.json({
    // ...existing fields...
    formatDetection: perFileFormatDetection,
});
```

#### 5c. `Workflow.tsx` — display in VALIDATE step

In the VALIDATE step JSX, add a format detection panel above the errors table. The panel reads
`uploadData?.formatDetection` (a `Record<filename, FormatDetection>`):

```tsx
{uploadData?.formatDetection && Object.entries(uploadData.formatDetection).map(
    ([fname, fd]: [string, any]) => (
    <div key={fname} className="mb-4 rounded-2xl bg-white/5 border border-white/10 px-6 py-4">
        <div className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-400 mb-1">
            {fname}
        </div>
        <div className="flex items-center gap-3 text-xs">
            <span className="text-slate-300">Detected:</span>
            <span className="font-black text-blue-400">{fd.detectedFormat}</span>
        </div>
        {fd.unmappedColumns.length > 0 ? (
            <div className="mt-2">
                <span className="text-[10px] font-black text-amber-400 uppercase tracking-widest">
                    Unmapped columns:
                </span>
                <div className="flex flex-wrap gap-2 mt-1">
                    {fd.unmappedColumns.map((col: string) => (
                        <span key={col}
                            className="px-2 py-0.5 rounded bg-amber-500/10 border border-amber-500/20
                                       text-amber-300 text-[10px] font-mono">
                            {col}
                        </span>
                    ))}
                </div>
            </div>
        ) : (
            <div className="mt-1 text-[10px] text-emerald-400 font-black uppercase tracking-widest">
                ✓ All columns mapped
            </div>
        )}
    </div>
))}
```

---

## Data Models

### Extended `processBuyFile` options

```typescript
interface ProcessBuyFileOptions {
    manualPurchaseOrder?: string;
    manualDestination?: string;
    manualProductRange?: string;
    manualTemplate?: string;
    manualComments?: string;
    manualKeyDate?: string;
    manualBrand?: string;           // NEW
    defaultQuantityIfMissing?: boolean;
    productSheetMap?: Record<string, ProductSheetRow[]>;
}
```

### `FormatDetection` interface (new)

```typescript
export interface FormatDetection {
    detectedCustomer: string;   // e.g. "DEFAULT", "Columbia", "Unknown"
    detectedFormat: string;     // e.g. "COL buy file", "Unknown format"
    unmappedColumns: string[];  // raw header strings not in alias map or ignore list
}
```

### API response shape additions

```typescript
// Existing response gains:
formatDetection: Record<string, FormatDetection>;  // keyed by filename
```

---

## Correctness Properties

*A property is a characteristic or behavior that should hold true across all valid executions of a
system — essentially, a formal statement about what the system should do. Properties serve as the
bridge between human-readable specifications and machine-verifiable correctness guarantees.*

### Property 1: Alias Coverage

*For any* alias string in the expanded `getFallbackColumnAliases()` map, passing a buy file whose
header row contains exactly that string should result in the corresponding internal field being
populated in the output (no unmapped-column warning for that header).

**Validates: Requirements 1.1**

### Property 2: Case-Insensitive Alias Matching

*For any* alias string in the alias map, transforming it by arbitrary capitalisation or by adding
leading/trailing whitespace should still resolve to the same internal field as the original alias.

**Validates: Requirements 1.3**

### Property 3: Silent Ignore List Stability

*For any* header string on the `shouldSilentlyIgnoreHeader` list, processing a buy file containing
that header should produce zero unmapped-column warnings for that header.

**Validates: Requirements 1.4**

### Property 4: Manual Brand Override Precedence

*For any* buy file and any `manualBrand` value: rows that contain a non-empty brand cell should use
the row-level brand in the output; rows that have an empty brand cell should use `manualBrand` in
the output.

**Validates: Requirements 2.4, 2.5**

### Property 5: Manual Destination Override and ISO Expansion

*For any* buy file and any `manualDestination` value: rows that have no transport-location cell
should use `manualDestination` (expanded via `COUNTRY_NAME_MAP` if it is a two-letter code) in the
output; rows that already have a non-empty transport-location cell should use that cell's value.

**Validates: Requirements 3.3, 3.4**

### Property 6: Excel Serial Date Conversion

*For any* integer N in the range [1, 2958465], `parseDate(N)` should return a `Date` object whose
value equals the date obtained by adding N days to the Excel epoch 1899-12-30.

**Validates: Requirements 4.1**

### Property 7: Date Format Round-Trip

*For any* date value that `parseDate()` successfully parses (covering all formats listed in
Requirement 4.2), formatting the result with `formatDateString()` and then re-parsing the formatted
string should produce a `Date` equal to the original.

**Validates: Requirements 4.2, 4.5**

### Property 8: Unparseable Date Warning

*For any* non-empty string that does not match any supported date format, passing it as a date field
value should cause `processBuyFile()` to include at least one WARNING-severity error referencing
that field and row.

**Validates: Requirements 4.3**

### Property 9: Format Detection Output Completeness

*For any* buy file processed by the engine, the returned `formatDetection` object should contain a
non-null `detectedCustomer`, a non-empty `detectedFormat` string, and an `unmappedColumns` array
that contains exactly the set of header strings that were neither in the alias map nor on the
silent-ignore list.

**Validates: Requirements 5.1, 5.4, 5.6**

### Property 10: TS–Python Alias Map Parity

*For any* alias string present in `getFallbackColumnAliases()` in the TypeScript engine, the same
alias string (lowercased, whitespace-normalised) should be present in `HEADER_ALIASES` in the
Python generator, mapping to the equivalent field.

**Validates: Requirements 6.1**

### Property 11: TS–Python Full Output Parity

*For any* buy file and any combination of override options (manualBrand, manualDestination,
manualPurchaseOrder, manualTemplate, manualKeyDate), the TypeScript engine and the Python generator
should produce ORDERS, LINES, and ORDER_SIZES sheets with identical row counts and identical values
in all mapped fields.

**Validates: Requirements 6.2, 6.3, 6.4, 6.5**

---

## Error Handling

| Scenario | Severity | Location | Behaviour |
|---|---|---|---|
| Unmapped column header | WARNING | `processBuyFile` header scan | Emit warning; add to `unmappedColumns`; continue processing |
| Unparseable date string | WARNING | `processBuyFile` row loop | Emit warning with field name, row, raw value; leave date blank |
| Excel serial out of range | WARNING | `parseDate()` | Return null; caller emits warning |
| `manualBrand` provided but file also has brand | — | Row loop | Silently use file brand; no warning |
| `manualDestination` is unrecognised code | — | `normalizeTransportLocation` | Pass through as-is; no warning |
| `detectedCustomer` cannot be determined | — | `processBuyFile` | Set to `"Unknown"`; `detectedFormat` = `"Unknown format"` |
| `formatDetection` missing from API response | — | `Workflow.tsx` | Guard with `uploadData?.formatDetection` optional chaining; render nothing |

---

## Testing Strategy

### Unit Tests

Focus on specific examples and edge cases:

- `parseDate()` with each of the 9 supported format strings (one example per format)
- `parseDate()` with Excel serial 1 → 1900-01-01, serial 44927 → 2023-01-01
- `parseDate()` with an invalid string returns `null`
- `getFallbackColumnAliases()` contains all new alias keys
- `normalizeTransportLocation("VN")` returns `"Vietnam"`
- `buildFormatLabel("col")` returns `"COL buy file"`
- `buildFormatLabel("")` returns `"Unknown format"`

### Property-Based Tests

Use a property-based testing library (recommended: **fast-check** for TypeScript,
**hypothesis** for Python). Each property test must run a minimum of **100 iterations**.

Tag format: `Feature: multi-format-excel-support, Property {N}: {property_text}`

**Property 1 — Alias Coverage**
```
// Feature: multi-format-excel-support, Property 1: alias coverage
fc.assert(fc.property(
    fc.constantFrom(...Object.keys(getFallbackColumnAliases())),
    (alias) => {
        const map = getFallbackColumnAliases();
        return map[alias] !== undefined && map[alias] !== '';
    }
), { numRuns: 100 });
```

**Property 2 — Case-Insensitive Alias Matching**
```
// Feature: multi-format-excel-support, Property 2: case-insensitive alias matching
fc.assert(fc.property(
    fc.constantFrom(...Object.keys(getFallbackColumnAliases())),
    fc.boolean(), fc.boolean(),
    (alias, upper, addSpace) => {
        const variant = (addSpace ? ' ' : '') + (upper ? alias.toUpperCase() : alias);
        const map = getFallbackColumnAliases();
        return map[variant.trim().toLowerCase()] === map[alias];
    }
), { numRuns: 200 });
```

**Property 6 — Excel Serial Date Conversion**
```
// Feature: multi-format-excel-support, Property 6: Excel serial date conversion
fc.assert(fc.property(
    fc.integer({ min: 1, max: 2958465 }),
    (serial) => {
        const result = parseDate(serial);
        if (!result) return false;
        const EPOCH = new Date(1899, 11, 30).getTime();
        const expected = new Date(EPOCH + serial * 86400000);
        return result.getTime() === expected.getTime();
    }
), { numRuns: 500 });
```

**Property 7 — Date Format Round-Trip**
```
// Feature: multi-format-excel-support, Property 7: date format round-trip
fc.assert(fc.property(
    fc.date({ min: new Date(1900, 0, 1), max: new Date(2099, 11, 31) }),
    (date) => {
        const formatted = formatDateString(date);
        const reparsed = parseDate(formatted);
        if (!reparsed) return false;
        return reparsed.getFullYear() === date.getFullYear()
            && reparsed.getMonth() === date.getMonth()
            && reparsed.getDate() === date.getDate();
    }
), { numRuns: 500 });
```

**Property 8 — Unparseable Date Warning**
```
// Feature: multi-format-excel-support, Property 8: unparseable date warning
fc.assert(fc.property(
    fc.string().filter(s => s.trim().length > 0 && parseDate(s) === null),
    (badDate) => {
        const engine = new ExcelEngine();
        // inject bad date into a minimal buy file and process
        const { errors } = engine.processRowWithDate(badDate);
        return errors.some(e => e.severity === 'WARNING' && e.message.includes(badDate));
    }
), { numRuns: 100 });
```

**Property 4 — Manual Brand Override Precedence** (integration-level)
```
// Feature: multi-format-excel-support, Property 4: manual brand override precedence
// For rows with no brand: output brand === manualBrand
// For rows with a brand: output brand === row brand (not manualBrand)
```

**Python — Hypothesis**
```python
# Feature: multi-format-excel-support, Property 6: Excel serial date conversion
@given(st.integers(min_value=1, max_value=2958465))
@settings(max_examples=500)
def test_excel_serial_conversion(serial):
    result = _parse_date(serial)
    assert result is not None
    epoch = datetime(1899, 12, 30)
    expected = epoch + timedelta(days=serial)
    assert result == expected

# Feature: multi-format-excel-support, Property 7: date format round-trip
@given(st.dates(min_value=date(1900, 1, 1), max_value=date(2099, 12, 31)))
@settings(max_examples=500)
def test_date_round_trip(d):
    formatted = d.strftime("%m/%d/%Y")
    result = _parse_date(formatted)
    assert result is not None
    assert result.date() == d
```

---

## New Brand Format Designs

### 6. 511 Tactical Buy File

**Column mapping additions** (already present in `getFallbackColumnAliases()` — verified):
- `'po#'` → `purchaseOrder`
- `'item#'` → `product`
- `'color code'` → `colour`
- `'size'` → `sizeName`
- `'quantity'` → `quantity`
- `'updated planned exit date'` → `exFtyDate`
- `'wh'` → `plant`

**Plant code → destination** (in `PLANT_COUNTRY_MAP`):
- `'3020'` → `'Sweden'`
- `'5001'` → `'Hong Kong'`

**PO suffix logic** (already implemented): `${manualPO}-${plant}` e.g. `PO002951-3020`.

**Brand maps** (already present):
- `BRAND_SUPPLIER_MAP["511 tactical"]` = `"PT. UWU JUMP INDONESIA"`
- `BRAND_CUSTOMER_MAP["511 tactical"]` = not set — customer name comes from `manualCustomer` or file

**No engine changes required** — all aliases and maps are already in place. This section documents the confirmed working state for regression purposes.

---

### 7. Burton Buy File

**Column mapping additions** (already present in `getFallbackColumnAliases()`):
- `'po#'` → `purchaseOrder`
- `'style'` → `product`
- `'color'` → `colour`
- `'size'` → `sizeName`
- `'quantity'` → `quantity`
- `'ex-factory date'` → `exFtyDate`
- `'country/region'` → `transportLocation`
- `'ship mode description'` → `transportMethod`
- `'factory'` → `vendorName`

**Country normalisation** (already in `COUNTRY_NAME_MAP`):
- `"CZECHIA"` → `"Czech Republic"`
- `"GREAT BRITAIN"` → `"UK"`

**PO format:** Burton PO values already contain the destination suffix (e.g. `PO002936-USA`). The engine reads them as-is from the `PO#` column. No plant-suffix logic applies.

**Silent ignore additions needed** — `'final destination'` is already on the ignore list (it contains internal DC codes, not country names). Verify `'seller name'`, `'supplier party id'`, `'sku #'`, `'material number'` are also silently ignored.

**Brand maps** (already present):
- `BRAND_SUPPLIER_MAP["burton"]` = `"PT. UWU JUMP INDONESIA"`
- `BRAND_CUSTOMER_MAP["burton"]` = `"Burton"`

**No engine changes required** for basic column mapping. The `'final destination'` ignore is already in place.

---

### 8. EVO/Oyuki Pivot-Format Buy File

This is the only format that requires new engine logic. EVO buy files use a **pivot layout** where:
- Fixed columns: `Style#`, `SKU`, `UPC`, `Name`, `Color`, `Size`, `JPY price`
- Pivot columns: one column per destination/order (e.g. `Japan Mountain Merch`, or `PO002954`)
- Each pivot cell value = quantity for that style/colour/size at that destination

**Detection logic** — add a new private method `detectPivotFormat()`:

```typescript
private detectPivotFormat(
    headerRow: ExcelJS.Row,
    aliases: Record<string, string>
): { isPivot: boolean; fixedCols: Record<string, number>; pivotCols: Array<{ colNumber: number; label: string }> } {
    const fixedCols: Record<string, number> = {};
    const pivotCols: Array<{ colNumber: number; label: string }> = [];
    const FIXED_PIVOT_HEADERS = new Set(['style#', 'sku', 'upc', 'name', 'color', 'colour', 'size', 'jpy price', 'bulk qty']);

    headerRow.eachCell((cell, colNumber) => {
        const raw = cell.value?.toString().trim() || '';
        const key = raw.toLowerCase();
        if (aliases[key]) {
            fixedCols[aliases[key]] = colNumber;
        } else if (FIXED_PIVOT_HEADERS.has(key)) {
            fixedCols[key] = colNumber;
        } else if (raw && !this.shouldSilentlyIgnoreHeader(raw)) {
            // Any non-empty, non-ignored, non-aliased column is a pivot column
            pivotCols.push({ colNumber, label: raw });
        }
    });

    const isPivot = pivotCols.length >= 1 &&
        (fixedCols['style#'] !== undefined || fixedCols['product'] !== undefined) &&
        (fixedCols['color'] !== undefined || fixedCols['colour'] !== undefined);

    return { isPivot, fixedCols, pivotCols };
}
```

**Row expansion logic** — when pivot format is detected, for each data row, iterate over `pivotCols` and emit one PO entry per pivot column where the cell value is a positive number:

```typescript
// Inside processBuyFile(), after header scanning:
const pivotResult = this.detectPivotFormat(headerRow, fallbackAliases);
if (pivotResult.isPivot) {
    // Process as pivot: each pivotCol becomes a separate PO
    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber <= headerRowNumber) return;
        for (const { colNumber, label } of pivotResult.pivotCols) {
            const qty = parseFloat(row.getCell(colNumber).value?.toString() || '0');
            if (!qty || qty <= 0) continue;
            const poNumber = manualPurchaseOrder || label;
            const destination = manualDestination || this.normalizeTransportLocation(label);
            // ... build POLine and POSize using fixedCols for style/colour/size
        }
    });
    return { data: Array.from(results.values()), errors: this.errors, formatDetection };
}
```

**`manualPurchaseOrder` + `manualDestination` are required** for EVO files unless the pivot column headers are already valid PO numbers. The UI should guide the operator to provide these.

**Python parity:** Add equivalent pivot detection and expansion to `ng_excel_generator.py`.

---

### 9. Fox Racing Buy File

**Column mapping** (already present in `getFallbackColumnAliases()`):
- `'purchasing document number'` → `purchaseOrder`
- `'plant'` → `plant`
- `'material description'` → `colour`
- `'grid value'` → `sizeName`
- `'order qty'` → `quantity`
- `'ex factory date'` → `exFtyDate`
- `'so order cancel date'` → `cancelDate`
- `'shipping instructions'` → `transportMethod`
- `'goods supplier name'` → `vendorName`

**Plant code destinations — NEEDS CONFIRMATION** (currently empty strings in `PLANT_COUNTRY_MAP`):

```typescript
// Fox Racing plant codes — destinations TBD, need confirmation from operator
'10':   '',   // TODO: confirm
'11':   '',   // TODO: confirm
'40':   '',   // TODO: confirm
'50':   '',   // TODO: confirm
'60':   '',   // TODO: confirm
```

Once confirmed, update `PLANT_COUNTRY_MAP` with the correct country strings. Until confirmed, the engine emits a WARNING for rows with these plant codes.

**Silent ignore additions** (already present): `'item number of purchasing document'`, `'vendor name'`, `'vendor number'`, `'goods supplier'`, `'purchasing group'`, etc.

**Brand maps** (already present):
- `BRAND_SUPPLIER_MAP["fox racing"]` = `"PT. UWU JUMP INDONESIA"`
- `BRAND_CUSTOMER_MAP["fox racing"]` = `"Fox Racing"`

**Action required:** Confirm Fox Racing plant code → country mapping with the operator, then update `PLANT_COUNTRY_MAP`.

---

### 10. Haglofs Buy File

**Column mapping** (already present in `getFallbackColumnAliases()`):
- `'po number'` → `purchaseOrder`
- `'style number'` → `product`
- `'style color'` → `colour`
- `'size'` → `sizeName`
- `'quantity'` → `quantity`
- `'delivery date'` → `exFtyDate`
- `'destination'` → `transportLocation`
- `'season'` → `season`
- `'customer'` → `customerName`

**Country normalisation** (already in `COUNTRY_NAME_MAP`): `SWEDEN`, `KOREA`, `JAPAN`, `HONG KONG`, `GERMANY`, etc. are all mapped.

**Brand maps** (already present):
- `BRAND_SUPPLIER_MAP["haglofs"]` = `"PT. UWU JUMP INDONESIA"`
- `BRAND_CUSTOMER_MAP["haglofs"]` = `"Haglofs"`

**No engine changes required** — all aliases and maps are already in place.

---

### 11. Extensible Brand Registration Pattern

To add a new brand, a developer must update only these locations:

| What to add | Location |
|---|---|
| Column header aliases | `getFallbackColumnAliases()` in `excel-engine.ts` |
| Supplier name | `BRAND_SUPPLIER_MAP` |
| Customer display name | `BRAND_CUSTOMER_MAP` |
| Orders template | `BRAND_ORDERS_TEMPLATE_MAP` |
| Lines template | `BRAND_LINES_TEMPLATE_MAP` |
| Key users | `BRAND_KEYUSER_MAP` |
| Plant code → country | `PLANT_COUNTRY_MAP` |
| Transport method strings | `TRANSPORT_MAP` |
| Python aliases | `HEADER_ALIASES` in `ng_excel_generator.py` |
| Python brand maps | Equivalent dicts in `ng_excel_generator.py` |

No changes to `processBuyFile()`, `resolveSupplier()`, `resolveCustomer()`, or any other method are needed for a standard new brand.

---

## Updated Correctness Properties

### Property 12: 511 Tactical Plant Code Suffix

*For any* 511 Tactical buy file row with a `WH` value present in `PLANT_COUNTRY_MAP`, the output PO number SHALL contain the `WH` value as a suffix (e.g. `PO002951-3020`) and `transportLocation` SHALL equal the mapped country.

**Validates: Requirement 7.2**

### Property 13: Burton Country Normalisation

*For any* Burton buy file row where `Country/Region` contains a value present in `COUNTRY_NAME_MAP`, the output `transportLocation` SHALL equal the normalised country name.

**Validates: Requirement 8.2**

### Property 14: EVO Pivot Expansion

*For any* EVO pivot-format buy file, the total number of output PO size rows SHALL equal the sum of all positive-quantity cells across all pivot columns.

**Validates: Requirement 9.2**

### Property 15: Fox Racing Plant Code Warning

*For any* Fox Racing buy file row where the `Plant` value maps to an empty string in `PLANT_COUNTRY_MAP`, the output SHALL contain at least one WARNING-severity error referencing that plant code.

**Validates: Requirement 12.3**

### Property 16: Brand Map Completeness

*For any* brand key present in `BRAND_SUPPLIER_MAP`, the same key SHALL also be present in `BRAND_CUSTOMER_MAP`, ensuring no brand produces a supplier without a customer name.

**Validates: Requirement 13.2**
