# Requirements Document

## Introduction

The multi-format-excel-support feature extends the PO cutting automation system to reliably ingest buy files from multiple brands (TNF, Columbia, Arcteryx, 511 Tactical, Burton, EVO/Oyuki, Fox Racing, Haglofs) and any future customers whose Excel column naming conventions differ from the current defaults. It adds UI controls for manual brand and destination overrides, improves date parsing to handle Excel serial numbers and additional string formats, expands the column alias map in both the TypeScript engine and the Python generator, and surfaces format-detection feedback in the Audit step so operators can immediately see what was detected and what columns were unrecognised. All engine changes must be kept in sync between `src/lib/excel-engine.ts` (and its supporting files) and `ng_excel_generator.py`.

## Glossary

- **Engine**: The TypeScript class `ExcelEngine` in `src/lib/excel-engine.ts` that parses buy files and produces NextGen PO output.
- **Python_Generator**: The Python script `ng_excel_generator.py` that performs equivalent processing for CLI/batch use.
- **Buy_File**: An `.xlsx` file uploaded by the operator containing raw purchase order data from a brand or customer.
- **Alias_Map**: The lookup table (in-code and/or DB) that maps raw buy-file column headers to internal field names.
- **MOCK_COLUMNS**: The active column mapping source in `src/lib/db/columnMapping.ts` used when `isMock=true`.
- **Header_Detection**: The logic that scans up to row 80 of a worksheet to find the row most likely to be the header row.
- **Detected_Format**: The customer/brand identity inferred by the Engine from the buy file's column headers and cell values.
- **Unmapped_Column**: A column header present in the buy file that has no entry in the Alias_Map and is not on the silent-ignore list.
- **Manual_Brand**: A brand identifier supplied by the operator via the UI to override or supplement brand detection.
- **Manual_Destination**: A transport destination (country name or code) supplied by the operator via the UI to override the value parsed from the buy file.
- **Excel_Serial_Date**: An integer or float stored in an Excel cell representing a date as the number of days since 1900-01-01.
- **Audit_Step**: The "Validate" stage of the Workflow UI where errors, warnings, and format-detection results are displayed.
- **NextGen_Template**: The standardised output workbook containing ORDERS, LINES, and ORDER_SIZES sheets.
- **TS_Python_Parity**: The requirement that any logic change in the Engine is mirrored in the Python_Generator and vice versa.

---

## Requirements

### Requirement 1: Expanded Column Alias Map

**User Story:** As an operator, I want the system to recognise more column name variants from real-world buy files, so that I do not have to manually rename columns before uploading.

#### Acceptance Criteria

1. THE Engine SHALL resolve each of the following additional column header variants to the correct internal field, in addition to all currently mapped aliases:
   - `"Style"`, `"Style #"`, `"Style No."` → `product`
   - `"Colour Code"`, `"Colour Desc"`, `"Colour Description"` → `colour`
   - `"Color Code"`, `"Color Desc"` → `colour`
   - `"Ship Date"`, `"Ship Window"`, `"Planned Ship Date"` → `exFtyDate`
   - `"In-DC Date"`, `"In DC Date"`, `"DC Arrival Date"` → `exFtyDate`
   - `"PO Date"`, `"Issue Date"` → `poIssuanceDate`
   - `"Cancel"`, `"Cancellation Date"` → `cancelDate`
   - `"Qty Ordered"`, `"Total Qty"`, `"Units"` → `quantity`
   - `"Vendor"`, `"Supplier Code"`, `"Mfr Code"` → `productSupplier`
   - `"Destination"`, `"Ship To"`, `"Ship To Country"` → `transportLocation`
   - `"Mode"`, `"Freight Mode"`, `"Shipping Method"` → `transportMethod`
   - `"Division"`, `"Product Division"` → `category`
   - `"Gender"`, `"Gender Code"` → `category`
   - `"Season Code"` → `season`
   - `"Buyer PO"`, `"Buyer PO #"`, `"Customer PO"` → `buyerPoNumber`
2. THE Python_Generator SHALL apply the same alias expansions as the Engine (TS_Python_Parity).
3. WHEN a buy file column header matches an alias via case-insensitive, whitespace-normalised comparison, THE Engine SHALL map it to the correct internal field without requiring an exact-case match.
4. THE Engine SHALL continue to silently ignore all headers currently on the silent-ignore list; the expanded alias map SHALL NOT cause previously-ignored headers to generate unmapped-column warnings.

---

### Requirement 2: Manual Brand Override

**User Story:** As an operator, I want to enter a brand name in the UI before uploading, so that I can force the correct brand when the buy file has no brand column or contains an ambiguous value.

#### Acceptance Criteria

1. THE Workflow_UI SHALL display a "Manual Brand" text input field on the Acquisition (UPLOAD) step, alongside the existing manual override fields.
2. WHEN the operator enters a value in the Manual Brand field and submits the form, THE Workflow_UI SHALL include the value as a `manualBrand` form field in the multipart POST request to `/api/upload`.
3. WHEN `manualBrand` is present in the request, THE Upload_Route SHALL pass it to the Engine's `processBuyFile` options.
4. WHEN `manualBrand` is provided and the buy file row does not contain a non-empty brand value, THE Engine SHALL use `manualBrand` as the brand for all rows in that file.
5. WHEN `manualBrand` is provided and a buy file row also contains a non-empty brand value, THE Engine SHALL use the row-level brand value (manual override does not overwrite explicit file data).
6. THE Python_Generator SHALL accept a `--manual-brand` CLI argument that applies the same override logic (TS_Python_Parity).

---

### Requirement 3: Manual Destination Override UI

**User Story:** As an operator, I want to enter a transport destination in the UI before uploading, so that I can set the destination when the buy file omits it or uses an unrecognised code.

#### Acceptance Criteria

1. THE Workflow_UI SHALL display a "Manual Destination" text input field on the Acquisition (UPLOAD) step.
2. WHEN the operator enters a value in the Manual Destination field and submits the form, THE Workflow_UI SHALL include the value as a `manualDestination` form field in the multipart POST request to `/api/upload`.
3. THE Upload_Route already reads `manualDestination` from the form data and passes it to the Engine; THE Engine SHALL apply it as the `transportLocation` for all rows where the buy file does not supply a non-empty destination value.
4. WHEN `manualDestination` is a two-letter ISO country code present in the Engine's `COUNTRY_NAME_MAP`, THE Engine SHALL expand it to the full country name before writing it to the output.
5. THE Python_Generator SHALL accept a `--manual-destination` CLI argument that applies the same override logic (TS_Python_Parity).

---

### Requirement 4: Robust Date Parsing

**User Story:** As an operator, I want the system to correctly parse dates regardless of whether they are stored as Excel serial numbers, ISO strings, US-format strings, or human-readable strings like "17-Jun-2026", so that date fields in the output are never blank due to an unrecognised format.

#### Acceptance Criteria

1. WHEN a date cell contains an Excel serial number (a numeric value ≥ 1 and ≤ 2958465, representing dates between 1900-01-01 and 9999-12-31), THE Engine SHALL convert it to a `Date` object using the standard Excel epoch (1899-12-30 base) before formatting.
2. THE Engine SHALL parse date strings in at least the following formats without error:
   - `YYYY-MM-DD` (ISO)
   - `M/D/YYYY` and `MM/DD/YYYY` (US slash)
   - `M-D-YYYY` and `MM-DD-YYYY` (US dash)
   - `DD-Mon-YYYY` (e.g. `17-Jun-2026`)
   - `Mon DD YYYY` (e.g. `Jun 17 2026`)
   - `DD/MM/YYYY` (European slash)
   - `YYYY/MM/DD` (ISO slash variant)
3. WHEN a date string does not match any supported format, THE Engine SHALL emit a WARNING-severity validation error identifying the field name, row number, and the raw value that could not be parsed, rather than silently producing a blank date.
4. THE Python_Generator SHALL support the same set of date formats and emit equivalent warnings for unparseable values (TS_Python_Parity).
5. FOR ALL date values that can be parsed by the Engine, formatting then re-parsing the formatted string SHALL produce a date equal to the original (round-trip property).

---

### Requirement 5: Detected Format Feedback in Audit Step

**User Story:** As an operator, I want the Audit step to show me what customer/format was detected from the uploaded file and list any column headers that were not recognised, so that I can identify mapping gaps before committing the output.

#### Acceptance Criteria

1. WHEN the Engine processes a buy file, THE Engine SHALL record the detected customer/brand identity (the value used as `Detected_Format`) and the list of column headers that were neither mapped nor silently ignored.
2. THE Upload_Route SHALL include a `formatDetection` object in the JSON response for each processed file, containing:
   - `detectedCustomer`: the resolved customer/brand string
   - `detectedFormat`: a human-readable label (e.g. `"COL buy file"`, `"Arcteryx buy file"`, `"Unknown"`)
   - `unmappedColumns`: an array of column header strings that were not resolved
3. THE Workflow_UI SHALL display the `formatDetection` data on the Audit (VALIDATE) step, showing the detected format label and, if `unmappedColumns` is non-empty, a list of those column names with a WARNING indicator.
4. WHEN `unmappedColumns` is non-empty, THE Workflow_UI SHALL display each unmapped column name so the operator can decide whether to add a mapping or ignore it.
5. WHEN all columns in the buy file are either mapped or silently ignored, THE Workflow_UI SHALL display a confirmation that no unmapped columns were found.
6. IF the Engine cannot determine a customer/brand from the file, THEN THE Engine SHALL set `detectedCustomer` to `"Unknown"` and `detectedFormat` to `"Unknown format"`.

---

### Requirement 6: TS–Python Parity Enforcement

**User Story:** As a developer, I want all engine logic changes to be applied to both the TypeScript engine and the Python generator, so that CLI batch runs and web-based runs produce identical output for the same input file.

#### Acceptance Criteria

1. THE Engine and THE Python_Generator SHALL share equivalent alias maps: every alias entry present in one SHALL have a corresponding entry in the other.
2. THE Engine and THE Python_Generator SHALL apply the same date-parsing logic, including support for all formats listed in Requirement 4 and the same Excel serial number conversion.
3. THE Engine and THE Python_Generator SHALL apply the same manual-brand and manual-destination override logic as specified in Requirements 2 and 3.
4. THE Engine and THE Python_Generator SHALL produce `detectedCustomer` and `unmappedColumns` values using the same detection algorithm.
5. WHEN given the same buy file and the same override parameters, THE Engine and THE Python_Generator SHALL produce output ORDERS, LINES, and ORDER_SIZES sheets with identical row counts and identical values in all mapped fields.

---

### Requirement 7: 511 Tactical Buy File Support

**User Story:** As an operator, I want to upload a 511 Tactical buy file and have the system correctly read the PO number, warehouse code, style, colour, size, quantity, and ex-factory date, so that I do not need to manually rename any columns.

#### Acceptance Criteria

1. THE Engine SHALL recognise the following 511 Tactical column headers and map them to the correct internal fields:
   - `PO#` → `purchaseOrder`
   - `Item#` → `product`
   - `Color Code` → `colour`
   - `Size` → `sizeName`
   - `Quantity` → `quantity`
   - `Updated Planned Exit Date` → `exFtyDate`
   - `WH` → `plant` (warehouse code used to derive destination country)
2. WHEN the `WH` column contains a known warehouse code (e.g. `3020` → Sweden, `5001` → Hong Kong), THE Engine SHALL derive the `transportLocation` from `PLANT_COUNTRY_MAP` and append the plant code to the PO number suffix (e.g. `PO002951-3020`).
3. WHEN `manualPurchaseOrder` is provided, THE Engine SHALL use it as the base PO number and append the plant code suffix from the `WH` column.
4. THE Engine SHALL set `productSupplier` to `"PT. UWU JUMP INDONESIA"` for all 511 Tactical rows (via `BRAND_SUPPLIER_MAP["511 tactical"]`).
5. THE Python_Generator SHALL apply the same column mappings and plant-code suffix logic (TS_Python_Parity).

---

### Requirement 8: Burton Buy File Support

**User Story:** As an operator, I want to upload a Burton buy file and have the system correctly read the PO number, destination country, style, colour, size, quantity, and ex-factory date, including PO values that already contain a destination suffix (e.g. `PO002936-USA`).

#### Acceptance Criteria

1. THE Engine SHALL recognise the following Burton column headers and map them to the correct internal fields:
   - `PO#` → `purchaseOrder`
   - `Style` → `product`
   - `Color` → `colour`
   - `Size` → `sizeName`
   - `Quantity` → `quantity`
   - `Ex-Factory Date` → `exFtyDate`
   - `Country/Region` → `transportLocation`
   - `Ship Mode Description` → `transportMethod`
   - `Factory` → `vendorName`
2. WHEN the `Country/Region` column contains a full country name (e.g. `CZECHIA`, `GREAT BRITAIN`), THE Engine SHALL normalise it to the canonical name via `COUNTRY_NAME_MAP` (e.g. `Czech Republic`, `UK`).
3. WHEN the PO value from the file already contains a destination suffix (e.g. `PO002936-USA`), THE Engine SHALL use it as-is without appending an additional plant suffix.
4. THE Engine SHALL set `productSupplier` to `"PT. UWU JUMP INDONESIA"` for all Burton rows (via `BRAND_SUPPLIER_MAP["burton"]`).
5. THE Python_Generator SHALL apply the same column mappings and country normalisation (TS_Python_Parity).

---

### Requirement 9: EVO/Oyuki Pivot-Format Buy File Support

**User Story:** As an operator, I want to upload an EVO/Oyuki buy file where destination quantities are stored as pivot columns (one column per destination/order), and have the system expand each pivot column into a separate PO line with the correct destination and quantity.

#### Acceptance Criteria

1. THE Engine SHALL detect the EVO/Oyuki pivot format when a worksheet contains columns `Style#`, `Name`, `Color`, `Size`, and one or more columns whose header matches the pattern `PO\d{6}` or a known destination label (e.g. `Japan Mountain Merch`).
2. WHEN the pivot format is detected, THE Engine SHALL treat each pivot column header as a separate PO identifier and each non-zero cell value in that column as the quantity for that PO/style/colour/size combination.
3. WHEN `manualPurchaseOrder` is provided, THE Engine SHALL use it as the PO number for all pivot columns (overriding the column header as PO source).
4. WHEN `manualDestination` is provided, THE Engine SHALL use it as the `transportLocation` for all pivot columns.
5. THE Engine SHALL map `Bulk Qty` → `quantity` for non-pivot rows (standard EVO format).
6. THE Python_Generator SHALL apply the same pivot-detection and expansion logic (TS_Python_Parity).

---

### Requirement 10: Fox Racing Buy File Support

**User Story:** As an operator, I want to upload a Fox Racing buy file and have the system correctly read the purchasing document number, plant code, material description, order quantity, and ex-factory date, with plant codes resolved to destination countries.

#### Acceptance Criteria

1. THE Engine SHALL recognise the following Fox Racing column headers and map them to the correct internal fields:
   - `Purchasing Document Number` → `purchaseOrder`
   - `Plant` → `plant`
   - `Material Description` → `colour`
   - `Grid Value` → `sizeName`
   - `Order Qty` → `quantity`
   - `Ex Factory Date` → `exFtyDate`
   - `So Order Cancel Date` → `cancelDate`
   - `Shipping Instructions` → `transportMethod`
   - `Goods Supplier Name` → `vendorName`
2. WHEN the `Plant` column contains a Fox Racing plant code (`10`, `11`, `40`, `50`, `60`), THE Engine SHALL look up the destination country from `PLANT_COUNTRY_MAP` and append the plant code to the PO number suffix.
3. THE Engine SHALL silently ignore the `Item Number of Purchasing Document` column (it is a line sequence number, not a product code).
4. THE Engine SHALL set `productSupplier` to `"PT. UWU JUMP INDONESIA"` for all Fox Racing rows (via `BRAND_SUPPLIER_MAP["fox racing"]`).
5. THE Python_Generator SHALL apply the same column mappings and plant-code suffix logic (TS_Python_Parity).

---

### Requirement 11: Haglofs Buy File Support

**User Story:** As an operator, I want to upload a Haglofs buy file and have the system correctly read the PO number, destination, style number, colour, size, quantity, and delivery date.

#### Acceptance Criteria

1. THE Engine SHALL recognise the following Haglofs column headers and map them to the correct internal fields:
   - `PO Number` → `purchaseOrder`
   - `Style Number` → `product`
   - `Style Color` → `colour`
   - `Size` → `sizeName`
   - `Quantity` → `quantity`
   - `Delivery Date` → `exFtyDate`
   - `Destination` → `transportLocation`
   - `Season` → `season`
   - `Customer` → `customerName`
2. WHEN the `Destination` column contains a full uppercase country name (e.g. `SWEDEN`, `KOREA`), THE Engine SHALL normalise it to the canonical mixed-case name via `COUNTRY_NAME_MAP`.
3. THE Engine SHALL set `productSupplier` to `"PT. UWU JUMP INDONESIA"` for all Haglofs rows (via `BRAND_SUPPLIER_MAP["haglofs"]`).
4. THE Python_Generator SHALL apply the same column mappings and country normalisation (TS_Python_Parity).

---

### Requirement 12: Fox Racing Plant Code Destination Mapping

**User Story:** As an operator, I want Fox Racing plant codes (`10`, `11`, `40`, `50`, `60`) to resolve to the correct destination countries, so that the output PO has the correct `transportLocation`.

#### Acceptance Criteria

1. THE `PLANT_COUNTRY_MAP` in `excel-engine.ts` SHALL contain confirmed destination country values for Fox Racing plant codes `10`, `11`, `40`, `50`, and `60`.
2. WHEN a Fox Racing buy file row has a `Plant` value of `10`, `11`, `40`, `50`, or `60`, THE Engine SHALL resolve the destination country from `PLANT_COUNTRY_MAP` and write it to `transportLocation`.
3. WHEN a Fox Racing plant code has no confirmed destination (i.e. the map value is an empty string), THE Engine SHALL emit a WARNING-severity error indicating the plant code has no destination mapping, and leave `transportLocation` blank.
4. THE Python_Generator SHALL use the same plant code → country mapping (TS_Python_Parity).

---

### Requirement 13: Extensible Brand Registration

**User Story:** As a developer, I want to add support for a new brand's buy file format by adding entries to the alias map, brand maps, and plant map — without modifying any other engine logic — so that the system scales to new customers without code restructuring.

#### Acceptance Criteria

1. THE Engine's `getFallbackColumnAliases()` SHALL be the single source of truth for column header → internal field mapping; adding a new brand's column headers to this map SHALL be sufficient to enable column detection for that brand.
2. THE `BRAND_SUPPLIER_MAP`, `BRAND_CUSTOMER_MAP`, `BRAND_ORDERS_TEMPLATE_MAP`, and `BRAND_LINES_TEMPLATE_MAP` SHALL each accept new brand keys without requiring changes to any other method.
3. THE `PLANT_COUNTRY_MAP` SHALL accept new plant codes or plant name patterns without requiring changes to any other method.
4. THE `TRANSPORT_MAP` SHALL accept new transport method strings without requiring changes to any other method.
5. THE Python_Generator SHALL mirror all four maps (`HEADER_ALIASES`, supplier map, customer map, plant map) so that adding a new brand to the TS engine requires a corresponding update to the Python generator only in those map definitions.
