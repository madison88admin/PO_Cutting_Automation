                                                                                                                                     # Tasks: Multi-Format Excel Support

## Phase 1 — Column Alias Expansion (Requirements 1, 7–11, 13)

- [x] 1. Add missing alias entries to `getFallbackColumnAliases()` in `src/lib/excel-engine.ts`
  - [x] 1.1 Add product group: `'style #'`, `'style no.'`
  - [x] 1.2 Add colour group: `'colour code'`, `'colour desc'`, `'colour description'`, `'color code'`, `'color desc'`
  - [x] 1.3 Add exFtyDate group: `'ship date'`, `'ship window'`, `'planned ship date'`, `'in-dc date'`, `'in dc date'`, `'dc arrival date'`
  - [x] 1.4 Add poIssuanceDate group: `'po date'`, `'issue date'`
  - [x] 1.5 Add cancelDate group: `'cancellation date'`
  - [x] 1.6 Add quantity group: `'qty ordered'`, `'total qty'`, `'units'`
  - [x] 1.7 Add productSupplier group: `'supplier code'`, `'mfr code'`
  - [x] 1.8 Add transportLocation group: `'ship to country'`
  - [x] 1.9 Add transportMethod group: `'mode'`, `'freight mode'`, `'shipping method'`
  - [x] 1.10 Add category group: `'division'`, `'product division'`, `'gender'`, `'gender code'`
  - [x] 1.11 Add season group: `'season code'`
  - [x] 1.12 Add buyerPoNumber group: `'buyer po'`, `'buyer po #'`, `'customer po'`

- [x] 2. Mirror alias additions in `HEADER_ALIASES` in `ng_excel_generator.py` (TS_Python_Parity for Requirement 1)
  - [x] 2.1 Add all new aliases from Task 1 to the corresponding field lists in `HEADER_ALIASES`
  - [x] 2.2 Add `"buyer_po_number"` as a new top-level key in `HEADER_ALIASES` with values `["buyer po", "buyer po #", "customer po"]`
  - [x] 2.3 Ensure the row-extraction loop in `process_buy_file()` reads `buyer_po_number` from the row dict

## Phase 2 — Manual Brand Override (Requirement 2)

- [x] 3. Add `manualBrand` option to `ExcelEngine.processBuyFile()` in `src/lib/excel-engine.ts`
  - [x] 3.1 Add `manualBrand?: string` to the `options` parameter type
  - [x] 3.2 In the row loop, change `const brand = ...` to use `fileBrand || manualBrand` (file brand takes priority)

- [x] 4. Read and forward `manualBrand` in `src/app/api/upload/route.ts`
  - [x] 4.1 Extract `manualBrand` from `formData`
  - [x] 4.2 Pass `manualBrand` to `engine.processBuyFile()` options

- [x] 5. Add Manual Brand input field to `src/components/Workflow.tsx`
  - [x] 5.1 Add `manualBrand` state variable
  - [x] 5.2 Append `manualBrand` to `formData` in `handleStartUpload()`
  - [x] 5.3 Add input field in the UPLOAD step JSX

- [x] 6. Add `--manual-brand` CLI argument to `ng_excel_generator.py` (TS_Python_Parity)
  - [x] 6.1 Add `--manual-brand` to `argparse`
  - [x] 6.2 Apply same `fileBrand || manualBrand` precedence in the row loop

## Phase 3 — Manual Destination UI (Requirement 3)

- [x] 7. Add Manual Destination input field to `src/components/Workflow.tsx`
  - [x] 7.1 Add `manualDestination` state variable
  - [x] 7.2 Append `manualDestination` to `formData` in `handleStartUpload()`
  - [x] 7.3 Add input field in the UPLOAD step JSX (route already reads and forwards this field)

## Phase 4 — Robust Date Parsing (Requirement 4)

- [x] 8. Expand `parseDate()` in `src/lib/excel-engine.ts`
  - [x] 8.1 Add Excel serial number handling: numeric input in range [1, 2958465] → Date via epoch 1899-12-30
  - [x] 8.2 Add ISO slash format: `YYYY/MM/DD`
  - [x] 8.3 Add `Mon DD YYYY` format (e.g. `Jun 17 2026`)
  - [x] 8.4 Update `getCellValue()` to pass numeric cell values as `number` type (not string) when cell is not already a Date

- [x] 9. Add unparseable date warnings in `processBuyFile()` row loop
  - [x] 9.1 After reading `exFtyDate`, emit WARNING when raw value is non-empty but `parseDate()` returns null
  - [x] 9.2 After reading `cancelDate`, emit WARNING when raw value is non-empty but `parseDate()` returns null

- [x] 10. Expand `_parse_date()` in `ng_excel_generator.py` (TS_Python_Parity)
  - [x] 10.1 Add Excel serial number handling (same epoch, same range)
  - [x] 10.2 Add all format strings from Requirement 4.2 in US-first order (comment required explaining order)

## Phase 5 — Format Detection Feedback (Requirement 5)

- [x] 11. Add `FormatDetection` interface and update `processBuyFile()` return type in `src/lib/excel-engine.ts`
  - [x] 11.1 Export `FormatDetection` interface with `detectedCustomer`, `detectedFormat`, `unmappedColumns`
  - [x] 11.2 Collect `unmappedColumnNames` during header scan (alongside existing WARNING push)
  - [x] 11.3 Add private `buildFormatLabel(customer: string): string` method
  - [x] 11.4 Return `formatDetection` object from `processBuyFile()`

- [x] 12. Include `formatDetection` in API response in `src/app/api/upload/route.ts`
  - [x] 12.1 Collect per-file `formatDetection` from each `processBuyFile()` call
  - [x] 12.2 Add `formatDetection: perFileFormatDetection` to the JSON response

- [x] 13. Display format detection in VALIDATE step in `src/components/Workflow.tsx`
  - [x] 13.1 Add format detection panel above the errors table showing `detectedFormat` label
  - [x] 13.2 Show unmapped columns list with amber WARNING indicator when `unmappedColumns.length > 0`
  - [x] 13.3 Show green "All columns mapped" confirmation when `unmappedColumns` is empty

## Phase 6 — EVO/Oyuki Pivot Format (Requirement 9)

- [ ] 14. Implement pivot format detection in `src/lib/excel-engine.ts`
  - [ ] 14.1 Add private `detectPivotFormat()` method that identifies fixed columns vs pivot columns
  - [ ] 14.2 Pivot columns are any non-empty, non-aliased, non-ignored headers after the fixed columns

- [ ] 15. Implement pivot row expansion in `processBuyFile()`
  - [ ] 15.1 When pivot format detected, iterate pivot columns and emit one PO entry per column per row where qty > 0
  - [ ] 15.2 Use `manualPurchaseOrder` as PO number when provided; otherwise use pivot column header
  - [ ] 15.3 Use `manualDestination` as destination when provided; otherwise attempt to normalise pivot column header as country

- [ ] 16. Add pivot detection and expansion to `ng_excel_generator.py` (TS_Python_Parity)

## Phase 7 — Fox Racing Plant Code Destinations (Requirement 12)

- [ ] 17. Confirm and populate Fox Racing plant code destinations in `PLANT_COUNTRY_MAP`
  - [ ] 17.1 Obtain confirmed destination countries for plant codes `10`, `11`, `40`, `50`, `60` from operator
  - [ ] 17.2 Update `PLANT_COUNTRY_MAP` entries (currently empty strings) with confirmed values
  - [ ] 17.3 Add WARNING emission in `processBuyFile()` when plant code maps to empty string destination

- [ ] 18. Mirror `PLANT_COUNTRY_MAP` Fox Racing entries in `ng_excel_generator.py`

## Phase 8 — Property-Based Tests (Requirements 1–13)

- [ ] 19. Write property-based tests for alias coverage and date parsing in TypeScript
  - [ ] 19.1 Property 1: every key in `getFallbackColumnAliases()` maps to a non-empty field name
  - [ ] 19.2 Property 2: alias matching is case-insensitive and whitespace-tolerant
  - [ ] 19.3 Property 6: Excel serial [1, 2958465] → correct Date via epoch 1899-12-30
  - [ ] 19.4 Property 7: date format round-trip (format then re-parse = original date)
  - [ ] 19.5 Property 8: non-empty unparseable date string → WARNING in `processBuyFile()` output
  - [ ] 19.6 Property 14: EVO pivot expansion — total output size rows = sum of positive pivot cells
  - [ ] 19.7 Property 16: every key in `BRAND_SUPPLIER_MAP` also exists in `BRAND_CUSTOMER_MAP`

- [ ] 20. Write property-based tests in Python (`ng_excel_generator.py` parity)
  - [ ] 20.1 Property 6 (Python): Excel serial conversion via Hypothesis
  - [ ] 20.2 Property 7 (Python): date round-trip via Hypothesis
  - [ ] 20.3 Property 10: every alias in `getFallbackColumnAliases()` (TS) has a matching entry in `HEADER_ALIASES` (Python)
