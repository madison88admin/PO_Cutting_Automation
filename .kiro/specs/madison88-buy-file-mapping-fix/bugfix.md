# Bugfix Requirements Document

## Introduction

Processing the Madison 88 buy file (`MAR - Buy File Indonesia MADISON 88, LTD.-2026-03-17.xlsx`) fails
across all 428 rows with three consistent errors. The root cause is a column header mismatch: the
Madison 88 file uses `Style#` for the product/style field and `Material` for the colour field, but the
engine's JDE Style lookup requires a column that maps to the internal `jdeStyle` field (header `JDE Style`),
which does not exist in this file. Without a `jdeStyle` value the PLM product-sheet lookup cannot be
keyed, causing style and colour lookups to fail for every row. Additionally, the size column header in
the Madison 88 file does not match any direct alias, so the engine falls back to an inferred mapping
(triggering a Row 1 WARNING) and then fails to find size data for every row.

The three errors reported are:
1. "JDE Style missing; PLM fields left blank" — `jdeStyle` column absent from the file
2. "JDE color [colorCode] not found in PLM sheet; PLM fields left blank" — colour key cannot be matched without a valid JDE style
3. "PO PO002951 missing: Size" (CRITICAL) — size column not reliably detected

---

## Bug Analysis

### Current Behavior (Defect)

1.1 WHEN the Madison 88 buy file is uploaded and the engine scans the header row, THEN the system does not find a column that maps to the internal `jdeStyle` field, because the file contains no `JDE Style` header and no alias for `Style#` resolves to `jdeStyle`.

1.2 WHEN `jdeStyle` is empty for a row and a product sheet map is present, THEN the system emits "JDE Style missing; PLM fields left blank" for every data row, leaving style number and colour blank.

1.3 WHEN the PLM lookup key cannot be formed (because `jdeStyle` is empty), THEN the system also emits "JDE [style] color [colorCode] not found in PLM sheet; PLM fields left blank" for every data row, even though the colour value from the `Material` column was successfully read.

1.4 WHEN the Madison 88 buy file is uploaded and the engine scans the header row for a size column, THEN the system does not find a direct alias match and falls back to an inferred size column, emitting a Row 1 WARNING "Inferred mapping: sizeName from size-like column".

1.5 WHEN the inferred size column does not contain usable size data for every row, THEN the system emits a CRITICAL error "PO [po] missing: Size" for every data row, causing all rows to be skipped.

### Expected Behavior (Correct)

2.1 WHEN the Madison 88 buy file is uploaded, THEN the system SHALL recognise `Style#` as an alias for the `product` field AND SHALL use the `product` field value as the style lookup key when no dedicated `jdeStyle` column is present, so that the PLM lookup can proceed using the style value from `Style#`.

2.2 WHEN `jdeStyle` is absent from the buy file but `product` is present, THEN the system SHALL attempt the PLM product-sheet lookup using the `product` field value as the style key, rather than immediately emitting "JDE Style missing" and skipping the lookup.

2.3 WHEN the PLM lookup succeeds using the `product`-field fallback, THEN the system SHALL NOT emit "JDE Style missing; PLM fields left blank" for rows where the product value is non-empty.

2.4 WHEN the Madison 88 buy file is uploaded, THEN the system SHALL recognise `Material` as an alias for the `colour` field (this alias already exists in `getFallbackColumnAliases()`) and SHALL correctly read the colour value for every row.

2.5 WHEN the Madison 88 buy file is uploaded and the size column header is present in the file, THEN the system SHALL resolve it to `sizeName` via a direct alias match (not an inferred match), so that no Row 1 "Inferred mapping" WARNING is emitted and size data is reliably read for every row.

2.6 WHEN size data is successfully read for every row, THEN the system SHALL NOT emit "PO [po] missing: Size" CRITICAL errors, and all 428 rows SHALL be processed without being skipped due to missing size.

### Unchanged Behavior (Regression Prevention)

3.1 WHEN a buy file that contains an explicit `JDE Style` column is uploaded, THEN the system SHALL CONTINUE TO use the `jdeStyle` column value as the primary PLM lookup key, with no change to existing behaviour.

3.2 WHEN a buy file that contains an explicit `JDE Style` column is uploaded and `jdeStyle` is empty for a row, THEN the system SHALL CONTINUE TO emit "JDE Style missing; PLM fields left blank" for that row.

3.3 WHEN a buy file for TNF, Columbia, or Arcteryx is uploaded with their existing column headers, THEN the system SHALL CONTINUE TO map all columns correctly and produce the same output as before this fix.

3.4 WHEN a buy file contains a column header that is already a direct alias for `sizeName` (e.g. `Size`, `Size Name`, `Size Code`), THEN the system SHALL CONTINUE TO map it directly without falling back to inference.

3.5 WHEN a buy file contains no size column at all, THEN the system SHALL CONTINUE TO emit the "No size column detected. Using default 'One Size' for all rows." WARNING and apply the `One Size` default bucket.

3.6 WHEN the PLM product-sheet lookup fails to find a match (colour key not found), THEN the system SHALL CONTINUE TO emit "JDE [style] color [colorCode] not found in PLM sheet; PLM fields left blank" for that row.
