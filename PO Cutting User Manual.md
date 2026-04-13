# User Manual

**PO Cutting Automation**

April 2026

**ISSUED BY**

Madison88 Business Solutions Asia Inc.  
304 Plaz@ B Building, Northgate Cyberzone,  
Filinvest, Alabang, Muntinlupa City,  
Metro Manila, Philippines 1781

---

## Overview

The PO Cutting Automation system converts customer buy files into standardized NextGen upload templates for cutting operations. It helps users upload one or more source Excel files, apply required manual values, validate data quality, review transformation results, and download final output files ready for NextGen ingestion.

The system includes:

- Buy file upload and multi-file merge support
- Product reference enrichment using Product Sheet / Product SHI export files
- Manual PO and fallback override fields
- Automated format detection and column-mapping checks
- Validation of critical and warning-level issues
- Generation of `ORDERS.xlsx`, `LINES.xlsx`, and `ORDER_SIZES.xlsx`
- Admin tools for maintaining factory, column, and MLO mappings
- Audit logging and run-history tracking in the backend

This manual covers both end-user workflow steps and administrator maintenance procedures.

---

## Table of Contents

Overview  
I. System Purpose and Outputs  
II. Accessing the System  
III. End-User Workflow  
IV. Upload Requirements  
V. Manual Fields and Advanced Overrides  
VI. Validation and Audit Results  
VII. Review and QA Checks  
VIII. Downloading Output Files  
IX. Admin Control Center  
X. Mapping Maintenance Guide  
XI. Security and Access Control  
XII. Troubleshooting and Support  
XIII. Best Practices  
XIV. Document Control

---

## I. System Purpose and Outputs

The main purpose of PO Cutting Automation is to transform raw customer buy files into the three Excel templates required by the NextGen process:

1. `ORDERS.xlsx`
   i. Contains PO header data such as Purchase Order, Product Supplier, Customer, templates, comments, and Key Users.
2. `LINES.xlsx`
   i. Contains line-level style and delivery data such as Line Item, Product, Product Range, and customer references.
3. `ORDER_SIZES.xlsx`
   i. Contains size-level breakdowns including Product, Size Name, and Quantity.

These files may be generated:

- As one merged output set for all uploaded buy files
- As separate per-file output packages when multiple buy files are uploaded

---

## II. Accessing the System

1. Open the PO Cutting Automation homepage.
2. The default landing view is the workflow screen.
3. The top-right navigation includes:
   i. `Admin Login`
   ii. `System Reference`
   iii. `Management Console`
4. The workflow can be used without admin access.
5. Admin access is required only for Control Center functions such as mapping maintenance.

Important notes:

- The system health status is shown on the main screen and refreshes automatically.
- If admin access is needed, use the administrator password provided by IT or the system owner.

---

## III. End-User Workflow

The standard user flow follows five stages:

1. `Acquisition`
   i. Upload buy file(s)
   ii. Upload the product reference file when product enrichment is required
   iii. Enter required manual values
2. `Engine`
   i. The system processes and transforms uploaded data
3. `Audit`
   i. The system lists validation findings and unmapped columns
4. `Review`
   i. The system summarizes merged output counts and QA checkpoints
5. `Export`
   i. Download the final Excel templates

Users should complete each stage in sequence before using the outputs.

---

## IV. Upload Requirements

### 1. Required Inputs

The following are required before starting upload:

- At least one buy file in Excel format
- `Manual PO`

### 2. Optional Inputs

The following may be supplied when available:

- Product Sheet / Product SHI export file
- Template
- Lines Template
- Comments
- Orders KeyDate
- Advanced override values

### 3. File Rules

The upload API currently enforces the following rules:

- Maximum of `5` files per request
- Maximum file size of `30 MB` per file
- Accepted file types: `.xlsx` and `.xls`
- Rate limit: `10` uploads per minute per workflow user

### 4. Upload Steps

1. Select one or more buy files.
2. Select the Product Sheet / Product SHI export file when style-to-product enrichment is needed.
3. Enter the `Manual PO`.
4. Fill in any additional required or fallback fields.
5. Click `START UPLOAD`.

Important note:

- If `Manual PO` is blank, the system will stop the workflow and show a critical validation error.

---

## V. Manual Fields and Advanced Overrides

### 1. Standard Manual Fields

Users may enter or select:

- `Manual PO`
- `Template`
- `Lines Template`
- `Comments`
- `Orders KeyDate`

### 2. Comment Options

The system includes preset comment choices for common buying scenarios. If the user selects `[Other]`, a custom comment can be entered manually.

### 3. Advanced Override Fields

The `Advanced overrides` panel is intended only for files with missing columns or incomplete source values.

Available override fields include:

- `KeyUser1`
- `KeyUser2`
- `KeyUser3`
- `KeyUser4`
- `KeyUser5`
- `Season`
- `Customer`
- `Brand`
- `Destination`

### 4. Auto-Inferred Values

The system can infer some values:

- `Season` may be inferred from the uploaded filename when possible
- `Customer` may be inferred for some known file naming patterns, such as Vuori files

Best practice:

- Keep advanced overrides empty unless there is a known gap in the source file.
- Use override values only to complete missing required data, not to replace valid source data without confirmation.
- For Vans and similar brands, attach the Product SHI export whenever the buy file contains buyer style or material values that must be converted into NextGen `Product` codes.

---

## VI. Validation and Audit Results

After processing, the system moves to the `Audit` stage.

### 1. What the Audit Screen Shows

The audit page displays:

- Validation findings
- Row references
- Severity level
- Diagnostic message
- Detected file format
- Unmapped columns per uploaded file

### 2. Severity Levels

1. `CRITICAL`
   i. A blocking issue that must be corrected before proceeding
2. `WARNING`
   i. A non-blocking issue that should still be reviewed carefully

### 3. Audit Behavior

- If critical errors exist, the `COMMIT TO REVIEW` button remains disabled.
- Users must return to the source file, correct the issue, and upload again.
- If no critical errors exist, the user may proceed to `Review`.

### 4. Common Validation Topics

The system may flag issues involving:

- Missing required values
- Unmapped columns
- Duplicate Purchase Orders across multiple files
- Blank pricing, delivery, or payment-related values
- Status conflicts such as unconfirmed vs confirmed values
- Customer mapping gaps

---

## VII. Review and QA Checks

The `Review` stage confirms that transformation is complete and gives the user a final chance to inspect the result before download.

### 1. What the Review Screen Shows

- Header count
- Merged order count
- Line count
- Size total
- File-level summary by uploaded filename
- Brand labels detected from processed data
- Error and warning totals

### 2. Critical Blocker Messages

The system may surface blocker warnings such as:

- Status conflict detected
- Customer mapping gap detected
- Blank pricing, delivery, or payment fields found
- Validation failure blocking progression

If any blocker appears, the user should verify the source workbook and the mapping setup before using the output files for cutting.

### 3. Review Steps

1. Confirm the PO, line, and size totals are reasonable.
2. Check file-level counts for each uploaded workbook.
3. Review any blocker messages.
4. If needed, return to `Audit`.
5. If acceptable, click `INITIALIZE TEMPLATE GENERATION`.

---

## VIII. Downloading Output Files

After review, the workflow moves to the `Export` stage.

### 1. Available Downloads

Users can download:

- `ORDERS`
- `LINES`
- `ORDER_SIZES`

### 2. Per-File Downloads

If multiple buy files were uploaded, the system may also display a `Per-file Template Export` section. This allows users to download:

- `ORDERS` for a specific source file
- `LINES` for a specific source file
- `SIZES` for a specific source file

### 3. Download Steps

1. Click the desired package button.
2. Wait for the browser download to begin.
3. Verify the file name and open the workbook.
4. Perform a final spot check before NextGen upload.

### 4. Vans Example: Correct Input Set and Output Pattern

For the Vans February buy example, the correct workflow should include these source files:

- Customer buy file: `Feb Buy Top up M88.xlsx`
- Product reference file: `Product Shi's Export (34).xlsx`

Expected behavior:

- The buy file provides destination, requested dates, buyer PO numbers, quantities, and buyer style context.
- The Product SHI export provides the NextGen `Product` code, style reference, size, cost, sell price, customer, and factory reference data.
- The final output should use the enriched `Product` values from the Product SHI reference instead of leaving the line product as the raw buyer material code from the buy sheet.

Expected output pattern for this Vans example:

1. `ORDERS.xlsx`
   i. One order header per destination / delivery grouping
   ii. Example order names:
      `PO002929-1004-Brampton`
      `PO002929-1023-South Ontario`
      `PO002929-D00028-Sun and Sand Sports`
   iii. Expected comments value:
      `Vans F26 Bulk Feb Buy`
   iv. Expected template:
      `Major Brand Bulk`
2. `LINES.xlsx`
   i. `Product` should resolve to SHI product codes such as `M88129633`, not raw buy-file material values such as `VN000QB4GRK`
   ii. `UDF-buyer_po_number` should carry the source PO number from the buy file
   iii. `Template` should match the approved Vans line template, such as `FOB Bulk EDI PO (New)`
3. `ORDER_SIZES.xlsx`
   i. Quantities must match the buy-file quantity by line
   ii. `Product` should match the same enriched SHI product code used in `LINES`
   iii. `Colour` should reflect the customer-facing color description used for the matched SHI product

QA note for this example:

- If the generated output shows raw style or material values like `VN000QB4GRK` in the `Product` column instead of SHI product codes like `M88129633`, the enrichment step did not complete correctly.
- If the order name and transport location disagree, such as an order name ending in `USA` while `TransportLocation` is `Canada`, the destination grouping logic must be reviewed before upload to NextGen.

---

## IX. Admin Control Center

The `Management Console` provides administrative tools for maintaining system logic.

### 1. Admin Access

1. Click `Admin Login`.
2. Enter the administrator password.
3. Open `Management Console`.

### 2. Main Admin Areas

The Control Center currently includes:

1. `Dashboard`
   i. Displays high-level operational cards and recent activity visuals
2. `Mappings`
   i. Maintains transformation logic tables
3. `Security`
   i. Displays active security controls and access-related information

Note:

- The admin interface is focused mainly on mapping maintenance for live operations.

---

## X. Mapping Maintenance Guide

The `Mappings` area contains three configuration sections.

### A. Factory Mapping

Purpose:

- Maps `Brand + Category` to `Product Supplier`

Typical use:

- When a brand/category combination should resolve to a different factory

Steps:

1. Open `Mappings`.
2. Select `Factory`.
3. Search existing rows if needed.
4. Click `NEW RECORD`.
5. Enter:
   i. Brand
   ii. Category
   iii. Product Supplier
6. Click `Save`.

### B. Column Aliases

Purpose:

- Maps incoming customer buy-file column headers to internal processing fields

Typical use:

- When a customer changes column names in the buy file
- When unmapped columns appear during validation

Steps:

1. Open `Mappings`.
2. Select `Column Aliases`.
3. Search or filter by customer.
4. Click `NEW ALIAS`.
5. Enter:
   i. Customer
   ii. Buy File Column
   iii. Internal Field
   iv. Optional Notes
6. Click `Save`.

Best practice:

- Add aliases only after confirming the source header meaning.

### C. MLO Config

Purpose:

- Maintains brand-level Key Users, templates, and valid statuses

Typical use:

- When a brand requires fixed KeyUser assignments
- When default order or line templates must be changed
- When valid processing statuses need to be updated

Steps:

1. Open `Mappings`.
2. Select `MLO Config`.
3. Click `NEW BRAND CONFIG` or edit an existing row.
4. Enter or update:
   i. Brand
   ii. KeyUser1
   iii. KeyUser2
   iv. KeyUser4
   v. KeyUser5
   vi. Orders Template
   vii. Lines Template
   viii. Valid Statuses
5. Click `Save`.

---

## XI. Security and Access Control

The system includes several access and governance controls.

### 1. Workflow Access

- Standard workflow usage is available without admin session access.

### 2. Protected Admin Actions

Admin-only endpoints require authorization for actions such as:

- Editing mapping tables
- Viewing audit logs
- Accessing protected admin data

### 3. Session Model

- Admin functions depend on an active admin session cookie.
- Unauthorized attempts are logged by the system.

### 4. Audit and Run Tracking

The backend records workflow events such as:

- Buy file uploaded
- Workflow started
- Data extraction completed
- Unauthorized access attempts

Run records also store processing status, error count, warning count, and row totals.

---

## XII. Troubleshooting and Support

### 1. “Manual PO is required.”

Cause:

- The required `Manual PO` field was not entered

Resolution:

- Return to `Acquisition` and provide the PO value before uploading

### 2. “No file uploaded.”

Cause:

- No buy file was attached

Resolution:

- Attach at least one valid Excel buy file

### 3. Unsupported MIME Type or File Format

Cause:

- The uploaded file is not recognized as Excel

Resolution:

- Re-save the file as `.xlsx` and upload again

### 4. Too Many Files

Cause:

- More than five files were submitted in one request

Resolution:

- Split the work into smaller upload batches

### 5. Critical Validation Errors

Cause:

- One or more blocking issues were detected during audit

Resolution:

1. Review the `Diagnostic Message`
2. Correct the source workbook or admin mapping
3. Re-upload the corrected file

### 6. Unmapped Columns

Cause:

- A source column header does not match any configured alias

Resolution:

- Ask an administrator to add a column alias if the field is valid and required

### 7. Download Failure

Cause:

- Output payload is missing or the workflow did not complete successfully

Resolution:

- Re-run the upload after checking the validation stage

### 8. Admin Login Failure

Cause:

- Invalid administrator password

Resolution:

- Confirm credentials with the authorized system owner or IT support

---

## XIII. Best Practices

1. Use the cleanest available buy file version before uploading.
2. Always enter the correct `Manual PO`.
3. Attach the Product Sheet / Product SHI file when richer item details and product-code enrichment are needed.
4. Review all critical and warning messages before export.
5. Do not rely on advanced overrides unless source data is incomplete.
6. Update column aliases immediately when customer headers change.
7. Verify customer mapping gaps before using the generated files for live cutting.
8. Perform a final spot check on `ORDERS`, `LINES`, and `ORDER_SIZES` before NextGen upload.

---

## XIV. Document Control

Document Title: `PO Cutting Automation User Manual`  
System Name: `PO Cutting Automation`  
Version: `1.0`  
Issue Date: `April 2026`  
Prepared For: `Madison88 Operations / Cutting / IT Admin Users`

Recommended review triggers:

- New brand onboarding
- New customer file layout
- Changes to template logic
- Changes to admin access or mapping governance
