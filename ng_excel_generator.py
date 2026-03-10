from __future__ import annotations

import argparse
from collections import defaultdict
from datetime import datetime, date
from pathlib import Path
import re
from typing import Any

from openpyxl import Workbook, load_workbook


TRANSPORT_MAP = {
    "Ocean": "Sea",
    "Air": "Air",
    "Sea": "Sea",
    # Arcteryx-specific transport codes
    "S1 - Seafreight": "Sea",
    "S1": "Sea",
    "A1 - Airfreight": "Air",
    "A1": "Air",
    "A2 - Airfreight": "Air",
    "A2": "Air",
    "Seafreight": "Sea",
    "Airfreight": "Air",
    "Sea Freight": "Sea",
    "Air Freight": "Air",
}

CURRENCY = "USD"
SUPPLIER_PROFILE = "DEFAULT_PROFILE"
DEFAULT_CUSTOMER = "COL"

# ─────────────────────────────────────────────────────────────────────────────
# Brand-level supplier fallback map.
# When buy file does not have an explicit supplier/vendor column or when
# vendor_name is provided instead of a code, we resolve it here.
# Keys are lowercase brand identifiers found in the buy file.
# ─────────────────────────────────────────────────────────────────────────────
BRAND_SUPPLIER_MAP: dict[str, str] = {
    "col": "MSO",
    "columbia": "MSO",
    "tnf": "PT. UWU JUMP INDONESIA",
    "the north face": "PT. UWU JUMP INDONESIA",
    "arcteryx": "PT UWU JUMP INDONESIA",
    "arc'teryx": "PT UWU JUMP INDONESIA",
}

# Brand → friendly customer name used in output files.
BRAND_CUSTOMER_MAP: dict[str, str] = {
    "col": "Columbia",
    "columbia": "Columbia",
    "tnf": "The North Face In-Line",
    "the north face": "The North Face In-Line",
    "arcteryx": "Arcteryx",
    "arc'teryx": "Arcteryx",
}

# Fallback fixed positions for the original COL buy file (1-based Excel columns)
DEFAULT_COL_MAP = {
    "po": 7,
    "product": 21,
    "product_alt": 22,
    "size": 0,
    "colour": 24,
    "qty": 28,
    "orig_ex_fac": 34,
    "trans_cond": 31,
    "buy_date": 9,
    "customer": 2,
    "brand": 2,
    "season": 6,
    "template": 8,
}

# ─────────────────────────────────────────────────────────────────────────────
# HEADER_ALIASES  –  extended to cover Arcteryx/Madison88, COL/INFOR, TNF
# ─────────────────────────────────────────────────────────────────────────────
HEADER_ALIASES: dict[str, list[str]] = {
    "po": [
        "po #", "po#", "po", "pono",
        "purchase order", "purchaseorder",
        "extraction po #", "extraction po#",
        # Arcteryx – per-shipment tracking code used as PO key
        "tracking number",
    ],
    "product": [
        "material style", "product", "style number", "style no",
        # Arcteryx – Article is the colourway-level style code
        "article",
        # Generic fallbacks
        "style", "sku", "item",
    ],
    "product_alt": [
        "jde style",
        # Arcteryx model = base style without colour suffix
        "model",
    ],
    "product_name": [
        "model description", "article name", "sku description",
        "material name", "style name", "description",
    ],
    "size": [
        "size", "size name", "sizename", "product size", "productsize",
        "size code", "size #",
    ],
    "customer": [
        "customer", "customer name", "brand",
        # Arcteryx
        "business unit description",
    ],
    "brand": [
        "brand",
        # Arcteryx
        "business unit description",
    ],
    "vendor_name": [
        "vendor name", "vendorname", "supplier name",
    ],
    "vendor_code": [
        "vendor code", "vendorcode", "vendor", "supplier",
        "product supplier",
    ],
    "season": [
        "season", "range", "productrange",
    ],
    "template": [
        "doc type", "template",
    ],
    "colour": [
        "color", "colour",
        # Arcteryx – Article Name holds the colour description
        "article name", "color description",
    ],
    "qty": [
        "ordered qty", "quantity", "qty",
        "open qty (pcs/prs)",
        # Arcteryx
        "requested qty",
    ],
    "orig_ex_fac": [
        "orig ex fac", "delivery date", "deliverydate",
        "negotiated ex fac date", "ex fac",
        # Arcteryx
        "ex-factory",
    ],
    "trans_cond": [
        "trans cond", "transport method", "transportmethod",
        # Arcteryx
        "transport mode",
    ],
    "buy_date": [
        "buy date", "keydate", "po issuance date",
        # Arcteryx
        "file date",
    ],
    "cancel_date": [
        "cancel date", "canceldate", "cancel",
    ],
    "status": [
        "status", "confirmation status",
        # Arcteryx
        "gsc type",
    ],
    "submit_buy": [
        "submit buy", "buy round",
    ],
    "category": [
        "product group description", "product line description",
        "planning category", "dept", "department",
        "capacity type",
    ],
}

ORDERS_HEADERS = [
    "PurchaseOrder", "ProductSupplier", "Status", "Customer",
    "TransportMethod", "TransportLocation", "PaymentTerm", "Template",
    "KeyDate", "ClosedDate", "DefaultDeliveryDate", "Comments", "Currency",
    "KeyUser1", "KeyUser2", "KeyUser3", "KeyUser4", "KeyUser5",
    "KeyUser6", "KeyUser7", "KeyUser8",
    "ArchiveDate", "PurchaseUOM", "SellingUOM",
    "ProductSupplierExt", "FindField_ProductSupplier",
]

LINES_HEADERS = [
    "PurchaseOrder", "LineItem", "ProductRange", "Product", "Customer",
    "DeliveryDate", "TransportMethod", "TransportLocation", "Status",
    "PurchasePrice", "SellingPrice", "Template", "KeyDate", "SupplierProfile",
    "ClosedDate", "Comments", "Currency", "ArchiveDate",
    "ProductExternalRef", "ProductCustomerRef", "PurchaseUOM", "SellingUOM",
    "UDF-buyer_po_number", "UDF-start_date", "UDF-canel_date",
    "UDF-Inspection result", "UDF-Report Type", "UDF-Inspector",
    "UDF-Approval Status", "UDF-Submitted inspection date",
    "FindField_Product",
]

SIZES_HEADERS = [
    "PurchaseOrder", "LineItem", "Range", "Product",
    "SizeName", "ProductSize", "Quantity", "Colour", "Customer",
    "Department", "CustomAttribute1", "CustomAttribute2", "CustomAttribute3",
    "LineRatio", "ColourExt", "CustomerExt", "DepartmentExt",
    "CustomAttribute1Ext", "CustomAttribute2Ext", "CustomAttribute3Ext",
    "ProductExternalRef", "ProductCustomerRef",
    "FindField_Colour", "FindField_Customer", "FindField_Department",
    "FindField_CustomAttribute1", "FindField_CustomAttribute2",
    "FindField_CustomAttribute3", "FindField_Product",
]


# ─────────────────────────────────────────────────────────────────────────────
# Utility helpers
# ─────────────────────────────────────────────────────────────────────────────

def _as_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _format_date(value: Any, fmt: str) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, datetime):
        return value.strftime(fmt)
    if isinstance(value, date):
        return value.strftime(fmt)
    raw = str(value).strip()
    if not raw:
        return ""
    for c in ("%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y", "%d-%b-%Y", "%d-%B-%Y"):
        try:
            return datetime.strptime(raw, c).strftime(fmt)
        except ValueError:
            continue
    return ""


def _parse_date(value: Any) -> datetime | None:
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    raw = str(value).strip()
    if not raw:
        return None
    for c in ("%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y", "%d-%b-%Y", "%d-%B-%Y"):
        try:
            return datetime.strptime(raw, c)
        except ValueError:
            continue
    return None


def _transport_method(value: Any) -> str:
    text = _as_text(value)
    return TRANSPORT_MAP.get(text, text or "Sea")


def _format_product_range(season: str) -> str:
    normalized = (season or "").strip()
    # Handle "FW26" → "FH:2026"  and "SS26" → "SH:2026" as well as "F26"/"S26"
    match = re.match(r"^([FS])(?:W|S)?(\d{2})$", normalized, flags=re.IGNORECASE)
    if match:
        half = "F" if match.group(1).upper() == "F" else "S"
        return f"{half}H:20{match.group(2)}"
    if normalized:
        return normalized
    return "FH:2026"


def _normalize_template(raw_template: str) -> str:
    normalized = (raw_template or "").strip().upper()
    template_map = {
        "OG": "BULK", "ZNB1": "BULK", "ZMF1": "BULK", "ZDS1": "BULK",
    }
    return template_map.get(normalized, (raw_template or "BULK").strip() or "BULK")


def _build_comments(brand: str, season: str, buy_date: Any, template: str, buy_round: str = "") -> str:
    b = (brand or "OUTPUT").strip() or "OUTPUT"
    s = (season or "NOS").strip() or "NOS"
    parsed = _parse_date(buy_date)
    if buy_round:
        return f"[{b}] {s} {buy_round} {template}"
    if parsed:
        mon_short = parsed.strftime("%b")
        day = parsed.strftime("%d")
        mon_upper = mon_short.upper()
        return f"[{b}] {s} {mon_short} Buy {day}-{mon_upper} {template}"
    return f"[{b}] {s}"


def _to_int_quantity(value: Any) -> int:
    if value in (None, ""):
        return 0
    if isinstance(value, bool):
        return 0
    if isinstance(value, (int, float)):
        return int(value)
    try:
        return int(float(str(value).strip()))
    except ValueError:
        return 0


def _append_row(ws, values: list[Any]) -> None:
    ws.append(["" if v is None else v for v in values])


def _normalize_header(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip().lower())


def _compact_text(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", value.lower())


def _header_matches(header: str, alias: str) -> bool:
    if not header or not alias:
        return False
    if header == alias:
        return True
    h_compact = _compact_text(header)
    a_compact = _compact_text(alias)
    if h_compact == a_compact:
        return True
    if a_compact and len(a_compact) > 3 and a_compact in h_compact:
        return True
    return False


def _normalize_po(value: Any) -> str:
    if value is None:
        return ""
    # Strip leading/trailing whitespace only — preserve internal spacing
    # (some PO numbers like "F  164860 OG" have intentional double spaces)
    candidate = str(value).strip()
    return candidate


def _resolve_supplier(vendor_code: str, vendor_name: str, brand: str) -> str:
    """Resolve product supplier from available fields, with brand-level fallback."""
    if vendor_code and len(vendor_code) > 2:
        return vendor_code.strip()
    if vendor_name and len(vendor_name) > 2:
        return vendor_name.strip()
    brand_key = (brand or "").strip().lower()
    return BRAND_SUPPLIER_MAP.get(brand_key, "MISSING_SUPPLIER")


def _resolve_customer(customer_raw: str, brand: str, fallback: str) -> str:
    """Resolve output customer name from available fields."""
    if customer_raw:
        brand_key = customer_raw.strip().lower()
        mapped = BRAND_CUSTOMER_MAP.get(brand_key)
        if mapped:
            return mapped
        # If it looks like a code (short, uppercase), try brand map too
        if len(customer_raw) <= 6 and customer_raw.isupper():
            return BRAND_CUSTOMER_MAP.get(brand_key, customer_raw)
        return customer_raw.strip()
    if brand:
        brand_key = brand.strip().lower()
        return BRAND_CUSTOMER_MAP.get(brand_key, brand.strip())
    return fallback


# ─────────────────────────────────────────────────────────────────────────────
# Layout detection  –  scans up to row 80 for the best header row
# ─────────────────────────────────────────────────────────────────────────────

def _detect_layout(ws) -> tuple[int, dict[str, int], str, int, int, set[str]]:
    best_row = 14
    best_map: dict[str, int] = {}
    best_score = -1
    best_nonempty_po_rows = 0

    for row_idx in range(1, min(81, ws.max_row + 1)):
        headers_by_col: dict[int, str] = {}
        for col in range(1, ws.max_column + 1):
            normalized = _normalize_header(ws.cell(row=row_idx, column=col).value)
            if normalized:
                headers_by_col[col] = normalized

        if not headers_by_col:
            continue

        col_map: dict[str, int] = {}
        for key, aliases in HEADER_ALIASES.items():
            for col, hdr in headers_by_col.items():
                if any(_header_matches(hdr, alias) for alias in aliases):
                    col_map[key] = col
                    break

        # Require at minimum: a PO-like field AND quantity AND a product field
        has_required = (
            ("po" in col_map)
            and ("qty" in col_map)
            and (("product" in col_map) or ("product_alt" in col_map))
        )
        if not has_required:
            continue

        score = len(col_map)

        # Probe how many rows below actually have PO data
        po_col = col_map["po"]
        probe_end = min(ws.max_row, row_idx + 300)
        nonempty_po_rows = sum(
            1 for r in range(row_idx + 1, probe_end + 1)
            if _normalize_po(ws.cell(row=r, column=po_col).value)
        )

        if (score > best_score) or (score == best_score and nonempty_po_rows > best_nonempty_po_rows):
            best_score = score
            best_nonempty_po_rows = nonempty_po_rows
            best_row = row_idx
            best_map = col_map

    if best_score >= 3:
        mode = f"auto-detected headers on row {best_row}"
        # Only merge DEFAULT_COL_MAP keys that were NOT detected from actual headers,
        # AND only when the default column index has a sensible header in this file.
        # This prevents COL-specific fixed positions from polluting other brand files.
        merged_map = dict(best_map)  # start from detected keys only
        return best_row + 1, merged_map, mode, best_score, best_nonempty_po_rows, set(best_map.keys())

    # Legacy fallback: headers row 14, data row 15
    return 15, DEFAULT_COL_MAP.copy(), "fallback fixed-column layout", 0, 0, set()


def _pick_source_sheet(wb, requested_sheet: str | None):
    if requested_sheet:
        if requested_sheet not in wb.sheetnames:
            available = ", ".join(wb.sheetnames)
            raise ValueError(f"Sheet '{requested_sheet}' not found. Available: {available}")
        ws = wb[requested_sheet]
        return (ws,) + _detect_layout(ws)

    best_ws = wb.active
    best_data_start = 15
    best_col_map = DEFAULT_COL_MAP.copy()
    best_layout_mode = "fallback fixed-column layout"
    best_score = -1
    best_nonempty = -1
    best_detected: set[str] = set()

    for ws in wb.worksheets:
        data_start, col_map, layout_mode, score, nonempty, detected = _detect_layout(ws)
        if (score > best_score) or (score == best_score and nonempty > best_nonempty):
            best_ws, best_data_start, best_col_map = ws, data_start, col_map
            best_layout_mode, best_score, best_nonempty, best_detected = layout_mode, score, nonempty, detected

    return best_ws, best_data_start, best_col_map, best_layout_mode, best_score, best_nonempty, best_detected


# ─────────────────────────────────────────────────────────────────────────────
# Main processing function
# ─────────────────────────────────────────────────────────────────────────────

def generate_templates(
    input_path: Path,
    output_dir: Path,
    sheet_name: str | None = None,
    customer_fallback: str = DEFAULT_CUSTOMER,
    strict: bool = False,
) -> None:
    wb = load_workbook(input_path, data_only=True)
    src, data_start_row, col_map, layout_mode, layout_score, probed_po_rows, detected_keys = \
        _pick_source_sheet(wb, sheet_name)

    orders_wb, lines_wb, sizes_wb = Workbook(), Workbook(), Workbook()
    orders_ws = orders_wb.active
    lines_ws  = lines_wb.active
    sizes_ws  = sizes_wb.active
    orders_ws.title, lines_ws.title, sizes_ws.title = "ORDERS", "LINES", "ORDER_SIZES"

    _append_row(orders_ws, ORDERS_HEADERS)
    _append_row(lines_ws,  LINES_HEADERS)
    _append_row(sizes_ws,  SIZES_HEADERS)

    seen_orders:         set[str]              = set()
    line_item_counter:   dict[str, int]        = defaultdict(int)
    po_to_lines:         dict[str, list[int]]  = defaultdict(list)
    po_to_nonzero_lines: dict[str, list[int]]  = defaultdict(list)
    po_to_sizes:         dict[str, list[int]]  = defaultdict(list)

    total_buy_rows = total_lines_rows = total_sizes_rows = 0
    skipped_empty_po = 0
    skipped_empty_po_samples: list[tuple] = []
    validation_warnings: list[str] = []
    validation_errors:   list[str] = []

    def add_warning(msg: str) -> None:
        if len(validation_warnings) < 200:
            validation_warnings.append(msg)

    def add_error(msg: str) -> None:
        if len(validation_errors) < 200:
            validation_errors.append(msg)

    # Column-level checks
    for required_key in ("season", "template", "customer", "size"):
        if required_key not in detected_keys:
            msg = f"Missing required column mapping for '{required_key}' (header not detected)."
            (add_error if strict else add_warning)(msg)

    if "brand" not in detected_keys:
        add_warning("Missing column mapping for 'brand' (comments will use fallback brand value).")

    def _cell(row_idx: int, key: str) -> Any:
        col_idx = col_map.get(key)
        if not col_idx:
            return None
        return src.cell(row=row_idx, column=col_idx).value

    for row_idx in range(data_start_row, src.max_row + 1):
        raw_po_val = _cell(row_idx, "po")
        po = _normalize_po(raw_po_val)
        if not po:
            skipped_empty_po += 1
            if len(skipped_empty_po_samples) < 10:
                skipped_empty_po_samples.append((
                    row_idx,
                    _as_text(src.cell(row=row_idx, column=1).value),
                    _as_text(src.cell(row=row_idx, column=2).value),
                    _as_text(src.cell(row=row_idx, column=3).value),
                    _as_text(raw_po_val),
                ))
            continue

        total_buy_rows += 1

        # ── Core fields ──────────────────────────────────────────────────────
        product = _as_text(_cell(row_idx, "product")) or _as_text(_cell(row_idx, "product_alt"))
        colour  = _as_text(_cell(row_idx, "colour"))
        qty     = _to_int_quantity(_cell(row_idx, "qty"))

        orig_ex_fac  = _cell(row_idx, "orig_ex_fac")
        buy_date     = _cell(row_idx, "buy_date")
        trans_cond   = _cell(row_idx, "trans_cond")
        season_raw   = _as_text(_cell(row_idx, "season"))
        template_raw = _as_text(_cell(row_idx, "template"))
        brand_value  = _as_text(_cell(row_idx, "brand"))
        customer_raw = _as_text(_cell(row_idx, "customer"))
        size_raw     = _as_text(_cell(row_idx, "size"))
        vendor_code  = _as_text(_cell(row_idx, "vendor_code"))
        vendor_name  = _as_text(_cell(row_idx, "vendor_name"))
        buy_round    = _as_text(_cell(row_idx, "submit_buy"))
        status_raw   = _as_text(_cell(row_idx, "status"))

        # ── Derived values ───────────────────────────────────────────────────
        trans_method       = _transport_method(trans_cond)
        key_date_obj       = _parse_date(buy_date)
        key_date_lines     = _format_date(buy_date, "%m/%d/%Y")
        delivery_date      = _format_date(orig_ex_fac, "%m/%d/%Y")
        cancel_date        = _format_date(_cell(row_idx, "cancel_date") or orig_ex_fac, "%m/%d/%Y")

        customer_value     = _resolve_customer(customer_raw, brand_value, customer_fallback)
        supplier_value     = _resolve_supplier(vendor_code, vendor_name, brand_value or customer_raw)
        size_value         = size_raw or "OS"
        product_range      = _format_product_range(season_raw)
        template_value     = _normalize_template(template_raw)
        status_value       = status_raw if status_raw else "Confirmed"
        comments_value     = _build_comments(
            brand_value or customer_raw, season_raw, buy_date, template_value, buy_round
        )

        # ── Row-level validation warnings ────────────────────────────────────
        for field_name, field_val, fallback_desc in [
            ("season",   season_raw,   f"ProductRange '{product_range}'"),
            ("template", template_raw, f"Template '{template_value}'"),
            ("customer", customer_raw, f"Customer '{customer_value}'"),
            ("size",     size_raw,     "Size 'OS'"),
        ]:
            if not field_val:
                msg = f"Row {row_idx} PO {po}: {field_name} is empty; fallback {fallback_desc} used."
                (add_error if strict else add_warning)(msg)

        if not brand_value:
            add_warning(f"Row {row_idx} PO {po}: brand is empty; comments use fallback brand value.")

        # ── ORDERS row (one per unique PO) ───────────────────────────────────
        if po not in seen_orders:
            seen_orders.add(po)
            _append_row(orders_ws, [
                po, supplier_value, "Confirmed", customer_value,
                trans_method, "", "", template_value,
                key_date_obj if key_date_obj else "",
                "", "", comments_value, CURRENCY,
                "", "", "", "", "", "", "", "", "", "", "", "", "",
            ])

        # ── LINES row (one per buy file row) ─────────────────────────────────
        line_item_counter[po] += 1
        line_item = line_item_counter[po]

        _append_row(lines_ws, [
            po, line_item, product_range, product, customer_value,
            delivery_date, trans_method, "", "", "", "",
            template_value, key_date_lines, SUPPLIER_PROFILE,
            "", "", CURRENCY, "", "", "", "", "",
            raw_po_val if raw_po_val is not None else po,
            delivery_date, cancel_date,
            "", "", "", "", "", "",
        ])
        po_to_lines[po].append(line_item)
        total_lines_rows += 1

        # ── ORDER_SIZES row (skip qty=0) ─────────────────────────────────────
        if qty == 0:
            continue

        po_to_nonzero_lines[po].append(line_item)
        _append_row(sizes_ws, [
            po, line_item, product_range, product,
            size_value, size_value, qty, colour,
            "", "", "", "", "", "", "", "", "", "", "", "", "", "",
            "", "", "", "", "", "", "",
        ])
        po_to_sizes[po].append(line_item)
        total_sizes_rows += 1

    # ── Post-processing integrity checks ─────────────────────────────────────
    unique_po_count = len(seen_orders)
    orders_count    = orders_ws.max_row - 1

    if orders_count != unique_po_count:
        raise ValueError(f"ORDERS row count mismatch. Expected {unique_po_count}, got {orders_count}.")

    if total_lines_rows != total_buy_rows:
        raise ValueError(f"LINES row count mismatch. Expected {total_buy_rows}, got {total_lines_rows}.")

    if total_buy_rows == 0:
        raise ValueError(
            "No usable buy rows detected. Check the sheet/header row, "
            "or pass --sheet with the exact worksheet name."
        )

    for po, line_items in po_to_lines.items():
        expected = list(range(1, len(line_items) + 1))
        if line_items != expected:
            raise ValueError(
                f"LineItem sequence invalid for PO {po}. "
                f"Expected {expected}, got {line_items}."
            )
        size_items          = po_to_sizes.get(po, [])
        expected_size_items = po_to_nonzero_lines.get(po, [])
        if size_items != expected_size_items:
            raise ValueError(
                f"LineItem mismatch LINES vs ORDER_SIZES for PO {po}. "
                f"Expected {expected_size_items}, got {size_items}."
            )

    if strict and validation_errors:
        _print_summary(
            src, layout_mode, layout_score, probed_po_rows, data_start_row,
            total_buy_rows, unique_po_count, orders_count,
            total_lines_rows, total_sizes_rows,
            skipped_empty_po, validation_warnings, validation_errors,
            skipped_empty_po_samples,
        )
        raise ValueError("Strict validation failed due to missing/blank required fields.")

    # ── Write output files ────────────────────────────────────────────────────
    output_dir.mkdir(parents=True, exist_ok=True)
    orders_wb.save(output_dir / "ORDERS.xlsx")
    lines_wb.save(output_dir  / "LINES.xlsx")
    sizes_wb.save(output_dir  / "ORDER_SIZES.xlsx")

    _print_summary(
        src, layout_mode, layout_score, probed_po_rows, data_start_row,
        total_buy_rows, unique_po_count, orders_count,
        total_lines_rows, total_sizes_rows,
        skipped_empty_po, validation_warnings, validation_errors,
        skipped_empty_po_samples,
    )
    print(f"\nGenerated files in: {output_dir.resolve()}")
    for fname in ("ORDERS.xlsx", "LINES.xlsx", "ORDER_SIZES.xlsx"):
        print(f"  - {fname}")


def _print_summary(
    src, layout_mode, layout_score, probed_po_rows, data_start_row,
    total_buy_rows, unique_po_count, orders_count,
    total_lines_rows, total_sizes_rows,
    skipped_empty_po, validation_warnings, validation_errors,
    skipped_empty_po_samples,
) -> None:
    lines_equals_sizes = total_lines_rows == total_sizes_rows
    print("\n── Validation Summary ──────────────────────────────────────────")
    print(f"  Source sheet   : {src.title}")
    print(f"  Layout mode    : {layout_mode}")
    print(f"  Layout score   : {layout_score}")
    print(f"  PO probe rows  : {probed_po_rows}")
    print(f"  Data start row : {data_start_row}")
    print(f"  Buy rows       : {total_buy_rows}")
    print(f"  Unique POs     : {unique_po_count}")
    print(f"  ORDERS rows    : {orders_count}")
    print(f"  LINES rows     : {total_lines_rows}")
    print(f"  ORDER_SIZES    : {total_sizes_rows}")
    print(f"  LINES==SIZES   : {'YES' if lines_equals_sizes else 'NO (qty=0 rows excluded in ORDER_SIZES)'}")
    print(f"  Skipped (no PO): {skipped_empty_po}")
    print(f"  Warnings       : {len(validation_warnings)}")
    print(f"  Errors         : {len(validation_errors)}")
    if validation_warnings:
        print("  Warning samples:")
        for msg in validation_warnings[:15]:
            print(f"    - {msg}")
    if validation_errors:
        print("  Error samples:")
        for msg in validation_errors[:15]:
            print(f"    - {msg}")
    if skipped_empty_po_samples:
        print("  Empty-PO samples (row, A, B, C, raw PO cell):")
        for row_idx, col_a, col_b, col_c, sample_val in skipped_empty_po_samples:
            print(f"    - {row_idx}: {col_a} | {col_b} | {col_c} | {sample_val}")
    print("────────────────────────────────────────────────────────────────")


# ─────────────────────────────────────────────────────────────────────────────
# CLI entry point
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Global NG buy-file processor → ORDERS / LINES / ORDER_SIZES"
    )
    parser.add_argument("--input",      dest="input_file",       required=True)
    parser.add_argument("--output-dir", dest="output_dir",       default=".")
    parser.add_argument("--sheet",      dest="sheet_name",       default=None)
    parser.add_argument("--customer",   dest="customer_fallback", default=DEFAULT_CUSTOMER)
    parser.add_argument("--strict",     dest="strict",           action="store_true")
    args = parser.parse_args()

    base           = Path.cwd()
    input_candidate = Path(args.input_file)
    input_path      = input_candidate if input_candidate.is_absolute() else (base / input_candidate)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    output_candidate = Path(args.output_dir)
    output_dir       = output_candidate if output_candidate.is_absolute() else (base / output_candidate)

    generate_templates(
        input_path        = input_path,
        output_dir        = output_dir,
        sheet_name        = args.sheet_name,
        customer_fallback = args.customer_fallback,
        strict            = args.strict,
    )