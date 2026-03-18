from __future__ import annotations

import argparse
from collections import defaultdict, Counter
from datetime import datetime, date
from pathlib import Path
import json
import os
import re
from typing import Any

from openpyxl import Workbook, load_workbook
from urllib.error import URLError, HTTPError
from urllib.parse import quote
from urllib.request import Request, urlopen


TRANSPORT_MAP = {
    "ocean": "Sea",
    "air": "Air",
    "sea": "Sea",
    # Arcteryx-specific transport codes
    "s1 - seafreight": "Sea",
    "s1": "Sea",
    "a1 - airfreight": "Air",
    "a1": "Air",
    "a2 - airfreight": "Air",
    "a2": "Air",
    "seafreight": "Sea",
    "airfreight": "Air",
    "sea freight": "Sea",
    "air freight": "Air",
    # Courier variants
    "courier": "Courier",
    "dhl": "Courier",
    "fedex": "Courier",
    "ups": "Courier",
}

# Valid mapped transport values after TRANSPORT_MAP resolution
VALID_TRANSPORT_VALUES = {"Sea", "Air", "Courier"}

# Country code → full name (used for TransportLocation)
COUNTRY_NAME_MAP = {
    "US": "United States",
    "KR": "Korea",
    "FR": "France",
    "GR": "Greece",
    "DE": "Germany",
    "IT": "Italy",
    "ES": "Spain",
    "CN": "China",
    "JP": "Japan",
    "VN": "Vietnam",
    "TH": "Thailand",
    "PH": "Philippines",
    "ID": "Indonesia",
    "TW": "Taiwan",
    "HK": "Hong Kong",
    "UK": "United Kingdom",
    "GB": "United Kingdom",
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

# ─────────────────────────────────────────────────────────────────────────────
# Customer subtype overrides.
# Keys are lowercase brand + subtype signal found in source data.
# Used by _resolve_customer_subtype() to map raw customer values that contain
# "RTO", "SMU", or other signals to the correct spelled-out customer name.
# ─────────────────────────────────────────────────────────────────────────────
TNF_CUSTOMER_SUBTYPE_MAP: dict[str, str] = {
    "the north face in-line": "The North Face In-Line",
    "the north face inline":  "The North Face In-Line",
    "the north face rto":     "The North Face RTO",
    "the north face smu":     "The North Face SMU",
    "tnf in-line":            "The North Face In-Line",
    "tnf inline":             "The North Face In-Line",
    "tnf rto":                "The North Face RTO",
    "tnf smu":                "The North Face SMU",
}

# ─────────────────────────────────────────────────────────────────────────────
# KeyUser map — per brand, the MLO team members that go into ORDERS.
# Keys are lowercase brand identifiers.
# Values are dicts keyed by KeyUser column name.
# Only KeyUser1 (Planning), KeyUser2 (Purchasing), KeyUser4 (Production),
# KeyUser5 (Logistics/Shipping) are populated per BRD; 3/6/7/8 left blank.
# ─────────────────────────────────────────────────────────────────────────────
BRAND_KEYUSER_MAP: dict[str, dict[str, str]] = {
    "tnf": {
        "KeyUser1": "Ron",
        "KeyUser2": "Maricar",
        "KeyUser3": "",
        "KeyUser4": "Ron",
        "KeyUser5": "Elaine Sanchez",
        "KeyUser6": "",
        "KeyUser7": "",
        "KeyUser8": "",
    },
    "the north face": {
        "KeyUser1": "Ron",
        "KeyUser2": "Maricar",
        "KeyUser3": "",
        "KeyUser4": "Ron",
        "KeyUser5": "Elaine Sanchez",
        "KeyUser6": "",
        "KeyUser7": "",
        "KeyUser8": "",
    },
    "col": {
        "KeyUser1": "",
        "KeyUser2": "",
        "KeyUser3": "",
        "KeyUser4": "",
        "KeyUser5": "",
        "KeyUser6": "",
        "KeyUser7": "",
        "KeyUser8": "",
    },
    "columbia": {
        "KeyUser1": "",
        "KeyUser2": "",
        "KeyUser3": "",
        "KeyUser4": "",
        "KeyUser5": "",
        "KeyUser6": "",
        "KeyUser7": "",
        "KeyUser8": "",
    },
    "arcteryx": {
        "KeyUser1": "",
        "KeyUser2": "",
        "KeyUser3": "",
        "KeyUser4": "",
        "KeyUser5": "",
        "KeyUser6": "",
        "KeyUser7": "",
        "KeyUser8": "",
    },
    "arc'teryx": {
        "KeyUser1": "",
        "KeyUser2": "",
        "KeyUser3": "",
        "KeyUser4": "",
        "KeyUser5": "",
        "KeyUser6": "",
        "KeyUser7": "",
        "KeyUser8": "",
    },
}

# Default KeyUser block (all blank) used when brand not found in map.
_DEFAULT_KEYUSERS: dict[str, str] = {
    k: "" for k in [
        "KeyUser1", "KeyUser2", "KeyUser3", "KeyUser4",
        "KeyUser5", "KeyUser6", "KeyUser7", "KeyUser8",
    ]
}

# ─────────────────────────────────────────────────────────────────────────────
# Template map — per brand, the NG template name to use.
# Overrides _normalize_template() default of "BULK" for brands with
# brand-specific template names (e.g. TNF uses "Major Brand Bulk").
# For LINES the template field uses a different, more granular value.
# ─────────────────────────────────────────────────────────────────────────────
BRAND_ORDERS_TEMPLATE_MAP: dict[str, str] = {
    "tnf":            "Major Brand Bulk",
    "the north face": "Major Brand Bulk",
    "col":            "BULK",
    "columbia":       "BULK",
    "arcteryx":       "BULK",
    "arc'teryx":      "BULK",
}

BRAND_LINES_TEMPLATE_MAP: dict[str, str] = {
    "tnf":            "FOB Bulk EDI PO (New)",
    "the north face": "FOB Bulk EDI PO (New)",
    "col":            "BULK",
    "columbia":       "BULK",
    "arcteryx":       "BULK",
    "arc'teryx":      "BULK",
}

# ─────────────────────────────────────────────────────────────────────────────
# Remote brand config (TS API) — optional source of truth.
# Python CLI will call this when available, otherwise fall back to maps above.
# ─────────────────────────────────────────────────────────────────────────────
_BRAND_CONFIG_CACHE: dict[str, dict[str, Any] | None] = {}


def _fetch_brand_config(brand: str) -> dict[str, Any] | None:
    base_url = os.getenv("BRAND_CONFIG_URL", "http://localhost:3000/api/brand-config")
    if not base_url or not brand:
        return None
    url = f"{base_url}?brand={quote(brand)}"
    try:
        req = Request(url, headers={"Accept": "application/json"})
        with urlopen(req, timeout=2) as resp:
            if resp.status != 200:
                return None
            payload = json.loads(resp.read().decode("utf-8"))
            return payload if isinstance(payload, dict) else None
    except (HTTPError, URLError, ValueError, OSError):
        return None


def _get_brand_config(brand: str) -> dict[str, Any] | None:
    key = (brand or "").strip().lower()
    if not key:
        return None
    if key in _BRAND_CONFIG_CACHE:
        return _BRAND_CONFIG_CACHE[key]
    config = _fetch_brand_config(brand)
    _BRAND_CONFIG_CACHE[key] = config
    return config

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
        "transport_location": [
            "transportlocation", "transport location",
            "destination", "dest country", "ult. destination",
        ],
    "plant": [
        "plant", "plant code",
    ],
    "po": [
        "po #", "po#", "po", "pono",
        "purchase order", "purchaseorder",
        "extraction po #", "extraction po#",
        # Arcteryx – per-shipment tracking code used as PO key
        "tracking number",
    ],
    "product": [
        "material style", "product", "style number", "style no",
        "product name",
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
    "product_external_ref": [
        "name",
        "product external ref",
    ],
    "product_customer_ref": [
        "buyer style number",
        "buyer style no",
        "buyer style #",
        "buyer style",
        "product customer ref",
    ],
    "product_name": [
        "model description", "article name", "sku description",
        "material name", "style name", "description",
        "buyer style name",
    ],
    "size": [
        "size", "size name", "sizename", "product size", "productsize",
        "size code", "size #", "size#", "size_name", "size-name",
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
        "factory",
    ],
    "vendor_code": [
        "vendor code", "vendorcode", "vendor", "supplier",
        "product supplier", "productsupplier",
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

# Expected column counts for output validation
EXPECTED_COL_COUNTS = {
    "ORDERS":      len(ORDERS_HEADERS),   # 26
    "LINES":       len(LINES_HEADERS),    # 31
    "ORDER_SIZES": len(SIZES_HEADERS),    # 29
}


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
    key = text.lower()
    return TRANSPORT_MAP.get(key, text or "Sea")


def _format_transport_location(value: Any) -> str:
    raw = _strip_brackets(_as_text(value))
    if not raw:
        return ""
    key = raw.strip().upper()
    return COUNTRY_NAME_MAP.get(key, raw.strip())


def _format_product_range(season: str) -> str:
    normalized = _strip_brackets((season or "").strip())
    # Handle "FW26" → "FH:2026"  and "SS26" → "SH:2026" as well as "F26"/"S26"
    match = re.match(r"^([FS])(?:W|S)?(\d{2})$", normalized, flags=re.IGNORECASE)
    if match:
        half = "F" if match.group(1).upper() == "F" else "S"
        return f"{half}H:20{match.group(2)}"
    if normalized:
        return normalized
    return "FH:2026"


def _normalize_template(raw_template: str) -> str:
    """Normalise raw doc-type codes from buy files to internal template tokens."""
    normalized = (raw_template or "").strip().upper()
    template_map = {
        "OG": "BULK", "ZNB1": "BULK", "ZMF1": "BULK", "ZDS1": "BULK",
        "SMS": "SMS",
    }
    return template_map.get(normalized, (raw_template or "BULK").strip() or "BULK")


def _resolve_orders_template(brand: str, raw_template: str, brand_config: dict[str, Any] | None = None) -> str:
    """Return the ORDERS Template value for this brand.
    Brand-config value takes priority; then brand-specific map; then _normalize_template()."""
    if brand_config:
        cfg = _as_text(brand_config.get("orders_template"))
        if cfg:
            return cfg
    brand_key = (brand or "").strip().lower()
    if brand_key in BRAND_ORDERS_TEMPLATE_MAP:
        return BRAND_ORDERS_TEMPLATE_MAP[brand_key]
    return _normalize_template(raw_template)


def _resolve_lines_template(brand: str, raw_template: str, brand_config: dict[str, Any] | None = None) -> str:
    """Return the LINES Template value for this brand."""
    if brand_config:
        cfg = _as_text(brand_config.get("lines_template"))
        if cfg:
            return cfg
    brand_key = (brand or "").strip().lower()
    if brand_key in BRAND_LINES_TEMPLATE_MAP:
        return BRAND_LINES_TEMPLATE_MAP[brand_key]
    return _normalize_template(raw_template)


def _build_comments(brand: str, season: str, buy_date: Any, template: str, buy_round: str = "") -> str:
    b = (brand or "OUTPUT").strip() or "OUTPUT"
    s = (season or "NOS").strip() or "NOS"
    parsed = _parse_date(buy_date)
    if buy_round:
        return f"{b} {s} {buy_round} {template}"
    if parsed:
        mon_short = parsed.strftime("%b")
        day = parsed.strftime("%d")
        mon_upper = mon_short.upper()
        suffix = f" {template}" if template else ""
        return f"{b} {s} {mon_short} Buy {day}-{mon_upper}{suffix}"
    return f"{b} {s}"


# Fix #4: parse_int for UDF-buyer_po_number, keep 0 as 0
def parse_int(value):
    try:
        return int(value)
    except Exception:
        return ""

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


def _strip_brackets(value: str) -> str:
    if not value:
        return value
    # Remove bracket characters but keep the inner text
    cleaned = re.sub(r"\[([^\]]+)\]", r"\1", value)
    cleaned = cleaned.replace("[", "").replace("]", "")
    return re.sub(r"\s+", " ", cleaned).strip()


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
    customer_raw = _strip_brackets(customer_raw)
    brand = _strip_brackets(brand)
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


def _resolve_customer_subtype(customer_raw: str, brand: str, fallback: str) -> str:
    """Resolve full customer name including In-Line / SMU / RTO subtype.

    Priority:
    1. Direct match in TNF_CUSTOMER_SUBTYPE_MAP (handles raw values already
       containing the subtype signal, e.g. 'The North Face RTO').
    2. BRAND_CUSTOMER_MAP default (In-Line) for the given brand.
    3. customer_raw as-is if non-empty.
    4. fallback.
    """
    customer_raw = _strip_brackets(customer_raw)
    brand = _strip_brackets(brand)
    if customer_raw:
        key = customer_raw.strip().lower()
        # Try explicit subtype map first
        if key in TNF_CUSTOMER_SUBTYPE_MAP:
            return TNF_CUSTOMER_SUBTYPE_MAP[key]
        # Try brand customer map
        brand_key = key
        if brand_key in BRAND_CUSTOMER_MAP:
            return BRAND_CUSTOMER_MAP[brand_key]
        # Short uppercase code → try brand map
        if len(customer_raw) <= 6 and customer_raw.isupper():
            brand_key2 = customer_raw.strip().lower()
            return BRAND_CUSTOMER_MAP.get(brand_key2, customer_raw)
        return customer_raw.strip()
    if brand:
        brand_key = brand.strip().lower()
        return BRAND_CUSTOMER_MAP.get(brand_key, brand.strip())
    return fallback


def _resolve_keyusers(brand: str) -> dict[str, str]:
    """Return KeyUser1-8 dict for the given brand."""
    brand_key = (brand or "").strip().lower()
    return dict(BRAND_KEYUSER_MAP.get(brand_key, _DEFAULT_KEYUSERS))


# ─────────────────────────────────────────────────────────────────────────────
# Output slice validation  (NEW)
# Validates the three generated workbooks against the reference format rules.
# Returns a list of (level, message) tuples — level is "ERROR" or "WARNING".
# ─────────────────────────────────────────────────────────────────────────────

def _validate_output_slices(
    orders_ws,
    lines_ws,
    sizes_ws,
    validate_sizes: bool = True,
) -> list[tuple[str, str]]:
    """
    Post-generation validation of the three output slices.

    Checks enforced
    ---------------
    1.  Row counts  : LINES rows == SIZES rows (data rows, excluding header)
    2.  Column counts: each sheet has the expected number of columns
    3.  PO universe : ORDERS, LINES, SIZES all share the exact same PO set
    4.  ORDERS      : one row per unique PO (no duplicates)
    5.  Qty total   : SIZES Quantity sum > 0
    6.  Transport   : every TransportMethod value in LINES and ORDERS maps to
                      Sea, Air, or Courier (catches unmapped raw values)
    7.  Customer    : no blank Customer values in LINES or ORDERS
    8.  Status      : ORDERS Status column contains only "Confirmed"
    9.  Duplicate PO+LineItem: LINES and SIZES must have no duplicate keys
    10. LINES/SIZES key alignment: same exact (PO, LineItem) pairs in both sheets
    """
    issues: list[tuple[str, str]] = []

    def col_values(ws, col_idx: int) -> list[str]:
        return [
            _as_text(ws.cell(row=r, column=col_idx).value)
            for r in range(2, ws.max_row + 1)
        ]

    def header_index(ws, name: str) -> int | None:
        for c in range(1, ws.max_column + 1):
            if _as_text(ws.cell(row=1, column=c).value).lower() == name.lower():
                return c
        return None

    # 1. Row counts (optional for ORDER_SIZES)
    lines_data_rows = lines_ws.max_row - 1
    sizes_data_rows = sizes_ws.max_row - 1

    if validate_sizes and lines_data_rows != sizes_data_rows:
        issues.append((
            "ERROR",
            f"Row count mismatch: LINES has {lines_data_rows} rows, "
            f"ORDER_SIZES has {sizes_data_rows} rows."
        ))

    # 2. Column counts
    sheet_checks = [(orders_ws, "ORDERS"), (lines_ws, "LINES")]
    if validate_sizes:
        sheet_checks.append((sizes_ws, "ORDER_SIZES"))
    for ws, sheet_key in sheet_checks:
        expected = EXPECTED_COL_COUNTS[sheet_key]
        actual = ws.max_column
        if actual != expected:
            issues.append((
                "ERROR",
                f"{sheet_key} column count is {actual}, expected {expected}."
            ))

    # 3 & 4. PO universe + ORDERS uniqueness
    orders_po_col = header_index(orders_ws, "PurchaseOrder") or 1
    orders_customer_col = header_index(orders_ws, "Customer")
    lines_po_col  = header_index(lines_ws,  "PurchaseOrder") or 1
    sizes_po_col  = header_index(sizes_ws,  "PurchaseOrder") or 1

    orders_pos = col_values(orders_ws, orders_po_col)
    lines_pos  = col_values(lines_ws,  lines_po_col)
    sizes_pos  = col_values(sizes_ws,  sizes_po_col)

    orders_po_set = set(orders_pos)
    lines_po_set  = set(lines_pos)
    sizes_po_set  = set(sizes_pos)

    if orders_customer_col:
        orders_customers = col_values(orders_ws, orders_customer_col)
        orders_keys = list(zip(orders_pos, orders_customers))
        dupes = [k for k, c in Counter(orders_keys).items() if c > 1]
        if dupes:
            issues.append(("ERROR", f"ORDERS has duplicate PO+Customer keys: {dupes[:5]}"))
    else:
        dupes = [p for p, c in Counter(orders_pos).items() if c > 1]
        if dupes:
            issues.append(("ERROR", f"ORDERS has duplicate POs: {dupes[:5]}"))

    if orders_po_set != lines_po_set:
        only_orders = orders_po_set - lines_po_set
        only_lines  = lines_po_set - orders_po_set
        if only_orders:
            issues.append(("ERROR", f"POs in ORDERS but not LINES: {list(only_orders)[:5]}"))
        if only_lines:
            issues.append(("ERROR", f"POs in LINES but not ORDERS: {list(only_lines)[:5]}"))

    if validate_sizes and orders_po_set != sizes_po_set:
        only_orders = orders_po_set - sizes_po_set
        only_sizes  = sizes_po_set - orders_po_set
        if only_orders:
            issues.append(("ERROR", f"POs in ORDERS but not ORDER_SIZES: {list(only_orders)[:5]}"))
        if only_sizes:
            issues.append(("ERROR", f"POs in ORDER_SIZES but not ORDERS: {list(only_sizes)[:5]}"))

    # 5. Qty total (optional)
    if validate_sizes:
        qty_col = header_index(sizes_ws, "Quantity") or 7
        total_qty = sum(
            _to_int_quantity(sizes_ws.cell(row=r, column=qty_col).value)
            for r in range(2, sizes_ws.max_row + 1)
        )
        if total_qty == 0:
            issues.append(("ERROR", "ORDER_SIZES total Quantity is 0 — no units written."))

    # 6. Transport mapping
    for ws, label in [(lines_ws, "LINES"), (orders_ws, "ORDERS")]:
        tm_col = header_index(ws, "TransportMethod")
        if tm_col is None:
            issues.append(("WARNING", f"{label}: TransportMethod column not found."))
            continue
        bad_transport = set()
        for r in range(2, ws.max_row + 1):
            val = _as_text(ws.cell(row=r, column=tm_col).value)
            if val and val not in VALID_TRANSPORT_VALUES:
                bad_transport.add(val)
        if bad_transport:
            issues.append((
                "ERROR",
                f"{label} TransportMethod has unmapped values (must be Sea, Air, or Courier): "
                f"{sorted(bad_transport)}"
            ))

    # 7. Customer not blank
    for ws, label in [(lines_ws, "LINES"), (orders_ws, "ORDERS")]:
        cust_col = header_index(ws, "Customer")
        if cust_col is None:
            issues.append(("WARNING", f"{label}: Customer column not found."))
            continue
        blank_rows = [
            r for r in range(2, ws.max_row + 1)
            if not _as_text(ws.cell(row=r, column=cust_col).value)
        ]
        if blank_rows:
            issues.append((
                "ERROR",
                f"{label} has {len(blank_rows)} blank Customer value(s) "
                f"(first 5 rows: {blank_rows[:5]})."
            ))

    # 8. Status = Confirmed
    status_col = header_index(orders_ws, "Status")
    if status_col:
        bad_status = set()
        for r in range(2, orders_ws.max_row + 1):
            val = _as_text(orders_ws.cell(row=r, column=status_col).value)
            if val and val != "Confirmed":
                bad_status.add(val)
        if bad_status:
            issues.append((
                "WARNING",
                f"ORDERS Status has unexpected values (expected 'Confirmed'): "
                f"{sorted(bad_status)} — OK if source INFOR view is 'Unconfirmed Only'."
            ))

    # 9. Duplicate PO+LineItem
    li_col_lines = header_index(lines_ws, "LineItem") or 2
    li_col_sizes = header_index(sizes_ws, "LineItem") or 2

    lines_keys = [
        (col_values(lines_ws, lines_po_col)[i], col_values(lines_ws, li_col_lines)[i])
        for i in range(lines_data_rows)
    ]
    sizes_keys = [
        (col_values(sizes_ws, sizes_po_col)[i], col_values(sizes_ws, li_col_sizes)[i])
        for i in range(sizes_data_rows)
    ]

    lines_key_set = set(lines_keys)
    sizes_key_set = set(sizes_keys)

    if len(lines_keys) != len(lines_key_set):
        issues.append(("ERROR", "LINES has duplicate (PurchaseOrder, LineItem) keys."))
    if validate_sizes and len(sizes_keys) != len(sizes_key_set):
        issues.append(("ERROR", "ORDER_SIZES has duplicate (PurchaseOrder, LineItem) keys."))

    # 10. LINES/SIZES key alignment (optional)
    if validate_sizes:
        only_lines_keys = lines_key_set - sizes_key_set
        only_sizes_keys = sizes_key_set - lines_key_set
        if only_lines_keys:
            issues.append((
                "ERROR",
                f"(PO, LineItem) keys in LINES but not ORDER_SIZES: "
                f"{sorted(only_lines_keys)[:5]}"
            ))
        if only_sizes_keys:
            issues.append((
                "ERROR",
                f"(PO, LineItem) keys in ORDER_SIZES but not LINES: "
                f"{sorted(only_sizes_keys)[:5]}"
            ))

    return issues


# ─────────────────────────────────────────────────────────────────────────────
# Layout detection  –  scans up to row 80 for the best header row
# ─────────────────────────────────────────────────────────────────────────────

def _detect_layout(
    ws,
    required_keys: set[str] | None = None,
) -> tuple[int, dict[str, int], str, int, int, set[str]]:
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

        required_keys = required_keys or {"po", "qty", "product"}
        # Require at minimum: required keys with special handling for product field
        has_required = True
        for key in required_keys:
            if key == "product":
                if not (("product" in col_map) or ("product_alt" in col_map)):
                    has_required = False
                    break
            elif key not in col_map:
                has_required = False
                break
        if not has_required:
            continue

        score = len(col_map)

        # Probe how many rows below actually have PO data (if PO exists)
        if "po" in col_map:
            po_col = col_map["po"]
            probe_end = min(ws.max_row, row_idx + 300)
            nonempty_po_rows = sum(
                1 for r in range(row_idx + 1, probe_end + 1)
                if _normalize_po(ws.cell(row=r, column=po_col).value)
            )
        else:
            nonempty_po_rows = 0

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
    return 15, DEFAULT_COL_MAP.copy(), "fallback fixed-column layout", 0, 0, set(DEFAULT_COL_MAP.keys())


def _pick_source_sheet(
    wb,
    requested_sheet: str | None,
    required_keys: set[str] | None = None,
):
    if requested_sheet:
        if requested_sheet not in wb.sheetnames:
            available = ", ".join(wb.sheetnames)
            raise ValueError(f"Sheet '{requested_sheet}' not found. Available: {available}")
        ws = wb[requested_sheet]
        return (ws,) + _detect_layout(ws, required_keys)

    best_ws = wb.active
    best_data_start = 15
    best_col_map = DEFAULT_COL_MAP.copy()
    best_layout_mode = "fallback fixed-column layout"
    best_score = -1
    best_nonempty = -1
    best_detected: set[str] = set()

    for ws in wb.worksheets:
        data_start, col_map, layout_mode, score, nonempty, detected = _detect_layout(ws, required_keys)
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
    manual_po: str | None = None,
    manual_destination: str | None = None,
    manual_product_range: str | None = None,
    validate_sizes: bool = True,
) -> None:
    wb = load_workbook(input_path, data_only=True)
    manual_po_norm = _normalize_po(manual_po) if manual_po else ""
    manual_product_range = (manual_product_range or "").strip() or ""
    manual_destination = (manual_destination or "").strip() or ""
    default_qty_if_missing = bool(manual_po_norm)

    required_keys = {"product"}
    if not manual_po_norm:
        required_keys.add("po")
    if not default_qty_if_missing:
        required_keys.add("qty")

    src, data_start_row, col_map, layout_mode, layout_score, probed_po_rows, detected_keys = \
        _pick_source_sheet(wb, sheet_name, required_keys)

    orders_wb, lines_wb, sizes_wb = Workbook(), Workbook(), Workbook()
    orders_ws = orders_wb.active
    lines_ws  = lines_wb.active
    sizes_ws  = sizes_wb.active
    orders_ws.title, lines_ws.title, sizes_ws.title = "ORDERS", "LINES", "ORDER_SIZES"

    _append_row(orders_ws, ORDERS_HEADERS)
    _append_row(lines_ws,  LINES_HEADERS)
    _append_row(sizes_ws,  SIZES_HEADERS)

    seen_orders:         set[tuple[str, str]]  = set()
    line_item_counter:   dict[str, int]        = defaultdict(int)
    po_to_lines:         dict[str, list[int]]  = defaultdict(list)
    po_to_nonzero_lines: dict[str, list[int]]  = defaultdict(list)
    po_to_sizes:         dict[str, list[int]]  = defaultdict(list)

    total_buy_rows = total_lines_rows = total_sizes_rows = 0
    skipped_no_colour = 0
    skipped_empty_po = 0
    skipped_empty_po_samples: list[tuple] = []
    validation_warnings: list[str] = []
    validation_errors:   list[str] = []
    qty_default_warned = False

    def add_warning(msg: str) -> None:
        if len(validation_warnings) < 200:
            validation_warnings.append(msg)

    def add_error(msg: str) -> None:
        if len(validation_errors) < 200:
            validation_errors.append(msg)

    # Column-level checks
    if "product" not in detected_keys and "product_alt" not in detected_keys:
        msg = "Missing required column mapping for 'product' (header not detected)."
        (add_error if strict else add_warning)(msg)
    if not manual_po_norm and "po" not in detected_keys:
        msg = "Missing required column mapping for 'po' (header not detected)."
        (add_error if strict else add_warning)(msg)
    if not manual_product_range and "season" not in detected_keys:
        msg = "Missing required column mapping for 'season' (header not detected)."
        (add_error if strict else add_warning)(msg)
    if not default_qty_if_missing and "qty" not in detected_keys:
        msg = "Missing required column mapping for 'qty' (header not detected)."
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
        po = manual_po_norm or _normalize_po(raw_po_val)
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

        # ── Core fields ──────────────────────────────────────────────────────
        product = _as_text(_cell(row_idx, "product")) or _as_text(_cell(row_idx, "product_alt"))
        if not product:
            add_warning(f"Row {row_idx} PO {po}: product is empty; row skipped.")
            continue
        product_external_ref = _as_text(_cell(row_idx, "product_external_ref"))
        product_customer_ref = _as_text(_cell(row_idx, "product_customer_ref"))
        colour  = _as_text(_cell(row_idx, "colour"))
        qty_cell = _cell(row_idx, "qty")
        if qty_cell is None and default_qty_if_missing:
            qty = 1
            if not qty_default_warned:
                add_warning("Quantity column missing; defaulting Quantity=1 for all rows.")
                qty_default_warned = True
        else:
            qty = _to_int_quantity(qty_cell)

        orig_ex_fac  = _cell(row_idx, "orig_ex_fac")
        buy_date     = _cell(row_idx, "buy_date")
        trans_cond   = _cell(row_idx, "trans_cond")
        # Fix #1: Use actual season/range value, raise error if missing
        season_raw = _as_text(_cell(row_idx, "season"))
        season_value = manual_product_range or season_raw
        if not season_value:
            add_warning(f"Row {row_idx} PO {po}: season/range is empty; row skipped.")
            continue
        template_raw = _as_text(_cell(row_idx, "template"))
        brand_value  = _as_text(_cell(row_idx, "brand"))
        customer_raw = _as_text(_cell(row_idx, "customer"))
        size_raw     = _as_text(_cell(row_idx, "size"))
        vendor_code  = _as_text(_cell(row_idx, "vendor_code"))
        vendor_name  = _as_text(_cell(row_idx, "vendor_name"))
        buy_round    = _as_text(_cell(row_idx, "submit_buy"))
        status_raw   = _as_text(_cell(row_idx, "status"))

        # Build PO suffix: PO-PLANT-DEST (Dest uses transport_location)
        plant_value = _as_text(_cell(row_idx, "plant"))
        dest_country_raw = manual_destination or _as_text(_cell(row_idx, "transport_location"))
        if plant_value or dest_country_raw:
            po = "-".join([po] + [p for p in [plant_value, dest_country_raw] if p])

        total_buy_rows += 1

        # ── Derived values ───────────────────────────────────────────────────
        trans_method       = _transport_method(trans_cond)
        if not _as_text(orig_ex_fac):
            add_warning(f"Row {row_idx} PO {po}: exFtyDate is empty; delivery/cancel dates left blank.")
        key_date_obj       = _parse_date(buy_date)
        key_date_lines     = _format_date(buy_date, "%m/%d/%Y")
        delivery_date      = _format_date(orig_ex_fac, "%m/%d/%Y")
        cancel_date        = _format_date(_cell(row_idx, "cancel_date") or orig_ex_fac, "%m/%d/%Y")

        brand_lookup       = brand_value or customer_raw
        brand_config       = _get_brand_config(brand_lookup)
        customer_value     = _resolve_customer_subtype(customer_raw, brand_value, customer_fallback)
        supplier_value     = _resolve_supplier(vendor_code, vendor_name, brand_lookup)
        size_value         = size_raw or "One Size"
        product_range      = _format_product_range(season_value)
        orders_template    = _resolve_orders_template(brand_lookup, template_raw, brand_config)
        lines_template     = _resolve_lines_template(brand_lookup, template_raw, brand_config)
        status_value       = status_raw if status_raw else "Confirmed"
        comments_value     = _build_comments(
            brand_lookup, product_range, buy_date, orders_template, buy_round
        )
        keyusers           = _resolve_keyusers(brand_lookup)
        if brand_config and isinstance(brand_config.get("keyusers"), dict):
            cfg_keyusers = brand_config.get("keyusers") or {}
            keyusers = {
                "KeyUser1": _as_text(cfg_keyusers.get("KeyUser1")) or keyusers.get("KeyUser1", ""),
                "KeyUser2": _as_text(cfg_keyusers.get("KeyUser2")) or keyusers.get("KeyUser2", ""),
                "KeyUser3": _as_text(cfg_keyusers.get("KeyUser3")) or keyusers.get("KeyUser3", ""),
                "KeyUser4": _as_text(cfg_keyusers.get("KeyUser4")) or keyusers.get("KeyUser4", ""),
                "KeyUser5": _as_text(cfg_keyusers.get("KeyUser5")) or keyusers.get("KeyUser5", ""),
                "KeyUser6": _as_text(cfg_keyusers.get("KeyUser6")) or keyusers.get("KeyUser6", ""),
                "KeyUser7": _as_text(cfg_keyusers.get("KeyUser7")) or keyusers.get("KeyUser7", ""),
                "KeyUser8": _as_text(cfg_keyusers.get("KeyUser8")) or keyusers.get("KeyUser8", ""),
            }
        valid_statuses = []
        if brand_config and isinstance(brand_config.get("valid_statuses"), list):
            valid_statuses = [str(s).strip().lower() for s in brand_config.get("valid_statuses") if s]

        # ── Row-level validation warnings ────────────────────────────────────
        for field_name, field_val, fallback_desc in [
            ("season",   season_value, f"ProductRange '{product_range}'"),
            ("product",  product,      "Product missing"),
            ("template", template_raw, f"Template '{orders_template}'"),
            ("customer", customer_raw, f"Customer '{customer_value}'"),
            ("size",     size_raw,     "Size 'One Size'"),
        ]:
            if not field_val:
                msg = f"Row {row_idx} PO {po}: {field_name} is empty; fallback {fallback_desc} used."
                (add_error if strict else add_warning)(msg)

        if not brand_value:
            add_warning(f"Row {row_idx} PO {po}: brand is empty; comments use fallback brand value.")
        if valid_statuses:
            if status_value and status_value.strip().lower() not in valid_statuses:
                add_warning(
                    f"Row {row_idx} PO {po}: status '{status_value}' not in valid statuses "
                    f"{valid_statuses}."
                )

        # ── ORDERS row (one per unique PO) ───────────────────────────────────
        order_key = (po, customer_value)
        if order_key not in seen_orders:
            seen_orders.add(order_key)
            # Fix #2: Map TransportLocation from source
            transport_location = _format_transport_location(
                manual_destination or _cell(row_idx, "transport_location")
            )
            _append_row(orders_ws, [
                po, supplier_value, status_value, customer_value,
                trans_method, transport_location, "", orders_template,
                _format_date(key_date_obj, "%m/%d/%Y") if key_date_obj else "",
                "", "", comments_value, CURRENCY,
                keyusers["KeyUser1"], keyusers["KeyUser2"], keyusers["KeyUser3"],
                keyusers["KeyUser4"], keyusers["KeyUser5"], keyusers["KeyUser6"],
                keyusers["KeyUser7"], keyusers["KeyUser8"],
                "", "", "", "", "",
            ])

        # Skip LINES / SIZES if Colour is blank
        if not colour:
            skipped_no_colour += 1
            add_warning(f"Row {row_idx} PO {po}: colour is empty; line/size skipped.")
            continue

        # ── LINES row (one per buy file row) ─────────────────────────────────
        line_item_counter[po] += 1
        line_item = line_item_counter[po]

        # Fix #2: Map TransportLocation from source for LINES
        transport_location = _format_transport_location(
            manual_destination or _cell(row_idx, "transport_location")
        )
        # Fix #3: KeyDate per line from DeliveryDate
        key_date_line = delivery_date
        _append_row(lines_ws, [
            po, line_item, product_range, product, customer_value,
            delivery_date, trans_method, transport_location, status_value, "", "",
            lines_template, key_date_line, SUPPLIER_PROFILE,
            "", "", CURRENCY, "", product_external_ref, product_customer_ref, "", "",
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
    unique_order_count = len(seen_orders)
    unique_po_count = len({po for po, _ in seen_orders})
    orders_count    = orders_ws.max_row - 1

    if orders_count != unique_order_count:
        raise ValueError(
            f"ORDERS row count mismatch. Expected {unique_order_count}, got {orders_count}."
        )

    if total_lines_rows != total_buy_rows - skipped_no_colour:
        raise ValueError(
            f"LINES row count mismatch. Expected {total_buy_rows - skipped_no_colour}, got {total_lines_rows}."
        )

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
            total_buy_rows, unique_po_count, unique_order_count, orders_count,
            total_lines_rows, total_sizes_rows,
            skipped_empty_po, validation_warnings, validation_errors,
            skipped_empty_po_samples,
            validate_sizes=validate_sizes,
        )
        raise ValueError("Strict validation failed due to missing/blank required fields.")

    # ── Output slice validation (NEW) ─────────────────────────────────────────
    slice_issues = _validate_output_slices(
        orders_ws, lines_ws, sizes_ws, validate_sizes=validate_sizes
    )
    slice_errors   = [(lvl, msg) for lvl, msg in slice_issues if lvl == "ERROR"]
    slice_warnings = [(lvl, msg) for lvl, msg in slice_issues if lvl == "WARNING"]

    # ── Write output files ────────────────────────────────────────────────────
    output_dir.mkdir(parents=True, exist_ok=True)
    orders_wb.save(output_dir / "orders.xlsx")
    lines_wb.save(output_dir  / "lines.xlsx")
    sizes_wb.save(output_dir  / "order_sizes.xlsx")

    _print_summary(
        src, layout_mode, layout_score, probed_po_rows, data_start_row,
        total_buy_rows, unique_po_count, unique_order_count, orders_count,
        total_lines_rows, total_sizes_rows,
        skipped_empty_po, validation_warnings, validation_errors,
        skipped_empty_po_samples,
        slice_issues=slice_issues,
        validate_sizes=validate_sizes,
    )

    if slice_errors and strict:
        raise ValueError(
            f"Output slice validation failed with {len(slice_errors)} error(s). "
            "Files written but review the report above."
        )
    print(f"\nGenerated files in: {output_dir.resolve()}")
    for fname in ("orders.xlsx", "lines.xlsx", "order_sizes.xlsx"):
        print(f"  - {fname}")


def _print_summary(
    src, layout_mode, layout_score, probed_po_rows, data_start_row,
    total_buy_rows, unique_po_count, unique_order_count, orders_count,
    total_lines_rows, total_sizes_rows,
    skipped_empty_po, validation_warnings, validation_errors,
    skipped_empty_po_samples,
    validate_sizes: bool = True,
    slice_issues: list[tuple[str, str]] | None = None,
) -> None:
    lines_equals_sizes = total_lines_rows == total_sizes_rows
    slice_issues = slice_issues or []
    slice_errors   = [msg for lvl, msg in slice_issues if lvl == "ERROR"]
    slice_warnings = [msg for lvl, msg in slice_issues if lvl == "WARNING"]
    overall_status = "✓ PASS" if not slice_errors else f"✗ FAIL ({len(slice_errors)} error(s))"

    print("\n── Generation Summary ──────────────────────────────────────────")
    print(f"  Source sheet   : {src.title}")
    print(f"  Layout mode    : {layout_mode}")
    print(f"  Layout score   : {layout_score}")
    print(f"  PO probe rows  : {probed_po_rows}")
    print(f"  Data start row : {data_start_row}")
    print(f"  Buy rows       : {total_buy_rows}")
    print(f"  Unique POs     : {unique_po_count}")
    print(f"  Unique Orders  : {unique_order_count} (PO+Customer)")
    print(f"  ORDERS rows    : {orders_count}")
    print(f"  LINES rows     : {total_lines_rows}")
    if skipped_no_colour:
        print(f"  Skipped (no Colour): {skipped_no_colour}")
    print(f"  ORDER_SIZES    : {total_sizes_rows}")
    if validate_sizes:
        print(f"  LINES==SIZES   : {'YES' if lines_equals_sizes else 'NO (qty=0 rows excluded in ORDER_SIZES)'}")
    else:
        print("  LINES==SIZES   : SKIPPED (sizes validation disabled)")
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

    print("\n── Output Slice Validation ─────────────────────────────────────")
    print(f"  Overall        : {overall_status}")
    checks = [
        "Column counts (ORDERS=26, LINES=31)",
        "ORDERS has no duplicate PO+Customer keys",
        "TransportMethod mapped to Sea, Air, or Courier only",
        "No blank Customer values",
        "ORDERS Status = Confirmed",
        "No duplicate PO+LineItem keys",
    ]
    if validate_sizes:
        checks = [
            "Row counts (LINES == SIZES)",
            "Column counts (ORDERS=26, LINES=31, SIZES=29)",
            "PO universe matches across all 3 files",
            "ORDERS has no duplicate PO+Customer keys",
            "ORDER_SIZES total Quantity > 0",
            "TransportMethod mapped to Sea, Air, or Courier only",
            "No blank Customer values",
            "ORDERS Status = Confirmed",
            "No duplicate PO+LineItem keys",
            "LINES/SIZES key alignment",
        ]
    issue_msgs = [msg for _, msg in slice_issues]
    for check in checks:
        hit = any(
            any(keyword in msg for keyword in check.lower().split())
            for msg in issue_msgs
        )
        print(f"  {'✗' if hit else '✓'} {check}")
    if slice_errors:
        print("  Errors:")
        for msg in slice_errors:
            print(f"    ✗ {msg}")
    if slice_warnings:
        print("  Warnings:")
        for msg in slice_warnings:
            print(f"    ⚠ {msg}")
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
    parser.add_argument("--po",         dest="manual_po",        default=None)
    parser.add_argument("--destination", dest="manual_destination", default=None)
    parser.add_argument("--product-range", dest="manual_product_range", default=None)
    parser.add_argument("--validate-sizes", dest="validate_sizes", action="store_true")
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
        manual_po         = args.manual_po,
        manual_destination = args.manual_destination,
        manual_product_range = args.manual_product_range,
        validate_sizes    = args.validate_sizes,
    )
