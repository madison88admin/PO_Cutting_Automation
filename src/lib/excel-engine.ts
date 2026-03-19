import ExcelJS from "exceljs";
import { logEvent } from "@/lib/audit";
import { getFactoryMapping, getMloMapping, getColumnMapping, getAllColumnMappings } from "@/lib/data-loader";
import { updateRun } from "@/lib/db/runHistory";

export interface POHeader {
    purchaseOrder: string;
    productSupplier: string;
    status: string;
    customer: string;
    transportMethod: string;
    transportLocation: string;
    ordersTemplate: string;
    linesTemplate: string;
    keyDate: string | Date;
    keyDateFormat?: "manual" | "standard";
    comments: string;
    currency: string;
    keyUser1: string;
    keyUser2: string;
    keyUser3: string;
    keyUser4: string;
    keyUser5: string;
    keyUser6: string;
    keyUser7: string;
    keyUser8: string;
}

export interface POLine {
    lineItem: number;
    productRange: string;
    styleNumber: string;
    supplierProfile: string;
    buyerPoNumber: string | number;
    startDate: string | Date;
    cancelDate: string | Date;
    cost?: string | number;
    colour: string;
    productExternalRef: string;
    productCustomerRef: string;
}

export interface POSize {
    productSize: string;
    quantity: number;
}

interface ProductSheetRow {
    colour: string;
    factory?: string;
    cost?: string | number;
    customerName?: string;
    productName?: string;
    buyerStyleNumber?: string;
}

export interface ValidationError {
    field: string;
    row: number;
    message: string;
    severity: "CRITICAL" | "WARNING";
}

export interface ProcessedPO {
    header: POHeader;
    lines: POLine[];
    sizes: Record<number, POSize[]>;
}

// ─────────────────────────────────────────────────────────────────────────────
// Brand-level lookup tables (used when buy file provides brand but not
// an explicit supplier / customer name).
// ─────────────────────────────────────────────────────────────────────────────
const BRAND_SUPPLIER_MAP: Record<string, string> = {
    col: "MSO",
    columbia: "MSO",
    tnf: "PT. UWU JUMP INDONESIA",
    "the north face": "PT. UWU JUMP INDONESIA",
    arcteryx: "PT UWU JUMP INDONESIA",
    "arc'teryx": "PT UWU JUMP INDONESIA",
};

const BRAND_CUSTOMER_MAP: Record<string, string> = {
    col: "Columbia",
    columbia: "Columbia",
    tnf: "The North Face In-Line",
    "the north face": "The North Face In-Line",
    arcteryx: "Arcteryx",
    "arc'teryx": "Arcteryx",
};

const TNF_CUSTOMER_SUBTYPE_MAP: Record<string, string> = {
    "the north face in-line": "The North Face In-Line",
    "the north face inline": "The North Face In-Line",
    "the north face rto": "The North Face RTO",
    "the north face smu": "The North Face SMU",
    "tnf in-line": "The North Face In-Line",
    "tnf inline": "The North Face In-Line",
    "tnf rto": "The North Face RTO",
    "tnf smu": "The North Face SMU",
};

const TRANSPORT_MAP: Record<string, string> = {
    "ocean": "Sea",
    "sea": "Sea",
    "sea freight": "Sea",
    "seafreight": "Sea",
    "s1 - seafreight": "Sea",
    "s1": "Sea",
    "air": "Air",
    "air freight": "Air",
    "airfreight": "Air",
    "a1 - airfreight": "Air",
    "a1": "Air",
    "a2 - airfreight": "Air",
    "a2": "Air",
    "courier": "Courier",
    "dhl": "Courier",
    "fedex": "Courier",
    "ups": "Courier",
};

const VALID_TRANSPORT_VALUES = new Set(["Sea", "Air", "Courier"]);

const COUNTRY_NAME_MAP: Record<string, string> = {
    AE: "UAE",
    AR: "Argentina",
    AT: "Austria",
    AU: "Australia",
    BR: "Brazil",
    CA: "Canada",
    CH: "Switzerland",         // ← BUG FIX (was "China")
    CL: "Chile",
    CN: "China",
    CZ: "Czech Republic",
    DE: "Germany",
    DK: "Denmark",
    EC: "Ecuador",
    ES: "Spain",
    FR: "France",
    GB: "UK",
    GR: "Greece",
    HK: "Hong Kong",
    HR: "Croatia",
    HU: "Hungary",
    ID: "Indonesia",
    IL: "Israel",
    IN: "India",
    IT: "Italy",
    JP: "Japan",
    KR: "Korea",
    MN: "Mongolia",
    NP: "Nepal",
    MT: "Malta",
    MX: "Mexico",
    MY: "Malaysia",
    PA: "Panama",
    PE: "Peru",
    PH: "Philippines",
    PL: "Poland",
    RS: "Serbia",
    RU: "Russia",
    TH: "Thailand",
    TR: "Turkey",
    TW: "Taiwan",
    UK: "UK",
    US: "USA",
    "UNITED KINGDOM": "UK",
    "UNITED ARAB EMIRATES": "UAE",
    "UNITED STATES": "USA",
    UY: "Uruguay",
    VN: "Vietnam",
    ZA: "South Africa",
};

// ─────────────────────────────────────────────────────────────────────────────
// FIX 1: KeyUser map — per brand, the MLO team members for ORDERS.
// KeyUser1 = Planning, KeyUser2 = Purchasing, KeyUser4 = Production,
// KeyUser5 = Logistics/Shipping. KeyUser3/6/7/8 left blank per BRD.
// ─────────────────────────────────────────────────────────────────────────────
interface KeyUsers {
    k1: string; k2: string; k3: string;
    k4: string; k5: string; k6: string;
    k7: string; k8: string;
}

const BRAND_KEYUSER_MAP: Record<string, KeyUsers> = {
    tnf: {
        k1: "Ron", k2: "Maricar", k3: "",
        k4: "Ron", k5: "Elaine Sanchez", k6: "", k7: "", k8: "",
    },
    "the north face": {
        k1: "Ron", k2: "Maricar", k3: "",
        k4: "Ron", k5: "Elaine Sanchez", k6: "", k7: "", k8: "",
    },
    col:      { k1: "", k2: "", k3: "", k4: "", k5: "", k6: "", k7: "", k8: "" },
    columbia: { k1: "", k2: "", k3: "", k4: "", k5: "", k6: "", k7: "", k8: "" },
    arcteryx: { k1: "", k2: "", k3: "", k4: "", k5: "", k6: "", k7: "", k8: "" },
    "arc'teryx": { k1: "", k2: "", k3: "", k4: "", k5: "", k6: "", k7: "", k8: "" },
};

const DEFAULT_KEYUSERS: KeyUsers = { k1: "", k2: "", k3: "", k4: "", k5: "", k6: "", k7: "", k8: "" };

// ─────────────────────────────────────────────────────────────────────────────
// FIX 2: Separate template maps for ORDERS and LINES per brand.
// ORDERS and LINES use different template values for TNF.
// ─────────────────────────────────────────────────────────────────────────────
const BRAND_ORDERS_TEMPLATE_MAP: Record<string, string> = {
    tnf:              "Major Brand Bulk",
    "the north face": "Major Brand Bulk",
    col:              "BULK",
    columbia:         "BULK",
    arcteryx:         "BULK",
    "arc'teryx":      "BULK",
};

const BRAND_LINES_TEMPLATE_MAP: Record<string, string> = {
    tnf:              "FOB Bulk EDI PO (New)",
    "the north face": "FOB Bulk EDI PO (New)",
    col:              "BULK",
    columbia:         "BULK",
    arcteryx:         "BULK",
    "arc'teryx":      "BULK",
};

export class ExcelEngine {
    private errors: ValidationError[] = [];
    private runId: string | null = null;
    private userId: string | null = null;

    constructor(runId?: string, userId?: string) {
        this.runId = runId || null;
        this.userId = userId || null;
    }

    // ── Header row detection (scans up to row 80, same as Python version) ──

    private detectHeaderRow(worksheet: ExcelJS.Worksheet): number {
        const KNOWN_HEADERS = [
            'erp ind', 'brand', 'po #', 'pono', 'purchase order',
            'purchaseorder', 'lineitem', 'productrange', 'company code', 'vendor code',
            'material style', 'jde style', 'doc type', 'orig ex fac', 'trans cond',
            'ordered qty', 'buy date', 'color', 'season',
            'tracking number', 'article', 'business unit description',
            'requested qty', 'ex-factory', 'transport mode',
            'qty', 'quantity', 'size', 'colour',
            'product name', 'buyer style number', 'buyer style name', 'customer name', 'factory',
        ];

        let bestRow = 1;
        let bestMatches = 0;

        for (let i = 1; i <= Math.min(80, worksheet.rowCount); i++) {
            const row = worksheet.getRow(i);
            const values: string[] = [];
            row.eachCell(cell => {
                const val = cell.value?.toString().toLowerCase().trim() || '';
                if (val) values.push(val);
            });

            const matches = KNOWN_HEADERS.filter(h => values.includes(h)).length;
            if (matches > bestMatches) {
                bestMatches = matches;
                bestRow = i;
            }
            if (matches >= 8) break;
        }

        return bestRow;
    }

    // ── Comprehensive global alias map ───────────────────────────────────────

    private getFallbackColumnAliases(): Record<string, string> {
        return {
            'transportlocation': 'transportLocation',
            'transport location': 'transportLocation',
            'destination': 'transportLocation',
            'dest country': 'transportLocation',
            'plant': 'plant',
            'plant code': 'plant',
            'ult. destination': 'transportLocation',
            'purchaseorder': 'purchaseOrder',
            'po': 'purchaseOrder',
            'po #': 'purchaseOrder',
            'po#': 'purchaseOrder',
            'pono': 'purchaseOrder',
            'purchase order': 'purchaseOrder',
            'tracking number': 'purchaseOrder',
            'extraction po #': 'buyerPoNumber',
            'extraction po#': 'buyerPoNumber',
            'style number': 'product',
            'style no': 'product',
            'material style': 'product',
            'style': 'product',
            'product': 'product',
            'sku': 'product',
            'item': 'product',
            'article': 'product',
            'model': 'product',
            'product name': 'product',
            'buyer style number': 'productCustomerRef',
            'buyer style no': 'productCustomerRef',
            'buyer style #': 'productCustomerRef',
            'buyer style': 'productCustomerRef',
            'name': 'productExternalRef',
            'product external ref': 'productExternalRef',
            'product customer ref': 'productCustomerRef',
            'vendor code': 'productSupplier',
            'vendorcode': 'productSupplier',
            'vendor': 'productSupplier',
            'supplier': 'productSupplier',
            'product supplier': 'productSupplier',
            'productsupplier': 'productSupplier',
            'vendor name': 'vendorName',
            'vendorname': 'vendorName',
            'supplier name': 'vendorName',
            'factory': 'vendorName',
            'size': 'sizeName',
            'size name': 'sizeName',
            'sizename': 'sizeName',
            'productsize': 'sizeName',
            'product size': 'sizeName',
            'size #': 'sizeName',
            'size#': 'sizeName',
            'size code': 'sizeName',
            'size_name': 'sizeName',
            'size-name': 'sizeName',
            'ordered qty': 'quantity',
            'open qty (pcs/prs)': 'quantity',
            'requested qty': 'quantity',
            'quantity': 'quantity',
            'qty': 'quantity',
            'deliverydate': 'exFtyDate',
            'delivery date': 'exFtyDate',
            'orig ex fac': 'exFtyDate',
            'negotiated ex fac date': 'exFtyDate',
            'ex fac': 'exFtyDate',
            'ex-factory': 'exFtyDate',
            'confirmed fty ex fac': 'confirmedExFac',
            'confirmed ex fac': 'confirmedExFac',
            'fty ex fac': 'confirmedExFac',
            'keydate': 'poIssuanceDate',
            'buy date': 'buyDate',
            'file date': 'buyDate',
            'cancel date': 'cancelDate',
            'canceldate': 'cancelDate',
            'cancel': 'cancelDate',
            'po issuance date': 'poIssuanceDate',
            'transportmethod': 'transportMethod',
            'transport method': 'transportMethod',
            'trans cond': 'transportMethod',
            'transport mode': 'transportMethod',
            'doc type': 'template',
            'template': 'template',
            'range': 'season',
            'productrange': 'season',
            'season': 'season',
            'brand': 'brand',
            'business unit description': 'brand',
            'customer': 'customerName',
            'customer name': 'customerName',
            'plm customer name': 'customerName',
            'status': 'status',
            'confirmation status': 'status',
            'gsc type': 'status',
            'colour': 'colour',
            'color': 'colour',
            'color name': 'colour',
            'article name': 'colour',
            'color description': 'colour',
            'submit buy': 'buyRound',
            'buy round': 'buyRound',
            'buyer style name': 'ignore',
            'jde style': 'jdeStyle',
            'udf-buyer_po_number': 'buyerPoNumber',
            'udf-start_date': 'exFtyDate',
            'udf-canel_date': 'cancelDate',
        };
    }

    private getProductSheetAliases(): Record<string, string> {
        return {
            'color name': 'colour',
            'colour name': 'colour',
            'color': 'colour',
            'colour': 'colour',
            'factory': 'factory',
            'vendor code': 'factory',
            'vendorcode': 'factory',
            'cost': 'cost',
            'customer name': 'customerName',
            'customer': 'customerName',
            'product name': 'productName',
            'product': 'productName',
            'buyer style number': 'buyerStyleNumber',
            'buyer style no': 'buyerStyleNumber',
            'buyer style #': 'buyerStyleNumber',
            'buyer style': 'buyerStyleNumber',
        };
    }

    private detectProductSheet(worksheet: ExcelJS.Worksheet): { isProductSheet: boolean; headerRow: number } {
        const headerRow = this.detectHeaderRow(worksheet);
        const header = worksheet.getRow(headerRow);
        const aliases = this.getProductSheetAliases();
        const productHeaders = new Set(Object.keys(aliases));
        const buyHeaders = new Set([
            'po #', 'pono', 'purchase order', 'purchaseorder', 'lineitem',
            'quantity', 'qty', 'size', 'season', 'brand', 'productrange',
        ]);

        let productScore = 0;
        let buyScore = 0;

        header.eachCell(cell => {
            const val = cell.value?.toString().toLowerCase().trim() || '';
            if (productHeaders.has(val)) productScore++;
            if (buyHeaders.has(val)) buyScore++;
        });

        return { isProductSheet: productScore >= 3 && buyScore <= 1, headerRow };
    }

    private normalizeColourKey(value: string): string {
        const raw = this.stripBrackets(value || '').toLowerCase().trim();
        if (!raw) return '';
        if (/^\d+(\.\d+)?$/.test(raw)) {
            const num = Number(raw);
            if (Number.isFinite(num)) return String(Math.trunc(num));
        }
        const digits = raw.match(/\d+/);
        if (digits && digits[0]) {
            const normalized = digits[0].replace(/^0+/, '');
            return normalized || '0';
        }
        return raw;
    }

    private extractProductSheetMapFromWorkbook(
        workbook: ExcelJS.Workbook,
    ): Record<string, ProductSheetRow[]> {
        const result: Record<string, ProductSheetRow[]> = {};
        const aliases = this.getProductSheetAliases();
        const seenEntries = new Set<string>(); // deduplicate across worksheets

        for (const ws of workbook.worksheets) {
            const { isProductSheet, headerRow } = this.detectProductSheet(ws);
            if (!isProductSheet) continue;

            const header = ws.getRow(headerRow);
            const headerMap: Record<string, number> = {};
            header.eachCell((cell, colNumber) => {
                const key = cell.value?.toString().toLowerCase().trim() || '';
                const mapped = aliases[key];
                if (mapped && !headerMap[mapped]) headerMap[mapped] = colNumber;
            });

            if (!headerMap['colour']) continue;

            ws.eachRow((row, rowNumber) => {
                if (rowNumber <= headerRow) return;
                const getRaw = (field: string) => {
                    const col = headerMap[field];
                    if (!col) return undefined;
                    return this.getCellValue(row.getCell(col));
                };
                const colourRaw = getRaw('colour')?.toString().trim() || '';
                const colourKey = this.normalizeColourKey(colourRaw);
                const buyerStyleNumber = getRaw('buyerStyleNumber')?.toString().trim() || '';
                if (!colourKey || !buyerStyleNumber) return;

                const factoryRaw = getRaw('factory');
                const costRaw = getRaw('cost');
                const customerRaw = getRaw('customerName');
                const productRaw = getRaw('productName');

                const entry: ProductSheetRow = {
                    colour: colourRaw,
                    factory: factoryRaw?.toString().trim() || '',
                    cost: typeof costRaw === 'number' ? costRaw : costRaw?.toString().trim(),
                    customerName: customerRaw?.toString().trim() || '',
                    productName: productRaw?.toString().trim() || '',
                    buyerStyleNumber,
                };
                // Build lookup keys: raw buyer_style_number AND each slash-separated segment
                // e.g. "217554/CU2279" → also index under "CU2279" to match buy file JDE Style
                // Exact matches are inserted first so they win over slash-segment duplicates
                const lookupKeys = new Map<string, boolean>(); // key → isExact
                lookupKeys.set(buyerStyleNumber, true);
                buyerStyleNumber.split('/').forEach(part => {
                    const p = part.trim();
                    if (p && p !== buyerStyleNumber) lookupKeys.set(p, false);
                });
                for (const [lk, isExact] of lookupKeys) {
                    const lkKey = `${lk}|${colourKey}`;
                    const dedupKey = `${lkKey}|${entry.colour}|${entry.factory}|${entry.productName}|${entry.customerName}`;
                    if (seenEntries.has(dedupKey)) continue;
                    seenEntries.add(dedupKey);
                    if (!result[lkKey]) result[lkKey] = [];
                    // Exact full buyer_style_number matches go first
                    if (isExact) {
                        result[lkKey].unshift(entry);
                    } else {
                        result[lkKey].push(entry);
                    }
                }
            });
        }

        return result;
    }

    // ── Supplier resolution with brand fallback ───────────────────────────────

    private resolveSupplier(
        vendorCode: string | undefined,
        vendorName: string | undefined,
        brand: string | undefined,
        category: string | undefined,
        factoryMap: any[],
    ): string {
        const vCode = this.stripBrackets(vendorCode || '').trim();
        const vName = this.stripBrackets(vendorName || '').trim();
        const b = this.stripBrackets(brand || '').trim();
        const cat = this.stripBrackets(category || '').trim();
        if (vCode && vCode.length > 2) return vCode;
        if (vName && vName.length > 2) return vName;

        // Factory mapping: brand + category → product_supplier
        if (b && cat) {
            const match = factoryMap.find(
                (f: any) =>
                    f.brand?.toLowerCase() === b.toLowerCase() &&
                    f.category?.toLowerCase() === cat.toLowerCase()
            );
            if (match?.product_supplier) return match.product_supplier;
        }

        // Brand-only factory match (when category is unknown)
        if (b) {
            const brandMatches = factoryMap.filter(
                (f: any) => f.brand?.toLowerCase() === b.toLowerCase() && f.product_supplier
            );
            if (brandMatches.length === 1) return brandMatches[0].product_supplier;
        }

        // Hardcoded brand fallback
        const key = b.toLowerCase();
        return BRAND_SUPPLIER_MAP[key] || 'MISSING_SUPPLIER';
    }

    // ── Customer name resolution with subtype support ─────────────────────────
    // FIX 3: DB mappedCustomerName no longer takes blind priority.
    // Subtype map is checked first so RTO/SMU values from the buy file
    // are never overwritten by a base-brand DB entry.

    private resolveCustomer(
        customerRaw: string | undefined,
        brand: string | undefined,
        detectedCustomer: string,
        mappedCustomerName: string | undefined,
    ): string {
        const raw = this.stripBrackets(customerRaw || '').trim();
        const brandClean = this.stripBrackets(brand || '').trim();
        const mapped = this.stripBrackets(mappedCustomerName || '').trim();

        // 1. Check subtype map first (handles RTO / SMU / In-Line variants)
        if (raw) {
            const key = raw.toLowerCase();
            if (TNF_CUSTOMER_SUBTYPE_MAP[key]) return TNF_CUSTOMER_SUBTYPE_MAP[key];
        }

        // 2. DB-mapped customer name — only use if it looks like a subtype-aware
        //    value (i.e. contains RTO / SMU / In-Line) OR if raw is empty.
        //    This prevents a base-brand DB entry ("The North Face") from
        //    overwriting a correctly detected subtype from the buy file.
        if (mapped) {
            const mappedKey = mapped.toLowerCase();
            const hasSubtype = mappedKey.includes('rto') || mappedKey.includes('smu') || mappedKey.includes('in-line') || mappedKey.includes('inline');
            if (!raw || hasSubtype) return mapped;
        }

        // 3. BRAND_CUSTOMER_MAP from raw value
        if (raw) {
            const key = raw.toLowerCase();
            if (BRAND_CUSTOMER_MAP[key]) return BRAND_CUSTOMER_MAP[key];
            return raw;
        }

        // 4. Brand fallback
        if (brandClean) {
            const key = brandClean.toLowerCase();
            if (BRAND_CUSTOMER_MAP[key]) return BRAND_CUSTOMER_MAP[key];
        }

        // 5. Detected customer from DB registration
        if (detectedCustomer && detectedCustomer !== 'DEFAULT') return detectedCustomer;

        return brandClean.toUpperCase() || 'COL';
    }

    // FIX 1: KeyUser resolution with hardcoded brand fallback ─────────────────
    private resolveKeyUsers(
        brand: string | undefined,
        manualK1: string | undefined,
        manualK2: string | undefined,
        manualK3: string | undefined,
        manualK4: string | undefined,
        manualK5: string | undefined,
        providedK1: string | undefined,
        providedK2: string | undefined,
        providedK4: string | undefined,
        providedK5: string | undefined,
        mloRow: any,
    ): KeyUsers {
        const hasManual =
            !!(manualK1 || manualK2 || manualK3 || manualK4 || manualK5);
        if (hasManual) {
            return {
                k1: manualK1 || '',
                k2: manualK2 || '',
                k3: manualK3 || '',
                k4: manualK4 || '',
                k5: manualK5 || '',
                k6: '', k7: '', k8: '',
            };
        }
        // If buy file already has explicit KeyUser values, use them
        if (providedK1 || providedK2) {
            return {
                k1: providedK1 || '',
                k2: providedK2 || '',
                k3: '',
                k4: providedK4 || '',
                k5: providedK5 || '',
                k6: '', k7: '', k8: '',
            };
        }

        // DB MLO map takes next priority
        if (mloRow) {
            return {
                k1: mloRow.keyuser1 || '',
                k2: mloRow.keyuser2 || '',
                k3: '',
                k4: mloRow.keyuser4 || '',
                k5: mloRow.keyuser5 || '',
                k6: '', k7: '', k8: '',
            };
        }

        // Hardcoded brand fallback — so TNF always gets correct values
        // even when DB MLO table is empty
        const key = (brand || '').trim().toLowerCase();
        return { ...( BRAND_KEYUSER_MAP[key] || DEFAULT_KEYUSERS) };
    }

    // FIX 2: Separate template resolvers for ORDERS and LINES ─────────────────
    private resolveOrdersTemplate(brand: string | undefined, rawTemplate: string): string {
        const key = (brand || '').trim().toLowerCase();
        if (BRAND_ORDERS_TEMPLATE_MAP[key]) return BRAND_ORDERS_TEMPLATE_MAP[key];
        return this.normalizeTemplate(rawTemplate);
    }

    private resolveLinesTemplate(brand: string | undefined, rawTemplate: string): string {
        const key = (brand || '').trim().toLowerCase();
        if (BRAND_LINES_TEMPLATE_MAP[key]) return BRAND_LINES_TEMPLATE_MAP[key];
        return this.normalizeTemplate(rawTemplate);
    }

    private looksLikeSizeHeader(headerText: string): boolean {
        const normalized = headerText.trim().toLowerCase();
        const directMatches = new Set([
            'size', 'size name', 'sizename', 'productsize', 'product size',
            'size #', 'size#', 'size code', 'size_name', 'size-name',
        ]);
        if (directMatches.has(normalized)) return true;
        return normalized.includes('size') && !normalized.includes('status');
    }

    private shouldSilentlyIgnoreHeader(headerText: string): boolean {
        const normalized = headerText.trim().toLowerCase();
        const exactIgnore = new Set([
            'unit total', 'confirmed unit total', 'vendor comments', 'vendor confirmed',
            'csc/lo comments', 'lo reviewed', 'lo rejected', 'csc confirmed',
            'csc rejected', 'last collab status date', 'hashcode', 'linehashcode',
            'mainitem_id', 'activity_info', 'modifyrivision', 'rawinfo',
            'writablecells', 'rowsuffix',
            'vendor price chg 1', 'price chg type 1',
            'vendor price chg 2', 'price chg type 2',
            'vendor price chg 3', 'price chg type 3',
            'net price chg', 'absolute price chg',
            'line #s 2', 'line #s',
            'lineitem', 'purchaseprice', 'sellingprice',
            'supplierprofile', 'closeddate', 'comments', 'currency', 'archivedate',
            'productexternalref', 'productcustomerref', 'purchaseuom', 'sellinguom',
            'paymentterm', 'defaultdeliverydate', 'productsupplierext',
            'keyuser1', 'keyuser2', 'keyuser3', 'keyuser4', 'keyuser5',
            'keyuser6', 'keyuser7', 'keyuser8',
            'department', 'customattribute1', 'customattribute2', 'customattribute3',
            'lineratio', 'colourext', 'customerext', 'departmentext',
            'customattribute1ext', 'customattribute2ext', 'customattribute3ext',
            'file date', 'sku', 'sku description', 'model description', 'gsc type',
            'product group description', 'product line description', 'planning category',
            'transit vendor destination', 'official buy',
            'storage location', 'stock segment',
            'erp ind', 'company code', 'ab number', 'gtn issue date',
            'sku status', 'slo', 'plo',
            'priority flag', 'lb', 'tooling code', 'vas', 'capacity type',
        ]);
        if (exactIgnore.has(normalized)) return true;
        const ignorePrefixes = ['findfield_', 'udf-inspection', 'udf-report', 'udf-inspector',
            'udf-approval', 'udf-submitted'];
        return ignorePrefixes.some(p => normalized.startsWith(p));
    }

    private inferCategoryFromFactoryMap(brand: string | undefined, factoryMap: any[]): string | undefined {
        if (!brand) return undefined;
        const matches = factoryMap
            .filter((f: any) => f.brand?.toLowerCase() === brand.toLowerCase())
            .map((f: any) => f.category)
            .filter(Boolean);
        const unique = Array.from(new Set(matches));
        return unique.length === 1 ? unique[0] : undefined;
    }

    private formatProductRange(season: string): string {
        const normalized = this.stripBrackets(season || '').trim();
        const m = normalized.match(/^([FS])(?:W|S)?(\d{2})$/i);
        if (m) return `${m[1].toUpperCase()}H:20${m[2]}`;
        if (normalized) return normalized;
        return 'FH:2026';
    }

    private normalizeTemplate(rawTemplate: string): string {
        const normalized = (rawTemplate || '').trim().toUpperCase();
        const map: Record<string, string> = {
            OG: 'BULK', ZNB1: 'BULK', ZMF1: 'BULK', ZDS1: 'BULK', SMS: 'SMS',
        };
        return map[normalized] || (rawTemplate || 'BULK').trim() || 'BULK';
    }

    private normalizeTransportMethod(raw: string | undefined): string {
        const key = (raw || '').trim().toLowerCase();
        const mapped = TRANSPORT_MAP[key];
        if (mapped) return mapped;
        // Warn if non-empty raw value didn't map
        return raw ? raw.trim() : 'Sea';
    }

    private normalizeTransportLocation(raw: string | undefined): string {
        const cleaned = this.stripBrackets(raw || '').trim();
        if (!cleaned) return '';
        const key = cleaned.toUpperCase();
        return COUNTRY_NAME_MAP[key] || cleaned;
    }

    private parseDate(raw: string | Date | undefined): Date | null {
        if (!raw) return null;
        if (raw instanceof Date) return isNaN(raw.getTime()) ? null : raw;
        const text = String(raw).trim();
        if (!text) return null;

        const isoMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (isoMatch) {
            const date = new Date(Number(isoMatch[1]), Number(isoMatch[2]) - 1, Number(isoMatch[3]));
            return isNaN(date.getTime()) ? null : date;
        }

        const usMatch = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
        if (usMatch) {
            const date = new Date(Number(usMatch[3]), Number(usMatch[1]) - 1, Number(usMatch[2]));
            return isNaN(date.getTime()) ? null : date;
        }

        const monMatch = text.match(/^(\d{1,2})-([A-Za-z]+)-(\d{4})$/);
        if (monMatch) {
            const months = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
            const monthIndex = months.findIndex(m => monMatch[2].toLowerCase().startsWith(m));
            if (monthIndex >= 0) {
                const date = new Date(Number(monMatch[3]), monthIndex, Number(monMatch[1]));
                return isNaN(date.getTime()) ? null : date;
            }
        }

        return null;
    }

    private formatDateString(raw: string | Date | undefined): string {
        const date = this.parseDate(raw);
        if (!date) return '';
        const mm = String(date.getMonth() + 1).padStart(2, '0');
        const dd = String(date.getDate()).padStart(2, '0');
        return `${mm}/${dd}/${date.getFullYear()}`;
    }

    private formatManualDateString(raw: string | Date | undefined): string {
        const date = this.parseDate(raw);
        if (!date) return '';
        return `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
    }

    private stripBrackets(value: string): string {
        if (!value) return value;
        const cleaned = value.replace(/\[([^\]]+)\]/g, '$1').replace(/\[|\]/g, '');
        return cleaned.replace(/\s+/g, ' ').trim();
    }

    private buildComments(brand: string | undefined, season: string, buyRound: string, buyDateRaw: string | undefined, template: string): string {
        const b = this.stripBrackets(brand || 'OUTPUT');
        const s = this.stripBrackets(season || 'NOS');
        const round = this.stripBrackets(buyRound || '');
        const tmpl = this.stripBrackets(template || '');
        if (round) return `${b} ${s} ${round} ${tmpl}`.trim();
        const parsed = this.parseDate(buyDateRaw);
        if (parsed) {
            const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
            const monShort = months[parsed.getMonth()];
            const day = String(parsed.getDate()).padStart(2, '0');
            const suffix = tmpl ? ` ${tmpl}` : '';
            return `${b} ${s} ${monShort} Buy ${day}-${monShort.toUpperCase()}${suffix}`.trim();
        }
        return `${b} ${s}`.trim();
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Core processing method
    // ─────────────────────────────────────────────────────────────────────────

    async extractProductSheetMap(buffer: any): Promise<Record<string, ProductSheetRow[]>> {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);
        return this.extractProductSheetMapFromWorkbook(workbook);
    }

    async analyzeWorkbook(buffer: any): Promise<{ productSheetMap: Record<string, ProductSheetRow[]>; hasBuySheet: boolean }> {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);

        const productSheetMap = this.extractProductSheetMapFromWorkbook(workbook);
        const aliases = this.getFallbackColumnAliases();
        let hasBuySheet = false;

        for (const ws of workbook.worksheets) {
            const { isProductSheet, headerRow } = this.detectProductSheet(ws);
            if (isProductSheet) continue;

            const row = ws.getRow(headerRow);
            let score = 0;
            row.eachCell(cell => {
                const v = cell.value?.toString().toLowerCase().trim() || '';
                if (aliases[v]) score++;
            });
            if (score >= 4) {
                hasBuySheet = true;
                break;
            }
        }

        return { productSheetMap, hasBuySheet };
    }

    async processBuyFile(
        buffer: any,
        options?: {
            manualPurchaseOrder?: string;
            manualDestination?: string;
            manualProductRange?: string;
            manualTemplate?: string;
            manualComments?: string;
            manualKeyDate?: string;
            manualKeyUser1?: string;
            manualKeyUser2?: string;
            manualKeyUser3?: string;
            manualKeyUser4?: string;
            manualKeyUser5?: string;
            defaultQuantityIfMissing?: boolean;
            productSheetMap?: Record<string, ProductSheetRow[]>;
        },
    ): Promise<{ data: ProcessedPO[]; errors: ValidationError[] }> {
        const manualPurchaseOrder = options?.manualPurchaseOrder?.toString().trim() || '';
        const manualDestination = options?.manualDestination?.toString().trim() || '';
        const manualProductRange = options?.manualProductRange?.toString().trim() || '';
        const manualTemplate = options?.manualTemplate?.toString().trim() || '';
        const manualComments = options?.manualComments?.toString().trim() || '';
        const manualKeyDate = options?.manualKeyDate?.toString().trim() || '';
        const manualKeyUser1 = options?.manualKeyUser1?.toString().trim() || '';
        const manualKeyUser2 = options?.manualKeyUser2?.toString().trim() || '';
        const manualKeyUser3 = options?.manualKeyUser3?.toString().trim() || '';
        const manualKeyUser4 = options?.manualKeyUser4?.toString().trim() || '';
        const manualKeyUser5 = options?.manualKeyUser5?.toString().trim() || '';
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);

        const workbookProductMap = this.extractProductSheetMapFromWorkbook(workbook);
        const productSheetMap: Record<string, ProductSheetRow[]> = {
            ...(options?.productSheetMap || {}),
            ...workbookProductMap,
        };

        let worksheet = workbook.worksheets[0];
        let headerRowNumber = this.detectHeaderRow(worksheet);
        let bestScore = -1;

        for (const ws of workbook.worksheets) {
            const candidate = this.detectHeaderRow(ws);
            const row = ws.getRow(candidate);
            let score = 0;
            const aliases = this.getFallbackColumnAliases();
            row.eachCell(cell => {
                const v = cell.value?.toString().toLowerCase().trim() || '';
                if (aliases[v]) score++;
            });
            if (score > bestScore) {
                bestScore = score;
                worksheet = ws;
                headerRowNumber = candidate;
            }
        }

        const firstDataRow = worksheet.getRow(headerRowNumber + 1);
        const allMappings = await getAllColumnMappings();
        const knownCustomers = Array.from(new Set(allMappings.map((m: any) => m.customer)));
        const lowerKnown = knownCustomers.map((c: string) => c.toLowerCase());

        let detectedCustomer = 'DEFAULT';
        firstDataRow.eachCell(cell => {
            const val = cell.value?.toString().trim();
            if (val && lowerKnown.includes(val.toLowerCase())) {
                detectedCustomer = knownCustomers.find((c: string) => c.toLowerCase() === val.toLowerCase()) || 'DEFAULT';
            }
        });

        const colMap = await getColumnMapping(detectedCustomer);
        const normalizedColMap: Record<string, string> = {};
        Object.entries(colMap).forEach(([k, v]) => {
            normalizedColMap[k.toLowerCase()] = v as string;
        });

        const fallbackAliases = this.getFallbackColumnAliases();
        Object.entries(fallbackAliases).forEach(([k, v]) => {
            if (!normalizedColMap[k]) normalizedColMap[k] = v;
        });

        const headerRow = worksheet.getRow(headerRowNumber);
        const headerMap: Record<string, number> = {};
        let inferredSizeCol: number | null = null;

        headerRow.eachCell((cell, colNumber) => {
            const headerText = cell.value?.toString().trim();
            if (!headerText) return;

            const headerKey = headerText.toLowerCase();
            const internalField = normalizedColMap[headerKey];
            const fallbackField = fallbackAliases[headerKey];
            if (internalField && internalField !== 'ignore') {
                if (!headerMap[internalField]) headerMap[internalField] = colNumber;
            } else if (internalField === 'ignore') {
                if (fallbackField === 'transportLocation') {
                    if (!headerMap['transportLocation']) headerMap['transportLocation'] = colNumber;
                }
                return;
            } else {
                if (!headerMap['sizeName'] && inferredSizeCol === null && this.looksLikeSizeHeader(headerText)) {
                    inferredSizeCol = colNumber;
                    return;
                }
                if (!this.shouldSilentlyIgnoreHeader(headerText)) {
                    this.errors.push({
                        field: 'Mapping', row: 1,
                        message: `Unmapped column ignored: ${headerText}`,
                        severity: 'WARNING',
                    });
                }
            }
        });

        if (!headerMap['sizeName'] && inferredSizeCol !== null) {
            headerMap['sizeName'] = inferredSizeCol;
            this.errors.push({ field: 'Mapping', row: 1, message: 'Inferred mapping: sizeName from size-like column.', severity: 'WARNING' });
        }

        const useDefaultSizeBucket = !headerMap['sizeName'];
        if (useDefaultSizeBucket) {
            this.errors.push({ field: 'Mapping', row: 1, message: "No size column detected. Using default 'One Size' for all rows.", severity: 'WARNING' });
        }

        const allowMissingPurchaseOrder = !!manualPurchaseOrder;
        const allowMissingQuantity = !!options?.defaultQuantityIfMissing && !headerMap['quantity'];
        const MANDATORY = ['product'];
        if (!allowMissingPurchaseOrder) MANDATORY.push('purchaseOrder');
        if (!allowMissingQuantity) MANDATORY.push('quantity');
        const missing = MANDATORY.filter(f => !headerMap[f]);

        const isOutputFile = !!(headerMap['lineItem'] || headerMap['productRange']);
        if (missing.length > 0) {
            let msg = `Missing column mappings: ${missing.join(', ')} for customer ${detectedCustomer}.`;
            if (isOutputFile) msg += ' ⚠️ Looks like an NG Output File was uploaded instead of a raw buy file.';
            else msg += ' Please update column_mapping table.';
            this.errors.push({ field: 'File Format', row: 1, message: msg, severity: 'CRITICAL' });

            if (this.runId) {
                await logEvent({ eventName: 'VALIDATION_FAILED', userId: this.userId || 'system', runId: this.runId, metadata: { error_type: 'MISSING_MAPPING', customer: detectedCustomer, missing_fields: missing } });
            }
            return { data: [], errors: this.errors };
        }

        const [factoryMap, mloMap] = await Promise.all([getFactoryMapping(), getMloMapping()]);
        const results: Map<string, ProcessedPO> = new Map();
        const warnedInferredCategory = new Set<string>();
        let skippedMissingSeason = 0;
        let warnedDefaultQty = false;

        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber <= headerRowNumber) return;

            const getRawVal = (field: string) => {
                const col = headerMap[field];
                if (!col) return undefined;
                return this.getCellValue(row.getCell(col));
            };
            const getVal = (field: string) => {
                const raw = getRawVal(field);
                if (raw instanceof Date) return raw.toISOString().split('T')[0];
                return raw?.toString().trim();
            };

            const rawFilePo = getVal('purchaseOrder'); // always reads from file column
            const poNumberRaw = manualPurchaseOrder || rawFilePo;
            if (!poNumberRaw) return;

            const plant = this.stripBrackets(getVal('plant') || '');
            const destCountryRaw = this.stripBrackets(manualDestination || getVal('transportLocation') || '');
            const destCountry = destCountryRaw
                ? (COUNTRY_NAME_MAP[destCountryRaw.toUpperCase()] || destCountryRaw)
                : '';
            const poSuffixParts = [plant, destCountry].filter(Boolean);
            const poNumber = poSuffixParts.length > 0
                ? `${poNumberRaw}-${poSuffixParts.join('-')}`
                : poNumberRaw;

            const brand = this.stripBrackets(getVal('brand') || '');
            const categoryRaw = this.stripBrackets(getVal('category') || '');
            const inferredCat = categoryRaw || this.inferCategoryFromFactoryMap(brand, factoryMap);
            const productExternalRef = this.stripBrackets(getVal('productExternalRef') || '');
            const productCustomerRef = this.stripBrackets(getVal('productCustomerRef') || '');
            const sizeRaw = this.stripBrackets(getVal('sizeName') || '');
            const size = sizeRaw || (useDefaultSizeBucket ? 'One Size' : undefined);
            let qty = parseFloat(getVal('quantity') || '0');
            if (!headerMap['quantity'] && options?.defaultQuantityIfMissing) {
                qty = 1;
                if (!warnedDefaultQty) {
                    warnedDefaultQty = true;
                    this.errors.push({
                        field: 'quantity',
                        row: 1,
                        message: "Quantity column missing. Defaulting Quantity=1 for all rows.",
                        severity: 'WARNING',
                    });
                }
            }
            const colour = this.stripBrackets(getVal('colour') || '').trim();
            if (!colour) {
                this.errors.push({
                    field: 'colour', row: rowNumber,
                    message: `Row ${rowNumber} PO ${poNumber}: colour is empty; line/size skipped.`,
                    severity: 'WARNING',
                });
                return;
            }
            if (colour.trim().toLowerCase() === 'not set') {
                this.errors.push({
                    field: 'colour', row: rowNumber,
                    message: `Row ${rowNumber} PO ${poNumber}: colour is "Not Set"; line/size skipped.`,
                    severity: 'WARNING',
                });
                return;
            }

            const colourKey = this.normalizeColourKey(colour);
            const jdeStyle = this.stripBrackets(getVal('jdeStyle') || '').trim();
            const lookupKey = jdeStyle && colourKey ? `${jdeStyle}|${colourKey}` : '';
            let productMatches = lookupKey ? (productSheetMap[lookupKey] || []) : [];
            // If multiple matches, use the first (exact buyer_style_number matches are inserted first)
            if (productMatches.length > 1) {
                productMatches = [productMatches[0]];
            }
            const hasPlmMap = Object.keys(productSheetMap).length > 0;
            let plmMissing = false;
            if (!jdeStyle && hasPlmMap) {
                this.errors.push({
                    field: 'PLM', row: rowNumber,
                    message: `Row ${rowNumber} PO ${poNumber}: JDE Style missing; PLM fields left blank.`,
                    severity: 'WARNING',
                });
                plmMissing = true;
            }
            if (productMatches.length === 0 && hasPlmMap) {
                this.errors.push({
                    field: 'PLM', row: rowNumber,
                    message: `Row ${rowNumber} PO ${poNumber}: JDE ${jdeStyle} color ${colour} not found in PLM sheet; PLM fields left blank.`,
                    severity: 'WARNING',
                });
                plmMissing = true;
            }
            const productMatch = !plmMissing && productMatches.length === 1 ? productMatches[0] : undefined;
            if (productMatch && productMatch.colour && productMatch.colour.trim().toLowerCase() === 'not set') {
                this.errors.push({
                    field: 'Colour', row: rowNumber,
                    message: `Row ${rowNumber} PO ${poNumber}: PLM Color Name is "Not Set"; line/size skipped.`,
                    severity: 'WARNING',
                });
                return;
            }
            const styleNumber = plmMissing
                ? this.stripBrackets(getVal('product') || getVal('jdeStyle') || '')
                : this.stripBrackets(productMatch?.productName || getVal('product') || getVal('jdeStyle') || '');
            const cost = undefined; // purchase price not captured

            const season = this.stripBrackets(getVal('season') || manualProductRange);
            if (!season) {
                skippedMissingSeason += 1;
                this.errors.push({
                    field: 'season', row: rowNumber,
                    message: `Row ${rowNumber} PO ${poNumber}: No season/range value found.`,
                    severity: 'CRITICAL',
                });
                return;
            }

            const transportLocation = this.normalizeTransportLocation(
                manualDestination || getVal('transportLocation') || ''
            );
            const buyDate = getVal('buyDate');
            const buyRound = this.stripBrackets(getVal('buyRound') || '');
            const exFtyDate = (getRawVal('exFtyDate') || getRawVal('confirmedExFac') || undefined) as Date | string | undefined;
            if (!exFtyDate) {
                this.errors.push({
                    field: 'exFtyDate', row: rowNumber,
                    message: `Row ${rowNumber} PO ${poNumber}: exFtyDate is empty.`,
                    severity: 'WARNING',
                });
            }
            const cancelDate = (getRawVal('cancelDate') || exFtyDate || '') as Date | string;
            const poIssuanceDate = getVal('poIssuanceDate') || buyDate || exFtyDate || '';
            const statusRaw = this.stripBrackets(getVal('status') || 'Confirmed');
            const transportRaw = this.stripBrackets(getVal('transportMethod') || '');
            const templateRaw = this.stripBrackets(getVal('template') || '');
            const vendorCodeRaw = this.stripBrackets(plmMissing
                ? (getVal('productSupplier') || '')
                : (productMatch?.factory || getVal('productSupplier') || ''));
            const vendorNameRaw = this.stripBrackets(getVal('vendorName') || '');

            const buyerPoNumberCell = getRawVal('buyerPoNumber');
            const buyerPoNumber: string | number = (() => {
                if (typeof buyerPoNumberCell === 'number') return buyerPoNumberCell;
                const asText = buyerPoNumberCell?.toString().trim();
                if (asText) return asText;
                const poRaw = getRawVal('purchaseOrder');
                if (typeof poRaw === 'number') return poRaw;
                return rawFilePo || poNumberRaw;
            })();

            const productSupplier = this.resolveSupplier(vendorCodeRaw, vendorNameRaw, brand, inferredCat, factoryMap);
            const customerName = plmMissing
                ? this.resolveCustomer(getVal('customerName'), brand, detectedCustomer, undefined)
                : this.resolveCustomer(productMatch?.customerName || getVal('customerName'), brand, detectedCustomer, undefined);
            const transportMethod = this.normalizeTransportMethod(transportRaw);

            const brandKey = (brand || '').trim().toLowerCase();
            const brandConfig = mloMap.find((m: any) => (m.brand || '').trim().toLowerCase() === brandKey);

            // Resolve separate templates for ORDERS and LINES
            const ordersTemplate = manualTemplate
                || brandConfig?.orders_template?.trim()
                || this.resolveOrdersTemplate(brand, templateRaw);
            const linesTemplate = manualTemplate
                || brandConfig?.lines_template?.trim()
                || this.resolveLinesTemplate(brand, templateRaw);
            const productRange = this.formatProductRange(season);
            const resolvedColour = plmMissing ? colour : (productMatch?.colour || colour);
            const keyDate = manualKeyDate || poIssuanceDate;
            const keyDateFormat: "manual" | "standard" = manualKeyDate ? "manual" : "standard";

            const validStatuses = Array.isArray(brandConfig?.valid_statuses)
                ? brandConfig!.valid_statuses!.map((s: string) => s.toLowerCase())
                : [];

            // Validate transport value
            if (transportMethod && !VALID_TRANSPORT_VALUES.has(transportMethod)) {
                this.errors.push({
                    field: 'TransportMethod', row: rowNumber,
                    message: `Row ${rowNumber} PO ${poNumber}: unmapped transport "${transportMethod}" — expected Sea, Air, or Courier.`,
                    severity: 'WARNING',
                });
            }

            const missingData: string[] = [];
            if (!styleNumber && !plmMissing) missingData.push('Product/Style');
            if (!size) missingData.push('Size');
            if (isNaN(qty)) missingData.push('Quantity');

            if (!categoryRaw && inferredCat && brand && !warnedInferredCategory.has(brand.toLowerCase())) {
                warnedInferredCategory.add(brand.toLowerCase());
                this.errors.push({ field: 'Mapping', row: rowNumber, message: `Category inferred from factory mapping for Brand: ${brand}`, severity: 'WARNING' });
            }

            if (missingData.length > 0) {
                this.errors.push({ field: 'Missing Data', row: rowNumber, message: `PO ${poNumberRaw} missing: ${missingData.join(', ')}.`, severity: 'CRITICAL' });
                return;
            }

            if (validStatuses.length > 0 && statusRaw) {
                const normalizedStatus = statusRaw.toLowerCase();
                if (!validStatuses.includes(normalizedStatus)) {
                    this.errors.push({
                        field: 'Status',
                        row: rowNumber,
                        message: `PO ${poNumberRaw} has status "${statusRaw}" not in valid statuses: ${validStatuses.join(', ')}.`,
                        severity: 'WARNING',
                    });
                }
            }

            // FIX 1: Use resolveKeyUsers() with hardcoded brand fallback
            const mloRow = brandConfig;
            const keyUsers = this.resolveKeyUsers(
                brand,
                manualKeyUser1,
                manualKeyUser2,
                manualKeyUser3,
                manualKeyUser4,
                manualKeyUser5,
                getVal('keyUser1'), getVal('keyUser2'),
                getVal('keyUser4'), getVal('keyUser5'),
                mloRow,
            );

            const customerKey = customerName || detectedCustomer;
            const poKey = `${poNumber}||${customerKey}`;
            if (!results.has(poKey)) {
                results.set(poKey, {
                    header: {
                        purchaseOrder: poNumber,
                        productSupplier,
                        status: statusRaw,
                        customer: customerName,
                        transportMethod,
                        transportLocation,
                        ordersTemplate,
                        linesTemplate,
                        keyDate,
                        keyDateFormat,
                        comments: manualComments || this.buildComments(brand, productRange, buyRound, buyDate, ordersTemplate),
                        currency: 'USD',
                        keyUser1: keyUsers.k1,
                        keyUser2: keyUsers.k2,
                        keyUser3: keyUsers.k3,
                        keyUser4: keyUsers.k4,
                        keyUser5: keyUsers.k5,
                        keyUser6: keyUsers.k6,
                        keyUser7: keyUsers.k7,
                        keyUser8: keyUsers.k8,
                    },
                    lines: [],
                    sizes: {},
                });
            }

            const po = results.get(poKey)!;

            let lineItemNum = 0;
            const rawLineItem = getRawVal('lineItem');
            if (rawLineItem !== undefined && rawLineItem !== null) {
                const maybe = Number(rawLineItem);
                if (Number.isFinite(maybe) && maybe > 0) lineItemNum = Math.round(maybe);
            }
            if (lineItemNum <= 0) {
                lineItemNum = po.lines.length > 0 ? Math.max(...po.lines.map(l => l.lineItem)) + 1 : 1;
            }

            let existingLine = po.lines.find(line => line.lineItem === lineItemNum);
            if (!existingLine) {
                existingLine = {
                    lineItem: lineItemNum,
                    productRange,
                    styleNumber: styleNumber || '',
                    supplierProfile: 'DEFAULT_PROFILE',
                    buyerPoNumber,
                    startDate: (exFtyDate || '') as Date | string,
                    cancelDate: (cancelDate || '') as Date | string,
                    cost,
                    colour: resolvedColour || '',
                    productExternalRef,
                    productCustomerRef,
                };
                po.lines.push(existingLine);
            } else {
                if (styleNumber && existingLine.styleNumber && styleNumber !== existingLine.styleNumber) {
                    this.errors.push({
                        field: 'LineItem', row: rowNumber,
                        message: `PO ${poNumber} line ${lineItemNum} product mismatch: existing ${existingLine.styleNumber}, row ${styleNumber}.`,
                        severity: 'CRITICAL',
                    });
                }
                if (!existingLine.styleNumber && styleNumber) existingLine.styleNumber = styleNumber;
                if (!existingLine.productExternalRef && productExternalRef) {
                    existingLine.productExternalRef = productExternalRef;
                }
                if (!existingLine.productCustomerRef && productCustomerRef) {
                    existingLine.productCustomerRef = productCustomerRef;
                }
                if ((existingLine.cost === undefined || existingLine.cost === '') && cost !== undefined && cost !== '') {
                    existingLine.cost = cost;
                }
            }

            if (!po.sizes[lineItemNum]) po.sizes[lineItemNum] = [];
            po.sizes[lineItemNum].push({ productSize: size || 'One Size', quantity: qty });
            if (qty <= 0) {
                this.errors.push({
                    field: 'Quantity',
                    row: rowNumber,
                    message: `Qty for ${styleNumber} size ${size} is ${qty} (included).`,
                    severity: 'WARNING',
                });
            }
        });

        // Post-process per-PO consistency checks
        for (const [poNumber, po] of results.entries()) {
            po.lines.sort((a, b) => a.lineItem - b.lineItem);
            const lineIds = po.lines.map(l => l.lineItem);
            if (lineIds.length > 0) {
                const minLine = Math.min(...lineIds);
                const maxLine = Math.max(...lineIds);
                if (minLine !== 1) {
                    this.errors.push({ field: 'LineItem', row: 1, message: `PO ${poNumber} starts at LineItem ${minLine} (should start at 1).`, severity: 'WARNING' });
                }
                for (let expected = minLine; expected <= maxLine; expected++) {
                    if (!lineIds.includes(expected)) {
                        this.errors.push({ field: 'LineItem', row: 1, message: `PO ${poNumber} missing LineItem ${expected}.`, severity: 'WARNING' });
                    }
                }
            }

            for (const line of po.lines) {
                const sizesForLine = po.sizes[line.lineItem] || [];
                const totalSizeQty = sizesForLine.reduce((acc, s) => acc + (Number.isFinite(s.quantity) ? s.quantity : 0), 0);
                if (sizesForLine.length === 0) {
                    this.errors.push({ field: 'LineItem', row: 1, message: `PO ${poNumber} line ${line.lineItem} (${line.styleNumber}) has no sizes attached.`, severity: 'WARNING' });
                }
                if (totalSizeQty <= 0) {
                    this.errors.push({ field: 'LineItem', row: 1, message: `PO ${poNumber} line ${line.lineItem} (${line.styleNumber}) has zero total size quantity.`, severity: 'WARNING' });
                }
                const nextLine = po.lines.find(l => l.lineItem === line.lineItem + 1);
                if (nextLine && nextLine.styleNumber && line.styleNumber && line.styleNumber !== nextLine.styleNumber) {
                    const currentHasSizes = (po.sizes[line.lineItem] || []).length > 0;
                    const nextHasSizes = (po.sizes[nextLine.lineItem] || []).length > 0;
                    if (!currentHasSizes && nextHasSizes) {
                        this.errors.push({ field: 'LineItem', row: 1, message: `PO ${poNumber} missing sizes for line ${line.lineItem} while line ${nextLine.lineItem} has sizes: possible row-offset.`, severity: 'WARNING' });
                    }
                }
            }
        }

        const processedData = Array.from(results.values());
        if (skippedMissingSeason > 0) {
            this.errors.push({ field: 'season', row: 1, message: `${skippedMissingSeason} row(s) skipped due to missing season/range.`, severity: 'WARNING' });
        }
        if (processedData.length === 0 && skippedMissingSeason > 0) {
            this.errors.push({ field: 'File Format', row: 1, message: 'No usable rows remain after skipping rows with missing season/range.', severity: 'CRITICAL' });
        }

        const errorCount = this.errors.filter(e => e.severity === 'CRITICAL').length;
        const warningCount = this.errors.filter(e => e.severity === 'WARNING').length;

        if (this.runId) {
            await updateRun(this.runId, {
                status: errorCount > 0 ? 'Validation Failed' : 'Pending Review',
                error_count: errorCount,
                warning_count: warningCount,
                orders_rows: processedData.length,
                lines_rows: processedData.reduce((a, p) => a + p.lines.length, 0),
                order_sizes_rows: processedData.reduce((a, p) =>
                    a + Object.values(p.sizes).reduce((b, s) => b + s.length, 0), 0),
                completed_at: new Date().toISOString(),
            });
            await logEvent({
                eventName: errorCount > 0 ? 'VALIDATION_FAILED' : 'VALIDATION_PASSED',
                userId: this.userId || 'system',
                runId: this.runId,
                metadata: { errorCount, warningCount, customer: detectedCustomer },
            });
        }

        return { data: processedData, errors: this.errors };
    }

    private getCellValue(cell: ExcelJS.Cell) {
        if (cell.isMerged && cell.master) return cell.master.value;
        return cell.value;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Output generation
    // ─────────────────────────────────────────────────────────────────────────

    async generateOutputs(data: ProcessedPO[]) {
        const ordersWb = new ExcelJS.Workbook();
        const linesWb = new ExcelJS.Workbook();
        const sizesWb = new ExcelJS.Workbook();

        const ordersSheet = ordersWb.addWorksheet('ORDERS');
        const linesSheet = linesWb.addWorksheet('LINES');
        const sizesSheet = sizesWb.addWorksheet('ORDER_SIZES');

        ordersSheet.columns = [
            { header: 'PurchaseOrder', key: 'purchaseOrder' },
            { header: 'ProductSupplier', key: 'productSupplier' },
            { header: 'Status', key: 'status' },
            { header: 'Customer', key: 'customer' },
            { header: 'TransportMethod', key: 'transportMethod' },
            { header: 'TransportLocation', key: 'transportLocation' },
            { header: 'PaymentTerm', key: 'paymentTerm' },
            { header: 'Template', key: 'template' },
            { header: 'KeyDate', key: 'keyDate' },
            { header: 'ClosedDate', key: 'closedDate' },
            { header: 'DefaultDeliveryDate', key: 'defaultDeliveryDate' },
            { header: 'Comments', key: 'comments' },
            { header: 'Currency', key: 'currency' },
            { header: 'KeyUser1', key: 'keyUser1' },
            { header: 'KeyUser2', key: 'keyUser2' },
            { header: 'KeyUser3', key: 'keyUser3' },
            { header: 'KeyUser4', key: 'keyUser4' },
            { header: 'KeyUser5', key: 'keyUser5' },
            { header: 'KeyUser6', key: 'keyUser6' },
            { header: 'KeyUser7', key: 'keyUser7' },
            { header: 'KeyUser8', key: 'keyUser8' },
            { header: 'ArchiveDate', key: 'archiveDate' },
            { header: 'PurchaseUOM', key: 'purchaseUOM' },
            { header: 'SellingUOM', key: 'sellingUOM' },
            { header: 'ProductSupplierExt', key: 'productSupplierExt' },
            { header: 'FindField_ProductSupplier', key: 'findField_ProductSupplier' },
        ];

        linesSheet.columns = [
            { header: 'PurchaseOrder', key: 'purchaseOrder' },
            { header: 'LineItem', key: 'lineItem' },
            { header: 'ProductRange', key: 'productRange' },
            { header: 'Product', key: 'product' },
            { header: 'Customer', key: 'customer' },
            { header: 'DeliveryDate', key: 'deliveryDate' },
            { header: 'TransportMethod', key: 'transportMethod' },
            { header: 'TransportLocation', key: 'transportLocation' },
            { header: 'Status', key: 'status' },
            { header: 'PurchasePrice', key: 'purchasePrice' },
            { header: 'SellingPrice', key: 'sellingPrice' },
            { header: 'Template', key: 'template' },
            { header: 'KeyDate', key: 'keyDate' },
            { header: 'SupplierProfile', key: 'supplierProfile' },
            { header: 'ClosedDate', key: 'closedDate' },
            { header: 'Comments', key: 'comments' },
            { header: 'Currency', key: 'currency' },
            { header: 'ArchiveDate', key: 'archiveDate' },
            { header: 'ProductExternalRef', key: 'productExternalRef' },
            { header: 'ProductCustomerRef', key: 'productCustomerRef' },
            { header: 'PurchaseUOM', key: 'purchaseUOM' },
            { header: 'SellingUOM', key: 'sellingUOM' },
            { header: 'UDF-buyer_po_number', key: 'udfBuyerPoNumber' },
            { header: 'UDF-start_date', key: 'udfStartDate' },
            { header: 'UDF-canel_date', key: 'udfCanelDate' },
            { header: 'UDF-Inspection result', key: 'udfInspectionResult' },
            { header: 'UDF-Report Type', key: 'udfReportType' },
            { header: 'UDF-Inspector', key: 'udfInspector' },
            { header: 'UDF-Approval Status', key: 'udfApprovalStatus' },
            { header: 'UDF-Submitted inspection date', key: 'udfSubmittedInspectionDate' },
            { header: 'FindField_Product', key: 'findField_Product' },
        ];

        sizesSheet.columns = [
            { header: 'PurchaseOrder', key: 'purchaseOrder' },
            { header: 'LineItem', key: 'lineItem' },
            { header: 'Range', key: 'range' },
            { header: 'Product', key: 'product' },
            { header: 'SizeName', key: 'sizeName' },
            { header: 'ProductSize', key: 'productSize' },
            { header: 'Quantity', key: 'quantity' },
            { header: 'Colour', key: 'colour' },
            { header: 'Customer', key: 'customer' },
            { header: 'Department', key: 'department' },
            { header: 'CustomAttribute1', key: 'customAttribute1' },
            { header: 'CustomAttribute2', key: 'customAttribute2' },
            { header: 'CustomAttribute3', key: 'customAttribute3' },
            { header: 'LineRatio', key: 'lineRatio' },
            { header: 'ColourExt', key: 'colourExt' },
            { header: 'CustomerExt', key: 'customerExt' },
            { header: 'DepartmentExt', key: 'departmentExt' },
            { header: 'CustomAttribute1Ext', key: 'customAttribute1Ext' },
            { header: 'CustomAttribute2Ext', key: 'customAttribute2Ext' },
            { header: 'CustomAttribute3Ext', key: 'customAttribute3Ext' },
            { header: 'ProductExternalRef', key: 'productExternalRef' },
            { header: 'ProductCustomerRef', key: 'productCustomerRef' },
            { header: 'FindField_Colour', key: 'findField_Colour' },
            { header: 'FindField_Customer', key: 'findField_Customer' },
            { header: 'FindField_Department', key: 'findField_Department' },
            { header: 'FindField_CustomAttribute1', key: 'findField_CustomAttribute1' },
            { header: 'FindField_CustomAttribute2', key: 'findField_CustomAttribute2' },
            { header: 'FindField_CustomAttribute3', key: 'findField_CustomAttribute3' },
            { header: 'FindField_Product', key: 'findField_Product' },
        ];

        if (data && data.length > 0) {
            data.forEach(po => {
                // FIX 2: ORDERS uses ordersTemplate, LINES uses linesTemplate
                ordersSheet.addRow({
                    purchaseOrder: po.header.purchaseOrder,
                    productSupplier: po.header.productSupplier,
                    status: po.header.status,
                    customer: po.header.customer,
                    transportMethod: po.header.transportMethod,
                    transportLocation: po.header.transportLocation,
                    paymentTerm: '',
                    template: po.header.ordersTemplate,
                    keyDate: po.header.keyDateFormat === 'manual'
                        ? this.formatManualDateString(po.header.keyDate)
                        : this.formatDateString(po.header.keyDate),
                    closedDate: '',
                    defaultDeliveryDate: '',
                    comments: po.header.comments,
                    currency: 'USD',
                    keyUser1: po.header.keyUser1,
                    keyUser2: po.header.keyUser2,
                    keyUser3: po.header.keyUser3,
                    keyUser4: po.header.keyUser4,
                    keyUser5: po.header.keyUser5,
                    keyUser6: po.header.keyUser6,
                    keyUser7: po.header.keyUser7,
                    keyUser8: po.header.keyUser8,
                    archiveDate: '',
                    purchaseUOM: '',
                    sellingUOM: '',
                    productSupplierExt: '',
                    findField_ProductSupplier: '',
                });
            });

            data.forEach(po => {
                po.lines.forEach(line => {
                    linesSheet.addRow({
                        purchaseOrder: po.header.purchaseOrder,
                        lineItem: line.lineItem,
                        productRange: line.productRange,
                        product: line.styleNumber,
                        customer: po.header.customer,
                        deliveryDate: this.formatDateString(line.startDate),
                        transportMethod: po.header.transportMethod,
                        transportLocation: po.header.transportLocation,
                        status: po.header.status,
                        purchasePrice: line.cost ?? '',
                        sellingPrice: '',
                        template: po.header.linesTemplate,
                        keyDate: this.formatDateString(line.startDate),
                        supplierProfile: line.supplierProfile,
                        closedDate: '',
                        comments: '',
                        currency: 'USD',
                        archiveDate: '',
                        productExternalRef: line.productExternalRef || '',
                        productCustomerRef: line.productCustomerRef || '',
                        purchaseUOM: '',
                        sellingUOM: '',
                        udfBuyerPoNumber: line.buyerPoNumber,
                        udfStartDate: this.formatDateString(line.startDate),
                        udfCanelDate: this.formatDateString(line.cancelDate),
                        udfInspectionResult: '',
                        udfReportType: '',
                        udfInspector: '',
                        udfApprovalStatus: '',
                        udfSubmittedInspectionDate: '',
                        findField_Product: '',
                    }).commit();
                });
            });

            data.forEach(po => {
                po.lines.forEach(line => {
                    (po.sizes[line.lineItem] || []).forEach(sz => {
                        sizesSheet.addRow({
                            purchaseOrder: po.header.purchaseOrder,
                            lineItem: line.lineItem,
                            range: line.productRange,
                            product: line.styleNumber,
                            sizeName: sz.productSize,
                            productSize: sz.productSize,
                            quantity: sz.quantity,
                            colour: line.colour,
                            customer: '',
                            department: '',
                            customAttribute1: '', customAttribute2: '', customAttribute3: '',
                            lineRatio: '', colourExt: '', customerExt: '', departmentExt: '',
                            customAttribute1Ext: '', customAttribute2Ext: '', customAttribute3Ext: '',
                            productExternalRef: '', productCustomerRef: '',
                            findField_Colour: '', findField_Customer: '', findField_Department: '',
                            findField_CustomAttribute1: '', findField_CustomAttribute2: '',
                            findField_CustomAttribute3: '', findField_Product: '',
                        });
                    });
                });
            });
        }

        if (this.runId && this.userId) {
            await logEvent({ eventName: 'OUTPUT_GENERATED', userId: this.userId, runId: this.runId, metadata: { orders_count: data.length } });
        }

        return {
            orders: await ordersWb.xlsx.writeBuffer(),
            lines: await linesWb.xlsx.writeBuffer(),
            sizes: await sizesWb.xlsx.writeBuffer(),
        };
    }
}
