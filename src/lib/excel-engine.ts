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
    startDate: string;
    cancelDate: string;
    colour: string;
}

export interface POSize {
    productSize: string;
    quantity: number;
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
            'jde style': 'product',
            'style': 'product',
            'product': 'product',
            'sku': 'product',
            'item': 'product',
            'article': 'product',
            'model': 'product',
            'vendor code': 'productSupplier',
            'vendorcode': 'productSupplier',
            'vendor': 'productSupplier',
            'supplier': 'productSupplier',
            'product supplier': 'productSupplier',
            'productsupplier': 'productSupplier',
            'vendor name': 'vendorName',
            'vendorname': 'vendorName',
            'supplier name': 'vendorName',
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
            'status': 'status',
            'confirmation status': 'status',
            'gsc type': 'status',
            'colour': 'colour',
            'color': 'colour',
            'article name': 'colour',
            'color description': 'colour',
            'submit buy': 'buyRound',
            'buy round': 'buyRound',
            'udf-buyer_po_number': 'buyerPoNumber',
            'udf-start_date': 'exFtyDate',
            'udf-canel_date': 'cancelDate',
        };
    }

    // ── Supplier resolution with brand fallback ───────────────────────────────

    private resolveSupplier(vendorCode: string | undefined, vendorName: string | undefined, brand: string | undefined): string {
        if (vendorCode && vendorCode.trim().length > 2) return vendorCode.trim();
        if (vendorName && vendorName.trim().length > 2) return vendorName.trim();
        const key = (brand || '').trim().toLowerCase();
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
        const raw = (customerRaw || '').trim();

        // 1. Check subtype map first (handles RTO / SMU / In-Line variants)
        if (raw) {
            const key = raw.toLowerCase();
            if (TNF_CUSTOMER_SUBTYPE_MAP[key]) return TNF_CUSTOMER_SUBTYPE_MAP[key];
        }

        // 2. DB-mapped customer name — only use if it looks like a subtype-aware
        //    value (i.e. contains RTO / SMU / In-Line) OR if raw is empty.
        //    This prevents a base-brand DB entry ("The North Face") from
        //    overwriting a correctly detected subtype from the buy file.
        if (mappedCustomerName?.trim()) {
            const mappedKey = mappedCustomerName.trim().toLowerCase();
            const hasSubtype = mappedKey.includes('rto') || mappedKey.includes('smu') || mappedKey.includes('in-line') || mappedKey.includes('inline');
            if (!raw || hasSubtype) return mappedCustomerName.trim();
        }

        // 3. BRAND_CUSTOMER_MAP from raw value
        if (raw) {
            const key = raw.toLowerCase();
            if (BRAND_CUSTOMER_MAP[key]) return BRAND_CUSTOMER_MAP[key];
            return raw;
        }

        // 4. Brand fallback
        if (brand) {
            const key = brand.trim().toLowerCase();
            if (BRAND_CUSTOMER_MAP[key]) return BRAND_CUSTOMER_MAP[key];
        }

        // 5. Detected customer from DB registration
        if (detectedCustomer && detectedCustomer !== 'DEFAULT') return detectedCustomer;

        return brand?.toUpperCase() || 'COL';
    }

    // FIX 1: KeyUser resolution with hardcoded brand fallback ─────────────────
    private resolveKeyUsers(
        brand: string | undefined,
        providedK1: string | undefined,
        providedK2: string | undefined,
        providedK4: string | undefined,
        providedK5: string | undefined,
        mloRow: any,
    ): KeyUsers {
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
            'transit vendor destination', 'official buy', 'plant',
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
        const normalized = (season || '').trim();
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

    private buildComments(brand: string | undefined, season: string, buyRound: string, buyDateRaw: string | undefined, template: string): string {
        const b = brand || 'OUTPUT';
        const s = season || 'NOS';
        if (buyRound) return `[${b}] ${s} ${buyRound} ${template}`;
        const parsed = this.parseDate(buyDateRaw);
        if (parsed) {
            const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
            const monShort = months[parsed.getMonth()];
            const day = String(parsed.getDate()).padStart(2, '0');
            const suffix = template ? ` ${template}` : '';
            return `[${b}] ${s} ${monShort} Buy ${day}-${monShort.toUpperCase()}${suffix}`;
        }
        return `[${b}] ${s}`;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Core processing method
    // ─────────────────────────────────────────────────────────────────────────

    async processBuyFile(buffer: any): Promise<{ data: ProcessedPO[]; errors: ValidationError[] }> {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);

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

        const MANDATORY = ['purchaseOrder', 'product', 'quantity'];
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

        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber <= headerRowNumber) return;

            const getRawVal = (field: string) => {
                const col = headerMap[field];
                if (!col) return undefined;
                return this.getCellValue(row.getCell(col));
            };
            const getVal = (field: string) => getRawVal(field)?.toString().trim();

            const poNumberRaw = getVal('purchaseOrder');
            if (!poNumberRaw) return;

            const dcCode = getVal('dcCode');
            const poNumber = poNumberRaw + (dcCode ? `-${dcCode}` : '');

            const brand = getVal('brand');
            const categoryRaw = getVal('category');
            const inferredCat = categoryRaw || this.inferCategoryFromFactoryMap(brand, factoryMap);
            const styleNumber = getVal('product');
            const size = getVal('sizeName') || (useDefaultSizeBucket ? 'One Size' : undefined);
            const qty = parseFloat(getVal('quantity') || '0');
            const colour = (getVal('colour') || '').trim();

            const season = getVal('season');
            if (!season) {
                skippedMissingSeason += 1;
                this.errors.push({
                    field: 'season', row: rowNumber,
                    message: `Row ${rowNumber} PO ${poNumber}: No season/range value found.`,
                    severity: 'CRITICAL',
                });
                return;
            }

            const transportLocation = getVal('transportLocation') || '';
            const buyDate = getVal('buyDate');
            const buyRound = getVal('buyRound') || '';
            const exFtyDate = getVal('exFtyDate') || undefined;
            if (!exFtyDate) {
                this.errors.push({
                    field: 'exFtyDate', row: rowNumber,
                    message: `Row ${rowNumber} PO ${poNumber}: exFtyDate is empty.`,
                    severity: 'WARNING',
                });
            }
            const cancelDate = getVal('cancelDate') || exFtyDate || '';
            const poIssuanceDate = getVal('poIssuanceDate') || buyDate || exFtyDate || '';
            const statusRaw = getVal('status') || 'Confirmed';
            const transportRaw = getVal('transportMethod');
            const templateRaw = getVal('template') || '';
            const vendorCodeRaw = getVal('productSupplier');
            const vendorNameRaw = getVal('vendorName');

            const buyerPoNumberCell = getRawVal('buyerPoNumber');
            const buyerPoNumber: string | number = (() => {
                if (typeof buyerPoNumberCell === 'number') return buyerPoNumberCell;
                const asText = buyerPoNumberCell?.toString().trim();
                if (asText) return asText;
                const poRaw = getRawVal('purchaseOrder');
                if (typeof poRaw === 'number') return poRaw;
                return poNumberRaw;
            })();

            const productSupplier = this.resolveSupplier(vendorCodeRaw, vendorNameRaw, brand);
            const customerName = this.resolveCustomer(getVal('customerName'), brand, detectedCustomer, undefined);
            const transportMethod = this.normalizeTransportMethod(transportRaw);

            const brandKey = (brand || '').trim().toLowerCase();
            const brandConfig = mloMap.find((m: any) => (m.brand || '').trim().toLowerCase() === brandKey);

            // Resolve separate templates for ORDERS and LINES
            const ordersTemplate = brandConfig?.orders_template?.trim()
                || this.resolveOrdersTemplate(brand, templateRaw);
            const linesTemplate = brandConfig?.lines_template?.trim()
                || this.resolveLinesTemplate(brand, templateRaw);

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
            if (!styleNumber) missingData.push('Product/Style');
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
                getVal('keyUser1'), getVal('keyUser2'),
                getVal('keyUser4'), getVal('keyUser5'),
                mloRow,
            );

            if (!results.has(poNumber)) {
                results.set(poNumber, {
                    header: {
                        purchaseOrder: poNumber,
                        productSupplier,
                        status: statusRaw,
                        customer: customerName,
                        transportMethod,
                        transportLocation,
                        ordersTemplate,
                        linesTemplate,
                        keyDate: poIssuanceDate,
                        comments: this.buildComments(brand, season, buyRound, buyDate, ordersTemplate),
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

            const po = results.get(poNumber)!;

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
                    productRange: this.formatProductRange(season),
                    styleNumber: styleNumber || '',
                    supplierProfile: 'DEFAULT_PROFILE',
                    buyerPoNumber,
                    startDate: exFtyDate || '',
                    cancelDate: cancelDate || '',
                    colour: colour || '',
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
            }

            if (qty > 0) {
                if (!po.sizes[lineItemNum]) po.sizes[lineItemNum] = [];
                po.sizes[lineItemNum].push({ productSize: size || 'One Size', quantity: qty });
            } else {
                this.errors.push({ field: 'Quantity', row: rowNumber, message: `Qty for ${styleNumber} size ${size} is ${qty} (excluded).`, severity: 'WARNING' });
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
                    keyDate: this.formatDateString(po.header.keyDate),
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
                        purchasePrice: '',
                        sellingPrice: '',
                        template: po.header.linesTemplate,
                        keyDate: this.formatDateString(line.startDate),
                        supplierProfile: line.supplierProfile,
                        closedDate: '',
                        comments: '',
                        currency: 'USD',
                        archiveDate: '',
                        productExternalRef: '',
                        productCustomerRef: '',
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
                        if (sz.quantity <= 0) return;
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
