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
    template: string;
    keyDate: string | Date;
    comments: string;
    currency: string;
    keyUser1: string;
    keyUser2: string;
    keyUser4: string;
    keyUser5: string;
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
            // Standard NG / COL INFOR
            'erp ind', 'brand', 'po #', 'pono', 'purchase order',
            'purchaseorder', 'lineitem', 'productrange', 'company code', 'vendor code',
            // COL INFOR specific
            'material style', 'jde style', 'doc type', 'orig ex fac', 'trans cond',
            'ordered qty', 'buy date', 'color', 'season',
            // Arcteryx / Madison88
            'tracking number', 'article', 'business unit description',
            'requested qty', 'ex-factory', 'transport mode',
            // Generic
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
            // CHANGED: was >= 4, now >= 8
            if (matches >= 8) break;
        }

        return bestRow;
    }

    // ── Comprehensive global alias map ───────────────────────────────────────

    private getFallbackColumnAliases(): Record<string, string> {
        return {
            // ── Transport Location ──────────────────────────────────────────
            'transportlocation': 'transportLocation',
            'transport location': 'transportLocation',
            'destination': 'transportLocation',
            'dest country': 'transportLocation',
            'ult. destination': 'transportLocation',
            // ── PO / Order identifier ────────────────────────────────────────
            'purchaseorder': 'purchaseOrder',
            'po #': 'purchaseOrder',
            'po#': 'purchaseOrder',
            'purchase order': 'purchaseOrder',
            // Arcteryx – Tracking Number is the per-shipment PO key
            'tracking number': 'purchaseOrder',
            'extraction po #': 'buyerPoNumber',
            'extraction po#': 'buyerPoNumber',

            // ── Product / Style ──────────────────────────────────────────────
            'style number': 'product',
            'style no': 'product',
            'material style': 'product',
            'jde style': 'product',
            'style': 'product',
            // Arcteryx – Article is the colourway-level SKU
            'article': 'product',
            // Arcteryx – Model is the base style (fallback)
            'model': 'product',

            // ── Vendor / Supplier ────────────────────────────────────────────
            'vendor code': 'productSupplier',
            'vendorcode': 'productSupplier',
            'vendor': 'productSupplier',
            'product supplier': 'productSupplier',
            'productsupplier': 'productSupplier',
            // Arcteryx – Vendor Name is the full supplier name
            'vendor name': 'vendorName',
            'vendorname': 'vendorName',
            'supplier name': 'vendorName',

            // ── Size ─────────────────────────────────────────────────────────
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

            // ── Quantity ─────────────────────────────────────────────────────
            'ordered qty': 'quantity',
            'open qty (pcs/prs)': 'quantity',
            // Arcteryx
            'requested qty': 'quantity',
            'quantity': 'quantity',
            'qty': 'quantity',

            // ── Date fields ───────────────────────────────────────────────────
            'deliverydate': 'exFtyDate',
            'orig ex fac': 'exFtyDate',
            'negotiated ex fac date': 'exFtyDate',
            // Arcteryx
            'ex-factory': 'exFtyDate',
            'keydate': 'poIssuanceDate',
            'buy date': 'buyDate',
            // Arcteryx – File Date is the closest equivalent to buy date
            'file date': 'buyDate',
            'cancel date': 'cancelDate',
            'canceldate': 'cancelDate',

            // ── Transport ─────────────────────────────────────────────────────
            'transportmethod': 'transportMethod',
            'trans cond': 'transportMethod',
            // Arcteryx
            'transport mode': 'transportMethod',

            // ── Template / Doc Type ───────────────────────────────────────────
            'doc type': 'template',
            'template': 'template',

            // ── Season / Range ────────────────────────────────────────────────
            'range': 'season',
            'productrange': 'season',
            'season': 'season',

            // ── Brand / Customer ──────────────────────────────────────────────
            'brand': 'brand',
            // Arcteryx
            'business unit description': 'brand',
            'customer': 'customerName',
            'customer name': 'customerName',

            // ── Status ────────────────────────────────────────────────────────
            'status': 'status',
            'confirmation status': 'status',
            // Arcteryx
            'gsc type': 'status',

            // ── Colour ────────────────────────────────────────────────────────
            'colour': 'colour',
            'color': 'colour',
            // Arcteryx – Article Name is the colour description
            'article name': 'colour',

            // ── Buy round (Arcteryx: "Submit Buy") ────────────────────────────
            'submit buy': 'buyRound',
            'buy round': 'buyRound',

            // ── Standard NG output headers (ignored on re-upload) ─────────────
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

    // ── Customer name resolution with brand fallback ──────────────────────────

    private resolveCustomer(customerRaw: string | undefined, brand: string | undefined, detectedCustomer: string, mappedCustomerName: string | undefined): string {
        if (mappedCustomerName?.trim()) return mappedCustomerName.trim();

        const raw = (customerRaw || '').trim();
        if (raw) {
            const key = raw.toLowerCase();
            return BRAND_CUSTOMER_MAP[key] || raw;
        }

        if (brand) {
            const key = brand.trim().toLowerCase();
            if (BRAND_CUSTOMER_MAP[key]) return BRAND_CUSTOMER_MAP[key];
        }

        if (detectedCustomer && detectedCustomer !== 'DEFAULT') return detectedCustomer;
        return brand?.toUpperCase() || 'COL';
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
            // Standard NG output columns (safe to ignore on re-upload)
            'lineitem', 'purchaseprice', 'sellingprice',
            'supplierprofile', 'closeddate', 'comments', 'currency', 'archivedate',
            'productexternalref', 'productcustomerref', 'purchaseuom', 'sellinguom',
            'paymentterm', 'defaultdeliverydate', 'productsupplierext',
            'keyuser1', 'keyuser2', 'keyuser3', 'keyuser4', 'keyuser5',
            'keyuser6', 'keyuser7', 'keyuser8',
            'department', 'customattribute1', 'customattribute2', 'customattribute3',
            'lineratio', 'colourext', 'customerext', 'departmentext',
            'customattribute1ext', 'customattribute2ext', 'customattribute3ext',
            // Arcteryx-specific columns not needed for NG output
            'file date', 'sku', 'sku description', 'model description', 'gsc type', 'product group description',
            'product line description', 'planning category',
            'transit vendor destination', 'official buy', 'plant',
            'storage location', 'stock segment',
            // COL INFOR metadata
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
        // Handles F26, FW26, S26, SS26
        const m = normalized.match(/^([FS])(?:W|S)?(\d{2})$/i);
        if (m) return `${m[1].toUpperCase()}H:20${m[2]}`;
        if (normalized) return normalized;
        return 'FH:2026';
    }

    private normalizeTemplate(rawTemplate: string): string {
        const normalized = (rawTemplate || '').trim().toUpperCase();
        const map: Record<string, string> = {
            OG: 'BULK', ZNB1: 'BULK', ZMF1: 'BULK', ZDS1: 'BULK',
        };
        return map[normalized] || (rawTemplate || 'BULK').trim() || 'BULK';
    }

    private normalizeTransportMethod(raw: string | undefined): string {
        const key = (raw || '').trim().toLowerCase();
        return TRANSPORT_MAP[key] || (raw ? raw.trim() : 'Sea');
    }

    private parseDate(raw: string | Date | undefined): Date | null {
        if (!raw) return null;
        if (raw instanceof Date) return isNaN(raw.getTime()) ? null : raw;
        const text = String(raw).trim();
        if (!text) return null;

        // YYYY-MM-DD (optionally with time)
        const isoMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (isoMatch) {
            const year = Number(isoMatch[1]);
            const month = Number(isoMatch[2]);
            const day = Number(isoMatch[3]);
            const date = new Date(year, month - 1, day);
            return isNaN(date.getTime()) ? null : date;
        }

        // MM/DD/YYYY or MM-DD-YYYY (aligns with Python formats)
        const usMatch = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
        if (usMatch) {
            const month = Number(usMatch[1]);
            const day = Number(usMatch[2]);
            const year = Number(usMatch[3]);
            const date = new Date(year, month - 1, day);
            return isNaN(date.getTime()) ? null : date;
        }

        // DD-Mon-YYYY or DD-Month-YYYY
        const monMatch = text.match(/^(\d{1,2})-([A-Za-z]+)-(\d{4})$/);
        if (monMatch) {
            const day = Number(monMatch[1]);
            const monText = monMatch[2].toLowerCase();
            const year = Number(monMatch[3]);
            const months = [
                'jan', 'feb', 'mar', 'apr', 'may', 'jun',
                'jul', 'aug', 'sep', 'oct', 'nov', 'dec',
            ];
            const monthIndex = months.findIndex(m => monText.startsWith(m));
            if (monthIndex >= 0) {
                const date = new Date(year, monthIndex, day);
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
        const yyyy = date.getFullYear();
        return `${mm}/${dd}/${yyyy}`;
    }

    private toExcelDate(raw: string | Date | undefined): Date | '' {
        const date = this.parseDate(raw);
        return date || '';
    }

    private buildComments(brand: string | undefined, season: string, buyRound: string, buyDateRaw: string | undefined, template: string): string {
        const b = brand || 'OUTPUT';
        const s = season || 'NOS';

        // If the file has an explicit buy round label (e.g. "Buy 2"), use it directly
        if (buyRound) return `[${b}] ${s} ${buyRound} ${template}`;

        const parsed = this.parseDate(buyDateRaw);
        if (parsed) {
            const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
            const monShort = months[parsed.getMonth()];
            const day = String(parsed.getDate()).padStart(2, '0');
            const suffix = template ? ` ${template}` : '';
            return `[${b}] ${s} ${monShort} Buy ${day}-${monShort.toUpperCase()}${suffix}`;
        }

        return `[${b}] ${s}`;
    }

    private getMloFallbackByBrand(brand: string | undefined): { k1: string; k2: string; k4: string; k5: string } {
        return { k1: '', k2: '', k4: '', k5: '' };
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Core processing method
    // ─────────────────────────────────────────────────────────────────────────

    async processBuyFile(buffer: any): Promise<{ data: ProcessedPO[]; errors: ValidationError[] }> {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);

        // Pick the sheet with the best header score (same logic as Python)
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

        // 1. Detect Customer from first data row (matches DB-registered customer names)
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

        // 2. Fetch customer-specific column map from DB
        const colMap = await getColumnMapping(detectedCustomer);
        const normalizedColMap: Record<string, string> = {};
        Object.entries(colMap).forEach(([k, v]) => {
            normalizedColMap[k.toLowerCase()] = v as string;
        });

        // Merge fallback aliases (DB takes priority)
        const fallbackAliases = this.getFallbackColumnAliases();
        Object.entries(fallbackAliases).forEach(([k, v]) => {
            if (!normalizedColMap[k]) normalizedColMap[k] = v;
        });

        // 3. Build header → column index map
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
                // NOTE: DEFAULT column_mapping currently marks TransportLocation as "ignore".
                // If/when that DB entry is corrected to "transportLocation", remove this override.
                // Allow transportLocation fallback even if DB mapping marks ignore.
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

        // 4. Validate mandatory fields
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
        const warnedMissingCategory = new Set<string>();
        const warnedMissingFactory = new Set<string>();
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
            // Fix #1: Remove default 'NOS', log CRITICAL and skip row if missing
            const season = getVal('season');
            if (!season) {
                skippedMissingSeason += 1;
                this.errors.push({
                    field: 'season',
                    row: rowNumber,
                    message: `Row ${rowNumber} PO ${poNumber}: No season/range value found.`,
                    severity: 'CRITICAL',
                });
                return;
            }
            // Fix #2: TransportLocation mapping
            const transportLocation = getVal('transportLocation') || '';
            const buyDate = getVal('buyDate');
            const buyRound = getVal('buyRound') || '';
            const exFtyDate = getVal('exFtyDate') || undefined;
            if (!exFtyDate) {
                this.errors.push({
                    field: 'exFtyDate',
                    row: rowNumber,
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
            const template = this.normalizeTemplate(templateRaw);

            // Validation
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

            const providedKeyUsers = {
                k1: getVal('keyUser1'), k2: getVal('keyUser2'),
                k4: getVal('keyUser4'), k5: getVal('keyUser5'),
            };
            const mlo = (providedKeyUsers.k1 || providedKeyUsers.k2) ? null :
                mloMap.find((m: any) => m.brand?.toLowerCase() === brand?.toLowerCase());
            const mloFallback = this.getMloFallbackByBrand(brand);
            const keyUsers = {
                k1: providedKeyUsers.k1 || mlo?.keyuser1 || mloFallback.k1,
                k2: providedKeyUsers.k2 || mlo?.keyuser2 || mloFallback.k2,
                k4: providedKeyUsers.k4 || mlo?.keyuser4 || mloFallback.k4,
                k5: providedKeyUsers.k5 || mlo?.keyuser5 || mloFallback.k5,
            };

            if (!results.has(poNumber)) {
                results.set(poNumber, {
                    header: {
                        purchaseOrder: poNumber,
                        productSupplier,
                        status: statusRaw,
                        customer: customerName,
                        transportMethod,
                        transportLocation,
                        template,
                        keyDate: poIssuanceDate,
                        comments: this.buildComments(brand, season, buyRound, buyDate, template),
                        currency: 'USD',
                        keyUser1: keyUsers.k1 || '',
                        keyUser2: keyUsers.k2 || '',
                        keyUser4: keyUsers.k4 || '',
                        keyUser5: keyUsers.k5 || '',
                    },
                    lines: [],
                    sizes: {},
                });
            }

            const po = results.get(poNumber)!;
            // Determine line item index, allowing explicit line item values for ORDER_SIZES format.
            let lineItemNum = 0;
            const rawLineItem = getRawVal('lineItem');
            if (rawLineItem !== undefined && rawLineItem !== null) {
                const maybe = Number(rawLineItem);
                if (Number.isFinite(maybe) && maybe > 0) {
                    lineItemNum = Math.round(maybe);
                }
            }
            if (lineItemNum <= 0) {
                // fallback sequential line numbering
                lineItemNum = po.lines.length > 0 ? Math.max(...po.lines.map(line => line.lineItem)) + 1 : 1;
            }

            let existingLine = po.lines.find(line => line.lineItem === lineItemNum);
            if (!existingLine) {
                existingLine = {
                    lineItem: lineItemNum,
                    productRange: this.formatProductRange(season),
                    styleNumber: styleNumber || '',
                    supplierProfile: 'DEFAULT_PROFILE',
                    buyerPoNumber: buyerPoNumber,
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

                // Keep line style aligned to the first non-empty style seen for the line.
                if (!existingLine.styleNumber && styleNumber) {
                    existingLine.styleNumber = styleNumber;
                }
            }

            if (qty > 0) {
                if (!po.sizes[lineItemNum]) po.sizes[lineItemNum] = [];
                po.sizes[lineItemNum].push({ productSize: size || 'One Size', quantity: qty });
            } else {
                this.errors.push({ field: 'Quantity', row: rowNumber, message: `Qty for ${styleNumber} size ${size} is ${qty} (excluded).`, severity: 'WARNING' });
            }
        });

        // Post-process per-PO consistency checks (line gaps, size totals, product mismatches)
        for (const [poNumber, po] of results.entries()) {
            po.lines.sort((a, b) => a.lineItem - b.lineItem);
            const lineIds = po.lines.map(line => line.lineItem);
            if (lineIds.length > 0) {
                const minLine = Math.min(...lineIds);
                const maxLine = Math.max(...lineIds);
                if (minLine !== 1) {
                    this.errors.push({ field: 'LineItem', row: 1, message: `PO ${poNumber} starts at LineItem ${minLine} (should start at 1).`, severity: 'WARNING' });
                }
                for (let expected = minLine; expected <= maxLine; expected++) {
                    if (!lineIds.includes(expected)) {
                        this.errors.push({ field: 'LineItem', row: 1, message: `PO ${poNumber} missing LineItem ${expected}; possible row offset in sizes ordering.`, severity: 'WARNING' });
                    }
                }
            }

            const lineStyleToRows: Record<number, Set<string>> = {};
            for (const line of po.lines) {
                if (!lineStyleToRows[line.lineItem]) lineStyleToRows[line.lineItem] = new Set();
                if (line.styleNumber) lineStyleToRows[line.lineItem].add(line.styleNumber);
            }

            for (const [lineItem, styleSet] of Object.entries(lineStyleToRows)) {
                if (styleSet.size > 1) {
                    const styleList = Array.from(styleSet).join(', ');
                    this.errors.push({ field: 'LineItem', row: 1, message: `PO ${poNumber} line ${lineItem} has conflicting products across line records: ${styleList}.`, severity: 'CRITICAL' });
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

                // Heuristic detection for possible 1-line offset in size ordering.
                const nextLine = po.lines.find(l => l.lineItem === line.lineItem + 1);
                if (nextLine && nextLine.styleNumber && line.styleNumber && line.styleNumber !== nextLine.styleNumber) {
                    const currentHasSizes = (po.sizes[line.lineItem] || []).length > 0;
                    const nextHasSizes = (po.sizes[nextLine.lineItem] || []).length > 0;
                    if (!currentHasSizes && nextHasSizes) {
                        this.errors.push({ field: 'LineItem', row: 1, message: `PO ${poNumber} missing sizes for line ${line.lineItem} while line ${nextLine.lineItem} has sizes: possible row-offset in order of SIZES lines.`, severity: 'WARNING' });
                    }
                    if (currentHasSizes && !nextHasSizes) {
                        this.errors.push({ field: 'LineItem', row: 1, message: `PO ${poNumber} line ${line.lineItem} has sizes but next line ${nextLine.lineItem} has none: possible SIZES ordering issue.`, severity: 'WARNING' });
                    }
                }
            }
        }

        const processedData = Array.from(results.values());
        if (skippedMissingSeason > 0) {
            this.errors.push({
                field: 'season',
                row: 1,
                message: `${skippedMissingSeason} row(s) skipped due to missing season/range.`,
                severity: 'WARNING',
            });
        }
        if (processedData.length === 0 && skippedMissingSeason > 0) {
            this.errors.push({
                field: 'File Format',
                row: 1,
                message: 'No usable rows remain after skipping rows with missing season/range.',
                severity: 'CRITICAL',
            });
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
    // Output generation (unchanged schema — exact column order preserved)
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

        linesSheet.columns = [
            { header: 'PurchaseOrder',                    key: 'purchaseOrder' },
            { header: 'LineItem',                         key: 'lineItem' },
            { header: 'ProductRange',                     key: 'productRange' },
            { header: 'Product',                          key: 'product' },
            { header: 'Customer',                         key: 'customer' },
            { header: 'DeliveryDate',                     key: 'deliveryDate' },
            { header: 'TransportMethod',                  key: 'transportMethod' },
            { header: 'TransportLocation',                key: 'transportLocation' },
            { header: 'Status',                           key: 'status' },
            { header: 'PurchasePrice',                    key: 'purchasePrice' },
            { header: 'SellingPrice',                     key: 'sellingPrice' },
            { header: 'Template',                         key: 'template' },
            { header: 'KeyDate',                          key: 'keyDate' },
            { header: 'SupplierProfile',                  key: 'supplierProfile' },
            { header: 'ClosedDate',                       key: 'closedDate' },
            { header: 'Comments',                         key: 'comments' },
            { header: 'Currency',                         key: 'currency' },
            { header: 'ArchiveDate',                      key: 'archiveDate' },
            { header: 'ProductExternalRef',               key: 'productExternalRef' },
            { header: 'ProductCustomerRef',               key: 'productCustomerRef' },
            { header: 'PurchaseUOM',                      key: 'purchaseUOM' },
            { header: 'SellingUOM',                       key: 'sellingUOM' },
            { header: 'UDF-buyer_po_number',              key: 'udfBuyerPoNumber' },
            { header: 'UDF-start_date',                   key: 'udfStartDate' },
            { header: 'UDF-canel_date',                   key: 'udfCanelDate' },
            { header: 'UDF-Inspection result',            key: 'udfInspectionResult' },
            { header: 'UDF-Report Type',                  key: 'udfReportType' },
            { header: 'UDF-Inspector',                    key: 'udfInspector' },
            { header: 'UDF-Approval Status',              key: 'udfApprovalStatus' },
            { header: 'UDF-Submitted inspection date',    key: 'udfSubmittedInspectionDate' },
            { header: 'FindField_Product',                key: 'findField_Product' },
        ];

        if (data && data.length > 0) {
            data.forEach(po => {
                ordersSheet.addRow({
                    purchaseOrder: po.header.purchaseOrder,
                    productSupplier: po.header.productSupplier,
                    status: po.header.status,
                    customer: po.header.customer,
                    transportMethod: po.header.transportMethod,
                    transportLocation: po.header.transportLocation,
                    paymentTerm: '',
                    template: po.header.template,
                    keyDate: this.formatDateString(po.header.keyDate), 
                    closedDate: '',
                    defaultDeliveryDate: '',
                    comments: po.header.comments,
                    currency: 'USD',
                    keyUser1: po.header.keyUser1,
                    keyUser2: po.header.keyUser2,
                    keyUser3: '',
                    keyUser4: po.header.keyUser4,
                    keyUser5: po.header.keyUser5,
                    keyUser6: '', keyUser7: '', keyUser8: '',
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
                        template: po.header.template,
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
                            customer: '', department: '',
                            customAttribute1: '', customAttribute2: '', customAttribute3: '',
                            lineRatio: '', colourExt: '', customerExt: '',
                            departmentExt: '',
                            customAttribute1Ext: '', customAttribute2Ext: '', customAttribute3Ext: '',
                            productExternalRef: '', productCustomerRef: '',
                            findField_Colour: '', findField_Customer: '',
                            findField_Department: '',
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
