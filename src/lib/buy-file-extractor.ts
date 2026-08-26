import { readExcelFile } from '@/lib/excel/excel-reader';
import { detectHeaderRow } from '@/lib/ai/header-detector';
import { mapHeaders } from '@/lib/ai/header-mapper';
import { findMatchingTemplateSupabase } from '@/lib/templates/supabase-store';
import { getColumnMapping } from '@/lib/data-loader';
import { NextGenCachedClient } from '@/lib/nextgen/client';
import { ExcelEngine } from '@/lib/excel-engine';
import type { ProductSheetRow } from '@/lib/excel-engine';
import { mergeBuyFileWithNextGen } from '@/lib/merge/merge-buy-nextgen';
import { BuyFileItem, ColumnMapping, NextGenStyleInfo, ProductData } from '@/lib/types/buy-file';
import { lookupBrand, getAllBrandAliases } from '@/lib/brand-config';

const INTERNAL_TO_CANONICAL: Record<string, keyof ColumnMapping> = {
    purchaseOrder: 'po_number',
    buyerPoNumber: 'po_number',
    product: 'buyer_style_number',
    productCustomerRef: 'buyer_style_number',
    styleNumber: 'buyer_style_number',
    buyerStyleNumber: 'buyer_style_number',
    productExternalRef: 'sku',
    material: 'sku',
    productDescription: 'product_description',
    shortText: 'product_description',
    longText: 'product_description',
    colour: 'color',
    colorName: 'color',
    colourName: 'color',
    styleColor: 'color_code',
    colorCode: 'color_code',
    sizeName: 'size',
    productSize: 'size',
    gridValue: 'size',
    quantity: 'quantity',
    orderedQty: 'quantity',
    scheduledQuantity: 'quantity',
    exFtyDate: 'delivery_date',
    deliveryDate: 'delivery_date',
    confirmedExFac: 'delivery_date',
    vendorConfirmedETD: 'delivery_date',
    requestedDeliveryDate: 'delivery_date',
    finalXfDate: 'delivery_date',
    season: 'season',
    seasonCode: 'season',
    customerName: 'customer',
    customer: 'customer',
    soldTo: 'customer',
    brand: 'customer',
    vendorName: 'factory',
    factory: 'factory',
    supplierName: 'factory',
    finalVendorName: 'factory',
    finalFactoryName: 'factory',
    plant: 'factory',
    currency: 'currency',
    finalCurrency: 'currency',
    purchasePrice: 'unit_cost',
    confirmedUnitPrice: 'unit_cost',
    unitPrice: 'unit_cost',
    fob: 'unit_cost',
    netValue: 'unit_cost',
    sellingPrice: 'unit_cost',
    productionUpchargesUsd: 'unit_cost',
    materialUpchargesUsd: 'unit_cost',
    buyInformation: 'buy_information',
    buyInfo: 'buy_information',
};

function convertLegacyMapping(legacy: Record<string, string>): ColumnMapping {
    const mapping: ColumnMapping = {};
    for (const [buyFileColumn, internalField] of Object.entries(legacy)) {
        const canonical = INTERNAL_TO_CANONICAL[internalField];
        if (canonical && !(mapping as Record<string, string>)[canonical]) {
            (mapping as Record<string, string>)[canonical] = buyFileColumn;
        }
    }
    return mapping;
}

async function loadLegacyMapping(customer?: string): Promise<ColumnMapping> {
    try {
        const defaultMapping = convertLegacyMapping(await getColumnMapping('DEFAULT'));
        if (!customer || customer === 'DEFAULT') return defaultMapping;
        const customerMapping = convertLegacyMapping(await getColumnMapping(customer));
        return { ...defaultMapping, ...customerMapping };
    } catch (err) {
        console.warn('[buy-file-extractor] failed to load legacy mapping:', err);
        return {};
    }
}

async function buildProductSheetMap(buffers: ArrayBuffer[]): Promise<Record<string, ProductSheetRow[]>> {
    if (!buffers.length) return {};
    const engine = new ExcelEngine();
    const merged: Record<string, ProductSheetRow[]> = {};
    for (const buffer of buffers) {
        const map = await engine.extractProductSheetMap(buffer);
        for (const [key, rows] of Object.entries(map)) {
            if (!merged[key]) merged[key] = [];
            merged[key].push(...rows);
        }
    }
    return merged;
}

function enrichItemsWithProductSheet(
    items: BuyFileItem[],
    productSheetMap: Record<string, ProductSheetRow[]>
): BuyFileItem[] {
    if (!Object.keys(productSheetMap).length) return items;

    const engine = new ExcelEngine();
    return items.map((item) => {
        // If NextGen already matched this item, only fill gaps from product sheet
        const alreadyMatched = item.matchStatus === 'matched' || item.matchStatus === 'ambiguous';

        const styleRaw = engine.stripBrackets(item.style || '').trim();
        const colorRaw = item.colorCode || item.color || '';
        const colorKey = engine.normalizeColourKey(colorRaw);
        const styleCandidates = normalizeStyleCandidates(styleRaw);
        const exactMatches = styleCandidates.flatMap((style) => productSheetMap[`${style}|${colorKey}`] || []);
        const candidates = deduplicateProductRows(exactMatches);

        if (!candidates.length) {
            // Don't downgrade a NextGen match to unmatched just because product sheet has no data
            if (alreadyMatched) return item;
            return {
                ...item,
                matchStatus: 'unmatched',
                matchScore: 0,
                matchReason: `No product export match for style ${styleRaw || '(blank)'} and color ${colorKey || '(blank)'}`,
            };
        }

        const ranked = candidates
            .map((candidate) => scoreProductMatch(item, candidate))
            .sort((a, b) => b.score - a.score);
        const bestResult = ranked[0];
        const runnerUp = ranked[1];
        const ambiguous = Boolean(runnerUp && runnerUp.score === bestResult.score && productIdentity(runnerUp.row) !== productIdentity(bestResult.row));
        const best = bestResult.row;

        // Only fill fields that are still blank (don't overwrite NextGen data)
        return {
            ...item,
            product: item.product || best.productName || null,
            productExternalRef: item.productExternalRef || best.productExternalRef || null,
            costingReference: item.costingReference || best.costingReference || null,
            color: item.color || best.colour || null,
            colorName: item.colorName || best.colourName || null,
            sku: item.sku || best.productExternalRef || null,
            factory: item.factory || best.factory || null,
            customer: item.customer || best.customerName || null,
            season: item.season || best.season || null,
            poNumber: item.poNumber || best.poNumber || null,
            unitCost: item.unitCost ?? toFiniteNumber(best.cost) ?? null,
            matchStatus: alreadyMatched ? item.matchStatus : (ambiguous ? 'ambiguous' : 'matched'),
            matchScore: alreadyMatched ? item.matchScore : bestResult.score,
            matchReason: alreadyMatched
                ? item.matchReason
                : ambiguous
                    ? `Multiple product export records share the top score (${bestResult.score})`
                    : bestResult.reasons.join('; '),
        };
    });
}

function normalizeStyleCandidates(style: string): string[] {
    const cleaned = String(style || '')
        .replace(/\s*\([^)]*\)\s*/g, '')
        .replace(/[^a-z0-9]/gi, '')
        .toUpperCase();
    const candidates = new Set<string>([cleaned]);
    if (/^NF0/.test(cleaned)) candidates.add(cleaned.slice(3));
    if (/^NF[^0]/.test(cleaned)) candidates.add(cleaned.slice(2));
    return [...candidates].filter(Boolean);
}

function normalizeSize(value: string | null | undefined): string {
    const normalized = String(value || '')
        .toLowerCase()
        .replace(/[^a-z0-9]/g, '');
    if (['os', '0os', 'onesize'].includes(normalized)) return 'onesize';
    return normalized;
}

function toFiniteNumber(value: string | number | null | undefined): number | null {
    if (value === null || value === undefined || value === '') return null;
    const parsed = typeof value === 'number' ? value : Number(String(value).replace(/,/g, ''));
    return Number.isFinite(parsed) ? parsed : null;
}

function normalizedText(value: string | null | undefined): string {
    return String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function productIdentity(row: ProductSheetRow): string {
    return [
        row.productName,
        row.productExternalRef,
        row.costingReference,
        row.sell,
        row.sizeName,
    ].map((value) => String(value || '')).join('|');
}

function deduplicateProductRows(rows: ProductSheetRow[]): ProductSheetRow[] {
    const unique = new Map<string, ProductSheetRow>();
    for (const row of rows) unique.set(productIdentity(row), row);
    return [...unique.values()];
}

function scoreProductMatch(
    item: BuyFileItem,
    row: ProductSheetRow
): { row: ProductSheetRow; score: number; reasons: string[] } {
    let score = 60; // style + color are exact because candidates came from the exact lookup key
    const reasons = ['exact style and color'];

    const itemSize = normalizeSize(item.size);
    const productSize = normalizeSize(row.sizeName);
    if (itemSize && productSize) {
        if (itemSize === productSize) {
            score += 15;
            reasons.push('size matched');
        } else {
            score -= 20;
            reasons.push('size differs');
        }
    }

    const buyPrice = toFiniteNumber(item.unitCost);
    const sellPrice = toFiniteNumber(row.sell);
    if (buyPrice !== null && sellPrice !== null) {
        if (Math.abs(buyPrice - sellPrice) <= 0.005) {
            score += 20;
            reasons.push('FOB matched Nexgen Sell');
        } else {
            score -= 25;
            reasons.push(`FOB ${buyPrice.toFixed(2)} differs from Nexgen Sell ${sellPrice.toFixed(2)}`);
        }
    }

    const itemCustomer = normalizedText(item.customer);
    const productCustomer = normalizedText(row.customerName);
    if (itemCustomer && productCustomer && (itemCustomer.includes(productCustomer) || productCustomer.includes(itemCustomer))) {
        score += 3;
        reasons.push('customer matched');
    }

    const itemFactory = normalizedText(item.factory);
    const productFactory = normalizedText(row.factory);
    if (itemFactory && productFactory && (itemFactory.includes(productFactory) || productFactory.includes(itemFactory))) {
        score += 2;
        reasons.push('factory matched');
    }

    return { row, score: Math.max(0, Math.min(100, score)), reasons };
}

const BUY_FILE_KEYWORDS = [
    'style', 'style number', 'style no', 'style #', 'article', 'model',
    'po', 'po number', 'po no', 'po#', 'order', 'purchase order',
    'quantity', 'qty', 'units',
    'color', 'colour', 'color code', 'colour code', 'option',
    'size', 'size name', 'size scale',
    'sku', 'upc', 'ean', 'product code',
    'factory', 'vendor', 'supplier', 'manufacturer',
    'customer', 'brand', 'buyer',
    'season', 'year', 'delivery', 'ex factory', 'ex-fty', 'ship date', 'crdd',
    'unit cost', 'cost', 'price', 'currency',
];

const SUMMARY_KEYWORDS = [
    'sum of', 'row labels', 'total', 'grand total', 'decision', 'count of', 'average of', 'min of', 'max of',
];

function scoreHeaders(headers: string[]): number {
    const normalized = headers.map((h) => String(h).toLowerCase());
    let score = 0;
    for (const keyword of BUY_FILE_KEYWORDS) {
        if (normalized.some((h) => h.includes(keyword))) {
            score += 1;
        }
    }
    for (const keyword of SUMMARY_KEYWORDS) {
        if (normalized.some((h) => h.includes(keyword))) {
            score -= 5;
        }
    }
    return score;
}

function isGeneratedNextGenExport(headers: string[]): boolean {
    const normalized = new Set(headers.map((header) => String(header).trim().toLowerCase()));
    const linesSignature = [
        'purchaseorder', 'lineitem', 'productrange', 'product',
        'deliverydate', 'transportmethod', 'template', 'findfield_product',
    ];
    return linesSignature.filter((header) => normalized.has(header)).length >= 6;
}

export interface BuyFileExtractionResult {
    items: BuyFileItem[];
    productData: ProductData[];
    headerRow: number;
    headers: string[];
    mapping: ColumnMapping;
    templateUsed: boolean;
    unmappedColumns: string[];
}

export async function extractBuyFile(
    fileBuffer: ArrayBuffer,
    customerHint?: string,
    sharedNextgenClient?: NextGenCachedClient,
    productSheetBuffers: ArrayBuffer[] = []
): Promise<BuyFileExtractionResult> {
    console.log('[buy-file-extractor] reading workbook');
    const { worksheets } = await readExcelFile(fileBuffer);
    console.log('[buy-file-extractor] worksheets found:', worksheets.length);

    // Only explicit reference workbooks may act as a product sheet. Treating the
    // buyer file itself as a reference creates false-positive Nexgen matches.
    const productSheetMap = await buildProductSheetMap(productSheetBuffers);
    console.log('[buy-file-extractor] product sheet map keys:', Object.keys(productSheetMap).length);

    // 1. Find the best worksheet: score all sheets by buy-file header keywords
    const sheetCandidates: { sheet: typeof worksheets[0]; headerRow: number; headers: string[]; score: number }[] = [];

    for (const sheet of worksheets) {
        const preview = sheet.rows.slice(0, 10);
        if (!preview.length) continue;

        // Score rows locally first. Do not send low-information summary/pivot
        // tabs to the LLM: they can add minutes without helping sheet selection.
        const locallyScoredRows = preview
            .map((row, index) => ({
                index,
                score: scoreHeaders(row.map((cell) => String(cell || '')).filter(Boolean)),
            }))
            .sort((a, b) => b.score - a.score);
        const localBest = locallyScoredRows[0];

        // Deterministic fast path, else AI detection for tabular-looking
        // sheets even with zero keyword score (foreign-language layouts).
        const maxNonEmpty = preview.reduce((mx, r) => Math.max(
            mx,
            r.filter((c) => c !== null && c !== undefined && String(c).trim() !== '').length
        ), 0);
        // Header-row choice, structural first: a real header row is almost
        // entirely short text labels, while data rows mix numbers and dates.
        // Keyword scores are only a tie-breaker — raw substring scoring alone
        // misfires on product names ("Ridge Beanie" contains "ean").
        let bestIdx = -1;
        let bestRatio = 0;
        let bestKw = -1;
        preview.forEach((r, idx) => {
            const cells = r.filter((c) => c !== null && c !== undefined && String(c).trim() !== '');
            if (cells.length < 4) return;
            const strCells = cells.filter((c) =>
                !(c instanceof Date) && typeof c === 'string' && c.length <= 60 && isNaN(Number(c))
            );
            const ratio = strCells.length / cells.length;
            const kw = scoreHeaders(r.map(String));
            if (
                ratio > bestRatio + 0.001 ||
                (Math.abs(ratio - bestRatio) <= 0.001 && kw > bestKw)
            ) {
                bestRatio = ratio;
                bestKw = kw;
                bestIdx = idx;
            }
        });

        let detected: { headerRow: number } | null = null;
        if (bestIdx >= 0 && bestRatio >= 0.85) {
            detected = { headerRow: bestIdx + 1 };
        } else if (maxNonEmpty >= 4) {
            detected = await detectHeaderRow(preview);
        } else {
            continue;
        }

        let headerRow = sheet.rows[detected.headerRow - 1] || [];
        let headers = headerRow.map((h) => String(h || '')).filter(Boolean);
        let headerRowIndex = detected.headerRow;

        // Sanity guard: if the chosen "header" row is mostly numbers/dates it
        // is a data row, not a header (AI can mis-pick on foreign layouts).
        // Fall back to the densest preview row in that case.
        const filledCells = headerRow.filter((c) => c !== null && c !== undefined && String(c).trim() !== '');
        const numericish = filledCells.filter((c) => c instanceof Date || !isNaN(Number(c))).length;
        if (filledCells.length >= 4 && numericish / filledCells.length > 0.5) {
            let bestRow = 0;
            let bestCount = 0;
            preview.forEach((row, idx) => {
                const count = row.filter((cell) => cell !== null && cell !== undefined && String(cell).trim() !== '').length;
                if (count > bestCount) {
                    bestCount = count;
                    bestRow = idx;
                }
            });
            headerRow = sheet.rows[bestRow] || [];
            headers = headerRow.map((h) => String(h || '')).filter(Boolean);
            headerRowIndex = bestRow + 1;
        }

        // Fallback: if AI-selected row is empty, pick the row with most non-empty cells
        if (!headers.length) {
            let bestRow = 0;
            let bestCount = 0;
            preview.forEach((row, idx) => {
                const count = row.filter((cell) => cell !== null && cell !== undefined && String(cell).trim() !== '').length;
                if (count > bestCount) {
                    bestCount = count;
                    bestRow = idx;
                }
            });
            headerRow = sheet.rows[bestRow] || [];
            headers = headerRow.map((h) => String(h || '')).filter(Boolean);
            headerRowIndex = bestRow + 1;
        }

        if (headers.length) {
            sheetCandidates.push({ sheet, headerRow: headerRowIndex, headers, score: scoreHeaders(headers) });
        }
    }

    if (!sheetCandidates.length) {
        throw new Error('Could not detect headers in any worksheet');
    }

    // Pick sheet with highest score; if tie, prefer the one with more data rows
    sheetCandidates.sort((a, b) => {
        if (b.score !== a.score) return b.score - a.score;
        return b.sheet.rows.length - a.sheet.rows.length;
    });

    console.log('[buy-file-extractor] sheet candidates:', sheetCandidates.map((c) => ({ name: c.sheet.name, score: c.score, rows: c.sheet.rows.length })));

    // Try candidates in order until we find one with actual items
    let lastError: Error | null = null;
    for (const candidate of sheetCandidates) {
        try {
            console.log(`[buy-file-extractor] trying sheet: "${candidate.sheet.name}" header row: ${candidate.headerRow} score: ${candidate.score}`);
            const result = await extractFromSheet(
                candidate.sheet,
                candidate.headerRow,
                candidate.headers,
                customerHint,
                sharedNextgenClient,
                productSheetMap
            );
            if (result.items.length > 0) {
                console.log(`[buy-file-extractor] selected sheet: "${candidate.sheet.name}" with ${result.items.length} items`);
                return result;
            }
            console.warn(`[buy-file-extractor] sheet "${candidate.sheet.name}" produced 0 items, trying next`);
        } catch (err) {
            lastError = err instanceof Error ? err : new Error(String(err));
            console.warn(`[buy-file-extractor] sheet "${candidate.sheet.name}" failed:`, lastError.message);
        }
    }

    if (lastError) {
        throw lastError;
    }

    throw new Error('No data found in any worksheet');
}

// ---------------------------------------------------------------------------
// Content-based header inference
//
// If we know what the cell values look like, we can infer which column is
// which — even if the header name is ambiguous or the fuzzy mapping got it
// wrong. This is brand-agnostic and works for any Excel file structure.
//
// Strategy:
//   1. Sample the first N data rows
//   2. For each column, classify the content (style code? style name? color
//      code? quantity? date?)
//   3. If a column's content strongly suggests a canonical field that wasn't
//      mapped (or was mapped to the wrong field), update the mapping
// ---------------------------------------------------------------------------

function isStyleCodeValue(v: string): boolean {
    const s = String(v || '').trim();
    if (!s || s.length > 20) return false;
    return /\d/.test(s) && !/\s/.test(s);
}

function isStyleNameValue(v: string): boolean {
    const s = String(v || '').trim();
    if (!s || s.length < 3) return false;
    // Style names have spaces and are mostly alpha
    return /\s/.test(s) && !/^\d+$/.test(s);
}

function isColorCodeValue(v: string): boolean {
    const s = String(v || '').trim();
    if (!s || s.length > 20) return false;
    // Color codes: short, alphanumeric, no spaces, often 3-10 chars (E8J, JK3, BOA)
    return s.length >= 2 && s.length <= 12 && /[a-zA-Z]/.test(s) && /\d/.test(s) && !/\s/.test(s);
}

function isQuantityValue(v: string | number): boolean {
    const n = Number(v);
    return !isNaN(n) && n > 0 && Number.isFinite(n) && n === Math.floor(n);
}

function isDateValue(v: any): boolean {
    return v instanceof Date || (typeof v === 'string' && /\d{4}[-/]\d{2}[-/]\d{2}/.test(v));
}

/**
 * Looks at cell values in the first N data rows and infers which column
 * maps to which canonical field. Overrides or supplements the header-based
 * mapping.
 *
 * Returns an updated ColumnMapping.
 */
function inferHeadersFromContent(
    rows: (string | number | Date | null)[][],
    headerRowIndex: number,
    headers: string[],
    existingMapping: ColumnMapping,
): ColumnMapping {
    const mapping = { ...existingMapping };
    const sampleSize = Math.min(20, rows.length - headerRowIndex - 1);
    if (sampleSize < 3) return mapping; // not enough data to infer

    // Track which headers are already mapped
    const mappedHeaders = new Set(Object.values(mapping as Record<string, string>));
    const mappedFields = new Set(Object.keys(mapping as Record<string, string>));

    // Analyze each column
    const columnStats: Record<number, {
        styleCodeCount: number;
        styleNameCount: number;
        colorCodeCount: number;
        quantityCount: number;
        dateCount: number;
        nonEmptyCount: number;
        totalLength: number;
        alphaCount: number;
    }> = {};

    // The `headers` array passed here may be FILTERED (blank header cells
    // removed), while `rows` keep raw Excel column positions. Build the
    // filtered→raw index translation once so every cell read below lands on
    // the column the label actually belongs to.
    const rawHeaderRow = rows[headerRowIndex - 1] || [];
    const rawIdxOfFiltered: number[] = [];
    {
        let k = 0;
        for (let i = 0; i < rawHeaderRow.length; i++) {
            const v = rawHeaderRow[i];
            if (v === null || v === undefined || String(v).trim() === '') continue;
            rawIdxOfFiltered[k++] = i;
        }
    }
    const rawColOf = (filteredIdx: number): number =>
        rawIdxOfFiltered[filteredIdx] !== undefined ? rawIdxOfFiltered[filteredIdx] : filteredIdx;

    for (let c = 0; c < headers.length; c++) {
        columnStats[c] = {
            styleCodeCount: 0,
            styleNameCount: 0,
            colorCodeCount: 0,
            quantityCount: 0,
            dateCount: 0,
            nonEmptyCount: 0,
            totalLength: 0,
            alphaCount: 0,
        };

        const rawC = rawColOf(c);
        for (let r = headerRowIndex + 1; r <= headerRowIndex + sampleSize && r < rows.length; r++) {
            const row = rows[r];
            if (!row) continue;
            const val = row[rawC];
            if (val === null || val === undefined || val === '') continue;

            const stats = columnStats[c];
            stats.nonEmptyCount++;
            const strVal = String(val).trim();

            if (/[a-zA-Z]/.test(strVal)) stats.alphaCount++;
            if (isStyleCodeValue(strVal)) stats.styleCodeCount++;
            else if (isStyleNameValue(strVal)) stats.styleNameCount++;

            if (isColorCodeValue(strVal)) stats.colorCodeCount++;
            if (isQuantityValue(val as string | number)) stats.quantityCount++;
            if (isDateValue(val)) stats.dateCount++;
            stats.totalLength += strVal.length;
        }
    }

    // Header-based veto lists. A column whose HEADER clearly declares another
    // semantic must never be claimed by content inference, no matter how its
    // values happen to classify (e.g. "Document Date" holding numeric Excel
    // serials looks like quantities; "Vendor Account" looks like a style code).
    const DATEISH_HEADER = /\b(date|etd|crd|ped|xf|arrival|handover|issued)\b|ex[-\s]?fac/i;
    const ADMIN_HEADER = /\b(account|email|planner|sbu|category|type|status|priority|remark|comment|approval|warehouse|^site\b|round|surcharge|terms|payment|incoterm|reference|instr\.?\b|mode\b)\b/i;
    const ORG_NAME_HEADER = /\b(vendor|supplier|factory|plant|warehouse|customer|buyer|company)\b.*\b(name|account|code|plnt)\b|\bname\b/i;
    const STYLEISH_HEADER = /\b(style|item|material|article|model|sku|product|dev\s*#)\b/i;

    // Score each column for each canonical field
    const fieldCandidates: Record<string, { colIndex: number; score: number; header: string }[]> = {
        buyer_style_number: [],
        buyer_style_name: [],
        color_code: [],
        quantity: [],
        delivery_date: [],
    };

    const ratio = (stats: { nonEmptyCount: number }, count: number) => stats.nonEmptyCount > 0 ? count / stats.nonEmptyCount : 0;

    for (let c = 0; c < headers.length; c++) {
        const stats = columnStats[c];
        if (stats.nonEmptyCount < 2) continue;
        const header = headers[c];
        const nh = String(header).toLowerCase().replace(/\s+/g, ' ').trim();
        const isDateishHeader = DATEISH_HEADER.test(nh);
        const isAdminHeader = ADMIN_HEADER.test(nh);
        const isOrgNameHeader = ORG_NAME_HEADER.test(nh);

        // Style code: most values look like codes (NF0A8CGZ, VN0A3XYZ)
        if (
            ratio(stats, stats.styleCodeCount) >= 0.6
            && !isDateishHeader && !isAdminHeader && !isOrgNameHeader
            // Purely numeric columns must at least have a style-ish header,
            // otherwise any ID or date-serial column qualifies.
            && (stats.alphaCount > 0 || STYLEISH_HEADER.test(nh))
        ) {
            fieldCandidates.buyer_style_number.push({
                colIndex: c,
                score: ratio(stats, stats.styleCodeCount),
                header,
            });
        }

        // Style name: most values look like names (SALTY LINED BEANIE)
        if (
            ratio(stats, stats.styleNameCount) >= 0.6
            && !isDateishHeader && !isAdminHeader && !isOrgNameHeader
            && !/\b(po\b|number|code|qty|quantity|season|status)\b/i.test(nh)
        ) {
            fieldCandidates.buyer_style_name.push({
                colIndex: c,
                score: ratio(stats, stats.styleNameCount),
                header,
            });
        }

        // Color code: most values look like color codes (E8J, JK3)
        if (
            ratio(stats, stats.colorCodeCount) >= 0.5
            && stats.nonEmptyCount >= 3
            && !isDateishHeader && !isAdminHeader && !isOrgNameHeader
        ) {
            fieldCandidates.color_code.push({
                colIndex: c,
                score: ratio(stats, stats.colorCodeCount),
                header,
            });
        }

        // Quantity: most values are positive integers
        if (
            ratio(stats, stats.quantityCount) >= 0.7
            && !isDateishHeader
            && !/\b(price|cost|fob|year|surcharge|uc)\b/i.test(nh)
        ) {
            fieldCandidates.quantity.push({
                colIndex: c,
                score: ratio(stats, stats.quantityCount),
                header,
            });
        }

        // Delivery date: most values are dates; veto clearly non-date columns
        if (
            ratio(stats, stats.dateCount) >= 0.6
            && !/\b(qty|quantity|price|cost|fob)\b/i.test(nh)
        ) {
            fieldCandidates.delivery_date.push({
                colIndex: c,
                score: ratio(stats, stats.dateCount),
                header,
            });
        }
    }

    // Sort candidates by score (highest first)
    for (const field of Object.keys(fieldCandidates)) {
        fieldCandidates[field].sort((a, b) => b.score - a.score);
    }

    // PO-number disambiguation: several layouts carry multiple PO-like columns
    // (often one legacy/empty and one filled). Prefer the PO-ish column that
    // actually contains data; break ties by canonical pattern rank, then
    // leftmost position. Only overrides the deterministic pick when the
    // currently-mapped column is substantially emptier than the alternative.
    {
        const POISH_HEADER = /^(po#|po #|po no\.?|po number|purchase order( no| number|#)?|purchasing document|final po cut|master po#?|new po|so#)$/i;
        const PO_RANK = ['final po cut', 'master po', 'po number', 'purchase order', 'purchasing document', 'new po', 'po#'];
        const rankOf = (h: string) => {
            const idx = PO_RANK.findIndex((r) => h.startsWith(r));
            return idx === -1 ? PO_RANK.length : idx;
        };
        const poCols: { c: number; nh: string; fill: number }[] = [];
        for (let c = 0; c < headers.length; c++) {
            const st = columnStats[c];
            if (!st || st.nonEmptyCount < 1) continue;
            const nh = String(headers[c]).toLowerCase().replace(/\s+/g, ' ').trim();
            if (!POISH_HEADER.test(nh)) continue;
            poCols.push({ c, nh, fill: Math.min(1, st.nonEmptyCount / sampleSize) });
        }
        if (poCols.length >= 1 && sampleSize > 0) {
            poCols.sort((a, b) => b.fill - a.fill || rankOf(a.nh) - rankOf(b.nh) || a.c - b.c);
            const best = poCols[0];
            const cur = (mapping as Record<string, string>).po_number;
            const curIdx = cur ? headers.indexOf(cur) : -1;
            const curFill = curIdx >= 0 && columnStats[curIdx] ? Math.min(1, columnStats[curIdx].nonEmptyCount / sampleSize) : 0;
            if (best.c !== curIdx && best.fill > curFill + 0.2) {
                (mapping as Record<string, string>).po_number = headers[best.c];
                console.log(`[content-inference] po_number <= "${headers[best.c]}" (fill ${(best.fill * 100).toFixed(0)}% vs ${Math.round(curFill * 100)}%)`);
            }
        }
    }

    // Apply content-based inference:
    // 1. If buyer_style_number is NOT mapped but we found a style code column → map it
    // 2. If buyer_style_number IS mapped but to a style NAME column, and we found a
    //    better style CODE column → override it
    // 3. Same logic for other fields

    const usedColumns = new Set<number>();

    // Mark already-mapped columns as used
    for (let c = 0; c < headers.length; c++) {
        if (mappedHeaders.has(headers[c])) usedColumns.add(c);
    }

    for (const [field, candidates] of Object.entries(fieldCandidates)) {
        if (candidates.length === 0) continue;

        const best = candidates[0];
        if (usedColumns.has(best.colIndex)) continue;

        const currentMapping = (mapping as Record<string, string>)[field];

        if (!currentMapping) {
            // Field not mapped — use content-based inference
            (mapping as Record<string, string>)[field] = best.header;
            usedColumns.add(best.colIndex);
            mappedHeaders.add(best.header);
            console.log(`[content-inference] ${field} <= "${best.header}" (content-based, score: ${(best.score * 100).toFixed(0)}%)`);
        } else if (field === 'buyer_style_number') {
            // Special case: if buyer_style_number is mapped to a column whose
            // content looks like style NAMES (not codes), and we found a column
            // whose content looks like style CODES, override it.
            const currentColIndex = headers.indexOf(currentMapping);
            const currentStats = columnStats[currentColIndex];
            if (currentStats && ratio(currentStats, currentStats.styleNameCount) >= 0.5 && ratio(currentStats, currentStats.styleCodeCount) < 0.2) {
                // Current mapping points to a style NAME column — override with style CODE column
                (mapping as Record<string, string>)[field] = best.header;
                usedColumns.add(best.colIndex);
                mappedHeaders.add(best.header);
                mappedHeaders.delete(currentMapping);
                usedColumns.delete(currentColIndex);
                // Try to remap the old column to buyer_style_name
                if (!mappedFields.has('buyer_style_name')) {
                    (mapping as Record<string, string>)['buyer_style_name'] = currentMapping;
                    mappedHeaders.add(currentMapping);
                    usedColumns.add(currentColIndex);
                    mappedFields.add('buyer_style_name');
                    console.log(`[content-inference] buyer_style_name <= "${currentMapping}" (remapped from buyer_style_number)`);
                }
                console.log(`[content-inference] buyer_style_number <= "${best.header}" (OVERRIDE: was "${currentMapping}", score: ${(best.score * 100).toFixed(0)}%)`);
            }
        }
    }

    // Contradiction overrides: an LLM/heuristic may have bound a field to a
    // column whose CONTENT flatly contradicts the field type (e.g. quantity ←
    // a free-text column). If some unused column matches the field's content
    // almost perfectly, steal the binding.
    const stealIfContradicts = (
        field: string,
        ratioOf: (s: { quantityCount: number; dateCount: number; nonEmptyCount: number }) => number,
        minNew: number,
        maxCur: number,
    ) => {
        const candidates = fieldCandidates[field];
        if (!candidates?.length) return;
        const currentMapping = (mapping as Record<string, string>)[field];
        if (!currentMapping) return;
        const curIdx = headers.indexOf(currentMapping);
        const curStats = curIdx >= 0 ? columnStats[curIdx] : undefined;
        const curRatio = curStats ? ratioOf(curStats) : 0;
        if (curRatio > maxCur) return;
        const best = candidates.find((c) => c.colIndex !== curIdx && !usedColumns.has(c.colIndex));
        if (!best) return;
        const bestStats = columnStats[best.colIndex];
        if (!bestStats || ratioOf(bestStats) < minNew) return;
        (mapping as Record<string, string>)[field] = best.header;
        usedColumns.add(best.colIndex);
        mappedHeaders.add(best.header);
        if (curIdx >= 0) {
            usedColumns.delete(curIdx);
            mappedHeaders.delete(currentMapping);
        }
        console.log(`[content-inference] ${field} <= "${best.header}" (CONTRADICTION override: was "${currentMapping}" at ${(curRatio * 100).toFixed(0)}%, new ${(ratioOf(bestStats) * 100).toFixed(0)}%)`);
    };

    stealIfContradicts('quantity', (s) => ratio(s, s.quantityCount), 0.8, 0.2);
    stealIfContradicts('delivery_date', (s) => ratio(s, s.dateCount), 0.8, 0.2);

    return mapping;
}

/**
 * Detect brand from headers, customer hint, and cell content.
 * Checks:
 *   1. customerHint (filename-derived) against brand aliases
 *   2. Header text against brand aliases
 *   3. Cell values in customer/brand columns against brand aliases
 *   4. Style code prefixes (NF0A = TNF, VN0A = Vans, etc.)
 */
function detectBrandFromContent(
    headers: string[],
    rows: (string | number | Date | null)[][],
    headerRowIndex: number,
    customerHint?: string,
): string | null {
    const aliases = getAllBrandAliases();

    // 1. Check customerHint
    if (customerHint) {
        const brand = lookupBrand(customerHint);
        if (brand) return brand;
        // Try partial match (e.g. "JUL - Buy File INDONESIA MADISON 88" contains no brand)
        // but "TNF Buy File" would match
        const lower = customerHint.toLowerCase();
        for (const alias of aliases) {
            if (lower.includes(alias.toLowerCase())) {
                const b = lookupBrand(alias);
                if (b) return b;
            }
        }
    }

    // 2. Check headers
    for (const header of headers) {
        const brand = lookupBrand(header);
        if (brand) return brand;
    }

    // 3. Check cell values in first few data rows
    const sampleSize = Math.min(10, rows.length - headerRowIndex - 1);
    for (let r = headerRowIndex + 1; r <= headerRowIndex + sampleSize && r < rows.length; r++) {
        const row = rows[r];
        if (!row) continue;
        for (const cell of row) {
            const val = String(cell || '').trim();
            if (!val) continue;
            const brand = lookupBrand(val);
            if (brand) return brand;
        }
    }

    // 4. Check style code prefixes in data
    for (let r = headerRowIndex + 1; r <= headerRowIndex + sampleSize && r < rows.length; r++) {
        const row = rows[r];
        if (!row) continue;
        for (const cell of row) {
            const val = String(cell || '').trim();
            if (!val || !/\d/.test(val) || /\s/.test(val)) continue;
            // NF0A... = TNF, VN0A... = Vans, etc.
            if (/^NF0[A-Z]/i.test(val)) return lookupBrand('tnf');
            if (/^VN0[A-Z]/i.test(val)) return lookupBrand('vans');
        }
    }

    return null;
}

async function extractFromSheet(
    sheet: { name: string; rows: (string | number | Date | null)[][] },
    headerRow: number,
    headers: string[],
    customerHint?: string,
    sharedNextgenClient?: NextGenCachedClient,
    productSheetMap: Record<string, ProductSheetRow[]> = {}
): Promise<BuyFileExtractionResult> {
    if (isGeneratedNextGenExport(headers)) {
        throw new Error('This workbook is a generated NextGen LINES output, not a raw Buy File. Upload the original brand Buy File instead.');
    }

    // NOTE: This function returns a single-sheet extraction result. productData is
    // built from rows within this sheet only.

    let mapping: ColumnMapping | null = null;
    let unmappedColumns: string[] = [];
    let templateUsed = false;

    // 2. Check for learned template

    console.log('[buy-file-extractor] sheet headers:', JSON.stringify(headers));

    const existingTemplate = await findMatchingTemplateSupabase(headers);
    if (existingTemplate) {
        console.log('[buy-file-extractor] using existing template', existingTemplate.id);
        mapping = existingTemplate.mapping;
        templateUsed = true;
    } else {
        // 3. Build mapping from legacy DB + AI + heuristic fallback
        const legacyMapping = await loadLegacyMapping(customerHint);
        const aiMappingResult = await mapHeaders(headers, customerHint);

        mapping = mergeMappings(headers, legacyMapping, aiMappingResult.mapping);
        unmappedColumns = headers.filter((h) => !Object.values(mapping as Record<string, string>).includes(h));
        templateUsed = false;
    }

    console.log('[buy-file-extractor] mapping:', JSON.stringify(mapping));
    console.log('[buy-file-extractor] unmappedColumns:', JSON.stringify(unmappedColumns));

    // 4a. Content-based header inference: look at cell values to fill gaps
    // or fix wrong mappings. This is brand-agnostic and works for any Excel file.
    if (mapping) {
        const mappingBeforeInference = JSON.stringify(mapping);
        mapping = inferHeadersFromContent(sheet.rows, headerRow, headers, mapping);
        if (JSON.stringify(mapping) !== mappingBeforeInference) {
            console.log('[buy-file-extractor] mapping after content inference:', JSON.stringify(mapping));
        }
    }

    // 4b. Read all rows locally
    console.log('[buy-file-extractor] reading all rows locally');

    // Detect brand from headers, customer hint, and cell content
    const detectedBrand = detectBrandFromContent(headers, sheet.rows, headerRow, customerHint);
    if (detectedBrand) {
        console.log('[buy-file-extractor] detected brand:', detectedBrand);
    }

    let items = readAllRows(sheet.rows, headerRow, mapping, sheet.name, customerHint, detectedBrand || undefined);
    console.log('[buy-file-extractor] extracted items:', items.length);

    // Colour derivation for concatenated article codes (e.g. Smartwool
    // "SW0026470011" = style "SW002647" + colour "0011", "SW002997Q671" =
    // style + colour "Q671"). Only fires when the item has a SKU that
    // literally starts with the style and the remainder is a short
    // alphanumeric colour code.
    items = items.map((item) => {
        if (item.colorCode || item.colorName || !item.sku || !item.style) return item;
        const sku = String(item.sku).toUpperCase();
        const style = String(item.style).toUpperCase();
        if (!sku.startsWith(style)) return item;
        const rest = sku.slice(style.length);
        if (/^[A-Z0-9]{2,6}$/.test(rest)) {
            return { ...item, colorCode: rest };
        }
        return item;
    });

    // PO fallback: some layouts carry a secondary dummy-PO column (e.g.
    // Smartwool "PR Number/ Dummy PO Number") that is filled when the primary
    // PO column is blank.
    {
        const usedValsPo = new Set(Object.values(mapping as Record<string, string>).map((v) => String(v).toLowerCase()));
        const poFallbackIdx = headers.findIndex(
            (h) => /\bdummy\s*po\b/i.test(String(h)) && !usedValsPo.has(String(h).toLowerCase())
        );
        if (poFallbackIdx >= 0) {
            items = items.map((item) => {
                if (item.poNumber) return item;
                const row = sheet.rows[item.sourceRow - 1];
                const val = row ? row[poFallbackIdx] : null;
                return val !== null && val !== undefined && String(val).trim()
                    ? { ...item, poNumber: String(val).trim() }
                    : item;
            });
        }
    }

    // Fill colour NAME from an unmapped "<Colour/Color> Name" sibling column
    // (e.g. Dynafit maps colour→"Colour" codes while "Color Name" carries
    // "DYN-0601 Smoke 0910"). Needed before the style derivation below.
    {
        const usedVals = new Set(Object.values(mapping as Record<string, string>).map((v) => String(v).toLowerCase()));
        const siblingIdx = headers.findIndex(
            (h) => /^(color|colour)\s*name$/i.test(String(h).trim()) && !usedVals.has(String(h).toLowerCase())
        );
        if (siblingIdx >= 0) {
            items = items.map((item) => {
                if (item.colorName) return item;
                const row = sheet.rows[item.sourceRow - 1];
                const val = row ? row[siblingIdx] : null;
                return val !== null && val !== undefined && String(val).trim()
                    ? { ...item, colorName: String(val).trim() }
                    : item;
            });
        }
    }

    // Style derivation for layouts without any style column (e.g. Dynafit
    // encodes it in the colour text: "DYN-0601 Smoke 0910" → "DYN-0601",
    // tolerating the "DYN- 2881" spacing variant).
    items = items.map((item) => {
        if (item.style) return item;
        const src = item.colorName || item.color;
        if (!src) return item;
        const token = String(src).trim().match(/^([A-Za-z]{2,6}-\s*[A-Za-z0-9]{2,10})/);
        if (token) {
            return { ...item, style: token[1].replace(/\s+/g, '').toUpperCase() };
        }
        return item;
    });

    // 5. Query NextGen FIRST (primary enrichment source).
    // NextGen now provides product, color, factory, cost, customer, season data.
    // Product sheet is only a fallback for fields NextGen couldn't fill.
    const uniqueStyles = [...new Set(items.map((item) => item.style || '').filter(Boolean))];
    console.log('[buy-file-extractor] unique styles:', uniqueStyles.length);

    const nextgenEnabled = process.env.NEXTGEN_ENABLED !== 'false';
    let nextgenInfo: Record<string, NextGenStyleInfo | null> = {};

    if (nextgenEnabled) {
        const nextgenClient = sharedNextgenClient || new NextGenCachedClient();
        const variants = [...new Map(items.map((item) => {
            const style = String(item.style || '').trim();
            const color = String(item.colorCode || item.color || '').trim();
            const brand = String(item.brand || '').trim();
            return [`${style.toLowerCase()}|${color.toLowerCase()}`, { style, color, brand }];
        })).entries()];
        for (const [key, variant] of variants) {
            nextgenInfo[key] = await nextgenClient.searchVariant(variant.style, variant.color, variant.brand);
        }
        // Enrich items with NextGen data (primary source)
        items = items.map((item) => {
            const variantKey = `${String(item.style || '').toLowerCase()}|${String(item.colorCode || item.color || '').toLowerCase()}`;
            const ngMatch = nextgenInfo[variantKey] || null;
            if (!ngMatch) {
                return {
                    ...item,
                    matchStatus: 'unmatched',
                    matchScore: 0,
                    matchReason: `Buyer style ${item.style || '(blank)'} was not found in Nexgen`,
                };
            }
            return {
                ...item,
                // NextGen enrichment — fills product, color, factory, cost, customer, season
                product: ngMatch.product || item.product || null,
                productExternalRef: ngMatch.productExternalRef || item.productExternalRef || item.sku || null,
                colorName: ngMatch.colorName || item.colorName || null,
                colorCode: ngMatch.colorCode || item.colorCode || null,
                factory: ngMatch.factory || item.factory || null,
                customer: ngMatch.customer || item.customer || null,
                season: ngMatch.season || item.season || null,
                currency: ngMatch.currency || item.currency || null,
                unitCost: ngMatch.unitCost ?? item.unitCost ?? null,
                costingReference: ngMatch.costingReference || item.costingReference || null,
                matchStatus: ngMatch.matchStatus || 'matched',
                matchScore: ngMatch.matchScore ?? 100,
                matchReason: ngMatch.matchReason || `Matched buyer style ${item.style} in Nexgen`,
            };
        });
        console.log('[buy-file-extractor] enriched items with NextGen:', items.length);
    } else {
        console.log('[buy-file-extractor] NextGen disabled, skipping style lookups');
    }

    // 4.5 Fallback: enrich with product sheet data for fields NextGen didn't fill.
    // This is now a secondary fallback, not the primary source.
    if (Object.keys(productSheetMap).length) {
        items = enrichItemsWithProductSheet(items, productSheetMap);
        console.log('[buy-file-extractor] fallback enrichment with product sheet:', items.length);
    }

    // 6. Merge into single source of truth (ProductData)
    const productData = mergeBuyFileWithNextGen(items, nextgenInfo);

    return {
        items,
        productData,
        headerRow,
        headers,
        mapping: mapping || {},
        templateUsed,
        unmappedColumns,
    };
}

function mergeMappings(headers: string[], ...mappings: (ColumnMapping | null)[]): ColumnMapping {
    const result: ColumnMapping = {};
    const headerSet = new Set(headers);
    for (const mapping of mappings) {
        if (!mapping) continue;
        for (const [field, header] of Object.entries(mapping)) {
            if (header && headerSet.has(header) && !(result as Record<string, string>)[field]) {
                (result as Record<string, string>)[field] = header;
            }
        }
    }
    return result;
}

function normalizeHeaderName(header: string): string {
    return String(header || '')
        .toLowerCase()
        .replace(/[^a-z0-9]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
}

function readAllRows(
    rows: (string | number | Date | null)[][],
    headerRowIndex: number,
    mapping: ColumnMapping | null,
    sheetName: string,
    customerHint?: string,
    brandHint?: string,
): BuyFileItem[] {
    if (!mapping) return [];

    const headerToIndex: Record<string, number> = {};
    const normalizedHeaderToIndex: Record<string, number> = {};
    const headerRow = rows[headerRowIndex - 1] || [];
    headerRow.forEach((cell, idx) => {
        const raw = String(cell || '').trim();
        headerToIndex[raw] = idx;
        normalizedHeaderToIndex[normalizeHeaderName(raw)] = idx;
    });

    const reverseMapping: Record<string, string> = {};
    for (const [canonicalField, headerName] of Object.entries(mapping)) {
        if (headerName) {
            reverseMapping[String(headerName)] = canonicalField;
            reverseMapping[normalizeHeaderName(String(headerName))] = canonicalField;
        }
    }

    const items: BuyFileItem[] = [];
    for (let i = headerRowIndex; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length === 0) continue;

        const get = (canonicalField: string): string | null => {
            const headerName = mapping?.[canonicalField as keyof ColumnMapping];
            if (!headerName) return null;
            let idx = headerToIndex[headerName];
            if (idx === undefined || idx < 0) {
                idx = normalizedHeaderToIndex[normalizeHeaderName(headerName)];
            }
            if (idx === undefined || idx < 0) return null;
            const val = row[idx];
            if (val === null || val === undefined) return null;
            return String(val).trim();
        };
        const getRawHeader = (...headerNames: string[]): string | null => {
            for (const headerName of headerNames) {
                let idx = headerToIndex[headerName];
                if (idx === undefined || idx < 0) {
                    idx = normalizedHeaderToIndex[normalizeHeaderName(headerName)];
                }
                if (idx === undefined || idx < 0) continue;
                const val = row[idx];
                if (val !== null && val !== undefined && String(val).trim()) {
                    return String(val).trim();
                }
            }
            return null;
        };

        const qtyStr = get('quantity');
        const qty = qtyStr ? Number(qtyStr.replace(/,/g, '')) : null;

        const unitCostStr = get('unit_cost');
        const unitCost = unitCostStr ? Number(unitCostStr.replace(/,/g, '')) : null;

        const style = get('buyer_style_number');
        const quantity = qty && !isNaN(qty) ? qty : null;
        if (!style && quantity === null) continue;
        const poNumber = get('po_number')
            || getRawHeader('FINAL PO CUT #', 'MASTER PO#', 'PURCHASE REQUISITION');

        items.push({
            style,
            styleName: get('buyer_style_name'),
            sku: get('sku'),
            description: get('product_description'),
            color: get('color'),
            colorCode: get('color_code'),
            colorName: null,
            size: get('size') || 'One Size',
            quantity,
            deliveryDate: get('delivery_date'),
            season: get('season'),
            customer: get('customer'),
            brand: brandHint || customerHint || null,
            factory: get('factory'),
            currency: get('currency') || 'USD',
            unitCost: unitCost && !isNaN(unitCost) ? unitCost : null,
            poNumber,
            product: null,
            productExternalRef: null,
            costingReference: null,
            buyInformation: get('buy_information'),
            matchStatus: 'not_checked',
            matchScore: null,
            matchReason: null,
            sourceSheet: sheetName,
            sourceRow: i + 1,
        });
    }

    return items;
}
