import { readExcelFile } from '@/lib/excel/excel-reader';
import { detectHeaderRow } from '@/lib/ai/header-detector-PF62XX6H';
import { mapHeaders, fallbackHeuristicMapping } from '@/lib/ai/header-mapper-PF62XX6H';
import { findMatchingTemplateSupabase } from '@/lib/templates/supabase-store';
import { getColumnMapping } from '@/lib/data-loader';
import { NextGenCachedClient } from '@/lib/nextgen/client';
import { ExcelEngine } from '@/lib/excel-engine';
import type { ProductSheetRow } from '@/lib/excel-engine';
import { mergeBuyFileWithNextGen } from '@/lib/merge/merge-buy-nextgen';
import { BuyFileItem, ColumnMapping, NextGenStyleInfo, ProductData } from '@/lib/types/buy-file';
import { detectBrand, resolveColorName } from '@/lib/colors/brand-colors';

const INTERNAL_TO_CANONICAL: Record<string, keyof ColumnMapping> = {
    purchaseOrder: 'po_number',
    buyerPoNumber: 'buyer_po_number',
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
    startDate: 'start_date',
    cancelDate: 'cancel_date',
    transportMethod: 'transport_method',
    buyInformation: 'buy_information',
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
        const styleRaw = engine.stripBrackets(item.style || '').trim();
        const colorRaw = item.color || item.colorCode || '';
        const colorKey = engine.normalizeColourKey(colorRaw);
        const lookupKey = `${styleRaw}|${colorKey}`;
        const matches = productSheetMap[lookupKey] || [];

        // Fallback: match by style only if color key failed
        const styleMatches = matches.length
            ? []
            : Object.entries(productSheetMap)
                  .filter(([key]) => key.startsWith(`${styleRaw}|`))
                  .flatMap(([_, rows]) => rows);

        const best = matches[0] || styleMatches[0];
        if (!best) return item;

        const cost = typeof best.cost === 'number' ? best.cost : null;

        return {
            ...item,
            styleName: item.styleName || best.productName || null,
            color: item.color || best.colour || null,
            colorName: item.colorName || best.colourName || null,
            sku: item.sku || best.productExternalRef || null,
            factory: item.factory || best.factory || null,
            customer: item.customer || best.customerName || null,
            season: item.season || best.season || null,
            unitCost: item.unitCost || cost,
            poNumber: item.poNumber || best.poNumber || null,
        };
    });
}

const BUY_FILE_KEYWORDS = [
    'style', 'style number', 'style no', 'style #', 'article', 'model',
    'po', 'po number', 'po no', 'po#', 'order', 'purchase order',
    'quantity', 'qty', 'units',
    'color', 'colour', 'color code', 'colour code', 'option', 'colorway',
    'size', 'size name', 'size scale', 'size desc',
    'sku', 'upc', 'ean', 'product code', 'material',
    'factory', 'vendor', 'supplier', 'manufacturer',
    'customer', 'brand', 'buyer', 'sold to',
    'season', 'year', 'delivery', 'ex factory', 'ex-fty', 'ship date', 'crd', 'crdd',
    'unit cost', 'cost', 'price', 'fob', 'currency',
    'transport', 'ship via', 'freight',
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
    const t0 = Date.now();
    const { worksheets } = await readExcelFile(fileBuffer);
    console.log('[buy-file-extractor] worksheets found:', worksheets.length, `(${Date.now() - t0}ms)`);

    // Build product sheet map from buy file itself and any provided product sheet files
    const t1 = Date.now();
    const productSheetMap = await buildProductSheetMap([fileBuffer, ...productSheetBuffers]);
    console.log('[buy-file-extractor] product sheet map keys:', Object.keys(productSheetMap).length, `(${Date.now() - t1}ms)`);

    // 1. Find the best worksheet: score all sheets by buy-file header keywords
    const sheetCandidates: { sheet: typeof worksheets[0]; headerRow: number; headers: string[]; score: number }[] = [];

    for (const sheet of worksheets) {
        const preview = sheet.rows.slice(0, 10);
        if (!preview.length) continue;

        const detected = await detectHeaderRow(preview);
        let headerRowIndex = detected.headerRow;
        let headerRow = sheet.rows[headerRowIndex - 1] || [];
        let headers = headerRow.map((h) => String(h || '')).filter(Boolean);

        // Multi-row header support: if headerRowCount > 1, merge sub-header rows
        // into the main header by concatenating (e.g., "PRODUCT INFO" + "Style" → "PRODUCT INFO Style")
        const headerRowCount = detected.headerRowCount || 1;
        if (headerRowCount > 1) {
            for (let r = 0; r < headerRowCount; r++) {
                const subRow = sheet.rows[headerRowIndex - 1 + r] || [];
                if (r === 0) continue; // first row already captured
                // Merge: if a column has a value in the sub-row but not in the main row, use it
                // If both have values, concatenate with space
                for (let c = 0; c < subRow.length; c++) {
                    const subVal = String(subRow[c] || '').trim();
                    if (!subVal) continue;
                    const mainVal = String(headerRow[c] || '').trim();
                    if (!mainVal) {
                        headerRow[c] = subVal;
                    } else {
                        headerRow[c] = `${mainVal} ${subVal}`;
                    }
                }
            }
            headers = headerRow.map((h) => String(h || '')).filter(Boolean);
            console.log(`[buy-file-extractor] merged ${headerRowCount}-row header:`, JSON.stringify(headers.slice(0, 10)));
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

async function extractFromSheet(
    sheet: { name: string; rows: (string | number | Date | null)[][] },
    headerRow: number,
    headers: string[],
    customerHint?: string,
    sharedNextgenClient?: NextGenCachedClient,
    productSheetMap: Record<string, ProductSheetRow[]> = {}
): Promise<BuyFileExtractionResult> {
    // NOTE: This function returns a single-sheet extraction result. productData is
    // built from rows within this sheet only.

    let mapping: ColumnMapping | null = null;
    let unmappedColumns: string[] = [];
    let templateUsed = false;

    // 2. Check for learned template

    console.log('[buy-file-extractor] sheet headers:', JSON.stringify(headers));

    const t2 = Date.now();
    const existingTemplate = await findMatchingTemplateSupabase(headers);
    console.log(`[buy-file-extractor] template lookup: ${Date.now() - t2}ms`);
    if (existingTemplate) {
        console.log('[buy-file-extractor] using existing template', existingTemplate.id);
        mapping = existingTemplate.mapping;
        templateUsed = true;
        // Merge in any deterministic pattern matches that the template might be missing
        // (e.g. new columns added to the buy file after the template was saved)
        const heuristicMapping = fallbackHeuristicMapping(headers);
        const beforeCount = Object.keys(mapping).length;
        mapping = mergeMappings(headers, mapping, heuristicMapping);
        const addedCount = Object.keys(mapping).length - beforeCount;
        if (addedCount > 0) {
            console.log(`[buy-file-extractor] merged ${addedCount} additional pattern matches into template mapping`);
        }
    } else {
        // 3. Build mapping from legacy DB + AI + heuristic fallback
        const t3 = Date.now();
        const legacyMapping = await loadLegacyMapping(customerHint);
        console.log(`[buy-file-extractor] legacy mapping: ${Date.now() - t3}ms`);
        const t4 = Date.now();
        const aiMappingResult = await mapHeaders(headers);
        console.log(`[buy-file-extractor] AI mapping: ${Date.now() - t4}ms`);

        mapping = mergeMappings(headers, legacyMapping, aiMappingResult.mapping);
        unmappedColumns = headers.filter((h) => !Object.values(mapping as Record<string, string>).includes(h));
        templateUsed = false;
    }

    console.log('[buy-file-extractor] mapping:', JSON.stringify(mapping));
    console.log('[buy-file-extractor] unmappedColumns:', JSON.stringify(unmappedColumns));

    // 4. Read all rows locally
    console.log('[buy-file-extractor] reading all rows locally');
    const t5 = Date.now();
    let items = readAllRows(sheet.rows, headerRow, mapping, sheet.name);
    console.log('[buy-file-extractor] extracted items:', items.length, `(${Date.now() - t5}ms)`);

    // 4.5 Enrich with product sheet data if available
    if (Object.keys(productSheetMap).length) {
        items = enrichItemsWithProductSheet(items, productSheetMap);
        console.log('[buy-file-extractor] enriched items with product sheet:', items.length);
    }

    // 4.6 Resolve color names from color codes using brand-specific mappings
    let colorsResolved = 0;
    for (const item of items) {
        if (!item.colorName) {
            // Priority 1: use the color name from the buy file (COLORWAY NAME)
            // since it's the authoritative source from the brand.
            if (item.color) {
                item.colorName = item.color;
                colorsResolved++;
                continue;
            }
            // Priority 2: resolve from color code using brand-specific map.
            // The colorCode field may be a composite "style+color" (e.g.,
            // "NF0A8CGZE8J" = style "NF0A8CGZ" + color "E8J"). Strip the style
            // prefix to get the actual color code for brand map lookup.
            if (item.colorCode) {
                const brand = detectBrand(item.style || '', item.customer || undefined);
                let rawCode = item.colorCode;
                const styleUpper = (item.style || '').toUpperCase();
                if (styleUpper && rawCode.toUpperCase().startsWith(styleUpper)) {
                    rawCode = rawCode.slice(styleUpper.length);
                }
                const resolved = resolveColorName(rawCode, brand);
                if (resolved) {
                    item.colorName = resolved;
                    colorsResolved++;
                }
            }
        }
    }
    if (colorsResolved > 0) {
        console.log(`[buy-file-extractor] resolved ${colorsResolved} color names from color codes`);
    }

    // 5. Query NextGen for unique styles (if enabled)
    const uniqueStyles = [...new Set(items.map((item) => item.style || '').filter(Boolean))];
    console.log('[buy-file-extractor] unique styles:', uniqueStyles.length);

    const nextgenEnabled = process.env.NEXTGEN_ENABLED !== 'false';
    let nextgenInfo: Record<string, NextGenStyleInfo | null> = {};

    if (nextgenEnabled) {
        const t6 = Date.now();
        const nextgenClient = sharedNextgenClient || new NextGenCachedClient();
        nextgenInfo = await nextgenClient.searchStyles(uniqueStyles);
        console.log(`[buy-file-extractor] NextGen searchStyles: ${Date.now() - t6}ms`);
    } else {
        console.log('[buy-file-extractor] NextGen disabled, skipping style lookups');
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
    sheetName: string
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

    // ── Size grid detection ──────────────────────────────────────────
    // Some buy files use a size matrix: columns like XS, S, M, L, XL, 2XL
    // instead of a single "size" + "quantity" pair. Detect this by checking
    // if any unmapped columns match common size names.
    const STANDARD_SIZES = ['xs', 's', 'm', 'l', 'xl', 'xxl', '2xl', '3xl', 'os', 'one size',
        'xxs', 'xxs', '4xl', '5xl', '6xl', 'xxxl', 'xxxl'];
    const sizeGridColumns: { header: string; index: number }[] = [];
    const mappedHeaders = new Set(Object.values(mapping).map((h) => String(h)));
    for (let c = 0; c < headerRow.length; c++) {
        const h = String(headerRow[c] || '').trim().toLowerCase();
        if (!h || mappedHeaders.has(String(headerRow[c] || ''))) continue;
        if (STANDARD_SIZES.includes(h) || /^\d{1,2}$/.test(h) || /^(x{0,3}s|x{0,3}l|xxl|2xl|3xl|4xl|5xl|6xl)$/i.test(h)) {
            sizeGridColumns.push({ header: String(headerRow[c] || ''), index: c });
        }
    }

    const isSizeGrid = sizeGridColumns.length >= 2 && !mapping.size;
    if (isSizeGrid) {
        console.log(`[buy-file-extractor] detected size grid matrix with ${sizeGridColumns.length} size columns:`,
            sizeGridColumns.map((s) => s.header).join(', '));
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

        // Read a column value by its raw header name (for unmapped columns like upcharges)
        const getRaw = (rawHeader: string): string | null => {
            let idx = headerToIndex[rawHeader];
            if (idx === undefined || idx < 0) {
                idx = normalizedHeaderToIndex[normalizeHeaderName(rawHeader)];
            }
            if (idx === undefined || idx < 0) return null;
            const val = row[idx];
            if (val === null || val === undefined) return null;
            return String(val as string).trim();
        };

        const unitCostStr = get('unit_cost');
        let unitCost = unitCostStr ? Number(unitCostStr.replace(/,/g, '')) : null;

        // If unit_cost (FOB) is empty, try computing from upcharges
        // FOB = production upcharges + material upcharges (in the file's currency)
        if (!unitCost || isNaN(unitCost)) {
            const prodUpchargeStr = getRaw('PRODUCTION UPCHARGES USD$') || getRaw('PRODUCTION UPCHARGES RMB$');
            const matUpchargeStr = getRaw('MATERIAL UPCHARGES USD$') || getRaw('MATERIAL UPCHARGES RMB$');
            const prodUpcharge = prodUpchargeStr ? Number(prodUpchargeStr.replace(/,/g, '')) : 0;
            const matUpcharge = matUpchargeStr ? Number(matUpchargeStr.replace(/,/g, '')) : 0;
            if ((prodUpcharge || matUpcharge) && !isNaN(prodUpcharge) && !isNaN(matUpcharge)) {
                unitCost = prodUpcharge + matUpcharge;
            }
        }

        const style = get('buyer_style_number');

        if (isSizeGrid && style) {
            // Unpivot: create one item per size column that has a quantity
            for (const sizeCol of sizeGridColumns) {
                const qtyVal = row[sizeCol.index];
                const qty = qtyVal ? Number(String(qtyVal).replace(/,/g, '')) : null;
                if (!qty || isNaN(qty) || qty <= 0) continue;

                items.push({
                    style,
                    styleName: get('buyer_style_name'),
                    sku: get('sku'),
                    description: get('product_description'),
                    color: get('color'),
                    colorCode: get('color_code'),
                    colorName: null,
                    size: sizeCol.header,
                    quantity: qty,
                    deliveryDate: get('delivery_date'),
                    season: get('season'),
                    customer: get('customer'),
                    factory: get('factory'),
                    currency: get('currency') || 'USD',
                    unitCost: unitCost && !isNaN(unitCost) ? unitCost : null,
                    poNumber: get('po_number'),
                    buyerPoNumber: get('buyer_po_number'),
                    startDate: get('start_date'),
                    cancelDate: get('cancel_date'),
                    transportMethod: get('transport_method'),
                    buyInformation: get('buy_information'),
                    sourceSheet: sheetName,
                    sourceRow: i + 1,
                });
            }
            continue;
        }

        const qtyStr = get('quantity');
        const qty = qtyStr ? Number(qtyStr.replace(/,/g, '')) : null;
        const quantity = qty && !isNaN(qty) ? qty : null;
        if (!style && quantity === null) continue;

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
            factory: get('factory'),
            currency: get('currency') || 'USD',
            unitCost: unitCost && !isNaN(unitCost) ? unitCost : null,
            poNumber: get('po_number'),
            buyerPoNumber: get('buyer_po_number'),
            startDate: get('start_date'),
            cancelDate: get('cancel_date'),
            transportMethod: get('transport_method'),
            buyInformation: get('buy_information'),
            sourceSheet: sheetName,
            sourceRow: i + 1,
        });
    }

    return items;
}
