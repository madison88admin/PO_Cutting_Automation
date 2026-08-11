import { NextGenClient as BaseNextGenClient } from '@/lib/nextgen';
import { NextGenCache } from './cache';
import { NextGenStyleInfo } from '@/lib/types/buy-file';
import { NextGenSearchClient } from './search-client';

const SEARCH_FIELDS = [
    'CommodityName',
    'ProductCode',
    'Product',
    'Material',
    'ProductExternalRef',
    'ExternalRef',
    'ProductCustomerRef',
    'CustomerRef',
    'Style',
    'StyleNumber',
    'BuyerStyleNumber',
    'SKU',
    'ItemNumber',
];

function normalizeKey(s: string): string {
    return String(s || '')
        .toLowerCase()
        .replace(/[^a-z0-9]/g, '')
        .trim();
}

function normalizeColor(s: string): string {
    return String(s || '')
        .toLowerCase()
        .replace(/[^a-z0-9]/g, '')
        .trim();
}

export interface NextGenSearchDiagnostics {
    recordsLoaded: number;
    searchFields: Record<string, number>;
    styleHits: Record<string, number>;
    colorHits: Record<string, number>;
}

export class NextGenCachedClient {
    private base: BaseNextGenClient;
    private search: NextGenSearchClient;
    private cache: NextGenCache;
    private records: any[] | null = null;

    constructor() {
        this.base = new BaseNextGenClient();
        this.search = new NextGenSearchClient(this.base);
        this.cache = new NextGenCache();
    }

    // Pre-load records so all lookups share a single PurchaseOrder/Read call (fallback path)
    async loadRecords(pageSize: number = 500): Promise<any[]> {
        if (!this.records) {
            this.records = await this.base.fetchRecentRecords(pageSize);
        }
        return this.records;
    }

    setRecords(records: any[]) {
        this.records = records;
    }

    async getDiagnostics(style: string, color?: string): Promise<NextGenSearchDiagnostics> {
        const records = await this.loadRecords();
        const targetStyle = normalizeKey(style);
        const targetColor = color ? normalizeColor(color) : '';
        const searchFields: Record<string, number> = {};
        const styleHits: Record<string, number> = {};
        const colorHits: Record<string, number> = {};

        for (const row of records) {
            for (const field of SEARCH_FIELDS) {
                const val = normalizeKey(String(row[field] || ''));
                if (!val) continue;
                if (val === targetStyle || val.includes(targetStyle)) {
                    searchFields[field] = (searchFields[field] || 0) + 1;
                    styleHits[style] = (styleHits[style] || 0) + 1;
                }
            }
            if (targetColor) {
                const rowColor = normalizeColor(String(row.OptionColourName || row.ColourName || row.ColorName || ''));
                if (rowColor && (rowColor === targetColor || rowColor.includes(targetColor) || targetColor.includes(rowColor))) {
                    colorHits[color!] = (colorHits[color!] || 0) + 1;
                }
            }
        }

        return { recordsLoaded: records.length, searchFields, styleHits, colorHits };
    }

    async searchStyle(style: string): Promise<NextGenStyleInfo | null> {
        if (this.cache.hasStyle(style)) {
            return this.cache.getStyle(style) || null;
        }

        try {
            // Primary: new dedicated search endpoint
            let info = await this.search.searchStyle(style);
            if (info) {
                console.log(`[nextgen-client] search endpoint matched style "${style}"`);
                this.cache.setStyle(style, info);
                return info;
            }
        } catch (err) {
            const message = err instanceof Error ? err.message : String(err);
            if (message.includes('auth') || message.includes('session') || message.includes('login')) {
                console.warn(`[nextgen-client] search endpoint auth error for "${style}", skipping fallback`);
                this.cache.setStyle(style, { style });
                return null;
            }
            console.warn(`[nextgen-client] search endpoint failed for "${style}", falling back`, err);
        }

        // Fallback: scan recent PurchaseOrder records
        return this.searchStyleFallback(style);
    }

    async searchStyleFallback(style: string): Promise<NextGenStyleInfo | null> {
        if (this.cache.hasStyle(style)) {
            return this.cache.getStyle(style) || null;
        }

        try {
            const results = await this.loadRecords();
            const targetStyle = normalizeKey(style);

            const fieldHits: Record<string, number> = {};
            for (const row of results) {
                for (const field of SEARCH_FIELDS) {
                    const val = normalizeKey(String(row[field] || ''));
                    if (val && (val === targetStyle || val.includes(targetStyle))) {
                        fieldHits[field] = (fieldHits[field] || 0) + 1;
                    }
                }
            }
            console.log(`[nextgen-client] fallback diagnostics for style "${style}":`, fieldHits);

            const bestField = Object.entries(fieldHits).sort((a, b) => b[1] - a[1])[0]?.[0];
            let matchingRows: any[] = [];
            if (bestField) {
                matchingRows = results.filter((row: any) => {
                    const val = normalizeKey(String(row[bestField] || ''));
                    return val && (val === targetStyle || val.includes(targetStyle));
                });
            }

            if (!matchingRows.length) {
                console.warn(`[nextgen-client] no NextGen match for style "${style}"`);
                this.cache.setStyle(style, { style });
                return null;
            }

            const firstRow = matchingRows.reduce((best, row) => {
                const populated = SEARCH_FIELDS.filter((f) => String(row[f] || '').trim()).length;
                const bestPopulated = SEARCH_FIELDS.filter((f) => String(best[f] || '').trim()).length;
                return populated > bestPopulated ? row : best;
            }, matchingRows[0]);

            const line = this.base.mapToPOLine(firstRow);
            const info: NextGenStyleInfo = {
                style: line.style || style,
                product: String(firstRow.Product || firstRow.ProductCode || firstRow.Material || line.style || ''),
                productRange: String(firstRow.ProductRange || firstRow.Range || ''),
                productExternalRef: String(firstRow.ProductExternalRef || firstRow.ExternalRef || ''),
                productCustomerRef: String(firstRow.ProductCustomerRef || firstRow.CustomerRef || ''),
                styleName: String(firstRow.StyleName || firstRow.StyleDescription || ''),
                brand: String(firstRow.Brand || firstRow.CustomerName || ''),
                season: line.season || '',
                department: String(firstRow.Department || firstRow.ProductDivision || ''),
                colorName: line.color || '',
                colorCode: String(firstRow.ColorCode || firstRow.OptionColourCode || ''),
                colorExt: String(firstRow.ColourExt || firstRow.ColorExt || ''),
                sizeScale: String(firstRow.SizeScale || firstRow.SizeRange || ''),
                purchaseUOM: String(firstRow.PurchaseUOM || firstRow.UOM || 'PCS'),
                sellingUOM: String(firstRow.SellingUOM || firstRow.UOM || 'PCS'),
                supplierProfile: String(firstRow.SupplierProfile || firstRow.SupplierCode || ''),
                customer: String(firstRow.CustomerName || firstRow.Customer || ''),
                factory: line.factory || '',
                currency: String(firstRow.Currency || ''),
            };

            this.cache.setStyle(style, info);
            return info;
        } catch (err) {
            console.error(`[nextgen-client] searchStyleFallback failed for ${style}:`, err);
            this.cache.setStyle(style, { style });
            return null;
        }
    }

    async searchStyles(styles: string[]): Promise<Record<string, NextGenStyleInfo | null>> {
        const unique = [...new Set(styles.map((s) => s.trim()).filter(Boolean))];
        const out: Record<string, NextGenStyleInfo | null> = {};
        if (!unique.length) return out;

        // 1. Load from persistent Supabase cache first
        const { loadCachedStyles, saveCachedStyles } = await import('@/lib/nextgen/style-cache');
        const cached = await loadCachedStyles(unique);
        const uncached = unique.filter((s) => !(s.toLowerCase().trim() in cached));

        // Copy cached results
        for (const [style, info] of Object.entries(cached)) {
            out[style] = info;
        }

        if (!uncached.length) {
            console.log(`[nextgen-client] all ${unique.length} styles found in persistent cache`);
            return out;
        }

        // 2. Search NextGen for uncached styles (parallel, batched)
        const batchSize = Number(process.env.NEXTGEN_SEARCH_BATCH_SIZE || '5');
        const newResults: Record<string, NextGenStyleInfo | null> = {};
        for (let i = 0; i < uncached.length; i += batchSize) {
            const batch = uncached.slice(i, i + batchSize);
            const results = await Promise.all(
                batch.map(async (style) => {
                    try {
                        return [style, await this.searchStyle(style)] as const;
                    } catch (err) {
                        console.error(`[nextgen-client] searchStyle failed for ${style}:`, err);
                        return [style, null] as const;
                    }
                })
            );
            for (const [style, info] of results) {
                out[style] = info;
                newResults[style] = info;
            }
        }

        // 3. Save new results to persistent cache
        if (Object.keys(newResults).length > 0) {
            saveCachedStyles(newResults).catch(() => {}); // fire and forget
        }

        return out;
    }

    async getLatestPO(): Promise<{ poNumber: string; id: string } | null> {
        if (this.records && this.records.length) {
            const match = this.records[0];
            const poNumber = this.base.getPONumberFromRecord(match);
            return {
                id: String(match.OrderId || match.Id || match.ID || match.id || ''),
                poNumber,
            };
        }
        return this.base.getLatestPO();
    }

    async lookupColorNames(skus: string[]): Promise<Record<string, string | null>> {
        const records = await this.loadRecords();
        const targets = skus.map((s) => normalizeKey(s)).filter(Boolean);
        const found: Record<string, string | null> = {};
        if (!targets.length) return found;

        for (const row of records) {
            for (const [key, value] of Object.entries(row)) {
                const valStr = normalizeKey(String(value || ''));
                for (const target of targets) {
                    if (found[target] !== undefined) continue;
                    if (valStr === target || valStr.includes(target)) {
                        const colorName = String(row.OptionColourName || row.ColourName || row.ColorName || '');
                        found[target] = colorName || null;
                    }
                }
            }
        }

        for (const target of targets) {
            if (found[target] === undefined) found[target] = null;
        }
        return found;
    }
}
