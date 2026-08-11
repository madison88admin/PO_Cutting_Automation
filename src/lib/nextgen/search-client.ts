import { NextGenClient } from '@/lib/nextgen';
import { NextGenStyleInfo } from '@/lib/types/buy-file';

const SEARCH_BASE_URL = process.env.NEXTGEN_SEARCH_BASE_URL || process.env.NEXTGEN_BASE_URL || 'https://nextgen.madison88.com';
const SEARCH_ENTITY_TYPES = process.env.NEXTGEN_SEARCH_ENTITY_TYPES || '0,5,6,138,80,121,9,222,163,69,23,41,139,42';

function normalizeKey(s: string): string {
    return String(s || '')
        .toLowerCase()
        .replace(/[^a-z0-9]/g, '')
        .trim();
}

function stripStylePrefix(style: string): string {
    const custom = (process.env.NEXTGEN_STYLE_PREFIX_STRIP || '').trim();
    let cleaned = style
        .replace(/\s*\([^)]*\)\s*/g, '')
        .trim();
    if (custom) {
        return cleaned.replace(new RegExp('^' + custom, 'i'), '').trim();
    }
    return cleaned
        .replace(/^(NF00|NF0|NF)/i, '')
        .trim();
}

interface SearchResult {
    Name: string;
    Id: number;
    ParentId: number;
    ParentEntityName: string | null;
    EntityType: number;
    SearchType: number;
    ExactMatch: boolean;
    FieldName: string | null;
    FieldValue: string | null;
    RangeDisplayName: string | null;
}

export class NextGenSearchClient {
    private base: NextGenClient;

    constructor(sharedBase?: NextGenClient) {
        this.base = sharedBase || new NextGenClient();
    }

    async searchStyle(style: string): Promise<NextGenStyleInfo | null> {
        const targetStyle = style.trim();
        if (!targetStyle) return null;

        try {
            await this.base.login();

            const entityTypes = SEARCH_ENTITY_TYPES.split(',').map((s) => s.trim()).filter(Boolean);
            const queries = entityTypes.map((et) => `searchEntityTypes=${et}`).join('&');

            const attempts = [targetStyle, stripStylePrefix(targetStyle)];
            const seen = new Set<string>();

            for (const term of attempts) {
                if (!term || seen.has(term)) continue;
                seen.add(term);

                const url = `${SEARCH_BASE_URL}/Search/GetSearchResults?criteria=${encodeURIComponent(term)}&${queries}`;
                console.log(`[nextgen-search] calling ${url}`);
                const response = await this.base.fetchWithCookie(url, { method: 'GET' }, true);

                const text = await response.text();
                console.log(`[nextgen-search] ${term} status: ${response.status}, body:`, text.slice(0, 500));

                if (response.status === 401 || response.status === 403) {
                    throw new Error(`NextGen search auth failed: ${response.status}`);
                }
                if (response.status === 302 && text.includes('/Account/Login')) {
                    throw new Error('NextGen search session expired');
                }
                if (!response.ok) continue;

                const data = text ? JSON.parse(text) : null;
                const info = this.mapSearchResponse(style, data);
                if (info) return info;
            }

            return null;
        } catch (err) {
            const message = err instanceof Error ? err.message : String(err);
            if (message.includes('auth') || message.includes('session') || message.includes('login')) {
                throw err;
            }
            console.warn(`[nextgen-search] failed for ${targetStyle}:`, message);
            return null;
        }
    }

    async searchStyles(styles: string[]): Promise<Record<string, NextGenStyleInfo | null>> {
        const unique = [...new Set(styles.map((s) => s.trim()).filter(Boolean))];
        const out: Record<string, NextGenStyleInfo | null> = {};
        if (!unique.length) return out;

        const results = await Promise.all(
            unique.map(async (style) => ({
                style,
                info: await this.searchStyle(style),
            }))
        );

        for (const { style, info } of results) {
            out[style] = info;
        }
        return out;
    }

    private mapSearchResponse(style: string, data: any): NextGenStyleInfo | null {
        if (!data || typeof data !== 'object') return null;
        const results = this.collectResults(data);
        if (!results.length) return null;

        const byType = this.groupByEntityType(results);
        const products = byType['5'] || [];
        const colors = byType['6'] || byType['138'] || [];
        const sizes = byType['9'] || byType['80'] || [];
        const pos = byType['163'] || [];

        const product = products.find((r) => r.ExactMatch) || products[0] || results.find((r) => r.ExactMatch) || results[0];
        const color = colors[0];
        const size = sizes[0];
        const po = pos[0];

        const season = this.parseSeason(product?.RangeDisplayName || '');

        return {
            style,
            product: product?.Name || null,
            productRange: product?.RangeDisplayName || null,
            productExternalRef: product?.FieldName && product.FieldName.toLowerCase().includes('buyer') ? product.FieldValue || style : style,
            productCustomerRef: style,
            styleName: null,
            brand: null,
            season: season || null,
            department: null,
            colorName: null,
            colorCode: null,
            colorExt: null,
            sizeScale: sizes.map((s) => s.Name).join(', ') || null,
            purchaseUOM: 'PCS',
            sellingUOM: 'PCS',
            supplierProfile: null,
            customer: this.extractCustomer(po?.Name || '') || null,
            factory: null,
            currency: null,
        };
    }

    private collectResults(data: any): SearchResult[] {
        const results: SearchResult[] = [];
        const nameResults = data.nameResults || {};
        for (const entityType of Object.values(nameResults)) {
            for (const searchType of Object.values(entityType as any)) {
                if (Array.isArray(searchType)) {
                    results.push(...searchType);
                }
            }
        }
        return results;
    }

    private groupByEntityType(results: SearchResult[]): Record<string, SearchResult[]> {
        return results.reduce((acc, r) => {
            const key = String(r.EntityType);
            if (!acc[key]) acc[key] = [];
            acc[key].push(r);
            return acc;
        }, {} as Record<string, SearchResult[]>);
    }

    private parseSeason(range: string): string | null {
        if (!range) return null;
        const match = range.match(/\b(S\d{2,4}|F?[WS]\d{2,4}|\d{4})\b/i);
        return match ? match[1].toUpperCase() : null;
    }

    private extractCustomer(name: string): string | null {
        if (!name) return null;
        const match = name.match(/-\s*([^-]+)$/);
        return match ? match[1].trim() : null;
    }
}

export function normalizeForSearch(s: string): string {
    return normalizeKey(s);
}

export function findBestStyleMatch(style: string, rows: any[]): { row: any; field: string } | null {
    const target = normalizeKey(style);
    const fields = [
        'style', 'Style', 'styleNumber', 'StyleNumber', 'buyerStyleNumber', 'BuyerStyleNumber',
        'productCode', 'ProductCode', 'product', 'Product', 'commodityName', 'CommodityName',
        'productExternalRef', 'ProductExternalRef', 'customerRef', 'CustomerRef', 'sku', 'SKU',
    ];

    const hits: Record<string, number> = {};
    for (const row of rows) {
        for (const field of fields) {
            const val = normalizeKey(String(row[field] || ''));
            if (val && (val === target || val.includes(target))) {
                hits[field] = (hits[field] || 0) + 1;
            }
        }
    }

    const best = Object.entries(hits).sort((a, b) => b[1] - a[1])[0];
    if (!best) return null;
    const bestField = best[0];
    const row = rows.find((r) => {
        const val = normalizeKey(String(r[bestField] || ''));
        return val && (val === target || val.includes(target));
    });
    return row ? { row, field: bestField } : null;
}

export function normalizeColorName(s: string): string {
    return String(s || '')
        .toLowerCase()
        .replace(/[^a-z0-9]/g, '')
        .trim();
}

export function colorsMatch(a: string, b: string): boolean {
    const na = normalizeColorName(a);
    const nb = normalizeColorName(b);
    if (!na || !nb) return false;
    return na === nb || na.includes(nb) || nb.includes(na);
}

export function sizesMatch(a: string, b: string): boolean {
    return normalizeKey(a) === normalizeKey(b);
}
