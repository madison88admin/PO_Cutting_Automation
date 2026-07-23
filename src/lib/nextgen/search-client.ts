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

function buildStyleSearchTerms(style: string): string[] {
    const original = String(style || '').trim();
    const withoutNotes = original.replace(/\s*\([^)]*\)\s*/g, '').trim();
    const compact = withoutNotes.replace(/[^a-z0-9]/gi, '');
    const configured = stripStylePrefix(withoutNotes);
    return [...new Set([original, withoutNotes, compact, configured].filter(Boolean))];
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
    private variantCatalogs = new Map<string, Promise<Array<{ product: SearchResult; options: any[] }>>>();

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

    async searchVariant(style: string, colorHint: string): Promise<NextGenStyleInfo | null> {
        const targetStyle = style.trim();
        const colorCode = this.extractColorCode(style, colorHint);
        if (!targetStyle || !String(colorHint || '').trim()) return this.searchStyle(style);

        const catalog = await this.getVariantCatalog(targetStyle);
        const variantMatches = catalog.flatMap(({ product, options }) => {
            const ranked = options
                .map((option) => ({
                    option,
                    score: this.optionMatchScore(option, colorHint, colorCode),
                }))
                .filter(({ score }) => score >= 55)
                .sort((a, b) => b.score - a.score);
            return ranked[0] ? [{
                product,
                ...ranked[0],
                ambiguousWithinProduct: Boolean(
                    ranked[1]
                    && ranked[1].score === ranked[0].score
                    && String(ranked[1].option?.ColourName || '') !== String(ranked[0].option?.ColourName || '')
                ),
            }] : [];
        });

        if (!variantMatches.length) return this.searchStyle(style);

        // Prefer the strongest colour match. If the same buyer style exists in
        // multiple historical products, use the newest product with that match.
        variantMatches.sort((a, b) => b.score - a.score || b.product.Id - a.product.Id);
        const selected = variantMatches[0];
        const isAmbiguous = selected.ambiguousWithinProduct;
        return {
            style,
            product: selected.product.Name,
            productRange: selected.product.RangeDisplayName || null,
            productExternalRef: style,
            productCustomerRef: selected.product.FieldValue || style,
            styleName: null,
            brand: selected.option.CustomerName || null,
            season: this.parseSeason(selected.option.ColourDescription || selected.product.RangeDisplayName || '') || null,
            department: selected.option.DepartmentName || null,
            colorName: selected.option.ColourName || null,
            colorCode,
            colorExt: selected.option.ColourExternalRef || null,
            sizeScale: null,
            purchaseUOM: 'PCS',
            sellingUOM: 'PCS',
            supplierProfile: null,
            customer: selected.option.CustomerName || null,
            factory: null,
            currency: null,
            matchStatus: isAmbiguous ? 'ambiguous' : 'matched',
            matchScore: selected.score,
            matchReason: isAmbiguous
                ? 'Multiple Nexgen colours have the same best score for this product.'
                : (colorCode ? 'Matched by Nexgen colour code.' : 'Matched by Nexgen colour description.'),
            candidateCount: variantMatches.length,
        };
    }

    private getVariantCatalog(style: string): Promise<Array<{ product: SearchResult; options: any[] }>> {
        const cacheKey = normalizeKey(style);
        const existing = this.variantCatalogs.get(cacheKey);
        if (existing) return existing;
        const promise = this.loadVariantCatalog(style);
        this.variantCatalogs.set(cacheKey, promise);
        return promise;
    }

    private async loadVariantCatalog(targetStyle: string): Promise<Array<{ product: SearchResult; options: any[] }>> {
        await this.base.login();
        // Search the buyer value as supplied first. Compact/configured variants
        // are fallbacks, so a new brand does not need a hard-coded prefix rule.
        const initialResultSets: SearchResult[][] = [];
        for (const term of buildStyleSearchTerms(targetStyle)) {
            const results = await this.fetchGlobalResults(term);
            initialResultSets.push(results);
            if ((this.groupByEntityType(results)['5'] || []).some((row) => row.ExactMatch)) break;
        }
        const firstResults = initialResultSets.flat();
        const firstProducts = this.groupByEntityType(firstResults)['5'] || [];

        // Nexgen's Buyer Style Number is the authoritative bridge between an
        // unfamiliar buyer format and every related Nexgen product revision.
        const expandedStyles = [...new Set(
            buildStyleSearchTerms(targetStyle)
                .map((term) => this.pickExpandedBuyerStyle(term, firstProducts))
                .filter((value): value is string => Boolean(value))
        )];
        const expandedResults = (await Promise.all(
            expandedStyles.map((expandedStyle) => this.fetchGlobalResults(expandedStyle))
        )).flat();
        const products = this.uniqueProducts([
            ...firstProducts,
            ...(this.groupByEntityType(expandedResults)['5'] || []),
        ]).filter((product) => this.productMatchesRequestedStyle(product, targetStyle));

        return Promise.all(products.map(async (product) => ({
            product,
            options: await this.fetchProductOptions(product.Id),
        })));
    }

    private async fetchGlobalResults(term: string): Promise<SearchResult[]> {
        const entityTypes = SEARCH_ENTITY_TYPES.split(',').map((s) => s.trim()).filter(Boolean);
        const queries = entityTypes.map((et) => `searchEntityTypes=${et}`).join('&');
        const url = `${SEARCH_BASE_URL}/Search/GetSearchResults?criteria=${encodeURIComponent(term)}&${queries}`;
        const response = await this.base.fetchWithCookie(url, { method: 'GET' }, true);
        if (!response.ok) return [];
        const data = await response.json();
        return this.collectResults(data);
    }

    private pickExpandedBuyerStyle(style: string, products: SearchResult[]): string | null {
        const candidates = products
            .filter((product) => /buyer style number/i.test(product.FieldName || ''))
            .map((product) => String(product.FieldValue || '').trim())
            .filter((value) => normalizeKey(value).startsWith(normalizeKey(style)) && normalizeKey(value) !== normalizeKey(style));
        if (!candidates.length) return null;
        const counts = candidates.reduce((map, value) => map.set(value, (map.get(value) || 0) + 1), new Map<string, number>());
        return [...counts.entries()].sort((a, b) => b[1] - a[1])[0][0];
    }

    private uniqueProducts(products: SearchResult[]): SearchResult[] {
        const unique = new Map<number, SearchResult>();
        for (const product of products.filter((result) => result.EntityType === 5)) {
            unique.set(product.Id, product);
        }
        return [...unique.values()];
    }

    private productMatchesRequestedStyle(product: SearchResult, requestedStyle: string): boolean {
        const terms = buildStyleSearchTerms(requestedStyle)
            .map(normalizeKey)
            .filter((term) => term.length >= 4);
        const productName = normalizeKey(product.Name);
        const buyerReference = normalizeKey(product.FieldValue || '');

        // A colour match alone is not enough. The Nexgen product must also
        // belong to the requested buyer-style root (or be the exact M product).
        return terms.some((term) =>
            productName === term
            || buyerReference === term
            || buyerReference.startsWith(term)
        );
    }

    private async fetchProductOptions(productId: number): Promise<any[]> {
        const body = new URLSearchParams({
            commodityId: String(productId),
            page: '1',
            pageSize: '100',
            sort: '',
            group: '',
            filter: '',
            aggregates: '',
        });
        const response = await this.base.fetchWithCookie(
            `${SEARCH_BASE_URL}/ProductOption/ProductOptionsGridRead`,
            {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' },
                body: body.toString(),
            },
            true
        );
        if (!response.ok) return [];
        const data = await response.json();
        return Array.isArray(data?.Data) ? data.Data : [];
    }

    private extractColorCode(style: string, value: string): string {
        const text = String(value || '').toUpperCase();
        const named = text.match(/\bTNF[\s-]+([A-Z0-9]{3})\b/);
        if (named) return named[1];
        const compactStyle = String(style || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
        const compactValue = text.replace(/[^A-Z0-9]/g, '');
        if (compactStyle && compactValue.startsWith(compactStyle)) {
            return compactValue.slice(compactStyle.length, compactStyle.length + 3);
        }
        return '';
    }

    private optionMatchScore(option: any, hint: string, code: string): number {
        const optionText = [
            option?.ColourName,
            option?.ColourDescription,
            option?.ColourExternalRef,
            option?.ColourCode,
        ].filter(Boolean).join(' ');
        const upperOption = optionText.toUpperCase();

        if (code && new RegExp(`(?:^|[^A-Z0-9])${code}(?:[^A-Z0-9]|$)`).test(upperOption)) {
            return 100;
        }

        const normalizedHint = normalizeKey(hint);
        const normalizedOption = normalizeKey(optionText);
        if (normalizedHint.length >= 3 && normalizedOption === normalizedHint) return 95;
        if (
            normalizedHint.length >= 4
            && (normalizedOption.includes(normalizedHint) || normalizedHint.includes(normalizedOption))
        ) return 85;

        // Generic fallback for brands whose buyer files contain colour words
        // rather than compact colour codes (for example "Citron and Oasis").
        const ignored = new Set(['and', 'with', 'color', 'colour', 'combo', 'the']);
        const tokens = String(hint || '')
            .toLowerCase()
            .split(/[^a-z0-9]+/)
            .filter((token) => token.length >= 3 && !ignored.has(token) && !/^\d+$/.test(token));
        const uniqueTokens = [...new Set(tokens)];
        if (!uniqueTokens.length) return 0;
        const matched = uniqueTokens.filter((token) => normalizedOption.includes(normalizeKey(token))).length;
        const ratio = matched / uniqueTokens.length;
        return matched > 0 && ratio >= 0.75 ? 55 + Math.round(ratio * 25) : 0;
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
