import { NextGenClient } from '@/lib/nextgen';
import { NextGenStyleInfo } from '@/lib/types/buy-file';
import { findClosestStyleMatch, sizesEquivalent, colorsEquivalent } from '@/lib/fuzzy';
import { getCachedNextGenMatch, saveNextGenMatch, getUserCorrection, getCachedColorMapping } from '@/lib/learning/cache';

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
    // Use brand config for prefix stripping — supports all configured brands
    // Falls back to TNF defaults for backward compatibility
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

    async searchVariant(style: string, colorHint: string, brand?: string): Promise<NextGenStyleInfo | null> {
        const targetStyle = style.trim();
        const colorCode = this.extractColorCode(style, colorHint);
        if (!targetStyle || !String(colorHint || '').trim()) return this.searchStyle(style);

        // Defaults — may be enriched by color mapping cache
        let effectiveColorHint = colorHint;
        let effectiveColorCode = colorCode;

        // --- Learning Layer: check cache before hitting NextGen ---
        const brandKey = (brand || '').toLowerCase();
        if (brandKey) {
            // 1. Check user corrections first (highest priority)
            const correction = await getUserCorrection(targetStyle, colorHint, brandKey);
            if (correction) {
                console.log(`[learning-layer] Using user correction: ${targetStyle}/${colorHint} → ${correction.corrected_nextgen_product}`);
                return {
                    style,
                    product: correction.corrected_nextgen_product,
                    productRange: null,
                    productExternalRef: style,
                    productCustomerRef: style,
                    styleName: null,
                    brand: brandKey,
                    season: null,
                    department: null,
                    colorName: correction.corrected_color_name || null,
                    colorCode: correction.corrected_color_code || colorCode,
                    colorExt: null,
                    sizeScale: null,
                    purchaseUOM: 'PCS',
                    sellingUOM: 'PCS',
                    supplierProfile: null,
                    customer: null,
                    factory: null,
                    currency: null,
                    unitCost: null,
                    costingReference: null,
                    sellPrice: null,
                    matchStatus: 'matched',
                    matchScore: 100,
                    matchReason: 'User-corrected match (cached).',
                    candidateCount: 1,
                    candidates: [],
                };
            }

            // 2. Check color mapping cache to improve color hint
            const colorMapping = await getCachedColorMapping(brandKey, colorHint);
            if (colorMapping) {
                console.log(`[learning-layer] Color mapping hit: ${colorHint} → ${colorMapping.nextgen_color_name} (${colorMapping.nextgen_color_code})`);
                effectiveColorHint = colorMapping.nextgen_color_name || colorHint;
                effectiveColorCode = colorMapping.nextgen_color_code || colorCode;
            }

            // 3. Check NextGen match cache
            const cached = await getCachedNextGenMatch(targetStyle, colorHint, brandKey);
            if (cached) {
                console.log(`[learning-layer] Cache hit: ${targetStyle}/${colorHint} → ${cached.nextgen_product} (hits: ${cached.hit_count})`);
                return {
                    style,
                    product: cached.nextgen_product,
                    productRange: null,
                    productExternalRef: style,
                    productCustomerRef: style,
                    styleName: cached.nextgen_style_name || null,
                    brand: brandKey,
                    season: null,
                    department: null,
                    colorName: cached.nextgen_color_name || null,
                    colorCode: cached.nextgen_color_code || colorCode,
                    colorExt: null,
                    sizeScale: null,
                    purchaseUOM: 'PCS',
                    sellingUOM: 'PCS',
                    supplierProfile: null,
                    customer: null,
                    factory: cached.factory,
                    currency: cached.currency,
                    unitCost: cached.unit_cost,
                    costingReference: cached.costing_reference,
                    sellPrice: null,
                    matchStatus: 'matched',
                    matchScore: cached.match_score,
                    matchReason: `Cached match (hits: ${cached.hit_count}). ${cached.match_reason}`,
                    candidateCount: 1,
                    candidates: [],
                };
            }
        }
        // --- End Learning Layer cache check ---

        const catalog = await this.getVariantCatalog(targetStyle);
        const variantMatches = catalog.flatMap(({ product, options }) => {
            const ranked = options
                .map((option) => ({
                    option,
                    score: this.optionMatchScore(option, effectiveColorHint, effectiveColorCode),
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
        const runnerUp = variantMatches[1];
        const isAmbiguous = selected.ambiguousWithinProduct || Boolean(
            runnerUp
            && runnerUp.score === selected.score
            && String(runnerUp.product.Name || "") !== String(selected.product.Name || "")
        );
        const candidates = variantMatches
            .map(({ product, option, score }) => ({
                product: String(product.Name || ''),
                colorName: String(option.ColourName || ''),
                score,
                productRange: product.RangeDisplayName || null,
                productExternalRef: style,
                productCustomerRef: product.FieldValue || style,
                colorCode: colorCode || option.ColourCode || null,
                colorExt: option.ColourExternalRef || null,
                customer: option.CustomerName || null,
            }))
            .filter((candidate, index, rows) =>
                Boolean(candidate.product && candidate.colorName)
                && rows.findIndex((row) =>
                    row.product === candidate.product && row.colorName === candidate.colorName
                ) === index
            )
            .slice(0, 8);
        const result: NextGenStyleInfo = {
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
            // Extract factory/cost from NextGen option data — replaces product sheet
            factory: this.extractOptionField(selected.option, ['Factory', 'OrderSupplierName', 'Supplier', 'Vendor', 'SupplierName']) || null,
            currency: this.extractOptionField(selected.option, ['Currency', 'OrderCurrency', 'CurrencyCode']) || null,
            unitCost: this.extractOptionCost(selected.option),
            costingReference: this.extractOptionField(selected.option, ['CostingReference', 'CostReference', 'CostRef']) || null,
            sellPrice: this.extractOptionSellPrice(selected.option),
            matchStatus: isAmbiguous ? 'ambiguous' : 'matched',
            matchScore: selected.score,
            matchReason: isAmbiguous
                ? 'Multiple Nexgen product or colour records have the same best score.'
                : (colorCode ? 'Matched by Nexgen colour code.' : 'Matched by Nexgen colour description.'),
            candidateCount: variantMatches.length,
            candidates,
        };

        // --- Learning Layer: save successful match to cache ---
        if (brandKey && result.matchStatus === 'matched' && result.product) {
            try {
                await saveNextGenMatch(targetStyle, colorHint, brandKey, result);
            } catch (err) {
                console.warn('[learning-layer] Failed to save NextGen match:', err);
            }
        }
        // --- End Learning Layer ---

        return result;
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

    /**
     * Extract a string field from a NextGen option object by trying
     * multiple possible field names. NextGen option data is not strictly
     * typed, so we probe for common variants.
     */
    private extractOptionField(option: any, fieldNames: string[]): string | null {
        if (!option || typeof option !== 'object') return null;
        for (const name of fieldNames) {
            const val = option[name];
            if (val !== null && val !== undefined && String(val).trim()) {
                return String(val).trim();
            }
        }
        return null;
    }

    /**
     * Extract FOB/unit cost from a NextGen option object.
     * Probes common field names used by NextGen for cost data.
     */
    private extractOptionCost(option: any): number | null {
        if (!option || typeof option !== 'object') return null;
        const costFields = [
            'FOB', 'Fob', 'fob', 'UnitCost', 'unitCost', 'Cost', 'cost',
            'OrderUnitCost', 'PurchasePrice', 'LandedCost', 'FactoryCost',
            'PrimaryUserDefinedFieldValuesTextUdf3', 'UDF3',
        ];
        for (const field of costFields) {
            const val = option[field];
            if (val === null || val === undefined || val === '') continue;
            const num = typeof val === 'number' ? val : Number(String(val).replace(/[^0-9.-]/g, ''));
            if (Number.isFinite(num) && num > 0) return num;
        }
        return null;
    }

    /**
     * Extract sell price from a NextGen option object.
     */
    private extractOptionSellPrice(option: any): number | null {
        if (!option || typeof option !== 'object') return null;
        const priceFields = [
            'SellPrice', 'sellPrice', 'WholesalePrice', 'MSRP', 'RetailPrice',
            'Price', 'price', 'SellingPrice',
        ];
        for (const field of priceFields) {
            const val = option[field];
            if (val === null || val === undefined || val === '') continue;
            const num = typeof val === 'number' ? val : Number(String(val).replace(/[^0-9.-]/g, ''));
            if (Number.isFinite(num) && num > 0) return num;
        }
        return null;
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

    // Layer 1: exact + substring match (original logic)
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
    if (best) {
        const bestField = best[0];
        const row = rows.find((r) => {
            const val = normalizeKey(String(r[bestField] || ''));
            return val && (val === target || val.includes(target));
        });
        if (row) return { row, field: bestField };
    }

    // Layer 2: Levenshtein fallback — handles 1-2 char typos in style codes
    // Only triggers if exact/substring matching found nothing.
    if (target.length < 4) return null; // skip for very short codes (too many false positives)

    let bestMatch: { row: any; field: string; distance: number } | null = null;
    for (const row of rows) {
        for (const field of fields) {
            const val = String(row[field] || '').trim();
            if (!val || val.length < 4) continue;
            const result = findClosestStyleMatch(target, [normalizeKey(val)], 2);
            if (!result) continue;
            if (!bestMatch || result.distance < bestMatch.distance) {
                bestMatch = { row, field, distance: result.distance };
            }
        }
    }

    if (bestMatch) {
        console.log(`[nextgen-search] Levenshtein fallback matched style "${style}" via field "${bestMatch.field}" (distance ${bestMatch.distance})`);
        return { row: bestMatch.row, field: bestMatch.field };
    }

    return null;
}

export function normalizeColorName(s: string): string {
    return String(s || '')
        .toLowerCase()
        .replace(/[^a-z0-9]/g, '')
        .trim();
}

export function colorsMatch(a: string, b: string): boolean {
    // Use canonical color vocabulary first, then fall back to substring
    if (colorsEquivalent(a, b)) return true;
    // Legacy fallback: raw normalized substring
    const na = normalizeColorName(a);
    const nb = normalizeColorName(b);
    if (!na || !nb) return false;
    return na === nb || na.includes(nb) || nb.includes(na);
}

export function sizesMatch(a: string, b: string): boolean {
    // Use canonical size normalization (M <-> Medium, L <-> Large, etc.)
    if (sizesEquivalent(a, b)) return true;
    // Legacy fallback: raw normalized exact match
    return normalizeKey(a) === normalizeKey(b);
}
