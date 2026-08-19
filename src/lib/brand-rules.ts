/**
 * Generic Brand Rule Engine
 *
 * Applies config-driven transformations that were previously hard-coded
 * as brand-specific if/else branches in excel-engine.ts.
 *
 * Rules are defined per-brand in brand-config.json under the "rules" key.
 * Unknown brands fall back to the "default" rule set, making the system
 * fully flexible for new brands without code changes.
 *
 * Rule categories:
 *   - styleNumber: ordered source field priority + preferCodeLike flag
 *   - color: ordered source field priority
 *   - size: regex → canonical value mappings
 *   - status: regex → canonical value mappings + default
 *   - poNumber: format type ("standard" | "vans" | "cotopaxi")
 *   - destination: default country when none detected
 */

import { getBrand } from '@/lib/brand-config';

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface RegexMapping {
    /** Regex pattern to test against the raw value */
    pattern: string;
    /** Regex flags (e.g. "i" for case-insensitive) */
    flags?: string;
    /** Canonical value to return if the pattern matches */
    value: string;
}

export interface BrandRules {
    /** Style number source priority — ordered list of field names to try */
    styleNumber?: {
        source?: string[];
        preferCodeLike?: boolean;
    };
    /** Color source priority — ordered list of field names to try */
    color?: {
        source?: string[];
    };
    /** Size normalization — regex mappings applied in order */
    size?: {
        mappings?: RegexMapping[];
        default?: string;
    };
    /** Status normalization — regex mappings + default value */
    status?: {
        mappings?: RegexMapping[];
        default?: string;
    };
    /** PO number format type */
    poNumber?: {
        format?: 'standard' | 'vans' | 'cotopaxi';
    };
    /** Destination fallback when none detected */
    destination?: {
        default?: string;
    };
}

// ---------------------------------------------------------------------------
// Default rules (used when a brand has no specific rules)
// ---------------------------------------------------------------------------

const DEFAULT_RULES: BrandRules = {
    styleNumber: {
        source: ['product', 'productCustomerRef', 'matchedStyleKey', 'buyProductName', 'jdeStyle'],
        preferCodeLike: true,
    },
    color: {
        source: ['inlineColorDescription', 'colourNameRaw', 'rawColour'],
    },
    size: {
        mappings: [
            { pattern: '^ons$', flags: 'i', value: 'One Size' },
            { pattern: '^one\\s*size$', flags: 'i', value: 'One Size' },
            { pattern: '^onesize$', flags: 'i', value: 'One Size' },
            { pattern: '-?1sz$', flags: 'i', value: 'One Size' },
            { pattern: '\\b1sz\\b', flags: 'i', value: 'One Size' },
            { pattern: '^reg\\s*fit$', flags: 'i', value: 'Reg Fit' },
            { pattern: '\\brg1sz\\b', flags: 'i', value: 'Reg Fit' },
            { pattern: '\\breg(?:ular)?\\s*fit\\b', flags: 'i', value: 'Reg Fit' },
            { pattern: '^helmet\\s*fit$', flags: 'i', value: 'Helmet Fit' },
            { pattern: '\\bhs\\d*\\b', flags: 'i', value: 'Helmet Fit' },
            { pattern: '\\bhelmet\\s*fit\\b', flags: 'i', value: 'Helmet Fit' },
        ],
    },
    status: {
        mappings: [],
        default: 'Confirmed',
    },
    poNumber: {
        format: 'standard',
    },
    destination: {
        default: '',
    },
};

// ---------------------------------------------------------------------------
// Rule lookup
// ---------------------------------------------------------------------------

/**
 * Get the effective rules for a brand.
 * Merges brand-specific rules over defaults.
 */
export function getBrandRules(brandName: string): BrandRules {
    const brand = getBrand(brandName);
    if (!brand || !(brand as any).rules) return DEFAULT_RULES;

    const brandRules = (brand as any).rules as BrandRules;
    return {
        styleNumber: { ...DEFAULT_RULES.styleNumber, ...brandRules.styleNumber },
        color: { ...DEFAULT_RULES.color, ...brandRules.color },
        size: { ...DEFAULT_RULES.size, ...brandRules.size },
        status: { ...DEFAULT_RULES.status, ...brandRules.status },
        poNumber: { ...DEFAULT_RULES.poNumber, ...brandRules.poNumber },
        destination: { ...DEFAULT_RULES.destination, ...brandRules.destination },
    };
}

// ---------------------------------------------------------------------------
// Rule application functions
// ---------------------------------------------------------------------------

/**
 * Check if a string looks like a style code (has digit, no space).
 */
export function isLikelyStyleCode(value: string): boolean {
    const s = String(value || '').trim();
    if (!s || s.length > 20) return false;
    return /\d/.test(s) && !/\s/.test(s);
}

/**
 * Resolve a field value using source priority rules.
 * Tries each source field in order, returns the first non-empty value.
 *
 * @param sources Ordered list of field names to try
 * @param getField Function that returns the value for a given field name
 * @param preferCodeLike If true, prefer the first code-like value over name-like values
 * @returns The resolved value, or empty string
 */
export function resolveBySourcePriority(
    sources: string[],
    getField: (name: string) => string | undefined,
    preferCodeLike?: boolean,
): string {
    if (preferCodeLike) {
        // First pass: find a code-like value
        for (const src of sources) {
            const val = String(getField(src) || '').trim();
            if (val && isLikelyStyleCode(val)) return val;
        }
    }
    // Second pass: return first non-empty value
    for (const src of sources) {
        const val = String(getField(src) || '').trim();
        if (val) return val;
    }
    return '';
}

/**
 * Apply regex mappings to a value.
 * Returns the first matching mapping's canonical value, or the original value.
 *
 * @param value Raw value to normalize
 * @param mappings Ordered list of regex mappings
 * @param testKey Optional: test against a different string (e.g. a normalized key)
 * @returns Canonical value or original
 */
export function applyRegexMappings(
    value: string,
    mappings: RegexMapping[] | undefined,
    testKey?: string,
): string {
    if (!mappings || !mappings.length) return value;
    const testValue = testKey || value;
    for (const mapping of mappings) {
        try {
            const regex = new RegExp(mapping.pattern, mapping.flags || '');
            if (regex.test(testValue)) return mapping.value;
        } catch {
            // Invalid regex — skip
        }
    }
    return value;
}

/**
 * Resolve style number using brand rules.
 *
 * @param brandName Brand key
 * @param getField Function that returns the value for a given field name
 * @returns Resolved style number
 */
export function resolveStyleNumber(
    brandName: string,
    getField: (name: string) => string | undefined,
): string {
    const rules = getBrandRules(brandName);
    const sources = rules.styleNumber?.source || ['product', 'buyProductName'];
    const preferCodeLike = rules.styleNumber?.preferCodeLike ?? true;
    return resolveBySourcePriority(sources, getField, preferCodeLike);
}

/**
 * Resolve color using brand rules.
 */
export function resolveColor(
    brandName: string,
    getField: (name: string) => string | undefined,
): string {
    const rules = getBrandRules(brandName);
    const sources = rules.color?.source || ['rawColour'];
    return resolveBySourcePriority(sources, getField, false);
}

/**
 * Normalize size using brand rules.
 */
export function normalizeSizeByRules(
    size: string,
    sizeKey: string,
    brandName: string,
): string {
    const rules = getBrandRules(brandName);
    const result = applyRegexMappings(size, rules.size?.mappings, sizeKey);
    if (result !== size) return result;
    // Try sizeKey as well
    const keyResult = applyRegexMappings(sizeKey, rules.size?.mappings);
    if (keyResult !== sizeKey) return keyResult;
    return rules.size?.default || size;
}

/**
 * Normalize status using brand rules.
 */
export function normalizeStatusByRules(
    status: string,
    brandName: string,
): string {
    const rules = getBrandRules(brandName);
    const result = applyRegexMappings(status, rules.status?.mappings);
    if (result !== status) return result;
    return rules.status?.default || status || 'Confirmed';
}

/**
 * Get PO number format type for a brand.
 */
export function getPoNumberFormat(brandName: string): 'standard' | 'vans' | 'cotopaxi' {
    const rules = getBrandRules(brandName);
    return rules.poNumber?.format || 'standard';
}

/**
 * Get destination default for a brand.
 */
export function getDestinationDefault(brandName: string): string {
    const rules = getBrandRules(brandName);
    return rules.destination?.default || '';
}
