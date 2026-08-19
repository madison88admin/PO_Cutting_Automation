/**
 * Unified fuzzy matching utilities for PO Cutting Automation.
 *
 * Provides:
 *  - Fuse.js-based header/column fuzzy matching (threshold 0.15 = ~85% similarity)
 *  - Levenshtein distance for style/product code matching (handles typos)
 *  - Size normalization map (M <-> Medium, L <-> Large, etc.)
 *  - Canonical color vocabulary for apparel
 *
 * Design principle: deterministic logic first, fuzzy fallback, AI last.
 */

import Fuse, { IFuseOptions } from 'fuse.js';
import levenshtein from 'fast-levenshtein';

/* ------------------------------------------------------------------ *
 * 1. HEADER / COLUMN FUZZY MATCHING (Fuse.js)
 * ------------------------------------------------------------------ */

export interface FuzzyMatchResult {
    /** The matched canonical field/key. */
    key: string;
    /** The original input string that was matched. */
    input: string;
    /** Fuse.js score (0 = perfect, 1 = no match). Convert to % via (1 - score) * 100. */
    score: number;
    /** Similarity percentage 0-100. */
    similarity: number;
}

const HEADER_FUSE_OPTIONS: IFuseOptions<string> = {
    includeScore: true,
    threshold: 0.15, // ~85% similarity required
    ignoreLocation: true,
    minMatchCharLength: 2,
    isCaseSensitive: false,
};

/**
 * Find the best fuzzy match for an input header against a list of known
 * canonical patterns/aliases. Returns null if no match meets the threshold.
 *
 * @example
 * fuzzyMatchHeader("Styl Number", ["style number", "style no", "style#"])
 * // => { key: "style number", similarity: 92 }
 */
export function fuzzyMatchHeader(
    input: string,
    candidates: string[],
    threshold: number = 0.15,
): FuzzyMatchResult | null {
    const normalized = input.trim();
    if (!normalized || !candidates.length) return null;

    const fuse = new Fuse(candidates, { ...HEADER_FUSE_OPTIONS, threshold });
    const result = fuse.search(normalized);
    if (!result.length) return null;

    const best = result[0];
    const similarity = Math.round((1 - (best.score ?? 1)) * 100);
    if (similarity < 80) return null; // hard floor below 80%

    return {
        key: best.item,
        input: normalized,
        score: best.score ?? 1,
        similarity,
    };
}

/**
 * Batch fuzzy-match multiple input headers against a candidate dictionary.
 * Returns a map of input -> matched canonical key for matches above threshold.
 */
export function fuzzyMatchHeaders(
    inputs: string[],
    candidateMap: Record<string, string>,
    threshold: number = 0.15,
): Record<string, { canonical: string; similarity: number }> {
    const candidates = Object.keys(candidateMap);
    const result: Record<string, { canonical: string; similarity: number }> = {};

    for (const input of inputs) {
        const match = fuzzyMatchHeader(input, candidates, threshold);
        if (match) {
            result[input] = {
                canonical: candidateMap[match.key],
                similarity: match.similarity,
            };
        }
    }

    return result;
}

/* ------------------------------------------------------------------ *
 * 2. LEVENSHTEIN DISTANCE FOR STYLE / PRODUCT CODE MATCHING
 * ------------------------------------------------------------------ */

/**
 * Compute Levenshtein edit distance between two strings.
 */
export function editDistance(a: string, b: string): number {
    return levenshtein.get(a, b);
}

/**
 * Compute similarity ratio 0-1 based on Levenshtein distance.
 * 1 = identical, 0 = completely different.
 */
export function levenshteinSimilarity(a: string, b: string): number {
    const s1 = String(a || '');
    const s2 = String(b || '');
    if (!s1 && !s2) return 1;
    if (!s1 || !s2) return 0;
    const maxLen = Math.max(s1.length, s2.length);
    if (maxLen === 0) return 1;
    return 1 - editDistance(s1, s2) / maxLen;
}

export interface StyleMatchResult {
    candidate: string;
    distance: number;
    similarity: number;
}

/**
 * Find the closest style/product code match using Levenshtein distance.
 * Useful as a fallback after exact and prefix-strip matching fails.
 *
 * @param target  The style code from the buy file (e.g. "NF00E6Q")
 * @param candidates  Known style codes from NextGen or product sheet
 * @param maxDistance  Max allowed edit distance (default 2 — handles 1-2 char typos)
 */
export function findClosestStyleMatch(
    target: string,
    candidates: string[],
    maxDistance: number = 2,
): StyleMatchResult | null {
    const t = String(target || '').trim();
    if (!t || !candidates.length) return null;

    let best: StyleMatchResult | null = null;
    for (const candidate of candidates) {
        const c = String(candidate || '').trim();
        if (!c) continue;
        const dist = editDistance(t, c);
        if (dist > maxDistance) continue;
        const sim = 1 - dist / Math.max(t.length, c.length);
        if (!best || dist < best.distance) {
            best = { candidate: c, distance: dist, similarity: sim };
        }
    }
    return best;
}

/* ------------------------------------------------------------------ *
 * 3. SIZE NORMALIZATION
 * ------------------------------------------------------------------ */

/**
 * Canonical size mapping for apparel.
 * Maps abbreviations and variants to a canonical size name.
 */
export const SIZE_CANONICAL_MAP: Record<string, string> = {
    // Standard abbreviations
    'xs': 'XS',
    's': 'S',
    'm': 'M',
    'l': 'L',
    'xl': 'XL',
    'xxl': 'XXL',
    '2xl': 'XXL',
    'xxxl': 'XXXL',
    '3xl': 'XXXL',
    'xxxxl': 'XXXXL',
    '4xl': 'XXXXL',
    'xxxxxl': 'XXXXXL',
    '5xl': 'XXXXXL',
    // Full word forms
    'x small': 'XS',
    'xsmall': 'XS',
    'extra small': 'XS',
    'extrasmall': 'XS',
    'small': 'S',
    'medium': 'M',
    'med': 'M',
    'large': 'L',
    'lg': 'L',
    'x large': 'XL',
    'xlarge': 'XL',
    'extra large': 'XL',
    'extralarge': 'XL',
    'xx large': 'XXL',
    'xxlarge': 'XXL',
    '2x large': 'XXL',
    'xxx large': 'XXXL',
    'xxxlarge': 'XXXL',
    '3x large': 'XXXL',
    // One size / free size
    'os': 'ONE SIZE',
    'o/s': 'ONE SIZE',
    'one size': 'ONE SIZE',
    'onesize': 'ONE SIZE',
    'free': 'ONE SIZE',
    'free size': 'ONE SIZE',
    'freesize': 'ONE SIZE',
    'tu': 'ONE SIZE',
    '0os': 'ONE SIZE', // TNF format: "0OS"
    // Numeric sizes stay as-is (handled by normalizeSize)
    // Youth / Kids
    'ys': 'YS',
    'ym': 'YM',
    'yl': 'YL',
    'yxs': 'YXS',
    'yxl': 'YXL',
    'youth s': 'YS',
    'youth m': 'YM',
    'youth l': 'YL',
    'youth xs': 'YXS',
    'youth xl': 'YXL',
    // Toddler / Baby
    't1': '1T',
    't2': '2T',
    't3': '3T',
    't4': '4T',
    't5': '5T',
    '1t': '1T',
    '2t': '2T',
    '3t': '3T',
    '4t': '4T',
    '5t': '5T',
};

/**
 * Normalize a size string to its canonical form.
 * - Numeric sizes (e.g. "32", "38") are returned as-is
 * - Abbreviations are mapped to canonical names
 * - Unknown sizes are returned uppercased and trimmed
 */
export function normalizeSize(size: string | null | undefined): string {
    if (!size) return '';
    const raw = String(size).trim();
    if (!raw) return '';
    const key = raw.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (!key) return raw.toUpperCase();

    // Pure numeric → return as-is (e.g. "32", "38", "10.5")
    if (/^\d+(\.\d+)?$/.test(key)) return raw;

    // Check canonical map
    if (SIZE_CANONICAL_MAP[key]) return SIZE_CANONICAL_MAP[key];

    // Unknown → return uppercased original
    return raw.toUpperCase();
}

/**
 * Check if two size strings are equivalent after normalization.
 */
export function sizesEquivalent(a: string | null | undefined, b: string | null | undefined): boolean {
    const na = normalizeSize(a);
    const nb = normalizeSize(b);
    if (!na || !nb) return false;
    return na === nb;
}

/* ------------------------------------------------------------------ *
 * 4. CANONICAL COLOR VOCABULARY
 * ------------------------------------------------------------------ */

/**
 * Canonical color mapping for apparel.
 * Maps common synonyms, abbreviations, and variants to a canonical color name.
 */
export const COLOR_CANONICAL_MAP: Record<string, string> = {
    // Black
    'blk': 'Black',
    'blac': 'Black',
    'bk': 'Black',
    'bl': 'Black',
    'b': 'Black',
    'nero': 'Black',
    'noir': 'Black',
    'schwarz': 'Black',
    // White
    'wht': 'White',
    'wh': 'White',
    'w': 'White',
    'bianco': 'White',
    'blanc': 'White',
    'weiss': 'White',
    // Grey / Gray
    'gry': 'Grey',
    'gr': 'Grey',
    'gray': 'Grey',
    'grey': 'Grey',
    'charcoal': 'Charcoal Grey',
    'heather grey': 'Heather Grey',
    'heather gray': 'Heather Grey',
    'slate': 'Slate Grey',
    // Navy / Blue
    'navy': 'Navy',
    'nv': 'Navy',
    'nvy': 'Navy',
    'navy blue': 'Navy',
    'dark blue': 'Navy',
    'dk blue': 'Navy',
    'royal': 'Royal Blue',
    'royal blue': 'Royal Blue',
    'cobalt': 'Cobalt Blue',
    'indigo': 'Indigo',
    'teal': 'Teal',
    'aqua': 'Aqua',
    'turquoise': 'Turquoise',
    'sky': 'Sky Blue',
    'sky blue': 'Sky Blue',
    'blue': 'Blue',
    // Red
    'rd': 'Red',
    'r': 'Red',
    'rosso': 'Red',
    'rouge': 'Red',
    'rot': 'Red',
    'burgundy': 'Burgundy',
    'maroon': 'Maroon',
    'wine': 'Wine',
    'crimson': 'Crimson',
    // Green
    'grn': 'Green',
    'gn': 'Green',
    'verde': 'Green',
    'vert': 'Green',
    'olive': 'Olive',
    'olive green': 'Olive',
    'forest': 'Forest Green',
    'forest green': 'Forest Green',
    'khaki': 'Khaki',
    'sage': 'Sage Green',
    'lime': 'Lime',
    'mint': 'Mint',
    // Yellow / Orange
    'ylw': 'Yellow',
    'yl': 'Yellow',
    'yellow': 'Yellow',
    'gold': 'Gold',
    'mustard': 'Mustard',
    'org': 'Orange',
    'orange': 'Orange',
    'coral': 'Coral',
    'peach': 'Peach',
    // Pink / Purple
    'pk': 'Pink',
    'pink': 'Pink',
    'rose': 'Rose',
    'fuchsia': 'Fuchsia',
    'magenta': 'Magenta',
    'pur': 'Purple',
    'purple': 'Purple',
    'violet': 'Violet',
    'lavender': 'Lavender',
    'plum': 'Plum',
    'lilac': 'Lilac',
    // Brown / Tan
    'brn': 'Brown',
    'br': 'Brown',
    'brown': 'Brown',
    'chocolate': 'Chocolate',
    'espresso': 'Espresso',
    'mocha': 'Mocha',
    'tan': 'Tan',
    'sand': 'Sand',
    'beige': 'Beige',
    'camel': 'Camel',
    'taupe': 'Taupe',
    // Multi / Patterns
    'multi': 'Multi',
    'multicolour': 'Multi',
    'multicolor': 'Multi',
    'assorted': 'Assorted',
    'combo': 'Combo',
    'mixed': 'Mixed',
    'print': 'Print',
    'floral': 'Floral',
    'stripe': 'Stripe',
    'stripes': 'Stripe',
    'camo': 'Camouflage',
    'camouflage': 'Camouflage',
};

/**
 * Normalize a color name to its canonical form.
 * Strips brand prefixes, color/colour words, and maps abbreviations.
 */
export function normalizeColorName(s: string | null | undefined): string {
    if (!s) return '';
    let raw = String(s).trim().toLowerCase();
    if (!raw) return '';

    // Strip common brand prefixes: "VANS - ", "TNF ", "JW - ", etc.
    raw = raw.replace(/^(vans|tnf|jw|vuori|hh|dyn|pp|cot|marmot|llb)\s*[-:]?\s*/i, '');

    // Strip "color"/"colour" words
    raw = raw.replace(/\b(colou?r)\b/gi, '');

    // Normalize whitespace and strip non-alphanumeric
    const key = raw.replace(/[^a-z0-9]/g, '').trim();
    if (!key) return '';

    // Check canonical map
    if (COLOR_CANONICAL_MAP[key]) return COLOR_CANONICAL_MAP[key];

    // Return title-cased original (stripped of prefixes)
    const cleaned = raw.replace(/[^a-z0-9\s]/gi, '').replace(/\s+/g, ' ').trim();
    return cleaned
        .split(' ')
        .filter(Boolean)
        .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
        .join(' ');
}

/**
 * Check if two color names are equivalent after normalization.
 * Uses canonical mapping first, then falls back to substring match.
 */
export function colorsEquivalent(a: string | null | undefined, b: string | null | undefined): boolean {
    const na = normalizeColorName(a);
    const nb = normalizeColorName(b);
    if (!na || !nb) return false;
    if (na === nb) return true;
    // Fallback: substring match on normalized lowercase
    const la = na.toLowerCase();
    const lb = nb.toLowerCase();
    return la.includes(lb) || lb.includes(la);
}
