/**
 * Learning Layer — caches successful matches and user corrections
 * so the system gets faster and more accurate over time.
 *
 * Four caches:
 *   1. NextGen match cache  — style+color → NextGen product info
 *   2. User correction memory — user overrides for wrong matches
 *   3. Header mapping cache  — header layout → canonical mapping per brand
 *   4. Color mapping cache   — raw color → canonical color per brand
 *
 * Storage:
 *   - Supabase (when available) for shared/persistent storage
 *   - JSON file fallback (when no DB) for local development
 *
 * All caches are keyed by normalized values so lookups are brand-agnostic.
 */

import { supabaseAdmin, isMock } from '@/lib/supabase';
import { NextGenStyleInfo } from '@/lib/types/buy-file';

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface NextGenMatchCacheEntry {
    id?: string;
    style_key: string;           // normalized style code (e.g. "nf0a8cgz")
    color_key: string;            // normalized color code (e.g. "e8j")
    brand: string;                // brand key (e.g. "tnf")
    nextgen_product: string;      // M88131054
    nextgen_product_id: number;   // NextGen internal ID
    nextgen_style_name: string;   // "Salty Lined Beanie"
    nextgen_color_name: string;   // "Wax Paper-A5 Gryphon Patch"
    nextgen_color_code: string;   // "E8J"
    unit_cost: number | null;
    currency: string | null;
    factory: string | null;
    costing_reference: string | null;
    match_score: number;          // confidence score 0-100
    match_reason: string;         // why it matched
    hit_count: number;            // how many times this cache entry was used
    created_at: string;
    updated_at: string;
}

export interface UserCorrectionEntry {
    id?: string;
    style_key: string;
    color_key: string;
    brand: string;
    original_nextgen_product: string;   // what the system matched
    corrected_nextgen_product: string;  // what the user chose
    corrected_nextgen_product_id: number;
    corrected_color_name: string;
    corrected_color_code: string;
    corrected_by: string;        // user ID
    created_at: string;
}

export interface HeaderMappingCacheEntry {
    id?: string;
    brand: string;
    file_signature: string;       // hash of sorted header names
    raw_headers: string[];        // original headers
    mapped_headers: Record<string, string>;  // canonical → column name
    confidence: number;
    hit_count: number;
    created_at: string;
    updated_at: string;
}

export interface ColorMappingCacheEntry {
    id?: string;
    brand: string;
    raw_color: string;            // raw color from buy file
    canonical_color: string;      // normalized color
    nextgen_color_name: string;   // matched NextGen color name
    nextgen_color_code: string;   // matched NextGen color code
    hit_count: number;
    created_at: string;
    updated_at: string;
}

// ---------------------------------------------------------------------------
// Normalization helpers
// ---------------------------------------------------------------------------

function normKey(s: string): string {
    return String(s || '')
        .toLowerCase()
        .replace(/[^a-z0-9]/g, '')
        .trim();
}

function fileSignature(headers: string[]): string {
    return headers
        .map(h => normKey(h))
        .filter(Boolean)
        .sort()
        .join('|');
}

function now(): string {
    return new Date().toISOString();
}

// ---------------------------------------------------------------------------
// In-memory cache (always available, even in mock mode)
// ---------------------------------------------------------------------------

const memNextGenCache = new Map<string, NextGenMatchCacheEntry>();
const memCorrections = new Map<string, UserCorrectionEntry>();
const memHeaderCache = new Map<string, HeaderMappingCacheEntry>();
const memColorCache = new Map<string, ColorMappingCacheEntry>();

function memKey(style: string, color: string, brand: string): string {
    return `${brand}:${normKey(style)}:${normKey(color)}`;
}

// ---------------------------------------------------------------------------
// 1. NextGen Match Cache
// ---------------------------------------------------------------------------

/**
 * Look up a cached NextGen match for a style+color+brand combination.
 * Returns null if not cached or cache is stale.
 */
export async function getCachedNextGenMatch(
    style: string,
    color: string,
    brand: string,
): Promise<NextGenMatchCacheEntry | null> {
    const key = memKey(style, color, brand);

    // Check in-memory first (fastest)
    const memHit = memNextGenCache.get(key);
    if (memHit) {
        memHit.hit_count++;
        memHit.updated_at = now();
        return memHit;
    }

    if (isMock) return null;

    try {
        const { data, error } = await supabaseAdmin
            .from('nextgen_match_cache')
            .select('*')
            .eq('style_key', normKey(style))
            .eq('color_key', normKey(color))
            .eq('brand', brand.toLowerCase())
            .limit(1);

        if (error || !data || data.length === 0) return null;

        // Populate in-memory cache
        memNextGenCache.set(key, data[0]);
        return data[0];
    } catch {
        return null;
    }
}

/**
 * Save a successful NextGen match to the cache.
 * Called after searchVariant() finds a reliable match.
 */
export async function saveNextGenMatch(
    style: string,
    color: string,
    brand: string,
    info: NextGenStyleInfo,
): Promise<void> {
    const key = memKey(style, color, brand);
    const entry: NextGenMatchCacheEntry = {
        style_key: normKey(style),
        color_key: normKey(color),
        brand: brand.toLowerCase(),
        nextgen_product: info.product || '',
        nextgen_product_id: info.style ? Number(info.style) || 0 : 0,
        nextgen_style_name: info.styleName || '',
        nextgen_color_name: info.colorName || '',
        nextgen_color_code: info.colorCode || '',
        unit_cost: info.unitCost ?? null,
        currency: info.currency ?? null,
        factory: info.factory ?? null,
        costing_reference: info.costingReference ?? null,
        match_score: info.matchScore ?? 0,
        match_reason: info.matchReason || '',
        hit_count: 1,
        created_at: now(),
        updated_at: now(),
    };

    // Always save to in-memory
    memNextGenCache.set(key, entry);

    if (isMock) return;

    try {
        const { error } = await supabaseAdmin
            .from('nextgen_match_cache')
            .upsert({
                ...entry,
                updated_at: now(),
            }, {
                onConflict: 'style_key,color_key,brand',
            });
        if (error) {
            console.warn('[learning-layer] Failed to save NextGen match to Supabase:', error.message);
        }
    } catch (err) {
        console.warn('[learning-layer] Failed to save NextGen match to Supabase:', err);
    }
}

// ---------------------------------------------------------------------------
// 2. User Correction Memory
// ---------------------------------------------------------------------------

/**
 * Get a user correction for a style+color+brand.
 * If a correction exists, it overrides the cached match.
 */
export async function getUserCorrection(
    style: string,
    color: string,
    brand: string,
): Promise<UserCorrectionEntry | null> {
    const key = memKey(style, color, brand);

    const memHit = memCorrections.get(key);
    if (memHit) {
        console.log(`[learning-layer] Correction mem hit: ${style}/${color}/${brand} → ${memHit.corrected_nextgen_product}`);
        return memHit;
    }

    if (isMock) return null;

    try {
        const ns = normKey(style);
        const nc = normKey(color);
        const nb = brand.toLowerCase();
        console.log(`[learning-layer] Looking up correction: style_key=${ns}, color_key=${nc}, brand=${nb}`);
        const { data, error } = await supabaseAdmin
            .from('nextgen_user_corrections')
            .select('*')
            .eq('style_key', ns)
            .eq('color_key', nc)
            .eq('brand', nb)
            .order('created_at', { ascending: false })
            .limit(1);

        if (error) {
            console.log(`[learning-layer] Correction lookup error: ${error.message}`);
            return null;
        }
        if (!data || data.length === 0) {
            console.log(`[learning-layer] No correction found for ${ns}/${nc}/${nb}`);
            return null;
        }
        const row = data[0];

        console.log(`[learning-layer] Correction DB hit: ${ns}/${nc}/${nb} → ${row.corrected_nextgen_product}`);
        memCorrections.set(key, row);
        return row;
    } catch (err) {
        console.log(`[learning-layer] Correction lookup exception: ${err}`);
        return null;
    }
}

/**
 * Save a user correction. This is called when the user manually overrides
 * a wrong NextGen match.
 */
export async function saveUserCorrection(
    style: string,
    color: string,
    brand: string,
    originalProduct: string,
    correctedProduct: string,
    correctedProductId: number,
    correctedColorName: string,
    correctedColorCode: string,
    userId: string,
): Promise<void> {
    const key = memKey(style, color, brand);
    const entry: UserCorrectionEntry = {
        style_key: normKey(style),
        color_key: normKey(color),
        brand: brand.toLowerCase(),
        original_nextgen_product: originalProduct,
        corrected_nextgen_product: correctedProduct,
        corrected_nextgen_product_id: correctedProductId,
        corrected_color_name: correctedColorName,
        corrected_color_code: correctedColorCode,
        corrected_by: userId,
        created_at: now(),
    };

    memCorrections.set(key, entry);

    if (isMock) return;

    try {
        const { error } = await supabaseAdmin
            .from('nextgen_user_corrections')
            .insert(entry);
        if (error) {
            console.warn('[learning-layer] Failed to save user correction to Supabase:', error.message);
        } else {
            console.log(`[learning-layer] Correction persisted to Supabase: ${normKey(style)}/${normKey(color)}/${brand.toLowerCase()} → ${correctedProduct}`);
        }
    } catch (err) {
        console.warn('[learning-layer] Failed to save user correction:', err);
    }
}

// ---------------------------------------------------------------------------
// 3. Header Mapping Cache
// ---------------------------------------------------------------------------

/**
 * Look up a cached header mapping for a brand + file signature.
 * Returns the cached mapping if the same header layout was seen before.
 */
export async function getCachedHeaderMapping(
    brand: string,
    headers: string[],
): Promise<HeaderMappingCacheEntry | null> {
    const sig = fileSignature(headers);
    const key = `${brand.toLowerCase()}:${sig}`;

    const memHit = memHeaderCache.get(key);
    if (memHit) {
        memHit.hit_count++;
        memHit.updated_at = now();
        return memHit;
    }

    if (isMock) return null;

    try {
        const { data, error } = await supabaseAdmin
            .from('header_mapping_cache')
            .select('*')
            .eq('brand', brand.toLowerCase())
            .eq('file_signature', sig)
            .limit(1);

        if (error || !data || data.length === 0) return null;

        memHeaderCache.set(key, data[0]);
        return data[0];
    } catch {
        return null;
    }
}

/**
 * Save a successful header mapping to the cache.
 */
export async function saveHeaderMapping(
    brand: string,
    headers: string[],
    mapping: Record<string, string>,
    confidence: number,
): Promise<void> {
    const sig = fileSignature(headers);
    const key = `${brand.toLowerCase()}:${sig}`;
    const entry: HeaderMappingCacheEntry = {
        brand: brand.toLowerCase(),
        file_signature: sig,
        raw_headers: headers,
        mapped_headers: mapping,
        confidence,
        hit_count: 1,
        created_at: now(),
        updated_at: now(),
    };

    memHeaderCache.set(key, entry);

    if (isMock) return;

    try {
        const { error } = await supabaseAdmin
            .from('header_mapping_cache')
            .upsert({
                ...entry,
                updated_at: now(),
            }, {
                onConflict: 'brand,file_signature',
            });
        if (error) {
            console.warn('[learning-layer] Failed to save header mapping to Supabase:', error.message);
        }
    } catch (err) {
        console.warn('[learning-layer] Failed to save header mapping:', err);
    }
}

// ---------------------------------------------------------------------------
// 4. Color Mapping Cache
// ---------------------------------------------------------------------------

/**
 * Look up a cached color mapping for a brand + raw color.
 */
export async function getCachedColorMapping(
    brand: string,
    rawColor: string,
): Promise<ColorMappingCacheEntry | null> {
    const key = `${brand.toLowerCase()}:${normKey(rawColor)}`;

    const memHit = memColorCache.get(key);
    if (memHit) {
        memHit.hit_count++;
        memHit.updated_at = now();
        return memHit;
    }

    if (isMock) return null;

    try {
        const { data, error } = await supabaseAdmin
            .from('color_mapping_cache')
            .select('*')
            .eq('brand', brand.toLowerCase())
            .eq('raw_color', rawColor.toLowerCase().trim())
            .limit(1);

        if (error || !data || data.length === 0) return null;

        memColorCache.set(key, data[0]);
        return data[0];
    } catch {
        return null;
    }
}

/**
 * Save a successful color mapping to the cache.
 */
export async function saveColorMapping(
    brand: string,
    rawColor: string,
    canonicalColor: string,
    nextgenColorName: string,
    nextgenColorCode: string,
): Promise<void> {
    const key = `${brand.toLowerCase()}:${normKey(rawColor)}`;
    const entry: ColorMappingCacheEntry = {
        brand: brand.toLowerCase(),
        raw_color: rawColor.toLowerCase().trim(),
        canonical_color: canonicalColor,
        nextgen_color_name: nextgenColorName,
        nextgen_color_code: nextgenColorCode,
        hit_count: 1,
        created_at: now(),
        updated_at: now(),
    };

    memColorCache.set(key, entry);

    if (isMock) return;

    try {
        const { error } = await supabaseAdmin
            .from('color_mapping_cache')
            .upsert({
                ...entry,
                updated_at: now(),
            }, {
                onConflict: 'brand,raw_color',
            });
        if (error) {
            console.warn('[learning-layer] Failed to save color mapping to Supabase:', error.message);
        }
    } catch (err) {
        console.warn('[learning-layer] Failed to save color mapping:', err);
    }
}

// ---------------------------------------------------------------------------
// Cache stats (for debugging/UI)
// ---------------------------------------------------------------------------

export function getCacheStats() {
    return {
        nextgenMatches: memNextGenCache.size,
        userCorrections: memCorrections.size,
        headerMappings: memHeaderCache.size,
        colorMappings: memColorCache.size,
        isMock,
    };
}
