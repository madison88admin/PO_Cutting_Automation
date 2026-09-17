/**
 * Persistent cache for NextGen style search results in Supabase.
 * This avoids re-querying NextGen for the same styles on every extraction.
 */
import { supabaseAdmin } from '@/lib/supabase';
import { NextGenStyleInfo } from '@/lib/types/buy-file';

const CACHE_TABLE = 'nextgen_style_cache';
const CACHE_TTL_MS = 24 * 60 * 60 * 1000; // 24 hours

interface CacheRow {
    style: string;
    info: NextGenStyleInfo | null;
    updated_at: string;
}

/**
 * Load cached style search results from Supabase.
 * Returns a map of style → NextGenStyleInfo (or null if not found in NextGen).
 * Only returns entries newer than the TTL.
 */
export async function loadCachedStyles(styles: string[]): Promise<Record<string, NextGenStyleInfo | null>> {
    const result: Record<string, NextGenStyleInfo | null> = {};
    if (!styles.length) return result;

    try {
        const { data, error } = await supabaseAdmin
            .from(CACHE_TABLE)
            .select('style, info, updated_at')
            .in('style', styles.map((s) => s.toLowerCase().trim()));

        if (error) {
            console.warn('[nextgen-cache] load error:', error.message);
            return result;
        }

        if (!data?.length) return result;

        const now = Date.now();
        for (const row of data as unknown as CacheRow[]) {
            const updatedAt = new Date(row.updated_at).getTime();
            if (now - updatedAt > CACHE_TTL_MS) continue; // expired
            result[row.style] = row.info;
        }

        const cachedCount = Object.keys(result).length;
        if (cachedCount > 0) {
            console.log(`[nextgen-cache] loaded ${cachedCount}/${styles.length} styles from cache`);
        }
    } catch (err) {
        console.warn('[nextgen-cache] load exception:', err);
    }

    return result;
}

/**
 * Save style search results to Supabase for future use.
 */
export async function saveCachedStyles(results: Record<string, NextGenStyleInfo | null>): Promise<void> {
    const entries = Object.entries(results);
    if (!entries.length) return;

    try {
        const rows = entries.map(([style, info]) => ({
            style: style.toLowerCase().trim(),
            info,
            updated_at: new Date().toISOString(),
        }));

        const { error } = await supabaseAdmin
            .from(CACHE_TABLE)
            .upsert(rows, { onConflict: 'style' });

        if (error) {
            console.warn('[nextgen-cache] save error:', error.message);
        } else {
            console.log(`[nextgen-cache] saved ${rows.length} styles to cache`);
        }
    } catch (err) {
        console.warn('[nextgen-cache] save exception:', err);
    }
}
