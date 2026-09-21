/**
 * Server-side TTL cache for NextGen bulk records (PurchaseOrder/Read).
 *
 * Why: NextGen times out intermittently and every upload currently re-fetches
 * the same 500-row snapshot. This cache shares one snapshot across requests
 * for a configurable window, degrades to stale data when NextGen is down,
 * and collapses concurrent requests into a single in-flight fetch.
 *
 * Config (env):
 *   NEXTGEN_CACHE_TTL_MS      — fresh window (default 600000 = 10 min)
 *   NEXTGEN_CACHE_MAX_STALE_MS — how long stale data may still be served on
 *                                NextGen failure (default 3600000 = 1 hour)
 *
 * The cache lives on globalThis so it survives Next.js route-module reloads,
 * mirroring src/lib/processing-jobs.ts.
 */

export interface CacheStatus {
    hasCache: boolean;
    recordCount: number;
    ageMs: number;
    ttlMs: number;
    maxStaleMs: number;
    fresh: boolean;
    stale: boolean;
    inflight: boolean;
    lastError?: string;
    lastErrorAt?: number;
}

interface CacheState {
    cache?: { records: any[]; fetchedAt: number; pageSize: number };
    inflight?: Promise<any[]>;
    lastError?: string;
    lastErrorAt?: number;
}

const g = globalThis as typeof globalThis & {
    __nextgenRecordCache?: CacheState;
};

const state: CacheState = g.__nextgenRecordCache || (g.__nextgenRecordCache = {});

const TTL_MS = () => Number(process.env.NEXTGEN_CACHE_TTL_MS || 600000);
const MAX_STALE_MS = () => Number(process.env.NEXTGEN_CACHE_MAX_STALE_MS || 3600000);

function keyFor(pageSize: number): string {
    return `bulk-${pageSize}`;
}

function age(): number {
    return state.cache ? Date.now() - state.cache.fetchedAt : Infinity;
}

/**
 * Fetch bulk records through the cache.
 * - Fresh cache hit: no NextGen call at all.
 * - Concurrent callers share one in-flight fetch (single-flight).
 * - If NextGen fails but a stale snapshot exists (within MAX_STALE_MS),
 *   serve it instead of failing the request.
 */
export async function getRecordsCached(
    pageSize: number,
    fetcher: () => Promise<any[]>,
    opts: { bypass?: boolean } = {},
): Promise<{ records: any[]; source: 'fresh-cache' | 'network' | 'inflight' | 'stale'; ageMs: number }> {
    const key = keyFor(pageSize);

    if (!opts.bypass && state.cache && state.cache.pageSize === pageSize && age() <= TTL_MS()) {
        return { records: state.cache.records, source: 'fresh-cache', ageMs: age() };
    }

    if (state.inflight) {
        const records = await state.inflight;
        return { records, source: 'inflight', ageMs: age() };
    }

    const fetchPromise: Promise<any[]> = (async () => {
        try {
            const records = await fetcher();
            state.cache = { records, fetchedAt: Date.now(), pageSize };
            delete state.lastError;
            delete state.lastErrorAt;
            return records;
        } catch (err) {
            state.lastError = err instanceof Error ? err.message : String(err);
            state.lastErrorAt = Date.now();
            // Stale fallback: better to process with a recent snapshot than fail hard.
            if (state.cache && state.cache.pageSize === pageSize && age() <= MAX_STALE_MS()) {
                return state.cache.records;
            }
            throw err;
        } finally {
            // Only one fetch can be inflight at a time (guarded above), so an
            // unconditional clear is race-free.
            state.inflight = undefined;
        }
    })();

    state.inflight = fetchPromise;
    const records = await fetchPromise;
    // A just-set lastError (not yet cleared by a successful fetch) means this
    // call failed against NextGen and was served from the stale snapshot.
    const servedStale = Boolean(state.lastError);
    return { records, source: servedStale ? 'stale' : 'network', ageMs: age() };
}

export function getCacheStatus(): CacheStatus {
    const ttl = TTL_MS();
    const maxStale = MAX_STALE_MS();
    const ageMs = state.cache ? Date.now() - state.cache.fetchedAt : 0;
    return {
        hasCache: Boolean(state.cache),
        recordCount: state.cache?.records.length || 0,
        ageMs: state.cache ? ageMs : 0,
        ttlMs: ttl,
        maxStaleMs: maxStale,
        fresh: Boolean(state.cache) && ageMs <= ttl,
        stale: Boolean(state.cache) && ageMs > ttl,
        inflight: Boolean(state.inflight),
        lastError: state.lastError,
        lastErrorAt: state.lastErrorAt,
    };
}

export function clearCache(): void {
    delete state.cache;
    delete state.lastError;
    delete state.lastErrorAt;
}
