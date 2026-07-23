import { supabaseAdmin } from '@/lib/supabase';
import { ColumnMapping, ExtractedTemplate } from '@/lib/types/buy-file';

function normalizeHeader(header: string): string {
    return String(header || '')
        .toLowerCase()
        .replace(/[^a-z0-9]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
}

function normalizeHeaders(headers: string[]): string[] {
    return headers.map(normalizeHeader);
}

export async function findMatchingTemplateSupabase(
    headers: string[]
): Promise<ExtractedTemplate | null> {
    try {
        const normalized = normalizeHeaders(headers);
        const { data, error } = await supabaseAdmin
            .from('buy_file_templates')
            .select('*');

        if (error) {
            console.error('[supabase-store] findMatchingTemplate error:', error);
            return null;
        }

        if (!data?.length) return null;

        const headerSet = new Set(normalized.filter(Boolean));
        let best: ExtractedTemplate | null = null;
        let bestScore = 0;

        for (const row of data) {
            const rowNormalized = (row.normalized_headers || []) as string[];
            const rowSet = new Set(rowNormalized.filter(Boolean));
            const matches = normalized.filter((h) => rowSet.has(h)).length;
            const total = Math.max(headerSet.size, rowSet.size);
            if (total === 0) continue;
            const score = matches / total;
            if (score > 0.8 && score > bestScore) {
                bestScore = score;
                best = {
                    id: row.id,
                    customer: row.customer || null,
                    headers: (row.headers || []) as string[],
                    normalizedHeaders: rowNormalized,
                    mapping: row.mapping as ColumnMapping,
                    detectedAt: row.updated_at || row.created_at,
                };
            }
        }

        return best;
    } catch (err) {
        console.error('[supabase-store] findMatchingTemplate exception:', err);
        return null;
    }
}

export async function saveTemplateSupabase(
    headers: string[],
    mapping: ColumnMapping,
    customer: string | null
): Promise<ExtractedTemplate | null> {
    try {
        const existing = await findMatchingTemplateSupabase(headers);
        const payload = {
            customer,
            headers: headers.map((h) => String(h)),
            normalized_headers: normalizeHeaders(headers),
            mapping,
        };

        let result;
        if (existing) {
            const { data, error } = await supabaseAdmin
                .from('buy_file_templates')
                .update(payload)
                .eq('id', existing.id)
                .select()
                .single();
            if (error) throw error;
            result = data;
        } else {
            const { data, error } = await supabaseAdmin
                .from('buy_file_templates')
                .insert(payload)
                .select()
                .single();
            if (error) throw error;
            result = data;
        }

        if (!result) return null;
        return {
            id: result.id,
            customer: result.customer || null,
            headers: (result.headers || []) as string[],
            normalizedHeaders: (result.normalized_headers || []) as string[],
            mapping: result.mapping as ColumnMapping,
            detectedAt: result.updated_at || result.created_at,
        };
    } catch (err) {
        console.error('[supabase-store] saveTemplate exception:', err);
        return null;
    }
}
