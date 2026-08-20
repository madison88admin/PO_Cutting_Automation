import { jsonrepair } from 'jsonrepair';
import { GROQ_API_KEY } from '@/lib/constants';
import { chatWithOllamaJson } from '@/lib/ai/ollama-client';
import { ColumnMapping, HeaderMappingResult } from '@/lib/types/buy-file';
import { fuzzyMatchHeader } from '@/lib/fuzzy';
import { getCachedHeaderMapping, saveHeaderMapping } from '@/lib/learning/cache';

const CANONICAL_FIELDS = new Set([
    'buyer_style_number', 'buyer_style_name', 'sku', 'product_description',
    'color', 'color_code', 'size', 'quantity', 'delivery_date', 'season',
    'customer', 'factory', 'currency', 'unit_cost', 'po_number',
    'buyer_po_number', 'start_date', 'cancel_date', 'transport_method',
    'buy_information',
]);

function normalizeMapping(mapping: Record<string, string>): ColumnMapping {
    const normalized: ColumnMapping = {};
    if (!mapping || typeof mapping !== 'object') return normalized;

    for (const [key, value] of Object.entries(mapping)) {
        if (!value) continue;
        const keyLower = key.toLowerCase().trim();
        const valueLower = value.toLowerCase().trim();

        if (CANONICAL_FIELDS.has(keyLower)) {
            // Already correct: canonical -> header
            (normalized as Record<string, string>)[keyLower] = value;
        } else if (CANONICAL_FIELDS.has(valueLower)) {
            // Reversed: header -> canonical, flip it
            (normalized as Record<string, string>)[valueLower] = key;
        }
    }

    return normalized;
}

const HEADER_PATTERNS: { field: string; patterns: string[] }[] = [
    { field: 'po_number', patterns: ['final po cut', 'master po', 'po number', 'purchase order', 'order number', 'po no', 'po no.', 'po#', 'so number', 'so#', 'sales order', 'po'] },
    { field: 'buyer_po_number', patterns: ['buyer po number', 'buyer po #', 'buyer po', 'customer po number', 'customer po #', 'customer po', 'bp no', 'extraction po #'] },
    { field: 'buyer_style_number', patterns: ['style#', 'style #', 'style number', 'style no', 'style no.', 'style ref', 'style reference', 'article', 'model', 'model no', 'model number', 'style'] },
    { field: 'buyer_style_name', patterns: ['style name', 'style description', 'style desc', 'style nm', 'style narrative'] },
    { field: 'sku', patterns: ['material', 'sku', 'upc', 'ean', 'product code', 'article code', 'old sku', 'eu old sku'] },
    { field: 'product_description', patterns: ['longtext', 'material description', 'product description', 'item description', 'style description', 'description', 'desc'] },
    { field: 'color_code', patterns: ['style color', 'color code', 'colour code', 'colorway code', 'color no', 'color #'] },
    { field: 'color', patterns: ['colorway name', 'color name', 'colour name', 'colorway', 'color', 'colour'] },
    { field: 'size', patterns: ['size 1', 'size 2', 'size name', 'size scale', 'product size', 'size#', 'size'] },
    { field: 'quantity', patterns: ['total quantity', 'order qty', 'po qty', 'buy qty', '1st qty', '2nd qty', 'quantity', 'qty', 'units'] },
    { field: 'delivery_date', patterns: ['vendor confirmed crd', 'brand requested ped', 'planned ped', 'delivery date', 'crdd date', 'ex factory', 'ex-fty', 'ship date', 'delivery', 'crd', 'ped', 'target date'] },
    { field: 'start_date', patterns: ['udf-start_date', 'start date', 'order start date', 'valid from'] },
    { field: 'cancel_date', patterns: ['udf-canel_date', 'udf-cancel_date', 'cancel date', 'canel date', 'order cancel date', 'valid until'] },
    { field: 'season', patterns: ['season', 'buy season', 'season year', 'year'] },
    { field: 'customer', patterns: ['sold-to party', 'sold to party', 'sold to', 'customer', 'buyer', 'brand', 'sales market', 'sales org', 'company'] },
    { field: 'factory', patterns: ['final factory name', 'final factory', 'final vendor name', 'final vendor', 'factory name', 'factory', 'vendor', 'supplier', 'manufacturer'] },
    { field: 'currency', patterns: ['final currency', 'currency', 'curr'] },
    { field: 'unit_cost', patterns: ['fob', 'unit cost', 'unit price', 'factory cost', 'cost', 'price', 'production upcharges usd', 'material upcharges usd', 'upcharge', 'up charge'] },
    { field: 'transport_method', patterns: ['order transport', 'transport method', 'transportation mode', 'transportation mode description', 'transport mode', 'shipment method', 'shipment mode', 'shipping method', 'ship mode', 'ship via', 'mode of delivery', 'freight mode'] },
    { field: 'buy_information', patterns: ['buy information', 'buy info', 'buying information', 'buying info', 'purchase information', 'purchase info', 'po information', 'po info', 'buy details', 'buy detail'] },
    { field: 'status', patterns: ['status', 'po status', 'order status', 'line status', 'workflow status', 'decision'] },
];

function patternToRegex(pattern: string): RegExp {
    // Escape special regex chars, then allow optional whitespace between words
    const escaped = pattern.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const withOptionalSpaces = escaped.replace(/\s+/g, '\\s*');
    return new RegExp(`^${withOptionalSpaces}$`, 'i');
}

function fallbackHeuristicMapping(headers: string[]): ColumnMapping {
    const mapping: ColumnMapping = {};
    const usedHeaders = new Set<number>();
    const strippedHeaders = headers.map((h) => h.trim().replace(/[^a-zA-Z0-9]+$/, ''));

    for (const { field, patterns } of HEADER_PATTERNS) {
        for (const pattern of patterns) {
            const regex = patternToRegex(pattern);
            for (let i = 0; i < strippedHeaders.length; i++) {
                if (usedHeaders.has(i)) continue;
                if (regex.test(strippedHeaders[i])) {
                    (mapping as Record<string, string>)[field] = headers[i];
                    usedHeaders.add(i);
                    break;
                }
            }
            if ((mapping as Record<string, string>)[field]) break;
        }
    }

    return mapping;
}

/**
 * Fuzzy matching layer: uses Fuse.js to match headers that the exact
 * heuristic patterns missed. Runs BEFORE the AI/LLM fallback so that
 * simple typos and minor variations don't need an LLM round-trip.
 *
 * Threshold 0.15 (~85% similarity) per industry standard for column
 * header matching. Hard floor at 80% to avoid false positives.
 */
function fuzzyHeuristicMapping(
    headers: string[],
    alreadyMapped: ColumnMapping,
): { mapping: ColumnMapping; fuzzyConfidence: number } {
    const mapping: ColumnMapping = {};
    const mappedHeaders = new Set(Object.values(alreadyMapped as Record<string, string>));
    const usedFields = new Set(Object.keys(alreadyMapped as Record<string, string>));

    // Build candidate list: each canonical field -> all its patterns
    const candidateToField: Record<string, string> = {};
    for (const { field, patterns } of HEADER_PATTERNS) {
        for (const pattern of patterns) {
            candidateToField[pattern] = field;
        }
    }
    const candidates = Object.keys(candidateToField);

    let totalSimilarity = 0;
    let matchCount = 0;

    for (let i = 0; i < headers.length; i++) {
        if (mappedHeaders.has(headers[i])) continue;
        const match = fuzzyMatchHeader(headers[i], candidates, 0.15);
        if (!match) continue;

        const field = candidateToField[match.key];
        if (!field || usedFields.has(field)) continue;

        (mapping as Record<string, string>)[field] = headers[i];
        usedFields.add(field);
        mappedHeaders.add(headers[i]);
        totalSimilarity += match.similarity;
        matchCount++;
        console.log(`[header-mapper] fuzzy matched "${headers[i]}" -> ${field} (${match.similarity}% similarity)`);
    }

    const fuzzyConfidence = matchCount > 0 ? Math.round(totalSimilarity / matchCount) : 0;
    return { mapping, fuzzyConfidence };
}

const SYSTEM_PROMPT = `You are an expert at mapping spreadsheet column headers to a canonical schema.

Given the list of headers from one row, return a JSON object that maps each header to the canonical field name.

Canonical fields:
- buyer_style_number
- buyer_style_name
- sku
- product_description
- color
- color_code
- size
- quantity
- delivery_date
- season
- customer
- factory
- currency
- unit_cost
- po_number
- buyer_po_number
- start_date
- cancel_date
- transport_method
- buy_information

Return ONLY valid JSON with this exact structure (canonical field as key, original header as value):
{
  "mapping": {
    "buyer_style_number": "STYLE#",
    "color": "Colorway Name",
    "quantity": "Total Quantity"
  },
  "confidence": 85,
  "unmappedColumns": ["Some Random Column"]
}

Map headers ONLY when you are confident. If a header does not match any canonical field, omit it from mapping.

No markdown, no explanations, no comments.`;

export async function mapHeaders(headers: string[], brandHint?: string): Promise<HeaderMappingResult> {
    const apiKey = GROQ_API_KEY || process.env.GROQ_API_KEY || '';

    // --- Learning Layer: check header mapping cache ---
    if (brandHint) {
        const cached = await getCachedHeaderMapping(brandHint, headers);
        if (cached && cached.confidence >= 80) {
            console.log(`[header-mapper] Cache hit for brand="${brandHint}" (hits: ${cached.hit_count}, confidence: ${cached.confidence})`);
            const mappedHeaders = new Set(Object.values(cached.mapped_headers));
            const unmapped = headers.filter((h) => !mappedHeaders.has(h));
            return {
                mapping: cached.mapped_headers as ColumnMapping,
                confidence: cached.confidence,
                unmappedColumns: unmapped,
            };
        }
    }
    // --- End Learning Layer ---

    const fallback = fallbackHeuristicMapping(headers);

    // Fuzzy matching layer: catch typos and minor variations that exact
    // heuristic patterns missed. Runs before AI to avoid unnecessary LLM calls.
    const { mapping: fuzzyMapping, fuzzyConfidence } = fuzzyHeuristicMapping(headers, fallback);
    let mapping: ColumnMapping = { ...fallback, ...fuzzyMapping };
    let confidence = 90;
    let unmappedColumns: string[] = [];
    let aiFailed = false;

    const prompt = `${SYSTEM_PROMPT}\n\nHeaders:\n${JSON.stringify(headers)}`;
    const mappedFieldCount = Object.keys(mapping).length;
    const knownLayout = Boolean(
        mapping.buyer_style_number
        && mapping.quantity
        && (mapping.color || mapping.color_code)
        && mappedFieldCount >= 6
    );

    if (knownLayout) {
        console.log(`[header-mapper] known layout mapped deterministically + fuzzy (${mappedFieldCount} fields); skipping LLM`);
    } else {
        try {
            const rawText = await chatWithOllamaJson(SYSTEM_PROMPT, prompt);
            const parsed = parseMappingResponse(rawText);
            mapping = { ...fallback, ...normalizeMapping(parsed.mapping || {}) };
            confidence = Number(parsed.confidence) || 0;
            unmappedColumns = Array.isArray(parsed.unmappedColumns) ? parsed.unmappedColumns : [];
            console.log('[header-mapper] mapped unknown headers with Ollama');
        } catch (err) {
            console.warn('[header-mapper] Ollama mapping failed, trying Groq fallback:', err);
            aiFailed = true;
        }
    }

    if (!knownLayout && Object.keys(mapping).length <= mappedFieldCount && apiKey) {
        try {
            const response = await fetch('https://api.groq.com/openai/v1/chat/completions', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${apiKey}`,
                },
                body: JSON.stringify({
                    model: 'llama-3.3-70b-versatile',
                    messages: [
                        { role: 'system', content: SYSTEM_PROMPT },
                        { role: 'user', content: prompt },
                    ],
                    temperature: 0.1,
                    max_tokens: 512,
                }),
            });

            if (response.ok) {
                const data = await response.json();
                const rawText = data?.choices?.[0]?.message?.content || '';
                const parsed = parseMappingResponse(rawText);

                mapping = { ...fallback, ...normalizeMapping(parsed.mapping || {}) };
                confidence = Number(parsed.confidence) || 0;
                unmappedColumns = Array.isArray(parsed.unmappedColumns) ? parsed.unmappedColumns : [];
            } else {
                const text = await response.text();
                console.warn('[header-mapper] Groq error:', response.status, text);
                aiFailed = true;
            }
        } catch (err) {
            console.warn('[header-mapper] AI mapping failed:', err);
            aiFailed = true;
        }
    } else if (!knownLayout && !Object.keys(mapping).length) {
        console.warn('[header-mapper] GROQ_API_KEY not configured, using heuristic mapping');
        aiFailed = true;
    }

    // Merge with heuristic + fuzzy fallback for any missing canonical fields
    const combinedFallback = { ...fallback, ...fuzzyMapping };
    for (const [field, header] of Object.entries(combinedFallback)) {
        if (!(mapping as Record<string, string>)[field] && header) {
            (mapping as Record<string, string>)[field] = header;
            if (aiFailed) confidence = Math.max(60, fuzzyConfidence || 60);
        }
    }

    // Determine unmapped columns from headers not referenced in mapping
    const mappedHeaders = new Set(Object.values(mapping as Record<string, string>));
    unmappedColumns = headers.filter((h) => !mappedHeaders.has(h));

    // --- Learning Layer: save successful header mapping to cache ---
    if (brandHint && confidence >= 80 && Object.keys(mapping).length >= 6) {
        try {
            await saveHeaderMapping(brandHint, headers, mapping as Record<string, string>, confidence);
        } catch (err) {
            console.warn('[header-mapper] Failed to save header mapping to cache:', err);
        }
    }
    // --- End Learning Layer ---

    return { mapping, confidence, unmappedColumns };
}

function parseMappingResponse(rawText: string): {
    mapping: Record<string, string>;
    confidence: number;
    unmappedColumns: string[];
} {
    try {
        return JSON.parse(rawText);
    } catch {
        try {
            return JSON.parse(jsonrepair(rawText));
        } catch {
            const jsonMatch = rawText.match(/\{[\s\S]*\}/);
            if (!jsonMatch) {
                throw new Error('Could not parse header mapping response');
            }
            return JSON.parse(jsonrepair(jsonMatch[0]));
        }
    }
}
