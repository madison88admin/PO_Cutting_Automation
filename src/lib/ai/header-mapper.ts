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
    'buy_information', 'status', 'transport_location',
]);

/**
 * A mapping is only worth caching/reusing when it binds the core trio:
 * a quantity source, a style/sku reference, and some colour signal.
 * Guards against poisoning the learning cache with half-mappings that
 * would otherwise bypass improved deterministic logic on later runs.
 */
function hasCoreFields(m: Record<string, string>): boolean {
    return Boolean(
        m.quantity
        && (m.buyer_style_number || m.sku)
        && (m.color || m.color_code)
    );
}

function normalizeMapping(mapping: Record<string, string>): ColumnMapping {    const normalized: ColumnMapping = {};
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
    { field: 'po_number', patterns: ['final po cut', 'master po', 'po number', 'purchase order no', 'purchase order number', 'purchasing document', 'purchasing document number', 'purchase order', 'order number', 'po no', 'po no.', 'po#', 'so number', 'so#', 'sales order', 'po cut', 'po', 'tracking no', 'tracking number', 'buy 1 - tracking no', 'buy 2 - tracking no', 'bestellung', 'auftragsnummer', 'auftrag', 'bestellnummer', 'col a'] },
    { field: 'buyer_po_number', patterns: ['buyer po number', 'buyer po #', 'buyer po', 'customer po number', 'customer po #', 'customer po', 'bp no', 'extraction po #'] },
    { field: 'buyer_style_number', patterns: ['jde style', 'material style', 'style#', 'style #', 'style code', 'style number', 'style no', 'style no.', 'style ref', 'style reference', 'item number', 'item#', 'article', 'article code', 'model', 'model no', 'model number', 'model code', 'dev #', 'item', 'style', 'material', 'stilnummer', 'artikelnummer', 'artikel code', 'modell', 'field 2'] },
    { field: 'buyer_style_name', patterns: ['style name', 'style description', 'style desc', 'style nm', 'style narrative', 'article full colors', 'season plan model commercial name', 'stilname', 'misc 9'] },
    { field: 'sku', patterns: ['sku', 'upc', 'ean', 'product code', 'article code', 'style color article code (12 digits)', 'style color size article code (18 digits)', 'old sku', 'eu old sku', 'item code', 'material', 'vendor sku', 'merch - sku'] },
    { field: 'product_description', patterns: ['longtext', 'material description', 'product description', 'item description', 'style description', 'product name', 'short text', 'description', 'desc', 'text', 'beschreibung'] },
    { field: 'color_code', patterns: ['style color', 'color code', 'colour code', 'colorway code', 'color no', 'color #', 'primary color peak pdm code', 'colorway', 'farbcode'] },
    { field: 'color', patterns: ['colorway name', 'color name', 'colour name', 'colorway', 'colour way', 'color description', 'color', 'colour', 'merch - color', 'farbe', 'data x'] },
    { field: 'size', patterns: ['grid value', 'size 1', 'size 2', 'size name', 'size scale', 'product size', 'size#', 'size', 'dimension', 'merch - size', 'groesse', 'größe', 'val zz'] },
    { field: 'quantity', patterns: ['total quantity', 'scheduled quantity', 'tot qty', 'order qty', 'ordered qty', 'po qty', 'buy qty', '1st qty', '2nd qty', 'quantity|n', 'qty (lum)', 'consensus quantity', 'quantity', 'new qty', 'qty', 'units', 'bulk qty', 'sb = eb qty', 'final po qty', 'buy 1 agreed qty', 'buy 2 agreed qty', 'menge', 'anzahl', 'stueckzahl', 'stückzahl', 'num 7'] },
    { field: 'delivery_date', patterns: ['confirmed fty ex fac', 'vendor confirmed crd', 'brand requested ped', 'planned ped', 'delivery date', 'crdd date', 'ex factory', 'confirmed ex-factory date|n', 'requested etd|n', 'final delivery date', 'final xf date', 'best crd', 'brand requested crd', 'material arrival date', 'shipping date', 'handover date ordered', 'vendor confirmed etd', 'orig ex fac', 'ex-factory', 'ex-fty', 'ship date', 'etd', 'delivery', 'crd', 'ped', 'target date', 'date', 'exf date', 'cfm crd', 'lieferdatum', 'liefertermin', 'info 3'] },
    { field: 'start_date', patterns: ['udf-start_date', 'start date', 'order start date', 'valid from'] },
    { field: 'cancel_date', patterns: ['udf-canel_date', 'udf-cancel_date', 'cancel date', 'canel date', 'order cancel date', 'valid until'] },
    { field: 'season', patterns: ['season code', 'season', 'buy season', 'season year', 'year', 'saison'] },
    { field: 'transport_location', patterns: ['transport location', 'transportlocation', 'ult. destination', 'destination name', 'destination', 'dest country', 'ship to country', 'ship to', 'ship-to party name', 'country/region', 'final destination', 'country'] },
    { field: 'customer', patterns: ['sold-to party', 'sold to party', 'sold to', 'customer name', 'customer/market', 'customer', 'buyer', 'brand', 'sales market', 'sales org', 'company', 'kunde'] },
    { field: 'factory', patterns: ['final factory name', 'final factory', 'final vendor name', 'final vendor', 'factory name', 'erp factory code', 'factory code', 'vendor code', 'factory', 'vendor plnt (conf plnt)', 'confirmed vendor plnt', 'vendor name', 'vendor', 'supplier', 'manufacturer', 'production supplier name', 'fabrik', 'lieferant'] },
    { field: 'currency', patterns: ['final currency', 'currency', 'curr', 'waehrung', 'währung'] },
    { field: 'unit_cost', patterns: ['fob', 'unit cost', 'unit price', 'factory cost', 'net price', 'base fob m88', 'base fob', 'cost', 'price', 'production upcharges usd', 'material upcharges usd', 'upcharge', 'up charge', 'kosten', 'preis'] },
    { field: 'transport_method', patterns: ['order transport', 'transport method', 'transportation mode', 'transportation mode description', 'transport mode', 'shipment method', 'shipment mode', 'shipping method', 'ship mode', 'ship via', 'mode of delivery|n', 'mode of delivery', 'trans cond', 'freight mode', 'transportart'] },
    { field: 'buy_information', patterns: ['buy information', 'buy info', 'buying information', 'buying info', 'purchase information', 'purchase info', 'po information', 'po info', 'buy details', 'buy detail'] },
    { field: 'status', patterns: ['status', 'po status', 'order status', 'line status', 'workflow status', 'decision'] },
];

function normalizeHeaderKey(s: string): string {
    return s.toLowerCase().replace(/[^a-z0-9|]+/g, ' ').replace(/\s+/g, ' ').trim();
}

/**
 * NEW-vs-OLD quantity resolution (e.g. Haglofs "NEW QTY" vs "OLD QTY"):
 * the NEW quantity is the buy quantity; OLD/previous is history and must
 * never back the quantity field. Fixes mappings regardless of source
 * (heuristic, AI, template, or cache). Idempotent.
 */
export function isOldQtyHeader(header: string): boolean {
    const h = String(header || '').trim();
    return /^(old|previous|prior)\b/i.test(h) && /\b(qty|quantity)\b/i.test(h);
}

export function isNewQtyHeader(header: string): boolean {
    const h = String(header || '').trim();
    return /\bnew\b/i.test(h) && /\b(qty|quantity)\b/i.test(h);
}

export function resolveNewVsOldQuantity(headers: string[], mapping: ColumnMapping): void {
    const m = mapping as Record<string, string>;
    if (!m.quantity || !isOldQtyHeader(m.quantity)) return;
    const used = new Set(Object.values(m));
    const target = headers.find((h) => isNewQtyHeader(h) && !used.has(h))
        || headers.find((h) => isNewQtyHeader(h) && h !== m.quantity);
    if (target) {
        console.log(`[header-mapper] quantity "${m.quantity}" -> "${target}" (NEW wins over OLD)`);
        m.quantity = target;
    }
}

function patternToRegex(pattern: string): RegExp {
    // Normalize both sides consistently (punctuation collapses to spaces) then
    // allow optional whitespace between tokens so "PO #" == "PO#" == "po".
    // Regex specials (including the literal "|" used by headers like
    // "Quantity|N") must be escaped, otherwise they become alternations.
    const normalized = normalizeHeaderKey(pattern);
    const body = normalized
        .split(' ')
        .map((tok) => tok
            .split('')
            .map((ch) => ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'))
            .join('\\s*'))
        .join('\\s*');
    return new RegExp(`^${body}$`, 'i');
}

export function fallbackHeuristicMapping(headers: string[]): ColumnMapping {
    const mapping: ColumnMapping = {};
    const usedHeaders = new Set<number>();
    const normalizedHeaders = headers.map((h) => normalizeHeaderKey(h));

    for (const { field, patterns } of HEADER_PATTERNS) {
        for (const pattern of patterns) {
            const regex = patternToRegex(pattern);
            for (let i = 0; i < normalizedHeaders.length; i++) {
                if (usedHeaders.has(i)) continue;
                if (!normalizedHeaders[i]) continue;
                if (regex.test(normalizedHeaders[i])) {
                    (mapping as Record<string, string>)[field] = headers[i];
                    usedHeaders.add(i);
                    break;
                }
            }
            if ((mapping as Record<string, string>)[field]) break;
        }
        // Substring fallback: if header CONTAINS the pattern as token (e.g. "SB = EB QTY" contains "qty")
        if ((mapping as Record<string, string>)[field]) continue;
        for (const pattern of patterns) {
            const normPat = normalizeHeaderKey(pattern);
            // Only use substring for short generic keywords to avoid false positives
            if (normPat.length < 3) continue;
            // Require whole-token containment: pad with spaces
            const tokenRegex = new RegExp(`\\b${normPat.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`, 'i');
            for (let i = 0; i < normalizedHeaders.length; i++) {
                if (usedHeaders.has(i)) continue;
                if (!normalizedHeaders[i]) continue;
                if (tokenRegex.test(normalizedHeaders[i])) {
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

const SYSTEM_PROMPT = `You are an expert garment PO header mapper. Map ANY spreadsheet headers — English, German, or even gibberish — to canonical fields by using BOTH header names AND sample data values.

Canonical fields and what they look like:
- po_number: PO identifiers like "PO-9001", "4500106757", "263B108B" (often contains "PO", dashes, or is a tracking number)
- buyer_style_number: style codes like "QQ10001", "G81212010", "RLOMH02", "VN000QB4GRK" (alphanumeric, no spaces, often has letters+digits)
- buyer_style_name: descriptive names like "Alpine Shell Jacket", "Ridge Beanie" (words with spaces)
- sku: material/sku codes, often same as style or longer (12-18 digits)
- product_description: long text descriptions
- color / color_code: color values like "Navy", "Red", "2N3", "ARC-Aquaculture"
- size: size values like "S", "M", "L", "OS", "One Size", "1-SIZ"
- quantity: numeric quantities like "120", "340" (positive integers)
- delivery_date: dates like "2027-03-15", "Mon Jul 06 2026"
- season, customer, factory, currency, unit_cost, transport_method, status, transport_location, etc.
- transport_location: destination codes/names like "CA", "US", "Germany", "South Ontario DC" (headers: Destination, Dest Country, Ship To, Country)
- customer: ONLY the brand/buyer name (headers: Customer, Buyer, Brand, Sold-To). NEVER map Destination/Ship-To/Country columns to customer.

Rules:
- Destination/Ship-To/Country columns are ALWAYS transport_location, NEVER customer or factory — even if their values look like country codes.
- Headers may be in German: Bestellung=po_number, Stilnummer=buyer_style_number, Farbe=color, Größe=size, Menge=quantity, Lieferdatum=delivery_date
- Headers may be meaningless gibberish like "Col A", "Field 2" — then INFER from sample data values!
- Use sample data row values to verify/disambiguate. A column containing "PO-9001" is po_number even if header is "Col A".
- Return ONLY valid JSON: {"mapping":{"buyer_style_number":"STYLE#","quantity":"Total Qty"},"confidence":85,"unmappedColumns":[]}
- If unsure, omit the field. No markdown, no explanation.`;

function buildMappingPrompt(headers: string[], sampleRows?: unknown[][]): string {
    let prompt = `Headers: ${JSON.stringify(headers)}`;
    if (sampleRows && sampleRows.length > 0) {
        // Send 3 sample rows — CRITICAL for gibberish/foreign headers
        const preview = sampleRows.slice(0, 3).map((row: unknown[], idx: number) => {
            const cells = headers.map((h: string, i: number) => `${h}: ${row[i] ?? ''}`);
            return `Row${idx + 1}: ${cells.join(' | ')}`;
        });
        prompt += `\nSample data (use this to infer when headers are gibberish or foreign):\n${preview.join('\n')}`;
        prompt += `\n\nInstructions: Map each header to its canonical field. For gibberish headers, look at the sample values: e.g. a column with "PO-9001" is po_number, "QQ10001" is buyer_style_number, "Navy" is color, "120" is quantity, "2027-03-15" is delivery_date.`;
    }
    return prompt;
}

export async function mapHeaders(headers: string[], brandHint?: string, sampleRows?: unknown[][]): Promise<HeaderMappingResult> {
    const apiKey = GROQ_API_KEY || process.env.GROQ_API_KEY || '';

    // --- Learning Layer: check header mapping cache ---
    if (brandHint) {
        const cached = await getCachedHeaderMapping(brandHint, headers);
        const cachedMap = (cached?.mapped_headers || {}) as Record<string, string>;
        if (cached && cached.confidence >= 80 && hasCoreFields(cachedMap)) {
            console.log(`[header-mapper] Cache hit for brand="${brandHint}" (hits: ${cached.hit_count}, confidence: ${cached.confidence})`);
            const mappedHeaders = new Set(Object.values(cached.mapped_headers));
            const unmapped = headers.filter((h) => !mappedHeaders.has(h));
            return {
                mapping: cached.mapped_headers as ColumnMapping,
                confidence: cached.confidence,
                unmappedColumns: unmapped,
            };
        }
        if (cached) {
            console.log(`[header-mapper] cache entry for brand="${brandHint}" lacks core fields; recomputing`);
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

    const prompt = buildMappingPrompt(headers, sampleRows);
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
                    model: process.env.GROQ_MODEL || 'openai/gpt-oss-20b',
                    messages: [
                        { role: 'system', content: SYSTEM_PROMPT },
                        { role: 'user', content: prompt },
                    ],
                    temperature: 0.1,
                    max_tokens: 2048,
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

    // NEW wins over OLD for quantity (covers AI answers too)
    resolveNewVsOldQuantity(headers, mapping);

    // Determine unmapped columns from headers not referenced in mapping
    const mappedHeaders = new Set(Object.values(mapping as Record<string, string>));
    unmappedColumns = headers.filter((h) => !mappedHeaders.has(h));

    // --- Learning Layer: save successful header mapping to cache ---
    if (brandHint && confidence >= 80 && Object.keys(mapping).length >= 6 && hasCoreFields(mapping as Record<string, string>)) {
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
    const empty = { mapping: {}, confidence: 0, unmappedColumns: [] };
    try {
        return JSON.parse(rawText);
    } catch {
        try {
            return JSON.parse(jsonrepair(rawText));
        } catch {
            const jsonMatch = rawText.match(/\{[\s\S]*\}/);
            if (!jsonMatch) {
                console.warn('[header-mapper] Could not parse AI response (no JSON found):', rawText.substring(0, 200));
                return empty;
            }
            try {
                return JSON.parse(jsonrepair(jsonMatch[0]));
            } catch {
                console.warn('[header-mapper] Could not parse AI response (jsonrepair failed):', jsonMatch[0].substring(0, 200));
                return empty;
            }
        }
    }
}
