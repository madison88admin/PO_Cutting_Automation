import { jsonrepair } from 'jsonrepair';
import { GROQ_API_KEY } from '@/lib/constants';
import { ColumnMapping, HeaderMappingResult } from '@/lib/types/buy-file';

const CANONICAL_FIELDS = new Set([
    'buyer_style_number', 'buyer_style_name', 'sku', 'product_description',
    'color', 'color_code', 'size', 'quantity', 'delivery_date', 'season',
    'customer', 'factory', 'currency', 'unit_cost', 'po_number',
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
    { field: 'po_number', patterns: ['master po', 'po number', 'purchase order', 'order number', 'po no', 'po no.', 'po#', 'so number', 'so#', 'sales order', 'po'] },
    { field: 'buyer_style_number', patterns: ['style#', 'style #', 'style number', 'style no', 'style no.', 'style ref', 'style reference', 'article', 'model', 'model no', 'model number', 'style'] },
    { field: 'buyer_style_name', patterns: ['style name', 'style description', 'style desc', 'style nm', 'style narrative'] },
    { field: 'sku', patterns: ['sku', 'upc', 'ean', 'product code', 'article code', 'old sku', 'eu old sku'] },
    { field: 'product_description', patterns: ['longtext', 'material description', 'product description', 'item description', 'style description', 'description', 'desc'] },
    { field: 'color_code', patterns: ['style color', 'color code', 'colour code', 'colorway code', 'color no', 'color #'] },
    { field: 'color', patterns: ['colorway name', 'color name', 'colour name', 'colorway', 'color', 'colour'] },
    { field: 'size', patterns: ['size 1', 'size 2', 'size name', 'size scale', 'product size', 'size#', 'size'] },
    { field: 'quantity', patterns: ['total quantity', 'order qty', 'po qty', 'buy qty', '1st qty', '2nd qty', 'quantity', 'qty', 'units'] },
    { field: 'delivery_date', patterns: ['vendor confirmed crd', 'brand requested ped', 'planned ped', 'delivery date', 'crdd date', 'ex factory', 'ex-fty', 'ship date', 'delivery', 'crd', 'ped', 'target date'] },
    { field: 'season', patterns: ['season', 'buy season', 'season year', 'year'] },
    { field: 'customer', patterns: ['sold-to party', 'sold to party', 'sold to', 'customer', 'buyer', 'brand', 'sales market', 'sales org', 'company'] },
    { field: 'factory', patterns: ['final factory name', 'final factory', 'final vendor name', 'final vendor', 'factory name', 'factory', 'vendor', 'supplier', 'manufacturer'] },
    { field: 'currency', patterns: ['final currency', 'currency', 'curr'] },
    { field: 'unit_cost', patterns: ['production upcharges usd', 'material upcharges usd', 'unit cost', 'unit price', 'factory cost', 'fob', 'cost', 'price', 'upcharge', 'up charge'] },
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

export async function mapHeaders(headers: string[]): Promise<HeaderMappingResult> {
    const apiKey = GROQ_API_KEY || process.env.GROQ_API_KEY || '';
    let mapping: ColumnMapping = {};
    let confidence = 0;
    let unmappedColumns: string[] = [];
    let aiFailed = false;

    if (apiKey) {
        const prompt = `${SYSTEM_PROMPT}\n\nHeaders:\n${JSON.stringify(headers)}`;

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

                let parsed: { mapping: Record<string, string>; confidence: number; unmappedColumns: string[] };
                try {
                    parsed = JSON.parse(rawText);
                } catch (err) {
                    try {
                        parsed = JSON.parse(jsonrepair(rawText));
                    } catch (repairErr) {
                        const jsonMatch = rawText.match(/\{[\s\S]*\}/);
                        if (jsonMatch) {
                            parsed = JSON.parse(jsonMatch[0]);
                        } else {
                            throw new Error('Could not parse header mapping response');
                        }
                    }
                }

                mapping = normalizeMapping(parsed.mapping || {});
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
    } else {
        console.warn('[header-mapper] GROQ_API_KEY not configured, using heuristic mapping');
        aiFailed = true;
    }

    // Merge with heuristic fallback for any missing canonical fields
    const fallback = fallbackHeuristicMapping(headers);
    for (const [field, header] of Object.entries(fallback)) {
        if (!(mapping as Record<string, string>)[field] && header) {
            (mapping as Record<string, string>)[field] = header;
            if (aiFailed) confidence = 60;
        }
    }

    // Determine unmapped columns from headers not referenced in mapping
    const mappedHeaders = new Set(Object.values(mapping as Record<string, string>));
    unmappedColumns = headers.filter((h) => !mappedHeaders.has(h));

    return { mapping, confidence, unmappedColumns };
}
