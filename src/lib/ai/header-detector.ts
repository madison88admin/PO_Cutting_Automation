import { GROQ_API_KEY } from '@/lib/constants';
import { chatWithOllamaJson } from '@/lib/ai/ollama-client';
import { HeaderDetectionResult } from '@/lib/types/buy-file';

const SYSTEM_PROMPT = `You are an expert at reading spreadsheet layouts. Find the row number (1-indexed) that contains the table headers.

Return ONLY valid JSON:
{
  "header_row": 5
}

No markdown, no explanations, no comments.`;

export async function detectHeaderRow(firstRows: unknown[][]): Promise<HeaderDetectionResult> {
    // Limit to first 10 rows for token efficiency
    const previewRows = firstRows.slice(0, 10);
    const deterministicRow = detectKnownHeaderRow(previewRows);
    if (deterministicRow) {
        console.log('[header-detector] deterministic header row:', deterministicRow);
        return { headerRow: deterministicRow };
    }

    const rowsText = previewRows
        .map((row, idx) => `Row ${idx + 1}: ${JSON.stringify(row)}`)
        .join('\n');

    const prompt = `${SYSTEM_PROMPT}\n\nWhich row contains the headers?\n\n${rowsText}`;

    try {
        const rawText = await chatWithOllamaJson(SYSTEM_PROMPT, prompt);
        const parsed = parseHeaderDetection(rawText);
        console.log('[header-detector] Ollama selected row:', parsed.headerRow);
        return parsed;
    } catch (error) {
        console.warn('[header-detector] Ollama failed, trying Groq fallback:', error);
    }

    const apiKey = GROQ_API_KEY || process.env.GROQ_API_KEY || '';
    if (!apiKey) {
        console.warn('[header-detector] GROQ_API_KEY not configured; defaulting to row 1');
        return { headerRow: 1 };
    }

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
            max_tokens: 128,
        }),
    });

    if (!response.ok) {
        const text = await response.text();
        console.warn('[header-detector] Groq error:', response.status, text, '- defaulting to row 1');
        return { headerRow: 1 };
    }

    const data = await response.json();
    const rawText = data?.choices?.[0]?.message?.content || '';

    try {
        return parseHeaderDetection(rawText);
    } catch (err) {
        console.warn('[header-detector] Could not parse response; defaulting to row 1');
        return { headerRow: 1 };
    }
}

function detectKnownHeaderRow(rows: unknown[][]): number | null {
    const headerTerms = [
        'style', 'buyer style', 'product name', 'purchase order', 'po number',
        'quantity', 'total quantity', 'color', 'colour', 'size', 'season',
        'customer', 'factory', 'vendor', 'currency', 'cost', 'sell',
        'delivery date', 'crd', 'decision', 'material',
    ];
    let bestRow = 0;
    let bestScore = 0;

    rows.forEach((row, index) => {
        let score = 0;
        for (const value of row || []) {
            const normalized = String(value || '')
                .toLowerCase()
                .replace(/[^a-z0-9]+/g, ' ')
                .replace(/\s+/g, ' ')
                .trim();
            if (!normalized) continue;
            if (headerTerms.some((term) => normalized === term || normalized.includes(term))) {
                score += 2;
            }
        }
        if (score > bestScore) {
            bestScore = score;
            bestRow = index + 1;
        }
    });

    return bestScore >= 4 ? bestRow : null;
}

function parseHeaderDetection(rawText: string): HeaderDetectionResult {
    const jsonMatch = rawText.match(/\{[\s\S]*\}/);
    const parsed = JSON.parse(jsonMatch?.[0] || rawText) as HeaderDetectionResult & { header_row?: number };
    const selectedRow = Number(parsed.headerRow || parsed.header_row || 1);
    return {
        headerRow: Math.max(1, Math.min(10, selectedRow)),
    };
}
