import { GROQ_API_KEY } from '@/lib/constants';
import { HeaderDetectionResult } from '@/lib/types/buy-file';

const SYSTEM_PROMPT = `You are an expert at reading spreadsheet layouts. Find the row number (1-indexed) that contains the table headers.

Return ONLY valid JSON:
{
  "header_row": 5
}

No markdown, no explanations, no comments.`;

export async function detectHeaderRow(firstRows: unknown[][]): Promise<HeaderDetectionResult> {
    const apiKey = GROQ_API_KEY || process.env.GROQ_API_KEY || '';
    if (!apiKey) {
        console.warn('[header-detector] GROQ_API_KEY not configured; defaulting to row 1');
        return { headerRow: 1 };
    }

    // Limit to first 10 rows for token efficiency
    const previewRows = firstRows.slice(0, 10);
    const rowsText = previewRows
        .map((row, idx) => `Row ${idx + 1}: ${JSON.stringify(row)}`)
        .join('\n');

    const prompt = `${SYSTEM_PROMPT}\n\nWhich row contains the headers?\n\n${rowsText}`;

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

    let parsed: HeaderDetectionResult;
    try {
        parsed = JSON.parse(rawText);
    } catch (err) {
        const jsonMatch = rawText.match(/\{[\s\S]*\}/);
        if (!jsonMatch) {
            console.warn('[header-detector] Could not parse response; defaulting to row 1');
            return { headerRow: 1 };
        }
        parsed = JSON.parse(jsonMatch[0]);
    }

    return {
        headerRow: Math.max(1, Math.min(10, Number(parsed.headerRow) || 1)),
    };
}
