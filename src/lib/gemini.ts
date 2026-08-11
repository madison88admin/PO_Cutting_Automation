/**
 * Gemini OCR Client
 * Uses Google Gemini API for extracting structured PO data from images/PDFs
 */

interface GeminiConfig {
    apiKey: string;
    model: string;
}

export interface GeminiOCRResult {
    poNumber: string;
    style: string;
    color: string;
    size: string;
    quantity: number;
    factory: string;
    customer: string;
    season: string;
    exFtyDate: string;
    transportMethod: string;
    plant: string;
}

export class GeminiOCRClient {
    private config: GeminiConfig;

    constructor(config?: Partial<GeminiConfig>) {
        this.config = {
            apiKey: config?.apiKey || process.env.GEMINI_API_KEY || '',
            model: config?.model || 'gemini-2.5-flash',
        };
    }

    async extractFromFile(base64File: string, mimeType: string): Promise<GeminiOCRResult[]> {
        if (!this.config.apiKey) {
            throw new Error('GEMINI_API_KEY is not configured');
        }

        const url = `https://generativelanguage.googleapis.com/v1beta/models/${this.config.model}:generateContent?key=${this.config.apiKey}`;

        const prompt = `Extract all purchase order line items from this document. Return ONLY a JSON array of objects with these fields:
- poNumber
- style
- color
- size
- quantity (number)
- factory
- customer
- season
- exFtyDate
- transportMethod
- plant

If a field is not present, use empty string or 0. Return valid JSON only. Do not include markdown or explanation.`;

        const body = {
            contents: [
                {
                    parts: [
                        { text: prompt },
                        {
                            inline_data: {
                                mime_type: mimeType,
                                data: base64File,
                            },
                        },
                    ],
                },
            ],
            generationConfig: {
                responseMimeType: 'application/json',
                temperature: 0.1,
            },
        };

        const response = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body),
        });

        if (!response.ok) {
            const text = await response.text();
            throw new Error(`Gemini API error: ${response.status} ${text}`);
        }

        const data = await response.json();
        const rawText = data?.candidates?.[0]?.content?.parts?.[0]?.text || '';
        if (!rawText) {
            throw new Error('No response from Gemini OCR');
        }

        let parsed: GeminiOCRResult[];
        try {
            parsed = JSON.parse(rawText);
        } catch {
            const jsonMatch = rawText.match(/\[[\s\S]*\]/);
            if (!jsonMatch) throw new Error('Could not parse Gemini OCR response');
            parsed = JSON.parse(jsonMatch[0]);
        }

        return parsed.map((item) => ({
            poNumber: String(item.poNumber || ''),
            style: String(item.style || ''),
            color: String(item.color || ''),
            size: String(item.size || ''),
            quantity: Number(item.quantity) || 0,
            factory: String(item.factory || ''),
            customer: String(item.customer || ''),
            season: String(item.season || ''),
            exFtyDate: String(item.exFtyDate || ''),
            transportMethod: String(item.transportMethod || ''),
            plant: String(item.plant || ''),
        }));
    }
}
