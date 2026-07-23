import { ExtractedTemplate, ColumnMapping } from '@/lib/types/buy-file';

const STORAGE_KEY = 'po_cutting_templates_v1';

function generateUUID(): string {
    if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
        return crypto.randomUUID();
    }
    // Fallback for environments without crypto.randomUUID
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
        const r = (Math.random() * 16) | 0;
        const v = c === 'x' ? r : (r & 0x3) | 0x8;
        return v.toString(16);
    });
}

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

export function loadTemplates(): ExtractedTemplate[] {
    if (typeof window === 'undefined') return [];
    try {
        const raw = window.localStorage.getItem(STORAGE_KEY);
        if (!raw) return [];
        const parsed = JSON.parse(raw);
        return Array.isArray(parsed) ? parsed : [];
    } catch (err) {
        console.error('[template-store] Failed to load templates:', err);
        return [];
    }
}

export function saveTemplates(templates: ExtractedTemplate[]): void {
    if (typeof window === 'undefined') return;
    try {
        window.localStorage.setItem(STORAGE_KEY, JSON.stringify(templates));
    } catch (err) {
        console.error('[template-store] Failed to save templates:', err);
    }
}

export function findMatchingTemplate(
    headers: string[],
    templates: ExtractedTemplate[]
): ExtractedTemplate | null {
    const normalized = normalizeHeaders(headers);
    const headerSet = new Set(normalized.filter(Boolean));

    let best: ExtractedTemplate | null = null;
    let bestScore = 0;

    for (const template of templates) {
        const templateSet = new Set(template.normalizedHeaders.filter(Boolean));
        const matches = normalized.filter((h) => templateSet.has(h)).length;
        const total = Math.max(headerSet.size, templateSet.size);
        if (total === 0) continue;
        const score = matches / total;
        if (score > 0.8 && score > bestScore) {
            bestScore = score;
            best = template;
        }
    }

    return best;
}

export function saveTemplate(
    headers: string[],
    mapping: ColumnMapping,
    customer: string | null
): ExtractedTemplate {
    const templates = loadTemplates();
    const existing = findMatchingTemplate(headers, templates);

    const template: ExtractedTemplate = {
        id: existing?.id || generateUUID(),
        customer,
        headers: headers.map((h) => String(h)),
        normalizedHeaders: normalizeHeaders(headers),
        mapping,
        detectedAt: new Date().toISOString(),
    };

    const filtered = templates.filter((t) => t.id !== template.id);
    filtered.push(template);
    saveTemplates(filtered);
    return template;
}
