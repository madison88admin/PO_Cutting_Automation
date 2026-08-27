type Field = string;

const NUMERIC_FIELDS = new Set(['quantity', 'unit_cost']);
const DATE_FIELDS = new Set(['delivery_date', 'start_date', 'cancel_date']);
const REQUIRED_FIELDS = ['buyer_style_number', 'color', 'size', 'quantity'];

const text = (value: unknown) => String(value ?? '').trim();
const numeric = (value: unknown) => {
    const v = text(value).replace(/[$,]/g, '');
    return v !== '' && Number.isFinite(Number(v));
};
const dateLike = (value: unknown) => {
    if (value instanceof Date) return !Number.isNaN(value.getTime());
    const v = text(value);
    return !!v && (!Number.isNaN(Date.parse(v)) || /^\d{4,5}(?:\.\d+)?$/.test(v));
};

export interface SemanticValidation {
    fieldConfidence: Record<string, number>;
    warnings: string[];
}

/** Validate proposed mappings against real sample values below the headers. */
export function validateHeaderSemantics(
    headers: string[],
    mapping: Record<string, string>,
    sampleRows: unknown[][],
): SemanticValidation {
    const index = new Map(headers.map((header, i) => [header, i]));
    const scores: Record<string, number> = {};
    const warnings: string[] = [];
    for (const [field, header] of Object.entries(mapping)) {
        const column = index.get(header);
        if (column === undefined) continue;
        const values = sampleRows.map(row => row[column]).map(text).filter(Boolean);
        if (!values.length) { scores[field] = 45; warnings.push(`${header} has no sample values to validate ${field}.`); continue; }
        let score = 70;
        if (NUMERIC_FIELDS.has(field)) {
            const ratio = values.filter(numeric).length / values.length;
            score = Math.round(ratio * 100);
        } else if (DATE_FIELDS.has(field)) {
            const ratio = values.filter(dateLike).length / values.length;
            score = Math.round(ratio * 100);
        } else if (field === 'color') {
            const ratio = values.filter(v => /[A-Za-z]/.test(v) && !/^\d+(?:[.,]\d+)?$/.test(v)).length / values.length;
            score = Math.round(ratio * 100);
        } else if (field === 'size') {
            const ratio = values.filter(v => /^(?:xxs|xs|s|m|l|xl|xxl|one size|\d{1,2}(?:\.\d)?)$/i.test(v)).length / values.length;
            score = Math.round(ratio * 100);
        } else if (field === 'buyer_style_number' || field === 'sku') {
            const unique = new Set(values).size;
            score = Math.min(100, 70 + (unique / values.length) * 30);
        }
        scores[field] = score;
        if (score < 60) warnings.push(`${header} may not represent ${field} (sample validation ${score}%).`);
    }
    for (const field of REQUIRED_FIELDS) {
        if (!mapping[field]) warnings.push(`Required field is not mapped: ${field}.`);
    }
    const mappedHeaders = Object.entries(mapping).filter(([, h]) => h).map(([, h]) => h);
    const duplicates = mappedHeaders.filter((header, i) => mappedHeaders.indexOf(header) !== i);
    for (const header of new Set(duplicates)) warnings.push(`${header} is assigned to more than one target field.`);
    return { fieldConfidence: scores, warnings: Array.from(new Set(warnings)) };
}
