/** True when a value is only a numeric code, not a human-readable colour. */
export function isNumericOnlyColor(value: unknown): boolean {
    const text = String(value ?? '').trim();
    return text.length > 0 && /^[+-]?\d+(?:[.,]\d+)?$/.test(text);
}

/**
 * Select the first usable display colour. Numeric-only values remain valid as
 * matching codes elsewhere, but must never be emitted as a colour name.
 */
export function pickDisplayColor(...values: unknown[]): string {
    for (const value of values) {
        const text = String(value ?? '').trim();
        if (text && !isNumericOnlyColor(text)) return text;
    }
    return '';
}
