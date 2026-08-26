// Regression harness: run every workbook in buy-files/ through the
// deterministic extraction pipeline and report mapping gaps per brand.
// Usage: npx tsx regression-buy-files.ts [--with-nextgen]
import { readFileSync, readdirSync } from 'fs';
import path from 'path';

const envContent = readFileSync('.env.local', 'utf-8');
for (const line of envContent.split('\n')) {
    const match = line.match(/^\s*([A-Z0-9_]+)\s*=\s*(.*)\s*$/);
    if (match) process.env[match[1]] = match[2];
}

const withNextgen = process.argv.includes('--with-nextgen');
if (!withNextgen) process.env.NEXTGEN_ENABLED = 'false';

const BRAND_HINTS: [RegExp, string][] = [
    [/arcteryx/i, 'arcteryx'],
    [/dynafit/i, 'dynafit'],
    [/fh26 jan buy/i, 'jack wolfskin'],
    [/marmot/i, 'marmot'],
    [/feb buy top up/i, 'vans'],
    [/fjall/i, 'fjallraven'],
    [/fw2627 bulk buy/i, 'rossignol'],
    [/smartwool/i, 'smartwool'],
    [/w26_december order/i, 'mammut'],
    [/llb/i, 'll bean'],
];

interface FileReport {
    file: string;
    hint: string;
    error?: string;
    headerRow?: number;
    sheetName?: string;
    templateUsed?: boolean;
    itemCount?: number;
    productCount?: number;
    unmapped?: string[];
    mappingKeys?: string[];
    quality?: Record<string, number>;
    sample?: Record<string, unknown>;
}

async function main() {
    const { extractBuyFile } = await import('./src/lib/buy-file-extractor');
    const dir = path.resolve('buy-files');
    const files = readdirSync(dir).filter((f) => /\.xlsx$/i.test(f)).sort();
    const reports: FileReport[] = [];

    for (const file of files) {
        const hintMatch = BRAND_HINTS.find(([re]) => re.test(file));
        const hint = hintMatch ? hintMatch[1] : '';
        const report: FileReport = { file, hint };
        try {
            const buf = readFileSync(path.join(dir, file));
            const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
            const result = await extractBuyFile(ab, hint || undefined);
            report.sheetName = result.items[0]?.sourceSheet;
            report.headerRow = result.headerRow;
            report.templateUsed = result.templateUsed;
            report.itemCount = result.items.length;
            report.productCount = result.productData.length;
            report.unmapped = result.unmappedColumns;
            report.mappingKeys = Object.entries(result.mapping)
                .filter(([, v]) => v)
                .map(([k, v]) => `${k}<-${v}`);

            const n = result.items.length || 1;
            const pct = (fn: (it: (typeof result.items)[number]) => boolean) =>
                Math.round((result.items.filter(fn).length / n) * 100);
            report.quality = {
                style: pct((i) => !!i.style),
                color: pct((i) => !!(i.colorCode || i.colorName || i.color)),
                size: pct((i) => !!i.size),
                qtyPos: pct((i) => (i.quantity ?? 0) > 0),
                po: pct((i) => !!i.poNumber),
                date: pct((i) => !!i.deliveryDate),
                cost: pct((i) => i.unitCost != null),
                factory: pct((i) => !!i.factory),
                customer: pct((i) => !!i.customer),
            };
            const s = result.items[0];
            report.sample = s
                ? {
                      style: s.style,
                      color: s.colorCode || s.color,
                      colorName: s.colorName,
                      size: s.size,
                      qty: s.quantity,
                      po: s.poNumber,
                      brand: s.brand,
                      customer: s.customer,
                      factory: s.factory,
                  }
                : undefined;
        } catch (e) {
            report.error = e instanceof Error ? e.message : String(e);
        }
        reports.push(report);
        console.log(JSON.stringify(report, null, 1));
    }

    console.log('\n===== SUMMARY =====');
    for (const r of reports) {
        if (r.error) {
            console.log(`FAIL ${r.file}: ${r.error}`);
        } else {
            const q = r.quality!;
            console.log(
                `${r.hint || '(no hint)'} | ${r.file} | items=${r.itemCount} prods=${r.productCount} ` +
                    `tpl=${r.templateUsed} style=${q.style}% colour=${q.color}% size=${q.size}% qty=${q.qtyPos}% po=${q.po}% ` +
                    `unmapped=[${(r.unmapped || []).slice(0, 6).join('; ')}]`
            );
        }
    }
}

main().catch((e) => {
    console.error(e);
    process.exit(1);
});
