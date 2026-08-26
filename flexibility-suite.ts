// Flexibility suite: generates synthetic buy files with UNSEEN layouts/brands
// and checks whether the extraction pipeline adapts without per-brand rules.
import { readFileSync, writeFileSync, mkdirSync, rmSync } from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';

process.env.NEXTGEN_ENABLED = 'false';

const OUT = path.resolve('buy-files-flex');
rmSync(OUT, { recursive: true, force: true });
mkdirSync(OUT, { recursive: true });

type Row = (string | number | null)[];

async function make(name: string, sheets: { name: string; rows: Row[] }[]) {
    const wb = new ExcelJS.Workbook();
    for (const s of sheets) {
        const ws = wb.addWorksheet(s.name);
        for (const r of s.rows) ws.addRow(r);
    }
    const buf = await wb.xlsx.writeBuffer();
    writeFileSync(path.join(OUT, name), Buffer.from(buf));
}

// Canonical semantic order used by every synthetic layout below:
// [po, styleCode, colour, styleName, size, qty, date]
const DATA: Row[] = [
    ['PO-9001', 'QQ10001', 'Navy', 'Alpine Shell Jacket', 'S', 120, '2027-03-15'],
    ['PO-9001', 'QQ10001', 'Navy', 'Alpine Shell Jacket', 'M', 340, '2027-03-15'],
    ['PO-9001', 'QQ10002', 'Red', 'Ridge Beanie', 'OS', 75, '2027-03-20'],
    ['PO-9002', 'QQ10003', 'Black', 'Trail Glove', 'L', 210, '2027-04-01'],
];

async function main() {
    // 1. Unknown brand, mildly renamed headers
    await make('flex1-renamed-headers.xlsx', [
        { name: 'Sheet1', rows: [['PO Num', 'Style Code', 'Colour Way', 'Item Name', 'Size', 'Qty', 'Ship Date'], ...DATA] },
    ]);

    // 2. Banner rows above the real header (row 4)
    await make('flex2-banner-row4.xlsx', [
        { name: 'Buy', rows: [
            ['ACME SPORTSWEAR — FW27 BUY SHEET'],
            ['CONFIDENTIAL — INTERNAL USE ONLY'],
            [],
            ['Purchase Order', 'Article', 'Colorway', 'Description', 'Grid', 'Units', 'Ex-Fac Date'], ...DATA,
        ] },
    ]);

    // 3. Junk summary tab + real order tab
    await make('flex3-junk-summary-tab.xlsx', [
        { name: 'Summary', rows: [['Totals Report'], ['Region', 'Sum of Qty'], ['US', 5000], ['EU', 3000]] },
        { name: 'Order Details', rows: [['PO Number', 'Style Number', 'Colour', 'Size Name', 'Quantity', 'Delivery Date'], ...DATA.map(r => [r[0], r[1], r[2], r[4], r[5], r[6]])] },
    ]);

    // 4. Gibberish headers — content must speak for itself (AI fallback territory)
    await make('flex4-gibberish-headers.xlsx', [
        { name: 'Sheet1', rows: [['Col A', 'Field 2', 'Data X', 'Misc 9', 'Val ZZ', 'Num 7', 'Info 3'], ...DATA] },
    ]);

    // 5. Minimal layout — no PO column at all
    await make('flex5-minimal-no-po.xlsx', [
        { name: 'Sheet1', rows: [['Style', 'Color', 'Size', 'Qty', 'Date'], ...DATA.map(r => [r[1], r[2], r[4], r[5], r[6]])] },
    ]);

    // 6. Case/space/punctuation chaos
    await make('flex6-messy-formatting.xlsx', [
        { name: 'Sheet1', rows: [[' P.O.# ', 'STYLE_NUMBER', 'colour-code', 'Product Desc.', 'Size(N)', 'Quantity (PCS)', 'Delivery__Date'], ...DATA] },
    ]);

    // 7. Two competing buy sheets — documents which one wins
    await make('flex7-two-buy-sheets.xlsx', [
        { name: 'WHOLESALE', rows: [['PO', 'Style', 'Color', 'Size', 'Qty'], ['WS-1', 'QQ20001', 'Blue', 'M', 10]] },
        { name: 'E-COMM', rows: [['PO', 'Style', 'Color', 'Size', 'Qty'], ['EC-1', 'QQ30001', 'Green', 'L', 20]] },
    ]);

    // 8. Foreign-language headers (German) with clean data
    await make('flex8-german-headers.xlsx', [
        { name: 'Tabelle1', rows: [['Bestellung', 'Stilnummer', 'Farbe', 'Beschreibung', 'Größe', 'Menge', 'Lieferdatum'], ...DATA] },
    ]);

    const { extractBuyFile } = await import('./src/lib/buy-file-extractor');
    const files = ['flex1-renamed-headers.xlsx', 'flex2-banner-row4.xlsx', 'flex3-junk-summary-tab.xlsx', 'flex4-gibberish-headers.xlsx', 'flex5-minimal-no-po.xlsx', 'flex6-messy-formatting.xlsx', 'flex7-two-buy-sheets.xlsx', 'flex8-german-headers.xlsx'];

    interface Verdict { file: string; sheet?: string; headerRow?: number; items: number; fill: Record<string, number>; error?: string }
    const results: Verdict[] = [];
    for (const f of files) {
        const v: Verdict = { file: f, items: 0, fill: {} };
        try {
            const buf = readFileSync(path.join(OUT, f));
            const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
            const r = await extractBuyFile(ab);
            v.sheet = r.items[0]?.sourceSheet;
            v.headerRow = r.headerRow;
            v.items = r.items.length;
            const n = r.items.length || 1;
            const pct = (fn: (i: (typeof r.items)[number]) => boolean) => Math.round((r.items.filter(fn).length / n) * 100);
            v.fill = {
                style: pct((i) => !!i.style),
                colour: pct((i) => !!(i.colorCode || i.colorName || i.color)),
                size: pct((i) => !!i.size),
                qty: pct((i) => (i.quantity ?? 0) > 0),
                po: pct((i) => !!i.poNumber),
                date: pct((i) => !!i.deliveryDate),
            };
            console.log(`\n=== ${f}`);
            console.log(`sheet="${v.sheet}" headerRow=${v.headerRow} items=${v.items} mapping=${JSON.stringify(r.mapping)}`);
        } catch (e) {
            v.error = e instanceof Error ? e.message : String(e);
            console.log(`\n=== ${f}\nCRASH: ${v.error}`);
        }
        results.push(v);
    }

    console.log('\n================= VERDICTS =================');
    for (const v of results) {
        const q = v.fill;
        if (v.error) { console.log(`FAIL    ${v.file}: ${v.error}`); continue; }
        const core = Math.min(q.style ?? 0, q.qty ?? 0);
        const grade = v.items === 0 ? 'FAIL' : core >= 90 ? 'PASS ' : core >= 50 ? 'PARTIAL' : 'FAIL';
        console.log(`${grade}  ${v.file} | sheet=${v.sheet} hr=${v.headerRow} items=${v.items} style=${q.style}% colour=${q.colour}% size=${q.size}% qty=${q.qty}% po=${q.po}% date=${q.date}%`);
    }
    process.exit(0);
}

main().catch((e) => { console.error(e); process.exit(1); });
