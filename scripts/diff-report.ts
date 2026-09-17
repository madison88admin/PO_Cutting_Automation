/**
 * Diff Report — Buy File Upload vs PO Line Data Dump (NextGen ground truth)
 *
 * Compares the ORDERS/LINES rows produced by an uploaded buy file against the
 * NextGen "PO Line Data Dump" Excel export and writes an XLSX report with:
 *
 *   - Matched        : style+color(+PO) found in dump, qty/cost deltas shown
 *   - Missing        : upload lines with NO NextGen counterpart
 *   - Extra          : dump lines with no upload counterpart (never ordered?)
 *   - Cost Mismatch  : matched lines whose cost/qty differs beyond tolerance
 *   - Summary        : counts + totals per PO
 *
 * Usage:
 *   npx tsx scripts/diff-report.ts <buy-file.xlsx> [dump.xlsx]
 *   npx tsx scripts/diff-report.ts "buy-files/Copy of FW2627 BULK BUY 4 - MADISON88 Rev1.xlsx"
 *
 * The dump defaults to ./PO Line Data Dump.xlsx or the path in PO_LINE_DUMP env var.
 * Output: reports/diff-<buyfile>-<timestamp>.xlsx (plus a console summary).
 */

import './load-env'; // MUST be first: loads .env.local before project imports
import path from 'path';
import fs from 'fs';
import ExcelJS from 'exceljs';
import { ExcelEngine, ProcessedPO } from '../src/lib/excel-engine';
import { sizesEquivalent } from '../src/lib/fuzzy';

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

interface DumpLine {
    poNumber: string;        // Order (PO002489-...)
    buyerPoNumber: string;   // Buyer PO Number (M88PO2489-...)
    poLine: string;          // PO Line
    product: string;         // Product (M88128865)
    buyerStyle: string;      // Product Buyer Style Number (2KF3005)
    styleName: string;       // Product Buyer Style Name
    color: string;           // Color ("4363-BLOOM (VIOLET TULIP)")
    size: string;            // Size
    quantity: number;        // Quantity
    purchasePrice: number | null;      // Purchase Price
    linePurchasePrice: number | null;  // Line Purchase Price
    supplier: string;        // Supplier (factory)
    customer: string;        // Customer
    season: string;          // Season
    deliveryDate: string;    // Delivery Date
    status: string;          // Purchase Order Status
    sheetRow: number;
}

interface UploadLine {
    poNumber: string;
    lineItem: number;
    styleNumber: string;     // after NextGen enrichment = M-code
    styleColor: string;      // buyer style + color code concatenated
    colour: string;          // resolved NextGen colour name
    rawColour: string;
    size: string;
    quantity: number;
    cost: string | number | undefined;
    factory: string;
    po: ProcessedPO;
    line: ProcessedPO['lines'][number];
}

interface MatchRow {
    matchType: 'MATCHED' | 'COST_MISMATCH' | 'QTY_MISMATCH' | 'MISSING' | 'EXTRA';
    poNumber: string;
    lineRef: string;
    style: string;
    color: string;
    size: string;
    uploadQty: number | null;
    dumpQty: number | null;
    uploadCost: number | null;
    dumpCost: number | null;
    delta: number | null;
    uploadValue: string;
    dumpValue: string;
    note: string;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function normKey(s: unknown): string {
    return String(s ?? '')
        .toLowerCase()
        .replace(/[^a-z0-9]/g, '')
        .trim();
}

/** "4363-BLOOM (VIOLET TULIP)" → code "4363", name "BLOOM VIOLET TULIP" */
function splitColor(raw: string): { code: string; name: string } {
    const text = String(raw ?? '').trim();
    const m = text.match(/^([A-Z0-9]{2,6})\s*-\s*(.+)$/i);
    if (m) return { code: m[1].toLowerCase(), name: normKey(m[2]) };
    return { code: '', name: normKey(text) };
}

/** Extract the leading color code from an upload colour field. */
function uploadColorCode(line: UploadLine): string {
    const candidates = [line.rawColour, line.styleColor, line.colour];
    for (const value of candidates) {
        const text = String(value ?? '').trim();
        // trailing 3-4 alphanumerics after style, e.g. NF0A8CGZJK3 → JK3
        const m = text.match(/([A-Z0-9]{3,4})$/);
        if (m && /[A-Z]/.test(m[1]) && /\d/.test(m[1] + '0')) return m[1].toLowerCase();
    }
    return normKey(candidates.find((v) => String(v ?? '').trim()) || '');
}

function toNumber(value: unknown): number | null {
    if (value === null || value === undefined || value === '') return null;
    const num = typeof value === 'number' ? value : Number(String(value).replace(/[^0-9.-]/g, ''));
    return Number.isFinite(num) ? num : null;
}

// ---------------------------------------------------------------------------
// Loaders
// ---------------------------------------------------------------------------

async function findDumpFile(explicit?: string): Promise<string> {
    if (explicit) return explicit;
    const envPath = process.env.PO_LINE_DUMP;
    if (envPath && fs.existsSync(envPath)) return envPath;
    const candidates = [
        'PO Line Data Dump.xlsx',
        'po-line-dump.xlsx',
        path.join('buy-files', 'PO Line Data Dump.xlsx'),
    ];
    for (const candidate of candidates) {
        if (fs.existsSync(candidate)) return candidate;
    }
    throw new Error(
        'PO Line Data Dump not found. Pass it as the 2nd argument or set PO_LINE_DUMP env var.'
    );
}

async function loadDumpLines(dumpPath: string): Promise<DumpLine[]> {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(dumpPath);
    const ws = wb.worksheets[0];

    // Header row = row with most string cells among first 5 rows
    let headerRow = 1;
    let bestCount = 0;
    for (let r = 1; r <= Math.min(5, ws.rowCount); r++) {
        let strings = 0;
        ws.getRow(r).eachCell({ includeEmpty: false }, (cell) => {
            if (typeof cell.value === 'string' && cell.value.trim()) strings++;
        });
        if (strings > bestCount) { bestCount = strings; headerRow = r; }
    }

    const headerValues = (ws.getRow(headerRow).values as unknown[])?.slice(1) || [];
    const col = (name: string): number => {
        const target = normKey(name);
        // Exact match first, then fuzzy (dump exports may add suffixes like "New")
        const exact = headerValues.findIndex((h) => normKey(String(h ?? '')) === target);
        if (exact >= 0) return exact;
        return headerValues.findIndex((h) => {
            const normalized = normKey(String(h ?? ''));
            return normalized.startsWith(target) || normalized.includes(target);
        });
    };

    const idx = {
        poNumber: col('Order'),
        buyerPo: col('Buyer PO Number'),
        poLine: col('PO Line'),
        product: col('Product'),
        buyerStyle: col('Product Buyer Style Number'),
        styleName: col('Product Buyer Style Name'),
        color: col('Color'),
        size: col('Size'),
        quantity: col('Quantity'),
        purchasePrice: col('Purchase Price'),
        linePrice: col('Line Purchase Price'),
        supplier: col('Supplier'),
        customer: col('Customer'),
        season: col('Season'),
        delivery: col('Delivery Date'),
        status: col('Purchase Order Status'),
    };

    const required = ['product', 'color', 'quantity'];
    for (const key of required) {
        if (idx[key as keyof typeof idx] < 0) {
            throw new Error(`Dump missing required column "${key}" — check header names`);
        }
    }

    const lines: DumpLine[] = [];
    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
        const row = (ws.getRow(r).values as unknown[])?.slice(1) || [];
        const get = (i: number) => (i >= 0 ? row[i] : undefined);
        const buyerStyle = String(get(idx.buyerStyle) ?? '').trim();
        const color = String(get(idx.color) ?? '').trim();
        const product = String(get(idx.product) ?? '').trim();
        const qty = toNumber(get(idx.quantity));
        if (!buyerStyle && !color && !product) continue; // skip blank rows
        lines.push({
            poNumber: String(get(idx.poNumber) ?? '').trim(),
            buyerPoNumber: String(get(idx.buyerPo) ?? '').trim(),
            poLine: String(get(idx.poLine) ?? '').trim(),
            product: String(get(idx.product) ?? '').trim(),
            buyerStyle,
            styleName: String(get(idx.styleName) ?? '').trim(),
            color,
            size: String(get(idx.size) ?? '').trim(),
            quantity: qty ?? 0,
            purchasePrice: toNumber(get(idx.purchasePrice)),
            linePurchasePrice: toNumber(get(idx.linePrice)),
            supplier: String(get(idx.supplier) ?? '').trim(),
            customer: String(get(idx.customer) ?? '').trim(),
            season: String(get(idx.season) ?? '').trim(),
            deliveryDate: String(get(idx.delivery) ?? '').trim(),
            status: String(get(idx.status) ?? '').trim(),
            sheetRow: r,
        });
    }
    return lines;
}

async function loadUploadLines(buyFilePath: string, manualPo?: string): Promise<UploadLine[]> {
    const engine = new ExcelEngine();
    const buffer = fs.readFileSync(buyFilePath);
    const { data } = await engine.processBuyFile(buffer, {
        sourceFilename: path.basename(buyFilePath),
        manualPurchaseOrder: manualPo || undefined,
    });

    const lines: UploadLine[] = [];
    for (const po of data) {
        for (const line of po.lines) {
            for (const [lineItem, sizes] of Object.entries(po.sizes)) {
                if (Number(lineItem) !== line.lineItem) continue;
                for (const sizeEntry of sizes) {
                    lines.push({
                        poNumber: po.header.purchaseOrder,
                        lineItem: line.lineItem,
                        styleNumber: String(line.styleNumber ?? ''),
                        styleColor: String(line.styleColor ?? ''),
                        colour: String(line.colour ?? ''),
                        rawColour: String(line.rawColour ?? ''),
                        size: sizeEntry.productSize,
                        quantity: sizeEntry.quantity,
                        cost: line.cost,
                        factory: String(po.header.productSupplier ?? ''),
                        po,
                        line,
                    });
                }
            }
        }
    }
    return lines;
}

// ---------------------------------------------------------------------------
// Matching engine
// ---------------------------------------------------------------------------

function styleMatches(uploadLine: UploadLine, dumpLine: DumpLine): boolean {
    const dumpStyle = normKey(dumpLine.buyerStyle);
    const dumpProduct = normKey(dumpLine.product);
    // styleNumber after NextGen enrichment is the M-code; raw styleColor keeps buyer style
    const candidates = [uploadLine.styleColor, uploadLine.styleNumber, uploadLine.colour]
        .map((v) => normKey(String(v)))
        .filter(Boolean);
    return candidates.some((candidate) => {
        if (!candidate) return false;
        return candidate === dumpStyle
            || candidate === dumpProduct
            || (dumpStyle && candidate.includes(dumpStyle))
            || (dumpProduct && candidate.includes(dumpProduct) && dumpProduct.length >= 6);
    });
}

function colorMatches(uploadLine: UploadLine, dumpLine: DumpLine): boolean {
    const dump = splitColor(dumpLine.color);
    const upCode = uploadColorCode(uploadLine);
    const upName = normKey(uploadLine.colour) || normKey(uploadLine.rawColour);
    if (upCode && dump.code && (dump.code === upCode || upCode.includes(dump.code) || dump.code.includes(upCode))) return true;
    if (upName && dump.name && (upName === dump.name || upName.includes(dump.name) || dump.name.includes(upName))) return true;
    // Fall back to full normalized color comparison (handles "ROS-A01 OLIVE SHADOW" on both sides)
    const upFull = normKey(String(uploadLine.colour || uploadLine.rawColour || uploadLine.styleColor || ''));
    const dumpFull = normKey(dumpLine.color);
    if (upFull && dumpFull && (upFull === dumpFull || upFull.includes(dumpFull) || dumpFull.includes(upFull))) return true;
    return false;
}

function sizeMatches(uploadLine: UploadLine, dumpLine: DumpLine): boolean {
    try {
        return sizesEquivalent(uploadLine.size, dumpLine.size);
    } catch {
        return normKey(uploadLine.size) === normKey(dumpLine.size);
    }
}

function buildDiff(uploadLines: UploadLine[], dumpLines: DumpLine[], costTolerance: number) {
    const rows: MatchRow[] = [];
    const usedDump = new Set<number>();

    // Index dump lines by normalized buyer style AND by M-product code
    const byStyle = new Map<string, DumpLine[]>();
    for (let i = 0; i < dumpLines.length; i++) {
        (dumpLines[i] as DumpLine & { _idx?: number })._idx = i;
        const keys = new Set([normKey(dumpLines[i].buyerStyle), normKey(dumpLines[i].product)].filter(Boolean));
        for (const key of keys) {
            const list = byStyle.get(key) || [];
            list.push(dumpLines[i]);
            byStyle.set(key, list);
        }
    }

    for (const up of uploadLines) {
        const candidates = (byStyle.get(normKey(up.styleColor)) || [])
            .concat(byStyle.get(normKey(up.styleNumber)) || []);

        let best: DumpLine | null = null;
        for (const candidate of candidates) {
            if (!styleMatches(up, candidate)) continue;
            if (!colorMatches(up, candidate)) continue;
            if (!sizeMatches(up, candidate)) continue;
            best = candidate;
            break;
        }
        // Relaxed pass: ignore size (some dumps aggregate sizes)
        if (!best) {
            for (const candidate of candidates) {
                if (!styleMatches(up, candidate)) continue;
                if (!colorMatches(up, candidate)) continue;
                best = candidate;
                break;
            }
        }

        if (best) {
            usedDump.add((best as DumpLine & { _idx?: number })._idx ?? -1);
            const upCost = toNumber(up.cost) ?? (best.purchasePrice !== null ? best.purchasePrice : null);
            const dumpCost = best.linePurchasePrice ?? best.purchasePrice;
            const qtyDelta = up.quantity - best.quantity;
            const costDelta = upCost !== null && dumpCost !== null ? upCost - dumpCost : null;

            let matchType: MatchRow['matchType'] = 'MATCHED';
            let note = '';
            if (Math.abs(qtyDelta) > 0) {
                matchType = 'QTY_MISMATCH';
                note = `qty Δ ${qtyDelta > 0 ? '+' : ''}${qtyDelta}`;
            }
            if (costDelta !== null && Math.abs(costDelta) > costTolerance) {
                matchType = matchType === 'MATCHED' ? 'COST_MISMATCH' : matchType;
                note = `${note}${note ? '; ' : ''}cost Δ ${costDelta.toFixed(2)}`;
            }

            rows.push({
                matchType,
                poNumber: up.poNumber,
                lineRef: `L${up.lineItem}`,
                style: up.styleNumber || up.styleColor,
                color: up.colour || up.rawColour,
                size: up.size,
                uploadQty: up.quantity,
                dumpQty: best.quantity,
                uploadCost: upCost,
                dumpCost,
                delta: costDelta,
                uploadValue: up.poNumber,
                dumpValue: best.buyerPoNumber || best.poNumber,
                note,
            });
        } else {
            rows.push({
                matchType: 'MISSING',
                poNumber: up.poNumber,
                lineRef: `L${up.lineItem}`,
                style: up.styleNumber || up.styleColor,
                color: up.colour || up.rawColour,
                size: up.size,
                uploadQty: up.quantity,
                dumpQty: null,
                uploadCost: toNumber(up.cost),
                dumpCost: null,
                delta: null,
                uploadValue: up.poNumber,
                dumpValue: '',
                note: 'no style+color match in dump',
            });
        }
    }

    // Extra: dump lines never used
    for (let i = 0; i < dumpLines.length; i++) {
        if (usedDump.has(i)) continue;
        const dump = dumpLines[i];
        rows.push({
            matchType: 'EXTRA',
            poNumber: dump.poNumber || dump.buyerPoNumber,
            lineRef: dump.poLine || `row${dump.sheetRow}`,
            style: dump.buyerStyle || dump.product,
            color: dump.color,
            size: dump.size,
            uploadQty: null,
            dumpQty: dump.quantity,
            uploadCost: null,
            dumpCost: dump.linePurchasePrice ?? dump.purchasePrice,
            delta: null,
            uploadValue: '',
            dumpValue: dump.buyerPoNumber || dump.poNumber,
            note: 'in dump but not in upload',
        });
    }

    return rows;
}

// ---------------------------------------------------------------------------
// Report writer
// ---------------------------------------------------------------------------

const MATCH_COLORS: Record<MatchRow['matchType'], string> = {
    MATCHED: 'FFC6EFCE',       // green
    COST_MISMATCH: 'FFFFEB9C', // yellow
    QTY_MISMATCH: 'FFFFC7CE',  // red
    MISSING: 'FFFFC7CE',       // red
    EXTRA: 'FFD9E1F2',         // blue
};

async function writeReport(rows: MatchRow[], buyFile: string, dumpFile: string): Promise<string> {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Diff Report');

    const header = [
        'Match Type', 'PO (Upload)', 'Line', 'Style', 'Color', 'Size',
        'Qty (Upload)', 'Qty (NextGen)', 'Cost (Upload)', 'Cost (NextGen)', 'Cost Δ',
        'PO (NextGen)', 'Note',
    ];
    const titleRow = ws.addRow([`Buy File vs NextGen PO Line Dump — ${path.basename(buyFile)}`]);
    titleRow.font = { bold: true, size: 13 };
    ws.addRow([`Dump: ${path.basename(dumpFile)}  ·  Generated: ${new Date().toISOString()}`]);
    ws.addRow([]);
    ws.addRow(header).font = { bold: true };
    header.forEach((_, i) => {
        ws.getCell(4, i + 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
    });

    for (const row of rows) {
        const added = ws.addRow([
            row.matchType, row.poNumber, row.lineRef, row.style, row.color, row.size,
            row.uploadQty, row.dumpQty, row.uploadCost, row.dumpCost,
            row.delta !== null ? Number(row.delta.toFixed(2)) : null,
            row.dumpValue, row.note,
        ]);
        added.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: MATCH_COLORS[row.matchType] } };
    }

    // Summary sheet
    const summary = wb.addWorksheet('Summary');
    summary.addRow(['Match Type', 'Count']).font = { bold: true };
    const counts = rows.reduce<Record<string, number>>((acc, row) => {
        acc[row.matchType] = (acc[row.matchType] || 0) + 1;
        return acc;
    }, {});
    for (const [type, count] of Object.entries(counts)) summary.addRow([type, count]);
    summary.addRow([]);
    const totals = rows.reduce<Record<string, { upQty: number; dumpQty: number }>>((acc, row) => {
        acc[row.poNumber] = acc[row.poNumber] || { upQty: 0, dumpQty: 0 };
        acc[row.poNumber].upQty += row.uploadQty ?? 0;
        acc[row.poNumber].dumpQty += row.dumpQty ?? 0;
        return acc;
    }, {});
    summary.addRow(['PO', 'Qty Upload', 'Qty NextGen']).font = { bold: true };
    for (const [po, t] of Object.entries(totals)) summary.addRow([po, t.upQty, t.dumpQty]);

    ws.autoFilter = { from: { row: 4, column: 1 }, to: { row: 4, column: header.length } };
    ws.columns.forEach((c) => { c.width = 22; });

    const outDir = 'reports';
    fs.mkdirSync(outDir, { recursive: true });
    const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const outFile = path.join(outDir, `diff-${path.basename(buyFile, '.xlsx').replace(/[^a-z0-9]+/gi, '_')}-${stamp}.xlsx`);
    await wb.xlsx.writeFile(outFile);
    return outFile;
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main() {
    const buyFile = process.argv[2];
    if (!buyFile || !fs.existsSync(buyFile)) {
        console.error('Usage: npx tsx scripts/diff-report.ts <buy-file.xlsx> [dump.xlsx] [--po M88PO123]');
        process.exit(1);
    }
    const dumpFile = await findDumpFile(process.argv[3] && !process.argv[3].startsWith('--') ? process.argv[3] : undefined);
    const costTolerance = Number(process.env.DIFF_COST_TOLERANCE || '0.01');
    const poFlagIdx = process.argv.indexOf('--po');
    const manualPo = poFlagIdx >= 0 ? process.argv[poFlagIdx + 1] : undefined;

    console.log(`[diff] buy file : ${buyFile}`);
    console.log(`[diff] dump     : ${dumpFile}`);
    console.log('[diff] loading dump…');
    const dumpLines = await loadDumpLines(dumpFile);
    console.log(`[diff] dump lines loaded: ${dumpLines.length}`);

    console.log('[diff] processing buy file through ExcelEngine…');
    const uploadLines = await loadUploadLines(buyFile, manualPo);
    console.log(`[diff] upload size rows: ${uploadLines.length}`);

    const rows = buildDiff(uploadLines, dumpLines, costTolerance);

    const counts = rows.reduce<Record<string, number>>((acc, row) => {
        acc[row.matchType] = (acc[row.matchType] || 0) + 1;
        return acc;
    }, {});
    console.log('\n===== DIFF SUMMARY =====');
    for (const [type, count] of Object.entries(counts)) console.log(`  ${type.padEnd(14)} ${count}`);

    const outFile = await writeReport(rows, buyFile, dumpFile);
    console.log(`\n[diff] report written: ${outFile}`);
}

main().catch((err) => {
    console.error('[diff] failed:', err);
    process.exit(1);
});
