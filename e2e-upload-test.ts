// End-to-end production-path test: dev server /api/upload (ExcelEngine
// pipeline) with LIVE NextGen matching, then validate the returned
// ORDERS/LINES/SIZES workbooks by re-opening them with ExcelJS.
import { readFileSync, writeFileSync, mkdirSync } from 'fs';
import path from 'path';

const BASE = 'http://localhost:3000';
const OUT_DIR = path.resolve('e2e-outputs');
try { mkdirSync(OUT_DIR, { recursive: true }); } catch {}

const CASES = [
    { file: 'F26 COL BULK 03.04 Buy - FTY (1).xlsx', brand: 'columbia' },
    { file: 'Copy of Buy W26_DYNAFIT_BUY2_GPO0004127.xlsx', brand: 'dynafit' },
    { file: 'Copy of LLB 2926.xlsx', brand: 'llbean' },
];

interface CaseResult {
    brand: string;
    httpOk: boolean;
    canProceed?: boolean;
    errorCount?: number;
    criticalCount?: number;
    warningCount?: number;
    ordersRows?: number;
    linesRows?: number;
    sizesRows?: number;
    nexgenSummary?: unknown;
    needsAttention?: string[];
    outputKeys?: string[];
    sampleOrdersRow?: string[];
    criticalSamples?: string[];
    filesSaved?: string[];
    error?: string;
}

async function runCase(c: { file: string; brand: string }): Promise<CaseResult> {
    const res: CaseResult = { brand: c.brand, httpOk: false };
    try {
        const buf = readFileSync(c.file.includes('/') ? c.file : path.join('buy-files', c.file));
        const form = new FormData();
        form.append('file', new Blob([new Uint8Array(buf)], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        }), c.file);
        form.append('manualPo', `E2E-${c.brand.toUpperCase()}-001`);
        const resp = await fetch(`${BASE}/api/upload`, {
            method: 'POST',
            body: form,
            headers: { 'x-user-id': 'e2e-flex-test' },
            signal: AbortSignal.timeout(570000),
        });
        res.httpOk = resp.ok;
        if (!resp.ok) {
            res.error = `HTTP ${resp.status}: ${(await resp.text()).slice(0, 300)}`;
            return res;
        }
        const data: any = await resp.json();
        res.canProceed = data.canProceed;
        const errs: any[] = data.errors || [];
        res.errorCount = errs.length;
        res.criticalCount = errs.filter((e) => String(e.severity).toUpperCase() === 'CRITICAL').length;
        res.warningCount = errs.filter((e) => String(e.severity).toUpperCase() === 'WARNING').length;
        res.outputKeys = Object.keys(data);
        res.nexgenSummary = data.nexgenVariantSummary ?? null;
        res.needsAttention = (data.needsAttention || []).map((n: any) => n.code || n.reason || '?').slice(0, 8);
        res.criticalSamples = errs.filter((e) => String(e.severity).toUpperCase() === 'CRITICAL').slice(0, 3)
            .map((e) => `${e.field || ''}${e.row ? '@r' + e.row : ''}: ${e.message}`.trim());

        // decode + validate outputs
        const wbMod = await import('exceljs');
        const ExcelJS = wbMod.default || wbMod;
        const countRows = async (b64: string | undefined, name: string) => {
            if (!b64) return undefined;
            const fp = path.join(OUT_DIR, `${c.brand}-${name}.xlsx`);
            writeFileSync(fp, Buffer.from(b64, 'base64'));
            res.filesSaved = [...(res.filesSaved || []), path.basename(fp)];
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.readFile(fp);
            const ws = wb.worksheets[0];
            return ws.rowCount - 1; // minus header
        };
        res.ordersRows = await countRows(data.orders || data.files?.orders, 'ORDERS');
        res.linesRows = await countRows(data.lines || data.files?.lines, 'LINES');
        res.sizesRows = await countRows(data.sizes || data.files?.sizes, 'ORDER_SIZES');

        const ordersB64 = data.orders || data.files?.orders;
        if (ordersB64) {
            const fp = path.join(OUT_DIR, `${c.brand}-ORDERS.xlsx`);
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.readFile(fp);
            const ws = wb.worksheets[0];
            const row = ws.getRow(2);
            const cells: string[] = [];
            row.eachCell({ includeEmpty: false }, (cell, col) => {
                if (col <= 8) cells.push(String(cell.value ?? '').slice(0, 24));
            });
            res.sampleOrdersRow = cells;
        }
    } catch (e) {
        res.error = e instanceof Error ? e.message : String(e);
    }
    return res;
}

async function main() {
    for (const c of CASES) {
        console.log(`\n===== E2E ${c.brand} (${c.file})`);
        const t0 = Date.now();
        const r = await runCase(c);
        const secs = ((Date.now() - t0) / 1000).toFixed(1);
        console.log(JSON.stringify(r, null, 1));
        console.log(`[${c.brand}] done in ${secs}s`);
    }
    process.exit(0);
}

main().catch((e) => { console.error(e); process.exit(1); });
