import { NextRequest, NextResponse } from "next/server";
import fs from "fs";
import ExcelJS from "exceljs";
import { ExcelEngine, ProcessedPO } from "@/lib/excel-engine";
import { sizesEquivalent } from "@/lib/fuzzy";

/**
 * POST /api/diff-report — Buy file vs NextGen PO Line Data Dump diff.
 *
 * multipart/form-data:
 *   file          (required) buy file xlsx
 *   dump          (optional) PO Line Data Dump xlsx; falls back to PO_LINE_DUMP env path
 *   po            (optional) manual PO number for buy files with no PO column
 *   costTolerance (optional) cost delta tolerance, default 0.01
 *
 * Response: summary counts + per-line diff rows (no xlsx generation; use
 * scripts/diff-report.ts for the Excel report).
 */

interface DumpLine {
    poNumber: string;
    buyerPoNumber: string;
    poLine: string;
    product: string;
    buyerStyle: string;
    color: string;
    size: string;
    quantity: number;
    purchasePrice: number | null;
    linePurchasePrice: number | null;
    supplier: string;
    sheetRow: number;
    _idx?: number;
}

interface UploadLine {
    poNumber: string;
    lineItem: number;
    styleNumber: string;
    styleColor: string;
    colour: string;
    rawColour: string;
    size: string;
    quantity: number;
    cost: string | number | undefined;
}

interface MatchRow {
    matchType: "MATCHED" | "COST_MISMATCH" | "QTY_MISMATCH" | "MISSING" | "EXTRA";
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
    note: string;
}

function normKey(s: unknown): string {
    return String(s ?? "").toLowerCase().replace(/[^a-z0-9]/g, "").trim();
}

function splitColor(raw: string): { code: string; name: string } {
    const text = String(raw ?? "").trim();
    const m = text.match(/^([A-Z0-9]{2,6})\s*-\s*(.+)$/i);
    if (m) return { code: m[1].toLowerCase(), name: normKey(m[2]) };
    return { code: "", name: normKey(text) };
}

function uploadColorCode(line: UploadLine): string {
    for (const value of [line.rawColour, line.styleColor, line.colour]) {
        const text = String(value ?? "").trim();
        const m = text.match(/([A-Z0-9]{3,4})$/);
        if (m && /[A-Z]/.test(m[1]) && /\d/.test(m[1] + "0")) return m[1].toLowerCase();
    }
    return normKey([line.rawColour, line.styleColor, line.colour].find((v) => String(v ?? "").trim()) || "");
}

function toNumber(value: unknown): number | null {
    if (value === null || value === undefined || value === "") return null;
    const num = typeof value === "number" ? value : Number(String(value).replace(/[^0-9.-]/g, ""));
    return Number.isFinite(num) ? num : null;
}

async function loadDumpLines(dumpBuffer: ArrayBuffer): Promise<DumpLine[]> {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(dumpBuffer);
    const ws = wb.worksheets[0];
    if (!ws) throw new Error("Dump file has no worksheets");

    let headerRow = 1;
    let bestCount = 0;
    for (let r = 1; r <= Math.min(5, ws.rowCount); r++) {
        let strings = 0;
        ws.getRow(r).eachCell({ includeEmpty: false }, (cell) => {
            if (typeof cell.value === "string" && cell.value.trim()) strings++;
        });
        if (strings > bestCount) { bestCount = strings; headerRow = r; }
    }

    const headerValues = ((ws.getRow(headerRow).values as unknown[]) || []).slice(1) as string[];
    const col = (name: string): number => {
        const target = normKey(name);
        const exact = headerValues.findIndex((h) => normKey(String(h ?? "")) === target);
        if (exact >= 0) return exact;
        return headerValues.findIndex((h) => {
            const normalized = normKey(String(h ?? ""));
            return normalized.startsWith(target) || normalized.includes(target);
        });
    };

    const idx = {
        poNumber: col("Order"),
        buyerPo: col("Buyer PO Number"),
        poLine: col("PO Line"),
        product: col("Product"),
        buyerStyle: col("Product Buyer Style Number"),
        color: col("Color"),
        size: col("Size"),
        quantity: col("Quantity"),
        purchasePrice: col("Purchase Price"),
        linePrice: col("Line Purchase Price"),
        supplier: col("Supplier"),
    };
    if (idx.product < 0 || idx.color < 0 || idx.quantity < 0) {
        throw new Error('Dump missing required columns ("Product", "Color", "Quantity")');
    }

    const lines: DumpLine[] = [];
    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
        const row = ((ws.getRow(r).values as unknown[]) || []).slice(1);
        const get = (i: number) => (i >= 0 ? row[i] : undefined);
        const product = String(get(idx.product) ?? "").trim();
        const buyerStyle = String(get(idx.buyerStyle) ?? "").trim();
        const color = String(get(idx.color) ?? "").trim();
        if (!product && !buyerStyle && !color) continue;
        lines.push({
            poNumber: String(get(idx.poNumber) ?? "").trim(),
            buyerPoNumber: String(get(idx.buyerPo) ?? "").trim(),
            poLine: String(get(idx.poLine) ?? "").trim(),
            product,
            buyerStyle,
            color,
            size: String(get(idx.size) ?? "").trim(),
            quantity: toNumber(get(idx.quantity)) ?? 0,
            purchasePrice: toNumber(get(idx.purchasePrice)),
            linePurchasePrice: toNumber(get(idx.linePrice)),
            supplier: String(get(idx.supplier) ?? "").trim(),
            sheetRow: r,
        });
    }
    return lines;
}

async function loadUploadLines(buffer: Buffer, manualPo?: string, filename?: string): Promise<UploadLine[]> {
    const engine = new ExcelEngine();
    const { data } = await engine.processBuyFile(buffer, {
        sourceFilename: filename || "upload.xlsx",
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
                        styleNumber: String(line.styleNumber ?? ""),
                        styleColor: String(line.styleColor ?? ""),
                        colour: String(line.colour ?? ""),
                        rawColour: String(line.rawColour ?? ""),
                        size: sizeEntry.productSize,
                        quantity: sizeEntry.quantity,
                        cost: line.cost,
                    });
                }
            }
        }
    }
    return lines;
}

function styleMatches(up: UploadLine, dump: DumpLine): boolean {
    const dumpStyle = normKey(dump.buyerStyle);
    const dumpProduct = normKey(dump.product);
    const candidates = [up.styleColor, up.styleNumber, up.colour].map((v) => normKey(v)).filter(Boolean);
    return candidates.some((candidate) =>
        candidate === dumpStyle
        || candidate === dumpProduct
        || (dumpStyle && candidate.includes(dumpStyle))
        || (dumpProduct && candidate.includes(dumpProduct) && dumpProduct.length >= 6)
    );
}

function colorMatches(up: UploadLine, dump: DumpLine): boolean {
    const d = splitColor(dump.color);
    const upCode = uploadColorCode(up);
    const upName = normKey(up.colour) || normKey(up.rawColour);
    if (upCode && d.code && (d.code === upCode || upCode.includes(d.code) || d.code.includes(upCode))) return true;
    if (upName && d.name && (upName === d.name || upName.includes(d.name) || d.name.includes(upName))) return true;
    const upFull = normKey(String(up.colour || up.rawColour || up.styleColor || ""));
    const dumpFull = normKey(dump.color);
    if (upFull && dumpFull && (upFull === dumpFull || upFull.includes(dumpFull) || dumpFull.includes(upFull))) return true;
    return false;
}

function sizeMatches(up: UploadLine, dump: DumpLine): boolean {
    try {
        return sizesEquivalent(up.size, dump.size);
    } catch {
        return normKey(up.size) === normKey(dump.size);
    }
}

function buildDiff(uploadLines: UploadLine[], dumpLines: DumpLine[], costTolerance: number): MatchRow[] {
    const rows: MatchRow[] = [];
    const usedDump = new Set<number>();

    const byStyle = new Map<string, DumpLine[]>();
    for (let i = 0; i < dumpLines.length; i++) {
        dumpLines[i]._idx = i;
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
        if (!best) {
            for (const candidate of candidates) {
                if (!styleMatches(up, candidate)) continue;
                if (!colorMatches(up, candidate)) continue;
                best = candidate;
                break;
            }
        }

        if (best) {
            usedDump.add(best._idx ?? -1);
            const upCost = toNumber(up.cost);
            const dumpCost = best.linePurchasePrice ?? best.purchasePrice;
            const qtyDelta = up.quantity - best.quantity;
            const costDelta = upCost !== null && dumpCost !== null ? upCost - dumpCost : null;

            let matchType: MatchRow["matchType"] = "MATCHED";
            const notes: string[] = [];
            if (qtyDelta !== 0) { matchType = "QTY_MISMATCH"; notes.push(`qty Δ ${qtyDelta > 0 ? "+" : ""}${qtyDelta}`); }
            if (costDelta !== null && Math.abs(costDelta) > costTolerance) {
                if (matchType === "MATCHED") matchType = "COST_MISMATCH";
                notes.push(`cost Δ ${costDelta.toFixed(2)}`);
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
                note: notes.join("; "),
            });
        } else {
            rows.push({
                matchType: "MISSING",
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
                note: "no style+color match in dump",
            });
        }
    }

    for (let i = 0; i < dumpLines.length; i++) {
        if (usedDump.has(i)) continue;
        const dump = dumpLines[i];
        rows.push({
            matchType: "EXTRA",
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
            note: "in dump but not in upload",
        });
    }

    return rows;
}

export async function POST(req: NextRequest) {
    try {
        const formData = await req.formData();
        const buyFile = formData.get("file");
        if (!(buyFile instanceof File)) {
            return NextResponse.json({ error: "No buy file provided" }, { status: 400 });
        }

        const manualPo = String(formData.get("po") || "").trim();
        const costTolerance = Number(formData.get("costTolerance")) || 0.01;

        // Dump source: uploaded file, or env-configured path
        let dumpBuffer: ArrayBuffer | null = null;
        const dumpFile = formData.get("dump");
        if (dumpFile instanceof File) {
            dumpBuffer = await dumpFile.arrayBuffer();
        } else {
            const dumpPath = process.env.PO_LINE_DUMP || "PO Line Data Dump.xlsx";
            if (fs.existsSync(dumpPath)) {
                dumpBuffer = fs.readFileSync(dumpPath).buffer as ArrayBuffer;
            }
        }
        if (!dumpBuffer) {
            return NextResponse.json(
                { error: "No dump provided: upload a 'dump' file or set PO_LINE_DUMP env var" },
                { status: 400 }
            );
        }

        const buyBuffer = Buffer.from(await buyFile.arrayBuffer());
        const [dumpLines, uploadLines] = await Promise.all([
            loadDumpLines(dumpBuffer),
            loadUploadLines(buyBuffer, manualPo, buyFile.name),
        ]);

        if (!uploadLines.length) {
            return NextResponse.json(
                { error: "Buy file produced no lines — check the file or pass a manual 'po'" },
                { status: 422 }
            );
        }

        const rows = buildDiff(uploadLines, dumpLines, costTolerance);
        const summary = rows.reduce<Record<string, number>>((acc, row) => {
            acc[row.matchType] = (acc[row.matchType] || 0) + 1;
            return acc;
        }, {});

        return NextResponse.json({
            success: true,
            buyFile: buyFile.name,
            dumpLines: dumpLines.length,
            uploadLines: uploadLines.length,
            summary,
            // Cap the payload: with a 28K-row dump, EXTRA rows dominate.
            rows: rows.filter((r) => r.matchType !== "EXTRA").slice(0, 5000),
            extraCount: summary.EXTRA || 0,
        });
    } catch (error) {
        const message = error instanceof Error ? error.message : "Unknown error";
        console.error("[diff-report] error:", message);
        return NextResponse.json({ error: message }, { status: 500 });
    }
}
