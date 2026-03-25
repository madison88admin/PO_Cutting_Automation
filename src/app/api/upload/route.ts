import { NextRequest, NextResponse } from "next/server";
import { ExcelEngine, ProcessedPO, POLine, ValidationError, FormatDetection } from "@/lib/excel-engine";
import { logEvent } from "@/lib/audit";
import { createRun, updateRun } from "@/lib/db/runHistory";

const uploadRateLimitMap = new Map<string, number[]>();
const MAX_UPLOADS_PER_MINUTE = 10;
const MAX_FILE_COUNT = 5;
const MAX_FILE_SIZE = 30 * 1024 * 1024; // 30 MB per file
const ALLOWED_MIME = new Set(["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "application/vnd.ms-excel"]);

function checkRateLimit(userId: string): boolean {
    const now = Date.now();
    const windowStart = now - 60000;
    const entries = uploadRateLimitMap.get(userId) || [];
    const recent = entries.filter(ts => ts > windowStart);
    recent.push(now);
    uploadRateLimitMap.set(userId, recent);
    return recent.length <= MAX_UPLOADS_PER_MINUTE;
}

function mergePOs(primary: ProcessedPO, secondary: ProcessedPO): ProcessedPO {
    const merged: ProcessedPO = {
        header: primary.header,
        lines: [...primary.lines],
        sizes: { ...primary.sizes },
    };

    const lineMap = new Map<number, POLine>();
    primary.lines.forEach(line => lineMap.set(line.lineItem, line));
    secondary.lines.forEach(line => {
        const existing = lineMap.get(line.lineItem);
        if (!existing) {
            merged.lines.push(line);
            lineMap.set(line.lineItem, line);
        } else {
            if (!existing.styleNumber && line.styleNumber) existing.styleNumber = line.styleNumber;
            if (existing.colour === "" && line.colour) existing.colour = line.colour;
            if ((existing.cost === undefined || existing.cost === "") && line.cost !== undefined && line.cost !== "") {
                existing.cost = line.cost;
            }
        }
    });

    for (const [lineItem, szs] of Object.entries(secondary.sizes)) {
        const idx = Number(lineItem);
        if (!merged.sizes[idx]) merged.sizes[idx] = [];
        merged.sizes[idx] = [...merged.sizes[idx], ...szs];
    }

    return merged;
}

export async function POST(req: NextRequest) {
    let runId: string | null = null;
    try {
        // Simple auth guard (skipped if no UPLOAD_API_KEY is configured)
        const apiKey = req.headers.get('x-api-key') || '';
        if (process.env.UPLOAD_API_KEY && (!apiKey || apiKey !== process.env.UPLOAD_API_KEY)) {
            return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
        }

        const userId = req.headers.get('x-user-id') || 'public-workflow-user';
        if (!checkRateLimit(userId)) {
            return NextResponse.json({ error: "Rate limit exceeded" }, { status: 429 });
        }

        const formData = await req.formData();
        const files = formData.getAll("file") as File[];
        const manualPo = (formData.get("manualPo")?.toString() || "").trim();
        const manualDestination = (formData.get("manualDestination")?.toString() || "").trim();
        const manualProductRange = (formData.get("manualProductRange")?.toString() || "").trim();
        const manualTemplate = (formData.get("manualTemplate")?.toString() || "").trim();
        const manualLinesTemplate = (formData.get("manualLinesTemplate")?.toString() || "").trim();
        const manualComments = (formData.get("manualComments")?.toString() || "").trim();
        const manualKeyDate = (formData.get("manualKeyDate")?.toString() || "").trim();
        const manualKeyUser1 = (formData.get("manualKeyUser1")?.toString() || "").trim();
        const manualKeyUser2 = (formData.get("manualKeyUser2")?.toString() || "").trim();
        const manualKeyUser3 = (formData.get("manualKeyUser3")?.toString() || "").trim();
        const manualKeyUser4 = (formData.get("manualKeyUser4")?.toString() || "").trim();
        const manualKeyUser5 = (formData.get("manualKeyUser5")?.toString() || "").trim();
        const manualSeason = (formData.get("manualSeason")?.toString() || "").trim();
        const manualCustomer = (formData.get("manualCustomer")?.toString() || "").trim();
        const manualBrand = (formData.get("manualBrand")?.toString() || "").trim();

        if (!files || files.length === 0) {
            return NextResponse.json({ error: "No file uploaded" }, { status: 400 });
        }

        if (files.length > MAX_FILE_COUNT) {
            return NextResponse.json({ error: `Too many files. Maximum is ${MAX_FILE_COUNT}.` }, { status: 400 });
        }

        for (const file of files) {
            if (file.size > MAX_FILE_SIZE) {
                return NextResponse.json({ error: `File ${file.name} exceeds max size of ${MAX_FILE_SIZE} bytes.` }, { status: 400 });
            }
            const ext = file.name.split('.').pop()?.toLowerCase();
            const isExcelExt = ext === 'xlsx' || ext === 'xls';
            if (file.type && !ALLOWED_MIME.has(file.type) && !isExcelExt) {
                return NextResponse.json({ error: `File ${file.name} has unsupported MIME type ${file.type}.` }, { status: 400 });
            }
        }

        // Create run history record for this upload
        runId = await createRun({
            user_id: userId,
            filename: files.map((f) => f.name).join(", "),
            status: 'Processing'
        });

        const fileBuffers: Array<{ file: File; buffer: Buffer }> = [];
        for (const file of files) {
            await logEvent({
                eventName: "BUY_FILE_UPLOADED",
                userId,
                runId,
                metadata: { filename: file.name, size: file.size }
            });
            const buffer = Buffer.from(await file.arrayBuffer());
            fileBuffers.push({ file, buffer });
        }

        await logEvent({
            eventName: "WORKFLOW_STARTED",
            userId,
            runId,
            metadata: { files: files.map((f) => f.name) }
        });

        const mergedPOsMap = new Map<string, ProcessedPO>();
        const allErrors: ValidationError[] = [];
        const fileSummaries: Array<{ filename: string; orders: number; lines: number; sizes: number; errors: number; warnings: number; brands: string[]; }> = [];
        const perFileExports: Record<string, { orders: string; lines: string; sizes: string }> = {};
        const poMap = new Map<string, string>();
        const perFileFormatDetection: Record<string, FormatDetection> = {};

        let productSheetMap: Record<string, any> = {};
        const buyFiles: Array<{ file: File; buffer: Buffer }> = [];

        for (const entry of fileBuffers) {
            const engine = new ExcelEngine(runId || undefined, userId);
            const analysis = await engine.analyzeWorkbook(entry.buffer);
            productSheetMap = { ...productSheetMap, ...analysis.productSheetMap };
            if (analysis.hasBuySheet) {
                buyFiles.push(entry);
            }
        }

        for (const entry of buyFiles) {
            const { file, buffer } = entry;
            const engine = new ExcelEngine(runId || undefined, userId);
            const { data, errors, formatDetection } = await engine.processBuyFile(buffer, {
                manualPurchaseOrder: manualPo || undefined,
                manualDestination: manualDestination || undefined,
                manualProductRange: manualProductRange || undefined,
                manualTemplate: manualTemplate || undefined,
                manualLinesTemplate: manualLinesTemplate || undefined,
                manualComments: manualComments || undefined,
                manualKeyDate: manualKeyDate || undefined,
                manualKeyUser1: manualKeyUser1 || undefined,
                manualKeyUser2: manualKeyUser2 || undefined,
                manualKeyUser3: manualKeyUser3 || undefined,
                manualKeyUser4: manualKeyUser4 || undefined,
                manualKeyUser5: manualKeyUser5 || undefined,
                manualSeason: manualSeason || undefined,
                manualCustomer: manualCustomer || undefined,
                manualBrand: manualBrand || undefined,
                defaultQuantityIfMissing: !!manualPo,
                productSheetMap,
            });

            let allSizes = 0;
            let allLines = 0;
            data.forEach((po: ProcessedPO) => {
                allLines += po.lines.length;
                allSizes += Object.values(po.sizes).reduce((acc, s) => acc + s.length, 0);
            });

            const criticalCount = errors.filter((e: ValidationError) => e.severity === 'CRITICAL').length;
            const warningCount = errors.filter((e: ValidationError) => e.severity === 'WARNING').length;

            const brandSet = new Set<string>();
            data.forEach((po: ProcessedPO) => {
                const label = (po.header?.customer || '').trim();
                if (label) brandSet.add(label);
            });

            fileSummaries.push({
                filename: file.name,
                orders: data.length,
                lines: allLines,
                sizes: allSizes,
                errors: criticalCount,
                warnings: warningCount,
                brands: Array.from(brandSet),
            });

            for (const po of data) {
                // Cross-file duplicate PO detection
                const existingFile = poMap.get(po.header.purchaseOrder);
                if (existingFile && existingFile !== file.name) {
                    allErrors.push({
                        field: 'PurchaseOrder', row: 1,
                        message: `[${file.name}] Duplicate PO ${po.header.purchaseOrder} also appears in ${existingFile}.`,
                        severity: 'CRITICAL'
                    });
                }
                poMap.set(po.header.purchaseOrder, file.name);

                const existing = mergedPOsMap.get(po.header.purchaseOrder);
                if (!existing) {
                    mergedPOsMap.set(po.header.purchaseOrder, po);
                } else {
                    const merged = mergePOs(existing, po);
                    mergedPOsMap.set(po.header.purchaseOrder, merged);
                }
            }

            errors.forEach((e: ValidationError) => {
                const copy: ValidationError = { ...e, message: `[${file.name}] ${e.message}` };
                allErrors.push(copy);
            });

            const exported = await engine.generateOutputs(data);
            perFileExports[file.name] = {
                orders: Buffer.from(exported.orders as any).toString('base64'),
                lines: Buffer.from(exported.lines as any).toString('base64'),
                sizes: Buffer.from(exported.sizes as any).toString('base64'),
            };
            if (formatDetection) {
                perFileFormatDetection[file.name] = formatDetection;
            }
        }

        const mergedData = Array.from(mergedPOsMap.values());
        const engine = new ExcelEngine(runId || undefined, userId);
        const outputs = await engine.generateOutputs(mergedData);

        const hasCritical = allErrors.some(e => e.severity === "CRITICAL");

        if (runId) {
            await updateRun(runId, {
            status: hasCritical ? 'Validation Failed' : 'Pending Review',
            error_count: allErrors.filter(e => e.severity === 'CRITICAL').length,
            warning_count: allErrors.filter(e => e.severity === 'WARNING').length,
            orders_rows: mergedData.length,
            lines_rows: mergedData.reduce((a, p) => a + p.lines.length, 0),
            order_sizes_rows: mergedData.reduce((a, p) => a + Object.values(p.sizes).reduce((b, s) => b + s.length, 0), 0),
            completed_at: new Date().toISOString(),
            });
        }

        await logEvent({
            eventName: "DATA_EXTRACTION_COMPLETE",
            userId,
            runId: runId || undefined,
            metadata: { rows_extracted: mergedData.length, file_count: files.length }
        });

        const filesOut = {
            orders: Buffer.from(outputs.orders as any).toString('base64'),
            lines: Buffer.from(outputs.lines as any).toString('base64'),
            sizes: Buffer.from(outputs.sizes as any).toString('base64')
        };

        const summary = fileSummaries;
        return NextResponse.json({
            success: true,
            runId,
            dataCount: mergedData.length,
            errors: allErrors,
            canProceed: !hasCritical,
            files: filesOut,
            fileSummary: summary,
            mergedSummary: {
                orders: mergedData.length,
                lines: mergedData.reduce((a, p) => a + p.lines.length, 0),
                sizes: mergedData.reduce((a, p) => a + Object.values(p.sizes).reduce((b, s) => b + s.length, 0), 0),
                errors: allErrors.filter(e => e.severity === 'CRITICAL').length,
                warnings: allErrors.filter(e => e.severity === 'WARNING').length,
            },
            fileOutputs: perFileExports,
            formatDetection: perFileFormatDetection,
        });

    } catch (error: any) {
        console.error("Upload error:", error);
        if (runId) {
            await updateRun(runId, {
                status: 'Validation Failed',
                error_count: 1,
                warning_count: 0,
                completed_at: new Date().toISOString(),
            });
        }
        return NextResponse.json({ error: "Internal Server Error" }, { status: 500 });
    }
}
