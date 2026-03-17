import { NextRequest, NextResponse } from "next/server";
import { ExcelEngine, ProcessedPO, POLine, ValidationError } from "@/lib/excel-engine";
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
            if (file.type && !ALLOWED_MIME.has(file.type)) {
                return NextResponse.json({ error: `File ${file.name} has unsupported MIME type ${file.type}.` }, { status: 400 });
            }
        }

        // Create run history record for this upload
        runId = await createRun({
            user_id: userId,
            filename: files.map((f) => f.name).join(", "),
            status: 'Processing'
        });

        for (const file of files) {
            await logEvent({
                eventName: "BUY_FILE_UPLOADED",
                userId,
                runId,
                metadata: { filename: file.name, size: file.size }
            });
        }

        await logEvent({
            eventName: "WORKFLOW_STARTED",
            userId,
            runId,
            metadata: { files: files.map((f) => f.name) }
        });

        const mergedPOsMap = new Map<string, ProcessedPO>();
        const allErrors: ValidationError[] = [];
        const fileSummaries: Array<{ filename: string; orders: number; lines: number; sizes: number; errors: number; warnings: number; }> = [];
        const perFileExports: Record<string, { orders: string; lines: string; sizes: string }> = {};
        const poMap = new Map<string, string>();

        for (const file of files) {
            const buffer = Buffer.from(await file.arrayBuffer());
            const engine = new ExcelEngine(runId || undefined, userId);
            const { data, errors } = await engine.processBuyFile(buffer);

            let allSizes = 0;
            let allLines = 0;
            data.forEach(po => {
                allLines += po.lines.length;
                allSizes += Object.values(po.sizes).reduce((acc, s) => acc + s.length, 0);
            });

            const criticalCount = errors.filter(e => e.severity === 'CRITICAL').length;
            const warningCount = errors.filter(e => e.severity === 'WARNING').length;

            fileSummaries.push({
                filename: file.name,
                orders: data.length,
                lines: allLines,
                sizes: allSizes,
                errors: criticalCount,
                warnings: warningCount,
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

            errors.forEach(e => {
                const copy: ValidationError = { ...e, message: `[${file.name}] ${e.message}` };
                allErrors.push(copy);
            });

            const exported = await engine.generateOutputs(data);
            perFileExports[file.name] = {
                orders: Buffer.from(exported.orders as any).toString('base64'),
                lines: Buffer.from(exported.lines as any).toString('base64'),
                sizes: Buffer.from(exported.sizes as any).toString('base64'),
            };
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
