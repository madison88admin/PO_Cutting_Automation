import { NextRequest, NextResponse } from "next/server";
import { ExcelEngine, ProcessedPO, POLine, ValidationError, FormatDetection } from "@/lib/excel-engine";
import { logEvent } from "@/lib/audit";
import { createRun, updateRun } from "@/lib/db/runHistory";

const uploadRateLimitMap = new Map<string, number[]>();
const MAX_UPLOADS_PER_MINUTE = 10;
const MAX_FILE_COUNT = 5;
const MAX_FILE_SIZE = 30 * 1024 * 1024; // 30 MB per file
const ALLOWED_MIME = new Set(["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "application/vnd.ms-excel"]);
const FALLBACK_USER_ID = "00000000-0000-0000-0000-000000000001";

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
        orderKeys: [...(primary.orderKeys || [])],
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
                if (!existing.styleColor && line.styleColor) existing.styleColor = line.styleColor;
                if (!existing.ourReference && line.ourReference) existing.ourReference = line.ourReference;
                if (!existing.startDate && line.startDate) existing.startDate = line.startDate;
                if (!existing.cancelDate && line.cancelDate) existing.cancelDate = line.cancelDate;
                if (!existing.hhStartDate && line.hhStartDate) existing.hhStartDate = line.hhStartDate;
                if (!existing.hhCancelDate && line.hhCancelDate) existing.hhCancelDate = line.hhCancelDate;
                if (!existing.hhConfirmedDeliveryDate && line.hhConfirmedDeliveryDate) existing.hhConfirmedDeliveryDate = line.hhConfirmedDeliveryDate;
                if (!existing.transportLocation && line.transportLocation) existing.transportLocation = line.transportLocation;
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

    if (secondary.orderKeys && secondary.orderKeys.length > 0) {
        const existingKeys = new Set((merged.orderKeys || []).map(k => `${k.purchaseOrder}||${k.customer}||${k.transportLocation}`));
        secondary.orderKeys.forEach(key => {
            const signature = `${key.purchaseOrder}||${key.customer}||${key.transportLocation}`;
            if (!existingKeys.has(signature)) {
                merged.orderKeys = merged.orderKeys || [];
                merged.orderKeys.push(key);
                existingKeys.add(signature);
            }
        });
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

        const userIdHeader = req.headers.get('x-user-id') || '';
        const userId = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(userIdHeader)
            ? userIdHeader
            : FALLBACK_USER_ID;
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
        const effectiveManualProductRange = manualProductRange || manualSeason;
        const manualCustomer = (formData.get("manualCustomer")?.toString() || "").trim();
        const manualBrand = (formData.get("manualBrand")?.toString() || "").trim();
        const inferredManualCustomer = manualCustomer || (files.some((f) => /vuori/i.test(f.name)) ? "Vuori" : "");

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
        const referenceSizesBuffers: Buffer[] = [];
        for (const file of files) {
            await logEvent({
                eventName: "BUY_FILE_UPLOADED",
                userId,
                runId,
                metadata: { filename: file.name, size: file.size }
            });
            const buffer = Buffer.from(await file.arrayBuffer());
            fileBuffers.push({ file, buffer });
            const lowerName = file.name.toLowerCase();
            if (lowerName.includes('sizes') && !lowerName.includes('buy') && !lowerName.includes('product shi')) {
                referenceSizesBuffers.push(buffer);
            }
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
            const lowerName = entry.file.name.toLowerCase();
            const isProductShiReference = lowerName.includes('product shi');
            const isReferenceSizes = lowerName.includes('sizes') && !lowerName.includes('buy') && !lowerName.includes('product shi');
            if (analysis.hasBuySheet && !isReferenceSizes && !isProductShiReference) {
                buyFiles.push(entry);
            }
        }

        for (const entry of buyFiles) {
            const { file, buffer } = entry;
            const engine = new ExcelEngine(runId || undefined, userId);
            const { data, errors, formatDetection } = await engine.processBuyFile(buffer, {
                manualPurchaseOrder: manualPo || undefined,
                manualDestination: manualDestination || undefined,
                manualProductRange: effectiveManualProductRange || undefined,
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
                manualCustomer: inferredManualCustomer || undefined,
                manualBrand: manualBrand || undefined,
                defaultQuantityIfMissing: !!manualPo,
                productSheetMap,
                llBeanReferenceSizesBuffer: referenceSizesBuffers[0],
                sourceFilename: file.name,
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
