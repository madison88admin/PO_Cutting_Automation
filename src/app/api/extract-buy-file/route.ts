import { NextRequest, NextResponse } from "next/server";
import { extractBuyFile } from "@/lib/buy-file-extractor";
import { NextGenCachedClient } from "@/lib/nextgen/client";
import { ExcelEngine } from "@/lib/excel-engine";
import { generateOrdersSheet } from "@/lib/generator/orders";
import { generateLinesSheet } from "@/lib/generator/lines";
import { generateSizesSheet } from "@/lib/generator/sizes";

async function workbookToBase64(workbook: any): Promise<string> {
    const buffer = await workbook.xlsx.writeBuffer();
    return Buffer.from(buffer as any).toString('base64');
}

export async function POST(req: NextRequest) {
    const timers: Record<string, number> = {};
    const start = (label: string) => { timers[label] = Date.now(); };
    const end = (label: string) => {
        const elapsed = Date.now() - (timers[label] || Date.now());
        console.log(`[extract-buy-file] ${label}: ${elapsed}ms`);
        return elapsed;
    };

    try {
        start('total');
        const formData = await req.formData();
        const allFiles = [
            ...(formData.get("file") ? [formData.get("file") as File] : []),
            ...(formData.getAll("files") as File[]),
        ].filter((f): f is File => f instanceof File);

        if (allFiles.length === 0) {
            return NextResponse.json({ error: "No file provided" }, { status: 400 });
        }

        for (const file of allFiles) {
            const ext = file.name.split('.').pop()?.toLowerCase();
            if (!ext || (ext !== 'xlsx' && ext !== 'xls')) {
                return NextResponse.json({ error: `Only Excel files (.xlsx, .xls) are supported: ${file.name}` }, { status: 400 });
            }
        }

        console.log("[extract-buy-file] received files:", allFiles.map((f) => ({ name: f.name, size: f.size })));
        const fileBuffers = await Promise.all(allFiles.map(async (f) => ({ file: f, buffer: await f.arrayBuffer() })));

        // Identify the primary buy file: first file whose workbook has a buy sheet
        const engine = new ExcelEngine();
        let buyFileIndex = 0;
        for (let i = 0; i < fileBuffers.length; i++) {
            const analysis = await engine.analyzeWorkbook(fileBuffers[i].buffer);
            if (analysis.hasBuySheet) {
                buyFileIndex = i;
                break;
            }
        }
        const buyFile = fileBuffers[buyFileIndex];
        const productSheetBuffers = fileBuffers
            .filter((_, index) => index !== buyFileIndex)
            .map((fb) => fb.buffer);
        console.log(`[extract-buy-file] selected buy file: ${buyFile.file.name}, product sheet buffers: ${productSheetBuffers.length}`);

        const nextgenEnabled = process.env.NEXTGEN_ENABLED !== 'false';
        // Single NextGen client for this upload: all lookups share one session
        const nextgenClient = nextgenEnabled ? new NextGenCachedClient() : null;

        // New deterministic extraction pipeline with product sheet enrichment
        const customerHint = buyFile.file.name.split('.')[0] || undefined;
        start('extract');
        const extraction = await extractBuyFile(buyFile.buffer, customerHint, nextgenClient || undefined, productSheetBuffers);
        end('extract');
        console.log("[extract-buy-file] extracted:", extraction.items.length, "products:", extraction.productData.length);

        const matchIssues = extraction.items
            .filter((item) => item.matchStatus === 'ambiguous' || item.matchStatus === 'unmatched')
            .map((item) => ({
                field: 'Nexgen Product Match',
                row: item.sourceRow,
                severity: 'CRITICAL' as const,
                status: item.matchStatus,
                style: item.style,
                color: item.colorCode || item.color,
                size: item.size,
                quantity: item.quantity,
                message: item.matchReason || `Nexgen product ${item.matchStatus}`,
            }));
        const matchSummary = extraction.items.reduce(
            (summary, item) => {
                summary[item.matchStatus] += 1;
                summary.totalQuantity += item.quantity || 0;
                if (item.matchStatus === 'matched') summary.matchedQuantity += item.quantity || 0;
                return summary;
            },
            {
                matched: 0,
                ambiguous: 0,
                unmatched: 0,
                not_checked: 0,
                totalQuantity: 0,
                matchedQuantity: 0,
            }
        );

        // Generate output workbooks from single internal model
        start('generate');
        const ordersWb = generateOrdersSheet(extraction.productData);
        const linesWb = generateLinesSheet(extraction.productData);
        const sizesWb = generateSizesSheet(extraction.productData);
        const filesOut = {
            orders: await workbookToBase64(ordersWb),
            lines: await workbookToBase64(linesWb),
            sizes: await workbookToBase64(sizesWb),
        };
        end('generate');

        // Style searches above already provide the Nexgen product and colour
        // fields required for output. Avoid a second 500-row PO scan here.
        const latestPO = null;
        const colorNames: Record<string, string | null> = {};

        const totalTime = end('total');
        console.log(`[extract-buy-file] TOTAL: ${totalTime}ms`);

        return NextResponse.json({
            success: true,
            filename: buyFile.file.name,
            model: 'exceljs-local',
            result: {
                items: extraction.items,
                productData: extraction.productData,
                headerRow: extraction.headerRow,
                headers: extraction.headers,
                mapping: extraction.mapping,
                templateUsed: extraction.templateUsed,
                unmappedColumns: extraction.unmappedColumns,
                files: filesOut,
                latestPO,
                colorNames,
                matchSummary,
                matchIssues,
            },
        });
    } catch (error) {
        const message = error instanceof Error ? error.message : "Unknown error";
        console.error("[extract-buy-file] error:", message);
        console.error("[extract-buy-file] stack:", error instanceof Error ? error.stack : "");
        return NextResponse.json({ error: message }, { status: 500 });
    }
}
