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
        const productSheetBuffers = fileBuffers.map((fb) => fb.buffer);
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

        // Latest PO and color lookups — run in parallel (both use same cached records)
        // These are non-critical: if NextGen is down, extraction still succeeds
        start('nextgenExtras');
        let latestPO = null;
        let colorNames: Record<string, string | null> = {};
        if (nextgenClient) {
            try {
                const skus = extraction.productData.map((p) => p.productExternalRef || p.style).filter(Boolean) as string[];
                const [poResult, colorResult] = await Promise.allSettled([
                    nextgenClient.getLatestPO(),
                    nextgenClient.lookupColorNames(skus),
                ]);
                if (poResult.status === 'fulfilled') latestPO = poResult.value;
                else console.warn('[extract-buy-file] getLatestPO failed:', poResult.reason);
                if (colorResult.status === 'fulfilled') colorNames = colorResult.value;
                else console.warn('[extract-buy-file] lookupColorNames failed:', colorResult.reason);
            } catch (err) {
                console.warn('[extract-buy-file] nextgenExtras failed (non-critical):', err);
            }
        }
        end('nextgenExtras');

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
            },
        });
    } catch (error) {
        const message = error instanceof Error ? error.message : "Unknown error";
        console.error("[extract-buy-file] error:", message);
        console.error("[extract-buy-file] stack:", error instanceof Error ? error.stack : "");
        return NextResponse.json({ error: message }, { status: 500 });
    }
}
