import { NextRequest, NextResponse } from "next/server";
import { ExcelEngine } from "@/lib/excel-engine";
import { logEvent } from "@/lib/audit";
import { createRun } from "@/lib/db/runHistory";

export async function POST(req: NextRequest) {
    try {
        const formData = await req.formData();
        const file = formData.get("file") as File;
        const publicUserId = "public-workflow-user";

        if (!file) {
            return NextResponse.json({ error: "No file uploaded" }, { status: 400 });
        }

        const buffer = Buffer.from(await file.arrayBuffer());

        // Create Run History record for public cutting workflow
        const runId = await createRun({
            user_id: publicUserId,
            filename: file.name,
            status: 'Processing'
        });

        await logEvent({
            eventName: "BUY_FILE_UPLOADED",
            userId: publicUserId,
            runId,
            metadata: { filename: file.name, size: file.size }
        });

        await logEvent({
            eventName: "WORKFLOW_STARTED",
            userId: publicUserId,
            runId,
            metadata: { filename: file.name }
        });

        const engine = new ExcelEngine(runId, publicUserId);
        const { data, errors } = await engine.processBuyFile(buffer);

        // Generate the output buffers
        const outputs = await engine.generateOutputs(data);

        const hasCritical = errors.some(e => e.severity === "CRITICAL");

        await logEvent({
            eventName: "DATA_EXTRACTION_COMPLETE",
            userId: publicUserId,
            runId,
            metadata: { rows_extracted: data.length }
        });

        console.log("Generating Base64 outputs...");
        const files = {
            orders: Buffer.from(outputs.orders as any).toString('base64'),
            lines: Buffer.from(outputs.lines as any).toString('base64'),
            sizes: Buffer.from(outputs.sizes as any).toString('base64')
        };
        console.log(`Base64 generation complete. Orders: ${files.orders.length} chars, Lines: ${files.lines.length} chars, Sizes: ${files.sizes.length} chars`);

        return NextResponse.json({
            success: true,
            runId,
            dataCount: data.length,
            errors,
            canProceed: !hasCritical,
            files
        });

    } catch (error: any) {
        console.error("Upload error:", error);
        return NextResponse.json({ error: "Internal Server Error" }, { status: 500 });
    }
}
