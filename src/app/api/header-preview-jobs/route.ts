import { NextRequest, NextResponse } from "next/server";
import { createProcessingJob, updateProcessingJob } from "@/lib/processing-jobs";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

function authorized(req: NextRequest): boolean {
    const expected = process.env.ADMIN_PANEL_PASSWORD || "";
    return Boolean(expected && req.headers.get("x-processing-key") === expected);
}

export async function POST(req: NextRequest) {
    if (!authorized(req)) {
        return NextResponse.json({ error: "Unauthorized preview request" }, { status: 401 });
    }
    const formData = await req.formData();
    if (!(formData.get("file") instanceof File)) {
        return NextResponse.json({ error: "No buy file uploaded" }, { status: 400 });
    }

    const job = createProcessingJob();
    const internalBase = (process.env.PROCESSING_INTERNAL_BASE_URL || "http://127.0.0.1:3003").replace(/\/+$/, "");
    const allowedOrigin = (process.env.ALLOWED_UPLOAD_ORIGINS || "https://m88-po-cutting.netlify.app")
        .split(",")[0]
        .trim()
        .replace(/\/+$/, "");
    updateProcessingJob(job.id, { status: "processing" });

    void fetch(`${internalBase}/api/header-preview`, {
        method: "POST",
        headers: { Origin: allowedOrigin },
        body: formData,
    }).then(async (response) => {
        const result = await response.json();
        if (!response.ok || result?.error) {
            updateProcessingJob(job.id, {
                status: "failed",
                error: result?.error || `Preview failed with HTTP ${response.status}`,
                result,
            });
            return;
        }
        updateProcessingJob(job.id, { status: "completed", result });
    }).catch((error) => {
        updateProcessingJob(job.id, {
            status: "failed",
            error: error instanceof Error ? error.message : "Preview request failed",
        });
    });

    return NextResponse.json({ jobId: job.id, status: "processing" }, { status: 202 });
}
