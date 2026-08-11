import { NextRequest, NextResponse } from "next/server";
import { getProcessingJob } from "@/lib/processing-jobs";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

function authorized(req: NextRequest): boolean {
    const expected = process.env.ADMIN_PANEL_PASSWORD || "";
    return Boolean(expected && req.headers.get("x-processing-key") === expected);
}

export async function GET(
    req: NextRequest,
    context: { params: Promise<{ id: string }> },
) {
    if (!authorized(req)) {
        return NextResponse.json({ error: "Unauthorized processing request" }, { status: 401 });
    }
    const { id } = await context.params;
    const job = getProcessingJob(id);
    if (!job) {
        return NextResponse.json({ error: "Processing job not found or expired" }, { status: 404 });
    }
    return NextResponse.json(job, {
        headers: { "Cache-Control": "no-store" },
    });
}
