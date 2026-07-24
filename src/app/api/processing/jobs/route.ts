import { NextRequest, NextResponse } from "next/server";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 30;

const backendBase = (
    process.env.PROCESSING_API_URL
    || process.env.NEXT_PUBLIC_PROCESSING_API_URL
    || "https://po-cutting-api.5-223-78-194.sslip.io"
).replace(/\/+$/, "");

function backendHeaders() {
    return {
        "x-processing-key": process.env.ADMIN_PANEL_PASSWORD || "",
    };
}

export async function POST(req: NextRequest) {
    try {
        const formData = await req.formData();
        const response = await fetch(`${backendBase}/api/processing-jobs`, {
            method: "POST",
            headers: backendHeaders(),
            body: formData,
            cache: "no-store",
        });
        const body = await response.text();
        return new NextResponse(body, {
            status: response.status,
            headers: {
                "Content-Type": response.headers.get("content-type") || "application/json",
                "Cache-Control": "no-store",
            },
        });
    } catch (error) {
        return NextResponse.json({
            error: error instanceof Error ? error.message : "Could not start processing job",
        }, { status: 502 });
    }
}

export async function GET(req: NextRequest) {
    try {
        const id = req.nextUrl.searchParams.get("id") || "";
        if (!/^[0-9a-f-]{36}$/i.test(id)) {
            return NextResponse.json({ error: "Invalid processing job ID" }, { status: 400 });
        }
        const response = await fetch(`${backendBase}/api/processing-jobs/${id}`, {
            headers: backendHeaders(),
            cache: "no-store",
        });
        const body = await response.text();
        return new NextResponse(body, {
            status: response.status,
            headers: {
                "Content-Type": response.headers.get("content-type") || "application/json",
                "Cache-Control": "no-store",
            },
        });
    } catch (error) {
        return NextResponse.json({
            error: error instanceof Error ? error.message : "Could not check processing job",
        }, { status: 502 });
    }
}
