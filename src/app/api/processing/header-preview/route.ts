import { NextRequest, NextResponse } from "next/server";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 30;

const backendBase = (
    process.env.PROCESSING_API_URL
    || process.env.NEXT_PUBLIC_PROCESSING_API_URL
    || "https://po-cutting-api.5-223-78-194.sslip.io"
).replace(/\/+$/, "");

export async function POST(req: NextRequest) {
    try {
        const formData = await req.formData();
        const response = await fetch(`${backendBase}/api/header-preview`, {
            method: "POST",
            headers: { Origin: "https://m88-po-cutting.netlify.app" },
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
            error: error instanceof Error ? error.message : "Could not preview headers",
        }, { status: 502 });
    }
}
