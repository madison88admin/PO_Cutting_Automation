import { NextResponse } from "next/server";

export async function GET() {
    return NextResponse.json({
        status: "ok",
        service: "po-line",
        timestamp: new Date().toISOString(),
    });
}
