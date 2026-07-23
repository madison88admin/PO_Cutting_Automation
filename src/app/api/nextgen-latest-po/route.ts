import { NextResponse } from "next/server";
import { NextGenClient } from "@/lib/nextgen";

export async function GET() {
    try {
        const client = new NextGenClient();
        const latest = await client.getLatestPO();

        if (!latest) {
            return NextResponse.json({ poNumber: null, message: "No PO found in NextGen" });
        }

        return NextResponse.json({ poNumber: latest.poNumber });
    } catch (error) {
        const message = error instanceof Error ? error.message : "Unknown error";
        const stack = error instanceof Error ? error.stack : null;
        console.error("[nextgen-latest-po] error:", message, stack);
        return NextResponse.json({ error: message, stack }, { status: 500 });
    }
}
