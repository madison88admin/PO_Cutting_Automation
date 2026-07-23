import { NextRequest, NextResponse } from "next/server";
import { NextGenClient } from "@/lib/nextgen";

export async function POST(req: NextRequest) {
    try {
        const body = await req.json();
        const skus = Array.isArray(body.skus) ? body.skus : (body.sku ? [body.sku] : []);
        if (!skus.length) {
            return NextResponse.json({ error: "sku or skus is required" }, { status: 400 });
        }

        const client = new NextGenClient();
        const results = await client.lookupColorNames(skus.map(String));
        return NextResponse.json({ results });
    } catch (err: any) {
        console.error("[nextgen-color-lookup] error:", err);
        return NextResponse.json({ error: err.message || "Failed to lookup color" }, { status: 500 });
    }
}
