import { NextRequest, NextResponse } from "next/server";
import { NextGenClient } from "@/lib/nextgen";

export async function POST(req: NextRequest) {
    try {
        const body = await req.json();
        const { poNumber, lines } = body as {
            poNumber: string;
            lines: { style: string; color: string; size: string; quantity: number }[];
        };

        if (!poNumber || !Array.isArray(lines)) {
            return NextResponse.json(
                { error: "poNumber and lines are required" },
                { status: 400 }
            );
        }

        const client = new NextGenClient();
        const result = await client.validatePO(poNumber, lines);

        return NextResponse.json(result);
    } catch (error) {
        const message = error instanceof Error ? error.message : "Unknown error";
        console.error("[validate-nextgen] error:", message);
        return NextResponse.json(
            { error: message },
            { status: 500 }
        );
    }
}
