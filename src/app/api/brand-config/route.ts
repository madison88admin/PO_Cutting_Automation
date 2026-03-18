import { NextRequest, NextResponse } from "next/server";
import { getMloMappings } from "@/lib/db/mloMapping";

type BrandConfigResponse = {
    brand: string;
    orders_template: string | null;
    lines_template: string | null;
    valid_statuses: string[];
    keyusers: {
        KeyUser1: string;
        KeyUser2: string;
        KeyUser3: string;
        KeyUser4: string;
        KeyUser5: string;
        KeyUser6: string;
        KeyUser7: string;
        KeyUser8: string;
    };
};

export async function GET(req: NextRequest) {
    const url = new URL(req.url);
    const brandRaw = (url.searchParams.get("brand") || "").trim();
    const brandKey = brandRaw.toLowerCase();

    const mappings = await getMloMappings();
    const row = mappings.find(m => (m.brand || "").trim().toLowerCase() === brandKey);

    const response: BrandConfigResponse = {
        brand: row?.brand || brandRaw,
        orders_template: row?.orders_template?.trim() || null,
        lines_template: row?.lines_template?.trim() || null,
        valid_statuses: Array.isArray(row?.valid_statuses)
            ? row!.valid_statuses!.filter(Boolean)
            : [],
        keyusers: {
            KeyUser1: row?.keyuser1 || "",
            KeyUser2: row?.keyuser2 || "",
            KeyUser3: "",
            KeyUser4: row?.keyuser4 || "",
            KeyUser5: row?.keyuser5 || "",
            KeyUser6: "",
            KeyUser7: "",
            KeyUser8: "",
        },
    };

    return NextResponse.json(response);
}
