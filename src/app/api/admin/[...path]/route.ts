import { NextRequest, NextResponse } from "next/server";
import { withAuth } from "@/lib/auth";
import { getRunHistory } from "@/lib/db/runHistory";
import { getAuditLogs } from "@/lib/db/auditLog";
import { getFactoryMappings, upsertFactory } from "@/lib/db/factoryMapping";

export async function GET(req: NextRequest) {
    // Check path to determine what to return
    const url = new URL(req.url);
    const path = url.pathname;

    if (path.endsWith('/runs')) {
        return withAuth(req, "VIEW_AUDIT_LOGS", async (req, session) => {
            const data = await getRunHistory(session.userId, session.role);
            return NextResponse.json(data);
        });
    }

    if (path.endsWith('/audit')) {
        return withAuth(req, "VIEW_AUDIT_LOGS", async (req, session) => {
            const data = await getAuditLogs();
            return NextResponse.json(data);
        });
    }

    if (path.endsWith('/factory')) {
        return withAuth(req, "EDIT_MAPPING_TABLES", async (req, session) => {
            const data = await getFactoryMappings();
            return NextResponse.json(data);
        });
    }

    return NextResponse.json({ error: "Not Found" }, { status: 404 });
}

export async function POST(req: NextRequest) {
    const url = new URL(req.url);
    const path = url.pathname;

    if (path.endsWith('/factory')) {
        return withAuth(req, "EDIT_MAPPING_TABLES", async (_req, session) => {
            const body = await req.json();
            const brand = String(body?.brand || "").trim();
            const category = String(body?.category || "").trim();
            const productSupplier = String(body?.product_supplier || "").trim();

            if (!brand || !category || !productSupplier) {
                return NextResponse.json(
                    { error: "brand, category, and product_supplier are required" },
                    { status: 400 }
                );
            }

            await upsertFactory(
                {
                    brand,
                    category,
                    product_supplier: productSupplier,
                },
                session.userId
            );

            return NextResponse.json({ ok: true });
        });
    }

    return NextResponse.json({ error: "Not Found" }, { status: 404 });
}
