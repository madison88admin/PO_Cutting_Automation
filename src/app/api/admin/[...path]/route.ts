import { NextRequest, NextResponse } from "next/server";
import { withAuth } from "@/lib/auth";
import { getRunHistory } from "@/lib/db/runHistory";
import { getAuditLogs } from "@/lib/db/auditLog";
import { getFactoryMappings, upsertFactory, deleteFactory } from "@/lib/db/factoryMapping";
import { getColumnMappings, upsertColumn } from "@/lib/db/columnMapping";
import { getMloMappings, upsertMlo } from "@/lib/db/mloMapping";

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

    if (path.endsWith('/columns')) {
        return withAuth(req, "EDIT_MAPPING_TABLES", async (req, session) => {
            const url2 = new URL(req.url);
            const customer = url2.searchParams.get('customer') || undefined;
            const data = await getColumnMappings(customer);
            return NextResponse.json(data);
        });
    }

    if (path.endsWith('/mlo')) {
        return withAuth(req, "EDIT_MAPPING_TABLES", async (req, session) => {
            const data = await getMloMappings();
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

            await upsertFactory({ brand, category, product_supplier: productSupplier }, session.userId);
            return NextResponse.json({ ok: true });
        });
    }

    if (path.endsWith('/columns')) {
        return withAuth(req, "EDIT_MAPPING_TABLES", async (_req, session) => {
            const body = await req.json();
            const customer = String(body?.customer || "").trim();
            const buy_file_column = String(body?.buy_file_column || "").trim();
            const internal_field = String(body?.internal_field || "").trim();

            if (!customer || !buy_file_column || !internal_field) {
                return NextResponse.json(
                    { error: "customer, buy_file_column, and internal_field are required" },
                    { status: 400 }
                );
            }

            await upsertColumn({ customer, buy_file_column, internal_field, notes: body?.notes || "" }, session.userId);
            return NextResponse.json({ ok: true });
        });
    }

    if (path.endsWith('/mlo')) {
        return withAuth(req, "EDIT_MAPPING_TABLES", async (_req, session) => {
            const body = await req.json();
            const brand = String(body?.brand || "").trim();
            if (!brand) {
                return NextResponse.json({ error: "brand is required" }, { status: 400 });
            }

            await upsertMlo({
                brand,
                keyuser1: String(body?.keyuser1 || ""),
                keyuser2: String(body?.keyuser2 || ""),
                keyuser4: String(body?.keyuser4 || ""),
                keyuser5: String(body?.keyuser5 || ""),
                orders_template: body?.orders_template || null,
                lines_template: body?.lines_template || null,
                valid_statuses: Array.isArray(body?.valid_statuses) ? body.valid_statuses : [],
            }, session.userId);
            return NextResponse.json({ ok: true });
        });
    }

    return NextResponse.json({ error: "Not Found" }, { status: 404 });
}

export async function DELETE(req: NextRequest) {
    const url = new URL(req.url);
    const path = url.pathname;

    if (path.endsWith('/factory')) {
        return withAuth(req, "EDIT_MAPPING_TABLES", async (_req, session) => {
            const body = await req.json();
            const id = String(body?.id || "").trim();
            if (!id) return NextResponse.json({ error: "id is required" }, { status: 400 });
            await deleteFactory(id, session.userId);
            return NextResponse.json({ ok: true });
        });
    }

    return NextResponse.json({ error: "Not Found" }, { status: 404 });
}
