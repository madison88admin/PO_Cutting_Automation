import { NextRequest, NextResponse } from "next/server";
import { NextGenClient } from "@/lib/nextgen";

/**
 * GET /api/nextgen-po-lines
 *
 * Fetch PO lines from NextGen's PurchaseOrder/Read endpoint (line-level rows:
 * one row per style+colour+size, with PO header fields attached).
 *
 * Query params (GET) or JSON body (POST):
 *   poNumber   — exact PO number filter (e.g. M88PO2488)
 *   style      — style/code substring (matched across product fields)
 *   page       — 1-based page number (default 1)
 *   pageSize   — rows per page, 1..1000 (default 200)
 *
 * Response:
 * {
 *   poNumber, style, page, pageSize, count, totalRows, hasMore, lines: [...]
 * }
 *
 * Power BI / Power Query: pass page=1..N while hasMore=true, keep pageSize fixed.
 */
export async function GET(req: NextRequest) {
    return handle(req);
}

export async function POST(req: NextRequest) {
    return handle(req);
}

async function handle(req: NextRequest) {
    try {
        let poNumber = "";
        let style = "";
        let page = 1;
        let pageSize = 200;

        if (req.method === "POST") {
            const body = await req.json().catch(() => ({}));
            poNumber = String(body.poNumber || "");
            style = String(body.style || "");
            page = Number(body.page) || 1;
            pageSize = Number(body.pageSize) || 200;
        } else {
            poNumber = req.nextUrl.searchParams.get("poNumber") || "";
            style = req.nextUrl.searchParams.get("style") || "";
            page = Number(req.nextUrl.searchParams.get("page")) || 1;
            pageSize = Number(req.nextUrl.searchParams.get("pageSize")) || 200;
        }

        const client = new NextGenClient();
        const result = await client.fetchPOLines({ poNumber, style, page, pageSize });

        return NextResponse.json({
            poNumber: poNumber || null,
            style: style || null,
            page: result.page,
            pageSize: result.pageSize,
            count: result.lines.length,
            totalRows: result.totalRows,
            hasMore: result.hasMore,
            lines: result.lines,
        });
    } catch (error) {
        const message = error instanceof Error ? error.message : "Unknown error";
        console.error("[nextgen-po-lines] error:", message);
        return NextResponse.json({ error: message }, { status: 500 });
    }
}
