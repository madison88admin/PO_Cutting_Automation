import { NextRequest, NextResponse } from "next/server";
import { NextGenClient, createWriteClient, isWriteEnabled } from "@/lib/nextgen";

/**
 * GET /api/nextgen-discover
 * Query params:
 *   - pages: comma-separated list of NextGen page paths to fetch (default: standard PO pages)
 *   - env: "test" (default) or "prod" — which NextGen environment to discover from
 *
 * Fetches NextGen UI HTML pages (read-only GET) and parses them for
 * insert/update/read endpoint URLs. This does NOT create any data in NextGen.
 *
 * By default, discovers from the TEST environment (port 8443) to match
 * where writes will go. Use ?env=prod to discover from production.
 */
export async function GET(req: NextRequest) {
    try {
        const url = new URL(req.url);
        const pagesParam = url.searchParams.get('pages') || '';
        const env = url.searchParams.get('env') || 'test';
        const pages = pagesParam
            ? pagesParam.split(',').map(p => p.trim()).filter(Boolean)
            : [
                '/PurchaseOrder/Create',
                '/PurchaseOrder',
                '/PurchaseOrder/Index',
                '/PurchaseOrder/Edit',
            ];

        // Use test env by default (matches write target); prod for read-only discovery
        const client = env === 'prod'
            ? new NextGenClient()
            : (isWriteEnabled() ? createWriteClient() : new NextGenClient());
        const result = await client.discoverEndpoints(pages);

        // Filter endpoints to those likely related to insert/write operations
        const insertEndpoints = result.allEndpoints.filter(
            ep => ep.url.toLowerCase().includes('insert') ||
                  ep.url.toLowerCase().includes('create') ||
                  ep.url.toLowerCase().includes('update') ||
                  ep.url.toLowerCase().includes('save') ||
                  ep.url.toLowerCase().includes('delete')
        );

        const readEndpoints = result.allEndpoints.filter(
            ep => ep.url.toLowerCase().includes('read') ||
                  ep.url.toLowerCase().includes('list') ||
                  ep.url.toLowerCase().includes('index')
        );

        const lineEndpoints = result.allEndpoints.filter(
            ep => ep.url.toLowerCase().includes('line') ||
                  ep.url.toLowerCase().includes('item')
        );

        const sizeEndpoints = result.allEndpoints.filter(
            ep => ep.url.toLowerCase().includes('size') ||
                  ep.url.toLowerCase().includes('variant')
        );

        return NextResponse.json({
            success: true,
            pagesFetched: result.pages.map(p => ({ path: p.path, status: p.status, endpointCount: p.endpoints.length })),
            insertEndpoints,
            readEndpoints,
            lineEndpoints,
            sizeEndpoints,
            allEndpoints: result.allEndpoints,
            pageDetails: result.pages,
        });
    } catch (error) {
        const message = error instanceof Error ? error.message : "Unknown error";
        console.error("[nextgen-discover] error:", message);
        return NextResponse.json({ error: message }, { status: 500 });
    }
}
