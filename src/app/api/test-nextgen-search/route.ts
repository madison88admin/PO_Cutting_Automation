import { NextRequest, NextResponse } from "next/server";
import { NextGenClient } from "@/lib/nextgen";

const SEARCH_BASE_URL = process.env.NEXTGEN_SEARCH_BASE_URL || process.env.NEXTGEN_BASE_URL || 'https://nextgen.madison88.com';

export async function GET(req: NextRequest) {
    try {
        const url = new URL(req.url);
        const style = url.searchParams.get('style') || 'NF0A8KYV';
        const path = url.searchParams.get('path') || process.env.NEXTGEN_SEARCH_PATH || '/api/v1/nextgen/search';

        const client = new NextGenClient();
        await client.login();

        const searchUrl = `${SEARCH_BASE_URL}${path.replace(/\{style\}/g, encodeURIComponent(style))}`;
        console.log('[test-nextgen-search] calling:', searchUrl);

        const response = await client.fetchWithCookie(searchUrl, { method: 'GET' }, true);
        const text = await response.text();

        console.log('[test-nextgen-search] status:', response.status);
        console.log('[test-nextgen-search] body:', text.slice(0, 1000));

        let data: any = null;
        try {
            data = text ? JSON.parse(text) : null;
        } catch {
            data = null;
        }

        return NextResponse.json({
            success: response.ok,
            status: response.status,
            url: searchUrl,
            headers: Object.fromEntries(response.headers.entries()),
            body: text.slice(0, 2000),
            data,
        });
    } catch (error) {
        const message = error instanceof Error ? error.message : "Unknown error";
        console.error('[test-nextgen-search] error:', message);
        return NextResponse.json({ error: message }, { status: 500 });
    }
}
