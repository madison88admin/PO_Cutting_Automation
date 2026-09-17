import { NextRequest, NextResponse } from "next/server";
import { NextGenClient } from "@/lib/nextgen";

const BASE_URL = process.env.NEXTGEN_BASE_URL || 'https://nextgen.madison88.com';

export async function GET(req: NextRequest) {
    try {
        const url = new URL(req.url);
        const productId = url.searchParams.get('id') || '45809';
        const path = url.searchParams.get('path') || `/Product/Edit/${productId}`;

        const client = new NextGenClient();
        await client.login();

        const targetUrl = `${BASE_URL}${path}`;
        console.log('[test-nextgen-product] calling:', targetUrl);

        const response = await client.fetchWithCookie(targetUrl, { method: 'GET' }, true);
        const text = await response.text();

        console.log('[test-nextgen-product] status:', response.status);
        console.log('[test-nextgen-product] content-type:', response.headers.get('content-type'));
        console.log('[test-nextgen-product] body preview:', text.slice(0, 2000));

        let data: any = null;
        try {
            data = text ? JSON.parse(text) : null;
        } catch {
            data = null;
        }

        return NextResponse.json({
            success: response.ok,
            status: response.status,
            url: targetUrl,
            contentType: response.headers.get('content-type'),
            body: text.slice(0, 3000),
            data,
        });
    } catch (error) {
        const message = error instanceof Error ? error.message : "Unknown error";
        console.error('[test-nextgen-product] error:', message);
        return NextResponse.json({ error: message }, { status: 500 });
    }
}
