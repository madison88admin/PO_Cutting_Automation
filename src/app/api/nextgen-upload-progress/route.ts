/**
 * Server-Sent Events (SSE) endpoint for real-time NextGen upload progress.
 * The frontend can connect to /api/nextgen-upload-progress?id=xxx to receive
 * progress updates as POs/lines/sizes are created.
 */
import { NextRequest } from 'next/server';

export const dynamic = 'force-dynamic';
export const runtime = 'nodejs';

// In-memory progress store (per upload session)
const progressStore = new Map<string, {
    total: number;
    current: number;
    status: 'pending' | 'uploading' | 'complete' | 'error';
    message: string;
    poNumber: string;
    lineItem: string;
    errors: string[];
}>();

export function updateProgress(id: string, update: Partial<{
    total: number;
    current: number;
    status: 'pending' | 'uploading' | 'complete' | 'error';
    message: string;
    poNumber: string;
    lineItem: string;
    errors: string[];
}>) {
    const existing = progressStore.get(id);
    if (existing) {
        progressStore.set(id, { ...existing, ...update });
    } else {
        progressStore.set(id, {
            total: 0,
            current: 0,
            status: 'pending',
            message: '',
            poNumber: '',
            lineItem: '',
            errors: [],
            ...update,
        });
    }
}

export function getProgress(id: string) {
    return progressStore.get(id) || null;
}

export function clearProgress(id: string) {
    progressStore.delete(id);
}

export async function GET(req: NextRequest) {
    const id = req.nextUrl.searchParams.get('id');
    if (!id) {
        return new Response('Missing id parameter', { status: 400 });
    }

    const encoder = new TextEncoder();
    const stream = new ReadableStream({
        start(controller) {
            let closed = false;

            const send = (data: any) => {
                if (closed) return;
                try {
                    controller.enqueue(encoder.encode(`data: ${JSON.stringify(data)}\n\n`));
                } catch {
                    closed = true;
                }
            };

            // Poll progress every 500ms
            const interval = setInterval(() => {
                const progress = progressStore.get(id);
                if (progress) {
                    send(progress);
                    if (progress.status === 'complete' || progress.status === 'error') {
                        clearInterval(interval);
                        send({ ...progress, done: true });
                        controller.close();
                        closed = true;
                        clearProgress(id);
                    }
                }
            }, 500);

            // Clean up on abort
            req.signal.addEventListener('abort', () => {
                clearInterval(interval);
                closed = true;
                try { controller.close(); } catch {}
            });
        },
    });

    return new Response(stream, {
        headers: {
            'Content-Type': 'text/event-stream',
            'Cache-Control': 'no-cache, no-transform',
            'Connection': 'keep-alive',
        },
    });
}
