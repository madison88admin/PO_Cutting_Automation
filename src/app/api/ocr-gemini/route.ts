import { NextRequest, NextResponse } from "next/server";
import { GeminiOCRClient, GeminiOCRResult } from "@/lib/gemini";
import { NextGenClient } from "@/lib/nextgen";

export async function POST(req: NextRequest) {
    try {
        const formData = await req.formData();
        const file = formData.get("file") as File | null;
        const fillFromNextgen = formData.get("fillFromNextgen") === "true";

        if (!file) {
            return NextResponse.json({ error: "No file provided" }, { status: 400 });
        }

        const bytes = await file.arrayBuffer();
        const base64 = Buffer.from(bytes).toString("base64");
        const mimeType = file.type || "application/pdf";

        const gemini = new GeminiOCRClient();
        const ocrResults = await gemini.extractFromFile(base64, mimeType);

        let mergedResults: GeminiOCRResult[] = ocrResults;
        let nextgenUsed = false;
        let nextgenError: string | null = null;

        if (fillFromNextgen && ocrResults.length > 0) {
            const poNumber = ocrResults[0].poNumber || ocrResults.find(r => r.poNumber)?.poNumber || '';
            if (poNumber) {
                try {
                    const nextgen = new NextGenClient();
                    const validation = await nextgen.validatePO(poNumber, ocrResults.map(r => ({
                        style: r.style,
                        color: r.color,
                        size: r.size,
                        quantity: r.quantity,
                    })));

                    if (validation.exists && validation.lines.length > 0) {
                        mergedResults = ocrResults.map((ocr) => {
                            const match = validation.lines.find(
                                (ng) =>
                                    ng.style.toLowerCase().trim() === ocr.style.toLowerCase().trim() &&
                                    ng.color.toLowerCase().trim() === ocr.color.toLowerCase().trim() &&
                                    ng.size.toLowerCase().trim() === ocr.size.toLowerCase().trim()
                            );
                            if (!match) return ocr;
                            const fillString = (ocrValue: string, ngValue: unknown) => {
                                const value = typeof ngValue === 'string' ? ngValue : '';
                                return ocrValue || value || '';
                            };
                            return {
                                poNumber: ocr.poNumber,
                                style: ocr.style,
                                color: ocr.color,
                                size: ocr.size,
                                quantity: ocr.quantity,
                                factory: fillString(ocr.factory, match.factory),
                                customer: fillString(ocr.customer, match.customer),
                                season: fillString(ocr.season, match.season),
                                exFtyDate: fillString(ocr.exFtyDate, match.exFtyDate),
                                transportMethod: fillString(ocr.transportMethod, match.transportMethod),
                                plant: fillString(ocr.plant, match.plant),
                            };
                        });
                        nextgenUsed = true;
                    }
                } catch (err) {
                    nextgenError = err instanceof Error ? err.message : 'NextGen lookup failed';
                    console.error('[ocr-gemini] nextgen fill error:', nextgenError);
                }
            }
        }

        return NextResponse.json({
            ocrResults,
            mergedResults,
            filename: file.name,
            nextgenUsed,
            nextgenError,
        });
    } catch (error) {
        const message = error instanceof Error ? error.message : "Unknown error";
        console.error("[ocr-gemini] error:", message);
        return NextResponse.json({ error: message }, { status: 500 });
    }
}
