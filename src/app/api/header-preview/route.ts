import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";
import { detectHeaderRow } from "@/lib/ai/header-detector";
import { mapHeaders } from "@/lib/ai/header-mapper";
import { findMatchingTemplateSupabase } from "@/lib/templates/supabase-store";

export async function POST(req: NextRequest) {
    try {
        if (process.env.NODE_ENV === "production") {
            const origin = (req.headers.get("origin") || "").replace(/\/+$/, "");
            const allowed = (process.env.ALLOWED_UPLOAD_ORIGINS || "https://m88-po-cutting.netlify.app")
                .split(",")
                .map((value) => value.trim().replace(/\/+$/, ""))
                .filter(Boolean);
            if (origin && !allowed.includes(origin)) {
                return NextResponse.json({ error: "Invalid request origin" }, { status: 403 });
            }
        }
        const formData = await req.formData();
        const files = formData.getAll("file") as File[];
        if (!files.length) {
            return NextResponse.json({ error: "No Excel file provided" }, { status: 400 });
        }

        const previews = [];
        for (const file of files) {
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(await file.arrayBuffer());
            let best: any = null;

            for (const worksheet of workbook.worksheets) {
                const firstRows: unknown[][] = [];
                for (let rowNumber = 1; rowNumber <= Math.min(10, worksheet.rowCount); rowNumber += 1) {
                    const values = worksheet.getRow(rowNumber).values;
                    firstRows.push(Array.isArray(values) ? values.slice(1) : Object.values(values || {}));
                }
                const detected = await detectHeaderRow(firstRows);
                const detectedValues = worksheet.getRow(detected.headerRow).values;
                const headers = (Array.isArray(detectedValues) ? detectedValues.slice(1) : Object.values(detectedValues || {}))
                    .map((value) => String(value || "").trim())
                    .filter(Boolean);
                if (!headers.length) continue;

                const learned = await findMatchingTemplateSupabase(headers);
                const mapped = learned
                    ? { mapping: learned.mapping, confidence: 100, unmappedColumns: headers.filter((header) => !Object.values(learned.mapping).includes(header)) }
                    : await mapHeaders(headers);
                const score = Object.keys(mapped.mapping).length;
                if (!best || score > best.score) {
                    best = {
                        filename: file.name,
                        worksheet: worksheet.name,
                        headerRow: detected.headerRow,
                        headers,
                        mapping: mapped.mapping,
                        confidence: mapped.confidence,
                        unmappedColumns: mapped.unmappedColumns,
                        source: learned ? "learned template" : "Qwen + header aliases",
                        score,
                    };
                }
            }

            if (!best) {
                previews.push({ filename: file.name, error: "Could not detect a header row" });
            } else {
                delete best.score;
                previews.push(best);
            }
        }

        return NextResponse.json({ success: true, previews });
    } catch (error) {
        return NextResponse.json({
            error: error instanceof Error ? error.message : "Header preview failed",
        }, { status: 500 });
    }
}
