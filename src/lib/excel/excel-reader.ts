import ExcelJS from 'exceljs';

export interface WorksheetSnapshot {
    name: string;
    rows: CellValue[][];
}

export type CellValue = string | number | Date | null;

export interface ExcelReadResult {
    worksheets: WorksheetSnapshot[];
    firstSheet: WorksheetSnapshot;
}

export async function readExcelFile(buffer: ArrayBuffer): Promise<ExcelReadResult> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(Buffer.from(new Uint8Array(buffer)) as any);

    const worksheets: WorksheetSnapshot[] = workbook.worksheets.map((worksheet) => {
        const rows: CellValue[][] = [];
        worksheet.eachRow((row, rowNumber) => {
            const rowValues: CellValue[] = [];
            row.eachCell((cell, colNumber) => {
                // Ensure contiguous array index per column
                rowValues[colNumber - 1] = getCellValue(cell);
            });
            rows[rowNumber - 1] = rowValues;
        });

        return {
            name: worksheet.name,
            rows,
        };
    });

    if (!worksheets.length) {
        throw new Error('Excel file has no worksheets');
    }

    return {
        worksheets,
        firstSheet: worksheets[0],
    };
}

export function getCellValue(cell: ExcelJS.Cell): CellValue {
    if (cell.value === null || cell.value === undefined) return null;

    // ExcelJS stores rich text as object with text property
    if (typeof cell.value === 'object' && 'richText' in cell.value) {
        return (cell.value as ExcelJS.CellRichTextValue).richText.map((t) => t.text).join('');
    }

    if (cell.type === ExcelJS.ValueType.Formula) {
        const formulaValue = cell.value as ExcelJS.CellFormulaValue;
        return formatValue(formulaValue.result);
    }

    if (typeof cell.value === 'object' && (cell.value as any).text) {
        return String((cell.value as any).text || '').trim() || null;
    }

    return formatValue(cell.value);
}

function formatValue(value: unknown): CellValue {
    if (value === null || value === undefined) return null;
    if (typeof value === 'string') return value.trim();
    if (typeof value === 'number') return value;
    if (value instanceof Date) return value;
    return String(value).trim();
}
