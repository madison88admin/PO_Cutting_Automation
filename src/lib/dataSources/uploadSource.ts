import { ExcelEngine } from "../excel-engine";

export async function processFromFile(buffer: Buffer, runId: string, userId: string) {
    const engine = new ExcelEngine(runId, userId);
    return await engine.processBuyFile(buffer);
}
