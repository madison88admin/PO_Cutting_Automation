import { processFromFile } from "./uploadSource";
import { processFromApi } from "./apiSource";

export async function processData(source: "upload" | "api", payload: any, runId: string, userId: string) {
    if (source === "upload") {
        return await processFromFile(payload as Buffer, runId, userId);
    } else {
        return await processFromApi(runId, userId);
    }
}
