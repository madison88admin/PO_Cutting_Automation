import { randomUUID } from "crypto";

export type ProcessingJob = {
    id: string;
    status: "queued" | "processing" | "completed" | "failed";
    createdAt: number;
    updatedAt: number;
    result?: any;
    error?: string;
};

const globalJobs = globalThis as typeof globalThis & {
    __poCuttingJobs?: Map<string, ProcessingJob>;
};

const jobs = globalJobs.__poCuttingJobs || new Map<string, ProcessingJob>();
globalJobs.__poCuttingJobs = jobs;

export function createProcessingJob(): ProcessingJob {
    cleanupJobs();
    const now = Date.now();
    const job: ProcessingJob = {
        id: randomUUID(),
        status: "queued",
        createdAt: now,
        updatedAt: now,
    };
    jobs.set(job.id, job);
    return job;
}

export function updateProcessingJob(id: string, update: Partial<ProcessingJob>) {
    const current = jobs.get(id);
    if (!current) return;
    jobs.set(id, { ...current, ...update, id, updatedAt: Date.now() });
}

export function getProcessingJob(id: string): ProcessingJob | null {
    cleanupJobs();
    return jobs.get(id) || null;
}

function cleanupJobs() {
    const cutoff = Date.now() - 60 * 60 * 1000;
    for (const [id, job] of jobs.entries()) {
        if (job.updatedAt < cutoff) jobs.delete(id);
    }
}
