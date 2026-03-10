import { logEvent as dbLogEvent, AuditLogEntry } from './db/auditLog';

export interface AuditEntry {
    eventName: string;
    userId: string;
    timestamp?: string;
    runId?: string;
    details?: string;
    metadata?: any;
    ipAddress?: string;
}

export async function logEvent(entry: AuditEntry) {
    try {
        await dbLogEvent({
            event: entry.eventName,
            user_id: entry.userId,
            run_id: entry.runId,
            metadata: {
                details: entry.details,
                ...entry.metadata
            },
            ip_address: entry.ipAddress
        });

        if (entry.eventName === "UNAUTHORIZED_ACCESS") {
            console.warn(`[AUDIT] Security Alarm: Unauthorized access attempt by ${entry.userId}`);
        }
    } catch (error) {
        console.error("Failed to write to audit log:", error);
    }
}
