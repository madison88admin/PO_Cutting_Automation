import { supabaseAdmin, isMock } from '../supabase';

export interface AuditLogEntry {
    event: string;
    user_id: string;
    run_id?: string;
    metadata?: any;
    ip_address?: string;
    created_at?: string;
}

export async function logEvent(entry: AuditLogEntry): Promise<void> {
    if (isMock) {
        console.log(`[MOCK AUDIT] ${entry.event}:`, entry.metadata);
        return;
    }

    const { error } = await supabaseAdmin
        .from('audit_logs')
        .insert({
            event: entry.event,
            user_id: entry.user_id,
            run_id: entry.run_id,
            metadata: entry.metadata,
            ip_address: entry.ip_address
        });

    if (error) {
        console.error('Failed to log event to Supabase:', error);
    }
}

export async function getAuditLogs(filters?: { event?: string; user_id?: string; startDate?: string; endDate?: string }) {
    if (isMock) return [];

    let query = supabaseAdmin
        .from('audit_logs')
        .select('*, users(name, email)')
        .order('created_at', { ascending: false });

    if (filters?.event) query = query.eq('event', filters.event);
    if (filters?.user_id) query = query.eq('user_id', filters.user_id);
    if (filters?.startDate) query = query.gte('created_at', filters.startDate);
    if (filters?.endDate) query = query.lte('created_at', filters.endDate);

    const { data, error } = await query;
    if (error) throw error;
    return data || [];
}
