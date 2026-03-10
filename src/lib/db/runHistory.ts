import { supabaseAdmin, isMock } from '../supabase';

export interface RunHistory {
    id: string;
    user_id: string;
    filename: string;
    status: 'Processing' | 'Validation Failed' | 'Pending Review' | 'Approved' | 'Rejected';
    error_count: number;
    warning_count: number;
    orders_rows: number;
    lines_rows: number;
    order_sizes_rows: number;
    reviewed_by?: string;
    review_notes?: string;
    created_at: string;
    completed_at?: string;
}

const MOCK_RUNS: RunHistory[] = [
    {
        id: 'mock-run-1',
        user_id: '00000000-0000-0000-0000-000000000001',
        filename: 'PO_Extract_2026.xlsx',
        status: 'Approved',
        error_count: 0,
        warning_count: 2,
        orders_rows: 156,
        lines_rows: 432,
        order_sizes_rows: 864,
        created_at: new Date().toISOString(),
        completed_at: new Date().toISOString()
    }
];

export async function createRun(run: Partial<RunHistory>): Promise<string> {
    if (isMock) return "mock-run-" + Date.now();

    try {
        const { data, error } = await supabaseAdmin
            .from('run_history')
            .insert(run)
            .select('id')
            .single();

        if (error) throw error;
        return data.id;
    } catch (error) {
        if (process.env.NODE_ENV !== 'production') {
            console.warn('createRun failed, falling back to mock run id:', error);
            return "mock-run-" + Date.now();
        }
        throw error;
    }
}

export async function updateRun(id: string, updates: Partial<RunHistory>): Promise<void> {
    if (isMock) return;
    try {
        const { error } = await supabaseAdmin
            .from('run_history')
            .update(updates)
            .eq('id', id);

        if (error) throw error;
    } catch (error) {
        if (process.env.NODE_ENV !== 'production') {
            console.warn('updateRun failed in dev, continuing without DB update:', error);
            return;
        }
        throw error;
    }
}

export async function getRunHistory(userId: string, role: string): Promise<any[]> {
    if (isMock) return MOCK_RUNS;
    try {
        let query = supabaseAdmin
            .from('run_history')
            .select('*, users!user_id(name, email), reviewer:users!reviewed_by(name)')
            .order('created_at', { ascending: false });

        // Visibility Rules
        if (role === 'PBD Planner') {
            query = query.eq('user_id', userId);
        } else if (role === 'Read-Only') {
            // Read-Only sees summary stats only (this might be handled in the UI or a specific aggregate query)
            // For now, return empty or limit data if specific "summary" requirement is strict
        }
        // Admin, IT Manager, Reviewer see ALL (no filter)

        const { data, error } = await query;
        if (error) throw error;
        return data || [];
    } catch (error) {
        if (process.env.NODE_ENV !== 'production') {
            console.warn('getRunHistory failed in dev, returning mock data:', error);
            return MOCK_RUNS;
        }
        throw error;
    }
}

export async function getRunSummaryStats() {
    if (isMock) {
        return { totalRuns: 1, totalErrors: 0, totalWarnings: 2, Approved: 1 };
    }
    try {
        const { data, error } = await supabaseAdmin
            .from('run_history')
            .select('status, error_count, warning_count');

        if (error) throw error;

        // Aggregate stats
        return data.reduce((acc: any, run: any) => {
            acc.totalRuns++;
            acc.totalErrors += run.error_count;
            acc.totalWarnings += run.warning_count;
            acc[run.status] = (acc[run.status] || 0) + 1;
            return acc;
        }, { totalRuns: 0, totalErrors: 0, totalWarnings: 0 });
    } catch (error) {
        if (process.env.NODE_ENV !== 'production') {
            console.warn('getRunSummaryStats failed in dev, returning mock summary:', error);
            return { totalRuns: 1, totalErrors: 0, totalWarnings: 0 };
        }
        throw error;
    }
}
