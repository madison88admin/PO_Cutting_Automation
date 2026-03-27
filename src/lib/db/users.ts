import { supabaseAdmin, isMock } from '../supabase';

export interface User {
    id: string;
    name: string;
    email: string;
    role: 'Admin' | 'PBD Planner' | 'Reviewer' | 'IT Manager' | 'Read-Only';
    is_active: boolean;
    created_at?: string;
    updated_at?: string;
}

const MOCK_ADMIN: User = {
    id: '00000000-0000-0000-0000-000000000001',
    name: 'Mock Admin',
    email: 'admin@mock.local',
    role: 'Admin',
    is_active: true
};

export async function getUser(id: string): Promise<User | null> {
    if (isMock) return MOCK_ADMIN;

    const { data, error } = await supabaseAdmin
        .from('users')
        .select('*')
        .eq('id', id)
        .single();

    if (error) return null;
    return data;
}

export async function getUserByEmail(email: string): Promise<User | null> {
    if (isMock) return MOCK_ADMIN;
    const { data, error } = await supabaseAdmin
        .from('users')
        .select('*')
        .eq('email', email)
        .single();

    if (error) return null;
    return data;
}

export async function listUsers(): Promise<User[]> {
    if (isMock) return [MOCK_ADMIN];
    const { data, error } = await supabaseAdmin
        .from('users')
        .select('*')
        .order('name');

    if (error) return [];
    return data || [];
}

export async function getDefaultWorkflowUserId(): Promise<string | null> {
    if (isMock) return MOCK_ADMIN.id;

    const { data, error } = await supabaseAdmin
        .from('users')
        .select('id')
        .eq('is_active', true)
        .order('role', { ascending: true })
        .order('created_at', { ascending: true })
        .limit(1)
        .single();

    if (error || !data?.id) {
        return null;
    }

    return data.id;
}

export async function createUser(user: Partial<User>): Promise<User | null> {
    if (isMock) return { ...MOCK_ADMIN, ...user } as User;
    const { data, error } = await supabaseAdmin
        .from('users')
        .insert(user)
        .select()
        .single();

    if (error) throw error;
    return data;
}

export async function updateUserRole(id: string, role: User['role']): Promise<void> {
    if (isMock) return;
    const { error } = await supabaseAdmin
        .from('users')
        .update({ role, updated_at: new Date().toISOString() })
        .eq('id', id);

    if (error) throw error;
}
