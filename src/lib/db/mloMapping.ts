import { supabaseAdmin, isMock } from '../supabase';
import { logEvent } from './auditLog';

export interface MloMapping {
    id: string;
    brand: string;
    keyuser1: string;
    keyuser2: string;
    keyuser4: string;
    keyuser5: string;
    orders_template?: string | null;
    lines_template?: string | null;
    valid_statuses?: string[] | null;
    updated_by: string;
    updated_at: string;
}

const MOCK_MLO: MloMapping[] = [
    {
        id: '1',
        brand: 'BrandAlpha',
        keyuser1: 'KU-ALPHA-1',
        keyuser2: 'KU-ALPHA-2',
        keyuser4: 'KU-ALPHA-4',
        keyuser5: 'KU-ALPHA-5',
        orders_template: 'BULK',
        lines_template: 'BULK',
        valid_statuses: ['Confirmed'],
        updated_by: 'mock',
        updated_at: '',
    },
    {
        id: '2',
        brand: 'BrandBeta',
        keyuser1: 'KU-BETA-1',
        keyuser2: 'KU-BETA-2',
        keyuser4: 'KU-BETA-4',
        keyuser5: 'KU-BETA-5',
        orders_template: 'BULK',
        lines_template: 'BULK',
        valid_statuses: ['Confirmed'],
        updated_by: 'mock',
        updated_at: '',
    }
];

export async function getMloMappings(): Promise<MloMapping[]> {
    if (isMock) return MOCK_MLO;

    const { data, error } = await supabaseAdmin
        .from('mlo_mapping')
        .select('*')
        .order('brand');

    if (error) throw error;
    return data || [];
}

export async function upsertMlo(mapping: Partial<MloMapping>, userId: string): Promise<void> {
    const { data: oldData } = await supabaseAdmin
        .from('mlo_mapping')
        .select('*')
        .eq('brand', mapping.brand)
        .single();

    const { error } = await supabaseAdmin
        .from('mlo_mapping')
        .upsert({
            ...mapping,
            updated_by: userId,
            updated_at: new Date().toISOString()
        });

    if (error) throw error;

    await logEvent({
        event: 'MAPPING_TABLE_UPDATED',
        user_id: userId,
        metadata: {
            table: 'mlo_mapping',
            before: oldData,
            after: mapping
        }
    });
}
