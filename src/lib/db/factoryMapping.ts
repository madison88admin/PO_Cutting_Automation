import { supabaseAdmin, isMock } from '../supabase';
import { logEvent } from './auditLog';

export interface FactoryMapping {
    id: string;
    brand: string;
    category: string;
    product_supplier: string;
    updated_by: string;
    updated_at: string;
}

const MOCK_FACTORIES: FactoryMapping[] = [
    { id: '1', brand: 'BrandAlpha', category: 'Footwear', product_supplier: 'Factory_Alpha_Footwear', updated_by: 'mock', updated_at: '' },
    { id: '2', brand: 'BrandBeta', category: 'Apparel', product_supplier: 'Factory_Beta_Apparel', updated_by: 'mock', updated_at: '' }
];

export async function getFactoryMappings(): Promise<FactoryMapping[]> {
    if (isMock) return MOCK_FACTORIES;

    const { data, error } = await supabaseAdmin
        .from('factory_mapping')
        .select('*')
        .order('brand');

    if (error) throw error;
    return data || [];
}

export async function upsertFactory(mapping: Partial<FactoryMapping>, userId: string): Promise<void> {
    const { data: oldData } = await supabaseAdmin
        .from('factory_mapping')
        .select('*')
        .eq('brand', mapping.brand)
        .eq('category', mapping.category)
        .single();

    const { error } = await supabaseAdmin
        .from('factory_mapping')
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
            table: 'factory_mapping',
            before: oldData,
            after: mapping
        }
    });
}

export async function deleteFactory(id: string, userId: string): Promise<void> {
    const { error } = await supabaseAdmin
        .from('factory_mapping')
        .delete()
        .eq('id', id);

    if (error) throw error;

    await logEvent({
        event: 'MAPPING_TABLE_UPDATED',
        user_id: userId,
        metadata: {
            table: 'factory_mapping',
            action: 'delete',
            id
        }
    });
}
