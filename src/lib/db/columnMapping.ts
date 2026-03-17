import { supabaseAdmin, isMock } from '../supabase';
import { logEvent } from './auditLog';

export interface ColumnMapping {
    id: string;
    customer: string;
    buy_file_column: string;
    internal_field: string;
    notes?: string;
    updated_by: string;
    updated_at: string;
}

const MOCK_COLUMNS: ColumnMapping[] = [
    // ABC Retail Mappings
    { id: '1', customer: 'ABC Retail', buy_file_column: 'PO Number', internal_field: 'purchaseOrder', updated_by: 'mock', updated_at: '' },
    { id: '2', customer: 'ABC Retail', buy_file_column: 'Brand', internal_field: 'brand', updated_by: 'mock', updated_at: '' },
    { id: '3', customer: 'ABC Retail', buy_file_column: 'Category', internal_field: 'category', updated_by: 'mock', updated_at: '' },
    { id: '4', customer: 'ABC Retail', buy_file_column: 'Customer Name', internal_field: 'customerName', updated_by: 'mock', updated_at: '' },
    { id: '5', customer: 'ABC Retail', buy_file_column: 'Customer Type', internal_field: 'customerType', updated_by: 'mock', updated_at: '' },
    { id: '6', customer: 'ABC Retail', buy_file_column: 'Season', internal_field: 'season', updated_by: 'mock', updated_at: '' },
    { id: '7', customer: 'ABC Retail', buy_file_column: 'Buy Month', internal_field: 'buyMonth', updated_by: 'mock', updated_at: '' },
    { id: '8', customer: 'ABC Retail', buy_file_column: 'Buy Date', internal_field: 'buyDate', updated_by: 'mock', updated_at: '' },
    { id: '9', customer: 'ABC Retail', buy_file_column: 'Ex-Fty Date', internal_field: 'exFtyDate', updated_by: 'mock', updated_at: '' },
    { id: '10', customer: 'ABC Retail', buy_file_column: 'Cancel Date', internal_field: 'cancelDate', updated_by: 'mock', updated_at: '' },
    { id: '11', customer: 'ABC Retail', buy_file_column: 'PO Issuance Date', internal_field: 'poIssuanceDate', updated_by: 'mock', updated_at: '' },
    { id: '12', customer: 'ABC Retail', buy_file_column: 'Delivery Date', internal_field: 'deliveryDate', updated_by: 'mock', updated_at: '' },
    { id: '13', customer: 'ABC Retail', buy_file_column: 'DC Code', internal_field: 'dcCode', updated_by: 'mock', updated_at: '' },
    { id: '14', customer: 'ABC Retail', buy_file_column: 'Shipment Location', internal_field: 'transportLocation', updated_by: 'mock', updated_at: '' },
    { id: '15', customer: 'ABC Retail', buy_file_column: 'Transport Method', internal_field: 'transportMethod', updated_by: 'mock', updated_at: '' },
    { id: '16', customer: 'ABC Retail', buy_file_column: 'Template Type', internal_field: 'template', updated_by: 'mock', updated_at: '' },
    { id: '17', customer: 'ABC Retail', buy_file_column: 'Style Number', internal_field: 'product', updated_by: 'mock', updated_at: '' },
    { id: '18', customer: 'ABC Retail', buy_file_column: 'Color', internal_field: 'colour', updated_by: 'mock', updated_at: '' },
    { id: '19', customer: 'ABC Retail', buy_file_column: 'Size', internal_field: 'sizeName', updated_by: 'mock', updated_at: '' },
    { id: '20', customer: 'ABC Retail', buy_file_column: 'Quantity', internal_field: 'quantity', updated_by: 'mock', updated_at: '' },
    { id: '21', customer: 'ABC Retail', buy_file_column: 'Buyer PO Number', internal_field: 'buyerPoNumber', updated_by: 'mock', updated_at: '' },
    { id: '22', customer: 'ABC Retail', buy_file_column: 'Status', internal_field: 'status', updated_by: 'mock', updated_at: '' },
    // XYZ Sports Mappings
    { id: '23', customer: 'XYZ Sports', buy_file_column: 'PO Number', internal_field: 'purchaseOrder', updated_by: 'mock', updated_at: '' },
    { id: '24', customer: 'XYZ Sports', buy_file_column: 'Brand', internal_field: 'brand', updated_by: 'mock', updated_at: '' },
    { id: '25', customer: 'XYZ Sports', buy_file_column: 'Category', internal_field: 'category', updated_by: 'mock', updated_at: '' },
    { id: '26', customer: 'XYZ Sports', buy_file_column: 'Customer Name', internal_field: 'customerName', updated_by: 'mock', updated_at: '' },
    { id: '27', customer: 'XYZ Sports', buy_file_column: 'Customer Type', internal_field: 'customerType', updated_by: 'mock', updated_at: '' },
    { id: '28', customer: 'XYZ Sports', buy_file_column: 'Season', internal_field: 'season', updated_by: 'mock', updated_at: '' },
    { id: '29', customer: 'XYZ Sports', buy_file_column: 'Buy Month', internal_field: 'buyMonth', updated_by: 'mock', updated_at: '' },
    { id: '30', customer: 'XYZ Sports', buy_file_column: 'Buy Date', internal_field: 'buyDate', updated_by: 'mock', updated_at: '' },
    { id: '31', customer: 'XYZ Sports', buy_file_column: 'Ex-Fty Date', internal_field: 'exFtyDate', updated_by: 'mock', updated_at: '' },
    { id: '32', customer: 'XYZ Sports', buy_file_column: 'Cancel Date', internal_field: 'cancelDate', updated_by: 'mock', updated_at: '' },
    { id: '33', customer: 'XYZ Sports', buy_file_column: 'PO Issuance Date', internal_field: 'poIssuanceDate', updated_by: 'mock', updated_at: '' },
    { id: '34', customer: 'XYZ Sports', buy_file_column: 'Delivery Date', internal_field: 'deliveryDate', updated_by: 'mock', updated_at: '' },
    { id: '35', customer: 'XYZ Sports', buy_file_column: 'DC Code', internal_field: 'dcCode', updated_by: 'mock', updated_at: '' },
    { id: '36', customer: 'XYZ Sports', buy_file_column: 'Shipment Location', internal_field: 'transportLocation', updated_by: 'mock', updated_at: '' },
    { id: '37', customer: 'XYZ Sports', buy_file_column: 'Transport Method', internal_field: 'transportMethod', updated_by: 'mock', updated_at: '' },
    { id: '38', customer: 'XYZ Sports', buy_file_column: 'Template Type', internal_field: 'template', updated_by: 'mock', updated_at: '' },
    { id: '39', customer: 'XYZ Sports', buy_file_column: 'Style Number', internal_field: 'product', updated_by: 'mock', updated_at: '' },
    { id: '40', customer: 'XYZ Sports', buy_file_column: 'Color', internal_field: 'colour', updated_by: 'mock', updated_at: '' },
    { id: '41', customer: 'XYZ Sports', buy_file_column: 'Size', internal_field: 'sizeName', updated_by: 'mock', updated_at: '' },
    { id: '42', customer: 'XYZ Sports', buy_file_column: 'Quantity', internal_field: 'quantity', updated_by: 'mock', updated_at: '' },
    { id: '43', customer: 'XYZ Sports', buy_file_column: 'Buyer PO Number', internal_field: 'buyerPoNumber', updated_by: 'mock', updated_at: '' },
    { id: '44', customer: 'XYZ Sports', buy_file_column: 'Status', internal_field: 'status', updated_by: 'mock', updated_at: '' },
    // DEFAULT / NextGen Standard Mappings (Comprehensive for Round-Trip)
    { id: '45', customer: 'DEFAULT', buy_file_column: 'PurchaseOrder', internal_field: 'purchaseOrder', updated_by: 'mock', updated_at: '' },
    { id: '46', customer: 'DEFAULT', buy_file_column: 'PO Number', internal_field: 'purchaseOrder', updated_by: 'mock', updated_at: '' },
    { id: '47', customer: 'DEFAULT', buy_file_column: 'PO#', internal_field: 'purchaseOrder', updated_by: 'mock', updated_at: '' },
    { id: '48', customer: 'DEFAULT', buy_file_column: 'Brand', internal_field: 'brand', updated_by: 'mock', updated_at: '' },
    { id: '49', customer: 'DEFAULT', buy_file_column: 'Category', internal_field: 'category', updated_by: 'mock', updated_at: '' },
    { id: '50', customer: 'DEFAULT', buy_file_column: 'Dept', internal_field: 'category', updated_by: 'mock', updated_at: '' },
    { id: '51', customer: 'DEFAULT', buy_file_column: 'ProductSupplier', internal_field: 'productSupplier', updated_by: 'mock', updated_at: '' },
    { id: '52', customer: 'DEFAULT', buy_file_column: 'Customer', internal_field: 'customerName', updated_by: 'mock', updated_at: '' },
    { id: '53', customer: 'DEFAULT', buy_file_column: 'Customer Name', internal_field: 'customerName', updated_by: 'mock', updated_at: '' },
    { id: '54', customer: 'DEFAULT', buy_file_column: 'TransportMethod', internal_field: 'transportMethod', updated_by: 'mock', updated_at: '' },
    { id: '55', customer: 'DEFAULT', buy_file_column: 'Transport Method', internal_field: 'transportMethod', updated_by: 'mock', updated_at: '' },
    { id: '56', customer: 'DEFAULT', buy_file_column: 'Template', internal_field: 'template', updated_by: 'mock', updated_at: '' },
    { id: '57', customer: 'DEFAULT', buy_file_column: 'Template Type', internal_field: 'template', updated_by: 'mock', updated_at: '' },
    { id: '58', customer: 'DEFAULT', buy_file_column: 'KeyDate', internal_field: 'poIssuanceDate', updated_by: 'mock', updated_at: '' },
    { id: '59', customer: 'DEFAULT', buy_file_column: 'DeliveryDate', internal_field: 'exFtyDate', updated_by: 'mock', updated_at: '' },
    { id: '60', customer: 'DEFAULT', buy_file_column: 'KeyUser1', internal_field: 'keyUser1', updated_by: 'mock', updated_at: '' },
    { id: '61', customer: 'DEFAULT', buy_file_column: 'KeyUser2', internal_field: 'keyUser2', updated_by: 'mock', updated_at: '' },
    { id: '62', customer: 'DEFAULT', buy_file_column: 'KeyUser4', internal_field: 'keyUser4', updated_by: 'mock', updated_at: '' },
    { id: '63', customer: 'DEFAULT', buy_file_column: 'KeyUser5', internal_field: 'keyUser5', updated_by: 'mock', updated_at: '' },
    { id: '64', customer: 'DEFAULT', buy_file_column: 'Product', internal_field: 'product', updated_by: 'mock', updated_at: '' },
    { id: '65', customer: 'DEFAULT', buy_file_column: 'Style Number', internal_field: 'product', updated_by: 'mock', updated_at: '' },
    { id: '66', customer: 'DEFAULT', buy_file_column: 'SizeName', internal_field: 'sizeName', updated_by: 'mock', updated_at: '' },
    { id: '67', customer: 'DEFAULT', buy_file_column: 'ProductSize', internal_field: 'sizeName', updated_by: 'mock', updated_at: '' },
    { id: '68', customer: 'DEFAULT', buy_file_column: 'Quantity', internal_field: 'quantity', updated_by: 'mock', updated_at: '' },
    { id: '69', customer: 'DEFAULT', buy_file_column: 'Status', internal_field: 'status', updated_by: 'mock', updated_at: '' },
    { id: '70', customer: 'DEFAULT', buy_file_column: 'Colour', internal_field: 'colour', updated_by: 'mock', updated_at: '' },
    { id: '71', customer: 'DEFAULT', buy_file_column: 'Color', internal_field: 'colour', updated_by: 'mock', updated_at: '' },
    { id: '72', customer: 'DEFAULT', buy_file_column: 'LineItem', internal_field: 'lineItem', updated_by: 'mock', updated_at: '' },
    { id: '73', customer: 'DEFAULT', buy_file_column: 'ProductRange', internal_field: 'productRange', updated_by: 'mock', updated_at: '' },
    { id: '74', customer: 'DEFAULT', buy_file_column: 'Range', internal_field: 'productRange', updated_by: 'mock', updated_at: '' },
    { id: '75', customer: 'DEFAULT', buy_file_column: 'UDF-buyer_po_number', internal_field: 'buyerPoNumber', updated_by: 'mock', updated_at: '' },
    { id: '76', customer: 'DEFAULT', buy_file_column: 'UDF-start_date', internal_field: 'startDate', updated_by: 'mock', updated_at: '' },
    { id: '77', customer: 'DEFAULT', buy_file_column: 'UDF-canel_date', internal_field: 'cancelDate', updated_by: 'mock', updated_at: '' },
    // Silent Mapping for auxiliary system headers (No warning)
    { id: '78', customer: 'DEFAULT', buy_file_column: 'TransportLocation', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '79', customer: 'DEFAULT', buy_file_column: 'PaymentTerm', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '80', customer: 'DEFAULT', buy_file_column: 'ClosedDate', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '81', customer: 'DEFAULT', buy_file_column: 'ArchiveDate', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '82', customer: 'DEFAULT', buy_file_column: 'Currency', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '83', customer: 'DEFAULT', buy_file_column: 'FindField_Product', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '84', customer: 'DEFAULT', buy_file_column: 'FindField_ProductSupplier', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '85', customer: 'DEFAULT', buy_file_column: 'ERP Ind', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '86', customer: 'DEFAULT', buy_file_column: 'Material Style', internal_field: 'product', updated_by: 'mock', updated_at: '' },
    { id: '87', customer: 'DEFAULT', buy_file_column: 'JDE Style', internal_field: 'product', updated_by: 'mock', updated_at: '' },
    { id: '88', customer: 'DEFAULT', buy_file_column: 'Ordered Qty', internal_field: 'quantity', updated_by: 'mock', updated_at: '' },
    { id: '89', customer: 'DEFAULT', buy_file_column: 'Orig Ex Fac', internal_field: 'exFtyDate', updated_by: 'mock', updated_at: '' },
    { id: '90', customer: 'DEFAULT', buy_file_column: 'Buy Date', internal_field: 'buyDate', updated_by: 'mock', updated_at: '' },
    { id: '91', customer: 'DEFAULT', buy_file_column: 'Doc Type', internal_field: 'template', updated_by: 'mock', updated_at: '' },
    { id: '92', customer: 'DEFAULT', buy_file_column: 'Trans Cond', internal_field: 'transportMethod', updated_by: 'mock', updated_at: '' },
    { id: '93', customer: 'DEFAULT', buy_file_column: 'Ult. Destination', internal_field: 'transportLocation', updated_by: 'mock', updated_at: '' },
    { id: '94', customer: 'DEFAULT', buy_file_column: 'Company Code', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '95', customer: 'DEFAULT', buy_file_column: 'Vendor Code', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '96', customer: 'DEFAULT', buy_file_column: 'AB Number', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '97', customer: 'DEFAULT', buy_file_column: 'GTN Issue Date', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '98', customer: 'DEFAULT', buy_file_column: 'Confirmation Status', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '99', customer: 'DEFAULT', buy_file_column: 'SKU Status', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '100', customer: 'DEFAULT', buy_file_column: 'SLO', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '101', customer: 'DEFAULT', buy_file_column: 'PLO', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '102', customer: 'DEFAULT', buy_file_column: 'Plant', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '103', customer: 'DEFAULT', buy_file_column: 'Dest Country', internal_field: 'transportLocation', updated_by: 'mock', updated_at: '' },
    { id: '104', customer: 'DEFAULT', buy_file_column: 'Line #s', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '105', customer: 'DEFAULT', buy_file_column: 'Priority Flag', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '106', customer: 'DEFAULT', buy_file_column: 'Material Name', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '107', customer: 'DEFAULT', buy_file_column: 'LB', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '108', customer: 'DEFAULT', buy_file_column: 'Tooling Code', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '109', customer: 'DEFAULT', buy_file_column: 'VAS', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '110', customer: 'DEFAULT', buy_file_column: 'UOM', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '111', customer: 'DEFAULT', buy_file_column: 'Collaboration Status', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '112', customer: 'DEFAULT', buy_file_column: 'Last Modifier', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '113', customer: 'DEFAULT', buy_file_column: 'Last Updated', internal_field: 'ignore', updated_by: 'mock', updated_at: '' },
    { id: '114', customer: 'DEFAULT', buy_file_column: 'Locked', internal_field: 'ignore', updated_by: 'mock', updated_at: '' }
];

export async function getColumnMappings(customer?: string): Promise<ColumnMapping[]> {
    if (isMock) {
        if (!customer) return MOCK_COLUMNS;
        return MOCK_COLUMNS.filter(m => m.customer === customer);
    }

    let query = supabaseAdmin
        .from('column_mapping')
        .select('*')
        .order('customer');

    if (customer) query = query.eq('customer', customer);

    const { data, error } = await query;
    if (error) throw error;
    return data || [];
}

export async function upsertColumn(mapping: Partial<ColumnMapping>, userId: string): Promise<void> {
    const { data: oldData } = await supabaseAdmin
        .from('column_mapping')
        .select('*')
        .eq('customer', mapping.customer)
        .eq('buy_file_column', mapping.buy_file_column)
        .single();

    const { error } = await supabaseAdmin
        .from('column_mapping')
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
            table: 'column_mapping',
            before: oldData,
            after: mapping
        }
    });
}
