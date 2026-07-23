export interface BuyFileItem {
    style: string | null;
    styleName: string | null;
    sku: string | null;
    description: string | null;
    color: string | null;
    colorCode: string | null;
    colorName: string | null;
    size: string | null;
    quantity: number | null;
    deliveryDate: string | null;
    season: string | null;
    customer: string | null;
    factory: string | null;
    currency: string | null;
    unitCost: number | null;
    poNumber: string | null;
    product: string | null;
    productExternalRef: string | null;
    costingReference: string | null;
    matchStatus: 'matched' | 'ambiguous' | 'unmatched' | 'not_checked';
    matchScore: number | null;
    matchReason: string | null;
    sourceSheet: string;
    sourceRow: number;
}

export interface ColumnMapping {
    buyer_style_number?: string;
    buyer_style_name?: string;
    sku?: string;
    product_description?: string;
    color?: string;
    color_code?: string;
    size?: string;
    quantity?: string;
    delivery_date?: string;
    season?: string;
    customer?: string;
    factory?: string;
    currency?: string;
    unit_cost?: string;
    po_number?: string;
    buyer_po_number?: string;
    start_date?: string;
    cancel_date?: string;
    transport_method?: string;
}

export interface HeaderDetectionResult {
    headerRow: number;
}

export interface HeaderMappingResult {
    mapping: ColumnMapping;
    confidence: number;
    unmappedColumns: string[];
}

export interface ExtractedTemplate {
    id: string;
    customer: string | null;
    headers: string[];
    normalizedHeaders: string[];
    mapping: ColumnMapping;
    detectedAt: string;
}

// Rich product data returned by NextGen for a single style
export interface NextGenStyleInfo {
    style: string;
    product?: string | null;
    productRange?: string | null;
    productExternalRef?: string | null;
    productCustomerRef?: string | null;
    styleName?: string | null;
    brand?: string | null;
    season?: string | null;
    department?: string | null;
    colorName?: string | null;
    colorCode?: string | null;
    colorExt?: string | null;
    sizeScale?: string | null;
    purchaseUOM?: string | null;
    sellingUOM?: string | null;
    supplierProfile?: string | null;
    customer?: string | null;
    factory?: string | null;
    currency?: string | null;
    matchStatus?: 'matched' | 'ambiguous';
    matchScore?: number | null;
    matchReason?: string | null;
    candidateCount?: number;
}

// Single source of truth used by all generators
export interface ProductData {
    style: string;
    product: string | null;
    productRange: string | null;
    productExternalRef: string | null;
    productCustomerRef: string | null;
    styleName: string | null;
    brand: string | null;
    season: string | null;
    department: string | null;
    colorName: string | null;
    colorCode: string | null;
    colorExt: string | null;
    supplierProfile: string | null;
    customer: string | null;
    factory: string | null;
    currency: string | null;
    purchaseUOM: string | null;
    sellingUOM: string | null;
    poNumber: string | null;
    deliveryDate: string | null;
    unitCost: number | null;
    costingReference: string | null;
    matchStatus: BuyFileItem['matchStatus'];
    matchScore: number | null;
    matchReason: string | null;
    sizes: ProductSize[];
}

export interface ProductSize {
    size: string;
    quantity: number;
}
