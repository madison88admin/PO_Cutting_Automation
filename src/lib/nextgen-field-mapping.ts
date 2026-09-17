/**
 * NextGen InsertFormData field value mappings.
 *
 * The InsertFormData endpoint expects numeric IDs for dropdown fields,
 * not text names. These mappings were extracted from the NextGen
 * /PurchaseOrder/Insert page HTML (Kendo DropDownList dataSource values).
 *
 * For virtualized dropdowns (Customer, Supplier), the IDs must be resolved
 * at runtime via the NextGen valueMapper API.
 */

// Currency: Text -> Value (ID)
export const CURRENCY_MAP: Record<string, string> = {
    'USD': '1',
    'RMB': '2',
    'IDR': '3',
};

// Status: Text -> Value (ID)
export const STATUS_MAP: Record<string, string> = {
    'Confirmed': '2',
    'Unconfirmed': '1',
    'Cancelled': '3',
};

// Template: Text -> Value (ID)
export const TEMPLATE_MAP: Record<string, string> = {
    'Major Brand Bulk': '6',
    'Major Brand Bulk (COPY)': '46',
    'SMS PO Header': '29',
    'Specialty Bulk PO': '5',
};

// TransportMethod: Text -> Value (ID)
export const TRANSPORT_METHOD_MAP: Record<string, string> = {
    'Courier': '1',
    'Sea': '2',
    'Air': '3',
    'Truck': '4',
    'Air/Sea Difference': '5',
    'Boat-Truck': '6',
    'Sea-FCL': '7',
};

// Location (Deliver To): Text -> Value (ID)
// Full list from /PurchaseOrder/Insert page
export const LOCATION_MAP: Record<string, string> = {
    'Albania': '110',
    'APAC': '109',
    'Argentina': '51',
    'Australia': '42',
    'Austria': '92',
    'Bangladesh': '68',
    'Belgium': '83',
    'Benelux': '101',
    'Bersalles': '91',
    'Bolivia': '84',
    'Boston': '31',
    'Brazil': '55',
    'Busan': '25',
    'California': '29',
    'Cambodia': '125',
    'Canada': '41',
    'Central America': '90',
    'Chile': '45',
    'China': '37',
    'Columbia': '74',
    'Costa Rica': '115',
    'Croatia': '116',
    'Cyprus': '105',
    'Czech Republic': '113',
    'Denmark': '97',
    'Dubai': '69',
    'Ecuador': '98',
    'Egypt': '120',
    'El Salvador': '94',
    'Emerging Markets': '48',
    'Europe': '47',
    'Finland': '114',
    'Florence': '24',
    'FOB': '35',
    'France': '53',
    'Germany': '39',
    'Great Britain': '121',
    'Greece': '117',
    'Guang Dong': '30',
    'Hamburg': '124',
    'Holland': '96',
    'Hong Kong': '28',
    'Hungary': '56',
    'Iceland': '112',
    'India': '52',
    'Indonesia': '63',
    'Ireland': '122',
    'Israel': '64',
    'Italy': '89',
    'Japan': '67',
    'Kazakhstan': '119',
    'Korea': '32',
    'Kuwait': '86',
    'Latvia': '123',
    'Lebanon': '87',
    'Lima': '72',
    'London': '103',
    'Long Beach': '26',
    'Los Angeles': '27',
    'Luxembourg': '111',
    'M88 QC Team': '99',
    'Malaysia': '59',
    'Malta': '106',
    'Mexico': '66',
    'Monaco': '76',
    'Mongolia': '49',
    'Nepal': '60',
    'Netherlands': '46',
    'New York': '22',
    'New Zealand': '79',
    'Nigeria': '107',
    'Ningbo': '33',
    'Norway': '85',
    'Panama': '57',
    'Paraguay': '77',
    'Peru': '80',
    'Philippines': '58',
    'Poland': '78',
    'Portugal': '118',
    'Romania': '81',
    'Russia': '73',
    'Serbia': '75',
    'Shanghai': '23',
    'Singapore': '61',
    'Slovenia': '93',
    'South Africa': '70',
    'South Korea': '100',
    'Spain': '95',
    'Sri Lanka': '104',
    'Sweden': '71',
    'Switzerland': '38',
    'Taipei': '34',
    'Taiwan': '82',
    'TBC': '108',
    'Thailand': '54',
    'Tunisia': '102',
    'Turkey': '65',
    'UAE': '50',
    'UK': '40',
    'Ukraine': '88',
    'Uruguay': '62',
    'USA': '36',
    'Vietnam': '44',
    'Not Set': '43',
};

// Reverse lookup: Value (ID) -> Text
export const LOCATION_REVERSE_MAP: Record<string, string> = Object.fromEntries(
    Object.entries(LOCATION_MAP).map(([k, v]) => [v, k])
);

/**
 * Resolve a text value to its numeric ID using the provided map.
 * If the value is already numeric, return as-is.
 * If not found, return empty string (will cause NextGen to reject).
 */
export function resolveId(value: string, map: Record<string, string>): string {
    if (!value) return '';
    const trimmed = value.trim();
    // Already numeric ID
    if (/^\d+$/.test(trimmed)) return trimmed;
    // Try exact match
    if (map[trimmed]) return map[trimmed];
    // Try case-insensitive match
    const lower = trimmed.toLowerCase();
    for (const [key, id] of Object.entries(map)) {
        if (key.toLowerCase() === lower) return id;
    }
    // Try partial match (location name contains the value)
    for (const [key, id] of Object.entries(map)) {
        if (key.toLowerCase().includes(lower) || lower.includes(key.toLowerCase())) {
            return id;
        }
    }
    return '';
}
