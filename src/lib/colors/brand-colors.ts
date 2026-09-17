/**
 * Brand-specific color code → color name mappings.
 *
 * Different brands use different color code systems:
 * - Columbia: 3-digit numeric (010 = Black, 422 = Shark, etc.)
 * - The North Face: 3-char alphanumeric (E6Q = TNF Black, JK3 = ...)
 * - Arc'teryx: 3-digit numeric (001 = Black Sapphire, 890 = Forage)
 *
 * This module provides normalization so that color names can be resolved
 * even when the buy file only has a color code.
 */

// Columbia Sportswear color codes (3-digit)
const COLUMBIA_COLORS: Record<string, string> = {
    '010': 'Black',
    '011': 'White',
    '012': 'Graphite',
    '021': 'Collegiate Navy',
    '022': 'Dark Blue',
    '025': 'Slate Blue',
    '031': 'Red',
    '032': 'Bright Red',
    '035': 'Cardinal',
    '041': 'Forest',
    '042': 'Olive',
    '043': 'Sage',
    '051': 'Brown',
    '052': 'Dark Brown',
    '061': 'Grey',
    '062': 'Heather Grey',
    '071': 'Yellow',
    '072': 'Gold',
    '081': 'Orange',
    '082': 'Coral',
    '091': 'Purple',
    '092': 'Violet',
    '101': 'Teal',
    '102': 'Aqua',
    '111': 'Pink',
    '112': 'Rose',
    '121': 'Tan',
    '122': 'Khaki',
    '123': 'Sand',
    '131': 'Charcoal',
    '132': 'Slate',
    '133': 'Stone',
    '141': 'Green',
    '142': 'Lime',
    '143': 'Mint',
    '151': 'Sky',
    '152': 'Cyan',
    '153': 'Blue',
    '161': 'Maroon',
    '162': 'Burgundy',
    '171': 'Ivory',
    '172': 'Cream',
    '173': 'Bone',
    '422': 'Shark',
    '445': 'Wave',
    '464': 'Pine',
    '478': 'Frost',
    '493': 'Sunset',
    '553': 'Glacier',
    '614': 'Campfire',
    '625': 'Sedona',
    '632': 'Hops',
    '641': 'Sawtooth',
    '684': 'Pebble',
    '796': 'Jade',
    '837': 'Coral',
    '887': 'Lichen',
    '920': 'Cobalt',
    '922': 'Sapphire',
    '924': 'Amethyst',
    '935': 'Sunbeam',
    '944': 'Cinnamon',
    '952': 'Lagoon',
    '962': 'Hollyhock',
    '974': 'Goldenrod',
    '984': 'Mineral',
    '993': 'Ember',
};

// The North Face color codes (3-char alphanumeric)
const TNF_COLORS: Record<string, string> = {
    'E6Q': 'TNF Black',
    'JK3': 'TNF White',
    'J9W': 'Tnf Red',
    'K3J': 'Tnf Blue',
    'L5R': 'Asphalt Grey',
    'NF0': 'TNF Black',
    'NF1': 'TNF White',
    'NF2': 'TNF Red',
    'NF3': 'TNF Blue',
    'NF4': 'TNF Green',
    'NF5': 'TNF Yellow',
    'NF6': 'TNF Orange',
    'NF7': 'TNF Purple',
    'NF8': 'TNF Brown',
    'NF9': 'TNF Grey',
};

// Arc'teryx color codes (3-digit)
const ARCTERYX_COLORS: Record<string, string> = {
    '001': 'Black Sapphire',
    '002': 'Black',
    '003': 'Jet Black',
    '010': 'Black',
    '085': 'Bourbon',
    '100': 'White',
    '190': 'Raven',
    '195': 'Black / Raptor',
    '280': 'Hyper Blue',
    '290': 'Cosmic Blue',
    '300': 'Pegasus',
    '330': 'Triton',
    '350': 'Aether',
    '360': 'Aurora',
    '370': 'Borealis',
    '380': 'Galaxy',
    '390': 'Navy',
    '420': 'Pilot',
    '443': 'Pulsar',
    '460': 'Cobalt',
    '480': 'Savannah',
    '500': 'Conifer Green',
    '510': 'Forest',
    '520': 'Crocodile',
    '540': 'Moss',
    '550': 'Fern',
    '560': 'Boreal Green',
    '570': 'Tropic',
    '580': 'Tundra',
    '590': 'Lime',
    '600': 'Firecracker',
    '610': 'Flame',
    '620': 'Rouge',
    '630': 'Terra',
    '640': 'Tomato',
    '650': 'Ember',
    '660': 'Clay',
    '670': 'Rust',
    '680': 'Cayenne',
    '690': 'Coral',
    '700': 'Sunshine',
    '710': 'Gold',
    '720': 'Amber',
    '730': 'Honey',
    '740': 'Wheat',
    '750': 'Sand',
    '760': 'Bone',
    '770': 'Ivory',
    '780': 'Cream',
    '790': 'Fossil',
    '800': 'Pewter',
    '810': 'Slate',
    '820': 'Graphite',
    '830': 'Charcoal',
    '840': 'Ash',
    '850': 'Storm',
    '860': 'Cloud',
    '870': 'Glacier',
    '880': 'Ice',
    '890': 'Forage',
    '900': 'Magenta',
    '910': 'Plum',
    '920': 'Garnet',
    '930': 'Wine',
    '940': 'Bordeaux',
    '950': 'Orchid',
    '960': 'Lilac',
    '970': 'Violet',
    '980': 'Amethyst',
    '990': 'Purple',
};

// Generic fallback for common color codes
const GENERIC_COLORS: Record<string, string> = {
    'blk': 'Black',
    'wht': 'White',
    'red': 'Red',
    'blu': 'Blue',
    'grn': 'Green',
    'yel': 'Yellow',
    'org': 'Orange',
    'pur': 'Purple',
    'pnk': 'Pink',
    'gry': 'Grey',
    'brn': 'Brown',
    'tan': 'Tan',
    'nav': 'Navy',
    'teal': 'Teal',
    'char': 'Charcoal',
    'caml': 'Camel',
    'sand': 'Sand',
    'khk': 'Khaki',
    'olive': 'Olive',
    'crd': 'Cardinal',
    'roy': 'Royal',
    'sge': 'Sage',
    'ston': 'Stone',
    'slate': 'Slate',
    'heather': 'Heather',
};

export type Brand = 'columbia' | 'tnf' | 'arcteryx' | 'generic';

const BRAND_MAPS: Record<Brand, Record<string, string>> = {
    columbia: COLUMBIA_COLORS,
    tnf: TNF_COLORS,
    arcteryx: ARCTERYX_COLORS,
    generic: GENERIC_COLORS,
};

/**
 * Detect brand from style number or customer name.
 */
export function detectBrand(style: string, customer?: string): Brand {
    const s = (style || '').toUpperCase();
    const c = (customer || '').toUpperCase();

    if (s.startsWith('NF') || s.startsWith('NF0') || c.includes('NORTH FACE') || c.includes('TNF')) return 'tnf';
    if (s.startsWith('WM') || s.startsWith('AR') || s.startsWith('SL') || c.includes('COLUMBIA')) return 'columbia';
    if (s.startsWith('X0') || c.includes('ARCTERYX') || c.includes("ARC'TERYX")) return 'arcteryx';

    return 'generic';
}

/**
 * Resolve a color code to a color name using brand-specific mappings.
 * Falls back to generic mappings if the brand-specific map doesn't have it.
 */
export function resolveColorName(colorCode: string, brand?: Brand): string | null {
    const code = String(colorCode || '').trim().toUpperCase();
    if (!code) return null;

    // Try brand-specific first
    if (brand && brand !== 'generic') {
        const brandMap = BRAND_MAPS[brand];
        if (brandMap[code]) return brandMap[code];
        // Try case-insensitive
        const lower = code.toLowerCase();
        for (const [key, val] of Object.entries(brandMap)) {
            if (key.toUpperCase() === code) return val;
        }
    }

    // Try generic
    if (GENERIC_COLORS[code.toLowerCase()]) return GENERIC_COLORS[code.toLowerCase()];

    return null;
}
