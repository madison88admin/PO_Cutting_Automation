import ExcelJS from "exceljs";
import { logEvent } from "@/lib/audit";
import { getFactoryMapping, getMloMapping, getColumnMapping, getAllColumnMappings } from "@/lib/data-loader";
import { updateRun } from "@/lib/db/runHistory";
import {
    FALLBACK_COLUMN_ALIASES,
    detectPivotFormatFromHeaders,
    formatManualDate,
    formatStandardDate,
    isLikelyPivotSizeHeader,
    normalizeHeaderText,
    parseExcelEngineDate,
    type PivotFormatDetection,
} from "./excel-engine-helpers";

// Destination code mapping for Helly Hansen and similar
import destinationMapping from "@/config/destination-mapping.json";

export interface POHeader {
    purchaseOrder: string;
    brandKey?: string;
    productSupplier: string;
    status: string;
    customer: string;
    transportMethod: string;
    transportLocation: string;
    ordersTemplate: string;
    linesTemplate: string;
    keyDate: string | Date;
    keyDateFormat?: "manual" | "standard";
    comments: string;
    currency: string;
    keyUser1: string;
    keyUser2: string;
    keyUser3: string;
    keyUser4: string;
    keyUser5: string;
    keyUser6: string;
    keyUser7: string;
    keyUser8: string;
}

export interface POLine {
    lineItem: number;
    productRange: string;
    styleNumber: string;
    supplierProfile: string;
    buyerPoNumber: string | number;
    startDate: string | Date;
    cancelDate: string | Date;
    dynafitLineKeyDate?: string | Date;
    hhStartDate?: string;
    hhCancelDate?: string;
    hhConfirmedDeliveryDate?: string;
    transportLocation?: string;
    styleColor?: string;
    rawColour?: string;
    ourReference?: string;
    cost?: string | number;
    colour: string;
    productExternalRef: string;
    productCustomerRef: string;
}

export interface POSize {
    productSize: string;
    quantity: number;
}

interface ProductSheetRow {
    colour: string;
    colourName?: string;
    factory?: string;
    cost?: string | number;
    customerName?: string;
    productName?: string;
    productExternalRef?: string;
    buyerStyleNumber?: string;
    season?: string;
    crd?: string | Date;
    exFactory?: string | Date;
    poNumber?: string;
    destinationName?: string;
}

export interface ValidationError {
    field: string;
    row: number;
    message: string;
    severity: "CRITICAL" | "WARNING";
}

export interface FormatDetection {
    detectedCustomer: string;
    detectedFormat: string;
    unmappedColumns: string[];
}

export interface ProcessedPO {
    header: POHeader;
    lines: POLine[];
    sizes: Record<number, POSize[]>;
    orderKeys?: Array<{ purchaseOrder: string; customer: string; customerName: string | undefined; transportLocation: string; transportMethod: string; ordersTemplate: string }>;
    llBeanReferenceSizeRows?: Array<{ purchaseOrder: string; lineItem: number; range: string; product: string; sizeName: string; productSize: string; quantity: number; colour: string }>;
    manualKeyDate?: string;
}

const PLANT_COUNTRY_MAP: Record<string, string> = {
    'visalia dc':               'USA',
    'visalia':                  'USA',
    'jonestown dc':             'USA',
    'jonestown':                'USA',
    'brampton dc':              'Canada',
    'brampton':                 'Canada',
    'dropship us':              'USA',
    'dropship international':   'USA',
    'dropship dc':              'USA',
    'dropship ca':              'Canada',
    'vf outdoor mexico':        'Mexico',
    'vf outdoor mexico s de r l d': 'Mexico',
    'photoshooting':            'BELGIUM',
    'eu main':                  'BELGIUM',
    'eu uk':                    'UK',
    'eu':                       'EU',
    'japan':                    'Japan',
    'korea':                    'Korea',
    'australia':                'Australia',
    'hong kong':                'Hong Kong',
    'china':                    'China',
    'virtual':                  'Dubai',
    'argentina':                'Argentina',
    'brazil':                   'Brazil',
    'chile':                    'Chile',
    'guatemala':                'Guatemala',
    'panama':                   'Panama',
    'peru':                     'PERU',
    'uruguay':                  'URUGUAY',
    'united arab emirates':     'UNITED ARAB EMIRATES',
    'singapore':                'Singapore',
    'apdindc':                  'Singapore',
    'israel':                   'Israel',
    'south africa':             'South Africa',
    'taiwan':                   'Taiwan',
    'thailand':                 'Thailand',
    'malaysia':                 'Malaysia',
    'nepal':                    'Nepal',
    'indonesia':                'Indonesia',
    '1001': 'USA',
    '1000': 'USA',
    '1010': 'USA',
    '1020': 'USA',
    '1004': 'Canada',
    '1009': 'USA',
    '1005': 'Mexico',
    't909': 'Japan',
    'd060': 'BELGIUM',
    'd080': 'UK',
    'vd60': 'Dubai',
    '0010': 'USA',
    '0011': 'Canada',
    '126': 'Australia',
    '920': 'USA',
    '120': 'Iceland',
    '0040': 'Netherlands',
    '0050': 'Singapore',
    '0060': 'UK',
    '10':   'USA',
    '11':   'Canada',
    '40':   'Netherlands',
    '50':   'Singapore',
    '60':   'UK',
    '3020': 'Sweden',
    '5001': 'Hong Kong',
    '500025': 'Korea',
    '1656': 'Poland',
    // Vans DC Plant codes
    '1023': 'USA',
    'd010': 'Czech Republic',
    'vd10': 'UAE',
    'd00028': 'UAE',
    // Vans DC Plant name patterns
    'south ontario dc': 'Canada',
    'canada brampton dc': 'Canada',
    'vf prague dc cz': 'Czech Republic',
    'vf northern europe': 'UK',
    'vf northern europe(uk)': 'UK',
    'sun and sand sports': 'UAE',
    'sun and sand sports llc': 'UAE',
};

const BRAND_SUPPLIER_MAP: Record<string, string> = {
    col: "",
    columbia: "",
    tnf: "PT. UWU JUMP INDONESIA",
    "the north face": "PT. UWU JUMP INDONESIA",
    arcteryx: "PT. UWU JUMP INDONESIA",
    "arc'teryx": "PT. UWU JUMP INDONESIA",
    "fox racing": "PT. UWU JUMP INDONESIA",
    "511 tactical": "PT. UWU JUMP INDONESIA",
    "haglofs": "Hangzhou U-Jump",
    "obermeyer": "Hangzhou U-Jump Arts and Crafts",
    "on running": "PT. UWU JUMP INDONESIA",
    "on ag": "PT. UWU JUMP INDONESIA",
    "66 degrees north": "PT. UWU JUMP INDONESIA",
    "peak performance": "PT. UWU JUMP INDONESIA",
    "prana": "PT. UWU JUMP INDONESIA",
    "burton": "PT. UWU JUMP INDONESIA",
    "cotopaxi": "PT. UWU JUMP INDONESIA",
    "hunter": "PT. UWU JUMP INDONESIA",
    "vuori": "PT. UWU JUMP INDONESIA",
    "helly hansen": "PT. UWU JUMP INDONESIA",
    hh: "PT. UWU JUMP INDONESIA",
    "jack wolfskin": "PT. UWU JUMP INDONESIA",
    "ll bean": "PT. UWU JUMP INDONESIA",
    "l.l.bean": "PT. UWU JUMP INDONESIA",
    marmot: "PT. UWU JUMP INDONESIA",
    // New brands
    "dynafit": "Hangzhou U-Jump Arts and Crafts",
    "travis mathew": "PT. UWU JUMP INDONESIA",
    "vans": "PT. UWU JUMP INDONESIA",
    "rossignol": "PT. UWU JUMP INDONESIA",
    "roscoe": "PT. UWU JUMP INDONESIA",
    "mammut": "PT. UWU JUMP INDONESIA",
};

const BRAND_CUSTOMER_MAP: Record<string, string> = {
    col: "Columbia",
    columbia: "Columbia",
    tnf: "The North Face In-Line",
    "the north face": "The North Face In-Line",
    "peak performance": "Peak Performance",
    prana: "Prana",
    arcteryx: "Arcteryx",
    "arc'teryx": "Arcteryx",
    "511 tactical": "511 Tactical",
    evo: "Evo",
    "haglofs": "Haglofs",
    "obermeyer": "Obermeyer",
    "on running": "On AG",
    "on ag": "On AG",
    "66 degrees north": "66 Degrees North",
    "burton": "Burton",
    "cotopaxi": "Cotopaxi",
    "fox racing": "Fox Racing",
    "vuori": "Vuori",
    "helly hansen": "Helly Hansen",
    hh: "Helly Hansen",
    "helly hansen distributie b.v.": "Helly Hansen",
    "helly hansen aus - toll prestons": "Helly Hansen",
    "mainfreight / helly hansen nz": "Helly Hansen",
    "utendor spa": "Helly Hansen",
    "helly hansen (u.s.) inc.": "Helly Hansen",
    "helly hansen smu": "Helly Hansen",
    "jack wolfskin": "Jack Wolfskin",
    "ll bean": "LL Bean",
    "l.l.bean": "LL Bean",
    marmot: "Marmot",
    // New brands
    "dynafit": "Dynafit",
    "travis mathew": "Travis Mathew",
    "vans": "Vans",
    "rossignol": "Rossignol",
    "south ontario dc": "Vans",
    "canada brampton dc": "Vans",
    "vf prague dc cz": "Vans",
    "vf northern europe": "Vans",
    "vf northern europe(uk)": "Vans",
    "sun and sand sports": "Vans",
    "sun and sand sports llc": "Vans",
    "roscoe": "Roscoe",
    "mammut": "Mammut",
};

const TNF_CUSTOMER_SUBTYPE_MAP: Record<string, string> = {
    "the north face in-line": "The North Face In-Line",
    "the north face inline": "The North Face In-Line",
    "the north face rto": "The North Face RTO",
    "the north face smu": "The North Face SMU",
    "tnf in-line": "The North Face In-Line",
    "tnf inline": "The North Face In-Line",
    "tnf rto": "The North Face RTO",
    "tnf smu": "The North Face SMU",
};

const TRANSPORT_MAP: Record<string, string> = {
    "ocean": "Sea",
    "ocean freight (collect)": "Sea",
    "ocean freight collect": "Sea",
    "sea": "Sea",
    "vessel": "Sea",
    "sea freight": "Sea",
    "seafreight": "Sea",
    "s1 - seafreight": "Sea",
    "s1": "Sea",
    "v": "Sea",
    "air": "Air",
    "air freight": "Air",
    "airfreight": "Air",
    "a1 - airfreight": "Air",
    "a1": "Air",
    "a2 - airfreight": "Air",
    "a2": "Air",
    "courier": "Courier",
    "dhl": "Courier",
    "fedex": "Courier",
    "ups": "Courier",
    "private parcel": "Courier",
    "private parcel service": "Courier",
    "parcel": "Courier",
    "international distributor": "Sea",
    "maersk ocean": "Sea",
    "maersk": "Sea",
    "hapag-lloyd": "Sea",
    "hapag lloyd": "Sea",
    "msc": "Sea",
    "cma cgm": "Sea",
    "evergreen": "Sea",
    "cosco": "Sea",
    "yang ming": "Sea",
    "one": "Sea",
    "sos - hunter sos": "Sea",
    "fb - hunter - fob warehouse": "Sea",
    "sms - sample warehouse": "Sea",
    "dte - davies turner e-com warehouse": "Sea",
    "hm - hammer gmbh & co. kg": "Sea",
    "hmcd - hammer cross dock": "Sea",
    // EVO / Roscoe
    "exw": "Sea",
};

const VALID_TRANSPORT_VALUES = new Set(["Sea", "Air", "Courier"]);

const COUNTRY_NAME_MAP: Record<string, string> = {
    AE: "UAE",
    AR: "Argentina",
    AT: "Austria",
    AU: "Australia",
    BR: "Brazil",
    CA: "Canada",
    CH: "Switzerland",
    CL: "Chile",
    CN: "China",
    CZ: "Czech Republic",
    DE: "Germany",
    DK: "Denmark",
    EC: "Ecuador",
    ES: "Spain",
    FR: "France",
    GB: "UK",
    GR: "Greece",
    HK: "Hong Kong",
    HR: "Croatia",
    HU: "Hungary",
    ID: "Indonesia",
    IL: "Israel",
    IN: "India",
    IT: "Italy",
    JP: "Japan",
    KR: "Korea",
    MN: "Mongolia",
    NP: "Nepal",
    MT: "Malta",
    MX: "Mexico",
    MY: "Malaysia",
    "NEW ZEALAND": "New Zealand",
    PA: "Panama",
    PE: "Peru",
    PH: "Philippines",
    PL: "Poland",
    RS: "Serbia",
    RU: "Russia",
    TH: "Thailand",
    TR: "Turkey",
    TW: "Taiwan",
    UK: "UK",
    US: "USA",
    "US WHOLESALE 3PL": "USA",
    "US RETAIL 3PL": "USA",
    "US ECOMM": "USA",
    "UNITED KINGDOM": "UK",
    "UNITED ARAB EMIRATES": "UAE",
    "UNITED STATES": "USA",
    UY: "Uruguay",
    VN: "Vietnam",
    ZA: "South Africa",
    "500025": "Korea",
    "SWEDEN": "Sweden",
    "KOREA": "Korea",
    "JAPAN": "Japan",
    "HONG KONG": "Hong Kong",
    "GERMANY": "Germany",
    "FRANCE": "France",
    "ITALY": "Italy",
    "SPAIN": "Spain",
    "NETHERLANDS": "Netherlands",
    "BELGIUM": "Belgium",
    "SWITZERLAND": "Switzerland",
    "AUSTRIA": "Austria",
    "DENMARK": "Denmark",
    "NORWAY": "Norway",
    "FINLAND": "Finland",
    "POLAND": "Poland",
    "CZECH REPUBLIC": "Czech Republic",
    "AUSTRALIA": "Australia",
    "CANADA": "Canada",
    "CHINA": "China",
    "INDIA": "India",
    "INDONESIA": "Indonesia",
    "MALAYSIA": "Malaysia",
    "THAILAND": "Thailand",
    "VIETNAM": "Vietnam",
    "TAIWAN": "Taiwan",
    "SINGAPORE": "Singapore",
    "CZECHIA": "Czech Republic",
    "GREAT BRITAIN": "UK",
    "TBC": "",
    // Roscoe destination codes (short — only EU since CA/US/JP already exist as ISO codes)
    "EU": "EU",
};

interface KeyUsers {
    k1: string; k2: string; k3: string;
    k4: string; k5: string; k6: string;
    k7: string; k8: string;
}

const BRAND_KEYUSER_MAP: Record<string, KeyUsers> = {
    tnf: { k1: "Ron", k2: "Maricar", k3: "", k4: "Ron", k5: "Elaine Sanchez", k6: "", k7: "", k8: "" },
    "the north face": { k1: "Ron", k2: "Maricar", k3: "", k4: "Ron", k5: "Elaine Sanchez", k6: "", k7: "", k8: "" },
    col:      { k1: "", k2: "", k3: "", k4: "", k5: "", k6: "", k7: "", k8: "" },
    columbia: { k1: "", k2: "", k3: "", k4: "", k5: "", k6: "", k7: "", k8: "" },
    arcteryx: { k1: "", k2: "", k3: "", k4: "", k5: "", k6: "", k7: "", k8: "" },
    "arc'teryx": { k1: "", k2: "", k3: "", k4: "", k5: "", k6: "", k7: "", k8: "" },
    rossignol: { k1: "Via", k2: "April Joy", k3: "", k4: "Via", k5: "Elaine Sanchez", k6: "", k7: "", k8: "" },
    "fox racing": { k1: "Ron", k2: "Maricar", k3: "", k4: "Ron", k5: "Pam", k6: "", k7: "", k8: "" },
    "511 tactical": { k1: "Shania", k2: "Joy", k3: "", k4: "Ron", k5: "Jay", k6: "", k7: "", k8: "" },
    evo: { k1: "Shania", k2: "Mariane", k3: "", k4: "Ron", k5: "Edbert", k6: "", k7: "", k8: "" },
    haglofs: { k1: "Shania", k2: "Mariane", k3: "", k4: "Ron", k5: "Edbert", k6: "", k7: "", k8: "" },
    hh: { k1: "Angelah", k2: "Mariane", k3: "", k4: "Angelah", k5: "Jenica", k6: "", k7: "", k8: "" },
    "helly hansen": { k1: "Angelah", k2: "Mariane", k3: "", k4: "Angelah", k5: "Jenica", k6: "", k7: "", k8: "" },
    prana: { k1: "Jessie", k2: "Maricon Alvarez", k3: "", k4: "Deaunne", k5: "Elaine Sanchez", k6: "", k7: "", k8: "" },
    "jack wolfskin": { k1: "Via", k2: "Mary", k3: "", k4: "Via", k5: "Elaine Sanchez", k6: "", k7: "", k8: "" },
    dynafit: { k1: "Patrick", k2: "Sarah Jane", k3: "", k4: "Patrick", k5: "Edbert Suan", k6: "", k7: "", k8: "" },
    vuori: { k1: "Patrick", k2: "Mary", k3: "", k4: "Patrick", k5: "Elaine Sanchez", k6: "", k7: "", k8: "" },
    "ll bean": { k1: "", k2: "", k3: "", k4: "", k5: "", k6: "", k7: "", k8: "" },
    "l.l.bean": { k1: "", k2: "", k3: "", k4: "", k5: "", k6: "", k7: "", k8: "" },
    marmot: { k1: "", k2: "", k3: "", k4: "", k5: "", k6: "", k7: "", k8: "" },
};

const DEFAULT_KEYUSERS: KeyUsers = { k1: "", k2: "", k3: "", k4: "", k5: "", k6: "", k7: "", k8: "" };

const FACTORY_CODE_MAP: Record<string, string> = {
    '508582':   'PT. UWU JUMP INDONESIA',
    '1002436':  'PT. UWU JUMP INDONESIA',
    '8668:puj': 'PT. UWU JUMP INDONESIA',
    'mad001':   'PT. UWU JUMP INDONESIA',
};

const BRAND_ORDERS_TEMPLATE_MAP: Record<string, string> = {
    tnf:              "Major Brand Bulk",
    "the north face": "Major Brand Bulk",
    col:              "BULK",
    columbia:         "BULK",
    arcteryx:         "BULK",
    "arc'teryx":      "BULK",
    rossignol:        "Major Brand Bulk",
    hh:               "Major Brand Bulk",
    "helly hansen":  "Major Brand Bulk",
    "jack wolfskin":  "Major Brand Bulk",
    dynafit:          "SMS PO Header",
    vuori:            "Major Brand Bulk",
    evo:              "BULK",
    "511 tactical":   "BULK",
    haglofs:          "BULK",
    "fox racing":     "BULK",
    "ll bean":        "Major Brand Bulk",
    "l.l.bean":       "Major Brand Bulk",
    marmot:           "Major Brand Bulk",
    prana:            "Major Brand Bulk",
};

const BRAND_LINES_TEMPLATE_MAP: Record<string, string> = {
    tnf:              "FOB Bulk EDI PO (New)",
    "the north face": "FOB Bulk EDI PO (New)",
    col:              "BULK",
    columbia:         "BULK",
    arcteryx:         "BULK",
    "arc'teryx":      "BULK",
    rossignol:        "FOB Bulk EDI PO (New)",
    hh:               "FOB Bulk EDI PO (New)",
    "helly hansen":  "FOB Bulk EDI PO (New)",
    "jack wolfskin":  "FOB Bulk EDI PO (New)",
    dynafit:          "SMS Non EDI (New)",
    vuori:            "FOB Bulk Non EDI PO (New)",
    evo:              "BULK",
    "511 tactical":   "BULK",
    haglofs:          "BULK",
    "fox racing":     "BULK",
    "ll bean":        "FOB Bulk EDI PO (New)",
    "l.l.bean":       "FOB Bulk EDI PO (New)",
    marmot:           "FOB Bulk EDI PO (New)",
    prana:            "FOB Bulk Non EDI PO (New)",
};

export class ExcelEngine {
    private errors: ValidationError[] = [];
    private runId: string | null = null;
    private userId: string | null = null;

    constructor(runId?: string, userId?: string) {
        this.runId = runId || null;
        this.userId = userId || null;
    }

    private detectHeaderRow(worksheet: ExcelJS.Worksheet): number {
        const KNOWN_HEADERS = new Set([
            'erp ind', 'brand', 'po #', 'pono', 'purchase order',
            'purchaseorder', 'lineitem', 'productrange', 'company code', 'vendor code',
            'material style', 'jde style', 'doc type', 'orig ex fac', 'trans cond',
            'ordered qty', 'buy date', 'color', 'season',
            'tracking number', 'article', 'business unit description',
            'requested qty', 'ex-factory', 'vendor confirmed crd', 'transport mode',
            'qty', 'quantity', 'size', 'colour',
            'product name', 'buyer style number', 'buyer style name', 'customer name', 'factory',
            'warehouse name', 'item',
            // Rossignol / bulk buy layouts with title rows above the real header row
            'destination', 'product code', 'sku', 'shipping date', 'tot qty', 'm88 ref', 'color name', 'size name',
        ].map(h => normalizeHeaderText(h)));
        const fallbackAliases = this.getFallbackColumnAliases();
        let bestRow = 1;
        let bestScore = -1;
        for (let i = 1; i <= Math.min(80, worksheet.rowCount); i++) {
            const row = worksheet.getRow(i);
            let score = 0;
            row.eachCell(cell => {
                const raw = cell.value?.toString().trim() || '';
                if (!raw) return;
                const key = normalizeHeaderText(raw);
                if (KNOWN_HEADERS.has(key)) score += 2;
                if (fallbackAliases[key]) score += 3;
                if (this.looksLikeSizeHeader(raw)) score += 1;
            });
            if (score > bestScore) { bestScore = score; bestRow = i; }
            if (score >= 12) break;
        }
        return bestRow;
    }

    private getFallbackColumnAliases(): Record<string, string> {
        return FALLBACK_COLUMN_ALIASES;
    }

    private getProductSheetAliases(): Record<string, string> {
        return {
            'color name': 'colourName', 'colour name': 'colourName', 'color': 'colour', 'colour': 'colour',
            'factory': 'factory', 'vendor code': 'factory', 'vendorcode': 'factory',
            'cost': 'cost', 'customer name': 'customerName', 'customer': 'customerName', 'cust': 'customerName',
            'product name': 'productName', 'style name': 'productName', 'ng style name': 'productName', 'product': 'productName',
            'name': 'productExternalRef',
            'style': 'buyerStyleNumber',
            'buyer style number': 'buyerStyleNumber', 'buyer style no': 'buyerStyleNumber',
            'buyer style #': 'buyerStyleNumber', 'buyer style': 'buyerStyleNumber',
            'style nr': 'buyerStyleNumber', 'style number': 'buyerStyleNumber',
            'season': 'season',
            'crd': 'crd',
            'ex. factory': 'exFactory',
            'ex factory': 'exFactory',
            'destination name': 'destinationName',
            'po number': 'poNumber',
        };
    }

    private detectProductSheet(worksheet: ExcelJS.Worksheet): { isProductSheet: boolean; headerRow: number } {
        const headerRow = this.detectHeaderRow(worksheet);
        const header = worksheet.getRow(headerRow);
        const aliases = this.getProductSheetAliases();
        const productHeaders = new Set(Object.keys(aliases).map(h => normalizeHeaderText(h)));
        const buyHeaders = new Set(['po #', 'po', 'pono', 'purchase order', 'purchaseorder', 'lineitem', 'quantity', 'qty', 'size', 'season', 'brand', 'productrange'].map(h => normalizeHeaderText(h)));
        const strongBuyHeaders = new Set([
            'purchase order type',
            'requested delivery date',
            'm3 delivery method description',
            'agreement used head',
            'po company name',
            'supplier number',
            'supply planning team owner',
            'purchase price',
            'purchase price (m3)',
            'stylecolor',
            'qty jan buy size-split',
            'bp no',
            'vendor confirmed etd',
            'etd',
            'po number',
            'remark',
            'surcharges',
        ].map(h => normalizeHeaderText(h)));
        let productScore = 0;
        let buyScore = 0;
        let strongBuyScore = 0;
        const headerVals = new Set<string>();
        header.eachCell(cell => {
            const val = normalizeHeaderText(cell.value?.toString() || '');
            if (val) headerVals.add(val);
            if (productHeaders.has(val)) productScore++;
            if (buyHeaders.has(val)) buyScore++;
            if (strongBuyHeaders.has(val)) strongBuyScore++;
        });
        const looksLikePeakPerformanceBuySheet =
            headerVals.has('article code [sap]')
            && headerVals.has('model code [sap]')
            && headerVals.has('primary color peak pdm code')
            && headerVals.has('article full colors')
            && headerVals.has('production supplier name')
            && (headerVals.has('final po qty') || headerVals.has('buy 1 agreed qty') || headerVals.has('buy 1 - tracking no.'));
        const looksLikeVansBuySheet =
            headerVals.has('purchase order #')
            && headerVals.has('dc plant')
            && headerVals.has('mdm material (base)')
            && headerVals.has('color description')
            && headerVals.has('s4h variant')
            && headerVals.has('season')
            && headerVals.has('sap size 1');
        const looksLikeVuoriProductSheet = ['style nr', 'product name', 'color name', 'customer name'].every(h => headerVals.has(h));
        if (looksLikeVuoriProductSheet) productScore += 5;
        if (looksLikeVansBuySheet) return { isProductSheet: false, headerRow };
        if (looksLikePeakPerformanceBuySheet) return { isProductSheet: false, headerRow };
        if (strongBuyScore >= 2 && !looksLikeVuoriProductSheet) return { isProductSheet: false, headerRow };
        return { isProductSheet: productScore >= 3 && (buyScore <= 1 || looksLikeVuoriProductSheet), headerRow };
    }

    private normalizeColourKey(value: string): string {
        const raw = this.stripBrackets(value || '').toLowerCase().trim();
        if (!raw) return '';
        const vansCodeMatch = raw.match(/^(?:[a-z]{2,5})\s*-\s*([a-z0-9]{2,5})\b/);
        if (vansCodeMatch) return vansCodeMatch[1];
        const hhStyleColorMatch = raw.match(/^(\d+)\s*[_-]\s*([a-z0-9]{2,5})\b/);
        if (hhStyleColorMatch) return hhStyleColorMatch[2].replace(/^0+/, '') || hhStyleColorMatch[2];
        if (/^[a-z0-9]{2,6}$/i.test(raw)) return raw;
        if (/^\d+(\.\d+)?$/.test(raw)) {
            const num = Number(raw);
            if (Number.isFinite(num)) return String(Math.trunc(num));
        }
        const dashParts = raw.split('-');
        if (dashParts.length >= 2 && dashParts[0].trim() === 'tnf') return dashParts[1].trim();
        const upperRaw = value.trim().toUpperCase();
        const tnfMaterialMatch = upperRaw.match(/^NF0[A-Z0-9]{5}([A-Z0-9]{2,4})$/);
        if (tnfMaterialMatch) return tnfMaterialMatch[1].toLowerCase();
        const digits = raw.match(/\d+/);
        if (digits && digits[0]) {
            const normalized = digits[0].replace(/^0+/, '');
            return normalized || '0';
        }
        return raw;
    }

    private extractFoxBracketedColour(value: string): string {
        const text = this.stripBrackets(value || '').trim();
        if (!text) return '';
        const matches = [...text.matchAll(/\[([^\]]+)\]/g)];
        if (matches.length > 0) {
            const bracketed = matches[matches.length - 1][1].trim();
            if (bracketed) return bracketed;
        }
        return text;
    }

    private normalizeVuoriColourKey(value: string): string {
        const raw = this.stripBrackets(value || '').toLowerCase().trim();
        if (!raw) return '';
        return raw
            .replace(/^(?:vuo(?:ri)?)\s*-\s*/i, '')
            .replace(/^(?:vuo(?:ri)?)\s+/i, '')
            .replace(/\s{2,}/g, ' ')
            .trim();
    }

    private normalizeJackWolfskinColourKey(value: string): string {
        const raw = this.stripBrackets(value || '').toLowerCase().trim();
        if (!raw) return '';
        if (raw.includes('not set')) return 'not set';
        let normalized = raw
            .replace(/^jw\s*[-_]\s*/i, '')
            .replace(/^jw\s+/i, '')
            .replace(/\b([a-z])\d{4}\b/gi, '')
            .replace(/\b\d{4}\b/g, '')
            .replace(/\bcolor\b/g, '')
            .replace(/\bcolour\b/g, '')
            .replace(/[_-]+/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();
        const digitsOnly = normalized.match(/^\d+$/);
        if (digitsOnly) return digitsOnly[0].replace(/^0+/, '') || digitsOnly[0];
        return normalized;
    }

    private normalizeJackWolfskinStyleKey(value: string): string {
        const raw = this.stripBrackets(value || '').trim();
        if (!raw) return '';
        const compact = raw.replace(/\s+/g, '');
        const prefixMatch = compact.match(/^([A-Za-z0-9]+?)(?:[_-].*|$)/);
        if (prefixMatch?.[1]) return prefixMatch[1];
        return compact;
    }

    private normalizeLlBeanColourKey(value: string): string {
        const raw = this.stripBrackets(value || '').toLowerCase().trim();
        if (!raw) return '';
        const compact = raw.replace(/[^a-z0-9]+/g, '');
        const aliases: Record<string, string> = {
            black: '1black',
            '1black': '1black',
            '2756black': '1black',
            '8281gray': '8281gray',
            '33018nautnvy': 'nautnvy',
            mdnghtblfi: 'midnightblackfairisle',
            midnightblackfairisle: 'midnightblackfairisle',
            dprswdfi: 'deeprosewoodfairisle',
            deeprosewoodfairisle: 'deeprosewoodfairisle',
            dpstgrms: 'deepestgreenmoose',
            deepestgreenmoose: 'deepestgreenmoose',
            dpglcrblbr: 'deepglacierbluebear',
            deepglacierbluebear: 'deepglacierbluebear',
            dpglcrbl: 'deepglacierblue',
            deepglacierblue: 'deepglacierblue',
            charhthr: '2756charhthr',
            charheather: '2756charhthr',
            fadedsage: 'fadedsage',
            icedorchid: 'icedorchid',
            classicnavy: 'llbeanclassicnavy',
            llbeanclassicnavy: 'llbeanclassicnavy',
            llbclassicnavy: 'llbeanclassicnavy',
            classicnavyolivegrey: 'llbclassicnavyolivegrey',
            llbclassicnavyolivegrey: 'llbclassicnavyolivegrey',
            clsscnvyolg: 'llbclassicnavyolivegrey',
            clsscnvyoig: 'llbclassicnavyolivegrey',
            clssnvyoig: 'llbclassicnavyolivegrey',
            clsscnavyolivegrey: 'llbclassicnavyolivegrey',
            electricorng: 'electricorange',
            electricorange: 'electricorange',
            darkcinder: 'darkcinder',
            carbonnavy: 'llbeanclassicnavy',
            oatmealfig: 'oatmealfig',
            shrdrkcndr: 'shrdrkcndr',
            lvndicdkmrb: 'lvndicdkmrb',
            bone: '1267bone',
            lapisteal: 'lapisteal',
            antiquegreen: 'antiquegreen',
            frenchlilac: 'frenchlilac',
            crbnnvypsmo: 'crbnnvypsmo',
            lightgray: '767lightgray',
            darkbronze: '2756darkbronze',
            cream: 'cream',
            red: 'red',
            'dpstgrnsp': 'dpstgrnsp',
        };
        if (aliases[compact]) return aliases[compact];
        const cleaned = compact
            .replace(/^llbean/, '')
            .replace(/^llb/, '')
            .replace(/^\d+/, '');
        return aliases[cleaned] || cleaned;
    }

    private normalizeCotopaxiColourText(value: string): string {
        const raw = this.stripBrackets(value || '').toLowerCase().trim();
        if (!raw) return '';
        const withoutBrand = raw.replace(/^cotopaxi\s*-\s*/i, '');
        const withoutCodes = withoutBrand.replace(/^\d+(?:\s*\/\s*\d+)*(?:\s*-\s*)?/, '');
        return withoutCodes
            .replace(/\//g, ' ')
            .replace(/\band\b/g, ' ')
            .replace(/[^a-z0-9]+/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();
    }

    private normalizeMarmotColourText(value: string): string {
        const raw = this.stripBrackets(value || '').toLowerCase().trim();
        if (!raw) return '';
        const suffix = raw.includes('-') ? raw.split('-').pop()!.trim() : raw;
        const map: Record<string, string> = {
            blk: 'black',
            black: 'black',
            blcknd: 'blackened',
            blcrnt: 'blue currant',
            dskh: 'desert khaki',
            sltsrm: 'sleet storm',
            pnfrst: 'pine forest',
            glcrstrm: 'glacier stream',
            csmsrd: 'cosmos red',
            hzybl: 'hazy blue',
            papy: 'papyrus',
            actny: 'arctic navy',
            da: 'dark azure',
            olvgr: 'olive grove',
            dkaz: 'dark azure',
            dksp: 'dark spice',
            gnbrd: 'gingerbread',
            shgr: 'shale grey',
            stlo: 'steel onyx',
            thhd: 'thunderhead',
            clsthtr: 'coal heather',
            blkplm: 'black plum',
            brhwht: 'birch white',
            dkspcmrd: 'dark spice cardamom red',
            hckrntmrd: 'huckleberry nutmeg red',
            ngtflnvmrd: 'nightfall navy marled',
            gryh: 'grey heather',
            rtkhk: 'rustic khaki',
            nflnv: 'nightfall navy',
        };
        return suffix
            .split(/[\/\s]+/)
            .filter(Boolean)
            .map(token => map[token] || token)
            .join(' ')
            .replace(/\s+/g, ' ')
            .trim();
    }

    private getPeakPerformanceColourCandidates(value: string): string[] {
        const raw = this.stripBrackets(value || '').toLowerCase().trim();
        if (!raw) return [];
        const candidates = new Set<string>();
        candidates.add(raw);
        candidates.add(raw.replace(/\s+/g, ' '));
        candidates.add(raw.replace(/[^a-z0-9]+/g, ''));
        const normalized = this.normalizeColourKey(raw);
        if (normalized) candidates.add(normalized.toLowerCase().trim());
        const suffixMatch = raw.match(/\bpp\s*-\s*(.+)$/i);
        if (suffixMatch?.[1]) {
            const suffix = suffixMatch[1].toLowerCase().trim();
            candidates.add(suffix);
            candidates.add(suffix.replace(/[^a-z0-9]+/g, ''));
            const suffixNormalized = this.normalizeColourKey(suffix);
            if (suffixNormalized) candidates.add(suffixNormalized.toLowerCase().trim());
        }
        return Array.from(candidates).filter(Boolean);
    }

    private extractStyleColourCode(styleKey: string): string {
        const upper = (styleKey || '').trim().toUpperCase();
        const match = upper.match(/([A-Z0-9]{3})$/);
        return match ? match[1].toLowerCase() : '';
    }

    private normalizeStyleKeyCandidates(styleKey: string): string[] {
        const candidates: string[] = [styleKey];
        if (/^NF0/i.test(styleKey)) candidates.push(styleKey.slice(3));
        if (/^NF[^0]/i.test(styleKey)) candidates.push(styleKey.slice(2));
        if (/^V\d{3,}$/i.test(styleKey)) candidates.push(styleKey.slice(1));
        if (/^H\d{3,}$/i.test(styleKey)) candidates.push(styleKey.slice(1));
        if (/^2UF/i.test(styleKey) && styleKey.length > 7) candidates.push(styleKey.slice(0, 7));
        const dynafitMatch = styleKey.match(/(\d{6})$/);
        if (dynafitMatch) {
            candidates.push(dynafitMatch[1]);
            candidates.push(String(parseInt(dynafitMatch[1], 10)));
        }
        return candidates;
    }

    private extractProductSheetMapFromWorkbook(workbook: ExcelJS.Workbook): Record<string, ProductSheetRow[]> {
        const result: Record<string, ProductSheetRow[]> = {};
        const aliases = this.getProductSheetAliases();
        const seenEntries = new Set<string>();
        for (const ws of workbook.worksheets) {
            const { isProductSheet, headerRow } = this.detectProductSheet(ws);
            if (!isProductSheet) continue;
            const header = ws.getRow(headerRow);
            const headerMap: Record<string, number> = {};
            header.eachCell((cell, colNumber) => {
                const key = cell.value?.toString().toLowerCase().trim() || '';
                const mapped = aliases[key];
                if (mapped && !headerMap[mapped]) headerMap[mapped] = colNumber;
            });
            const colourHeaderCol = headerMap['colour'] || headerMap['colourName'];
            if (!colourHeaderCol) continue;
            ws.eachRow((row, rowNumber) => {
                if (rowNumber <= headerRow) return;
                const getRaw = (field: string) => {
                    const col = headerMap[field];
                    if (!col) return undefined;
                    return this.getCellValue(row.getCell(col));
                };
                const colourRaw = (colourHeaderCol ? this.getCellValue(row.getCell(colourHeaderCol)) : undefined)?.toString().trim() || '';
                const colourNameRaw = getRaw('colourName')?.toString().trim() || '';
                const customerNameRaw = getRaw('customerName')?.toString().trim() || '';
                const isVuoriCustomer = customerNameRaw.toLowerCase().includes('vuori');
                const isPeakPerformanceCustomer = customerNameRaw.toLowerCase().includes('peak performance');
                const isMarmotCustomer = customerNameRaw.toLowerCase().includes('marmot') || colourNameRaw.toLowerCase().startsWith('mar-') || colourRaw.toLowerCase().startsWith('mar-');
                const isLlBeanCustomer = customerNameRaw.toLowerCase().includes('ll bean');
                const peakPerformanceColourSource = this.stripBrackets(
                    getRaw('inlineColorName')?.toString().trim() || colourNameRaw || colourRaw
                ).trim();
                const peakPerformanceColourCandidates = this.getPeakPerformanceColourCandidates(peakPerformanceColourSource);
                const peakPerformanceColourKey = peakPerformanceColourCandidates[0]?.toLowerCase().trim() || '';
                const colourKey = isVuoriCustomer
                    ? this.normalizeVuoriColourKey(colourNameRaw || colourRaw)
                    : (isPeakPerformanceCustomer
                        ? peakPerformanceColourKey
                        : (isMarmotCustomer ? this.normalizeMarmotColourText(colourNameRaw || colourRaw) : (isLlBeanCustomer ? this.normalizeLlBeanColourKey(colourRaw) : this.normalizeColourKey(colourRaw))));
                const cotopaxiColourKey = this.normalizeCotopaxiColourText(colourRaw);
                const marmotColourKey = isMarmotCustomer ? this.normalizeMarmotColourText(colourNameRaw || colourRaw) : '';
                const jwsColourKey = (customerNameRaw.toLowerCase().includes('jack wolfskin'))
                    ? this.normalizeJackWolfskinColourKey(colourRaw)
                    : '';
                const llbColourKey = isLlBeanCustomer ? this.normalizeLlBeanColourKey(colourRaw) : '';
                const buyerStyleNumber = getRaw('buyerStyleNumber')?.toString().trim() || '';
                if (!colourKey || !buyerStyleNumber) return;
                const entry: ProductSheetRow = {
                    colour: isPeakPerformanceCustomer ? (peakPerformanceColourSource || colourRaw) : colourRaw,
                    colourName: getRaw('colourName')?.toString().trim() || '',
                    factory: getRaw('factory')?.toString().trim() || '',
                    cost: (() => { const c = getRaw('cost'); return typeof c === 'number' ? c : c?.toString().trim(); })(),
                    customerName: getRaw('customerName')?.toString().trim() || '',
                    productName: getRaw('productName')?.toString().trim() || '',
                    productExternalRef: getRaw('productExternalRef')?.toString().trim() || '',
                    buyerStyleNumber,
                    season: getRaw('season')?.toString().trim() || '',
                    crd: getRaw('crd') as string | Date | undefined,
                    exFactory: getRaw('exFactory') as string | Date | undefined,
                    poNumber: getRaw('poNumber')?.toString().trim() || '',
                    destinationName: getRaw('destinationName')?.toString().trim() || '',
                };
                const lookupKeys = new Map<string, boolean>();
                lookupKeys.set(buyerStyleNumber, true);
                buyerStyleNumber.split('/').forEach(part => { const p = part.trim(); if (p && p !== buyerStyleNumber) lookupKeys.set(p, false); });
                const styleBase = buyerStyleNumber.split(/\s*[\(\-]/)[0].trim();
                if (styleBase && styleBase !== buyerStyleNumber) lookupKeys.set(styleBase, false);
                if (isPeakPerformanceCustomer) {
                    for (const altColour of peakPerformanceColourCandidates) {
                        const altKey = this.normalizeColourKey(altColour) || altColour;
                        if (altKey) lookupKeys.set(`${buyerStyleNumber}|${altKey}`, true);
                    }
                }
                for (const [lk, isExact] of lookupKeys) {
                    const lkKey = lk.includes('|') ? lk : `${lk}|${colourKey}`;
                    const dedupKey = `${lkKey}|${entry.colour}|${entry.factory}|${entry.productName}|${entry.productExternalRef}|${entry.customerName}`;
                    if (seenEntries.has(dedupKey)) continue;
                    seenEntries.add(dedupKey);
                    if (!result[lkKey]) result[lkKey] = [];
                    if (isExact) result[lkKey].unshift(entry); else result[lkKey].push(entry);
                    if (jwsColourKey && jwsColourKey !== colourKey) {
                        const jwsAltKey = `${lk}|${jwsColourKey}`;
                        if (!result[jwsAltKey]) result[jwsAltKey] = [];
                        if (isExact) result[jwsAltKey].unshift(entry); else result[jwsAltKey].push(entry);
                    }
                    if (cotopaxiColourKey && cotopaxiColourKey !== colourKey) {
                        const cotopaxiAltKey = `${lk}|${cotopaxiColourKey}`;
                        if (!result[cotopaxiAltKey]) result[cotopaxiAltKey] = [];
                        if (isExact) result[cotopaxiAltKey].unshift(entry); else result[cotopaxiAltKey].push(entry);
                    }
                    if (marmotColourKey && marmotColourKey !== colourKey) {
                        const marmotAltKey = `${lk}|${marmotColourKey}`;
                        if (!result[marmotAltKey]) result[marmotAltKey] = [];
                        if (isExact) result[marmotAltKey].unshift(entry); else result[marmotAltKey].push(entry);
                    }
                    if (llbColourKey && llbColourKey !== colourKey) {
                        const llbAltKey = `${lk}|${llbColourKey}`;
                        if (!result[llbAltKey]) result[llbAltKey] = [];
                        if (isExact) result[llbAltKey].unshift(entry); else result[llbAltKey].push(entry);
                    }
                    const isCotopaxi = (entry.customerName || '').trim().toLowerCase().includes('cotopaxi');
                    if (isCotopaxi && cotopaxiColourKey && cotopaxiColourKey !== colourKey) {
                        const altKey = `${lk}|${cotopaxiColourKey}`;
                        if (!result[altKey]) result[altKey] = [];
                        if (isExact) result[altKey].unshift(entry); else result[altKey].push(entry);
                    }
                }
            });
        }
        return result;
    }

    private extractInlineProductSheetMapFromBuyWorkbook(workbook: ExcelJS.Workbook): Record<string, ProductSheetRow[]> {
        const result: Record<string, ProductSheetRow[]> = {};
        const seenEntries = new Set<string>();
        const buyAliases: Record<string, string[]> = {
            buyerStyleNumber: ['item#', 'item #', 'buyer item#', 'buyer item #', 'style code', 'style number', 'style nr', 'product', 'sku'],
            colour: ['color code', 'colour', 'color', 'colour code', 'color description', 'style color', 'stylecolor'],
            colourName: ['color name', 'colour name', 'color description', 'colour description'],
            factory: ['updated fty', 'factory', 'vendor name', 'supplier name'],
            productName: ['style name', 'product name'],
            productExternalRef: ['sku', 'fty-sku', 'sku-wh', 'po#', 'po #', 'purchase order'],
            productCustomerRef: ['buyer style number', 'buyer item#', 'buyer item #'],
            season: ['season', 'item market', 'type', 'segmentation'],
            crd: ['updated planned exit date', 'confirmed x-fty', 'delivery date', 'requested delivery date'],
            exFactory: ['confirmed x-fty', 'updated planned exit date', 'delivery date'],
            poNumber: ['po#', 'po #', 'purchase order', 'master po#'],
            destinationName: ['wh', 'whs', 'warehouse name', 'destination name', 'item market'],
        };

        for (const ws of workbook.worksheets) {
            const { isProductSheet } = this.detectProductSheet(ws);
            if (isProductSheet) continue;

            const headerRow = this.detectHeaderRow(ws);
            if (!this.isLikelyBuySheet(ws, headerRow, this.getFallbackColumnAliases())) continue;

            const header = ws.getRow(headerRow);
            const headerMap: Record<string, number> = {};
            header.eachCell((cell, colNumber) => {
                const key = normalizeHeaderText(cell.value?.toString() || '');
                if (!key) return;
                for (const [field, keys] of Object.entries(buyAliases)) {
                    if (!headerMap[field] && keys.includes(key)) {
                        headerMap[field] = colNumber;
                    }
                }
            });

            const getRaw = (row: ExcelJS.Row, field: keyof typeof buyAliases) => {
                const col = headerMap[field as string];
                if (!col) return undefined;
                return this.getCellValue(row.getCell(col));
            };

            ws.eachRow((row, rowNumber) => {
                if (rowNumber <= headerRow) return;

                const buyerStyleNumber = this.stripBrackets(
                    (getRaw(row, 'buyerStyleNumber') || getRaw(row, 'productCustomerRef') || getRaw(row, 'productName') || getRaw(row, 'productExternalRef') || getRaw(row, 'poNumber') || '')
                        .toString()
                ).trim();
                const colourRaw = this.stripBrackets((getRaw(row, 'colour') || getRaw(row, 'colourName') || '')?.toString() || '').trim();
                const colourKey = this.normalizeColourKey(colourRaw);
                if (!buyerStyleNumber || !colourKey) return;

                const entry: ProductSheetRow = {
                    colour: colourRaw,
                    colourName: this.stripBrackets((getRaw(row, 'colourName') || colourRaw || '')?.toString() || '').trim(),
                    factory: this.stripBrackets((getRaw(row, 'factory') || '')?.toString() || '').trim(),
                    cost: (() => {
                        const c = getRaw(row, 'poNumber');
                        return typeof c === 'number' ? c : (c?.toString().trim() || '');
                    })(),
                    customerName: '',
                    productName: this.stripBrackets((getRaw(row, 'productName') || '')?.toString() || '').trim(),
                    productExternalRef: this.stripBrackets((getRaw(row, 'productExternalRef') || '')?.toString() || '').trim(),
                    buyerStyleNumber,
                    season: this.stripBrackets((getRaw(row, 'season') || '')?.toString() || '').trim(),
                    crd: getRaw(row, 'crd') as string | Date | undefined,
                    exFactory: getRaw(row, 'exFactory') as string | Date | undefined,
                    poNumber: this.stripBrackets((getRaw(row, 'poNumber') || '')?.toString() || '').trim(),
                    destinationName: this.stripBrackets((getRaw(row, 'destinationName') || '')?.toString() || '').trim(),
                };

                const lookupKeys = new Set<string>();
                lookupKeys.add(`${buyerStyleNumber}|${colourKey}`);
                const styleBase = buyerStyleNumber.split(/\s*[\(\-]/)[0].trim();
                if (styleBase && styleBase !== buyerStyleNumber) lookupKeys.add(`${styleBase}|${colourKey}`);

                for (const lk of lookupKeys) {
                    const dedupKey = `${lk}|${entry.colour}|${entry.factory}|${entry.productName}|${entry.productExternalRef}|${entry.customerName}`;
                    if (seenEntries.has(dedupKey)) continue;
                    seenEntries.add(dedupKey);
                    if (!result[lk]) result[lk] = [];
                    result[lk].push(entry);
                }
            });
        }

        return result;
    }

    private resolveSupplier(vendorCode: string | undefined, vendorName: string | undefined, brand: string | undefined, category: string | undefined, factoryMap: any[]): string {
        const vCode = this.stripBrackets(vendorCode || '').trim();
        const vName = this.stripBrackets(vendorName || '').trim();
        const b = this.stripBrackets(brand || '').trim();
        const cat = this.stripBrackets(category || '').trim();
        if (vCode && FACTORY_CODE_MAP[vCode]) return FACTORY_CODE_MAP[vCode];
        if (vCode && vCode.length > 2 && !/^\d+$/.test(vCode)) return vCode;
        if (vName && vName.length > 2) {
            if (b.toLowerCase() === 'jack wolfskin' && /pt\s*uwu\s*jump\s*-\s*jw/i.test(vName)) return 'PT. UWU JUMP INDONESIA';
            return vName.replace(/^PT\s+(?!\.)/i, 'PT. ');
        }
        if (b && cat) {
            const match = factoryMap.find((f: any) => f.brand?.toLowerCase() === b.toLowerCase() && f.category?.toLowerCase() === cat.toLowerCase());
            if (match?.product_supplier) return match.product_supplier;
        }
        if (b) {
            const brandMatches = factoryMap.filter((f: any) => f.brand?.toLowerCase() === b.toLowerCase() && f.product_supplier);
            if (brandMatches.length === 1) return brandMatches[0].product_supplier;
        }
        return BRAND_SUPPLIER_MAP[b.toLowerCase()] || 'MISSING_SUPPLIER';
    }

    private resolveCustomer(customerRaw: string | undefined, brand: string | undefined, detectedCustomer: string, mappedCustomerName: string | undefined): string {
        const raw = this.stripBrackets(customerRaw || '').trim();
        const brandClean = this.stripBrackets(brand || '').trim();
        const mapped = this.stripBrackets(mappedCustomerName || '').trim();
        if (raw) { const key = raw.toLowerCase(); if (TNF_CUSTOMER_SUBTYPE_MAP[key]) return TNF_CUSTOMER_SUBTYPE_MAP[key]; }
        if (mapped) {
            const mappedKey = mapped.toLowerCase();
            const hasSubtype = mappedKey.includes('rto') || mappedKey.includes('smu') || mappedKey.includes('in-line') || mappedKey.includes('inline');
            if (!raw || hasSubtype) return mapped;
        }
        if (raw) { const key = raw.toLowerCase(); if (BRAND_CUSTOMER_MAP[key]) return BRAND_CUSTOMER_MAP[key]; return raw; }
        if (brandClean) { const key = brandClean.toLowerCase(); if (BRAND_CUSTOMER_MAP[key]) return BRAND_CUSTOMER_MAP[key]; }
        if (detectedCustomer && detectedCustomer !== 'DEFAULT') return detectedCustomer;
        return brandClean.toUpperCase() || 'COL';
    }

    private detectCustomerSubtype(rawCustomer: string | undefined): string | undefined {
        const text = (rawCustomer || '').toLowerCase();
        if (text.includes('smu')) return 'SMU';
        if (text.includes('rto')) return 'RTO';
        if (text.includes('outlet')) return 'Outlet';
        return undefined;
    }

    private normalizeVansPoSuffix(rawCustomer: string | undefined): string {
        const text = this.stripBrackets(rawCustomer || '').trim();
        const key = text.toLowerCase();
        if (!key) return '';
        if (key.includes('south ontario')) return 'South Ontario';
        if (key.includes('brampton')) return 'Brampton';
        if (key.includes('sun and sand sports')) return 'Sun and Sand Sports';
        if (key.includes('vf prague')) return 'VF Prague DC CZ';
        if (key.includes('vf northern europe')) return 'VF Northern Europe (UK)';
        return text.replace(/\s+dc$/i, '').trim();
    }

    private normalizeVansPlantCode(rawPlant: string | undefined): string {
        const plant = this.stripBrackets(rawPlant || '').trim();
        if (!plant) return '';
        if (plant.toLowerCase() === 'd00028') return 'VD10';
        return plant.toUpperCase();
    }

    private normalizeVansPlantLabel(rawPlantName: string | undefined, rawPlantCode: string | undefined): string {
        const plantName = this.stripBrackets(rawPlantName || '').trim();
        const plantCode = this.stripBrackets(rawPlantCode || '').trim().toUpperCase();
        const labelByCode: Record<string, string> = {
            '1023': 'South Ontario',
            '1004': 'Brampton',
            'D010': 'EMEA',
            'D080': 'EMEA',
            'VD10': 'EMEA',
            'D00028': 'Sun and Sand Sports',
        };
        if (plantName) {
            const directMap: Record<string, string> = {
                'south ontario dc': 'South Ontario',
                'bampton dc': 'Brampton',
                'brampton dc': 'Brampton',
                'vf northern europe': 'VF Northern Europe',
                'vf northern europe (uk)': 'VF Northern Europe',
                'vf prague dc cz': 'VF Prague',
                'sun and sand sports': 'Sun and Sand Sports',
                'sun and sand sports llc': 'Sun and Sand Sports',
                'emea': 'EMEA',
            };
            const key = plantName.toLowerCase();
            if (directMap[key]) return directMap[key];
            return plantName.replace(/\s+dc$/i, '').trim();
        }
        return labelByCode[plantCode] || plantCode || '';
    }

    private resolveHhDestinationCountry(companyName: string | undefined, shipTo: string | undefined, manualDestination: string | undefined, plantDerivedCountry: string | undefined): string {
        const resolvedCompany = this.resolveHhCompanyName(companyName);
        const raw = this.stripBrackets(resolvedCompany || '').toLowerCase().trim();
        const shipToKey = this.stripBrackets(shipTo || '').toLowerCase().trim();
        const manualKey = this.stripBrackets(manualDestination || '').toLowerCase().trim();
        const plantKey = this.stripBrackets(plantDerivedCountry || '').toLowerCase().trim();
        const source = raw || shipToKey || manualKey || plantKey;
        if (!source) return '';
        const exactMap: Record<string, string> = {
            'helly hansen distributie b.v.': 'Netherlands',
            'helly hansen aus - toll prestons': 'Australia',
            'mainfreight / helly hansen nz': 'New Zealand',
            'utendor spa': 'Chile',
            'helly hansen (u.s.) inc.': 'USA',
            'helly hansen smu': 'USA',
        };
        if (exactMap[source]) return exactMap[source];
        if (source.includes('new zealand') || source.includes(' nz') || source.endsWith(' nz') || source.includes('mainfreight')) return 'New Zealand';
        if (source.includes('australia') || source.includes(' aus ') || source.startsWith('aus ') || source.includes('prestons') || source.includes('sydney')) return 'Australia';
        if (source.includes('netherlands') || source.includes('distributie') || source.includes(' b.v.') || source.includes(' b v') || source.includes('houten') || source.includes('utrecht')) return 'Netherlands';
        if (source.includes('italy') || source.includes('utendor') || source.includes('udor') || source.includes('spa')) return 'Italy';
        if (source.includes('usa') || source.includes('u.s.') || source.includes(' united states') || source.includes('us ')) return 'USA';
        if (source.includes('uk') || source.includes('united kingdom') || source.includes(' great britain')) return 'UK';
        if (source.includes('canada') || source.includes(' brampton')) return 'Canada';
        return '';
    }

    private resolveHhCompanyName(rawCompanyName: string | undefined): string {
        const source = this.stripBrackets(rawCompanyName || '').toLowerCase().trim();
        if (!source) return '';
        const exactMap: Record<string, string> = {
            'whs 126': 'HELLY HANSEN AUS - TOLL Prestons',
            'whs 920': 'Helly Hansen (U.S.) Inc.',
            'whs 120': 'Helly Hansen Distributie B.V.',
        };
        return exactMap[source] || this.stripBrackets(rawCompanyName || '').trim();
    }

    private normalizeHhPlantCode(rawPlantCode: string | undefined): string {
        const text = this.stripBrackets(rawPlantCode || '').trim();
        if (!text) return '';
        const whsMatch = text.match(/^WHS\s*(.+)$/i);
        return whsMatch ? whsMatch[1].trim() : text;
    }

    private normalizeSizeName(rawSize: string | undefined, brand: string | undefined): string {
        const size = this.stripBrackets(rawSize || '').trim();
        if (!size) return 'One Size';
        if (size.toLowerCase() === 'os') return 'One Size';
        if (size.toLowerCase() === 'o/s') return 'One Size';
        if (size.toLowerCase() === 'ons') return 'One Size';
        if (/^one\s*size$/i.test(size) || /^onesize$/i.test(size)) return 'One Size';
        if ((brand || '').trim().toLowerCase() === 'on ag') return 'One Size';
        if ((brand || '').trim().toLowerCase() === 'vans' && /^one\s*size$/i.test(size)) return 'One Size';
        return size;
    }

    private normalizeStatus(rawStatus: string | undefined, brand: string | undefined): string {
        const status = this.stripBrackets(rawStatus || '').trim();
        const brandKey = (brand || '').trim().toLowerCase();
        if ((brandKey === 'hh' || brandKey === 'helly hansen') && (/^20$/.test(status) || /^confirmed$/i.test(status))) return 'Confirmed';
        if (brandKey === 'vans' && (!status || status.toLowerCase() === 'converted')) return 'Confirmed';
        return status || 'Confirmed';
    }

    private resolveKeyUsers(brand: string | undefined, manualK1: string | undefined, manualK2: string | undefined, manualK3: string | undefined, manualK4: string | undefined, manualK5: string | undefined, providedK1: string | undefined, providedK2: string | undefined, providedK4: string | undefined, providedK5: string | undefined, mloRow: any): KeyUsers {
        const hasManual = !!(manualK1 || manualK2 || manualK3 || manualK4 || manualK5);
        if (hasManual) return { k1: manualK1 || '', k2: manualK2 || '', k3: manualK3 || '', k4: manualK4 || '', k5: manualK5 || '', k6: '', k7: '', k8: '' };
        if (providedK1 || providedK2) return { k1: providedK1 || '', k2: providedK2 || '', k3: '', k4: providedK4 || '', k5: providedK5 || '', k6: '', k7: '', k8: '' };
        const key = (brand || '').trim().toLowerCase();
        if (key === 'dynafit') return { k1: 'Patrick', k2: 'Sarah Jane', k3: '', k4: 'Patrick', k5: 'Edbert Suan', k6: '', k7: '', k8: '' };
        if (key === 'll bean') return { k1: 'Angelah', k2: 'MJ', k3: '', k4: 'Angelah', k5: 'Pamela', k6: '', k7: '', k8: '' };
        if (key === 'fox racing') return { k1: 'Ron', k2: 'Maricar', k3: '', k4: 'Ron', k5: 'Pam', k6: '', k7: '', k8: '' };
        if (key === '511 tactical') return { k1: 'Shania', k2: 'Joy', k3: '', k4: 'Ron', k5: 'Jay', k6: '', k7: '', k8: '' };
        if (key === 'evo') return { k1: 'Shania', k2: 'Mariane', k3: '', k4: 'Ron', k5: 'Edbert', k6: '', k7: '', k8: '' };
        if (key === 'haglofs') return { k1: 'Shania', k2: 'Mariane', k3: '', k4: 'Ron', k5: 'Edbert', k6: '', k7: '', k8: '' };
        if (mloRow) return { k1: mloRow.keyuser1 || '', k2: mloRow.keyuser2 || '', k3: '', k4: mloRow.keyuser4 || '', k5: mloRow.keyuser5 || '', k6: '', k7: '', k8: '' };
        return { ...(BRAND_KEYUSER_MAP[key] || DEFAULT_KEYUSERS) };
    }

    private resolveOrdersTemplate(brand: string | undefined, rawTemplate: string): string {
        const key = (brand || '').trim().toLowerCase();
        if (BRAND_ORDERS_TEMPLATE_MAP[key]) return BRAND_ORDERS_TEMPLATE_MAP[key];
        return this.normalizeTemplate(rawTemplate);
    }

    private resolveLinesTemplate(brand: string | undefined, rawTemplate: string): string {
        const key = (brand || '').trim().toLowerCase();
        if (BRAND_LINES_TEMPLATE_MAP[key]) return BRAND_LINES_TEMPLATE_MAP[key];
        return this.normalizeTemplate(rawTemplate);
    }

    private looksLikeSizeHeader(headerText: string): boolean {
        const normalized = headerText.trim().toLowerCase();
        const directMatches = new Set(['size', 'size name', 'sizename', 'productsize', 'product size', 'size #', 'size#', 'size code', 'size_name', 'size-name']);
        if (directMatches.has(normalized)) return true;
        return isLikelyPivotSizeHeader(headerText);
    }

    private shouldSilentlyIgnoreHeader(headerText: string): boolean {
        const normalized = headerText.trim().toLowerCase();
        if (/^po\d{4,}$/i.test(headerText.trim())) return true;
        const exactIgnore = new Set([
            'unit total', 'confirmed unit total', 'vendor comments', 'vendor confirmed',
            'csc/lo comments', 'lo reviewed', 'lo rejected', 'csc confirmed', 'csc rejected',
            'last collab status date', 'hashcode', 'linehashcode', 'mainitem_id', 'activity_info',
            'modifyrivision', 'rawinfo', 'writablecells', 'rowsuffix',
            'vendor price chg 1', 'price chg type 1', 'vendor price chg 2', 'price chg type 2',
            'vendor price chg 3', 'price chg type 3', 'net price chg', 'absolute price chg',
            'line #s 2', 'line #s', 'lineitem', 'purchaseprice', 'sellingprice',
            'supplierprofile', 'closeddate', 'comments', 'currency', 'archivedate',
            'productexternalref', 'productcustomerref', 'purchaseuom', 'sellinguom',
            'paymentterm', 'defaultdeliverydate', 'productsupplierext',
            'keyuser1', 'keyuser2', 'keyuser3', 'keyuser4', 'keyuser5',
            'keyuser6', 'keyuser7', 'keyuser8',
            'department', 'customattribute1', 'customattribute2', 'customattribute3',
            'lineratio', 'colourext', 'customerext', 'departmentext',
            'customattribute1ext', 'customattribute2ext', 'customattribute3ext',
            'file date', 'sku', 'sku description', 'model description', 'gsc type',
            'product group description', 'product line description', 'planning category',
            'transit vendor destination', 'official buy', 'storage location', 'stock segment',
            'erp ind', 'company code', 'ab number', 'gtn issue date', 'sku status', 'slo', 'plo',
            'priority flag', 'lb', 'tooling code', 'vas', 'capacity type',
            'log#', 'analysis dc', 'year', 'season indicator', 'buy month',
            'material description short', 'grid value',
            'original qty', 'revised qty (0 if cancel, new qty if top up or reduce)',
            'final qty per material per factory',
            'final qty per material per factory(combine plus + regular qty)',
            'final qty per material per factory (regular vs regular + plus)',
            'purchase requisition', 'item of requisition', 'sales document', 'planner name', 'team',
            'total leadtime1', 'dim 1', 'dim 2', 'buy ready file lookup', 'buy ready feedback from vendor',
            'final factory name', 'final vendor', 'final factory coo',
            'change in fcty code?', 'new fcty code (if col cd is "yes")', 'reason for change, if any',
            'ori fob', 'correction to ori fob', 'reason for correction, if any', 'final original fob',
            'production upcharges usd$', 'material upcharges usd$', 'total surcharge usd$',
            'revised final fob', 'ori rmb fob', 'correction to ori rmb fob',
            'reason for correction, if any (#1)', 'final ori rmb',
            'production upcharges rmb$', 'material upcharges rmb$', 'total surcharge rmb$',
            'revised rmb fob', 'production upcharge %', 'material upcharge %', 'total upcharge %',
            'production surcharge confirmation status', 'material surcharge confirmation status',
            'comment category (core data check)', 'comment (core data check)',
            'production minimum order qty / absolute moq', 'moq related comments',
            'matl related comments', 'additional', 'vendor remarks', 'planner comments', 'decision', 'pped',
            'sbu - apparel or acc/equip', 'stock category',
            'order type', 'deliv date(from/to)', 'smu', "planner's comment",
            'eu old sku', 'lt2', 'calculated indc', 'final qty for pivot',
            'region grouping', 'transportation mode description', 'eu collection',
            'm88 ped', 'regular material', 'plus material', 'ship to', 'unit price',
            'xs', 's', 'm', 'l', 'xl',
            'grand total', 'confirmed unit price', 'total confirmed unit price',
            'moq', 'moq upcharge(%)', 'moq upcharge (%)', 'std',
            'final delivery date', 'customer request date', 'forecast', 'total bulk', 'variance',
        ]);
        if (exactIgnore.has(normalized)) return true;
        const ignorePrefixes = ['findfield_', 'udf-inspection', 'udf-report', 'udf-inspector', 'udf-approval', 'udf-submitted'];
        return ignorePrefixes.some(p => normalized.startsWith(p));
    }

    private inferCategoryFromFactoryMap(brand: string | undefined, factoryMap: any[]): string | undefined {
        if (!brand) return undefined;
        const matches = factoryMap.filter((f: any) => f.brand?.toLowerCase() === brand.toLowerCase()).map((f: any) => f.category).filter(Boolean);
        const unique = Array.from(new Set(matches));
        return unique.length === 1 ? unique[0] : undefined;
    }

    private formatProductRange(season: string): string {
        const normalized = this.stripBrackets(season || '').trim();
        const fhMatch = normalized.match(/^FH(\d{2})$/i);
        if (fhMatch) return `FH:20${fhMatch[1]}`;
        const faMatch = normalized.match(/^(\d{2})FA$/i);
        if (faMatch) return `FH:20${faMatch[1]}`;
        const m = normalized.match(/^([FS])(?:W|S)?(\d{2})$/i);
        if (m) return `${m[1].toUpperCase()}H:20${m[2]}`;
        const altMatch = normalized.match(/^(AW|FW|AH)(\d{2})$/i);
        if (altMatch) return `FH:20${altMatch[2]}`;
        const springMatch = normalized.match(/^(SS|SP)(\d{2})$/i);
        if (springMatch) return `SH:20${springMatch[2]}`;
        const winterTextMatch = normalized.match(/^WINTER\s+20(\d{2})$/i);
        if (winterTextMatch) return `FH:20${winterTextMatch[1]}`;
        const summerTextMatch = normalized.match(/^(SUMMER|SPRING)\s+20(\d{2})$/i);
        if (summerTextMatch) return `SH:20${summerTextMatch[2]}`;
        if (normalized) return normalized;
        return 'FH:2026';
    }

    private compactProductRange(raw: string): string {
        const normalized = this.stripBrackets(raw || '').trim();
        const fullYearMatch = normalized.match(/^([FS]H):?20(\d{2})$/i);
        if (fullYearMatch) return `${fullYearMatch[1].toUpperCase()}${fullYearMatch[2]}`;
        const shortMatch = normalized.match(/^([FS][HWS])[: ]?(\d{2,4})$/i);
        if (shortMatch) return `${shortMatch[1].toUpperCase()}${shortMatch[2].slice(-2)}`;
        return normalized.replace(/[^A-Za-z0-9]/g, '');
    }

    private resolveRossignolDestinationSuffix(raw: string): string {
        const key = this.stripBrackets(raw || '').trim().toUpperCase();
        if (!key) return '';
        const map: Record<string, string> = {
            CA: 'Canada',
            CANADA: 'Canada',
            EU: 'Europe',
            EUROPE: 'Europe',
            FRANCE: 'Europe',
            JP: 'Japan',
            JAPAN: 'Japan',
            US: 'USA',
            USA: 'USA',
            'UNITED STATES': 'USA',
        };
        return map[key] || this.stripBrackets(raw || '').trim();
    }

    private resolveOnAgCountryToken(raw: string): string {
        const normalized = this.normalizeTransportLocation(raw || '');
        return this.stripBrackets(normalized || raw || '').trim().toUpperCase();
    }

    private normalizeOnAgTransportLocation(raw: string): string {
        const normalized = this.normalizeTransportLocation(raw || '');
        const cleaned = this.stripBrackets(normalized || raw || '').trim();
        if (!cleaned) return '';
        if (/^[A-Z]{2,4}$/.test(cleaned)) return cleaned;
        if (/^[A-Z\s]+$/.test(cleaned)) {
            return cleaned
                .toLowerCase()
                .split(/\s+/)
                .map(part => part.charAt(0).toUpperCase() + part.slice(1))
                .join(' ');
        }
        return cleaned;
    }

    private extractOnAgDestinationCode(raw: string): string {
        const text = this.stripBrackets(raw || '').trim();
        if (!text) return '';
        const beforeDash = text.split(/\s+-\s+/, 1)[0]?.trim() || '';
        return beforeDash || text;
    }

    private normalizeHunterTransportLocation(raw: string | undefined, packingSplit: string | undefined, purchaseOrderRaw: string | undefined): string {
        const normalized = this.normalizeTransportLocation(raw || '');
        if (normalized && normalized.toUpperCase() !== 'TBC') return normalized;

        const token = (packingSplit || purchaseOrderRaw || '').toString().toUpperCase();
        if (!token) return normalized;
        if (token.includes('UKSOS') || token.includes('GOLDSEAL') || token.includes('DTE')) return 'Great Britain';
        if (/(^|[-\s])DE($|[-\s])/.test(token) || token.includes('ZALANDO')) return 'Germany';
        return normalized;
    }

    private normalizeHunterOrderTransportLocation(packingSplit: string | undefined, purchaseOrderRaw: string | undefined): string {
        const token = (packingSplit || purchaseOrderRaw || '').toString().toUpperCase();
        if (!token) return '';
        if (token.includes('UKSOS') || token.includes('SOS') || token.includes('GOLDSEAL') || token.includes('ECOM') || token.includes('DTE')) return 'Great Britain';
        if (/(^|[-\s])DE($|[-\s])/.test(token) || token.includes('ZALANDO')) return 'Germany';
        return '';
    }

    private extractPranaPlantLabel(raw: string | undefined): string {
        const text = this.stripBrackets(raw || '').trim();
        if (!text) return '';
        const firstSegment = text.split(',')[0]?.trim() || text;
        return firstSegment
            .replace(/\s{2,}/g, ' ')
            .replace(/\s*-\s*$/g, '')
            .trim();
    }

    private resolvePranaDestinationCountry(raws: Array<string | undefined>): string {
        const source = raws
            .map(value => this.stripBrackets(value || '').toLowerCase().trim())
            .find(Boolean) || '';
        if (!source) return '';
        if (source.includes('united states') || source.includes(' us ') || source.startsWith('us ') || source.endsWith(' us') || source.includes('usa') || source.includes('portland') || source.includes('oregon') || source.includes('prana us warehouse')) return 'USA';
        if (source.includes('canada')) return 'Canada';
        if (source.includes('mexico')) return 'Mexico';
        if (source.includes('uk') || source.includes('united kingdom')) return 'UK';
        if (source.includes('australia')) return 'Australia';
        return this.normalizeTransportLocation(source);
    }

    private formatIsoDateString(raw: string | Date | undefined): string {
        const date = this.parseDate(raw);
        if (!date) return '';
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    }

    private shiftDate(raw: string | number | Date | undefined, days: number): Date | null {
        const parsed = this.parseDate(raw);
        if (!parsed) return null;
        const shifted = new Date(parsed);
        shifted.setDate(shifted.getDate() + days);
        return shifted;
    }

    private inferSeasonFromWorksheet(worksheet: ExcelJS.Worksheet, headerRowNumber: number, sourceFilename?: string): string {
        const candidates = new Set<string>();
        candidates.add(worksheet.name || '');
        if (sourceFilename) candidates.add(sourceFilename);
        for (let rowNum = 1; rowNum <= Math.min(headerRowNumber, 5); rowNum++) {
            const row = worksheet.getRow(rowNum);
            row.eachCell(cell => {
                const text = this.stripBrackets(cell.value?.toString() || '').trim();
                if (text) candidates.add(text);
            });
        }

        for (const candidate of candidates) {
            const text = candidate.trim();
            if (!text) continue;
            const fhMatch = text.match(/\bFH(\d{2})\b/i);
            if (fhMatch) return `FH${fhMatch[1]}`.toUpperCase();
            const awMatch = text.match(/\b(?:AW|FW|AH)(\d{2})\b/i);
            if (awMatch) return `AW${awMatch[1]}`.toUpperCase();
            const ssMatch = text.match(/\b(?:SS|SP)(\d{2})\b/i);
            if (ssMatch) return `SS${ssMatch[1]}`.toUpperCase();
            const longMatch = text.match(/\b([FS])W?(\d{2})(\d{2})\b/i);
            if (longMatch) return `${longMatch[1].toUpperCase()}W${longMatch[2]}`;
            const shortMatch = text.match(/\b([FS]W?\d{2})\b/i);
            if (shortMatch) return shortMatch[1].toUpperCase();
            const titleMatch = text.match(/\b([FS])W(\d{2})\b/i);
            if (titleMatch) return `${titleMatch[1].toUpperCase()}W${titleMatch[2]}`;
        }
        return '';
    }

    private inferFoxSeasonFromStyle(raw: string): string {
        const text = this.stripBrackets(raw || '').trim();
        if (!text) return '';
        const match = text.match(/\b([FS]\d{2})[A-Z0-9]*\b/i);
        return match ? match[1].toUpperCase() : '';
    }

    private inferArcteryxSeason(rawDate: string | number | Date | undefined): string {
        const parsed = this.parseDate(rawDate);
        if (!parsed) return '';
        const year = parsed.getFullYear();
        if (!Number.isFinite(year) || year < 2000) return '';
        return `F${String(year).slice(-2)}`;
    }

    private inferFoxSeasonFromDate(rawDate: string | number | Date | undefined): string {
        const parsed = this.parseDate(rawDate);
        if (!parsed) return '';
        const year = parsed.getFullYear();
        if (!Number.isFinite(year) || year < 2000) return '';
        return `F${String(year).slice(-2)}`;
    }

    private normalizeTemplate(rawTemplate: string): string {
        const normalized = (rawTemplate || '').trim().toUpperCase();
        const map: Record<string, string> = { OG: 'BULK', ZNB1: 'BULK', ZMF1: 'BULK', ZDS1: 'BULK', SMS: 'SMS' };
        return map[normalized] || (rawTemplate || 'BULK').trim() || 'BULK';
    }

    private normalizeTransportMethod(raw: string | undefined): string {
        const key = (raw || '').trim().toLowerCase();
        const mapped = TRANSPORT_MAP[key];
        if (mapped) return mapped;
        const codeMatch = key.match(/^(?:v|a|s|c|o|01|1|2|3|4|5|6|7|8|9)\s*[-_.]?\s*(sea|air|courier)\b/);
        if (codeMatch) return codeMatch[1] === 'sea' ? 'Sea' : codeMatch[1] === 'air' ? 'Air' : 'Courier';
        if (/\bsea\b/.test(key)) return 'Sea';
        if (/\bair\b/.test(key)) return 'Air';
        if (/\bcourier\b/.test(key)) return 'Courier';
        return raw ? raw.trim() : 'Sea';
    }

    private normalizeTransportLocation(raw: string | undefined): string {
        const cleaned = this.stripBrackets(raw || '').trim();
        if (!cleaned) return '';
        const key = cleaned.toUpperCase();
        return COUNTRY_NAME_MAP[key] || cleaned;
    }

    private formatPurchaseOrder(basePo: string | undefined, plantPart: string | undefined, destinationPart: string | undefined): string {
        const base = this.stripBrackets(basePo || '').trim();
        const suffixParts = [plantPart, destinationPart]
            .map(part => this.stripBrackets(part || '').trim())
            .filter(Boolean);
        const uniqueSuffixParts = suffixParts.filter((part, index) =>
            suffixParts.findIndex(candidate => candidate.toLowerCase() === part.toLowerCase()) === index
        );

        // Normalize base by removing well-known hunter/origin suffix tokens that should be replaced by table values
        const baseNormalized = base.replace(/\s*-\s*(UKSOS|SOS|GOLDSEAL|DTE|ZALANDO)\s*$/i, '').trim();

        if (!baseNormalized) return this.collapseRepeatedPurchaseOrder(uniqueSuffixParts.join(' - '));
        if (uniqueSuffixParts.length === 0) return this.collapseRepeatedPurchaseOrder(baseNormalized);

        const lowerBase = baseNormalized.toLowerCase();
        const suffixAlreadyInBase = uniqueSuffixParts.every(part => lowerBase.includes(part.toLowerCase()));
        if (suffixAlreadyInBase) return this.collapseRepeatedPurchaseOrder(baseNormalized);

        return this.collapseRepeatedPurchaseOrder(`${baseNormalized} - ${uniqueSuffixParts.join(' - ')}`);
    }

    private collapseRepeatedPurchaseOrder(raw: string | undefined): string {
        const text = this.stripBrackets(raw || '').replace(/\s+/g, ' ').trim();
        if (!text) return '';
        const tokens = text.split(' ');
        const maxChunkSize = Math.floor(tokens.length / 2);
        for (let chunkSize = 1; chunkSize <= maxChunkSize; chunkSize++) {
            if (tokens.length % chunkSize !== 0) continue;
            const chunk = tokens.slice(0, chunkSize).join(' ');
            let repeated = true;
            for (let i = chunkSize; i < tokens.length; i += chunkSize) {
                if (tokens.slice(i, i + chunkSize).join(' ') !== chunk) {
                    repeated = false;
                    break;
                }
            }
            if (repeated) return chunk;
        }
        return text;
    }

    private isLikelyDestinationCountry(value: string | undefined, plantPart: string | undefined): boolean {
        const cleaned = this.stripBrackets(value || '').trim();
        if (!cleaned) return false;
        if (/^\d+$/.test(cleaned)) return false;
        const plant = this.stripBrackets(plantPart || '').trim().toLowerCase();
        const lowered = cleaned.toLowerCase();
        if (plant && lowered === plant) return false;
        if (plant && lowered.includes(plant)) return false;
        if (COUNTRY_NAME_MAP[cleaned.toUpperCase()]) return true;
        return cleaned.length >= 2;
    }

    private extractCountryFromPurchaseOrder(purchaseOrder: string | undefined): string {
        const text = this.stripBrackets(purchaseOrder || '').trim();
        if (!text) return '';
        const suffix = text.split('-').slice(1).join('-').trim();
        if (!suffix) return '';
        return this.normalizeTransportLocation(suffix);
    }

    private parseLooseNumber(raw: string | number | Date | undefined | null): number {
        if (raw === undefined || raw === null || raw === '') return NaN;
        if (typeof raw === 'number') return Number.isFinite(raw) ? raw : NaN;
        if (raw instanceof Date) return NaN;
        const text = String(raw).trim();
        if (!text) return NaN;
        const compact = text.replace(/\s+/g, '').replace(/[’']/g, '');
        const hasComma = compact.includes(',');
        const hasDot = compact.includes('.');
        let normalized = compact;
        if (hasComma && hasDot) {
            if (/^\d{1,3}(?:,\d{3})+(?:\.\d+)?$/.test(compact)) {
                normalized = compact.replace(/,/g, '');
            } else if (/^\d{1,3}(?:\.\d{3})+(?:,\d+)?$/.test(compact)) {
                normalized = compact.replace(/\./g, '').replace(',', '.');
            } else {
                normalized = compact.replace(/,/g, '');
            }
        } else if (hasComma) {
            if (/^\d{1,3}(?:,\d{3})+$/.test(compact)) normalized = compact.replace(/,/g, '');
            else normalized = compact.replace(/,/g, '.');
        } else if (hasDot) {
            if (/^\d{1,3}(?:\.\d{3})+$/.test(compact)) normalized = compact.replace(/\./g, '');
        }
        const parsed = Number(normalized);
        return Number.isFinite(parsed) ? parsed : NaN;
    }

    private parseDate(raw: string | number | Date | undefined): Date | null {
        if (!raw) return null;
        if (raw instanceof Date) return isNaN(raw.getTime()) ? null : raw;
        if (typeof raw === 'number' && Number.isFinite(raw)) {
            const excelEpoch = new Date(Date.UTC(1899, 11, 30));
            const wholeDays = Math.trunc(raw);
            const fractionalDay = raw - wholeDays;
            const millis = wholeDays * 86400000 + Math.round(fractionalDay * 86400000);
            const date = new Date(excelEpoch.getTime() + millis);
            return isNaN(date.getTime()) ? null : date;
        }
        const text = String(raw).trim();
        if (!text) return null;
        const isoMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (isoMatch) { const date = new Date(Number(isoMatch[1]), Number(isoMatch[2]) - 1, Number(isoMatch[3])); return isNaN(date.getTime()) ? null : date; }
        const usMatch = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
        if (usMatch) { const date = new Date(Number(usMatch[3]), Number(usMatch[1]) - 1, Number(usMatch[2])); return isNaN(date.getTime()) ? null : date; }
        const shortUsMatch = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2})$/);
        if (shortUsMatch) {
            const yy = Number(shortUsMatch[3]);
            const fullYear = yy >= 70 ? 1900 + yy : 2000 + yy;
            const date = new Date(fullYear, Number(shortUsMatch[1]) - 1, Number(shortUsMatch[2]));
            return isNaN(date.getTime()) ? null : date;
        }
        const monMatch = text.match(/^(\d{1,2})-([A-Za-z]+)-(\d{4})$/);
        if (monMatch) {
            const months = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
            const monthIndex = months.findIndex(m => monMatch[2].toLowerCase().startsWith(m));
            if (monthIndex >= 0) { const date = new Date(Number(monMatch[3]), monthIndex, Number(monMatch[1])); return isNaN(date.getTime()) ? null : date; }
        }
        const fallbackDate = new Date(text);
        if (!isNaN(fallbackDate.getTime())) return fallbackDate;
        return null;
    }

    private formatDateString(raw: string | Date | undefined): string {
        const date = this.parseDate(raw);
        if (!date) return '';
        const mm = String(date.getMonth() + 1).padStart(2, '0');
        const dd = String(date.getDate()).padStart(2, '0');
        return `${mm}/${dd}/${date.getFullYear()}`;
    }

    private formatManualDateString(raw: string | Date | undefined): string {
        const date = this.parseDate(raw);
        if (!date) return '';
        const mm = String(date.getMonth() + 1).padStart(2, '0');
        const dd = String(date.getDate()).padStart(2, '0');
        return `${mm}/${dd}/${date.getFullYear()}`;
    }

    private stripBrackets(value: string): string {
        if (!value) return value;
        return value.replace(/\[([^\]]+)\]/g, '$1').replace(/\[|\]/g, '').replace(/\s+/g, ' ').trim();
    }

    private buildComments(brand: string | undefined, season: string, buyRound: string, buyDateRaw: string | undefined, template: string): string {
        const brandKey = this.stripBrackets(brand || 'OUTPUT').toLowerCase();
        const b = brandKey === 'hh' || brandKey === 'helly hansen'
            ? 'HH'
            : this.stripBrackets(brand || 'OUTPUT');
        const s = this.stripBrackets(season || 'NOS');
        const round = this.stripBrackets(buyRound || '');
        const tmpl = this.stripBrackets(template || '');
        if (brandKey === 'dynafit') {
            const poMatch = (this.stripBrackets(round || '').match(/po\s*(\d+)/i) || this.stripBrackets(template || '').match(/po\s*(\d+)/i));
            const poToken = poMatch ? `PO${poMatch[1]}` : 'PO2956';
            const dynafitSeason = (() => {
                const upper = s.toUpperCase();
                const year = (upper.match(/\b(20\d{2}|\d{2})\b/) || [])[1] || '27';
                const yy = year.length === 4 ? year.slice(-2) : year;
                if (upper.includes('FH') || upper.includes('FW')) return `FW${yy}`;
                if (upper.includes('SH') || upper.includes('SS')) return `SS${yy}`;
                return `FW${yy}`;
            })();
            return `Dynafit ${dynafitSeason} SMS March 03.05 Buy ${poToken}`.trim();
        }
        if (brandKey === 'prana') {
            return `PRA ${s} Dec Buy 11-DEC Bulk`.trim();
        }
        if (brandKey === 'hh' || brandKey === 'helly hansen') {
            const compact = this.compactProductRange(s).replace(/^FH/i, 'F').replace(/^SH/i, 'S');
            const parts = ['HH', compact || 'F26', 'Bulk'];
            if (round) parts.push(round);
            return parts.join(' ').trim();
        }
        if (round) return `${b} ${s} ${round} ${tmpl}`.trim();
        const parsed = this.parseDate(buyDateRaw);
        if (parsed) {
            const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
            const monShort = months[parsed.getMonth()];
            const day = String(parsed.getDate()).padStart(2, '0');
            const suffix = tmpl ? ` ${tmpl}` : '';
            return `${b} ${s} ${monShort} Buy ${day}-${monShort.toUpperCase()}${suffix}`.trim();
        }
        return `${b} ${s}`.trim();
    }

    private resolveJackWolfskinKeyDate(season: string, fallback: string | Date | undefined): string | Date {
        const seasonText = this.stripBrackets(season || '').toUpperCase();
        const match = seasonText.match(/\b([FS])H?\s*:?\s*(\d{2,4})\b/);
        if (!match) return fallback || '';
        const prefix = match[1];
        const yearToken = match[2];
        const year = yearToken.length === 2 ? 2000 + Number(yearToken) : Number(yearToken);
        if (!Number.isFinite(year) || year < 2000) return fallback || '';
        const month = prefix === 'F' ? 9 : 2;
        const date = new Date(year, month, 1, 8, 0, 0, 0);
        while (date.getDay() !== 5) date.setDate(date.getDate() + 1);
        return date;
    }

    private resolveDynafitExportContext(args: {
        poNumberRaw: string;
        plantPart: string;
        rawFilePo: string;
        buyerPoNumber: string | number;
        productMatch?: ProductSheetRow;
        destinationFromFile: string;
        plantDerivedCountry: string;
        shipToRaw: string;
        transportLocationSource: string;
        effectiveTransportLocation: string;
        getRawVal: (field: string) => any;
        productSupplierFallback: string;
    }) {
        const destinationSuffix = this.stripBrackets(
            args.productMatch?.destinationName
                || args.destinationFromFile
                || args.plantDerivedCountry
                || args.shipToRaw
                || args.transportLocationSource
                || args.effectiveTransportLocation
                || ''
        ).trim();
        const exportPurchaseOrder = this.formatPurchaseOrder(args.poNumberRaw, args.plantPart, destinationSuffix);
        const buyerPoNumber = args.rawFilePo || args.buyerPoNumber?.toString?.().trim?.() || args.poNumberRaw || '';
        const crd = args.productMatch?.crd || args.getRawVal('crd') || args.getRawVal('dynafitLineKeyDate') || args.getRawVal('finalDeliveryDate') || args.getRawVal('exFtyDate') || args.getRawVal('confirmedExFac') || '';
        const exFactory = args.productMatch?.exFactory || args.getRawVal('exFactory') || args.getRawVal('confirmedExFac') || args.getRawVal('exFtyDate') || '';
        const deliveryDate = crd || exFactory || '';
        return {
            destinationSuffix,
            exportPurchaseOrder,
            buyerPoNumber,
            productSupplier: this.stripBrackets(args.productMatch?.factory || '').trim() || args.productSupplierFallback,
            productRange: args.productMatch?.season || 'FH:2027',
            transportMethod: 'Courier',
            ordersTemplate: 'SMS PO Header',
            linesTemplate: 'SMS Non EDI (New)',
            deliveryDate,
            startDate: deliveryDate,
            cancelDate: deliveryDate,
            lineKeyDate: exFactory || deliveryDate || '',
            resolvedColour: args.productMatch?.colour || '',
            crd,
            exFactory,
        };
    }

    async extractProductSheetMap(buffer: any): Promise<Record<string, ProductSheetRow[]>> {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);
        return this.extractProductSheetMapFromWorkbook(workbook);
    }

    private isLikelyBuySheet(worksheet: ExcelJS.Worksheet, headerRow: number, aliases: Record<string, string>): boolean {
        const row = worksheet.getRow(headerRow);
        const mappedFields = new Set<string>();
        const headerVals = new Set<string>();
        let hasGrandTotalHeader = false;
        row.eachCell(cell => {
            const v = normalizeHeaderText(cell.value?.toString() || '');
            const mapped = aliases[v];
            if (mapped) mappedFields.add(mapped);
            if (v === 'grand total') hasGrandTotalHeader = true;
            if (v) headerVals.add(v);
        });

        const hasProduct = mappedFields.has('product');
        const hasPurchaseOrder = mappedFields.has('purchaseOrder') || mappedFields.has('buyerPoNumber');
        const hasQuantity = mappedFields.has('quantity') || mappedFields.has('finalQty') || hasGrandTotalHeader;
        const hasSeason = mappedFields.has('season');
        const hasBuyStructure = mappedFields.has('transportMethod')
            || mappedFields.has('template')
            || mappedFields.has('exFtyDate')
            || mappedFields.has('buyDate')
            || mappedFields.has('status');
        const looksLikeRossignolBuySheet =
            headerVals.has('destination')
            && headerVals.has('product code')
            && headerVals.has('sku')
            && headerVals.has('shipping date')
            && headerVals.has('tot qty')
            && headerVals.has('m88 ref')
            && headerVals.has('color name')
            && headerVals.has('size name');
        const looksLikeFoxBuySheet =
            headerVals.has('purchasing document number')
            && headerVals.has('item number of purchasing document')
            && headerVals.has('material')
            && headerVals.has('material description')
            && headerVals.has('order qty')
            && headerVals.has('ex factory date')
            && (headerVals.has('vendor name') || headerVals.has('goods supplier name'));

        return (
            hasProduct && (
            (hasPurchaseOrder && hasQuantity)
            || (hasSeason && hasQuantity)
            || (hasPurchaseOrder && hasBuyStructure)
            )
            || looksLikeRossignolBuySheet
            || looksLikeFoxBuySheet
        );
    }

    async analyzeWorkbook(buffer: any): Promise<{ productSheetMap: Record<string, ProductSheetRow[]>; hasBuySheet: boolean }> {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);
        const productSheetMap = this.extractProductSheetMapFromWorkbook(workbook);
        const aliases = this.getFallbackColumnAliases();
        const sourceFilename = '';
        let hasBuySheet = false;
        let hasProductSheet = false;
        for (const ws of workbook.worksheets) {
            const { isProductSheet, headerRow } = this.detectProductSheet(ws);
            if (isProductSheet) { hasProductSheet = true; continue; }
            if (this.isLikelyBuySheet(ws, headerRow, aliases)) { hasBuySheet = true; break; }
        }
        if (!hasBuySheet && !hasProductSheet) hasBuySheet = true;
        return { productSheetMap, hasBuySheet };
    }

    async processBuyFile(buffer: any, options?: {
        manualPurchaseOrder?: string; manualDestination?: string; manualProductRange?: string;
        manualTemplate?: string; manualLinesTemplate?: string; manualComments?: string;
        manualKeyDate?: string; manualKeyUser1?: string; manualKeyUser2?: string;
        manualKeyUser3?: string; manualKeyUser4?: string; manualKeyUser5?: string;
        manualSeason?: string; manualCustomer?: string; manualBrand?: string;
        defaultQuantityIfMissing?: boolean; productSheetMap?: Record<string, ProductSheetRow[]>;
        llBeanReferenceSizesBuffer?: any;
        sourceFilename?: string;
    }): Promise<{ data: ProcessedPO[]; errors: ValidationError[]; formatDetection?: FormatDetection }> {
        const sourceFilename = (options?.sourceFilename || '').toLowerCase();
        if (sourceFilename.includes("product shi") && !sourceFilename.includes("dynafit")) {
            return { data: [], errors: this.errors };
        }
        const manualPurchaseOrder = options?.manualPurchaseOrder?.toString().trim() || '';
        const manualDestination = options?.manualDestination?.toString().trim() || '';
        const manualProductRange = options?.manualProductRange?.toString().trim() || '';
        const manualSeason = options?.manualSeason?.toString().trim() || '';
        const manualTemplate = options?.manualTemplate?.toString().trim() || '';
        const manualLinesTemplate = options?.manualLinesTemplate?.toString().trim() || '';
        const manualComments = options?.manualComments?.toString().trim() || '';
        const manualKeyDate = options?.manualKeyDate?.toString().trim() || '';
        const manualKeyUser1 = options?.manualKeyUser1?.toString().trim() || '';
        const manualKeyUser2 = options?.manualKeyUser2?.toString().trim() || '';
        const manualKeyUser3 = options?.manualKeyUser3?.toString().trim() || '';
        const manualKeyUser4 = options?.manualKeyUser4?.toString().trim() || '';
        const manualKeyUser5 = options?.manualKeyUser5?.toString().trim() || '';
        const manualCustomer = options?.manualCustomer?.toString().trim() || '';
        const manualBrand = options?.manualBrand?.toString().trim() || '';
        const sourceNameHint = (options?.sourceFilename || '').toLowerCase();
        const looksLikePranaSource = sourceNameHint.includes('prana');

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);
        const workbookProductMap = this.extractProductSheetMapFromWorkbook(workbook);
        const inlineProductMap = Object.keys(workbookProductMap).length === 0
            ? this.extractInlineProductSheetMapFromBuyWorkbook(workbook)
            : {};
        const productSheetMap: Record<string, ProductSheetRow[]> = { ...(options?.productSheetMap || {}), ...workbookProductMap, ...inlineProductMap };
        const fallbackAliases = this.getFallbackColumnAliases();
        const seasonOverride = manualSeason || manualProductRange;

        const buySheetCandidates = workbook.worksheets
            .map(ws => {
                const candidate = this.detectHeaderRow(ws);
                const productSheet = this.detectProductSheet(ws);
                const isLikelyBuy = !productSheet.isProductSheet && this.isLikelyBuySheet(ws, candidate, fallbackAliases);
                return { ws, headerRow: candidate, isLikelyBuy };
            })
            .filter(entry => entry.isLikelyBuy);

        if (buySheetCandidates.length === 0) {
            return { data: [], errors: this.errors };
        }

        let worksheet = workbook.worksheets[0];
        let headerRowNumber = this.detectHeaderRow(worksheet);
        let bestScore = -1;
        if (looksLikePranaSource) {
            const pranaCleanSheet = buySheetCandidates.find(candidateSheet => {
                const row = candidateSheet.ws.getRow(candidateSheet.headerRow);
                const headers = new Set<string>();
                row.eachCell(cell => {
                    const v = cell.value?.toString().toLowerCase().trim() || '';
                    if (v) headers.add(v);
                });
                return headers.has('po#')
                    && headers.has('style #')
                    && headers.has('style description')
                    && headers.has('style status')
                    && headers.has('qty')
                    && headers.has('ship via')
                    && headers.has('ship to')
                    && !headers.has('product name')
                    && !headers.has('customer name')
                    && !headers.has('factory');
            });
            if (pranaCleanSheet) {
                worksheet = pranaCleanSheet.ws;
                headerRowNumber = pranaCleanSheet.headerRow;
            } else {
                for (const candidateSheet of (buySheetCandidates.length > 0 ? buySheetCandidates : workbook.worksheets.map(ws => ({ ws, headerRow: this.detectHeaderRow(ws) })))) {
                    const ws = candidateSheet.ws;
                    const candidate = candidateSheet.headerRow;
                    const row = ws.getRow(candidate);
                    let score = 0;
                    row.eachCell(cell => { const v = cell.value?.toString().toLowerCase().trim() || ''; if (fallbackAliases[v]) score++; });
                    if (score > bestScore) { bestScore = score; worksheet = ws; headerRowNumber = candidate; }
                }
            }
        } else {
            for (const candidateSheet of (buySheetCandidates.length > 0 ? buySheetCandidates : workbook.worksheets.map(ws => ({ ws, headerRow: this.detectHeaderRow(ws) })))) {
                const ws = candidateSheet.ws;
                const candidate = candidateSheet.headerRow;
                const row = ws.getRow(candidate);
                let score = 0;
                row.eachCell(cell => { const v = cell.value?.toString().toLowerCase().trim() || ''; if (fallbackAliases[v]) score++; });
                if (score > bestScore) { bestScore = score; worksheet = ws; headerRowNumber = candidate; }
            }
        }

        let selectedSheetProductDetection = this.detectProductSheet(worksheet);
        if (selectedSheetProductDetection.isProductSheet) {
            return { data: [], errors: this.errors };
        }

        let inferredSeasonFromSheet = this.inferSeasonFromWorksheet(worksheet, headerRowNumber, options?.sourceFilename);
        const headerRow = worksheet.getRow(headerRowNumber);
        const headerKeysInRow = new Set<string>();
        headerRow.eachCell(cell => {
            const headerText = cell.value?.toString().trim();
            if (!headerText) return;
            headerKeysInRow.add(normalizeHeaderText(headerText));
        });
        const looksLikePeakPerformanceDetailedSheet =
            !looksLikePranaSource
            && headerKeysInRow.has('product name')
            && headerKeysInRow.has('color name')
            && headerKeysInRow.has('customer name')
            && headerKeysInRow.has('factory');
        const looksLikePeakPerformanceEarlyBuySheet =
            !looksLikePranaSource
            && headerKeysInRow.has('article code [sap]')
            && headerKeysInRow.has('model code [sap]')
            && headerKeysInRow.has('primary color peak pdm code')
            && headerKeysInRow.has('article full colors')
            && headerKeysInRow.has('production supplier name')
            && (headerKeysInRow.has('final po qty') || headerKeysInRow.has('buy 1 agreed qty') || headerKeysInRow.has('buy 1 - tracking no.'));
        const looksLikePeakPerformanceSheet = looksLikePeakPerformanceDetailedSheet || looksLikePeakPerformanceEarlyBuySheet;
        const firstDataRow = worksheet.getRow(headerRowNumber + 1);
        const allMappings = await getAllColumnMappings();
        const knownCustomers = Array.from(new Set(allMappings.map((m: any) => m.customer)));
        const lowerKnown = knownCustomers.map((c: string) => c.toLowerCase());
        let detectedCustomer = looksLikePeakPerformanceSheet ? 'Peak Performance' : 'DEFAULT';
        firstDataRow.eachCell(cell => {
            if (looksLikePeakPerformanceSheet) return;
            const val = cell.value?.toString().trim();
            if (!val) return;
            const lowerVal = val.toLowerCase();
            if (lowerVal.includes('pt uwu jump - jw') || lowerVal.includes('jack wolfskin') || lowerVal.includes(' jw')) {
                detectedCustomer = 'Jack Wolfskin';
                return;
            }
            if (lowerKnown.includes(lowerVal)) {
                detectedCustomer = knownCustomers.find((c: string) => c.toLowerCase() === lowerVal) || 'DEFAULT';
                return;
            }
            const mappedCustomer = BRAND_CUSTOMER_MAP[lowerVal];
            if (mappedCustomer && mappedCustomer !== 'DEFAULT') {
                detectedCustomer = mappedCustomer;
            }
        });
        if (!looksLikePeakPerformanceSheet) {
            if (sourceNameHint.includes('mammut')) detectedCustomer = 'Mammut';
            if (sourceNameHint.includes('vuori') || sourceNameHint.includes('podetails')) detectedCustomer = 'Vuori';
            if (sourceNameHint.includes('marmot')) detectedCustomer = 'Marmot';
        }

        if (detectedCustomer === 'Marmot') {
            const marmotPreferredSheet = workbook.worksheets.find(ws => normalizeHeaderText(ws.name) === 'all data');
            if (marmotPreferredSheet) {
                worksheet = marmotPreferredSheet;
                headerRowNumber = this.detectHeaderRow(worksheet);
                inferredSeasonFromSheet = this.inferSeasonFromWorksheet(worksheet, headerRowNumber, options?.sourceFilename);
                selectedSheetProductDetection = this.detectProductSheet(worksheet);
                if (selectedSheetProductDetection.isProductSheet) {
                    return { data: [], errors: this.errors };
                }
            }
        }

        const colMap = await getColumnMapping(detectedCustomer);
        const normalizedColMap: Record<string, string> = {};
        Object.entries(colMap).forEach(([k, v]) => { normalizedColMap[normalizeHeaderText(k)] = v as string; });
        Object.entries(fallbackAliases).forEach(([k, v]) => { if (!normalizedColMap[k]) normalizedColMap[k] = v; });

        const headerMap: Record<string, number> = {};
        let inferredSizeCol: number | null = null;
        const unmappedHeaders: { headerText: string; colNumber: number }[] = [];
        let maxColNumber = 0;
        let lastColHeaderText = '';
        const looksLikeJackWolfskinBuy = headerKeysInRow.has('stylecolor')
            && headerKeysInRow.has('qty jan buy size-split')
            && headerKeysInRow.has('bp no')
            && headerKeysInRow.has('vendor confirmed etd');
        const looksLikeVuoriBuy = headerKeysInRow.has('purchase order no')
            && headerKeysInRow.has('requested etd|n')
            && headerKeysInRow.has('confirmed ex-factory date|n')
            && headerKeysInRow.has('line number')
            && headerKeysInRow.has('product name')
            && headerKeysInRow.has('warehouse name');
        const looksLikePeakPerformanceBuy = headerKeysInRow.has('article code [sap]')
            && headerKeysInRow.has('model code [sap]')
            && headerKeysInRow.has('product name')
            && headerKeysInRow.has('color name')
            && headerKeysInRow.has('customer name')
            && !looksLikePranaSource;
        if (looksLikeVuoriBuy) detectedCustomer = 'Vuori';
        if (looksLikePeakPerformanceSheet || looksLikePeakPerformanceBuy) detectedCustomer = 'Peak Performance';
        if (looksLikePranaSource) detectedCustomer = 'Prana';
        if (sourceNameHint.includes('mammut')) detectedCustomer = 'Mammut';
        if (sourceNameHint.includes('vuori') || sourceNameHint.includes('podetails')) detectedCustomer = 'Vuori';
        if (sourceNameHint.includes('marmot')) detectedCustomer = 'Marmot';

        headerRow.eachCell((cell, colNumber) => {
            const headerText = cell.value?.toString().trim();
            if (!headerText) return;
            const isDynafitHint = /dynafit/i.test(sourceFilename) || (detectedCustomer || '').toLowerCase() === 'dynafit';
            if (colNumber > maxColNumber) { maxColNumber = colNumber; lastColHeaderText = headerText; }
            const headerKey = normalizeHeaderText(headerText);
            if (isDynafitHint && /^(?:po|p)\d{4,}$/i.test(headerText.trim()) && !headerMap['purchaseOrder']) {
                headerMap['purchaseOrder'] = colNumber;
            }
            if (isDynafitHint && headerKey === 'po number' && !headerMap['buyerPoNumber']) {
                headerMap['buyerPoNumber'] = colNumber;
                return;
            }
            if (isDynafitHint && headerKey === 'crd' && !headerMap['dynafitLineKeyDate']) {
                headerMap['dynafitLineKeyDate'] = colNumber;
                return;
            }
            if (headerKey === 'material' && !headerMap['foxMaterialCode']) {
                headerMap['foxMaterialCode'] = colNumber;
            }
            if (headerKey === 'material description' && !headerMap['foxMaterialDescription']) {
                headerMap['foxMaterialDescription'] = colNumber;
            }
            if (headerKey === 'product name' && !headerMap['onAgProductName']) {
                headerMap['onAgProductName'] = colNumber;
            }
            if (headerKey === 'product name' && !headerMap['inlineProductName']) {
                headerMap['inlineProductName'] = colNumber;
            }
            if (headerKey === 'm88 ref' && !headerMap['rossignolM88Ref']) {
                headerMap['rossignolM88Ref'] = colNumber;
            }
            if (headerKey === 'product code' && !headerMap['rossignolProductCode']) {
                headerMap['rossignolProductCode'] = colNumber;
            }
            if (headerKey === 'item' && !headerMap['productAlt']) {
                headerMap['productAlt'] = colNumber;
            }
            if (headerKey === 'primary color peak pdm code' && !headerMap['peakPerformanceBaseColour']) {
                headerMap['peakPerformanceBaseColour'] = colNumber;
            }
            if (headerKey === 'color name' && !headerMap['inlineColorName']) {
                headerMap['inlineColorName'] = colNumber;
            }
            if (looksLikePranaSource && headerKey === 'color code' && !headerMap['pranaColorCode']) {
                headerMap['pranaColorCode'] = colNumber;
                return;
            }
            if (looksLikePranaSource && headerKey === 'color' && !headerMap['pranaColorText']) {
                headerMap['pranaColorText'] = colNumber;
            }
            if (headerKey === 'color description' && !headerMap['inlineColorDescription']) {
                headerMap['inlineColorDescription'] = colNumber;
            }
            if (headerKey === 'short text' && !headerMap['marmotShortText']) {
                headerMap['marmotShortText'] = colNumber;
            }
            if (sourceNameHint.includes('marmot') && headerKey === 'style' && !headerMap['marmotStyle']) {
                headerMap['marmotStyle'] = colNumber;
            }
            if (headerKey === 'style color' && !headerMap['inlineStyleColor']) {
                headerMap['inlineStyleColor'] = colNumber;
            }
            if (headerKey === 'stylecolor' && !headerMap['inlineStyleColor']) {
                headerMap['inlineStyleColor'] = colNumber;
            }
            if (headerKey === 'your reference' && !headerMap['ourReference']) {
                headerMap['ourReference'] = colNumber;
            }
            if (headerKey === 'size name' && headerMap['sizeName'] && !headerMap['inlineSizeName']) {
                headerMap['inlineSizeName'] = colNumber;
            }
            if (headerKey === 'factory' && !headerMap['inlineFactory']) {
                headerMap['inlineFactory'] = colNumber;
            }
            if (headerKey === 'buyer item' && !headerMap['onAgBuyerItem']) {
                headerMap['onAgBuyerItem'] = colNumber;
            }
            if (headerKey === 'destination name' && !headerMap['onAgDestinationName']) {
                headerMap['onAgDestinationName'] = colNumber;
            }
            if (headerKey === 'destination name' && !headerMap['arcteryxDestinationName']) {
                headerMap['arcteryxDestinationName'] = colNumber;
            }
            if (headerKey === 'dest country' && !headerMap['transportLocation']) {
                headerMap['transportLocation'] = colNumber;
            }
            if (headerKey === 'ultimate destination' && !headerMap['ultimateDestination']) {
                headerMap['ultimateDestination'] = colNumber;
            }
            if (headerKey === 'ult. destination' && !headerMap['ultimateDestination']) {
                headerMap['ultimateDestination'] = colNumber;
            }
            if (headerKey === 'ult destination' && !headerMap['ultimateDestination']) {
                headerMap['ultimateDestination'] = colNumber;
            }
            if (headerKey === 'ship to' && !headerMap['shipTo']) {
                headerMap['shipTo'] = colNumber;
            }
            if (headerKey === 'whs' && !headerMap['whs']) {
                headerMap['whs'] = colNumber;
            }
            if (headerKey === 'whs' && !headerMap['plantName']) {
                headerMap['plantName'] = colNumber;
            }
            if (headerKey === 'goods supplier name' && !headerMap['goodsSupplierName']) {
                headerMap['goodsSupplierName'] = colNumber;
            }
            if (headerKey === 'warehouse name' && !headerMap['warehouseName']) {
                headerMap['warehouseName'] = colNumber;
            }
            if (headerKey === 'po company name' && !headerMap['hhCompanyName']) {
                headerMap['hhCompanyName'] = colNumber;
            }
            if (headerKey === 'final xf date 3.16' && !headerMap['hhFinalXfDate']) {
                headerMap['hhFinalXfDate'] = colNumber;
            }
            if (headerKey === 'confirmed delivery date' && !headerMap['hhConfirmedDeliveryDate']) {
                headerMap['hhConfirmedDeliveryDate'] = colNumber;
            }
            if (headerKey === 'vendor confirmed etd' && !headerMap['confirmedExFac']) {
                headerMap['confirmedExFac'] = colNumber;
            }
            if (headerKey === 'etd' && !headerMap['exFtyDate']) {
                headerMap['exFtyDate'] = colNumber;
            }
            if (headerKey === 'qty jan buy size-split' && !headerMap['quantity']) {
                headerMap['quantity'] = colNumber;
            }
            if (headerKey === 'bp no' && !headerMap['buyerPoNumber']) {
                headerMap['buyerPoNumber'] = colNumber;
            }
            if (headerKey === 'surcharges') {
                return;
            }
            if (looksLikeJackWolfskinBuy && headerKey === 'material') {
                if (!headerMap['jwsMaterialCode']) headerMap['jwsMaterialCode'] = colNumber;
                return;
            }
            if (headerKey === 'grand total' && !headerMap['finalQty']) {
                headerMap['finalQty'] = colNumber;
            }
            if (headerKey === 'po number' && !headerMap['arcteryxBuyerPo']) {
                headerMap['arcteryxBuyerPo'] = colNumber;
            }
            if (headerKey === 'packing splits' && !headerMap['hunterPackingSplit']) {
                headerMap['hunterPackingSplit'] = colNumber;
            }
            if (headerKey === 'xs' && !headerMap['hunterQtyXS']) {
                headerMap['hunterQtyXS'] = colNumber;
            }
            if (headerKey === 's' && !headerMap['hunterQtyS']) {
                headerMap['hunterQtyS'] = colNumber;
            }
            if (headerKey === 'm' && !headerMap['hunterQtyM']) {
                headerMap['hunterQtyM'] = colNumber;
            }
            if (headerKey === 'l' && !headerMap['hunterQtyL']) {
                headerMap['hunterQtyL'] = colNumber;
            }
            if (headerKey === 'xl' && !headerMap['hunterQtyXL']) {
                headerMap['hunterQtyXL'] = colNumber;
            }
            if (headerKey === 'crd' && !headerMap['pranaCrd']) {
                headerMap['pranaCrd'] = colNumber;
            }
            if (headerKey === 'confirmed crd dt (vendor) -(vendor confirmed crd dt)' && !headerMap['vansConfirmedVendorCrd']) {
                headerMap['vansConfirmedVendorCrd'] = colNumber;
            }
            if (headerKey === 'brand requested crd' && !headerMap['vansBrandRequestedCrd']) {
                headerMap['vansBrandRequestedCrd'] = colNumber;
            }
            const internalField = normalizedColMap[headerKey];
            const fallbackField = fallbackAliases[headerKey];
            if (internalField && internalField !== 'ignore') {
                if (!headerMap[internalField]) headerMap[internalField] = colNumber;
                else if (internalField === 'plant' && headerKey === 'dc plant' && !headerMap['plantName']) headerMap['plantName'] = colNumber;
            } else if (internalField === 'ignore') {
                if (fallbackField === 'transportLocation') {
                    if (!headerMap['transportLocation']) headerMap['transportLocation'] = colNumber;
                }
                return;
            } else if (headerKey === 'dc plant' && headerMap['plant'] && !headerMap['plantName']) {
                headerMap['plantName'] = colNumber;
            } else {
                if (!headerMap['sizeName'] && inferredSizeCol === null && this.looksLikeSizeHeader(headerText)) {
                    inferredSizeCol = colNumber;
                    return;
                }
                unmappedHeaders.push({ headerText, colNumber });
            }
        });

        const sourceNameKey = (options?.sourceFilename || '').toLowerCase();
        if (!headerMap['product'] && headerMap['foxMaterialCode'] && sourceNameKey.includes('fox racing')) {
            headerMap['product'] = headerMap['foxMaterialCode'];
        }
        if (sourceNameKey.includes('511 tactical') && !headerMap['quantity']) {
            headerRow.eachCell((cell, colNumber) => {
                const headerKey = normalizeHeaderText(cell.value?.toString().trim() || '');
                if (headerKey === 'total' && !headerMap['quantity']) {
                    headerMap['quantity'] = colNumber;
                }
            });
        }

        // Detect pre-computed NG PO in last column (ON AG INFOR, etc.)
        // If a manual PO was supplied, keep the real file PO column intact for buyer PO extraction.
        if (!manualPurchaseOrder && maxColNumber > 0 && /^(?:po|p)\d{4,}$/i.test(lastColHeaderText)) {
            headerMap['purchaseOrder'] = maxColNumber;
        }

        const precomputedPoColNumber = (!manualPurchaseOrder && /^(?:po|p)\d{4,}$/i.test(lastColHeaderText)) ? maxColNumber : null;
        const pivotFormat = detectPivotFormatFromHeaders(
            Array.from({ length: maxColNumber }, (_, i) => {
                const cell = headerRow.getCell(i + 1);
                return { colNumber: i + 1, headerText: cell.value?.toString().trim() || '' };
            }).filter(h => h.headerText),
            fallbackAliases,
            (h) => this.shouldSilentlyIgnoreHeader(h),
        );
        const hhWorkbookHint = /helly hansen|\bhh\b/i.test(options?.sourceFilename || '') || /helly hansen/i.test(detectedCustomer);
        const hhFirstSizeCol = hhWorkbookHint
            ? Array.from({ length: maxColNumber }, (_, i) => ({
                colNumber: i + 1,
                headerText: headerRow.getCell(i + 1).value?.toString().trim() || '',
            })).find(({ headerText }) => isLikelyPivotSizeHeader(headerText))?.colNumber || 0
            : 0;
        const hhPivotFallback = hhWorkbookHint
            ? Array.from({ length: maxColNumber }, (_, i) => {
                const colNumber = i + 1;
                const headerText = headerRow.getCell(colNumber).value?.toString().trim() || '';
                return { colNumber, headerText };
            }).filter(({ colNumber, headerText }) => {
                const normalized = normalizeHeaderText(headerText);
                if (!normalized) return false;
                if (this.shouldSilentlyIgnoreHeader(headerText)) return false;
                if (fallbackAliases[normalized]) return false;
                if (hhFirstSizeCol > 0 && colNumber >= hhFirstSizeCol) return true;
                return isLikelyPivotSizeHeader(headerText);
            })
            : [];
        const pivotColumnsByNumber = new Map<number, { colNumber: number; headerText: string }>();
        pivotFormat.pivotColumns.forEach(col => pivotColumnsByNumber.set(col.colNumber, col));
        hhPivotFallback.forEach(col => {
            if (!pivotColumnsByNumber.has(col.colNumber)) pivotColumnsByNumber.set(col.colNumber, col);
        });
        const effectivePivotFormat = pivotColumnsByNumber.size > 0
            ? { ...pivotFormat, isPivotFormat: pivotFormat.isPivotFormat || hhPivotFallback.length > 0, pivotColumns: Array.from(pivotColumnsByNumber.values()) }
            : pivotFormat;
        const pivotColumnNumbers = new Set(effectivePivotFormat.pivotColumns.map(col => col.colNumber));

        unmappedHeaders.forEach(({ headerText, colNumber }) => {
            if (pivotColumnNumbers.has(colNumber)) return;
            if (hhWorkbookHint && isLikelyPivotSizeHeader(headerText)) return;
            if (!this.shouldSilentlyIgnoreHeader(headerText)) {
                this.errors.push({ field: 'Mapping', row: 1, message: `Unmapped column ignored: ${headerText}`, severity: 'WARNING' });
            }
        });

        if (!headerMap['sizeName'] && inferredSizeCol !== null && !effectivePivotFormat.isPivotFormat) {
            headerMap['sizeName'] = inferredSizeCol;
            this.errors.push({ field: 'Mapping', row: 1, message: 'Inferred mapping: sizeName from size-like column.', severity: 'WARNING' });
        }

        const useDefaultSizeBucket = !headerMap['sizeName'] && !effectivePivotFormat.isPivotFormat && detectedCustomer !== 'Columbia';
        if (useDefaultSizeBucket) {
            this.errors.push({ field: 'Mapping', row: 1, message: "No size column detected. Using default 'One Size' for all rows.", severity: 'WARNING' });
        }

        const allowMissingPurchaseOrder = !!manualPurchaseOrder;
        const allowMissingQuantity = !!options?.defaultQuantityIfMissing && !headerMap['quantity'] && !headerMap['finalQty'];
        const MANDATORY = ['product'];
        if (!allowMissingPurchaseOrder) MANDATORY.push('purchaseOrder');
        if (!allowMissingQuantity) MANDATORY.push('quantity');
        const missing = MANDATORY.filter(f => !headerMap[f] && !headerMap['finalQty']);
        const isOutputFile = !!(headerMap['lineItem'] || headerMap['productRange']);
        if (missing.length > 0) {
            let msg = `Missing column mappings: ${missing.join(', ')} for customer ${detectedCustomer}.`;
            if (isOutputFile) msg += ' ⚠️ Looks like an NG Output File was uploaded instead of a raw buy file.';
            else msg += ' Please update column_mapping table.';
            this.errors.push({ field: 'File Format', row: 1, message: msg, severity: 'CRITICAL' });
            if (this.runId) {
                await logEvent({ eventName: 'VALIDATION_FAILED', userId: this.userId || 'system', runId: this.runId, metadata: { error_type: 'MISSING_MAPPING', customer: detectedCustomer, missing_fields: missing } });
            }
            return { data: [], errors: this.errors };
        }

        const [factoryMap, mloMap] = await Promise.all([getFactoryMapping(), getMloMapping()]);
        const results: Map<string, ProcessedPO> = new Map();
        const seenOrderKeys = new Set<string>();
        const warnedInferredCategory = new Set<string>();
        let skippedMissingSeason = 0;
        let warnedDefaultQty = false;
        const rossignolBuyToken = (() => {
            const headerText = this.stripBrackets(worksheet.getRow(1).getCell(1).text || worksheet.name || '').trim();
            const sourceText = headerText || (options?.sourceFilename || '');
            const match = sourceText.match(/\bBUY\s*\d+\b/i);
            return match ? match[0].replace(/\s+/g, ' ').toUpperCase() : '';
        })();

        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber <= headerRowNumber) return;
            const getRawVal = (field: string) => { const col = headerMap[field]; if (!col) return undefined; return this.getCellValue(row.getCell(col)); };
            const getVal = (field: string) => { const raw = getRawVal(field); if (raw instanceof Date) return raw.toISOString().split('T')[0]; return raw?.toString().trim(); };

            const rawFilePo = getVal('purchaseOrder');
            const brand = this.stripBrackets(getVal('brand') || '');
            const strongSourceBrand = (() => {
                if (looksLikePranaSource) return 'prana';
                if (looksLikePeakPerformanceSheet || looksLikePeakPerformanceBuy) return 'peak performance';
                const custRaw = (getVal('customerName') || '').toLowerCase();
                const sourceName = (options?.sourceFilename || '').toLowerCase();
                if (sourceName.includes('podetails')) return 'vuori';
                if ((BRAND_CUSTOMER_MAP[custRaw] || '').toLowerCase() === 'vans' || custRaw.includes('vans')) return 'vans';
                const prodRaw = (getVal('product') || '').trim();
                if (/^RL[A-Z0-9]/i.test(prodRaw)) return 'rossignol';
                if (sourceName.includes('vuori')) return 'vuori';
                if (sourceName.includes('helly hansen') || sourceName.startsWith('hh') || sourceName.includes('_hh') || sourceName.includes('hh_')) return 'hh';
                if (sourceName.includes('jack wolfskin')) return 'jack wolfskin';
                if (sourceName.includes('ll bean') || sourceName.includes('l.l.bean') || sourceName.startsWith('llb') || /\bllb\b/i.test(sourceName)) return 'll bean';
                if (sourceName.includes('marmot')) return 'marmot';
                if (sourceName.includes('burton')) return 'burton';
                if (sourceName.includes('dynafit')) return 'dynafit';
                if (sourceName.includes('mammut')) return 'mammut';
                if (sourceName.includes('fox racing')) return 'fox racing';
                if (sourceName.includes('511 tactical')) return '511 tactical';
                if (sourceName.includes('evo')) return 'evo';
                if (sourceName.includes('cotopaxi')) return 'cotopaxi';
                if (sourceName.includes('haglofs')) return 'haglofs';
                if (sourceName.includes('hunter')) return 'hunter';
                if (sourceName.includes('66 degrees north') || sourceName.includes('66north')) return '66 degrees north';
                if (custRaw.includes('burton')) return 'burton';
                if (custRaw.includes('dynafit')) return 'dynafit';
                if (custRaw.includes('on ag') || custRaw.includes('on running')) return 'on ag';
                if (custRaw.includes('peak performance')) return 'peak performance';
                if (custRaw.includes('prana')) return 'prana';
                if (custRaw.includes('cotopaxi')) return 'cotopaxi';
                if (custRaw.includes('vuori')) return 'vuori';
                if (custRaw.includes('helly hansen') || custRaw === 'hh') return 'helly hansen';
                if (custRaw.includes('jack wolfskin')) return 'jack wolfskin';
                if (custRaw.includes('ll bean') || custRaw.includes('l.l.bean') || custRaw.startsWith('llb')) return 'll bean';
                if (custRaw.includes('marmot')) return 'marmot';
                if (custRaw.includes('fox racing') || custRaw === 'fox') return 'fox racing';
                if (custRaw.includes('mammut')) return 'mammut';
                if (custRaw.includes('511 tactical')) return '511 tactical';
                if (custRaw.includes('evo')) return 'evo';
                if (custRaw.includes('haglofs')) return 'haglofs';
                if (custRaw.includes('hunter')) return 'hunter';
                if (custRaw.includes('66 degrees north') || custRaw.includes('66north')) return '66 degrees north';
                if (custRaw.includes('north face') || custRaw.includes('tnf')) return 'tnf';
                if (custRaw.includes('columbia') && custRaw.length > 0) return 'columbia';
                if (custRaw.includes('arcteryx') || custRaw.includes("arc'teryx")) return 'arcteryx';
                const suppRaw = (getVal('vendorName') || getVal('productSupplier') || '').toLowerCase();
                if (suppRaw.includes('pt uwu jump - jw') || suppRaw.includes('jack wolfskin')) return 'jack wolfskin';
                if (suppRaw.includes('madison 88')) return 'tnf';
                if (suppRaw.includes('uwu jump')) return looksLikePeakPerformanceSheet || looksLikePeakPerformanceBuy ? 'peak performance' : 'tnf';
                if (suppRaw.includes('llb') || suppRaw.includes('jaytex') || suppRaw.includes('ll bean')) return 'll bean';
                return '';
            })();
            const inferredBrand = this.stripBrackets(manualBrand || '') || (looksLikePranaSource ? 'prana' : ((looksLikePeakPerformanceSheet || looksLikePeakPerformanceBuy) ? 'peak performance' : strongSourceBrand || brand || (() => {
                const custRaw = (getVal('customerName') || '').toLowerCase();
                if (custRaw.includes('burton')) return 'burton';
                if (custRaw.includes('dynafit')) return 'dynafit';
                if (custRaw.includes('on ag') || custRaw.includes('on running')) return 'on ag';
                if (custRaw.includes('peak performance')) return 'peak performance';
                if (custRaw.includes('prana')) return 'prana';
                if (custRaw.includes('cotopaxi')) return 'cotopaxi';
                if (custRaw.includes('vuori')) return 'vuori';
                if (custRaw.includes('helly hansen') || custRaw === 'hh') return 'helly hansen';
                if (custRaw.includes('jack wolfskin')) return 'jack wolfskin';
                if (custRaw.includes('ll bean') || custRaw.includes('l.l.bean') || custRaw.startsWith('llb')) return 'll bean';
                if (custRaw.includes('marmot')) return 'marmot';
                if (custRaw.includes('fox racing') || custRaw === 'fox') return 'fox racing';
                if (custRaw.includes('mammut')) return 'mammut';
                if (custRaw.includes('511 tactical')) return '511 tactical';
                if (custRaw.includes('evo')) return 'evo';
                if (custRaw.includes('haglofs')) return 'haglofs';
                if (custRaw.includes('hunter')) return 'hunter';
                if (custRaw.includes('66 degrees north') || custRaw.includes('66north')) return '66 degrees north';
                if (custRaw.includes('north face') || custRaw.includes('tnf')) return 'tnf';
                if (custRaw.includes('columbia') && custRaw.length > 0) return 'columbia';
                if (custRaw.includes('arcteryx') || custRaw.includes("arc'teryx")) return 'arcteryx';
                const suppRaw = (getVal('vendorName') || getVal('productSupplier') || '').toLowerCase();
                if (suppRaw.includes('pt uwu jump - jw') || suppRaw.includes('jack wolfskin')) return 'jack wolfskin';
                if (suppRaw.includes('madison 88')) return 'tnf';
                if (suppRaw.includes('uwu jump')) return looksLikePeakPerformanceSheet || looksLikePeakPerformanceBuy ? 'peak performance' : 'tnf';
                if (suppRaw.includes('llb') || suppRaw.includes('jaytex') || suppRaw.includes('ll bean')) return 'll bean';
                return '';
            })()));
            let brandKey = looksLikePranaSource
                ? 'prana'
                : ((looksLikePeakPerformanceSheet || looksLikePeakPerformanceBuy)
                    ? 'peak performance'
                    : (inferredBrand || brand || '').trim().toLowerCase());
            if (brandKey === 'col') brandKey = 'columbia';
            const isHHBrand = brandKey === 'hh' || brandKey === 'helly hansen';
            const poNumberRaw = manualPurchaseOrder || rawFilePo || (brandKey === 'rossignol' ? rossignolBuyToken : '');
            if (!poNumberRaw) return;
            if (isHHBrand && headerMap['whs'] && !headerMap['plantName']) {
                headerMap['plantName'] = headerMap['whs'];
            }
            const rawPlant = this.stripBrackets(getVal('plant') || '');
            const plant = brandKey === 'vans' ? this.normalizeVansPlantCode(rawPlant) : rawPlant;
            const plantName = this.stripBrackets(getVal('plantName') || '');
            const warehouseName = this.stripBrackets(getVal('warehouseName') || '');
            const whsCode = this.stripBrackets(getVal('whs') || '');
            const customerNameRaw = getVal('customerName');
            const manualCustomerName = this.stripBrackets(manualCustomer || '');
            const hhCompanyNameRaw = getVal('hhCompanyName');
            const ultimateDestinationRaw = this.stripBrackets(getVal('ultimateDestination') || '');
            const shipToRaw = this.stripBrackets(getVal('shipTo') || '');
            const hhPlantSource = rawPlant || plantName || whsCode;
            const plantCountryKey = rawPlant.toLowerCase() || plantName.toLowerCase() || whsCode.toLowerCase();
            const plantDerivedCountry = PLANT_COUNTRY_MAP[plantCountryKey] !== undefined
                ? PLANT_COUNTRY_MAP[plantCountryKey]
                : (PLANT_COUNTRY_MAP[plantName.toLowerCase()] || '');
            const hasDestinationColumn = !!headerMap['transportLocation'];
            const destinationFromFile = this.stripBrackets(getVal('transportLocation') || '');
            const onAgDestinationName = this.stripBrackets(getVal('onAgDestinationName') || '');
            const arcteryxDestinationName = this.stripBrackets(getVal('arcteryxDestinationName') || '');
            const vuoriDestinationName = this.stripBrackets(getVal('vuoriDestinationName') || '');
            const hunterPackingSplit = this.stripBrackets(getVal('hunterPackingSplit') || '');
            const hhDestinationSource = destinationFromFile || shipToRaw || manualDestination || plantDerivedCountry;
            const destCountryRaw = this.stripBrackets(
                isHHBrand
                    ? hhDestinationSource
                    : (manualDestination || destinationFromFile || plantDerivedCountry)
            );
            const destCountry = destCountryRaw ? (COUNTRY_NAME_MAP[destCountryRaw.toUpperCase()] || destCountryRaw) : '';
            const onAgCountryToken = brandKey === 'on ag'
                ? this.resolveOnAgCountryToken(destinationFromFile || manualDestination || plantDerivedCountry)
                : '';
            const onAgDestinationCode = brandKey === 'on ag'
                ? this.extractOnAgDestinationCode(onAgDestinationName)
                : '';
            const burtonDestination = brandKey === 'burton'
                ? this.normalizeTransportLocation(destinationFromFile || manualDestination || plantDerivedCountry)
                : '';
            const rossignolDestinationSource = destinationFromFile || manualDestination || plantDerivedCountry;
            const rossignolPoSuffix = brandKey === 'rossignol'
                ? this.resolveRossignolDestinationSuffix(rossignolDestinationSource)
                : '';
            const vansPoSuffix = brandKey === 'vans' ? this.normalizeVansPoSuffix(customerNameRaw) : '';
            const llbDestination = brandKey === 'll bean'
                ? this.normalizeTransportLocation(destinationFromFile || shipToRaw || manualDestination || plantDerivedCountry)
                : '';
            const llbDestinationLabel = brandKey === 'll bean'
                ? (() => {
                    const shipKey = shipToRaw.toLowerCase().trim();
                    if (shipKey.includes('canada')) return 'Jaytex (Canada)';
                    if (shipKey.includes('usa') || shipKey.includes('united states')) return 'USA';
                    return this.stripBrackets(destinationFromFile || shipToRaw || manualDestination || plantDerivedCountry || '').trim();
                })()
                : '';
            const hhDestinationCountry = isHHBrand
                ? this.resolveHhDestinationCountry(hhCompanyNameRaw || customerNameRaw, shipToRaw, manualDestination, plantDerivedCountry)
                : '';
            const hhPoSuffix = isHHBrand
                ? [hhDestinationCountry || destCountry || hhDestinationSource].filter(Boolean)
                : [];
            const hhStartDateRaw = isHHBrand ? (getRawVal('hhFinalXfDate') || getRawVal('finalXfDate') || getRawVal('exFtyDate') || getRawVal('confirmedExFac') || '') : '';
            const hhCancelDateRaw = isHHBrand ? (getRawVal('hhConfirmedDeliveryDate') || getRawVal('confirmedExFac') || getRawVal('cancelDate') || '') : '';
            const jwsStartDateRaw = brandKey === 'jack wolfskin'
                ? (getRawVal('confirmedExFac') || getRawVal('finalDeliveryDate') || getRawVal('exFtyDate') || getRawVal('cancelDate') || '')
                : '';
            const jwsCancelDateRaw = brandKey === 'jack wolfskin'
                ? (getRawVal('confirmedExFac') || getRawVal('finalDeliveryDate') || getRawVal('exFtyDate') || getRawVal('cancelDate') || '')
                : '';
            const hhStartDate = isHHBrand ? this.formatDateString(hhStartDateRaw as any) : '';
            const hhCancelDate = isHHBrand ? this.formatDateString(hhCancelDateRaw as any) : '';
            const pranaPlantLabel = brandKey === 'prana'
                ? this.extractPranaPlantLabel(shipToRaw || plantName || plant || whsCode || manualDestination)
                : '';
            const vansPlantLabel = brandKey === 'vans'
                ? this.normalizeVansPlantLabel(plantName || plant || whsCode, plant)
                : '';
            const columbiaUltimateDestination = brandKey === 'columbia'
                ? this.stripBrackets(ultimateDestinationRaw || '').trim()
                : '';
            const columbiaDestinationCountry = brandKey === 'columbia'
                ? (this.normalizeTransportLocation(destCountryRaw || destinationFromFile || plantDerivedCountry || '') || this.stripBrackets(destCountryRaw || '').trim())
                : '';
            const pranaDestinationCountry = brandKey === 'prana'
                ? this.resolvePranaDestinationCountry([destinationFromFile, manualDestination, plantDerivedCountry, shipToRaw])
                : '';
            const poPlantPart = brandKey === 'prana'
                ? pranaPlantLabel
                : (brandKey === 'columbia'
                    ? columbiaUltimateDestination
                    : (brandKey === 'vans'
                        ? vansPlantLabel
                        : (isHHBrand
                            ? (this.normalizeHhPlantCode(whsCode || plant || ''))
                            : (whsCode || plant || plantName))));
            let poDestination = [
                destCountry,
                destinationFromFile,
                shipToRaw,
                manualDestination,
                plantDerivedCountry,
            ].find(candidate => this.isLikelyDestinationCountry(candidate, poPlantPart)) || '';
            if (brandKey === 'hunter') {
                const hunterPoDestination = this.normalizeHunterOrderTransportLocation(hunterPackingSplit, poNumberRaw)
                    || this.normalizeHunterTransportLocation(destinationFromFile || manualDestination || plantDerivedCountry, hunterPackingSplit, poNumberRaw);
                if (hunterPoDestination) {
                    poDestination = hunterPoDestination;
                }
            } else if (brandKey === 'jack wolfskin' && !poDestination) {
                poDestination = 'Germany';
            } else if (brandKey === 'prana') {
                poDestination = pranaDestinationCountry;
            } else if (brandKey === 'columbia') {
                poDestination = columbiaDestinationCountry;
            } else if (isHHBrand && plantDerivedCountry) {
                poDestination = plantDerivedCountry;
            }
            let poNumber = this.formatPurchaseOrder(poNumberRaw, poPlantPart, poDestination);
            const manualDestinationEffective = manualDestination;

            const categoryRaw = this.stripBrackets(getVal('category') || '');
            const inferredCat = categoryRaw || this.inferCategoryFromFactoryMap(brand, factoryMap);
            const productExternalRef = '';
            const productCustomerRef = this.stripBrackets(getVal('productCustomerRef') || '');
            const inlineSizeName = this.stripBrackets(getVal('inlineSizeName') || '').trim();
            const sizeRaw = this.stripBrackets(brandKey === 'burton' ? (inlineSizeName || getVal('sizeName') || '') : (getVal('sizeName') || ''));
            const size = this.normalizeSizeName(sizeRaw, inferredBrand || brand);

            const rowQty = this.parseLooseNumber(getVal('finalQty') || getVal('quantity') || getVal('grandTotal') || '0');
            let qty = rowQty;
            if ((!Number.isFinite(qty) || qty <= 0) && brandKey === 'hunter') {
                const normalizedSize = size.toUpperCase().trim();
                const hunterSizeBucketMap: Record<string, string> = {
                    'XS': 'hunterQtyXS',
                    'S': 'hunterQtyS',
                    'M': 'hunterQtyM',
                    'L': 'hunterQtyL',
                    'XL': 'hunterQtyXL',
                };
                const bucketField = hunterSizeBucketMap[normalizedSize];
                if (bucketField) {
                    const bucketRaw = getRawVal(bucketField);
                    const bucketQty = this.parseLooseNumber(bucketRaw?.toString().trim() || '0');
                    if (Number.isFinite(bucketQty) && bucketQty > 0) qty = bucketQty;
                }
            }
            const hasHunterSizeBuckets = !!(headerMap['hunterQtyXS'] || headerMap['hunterQtyS'] || headerMap['hunterQtyM'] || headerMap['hunterQtyL'] || headerMap['hunterQtyXL']);
            if (!headerMap['quantity'] && options?.defaultQuantityIfMissing && !(brandKey === 'hunter' && hasHunterSizeBuckets) && !effectivePivotFormat.isPivotFormat) {
                qty = 1;
                if (!warnedDefaultQty) { warnedDefaultQty = true; this.errors.push({ field: 'quantity', row: 1, message: "Quantity column missing. Defaulting Quantity=1 for all rows.", severity: 'WARNING' }); }
            }
            const looksLikePranaRow = brandKey === 'prana'
                || /prana/i.test(customerNameRaw || '')
                || /prana/i.test(detectedCustomer || '')
                || /prana/i.test(getVal('shipTo') || '');
            if (looksLikePranaRow && qty <= 0) return;

            const foxMaterialCode = this.stripBrackets(getVal('foxMaterialCode') || '').trim();
            const foxMaterialDescription = this.stripBrackets(getVal('foxMaterialDescription') || '').trim();
            const jdeStyle = this.stripBrackets(getVal('jdeStyle') || '').trim();
            const productField = this.stripBrackets(getVal('product') || '').trim();
            const rossignolM88Ref = this.stripBrackets(getVal('rossignolM88Ref') || '').trim();
            const rossignolProductCode = this.stripBrackets(getVal('rossignolProductCode') || '').trim();
            const productAltField = this.stripBrackets(getVal('productAlt') || '').trim();
            const onAgBuyerItem = this.stripBrackets(getVal('onAgBuyerItem') || '').trim();
            const onAgProductName = this.stripBrackets(getVal('onAgProductName') || '').trim();
            const inlineProductName = this.stripBrackets(getVal('inlineProductName') || '').trim();
            const inlineColorName = this.stripBrackets(getVal('inlineColorName') || '').trim();
            const inlineColorDescription = this.stripBrackets(getVal('inlineColorDescription') || '').trim();
            const inlineStyleColor = this.stripBrackets(getVal('inlineStyleColor') || '').trim();
            const marmotShortText = this.stripBrackets(getVal('marmotShortText') || '').trim();
            const marmotStyle = this.stripBrackets(getVal('marmotStyle') || '').trim();
            const ourReference = this.stripBrackets(getVal('ourReference') || '').trim();
            const inlineFactory = this.stripBrackets(getVal('inlineFactory') || '').trim();
            const jwsPlantCode = this.stripBrackets(
                (inlineFactory || poPlantPart)
                    .replace(/\bPT\.?\s*UWU\s*JUMP\b/ig, '')
                    .replace(/\bPT\.?\s*UWU\b/ig, '')
                    .replace(/\s*-\s*/g, '')
            ).trim();
            if (brandKey === 'jack wolfskin') {
                const jwsDestinationCode = this.stripBrackets(poDestination || 'Germany').trim() || 'Germany';
                // JWS policy: drop the JW plant code from the PO and keep only destination suffix.
                poNumber = `${this.stripBrackets(poNumberRaw || '').trim()}-${jwsDestinationCode}`;
            }
            const rawColour = this.stripBrackets(getVal('colour') || '').trim();
            const colourNameRaw = this.stripBrackets(getVal('colourName') || '').trim();
            const pranaColorCodeRaw = this.stripBrackets(getVal('pranaColorCode') || '').trim();
            const pranaColorTextRaw = this.stripBrackets(getVal('pranaColorText') || '').trim();
            const peakPerformanceColourSource = this.stripBrackets(getVal('inlineColorName') || colourNameRaw || rawColour).trim();
            const peakPerformanceColourCandidates = this.getPeakPerformanceColourCandidates(peakPerformanceColourSource);
            const peakPerformanceColourKey = peakPerformanceColourCandidates[0]?.toLowerCase().trim() || '';
            const colour = brandKey === 'vuori'
                ? (inlineColorDescription || colourNameRaw || rawColour)
                : brandKey === 'fox racing'
                ? (this.extractFoxBracketedColour(foxMaterialDescription || colourNameRaw || rawColour) || foxMaterialDescription || colourNameRaw || rawColour)
                : brandKey === 'marmot'
                ? (marmotShortText || rawColour)
                : brandKey === 'prana'
                ? (pranaColorTextRaw ? `PRA- ${pranaColorCodeRaw || rawColour} ${pranaColorTextRaw}`.replace(/\s+/g, ' ').trim() : (pranaColorCodeRaw || rawColour))
                : brandKey === 'rossignol'
                ? (inlineColorName || colourNameRaw || rawColour)
                : brandKey === 'peak performance'
                ? (peakPerformanceColourSource || colourNameRaw || rawColour)
                : isHHBrand
                ? (inlineColorDescription || inlineColorName || inlineStyleColor || rawColour)
                : rawColour;
            if (!colour && brandKey !== 'marmot') { this.errors.push({ field: 'colour', row: rowNumber, message: `Row ${rowNumber} PO ${poNumber}: colour is empty; line/size skipped.`, severity: 'WARNING' }); return; }
            if (colour && colour.trim().toLowerCase() === 'not set') { this.errors.push({ field: 'colour', row: rowNumber, message: `Row ${rowNumber} PO ${poNumber}: colour is "Not Set"; line/size skipped.`, severity: 'WARNING' }); return; }

            let colourKey = this.normalizeColourKey(colour);
            if (brandKey === 'vuori') {
                colourKey = this.normalizeVuoriColourKey(colour);
            } else if (brandKey === 'jack wolfskin') {
                colourKey = this.normalizeJackWolfskinColourKey(colour) || this.normalizeColourKey(colour);
            } else if (brandKey === 'prana') {
                colourKey = (pranaColorCodeRaw || rawColour || '').toLowerCase().trim();
            } else if (brandKey === 'peak performance') {
                colourKey = peakPerformanceColourKey || '';
            } else if (brandKey === 'll bean') {
                colourKey = this.normalizeLlBeanColourKey(colour);
            } else if (brandKey === 'cotopaxi') {
                colourKey = this.normalizeCotopaxiColourText(colour) || this.normalizeColourKey(colour);
            } else if (brandKey === 'marmot') {
                colourKey = this.normalizeMarmotColourText(colour) || this.normalizeColourKey(colour);
            }
            let effectiveStyle = '';
            if (brandKey === 'on ag') {
                effectiveStyle = onAgBuyerItem || jdeStyle || productField;
            } else if (brandKey === 'arcteryx') {
                effectiveStyle = productCustomerRef || jdeStyle || productField;
            } else if (isHHBrand) {
                effectiveStyle = productCustomerRef || jdeStyle || productField;
            } else if (brandKey === 'jack wolfskin') {
                effectiveStyle = this.normalizeJackWolfskinStyleKey(
                    getVal('inlineStyleColor') || getVal('inlineFactory') || getVal('productCustomerRef') || getVal('jdeStyle') || getVal('product') || ''
                ) || productCustomerRef || jdeStyle || productField;
            } else if (brandKey === 'vuori') {
                // Vuori buy files can expose the M-series in Item and the human style name in Product.
                // Prefer Item via productAlt so PLM matching resolves to the M-series code.
                effectiveStyle = productAltField || productCustomerRef || jdeStyle || productField;
            } else if (brandKey === 'marmot') {
                effectiveStyle = marmotStyle || productAltField || productCustomerRef || jdeStyle || productField;
            } else if (brandKey === 'rossignol') {
                effectiveStyle = rossignolM88Ref || rossignolProductCode || productField || productCustomerRef || jdeStyle;
            } else if (brandKey === 'll bean') {
                effectiveStyle = productCustomerRef || jdeStyle || productField;
            } else if (brandKey === 'fox racing') {
                effectiveStyle = foxMaterialCode || jdeStyle || productField;
            } else if (brandKey === 'peak performance') {
                effectiveStyle = productAltField || productField || jdeStyle;
            } else {
                effectiveStyle = jdeStyle || productField;
            }
            const styleKeyCandidates = effectiveStyle ? this.normalizeStyleKeyCandidates(effectiveStyle) : [];

            let productMatches: ProductSheetRow[] = [];
            let matchedStyleKey = effectiveStyle;
            for (const candidate of styleKeyCandidates) {
                const lk = candidate && colourKey ? `${candidate}|${colourKey}` : '';
                const matches = lk ? (productSheetMap[lk] || []) : [];
                if (matches.length > 0) { productMatches = matches; matchedStyleKey = candidate; break; }
            }

            if (productMatches.length === 0 && styleKeyCandidates.length > 0 && brandKey !== 'vuori') {
                for (const candidate of styleKeyCandidates) {
                    const styleColourCode = this.extractStyleColourCode(candidate);
                    const lk = candidate && styleColourCode ? `${candidate}|${styleColourCode}` : '';
                    const matches = lk ? (productSheetMap[lk] || []) : [];
                    if (matches.length > 0) { productMatches = matches; matchedStyleKey = candidate; break; }
                }
            }

            if (productMatches.length === 0 && colourKey && styleKeyCandidates.length > 0 && brandKey !== 'vuori') {
                for (const candidate of styleKeyCandidates) {
                    const allForStyle = Object.entries(productSheetMap).filter(([k]) => k.startsWith(`${candidate}|`)).flatMap(([, v]) => v);
                    if (allForStyle.length > 0) {
                        const fuzzy = allForStyle.find(e => { const ek = this.normalizeColourKey(e.colour); return ek === colourKey || ek.includes(colourKey) || colourKey.includes(ek); });
                        if (fuzzy) { productMatches = [fuzzy]; matchedStyleKey = candidate; break; }
                    }
                }
            }

            if (productMatches.length === 0 && brandKey === 'marmot' && styleKeyCandidates.length > 0) {
                for (const candidate of styleKeyCandidates) {
                    const allForStyle = Object.entries(productSheetMap).filter(([k]) => k.startsWith(`${candidate}|`)).flatMap(([, v]) => v);
                    if (allForStyle.length > 0) {
                        productMatches = [allForStyle[0]];
                        matchedStyleKey = candidate;
                        break;
                    }
                }
            }

            if (productMatches.length === 0 && brandKey === 'dynafit' && styleKeyCandidates.length > 0) {
                for (const candidate of styleKeyCandidates) {
                    const allForStyle = Object.entries(productSheetMap).filter(([k]) => k.startsWith(`${candidate}|`)).flatMap(([, v]) => v);
                    if (allForStyle.length > 0) {
                        productMatches = [allForStyle[0]];
                        matchedStyleKey = candidate;
                        break;
                    }
                }
            }

            if (productMatches.length > 1) productMatches = [productMatches[0]];

            const hasArcInlineProductData = brandKey === 'arcteryx' && !!(inlineProductName || inlineColorName || inlineFactory);
            const hasBurtonInlineProductData = brandKey === 'burton' && !!(inlineProductName || inlineColorName || inlineFactory);
            const has66NorthInlineProductData = brandKey === '66 degrees north' && !!(inlineProductName || inlineColorName || inlineFactory);
            const hasInlineProductData = hasArcInlineProductData || hasBurtonInlineProductData || has66NorthInlineProductData;
            const hasPlmMap = Object.keys(productSheetMap).length > 0;
            let plmMissing = false;
            if (!effectiveStyle && hasPlmMap && !hasInlineProductData) { this.errors.push({ field: 'PLM', row: rowNumber, message: `Row ${rowNumber} PO ${poNumber}: JDE Style missing; PLM fields left blank.`, severity: 'WARNING' }); plmMissing = true; }
            if (productMatches.length === 0 && !plmMissing && hasPlmMap && !hasInlineProductData) { this.errors.push({ field: 'PLM', row: rowNumber, message: `Row ${rowNumber} PO ${poNumber}: JDE ${effectiveStyle} color ${colour} not found in PLM sheet; PLM fields left blank.`, severity: 'WARNING' }); plmMissing = true; }

            const productMatch = !plmMissing && productMatches.length === 1 ? productMatches[0] : undefined;
            if (brandKey === 'dynafit' && !productMatch) {
                this.errors.push({
                    field: 'PLM',
                    row: rowNumber,
                    message: `Row ${rowNumber} PO ${poNumber}: DROPPED - no PLM match and not in confirmed manual order.`,
                    severity: 'WARNING',
                });
                return;
            }
            if (productMatch && productMatch.colour && productMatch.colour.trim().toLowerCase() === 'not set') {
                this.errors.push({ field: 'Colour', row: rowNumber, message: `Row ${rowNumber} PO ${poNumber}: PLM Color Name is "Not Set"; line/size skipped.`, severity: 'WARNING' }); return;
            }
            let dynafitBuyerPoNumber = '';

            const pivotSizeEntries = effectivePivotFormat.isPivotFormat
                ? effectivePivotFormat.pivotColumns
                    .map(({ colNumber, headerText }) => {
                        const rawPivotQty = this.getCellValue(row.getCell(colNumber));
                        const pivotQty = this.parseLooseNumber(rawPivotQty?.toString().trim() || row.getCell(colNumber).text || '0');
                        return {
                            sizeName: this.normalizeSizeName(headerText, inferredBrand || brand),
                            quantity: pivotQty,
                        };
                    })
                    .filter(entry => Number.isFinite(entry.quantity) && entry.quantity > 0)
                : [];
            const hasPivotSizes = pivotSizeEntries.length > 0;
            const pivotQtyTotal = hasPivotSizes ? pivotSizeEntries.reduce((acc, entry) => acc + (Number.isFinite(entry.quantity) ? entry.quantity : 0), 0) : 0;
            const usePivotSizesForRow = hasPivotSizes && pivotQtyTotal > 0;

            const buyProductName = this.stripBrackets(inlineProductName || getVal('productName') || '').trim();
            const plmProductName = this.stripBrackets(productMatch?.productName || '').trim();
            let styleNumber = plmProductName;
            if (!styleNumber) {
                if (brandKey === 'vans') {
                    styleNumber = this.stripBrackets(getVal('jdeStyle') || buyProductName || getVal('product') || matchedStyleKey || '');
                } else if (brandKey === 'rossignol') {
                    styleNumber = this.stripBrackets(rossignolM88Ref || rossignolProductCode || buyProductName || getVal('product') || matchedStyleKey || '');
                } else if (brandKey === 'on ag') {
                    styleNumber = this.stripBrackets(onAgProductName || '');
                } else if (brandKey === 'fox racing') {
                    styleNumber = this.stripBrackets(foxMaterialCode || buyProductName || getVal('product') || matchedStyleKey || '');
                } else if (brandKey === 'arcteryx') {
                    styleNumber = this.stripBrackets(inlineProductName || buyProductName || getVal('product') || matchedStyleKey || '');
                } else if (brandKey === 'jack wolfskin') {
                    styleNumber = this.stripBrackets(matchedStyleKey || buyProductName || getVal('product') || getVal('jdeStyle') || getVal('productCustomerRef') || getVal('productExternalRef') || '');
                } else if (brandKey === 'vuori') {
                    styleNumber = this.stripBrackets(buyProductName || getVal('product') || getVal('productCustomerRef') || getVal('jdeStyle') || matchedStyleKey || '');
                } else if (brandKey === 'dynafit') {
                    styleNumber = this.stripBrackets(matchedStyleKey || buyProductName || getVal('product') || getVal('jdeStyle') || getVal('productCustomerRef') || '');
                } else if (brandKey === 'll bean') {
                    styleNumber = this.stripBrackets(getVal('productCustomerRef') || buyProductName || getVal('product') || getVal('jdeStyle') || matchedStyleKey || '');
                } else if (isHHBrand) {
                    styleNumber = this.stripBrackets(buyProductName || getVal('product') || getVal('jdeStyle') || matchedStyleKey || '');
                } else if (brandKey === 'burton') {
                    styleNumber = this.stripBrackets(inlineProductName || buyProductName || getVal('product') || matchedStyleKey || '');
                } else if (brandKey === '66 degrees north') {
                    styleNumber = this.stripBrackets(buyProductName || inlineProductName || getVal('product') || matchedStyleKey || '');
                } else if (brandKey === 'prana') {
                    styleNumber = this.stripBrackets(inlineProductName || buyProductName || getVal('product') || getVal('jdeStyle') || matchedStyleKey || '');
                } else if (plmMissing) {
                    styleNumber = this.stripBrackets(buyProductName || getVal('product') || getVal('jdeStyle') || '');
                } else {
                    styleNumber = this.stripBrackets(buyProductName || getVal('product') || getVal('jdeStyle') || matchedStyleKey || '');
                }
            }
            if (process.env.DEBUG_EXPORT_TRACE === '1' && brandKey === 'cotopaxi' && rowNumber <= 3) {
                // Cotopaxi trace: helps verify whether PLM matching or fallback is producing the exported Product.
                console.log('[cotopaxi-trace]', {
                    rowNumber,
                    buyProductName,
                    productField,
                    jdeStyle,
                    colour,
                    colourKey,
                    effectiveStyle,
                    matchedStyleKey,
                    productMatchName: plmProductName,
                    finalStyleNumber: styleNumber,
                    poNumber,
                });
            }

            const foxSeasonFromStyle = brandKey === 'fox racing'
                ? this.inferFoxSeasonFromStyle(foxMaterialCode || productField || jdeStyle)
                : '';
            const foxSeasonDateSource = (getRawVal('exFtyDate') || getRawVal('confirmedExFac')) as string | number | Date | undefined;
            const foxSeasonFromDate = brandKey === 'fox racing'
                ? this.inferFoxSeasonFromDate(foxSeasonDateSource)
                : '';
            const arcteryxSeasonFromDate = brandKey === 'arcteryx'
                ? this.inferArcteryxSeason((getRawVal('exFtyDate') || getRawVal('confirmedExFac')) as Date | string | undefined)
                : '';
            const hunterSeasonRaw = this.stripBrackets(getVal('season') || '');
            const hunterEffectiveSeason = brandKey === 'hunter'
                ? ((hunterSeasonRaw && !/^AW\d{2}[_-]/i.test(hunterSeasonRaw)) ? hunterSeasonRaw : (seasonOverride || inferredSeasonFromSheet || hunterSeasonRaw))
                : '';
            const season = this.stripBrackets(
                brandKey === 'hunter'
                    ? (hunterEffectiveSeason || foxSeasonFromStyle || arcteryxSeasonFromDate)
                    : (getVal('season') || foxSeasonFromStyle || foxSeasonFromDate || arcteryxSeasonFromDate || seasonOverride || inferredSeasonFromSheet)
            );
            if (!season) { skippedMissingSeason += 1; this.errors.push({ field: 'season', row: rowNumber, message: `Row ${rowNumber} PO ${poNumber}: No season/range value found.`, severity: 'CRITICAL' }); return; }

            if (usePivotSizesForRow) qty = pivotQtyTotal;
            const rossignolDestinationRaw = manualDestinationEffective || destinationFromFile || plantDerivedCountry;
            const vendorNameRaw = this.stripBrackets(getVal('vendorName') || '');
            const pranaTransportLocation = (() => {
                if (brandKey !== 'prana') return '';
                const shipToKey = shipToRaw.toLowerCase();
                if (shipToKey.includes('united states') || shipToKey.includes(' us ') || shipToKey.includes('us warehouse') || shipToKey.includes('prana us warehouse')) return 'USA';
                return manualDestination || getVal('transportLocation') || plantDerivedCountry || '';
            })();
            const llbTransportLocation = (() => {
                if (brandKey !== 'll bean') return '';
                return this.normalizeTransportLocation(destinationFromFile || shipToRaw || manualDestination || plantDerivedCountry);
            })();
            const jwsTransportLocation = (() => {
                if (brandKey !== 'jack wolfskin') return '';
                return 'Germany';
            })();
            let transportLocationSource: string = manualDestination || getVal('transportLocation') || plantDerivedCountry || '';
            if (brandKey === 'vans') {
                transportLocationSource = manualDestination || plantDerivedCountry || getVal('transportLocation') || '';
            } else if (brandKey === 'on ag') {
                transportLocationSource = this.normalizeOnAgTransportLocation(manualDestinationEffective || destinationFromFile || plantDerivedCountry);
            } else if (brandKey === 'prana') {
                transportLocationSource = pranaTransportLocation;
            } else if (brandKey === 'll bean') {
                transportLocationSource = llbTransportLocation;
            } else if (brandKey === 'dynafit') {
                transportLocationSource = this.normalizeTransportLocation(destinationFromFile || manualDestination || plantDerivedCountry || shipToRaw || 'Germany') || 'Germany';
            } else if (brandKey === 'jack wolfskin') {
                transportLocationSource = jwsTransportLocation;
            } else if (brandKey === 'hunter') {
                transportLocationSource = this.normalizeHunterTransportLocation(manualDestination || destinationFromFile || plantDerivedCountry, hunterPackingSplit, poNumberRaw);
            } else if (brandKey === 'rossignol') {
                transportLocationSource = (((rossignolDestinationRaw || '').trim().toUpperCase() === 'EU') ? 'France' : rossignolDestinationRaw);
            } else if (isHHBrand) {
                transportLocationSource = hhDestinationCountry || hhDestinationSource;
            }
            const transportLocation = this.normalizeTransportLocation(transportLocationSource);
            const effectiveTransportLocation = brandKey === '66 degrees north'
                ? (transportLocation || 'Iceland')
                : transportLocation;
            if (brandKey === '66 degrees north' && !poDestination) {
                poDestination = effectiveTransportLocation || plantDerivedCountry || 'Iceland';
                poNumber = this.formatPurchaseOrder(poNumberRaw, poPlantPart, poDestination);
            }
            const dynafitContext = brandKey === 'dynafit'
                ? this.resolveDynafitExportContext({
                    poNumberRaw,
                    plantPart: whsCode || plant || plantName,
                    rawFilePo: rawFilePo || '',
                    buyerPoNumber: poNumberRaw,
                    productMatch,
                    destinationFromFile,
                    plantDerivedCountry,
                    shipToRaw,
                    transportLocationSource,
                    effectiveTransportLocation,
                    getRawVal,
                    productSupplierFallback: BRAND_SUPPLIER_MAP['dynafit'],
                })
                : undefined;
            const dynafitExportPurchaseOrder = brandKey === 'dynafit' ? (dynafitContext?.exportPurchaseOrder || poNumber) : poNumber;
            const hunterLineTransportLocation = brandKey === 'hunter'
                ? this.stripBrackets(getVal('transportLocation') || '').trim()
                : '';
            const buyDate = getVal('buyDate');
            const buyRound = this.stripBrackets(getVal('buyRound') || '');
            const pranaDateSource = getRawVal('pranaCrd') || getRawVal('exFtyDate') || getRawVal('confirmedExFac');
            const dynafitCrdRaw = brandKey === 'dynafit'
                ? (dynafitContext?.crd || getRawVal('crd') || getRawVal('dynafitLineKeyDate') || getRawVal('finalDeliveryDate'))
                : undefined;
            const dynafitExFactoryRaw = brandKey === 'dynafit'
                ? (dynafitContext?.exFactory || getRawVal('exFactory') || getRawVal('confirmedExFac') || getRawVal('exFtyDate'))
                : undefined;
            const exFtyDate = (() => {
                if (brandKey === 'vans') {
                    return getRawVal('vansConfirmedVendorCrd') || getRawVal('vansBrandRequestedCrd') || getRawVal('exFtyDate') || getRawVal('confirmedExFac');
                }
                if (brandKey === 'prana') return pranaDateSource;
                if (brandKey === 'dynafit') return dynafitCrdRaw || dynafitExFactoryRaw || getRawVal('exFtyDate') || getRawVal('confirmedExFac');
                if (isHHBrand) return getRawVal('finalXfDate') || getRawVal('exFtyDate') || getRawVal('confirmedExFac');
                return getRawVal('exFtyDate') || getRawVal('confirmedExFac');
            })() as Date | string | undefined;
            if (brandKey === 'prana' && typeof pranaDateSource === 'string' && pranaDateSource.includes('#')) return;
            if (brandKey === 'dynafit' && !dynafitCrdRaw) {
                this.errors.push({ field: 'DeliveryDate', row: rowNumber, message: `Row ${rowNumber} PO ${poNumber}: Dynafit CRD missing; falling back to buy-file ETD/EX. Factory.`, severity: 'WARNING' });
            }
            if (!exFtyDate) { this.errors.push({ field: 'exFtyDate', row: rowNumber, message: `Row ${rowNumber} PO ${poNumber}: exFtyDate is empty.`, severity: 'WARNING' }); }
            const llbFinalDeliveryDateCell = brandKey === 'll bean' ? (getRawVal('confirmedExFac') || getRawVal('finalDeliveryDate') || getRawVal('cancelDate')) : '';
            const llbFinalDeliveryDateRaw = brandKey === 'll bean'
                ? (
                    this.parseDate(llbFinalDeliveryDateCell as any)
                        ? llbFinalDeliveryDateCell
                        : (this.shiftDate(exFtyDate, 7) || llbFinalDeliveryDateCell || exFtyDate || '')
                )
                : '';

            const cancelDate = (isHHBrand
                ? (getRawVal('finalDeliveryDate') || getRawVal('confirmedExFac') || exFtyDate || '')
                : (brandKey === 'jack wolfskin'
                    ? (jwsCancelDateRaw || jwsStartDateRaw || getRawVal('cancelDate') || exFtyDate || '')
                    : (brandKey === 'll bean'
                        ? (llbFinalDeliveryDateRaw || exFtyDate || '')
                        : (getRawVal('cancelDate') || exFtyDate || '')))
            ) as Date | string;
            const poIssuanceDate = getVal('poIssuanceDate') || buyDate || exFtyDate || '';
            const statusRaw = this.normalizeStatus(getVal('status'), inferredBrand || brand);
            const transportRaw = this.stripBrackets(getVal('transportMethod') || '');
            const templateRaw = this.stripBrackets(getVal('template') || '');
            const vendorCodeRaw = this.stripBrackets(plmMissing ? (getVal('productSupplier') || '') : (productMatch?.factory || getVal('productSupplier') || ''));
            const buyerPoNumberCell = getRawVal('buyerPoNumber');
            const buyerPoNumber: string | number = (() => {
                const poRaw = getRawVal('purchaseOrder');
                if (brandKey === 'vans' && typeof buyerPoNumberCell === 'number') return buyerPoNumberCell;
                if (brandKey === 'll bean') {
                    const buyerPoText = buyerPoNumberCell?.toString().trim();
                    if (buyerPoText) return buyerPoText;
                    if (typeof rawFilePo === 'number') return rawFilePo;
                    const rawFilePoText = rawFilePo?.toString().trim();
                    if (rawFilePoText) return rawFilePoText;
                }
                if (brandKey === 'vans') {
                    const vansBuyerPo = buyerPoNumberCell?.toString().trim();
                    if (vansBuyerPo) return vansBuyerPo;
                }
                if (brandKey === 'rossignol') {
                    const buyerPoText = buyerPoNumberCell?.toString().trim();
                    if (buyerPoText) return buyerPoText;
                    const compactRange = this.compactProductRange(this.formatProductRange(season));
                    const suffix = rossignolPoSuffix || this.resolveRossignolDestinationSuffix(transportLocation);
                    if (compactRange && suffix) return `M88 ${poNumberRaw} ${compactRange} ROS-${suffix}-TBA`;
                    if (compactRange) return `M88 ${poNumberRaw} ${compactRange} ROS-TBA`;
                }
                if (brandKey === 'on ag') {
                    if (typeof rawFilePo === 'number') return rawFilePo;
                    const rawBuyerPo = (rawFilePo || '').toString().trim();
                    if (rawBuyerPo) return rawBuyerPo;
                    if (typeof buyerPoNumberCell === 'number') return buyerPoNumberCell;
                    const buyerPoText = buyerPoNumberCell?.toString().trim();
                    if (buyerPoText) return buyerPoText;
                }
                if (brandKey === 'prana') {
                    if (typeof buyerPoNumberCell === 'number') return buyerPoNumberCell;
                    const buyerPoText = buyerPoNumberCell?.toString().trim();
                    if (buyerPoText) return buyerPoText;
                }
                if (brandKey === 'arcteryx') {
                    const arcteryxBuyerPo = getRawVal('arcteryxBuyerPo');
                    if (typeof arcteryxBuyerPo === 'number') return arcteryxBuyerPo;
                    const arcteryxBuyerPoText = arcteryxBuyerPo?.toString().trim();
                    if (arcteryxBuyerPoText) return arcteryxBuyerPoText;
                    if (typeof rawFilePo === 'number') return rawFilePo;
                    const rawFilePoText = rawFilePo?.toString().trim();
                    if (rawFilePoText && !/^p\d{4,}-/i.test(rawFilePoText) && !/^po\d{4,}-/i.test(rawFilePoText)) return rawFilePoText;
                    if (typeof buyerPoNumberCell === 'number') return buyerPoNumberCell;
                    const buyerPoText = buyerPoNumberCell?.toString().trim();
                    if (buyerPoText) return buyerPoText;
                }
                if (typeof poRaw === 'number') return poRaw;
                if (poRaw?.toString().trim()) return poRaw.toString().trim();
                if (typeof buyerPoNumberCell === 'number') return buyerPoNumberCell;
                const asText = buyerPoNumberCell?.toString().trim();
                if (asText) return asText;
                return '';
            })();
            dynafitBuyerPoNumber = brandKey === 'dynafit'
                ? (rawFilePo || buyerPoNumber?.toString?.().trim?.() || poNumberRaw || '')
                : (buyerPoNumber?.toString?.().trim?.() || String(buyerPoNumber || ''));

            const productSupplier = brandKey === 'on ag'
                ? (this.stripBrackets(productMatch?.factory || '').trim() || BRAND_SUPPLIER_MAP['on ag'])
                : brandKey === 'arcteryx'
                    ? BRAND_SUPPLIER_MAP['arcteryx']
                    : brandKey === 'll bean'
                        ? (this.stripBrackets(productMatch?.factory || '').trim() || BRAND_SUPPLIER_MAP['ll bean'])
                    : brandKey === 'dynafit'
                        ? (this.stripBrackets(productMatch?.factory || '').trim() || BRAND_SUPPLIER_MAP['dynafit'])
                    : brandKey === 'fox racing'
                        ? (this.stripBrackets(getVal('goodsSupplierName') || '').trim() || BRAND_SUPPLIER_MAP['fox racing'])
                    : brandKey === 'burton'
                        ? (this.stripBrackets(inlineFactory || productMatch?.factory || '').trim() || BRAND_SUPPLIER_MAP['burton'])
                        : brandKey === '66 degrees north'
                            ? (this.stripBrackets(inlineFactory || productMatch?.factory || '').trim() || BRAND_SUPPLIER_MAP['66 degrees north'])
                            : brandKey === 'prana'
                                ? (this.stripBrackets(inlineFactory || productMatch?.factory || '').trim() || BRAND_SUPPLIER_MAP['prana'])
                                : this.resolveSupplier(vendorCodeRaw, vendorNameRaw, inferredBrand || brand, inferredCat, factoryMap);
            let resolvedColour = colour;
            if (brandKey === 'vans' || brandKey === 'rossignol') {
                resolvedColour = colour;
            } else if (brandKey === 'fox racing') {
                const foxColourFromDesc = this.extractFoxBracketedColour(foxMaterialDescription || colour);
                resolvedColour = foxColourFromDesc || (plmMissing ? colour : (productMatch?.colour || colour));
            } else if (brandKey === 'arcteryx') {
                resolvedColour = plmMissing ? (inlineColorName || colour) : (inlineColorName || productMatch?.colour || colour);
            } else if (brandKey === 'burton') {
                resolvedColour = plmMissing ? (inlineColorName || colour) : (inlineColorName || productMatch?.colour || colour);
            } else if (brandKey === '66 degrees north') {
                resolvedColour = plmMissing ? (inlineColorName || colour) : (inlineColorName || productMatch?.colour || colour);
            } else if (brandKey === 'prana') {
                resolvedColour = plmMissing ? (inlineColorName || colour) : (inlineColorName || productMatch?.colour || colour);
            } else if (brandKey === 'jack wolfskin') {
                resolvedColour = productMatch?.colour || colour;
            } else if (brandKey === 'vuori') {
                resolvedColour = productMatch?.colourName || productMatch?.colour || colour;
            } else if (brandKey === 'll bean') {
                resolvedColour = productMatch?.colour || colour;
            } else if (brandKey === 'marmot') {
                resolvedColour = productMatch?.colourName || productMatch?.colour || colour;
            } else if (brandKey === 'peak performance') {
                resolvedColour = plmMissing
                    ? (peakPerformanceColourSource || colour)
                    : (productMatch?.colourName || productMatch?.colour || peakPerformanceColourSource || colour);
            } else if (brandKey === 'dynafit') {
                resolvedColour = productMatch?.productName || productMatch?.colour || colour;
            } else {
                resolvedColour = plmMissing ? colour : (productMatch?.colour || colour);
            }
            const customerNameForResolve = manualCustomerName || customerNameRaw || (inferredBrand ? (BRAND_CUSTOMER_MAP[(inferredBrand).toLowerCase()] || '') : '') || (brand ? (BRAND_CUSTOMER_MAP[brand.toLowerCase()] || '') : '');
            const customerName = brandKey === 'prana'
                ? 'Prana'
                : brandKey === 'peak performance'
                ? 'Peak Performance'
                : plmMissing
                ? this.resolveCustomer(customerNameForResolve, inferredBrand || brand, detectedCustomer, undefined)
                : this.resolveCustomer(((brandKey === 'arcteryx' || brandKey === 'burton' || brandKey === '66 degrees north') ? (productMatch?.customerName || manualCustomerName || customerNameRaw) : productMatch?.customerName) || manualCustomerName || customerNameRaw, inferredBrand || brand, detectedCustomer, undefined);
            const llbCustomerName = brandKey === 'll bean' ? 'LL Bean' : customerName;

            const transportMethod = brandKey === 'dynafit'
                ? (dynafitContext?.transportMethod || 'Courier')
                : (brandKey === '66 degrees north'
                ? (transportRaw ? this.normalizeTransportMethod(transportRaw) : 'Air')
                : (brandKey === 'cotopaxi' && transportRaw.trim().toLowerCase() === 'international distributor'
                ? 'Courier'
                : this.normalizeTransportMethod(transportRaw)));
            const brandConfig = mloMap.find((m: any) => (m.brand || '').trim().toLowerCase() === brandKey);
            const ordersTemplate = brandKey === 'dynafit'
                ? (dynafitContext?.ordersTemplate || 'SMS PO Header')
                : (manualTemplate || brandConfig?.orders_template?.trim() || this.resolveOrdersTemplate(inferredBrand || brand, templateRaw));
            const linesTemplateBase = manualLinesTemplate || brandConfig?.lines_template?.trim() || this.resolveLinesTemplate(inferredBrand || brand, templateRaw);
            const hunterTemplateDate = this.formatIsoDateString(exFtyDate || '');
            const linesTemplate = brandKey === 'hunter'
                ? (hunterTemplateDate && !linesTemplateBase.includes(hunterTemplateDate)
                    ? `${linesTemplateBase} ${hunterTemplateDate}`.trim()
                    : linesTemplateBase)
                : linesTemplateBase;
            const productRange = brandKey === 'dynafit'
                ? (dynafitContext?.productRange || 'FH:2027')
                : this.formatProductRange(season);
            const keyDate = brandKey === 'hunter'
                ? ''
                : (brandKey === 'll bean'
                    ? (manualKeyDate || this.formatDateString(buyDate as any) || poIssuanceDate)
                    : (brandKey === 'dynafit'
                        ? (manualKeyDate || this.formatDateString(this.shiftDate(exFtyDate, -84) || exFtyDate) || poIssuanceDate)
                    : (brandKey === 'jack wolfskin'
                        ? this.resolveJackWolfskinKeyDate(season, manualKeyDate || poIssuanceDate)
                        : (manualKeyDate || poIssuanceDate))))
                ;
            const keyDateFormat: "manual" | "standard" = manualKeyDate ? "manual" : "standard";
            const commentBrand = isHHBrand ? 'HH' : (inferredBrand || brand || detectedCustomer);
            const commentBuyRound = isHHBrand && !buyRound
                ? ((options?.sourceFilename || '').toLowerCase().includes('feb bulk buy') || (options?.sourceFilename || '').toLowerCase().includes('febbuy')
                    ? 'FebBuy'
                    : buyRound)
                : buyRound;

            const customerSubtype = (brandKey === 'burton' || isHHBrand)
                ? undefined
                : this.detectCustomerSubtype(productMatch?.customerName || getVal('customerName') || getVal('brand') || detectedCustomer || '');
            if (brandKey === 'vuori') {
                const vuoriSuffix = warehouseName || plantName || this.stripBrackets(customerNameRaw || customerName || '').trim();
                if (vuoriSuffix && !poNumber.toLowerCase().endsWith(`-${vuoriSuffix.toLowerCase()}`)) {
                    poNumber = `${poNumber}-${vuoriSuffix}`;
                }
            } else if (brandKey === 'prana') {
                const pranaSuffixParts = [pranaPlantLabel, pranaDestinationCountry].filter(Boolean);
                poNumber = poNumberRaw;
                if (pranaSuffixParts.length > 0) {
                    poNumber = `${poNumberRaw} - ${pranaSuffixParts.join(' - ')}`;
                }
            } else if (brandKey === 'columbia') {
                const columbiaBasePo = this.stripBrackets(poNumberRaw || '').trim().split(/\s+-\s+/)[0].trim();
                const columbiaSuffixParts = [columbiaUltimateDestination, columbiaDestinationCountry]
                    .map(part => this.stripBrackets(part || '').trim())
                    .filter(Boolean);
                const uniqueColumbiaSuffixParts = columbiaSuffixParts.filter((part, index) =>
                    columbiaSuffixParts.findIndex(candidate => candidate.toLowerCase() === part.toLowerCase()) === index
                );
                poNumber = columbiaBasePo;
                if (uniqueColumbiaSuffixParts.length > 0) {
                    poNumber = `${columbiaBasePo} - ${uniqueColumbiaSuffixParts.join(' - ')}`;
                }
                if (customerSubtype && !poNumber.toLowerCase().endsWith(` ${customerSubtype.toLowerCase()}`)) {
                    poNumber = `${poNumber} ${customerSubtype}`;
                }
            } else if (brandKey !== 'columbia' && customerSubtype && !poNumber.toLowerCase().endsWith(` ${customerSubtype.toLowerCase()}`)) {
                poNumber = `${poNumber} ${customerSubtype}`;
            }

            poNumber = this.collapseRepeatedPurchaseOrder(poNumber);

            const validStatuses = Array.isArray(brandConfig?.valid_statuses) ? brandConfig!.valid_statuses!.map((s: string) => s.toLowerCase()) : [];
            if (transportMethod && !VALID_TRANSPORT_VALUES.has(transportMethod)) {
                this.errors.push({ field: 'TransportMethod', row: rowNumber, message: `Row ${rowNumber} PO ${poNumber}: unmapped transport "${transportMethod}" — expected Sea, Air, or Courier.`, severity: 'WARNING' });
            }

            if (brandKey === 'evo' && effectivePivotFormat.isPivotFormat) {
                const evoDestinationRow = Math.max(1, headerRowNumber - 2);
                const evoOrRow = Math.max(1, headerRowNumber - 1);
                const evoEntries = effectivePivotFormat.pivotColumns
                    .map(({ colNumber }) => {
                        const rawQty = this.getCellValue(row.getCell(colNumber));
                        const qtyValue = this.parseLooseNumber(rawQty?.toString().trim() || row.getCell(colNumber).text || '0');
                        const destinationLabel = this.stripBrackets(this.getCellValue(worksheet.getRow(evoDestinationRow).getCell(colNumber))?.toString() || '').trim();
                        const orNumber = this.stripBrackets(this.getCellValue(worksheet.getRow(evoOrRow).getCell(colNumber))?.toString() || '').trim();
                        return { qtyValue, destinationLabel, orNumber };
                    })
                    .filter(entry => Number.isFinite(entry.qtyValue) && entry.qtyValue > 0);

                if (evoEntries.length === 0) {
                    return;
                }

                for (const entry of evoEntries) {
                    const evoPoNumber = entry.orNumber || poNumberRaw || poNumber;
                    const evoCustomerKey = customerName || customerNameRaw || detectedCustomer;
                    const evoOrderKey = `${evoPoNumber}||${evoCustomerKey}||${entry.destinationLabel || ''}`;
                    const evoTransportLocation = entry.destinationLabel || effectiveTransportLocation || destinationFromFile || plantDerivedCountry || '';
                    const evoKeyUsers = this.resolveKeyUsers(inferredBrand || brand, manualKeyUser1, manualKeyUser2, manualKeyUser3, manualKeyUser4, manualKeyUser5, getVal('keyUser1'), getVal('keyUser2'), getVal('keyUser4'), getVal('keyUser5'), brandConfig);

                    if (!results.has(evoPoNumber)) {
                        results.set(evoPoNumber, {
                            header: {
                                purchaseOrder: evoPoNumber,
                                brandKey,
                                productSupplier,
                                status: 'Confirmed',
                                customer: customerName,
                                transportMethod,
                                transportLocation: evoTransportLocation,
                                ordersTemplate,
                                linesTemplate,
                                keyDate,
                                keyDateFormat,
                                comments: manualComments || this.buildComments(commentBrand, productRange, commentBuyRound, buyDate, ordersTemplate),
                                currency: 'USD',
                                keyUser1: evoKeyUsers.k1,
                                keyUser2: evoKeyUsers.k2,
                                keyUser3: evoKeyUsers.k3,
                                keyUser4: evoKeyUsers.k4,
                                keyUser5: evoKeyUsers.k5,
                                keyUser6: evoKeyUsers.k6,
                                keyUser7: evoKeyUsers.k7,
                                keyUser8: evoKeyUsers.k8,
                            },
                            lines: [],
                            sizes: {},
                            orderKeys: [],
                            manualKeyDate: manualKeyDate || undefined,
                        });
                    }

                    const evoPo = results.get(evoPoNumber)!;
                    if (!seenOrderKeys.has(evoOrderKey)) {
                        seenOrderKeys.add(evoOrderKey);
                        if (!evoPo.orderKeys) evoPo.orderKeys = [];
                        evoPo.orderKeys.push({
                            purchaseOrder: evoPoNumber,
                            customer: evoCustomerKey,
                            customerName: customerName || customerNameRaw,
                            transportLocation: evoTransportLocation,
                            transportMethod,
                            ordersTemplate
                        });
                    }

                    const evoLineItem = (evoPo.lines.length > 0 ? Math.max(...evoPo.lines.map(l => l.lineItem)) : 0) + 1;
                    evoPo.lines.push({
                        lineItem: evoLineItem,
                        productRange,
                        styleNumber: styleNumber || '',
                        supplierProfile: 'DEFAULT_PROFILE',
                        buyerPoNumber: evoPoNumber,
                        startDate: exFtyDate || '',
                        cancelDate: cancelDate || '',
                        transportLocation: evoTransportLocation,
                        styleColor: inlineStyleColor || undefined,
                        rawColour: colour || undefined,
                        ourReference: ourReference || undefined,
                        cost: undefined,
                        colour: productMatch?.colour || colour,
                        productExternalRef,
                        productCustomerRef,
                    } as POLine);

                    evoPo.sizes[evoLineItem] = [{ productSize: size || 'One Size', quantity: entry.qtyValue }];
                }

                return;
            }

            const missingData: string[] = [];
            if (!styleNumber && !plmMissing) missingData.push('Product/Style');
            if (!size) missingData.push('Size');
            if (isNaN(qty)) missingData.push('Quantity');
            if (!categoryRaw && inferredCat && brand && !warnedInferredCategory.has(brand.toLowerCase())) {
                warnedInferredCategory.add(brand.toLowerCase());
                this.errors.push({ field: 'Mapping', row: rowNumber, message: `Category inferred from factory mapping for Brand: ${brand}`, severity: 'WARNING' });
            }
            if (missingData.length > 0) { this.errors.push({ field: 'Missing Data', row: rowNumber, message: `PO ${poNumberRaw} missing: ${missingData.join(', ')}.`, severity: 'CRITICAL' }); return; }
            if (validStatuses.length > 0 && statusRaw) {
                if (!validStatuses.includes(statusRaw.toLowerCase())) {
                    this.errors.push({ field: 'Status', row: rowNumber, message: `PO ${poNumberRaw} has status "${statusRaw}" not in valid statuses: ${validStatuses.join(', ')}.`, severity: 'WARNING' });
                }
            }

            const mloRow = brandConfig;
            const keyUsers = this.resolveKeyUsers(inferredBrand || brand, manualKeyUser1, manualKeyUser2, manualKeyUser3, manualKeyUser4, manualKeyUser5, getVal('keyUser1'), getVal('keyUser2'), getVal('keyUser4'), getVal('keyUser5'), mloRow);
            const customerKey = llbCustomerName || customerName || detectedCustomer;
            const poKey = poNumber;
            let orderKey = `${poNumber}||${customerKey}`;
            if (isHHBrand) {
                orderKey = `${poNumber}||${customerKey}||${destCountry}`;
            } else if (brandKey === 'll bean') {
                orderKey = `${poNumber}||${customerKey}||${llbDestinationLabel || llbDestination}`;
            } else if (brandKey === 'vuori') {
                orderKey = `${poNumber}||${customerKey}||${vuoriDestinationName || destinationFromFile || plantDerivedCountry}`;
            }
            const hhOrderPurchaseOrder = poNumber;
            const dynafitBuyerPoNumberValue = brandKey === 'dynafit'
                ? (dynafitContext?.buyerPoNumber || rawFilePo || buyerPoNumber?.toString?.().trim?.() || poNumberRaw || '')
                : '';
            const dynafitResolvedColourValue = brandKey === 'dynafit'
                ? (dynafitContext?.resolvedColour || colour)
                : '';

            if (!results.has(poKey)) {
                results.set(poKey, {
                    header: {
                        purchaseOrder: brandKey === 'dynafit' ? dynafitExportPurchaseOrder : poNumber, brandKey, productSupplier, status: 'Confirmed', customer: llbCustomerName || customerName,
                        transportMethod, transportLocation: isHHBrand
                            ? (hhDestinationCountry || destCountry || hhDestinationSource || effectiveTransportLocation)
                            : effectiveTransportLocation, ordersTemplate, linesTemplate, keyDate, keyDateFormat,
                        comments: manualComments || this.buildComments(commentBrand, productRange, commentBuyRound, buyDate, ordersTemplate),
                        currency: 'USD', keyUser1: keyUsers.k1, keyUser2: keyUsers.k2, keyUser3: keyUsers.k3,
                        keyUser4: keyUsers.k4, keyUser5: keyUsers.k5, keyUser6: keyUsers.k6, keyUser7: keyUsers.k7, keyUser8: keyUsers.k8,
                    },
                    lines: [], sizes: {}, orderKeys: [],
                    manualKeyDate: manualKeyDate || undefined,
                });
            }

            const po = results.get(poKey)!;
            if (!seenOrderKeys.has(orderKey)) {
                seenOrderKeys.add(orderKey);
                if (!po.orderKeys) po.orderKeys = [];
                po.orderKeys.push({
                    purchaseOrder: brandKey === 'dynafit' ? dynafitExportPurchaseOrder : hhOrderPurchaseOrder,
                    customer: customerKey,
                    customerName: llbCustomerName || customerName,
                    transportLocation: brandKey === 'hunter'
                        ? this.normalizeHunterOrderTransportLocation(hunterPackingSplit, poNumberRaw)
                        : (isHHBrand ? (hhDestinationCountry || destCountry || hhDestinationSource || effectiveTransportLocation) : effectiveTransportLocation),
                    transportMethod,
                    ordersTemplate
                });
            }

            let lineItemNum = 0;
            const rawLineItem = getRawVal('lineItem');
            if (brandKey === 'vuori') {
                lineItemNum = (po.lines.length > 0 ? Math.max(...po.lines.map(l => l.lineItem)) : 0) + 1;
            } else if (brandKey !== 'cotopaxi' && rawLineItem !== undefined && rawLineItem !== null) {
                const maybe = Number(rawLineItem);
                if (Number.isFinite(maybe) && maybe > 0) lineItemNum = Math.round(maybe);
            }
            if (lineItemNum <= 0) lineItemNum = (po.lines.length > 0 ? Math.max(...po.lines.map(l => l.lineItem)) : 0) + 1;

            let existingLine = po.lines.find(line => line.lineItem === lineItemNum);
            if (!existingLine) {
                const dynafitLineKeyDateRaw = (brandKey === 'dynafit' ? getRawVal('dynafitLineKeyDate') : undefined) as string | Date | undefined;
                existingLine = {
                    lineItem: lineItemNum,
                    productRange,
                    styleNumber: styleNumber || '',
                    supplierProfile: 'DEFAULT_PROFILE',
                    buyerPoNumber: brandKey === 'dynafit' ? dynafitBuyerPoNumberValue : dynafitBuyerPoNumber,
                    dynafitLineKeyDate: brandKey === 'dynafit' ? ((dynafitContext?.lineKeyDate as string | Date | undefined) || dynafitLineKeyDateRaw || exFtyDate || undefined) : undefined,
                    startDate: (isHHBrand ? (hhStartDateRaw || '') : (brandKey === 'll bean' ? (llbFinalDeliveryDateRaw || '') : (brandKey === 'jack wolfskin' ? (jwsStartDateRaw || '') : (brandKey === 'dynafit' ? (dynafitContext?.startDate || dynafitCrdRaw || exFtyDate || '') : (exFtyDate || ''))))) as Date | string,
                    cancelDate: (isHHBrand ? (hhCancelDateRaw || '') : (brandKey === 'll bean' ? (llbFinalDeliveryDateRaw || '') : (brandKey === 'jack wolfskin' ? (jwsCancelDateRaw || '') : (brandKey === 'dynafit' ? (dynafitContext?.cancelDate || dynafitCrdRaw || exFtyDate || '') : (cancelDate || ''))))) as Date | string,
                    hhStartDate: hhStartDate || undefined,
                    hhCancelDate: hhCancelDate || undefined,
                    hhConfirmedDeliveryDate: hhCancelDate || undefined,
                    transportLocation: brandKey === 'hunter'
                        ? hunterLineTransportLocation
                        : (isHHBrand ? (hhDestinationCountry || destCountry || hhDestinationSource || effectiveTransportLocation) : effectiveTransportLocation),
                    styleColor: inlineStyleColor || undefined,
                    rawColour: colour || undefined,
                    ourReference: ourReference || undefined,
                    cost: undefined,
                    colour: brandKey === 'dynafit' ? (dynafitResolvedColourValue || '') : (resolvedColour || ''),
                    productExternalRef: (brandKey === 'arcteryx' || brandKey === 'hunter' || brandKey === 'burton') ? '' : productExternalRef,
                    productCustomerRef: brandKey === 'arcteryx' ? '' : productCustomerRef,
                };
                po.lines.push(existingLine as POLine);
            } else {
                if (styleNumber && existingLine.styleNumber && styleNumber !== existingLine.styleNumber) {
                    this.errors.push({ field: 'LineItem', row: rowNumber, message: `PO ${poNumber} line ${lineItemNum} product mismatch: existing ${existingLine.styleNumber}, row ${styleNumber}.`, severity: 'CRITICAL' });
                }
                if (!existingLine.styleNumber && styleNumber) existingLine.styleNumber = styleNumber;
                if (brandKey !== 'arcteryx' && brandKey !== 'burton') {
                    if (!existingLine.productExternalRef && productExternalRef) existingLine.productExternalRef = productExternalRef;
                    if (!existingLine.productCustomerRef && productCustomerRef) existingLine.productCustomerRef = productCustomerRef;
                }
                if (!existingLine.styleColor && inlineStyleColor) existingLine.styleColor = inlineStyleColor;
                if (!existingLine.rawColour && colour) existingLine.rawColour = colour;
                if (!existingLine.ourReference && ourReference) existingLine.ourReference = ourReference;
                if (isHHBrand) {
                    if (!existingLine.hhStartDate && hhStartDate) existingLine.hhStartDate = hhStartDate;
                    if (!existingLine.hhCancelDate && hhCancelDate) existingLine.hhCancelDate = hhCancelDate;
                    if (!existingLine.hhConfirmedDeliveryDate && hhCancelDate) existingLine.hhConfirmedDeliveryDate = hhCancelDate;
                    if (!existingLine.transportLocation && (hhDestinationCountry || destCountry || hhDestinationSource || effectiveTransportLocation)) {
                        existingLine.transportLocation = hhDestinationCountry || destCountry || hhDestinationSource || effectiveTransportLocation;
                    }
                }
                if (brandKey === 'dynafit' && !existingLine.dynafitLineKeyDate) {
                    existingLine.dynafitLineKeyDate = ((dynafitContext?.lineKeyDate as string | Date | undefined) || (getRawVal('dynafitLineKeyDate') as string | Date | undefined)) || undefined;
                }
                if (brandKey === 'dynafit' && (!existingLine.buyerPoNumber || existingLine.buyerPoNumber === buyerPoNumber)) {
                    existingLine.buyerPoNumber = dynafitBuyerPoNumberValue || dynafitBuyerPoNumber;
                }
            }

            if (!po.sizes[lineItemNum]) po.sizes[lineItemNum] = [];
            if (usePivotSizesForRow) {
                pivotSizeEntries.forEach(entry => {
                    po.sizes[lineItemNum].push({ productSize: entry.sizeName || size || 'One Size', quantity: entry.quantity });
                });
            } else {
                po.sizes[lineItemNum].push({ productSize: size || 'One Size', quantity: qty });
            }
            if (qty <= 0) { this.errors.push({ field: 'Quantity', row: rowNumber, message: `Qty for ${styleNumber} size ${size} is ${qty} (included).`, severity: 'WARNING' }); }
        });

        for (const [poNumber, po] of results.entries()) {
            po.lines.sort((a, b) => a.lineItem - b.lineItem);
            const lineIds = po.lines.map(l => l.lineItem);
            if (lineIds.length > 0) {
                const minLine = Math.min(...lineIds);
                const maxLine = Math.max(...lineIds);
                if (minLine !== 1) this.errors.push({ field: 'LineItem', row: 1, message: `PO ${poNumber} starts at LineItem ${minLine} (should start at 1).`, severity: 'WARNING' });
                for (let expected = minLine; expected <= maxLine; expected++) {
                    if (!lineIds.includes(expected)) this.errors.push({ field: 'LineItem', row: 1, message: `PO ${poNumber} missing LineItem ${expected}.`, severity: 'WARNING' });
                }
            }
            for (const line of po.lines) {
                const sizesForLine = po.sizes[line.lineItem] || [];
                const totalSizeQty = sizesForLine.reduce((acc, s) => acc + (Number.isFinite(s.quantity) ? s.quantity : 0), 0);
                if (sizesForLine.length === 0) this.errors.push({ field: 'LineItem', row: 1, message: `PO ${poNumber} line ${line.lineItem} (${line.styleNumber}) has no sizes attached.`, severity: 'WARNING' });
                if (totalSizeQty <= 0) this.errors.push({ field: 'LineItem', row: 1, message: `PO ${poNumber} line ${line.lineItem} (${line.styleNumber}) has zero total size quantity.`, severity: 'WARNING' });
            }
        }

        const processedData = Array.from(results.values());
        if (skippedMissingSeason > 0) this.errors.push({ field: 'season', row: 1, message: `${skippedMissingSeason} row(s) skipped due to missing season/range.`, severity: 'WARNING' });
        if (processedData.length === 0 && skippedMissingSeason > 0) this.errors.push({ field: 'File Format', row: 1, message: 'No usable rows remain after skipping rows with missing season/range.', severity: 'CRITICAL' });

        const errorCount = this.errors.filter(e => e.severity === 'CRITICAL').length;
        const warningCount = this.errors.filter(e => e.severity === 'WARNING').length;
        if (this.runId) {
            await updateRun(this.runId, { status: errorCount > 0 ? 'Validation Failed' : 'Pending Review', error_count: errorCount, warning_count: warningCount, orders_rows: processedData.length, lines_rows: processedData.reduce((a, p) => a + p.lines.length, 0), order_sizes_rows: processedData.reduce((a, p) => a + Object.values(p.sizes).reduce((b, s) => b + s.length, 0), 0), completed_at: new Date().toISOString() });
            await logEvent({ eventName: errorCount > 0 ? 'VALIDATION_FAILED' : 'VALIDATION_PASSED', userId: this.userId || 'system', runId: this.runId, metadata: { errorCount, warningCount, customer: detectedCustomer } });
        }

        const formatDetection: FormatDetection = {
            detectedCustomer,
            detectedFormat: effectivePivotFormat.isPivotFormat ? 'pivot' : 'standard',
            unmappedColumns: unmappedHeaders
                .map(h => h.headerText)
                .filter(h => !this.shouldSilentlyIgnoreHeader(h))
                .filter(h => !effectivePivotFormat.pivotColumns.some(col => col.headerText === h)),
        };
        if (options?.llBeanReferenceSizesBuffer) {
            await this.applyLlBeanReferenceSizes(processedData, options.llBeanReferenceSizesBuffer);
        }

        return { data: processedData, errors: this.errors, formatDetection };
    }

    private async applyLlBeanReferenceSizes(data: ProcessedPO[], buffer: any): Promise<void> {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);
        const rows = this.extractLlBeanReferenceSizeRows(workbook);
        if (rows.length === 0) return;

        for (const po of data) {
            const brandKey = (po.header.brandKey || '').trim().toLowerCase();
            const customerKey = (po.header.customer || '').trim().toLowerCase();
            if (brandKey !== 'll bean' && customerKey !== 'll bean') continue;
            po.llBeanReferenceSizeRows = rows.map(row => ({ ...row }));
        }
    }

    private extractLlBeanReferenceSizeRows(workbook: ExcelJS.Workbook): Array<{ purchaseOrder: string; lineItem: number; range: string; product: string; sizeName: string; productSize: string; quantity: number; colour: string }> {
        const rows: Array<{ purchaseOrder: string; lineItem: number; range: string; product: string; sizeName: string; productSize: string; quantity: number; colour: string }> = [];
        const aliases = this.getFallbackColumnAliases();
        for (const ws of workbook.worksheets) {
            const headerRowNumber = this.detectHeaderRow(ws);
            const header = ws.getRow(headerRowNumber);
            const headerMap: Record<string, number> = {};
            header.eachCell((cell, colNumber) => {
                const key = normalizeHeaderText(cell.value?.toString() || '');
                const mapped = aliases[key];
                if (mapped && !headerMap[mapped]) headerMap[mapped] = colNumber;
            });
            const required = ['lineItem', 'productRange', 'product', 'sizeName', 'productSize', 'quantity', 'colour'];
            if (!required.some(k => headerMap[k])) continue;
            ws.eachRow((row, rowNumber) => {
                if (rowNumber <= headerRowNumber) return;
                const get = (field: string) => {
                    const col = headerMap[field];
                    if (!col) return '';
                    return this.stripBrackets(this.getCellValue(row.getCell(col)) as any).toString().trim();
                };
                const purchaseOrder = get('purchaseOrder');
                const product = get('product');
                const colour = get('colour');
                const lineItem = Number(get('lineItem') || '0');
                const qty = this.parseLooseNumber(get('quantity') || '0');
                if (!product || !colour || !Number.isFinite(lineItem) || lineItem <= 0) return;
                rows.push({
                    purchaseOrder,
                    lineItem,
                    range: get('productRange'),
                    product,
                    sizeName: get('sizeName'),
                    productSize: get('productSize'),
                    quantity: Number.isFinite(qty) ? qty : 0,
                    colour,
                });
            });
        }
        return rows;
    }

    private getCellValue(cell: ExcelJS.Cell) {
        const value = cell.isMerged && cell.master ? cell.master.value : cell.value;
        if (value && typeof value === 'object') {
            if ('result' in value && value.result !== undefined && value.result !== null) return value.result;
            if ('text' in value && typeof value.text === 'string') return value.text;
            if ('richText' in value && Array.isArray(value.richText)) {
                return value.richText.map((part: any) => part?.text || '').join('');
            }
            if ('hyperlink' in value && typeof value.hyperlink === 'string' && 'text' in value && typeof value.text === 'string') {
                return value.text;
            }
            if (typeof cell.text === 'string' && cell.text.trim()) return cell.text;
        }
        return value;
    }

    async generateOutputs(data: ProcessedPO[]) {
        const ordersWb = new ExcelJS.Workbook();
        const linesWb = new ExcelJS.Workbook();
        const sizesWb = new ExcelJS.Workbook();
        const ordersSheet = ordersWb.addWorksheet('ORDERS');
        const linesSheet = linesWb.addWorksheet('LINES');
        const sizesSheet = sizesWb.addWorksheet('ORDER_SIZES');

        ordersSheet.columns = [
            { header: 'PurchaseOrder', key: 'purchaseOrder' }, { header: 'ProductSupplier', key: 'productSupplier' },
            { header: 'Status', key: 'status' }, { header: 'Customer', key: 'customer' },
            { header: 'TransportMethod', key: 'transportMethod' }, { header: 'TransportLocation', key: 'transportLocation' },
            { header: 'PaymentTerm', key: 'paymentTerm' }, { header: 'Template', key: 'template' },
            { header: 'KeyDate', key: 'keyDate' }, { header: 'ClosedDate', key: 'closedDate' },
            { header: 'DefaultDeliveryDate', key: 'defaultDeliveryDate' }, { header: 'Comments', key: 'comments' },
            { header: 'Currency', key: 'currency' }, { header: 'KeyUser1', key: 'keyUser1' },
            { header: 'KeyUser2', key: 'keyUser2' }, { header: 'KeyUser3', key: 'keyUser3' },
            { header: 'KeyUser4', key: 'keyUser4' }, { header: 'KeyUser5', key: 'keyUser5' },
            { header: 'KeyUser6', key: 'keyUser6' }, { header: 'KeyUser7', key: 'keyUser7' },
            { header: 'KeyUser8', key: 'keyUser8' }, { header: 'ArchiveDate', key: 'archiveDate' },
            { header: 'PurchaseUOM', key: 'purchaseUOM' }, { header: 'SellingUOM', key: 'sellingUOM' },
            { header: 'ProductSupplierExt', key: 'productSupplierExt' }, { header: 'FindField_ProductSupplier', key: 'findField_ProductSupplier' },
        ];

        linesSheet.columns = [
            { header: 'PurchaseOrder', key: 'purchaseOrder' }, { header: 'LineItem', key: 'lineItem' },
            { header: 'ProductRange', key: 'productRange' }, { header: 'Product', key: 'product' },
            { header: 'Customer', key: 'customer' }, { header: 'DeliveryDate', key: 'deliveryDate' },
            { header: 'TransportMethod', key: 'transportMethod' }, { header: 'TransportLocation', key: 'transportLocation' },
            { header: 'Status', key: 'status' }, { header: 'PurchasePrice', key: 'purchasePrice' },
            { header: 'SellingPrice', key: 'sellingPrice' }, { header: 'Template', key: 'template' },
            { header: 'KeyDate', key: 'keyDate' }, { header: 'SupplierProfile', key: 'supplierProfile' },
            { header: 'ClosedDate', key: 'closedDate' }, { header: 'Comments', key: 'comments' },
            { header: 'Currency', key: 'currency' }, { header: 'ArchiveDate', key: 'archiveDate' },
            { header: 'ProductExternalRef', key: 'productExternalRef' }, { header: 'ProductCustomerRef', key: 'productCustomerRef' },
            { header: 'PurchaseUOM', key: 'purchaseUOM' }, { header: 'SellingUOM', key: 'sellingUOM' },
            { header: 'UDF-buyer_po_number', key: 'udfBuyerPoNumber' }, { header: 'UDF-start_date', key: 'udfStartDate' },
            { header: 'UDF-canel_date', key: 'udfCanelDate' }, { header: 'UDF-Inspection result', key: 'udfInspectionResult' },
            { header: 'UDF-Report Type', key: 'udfReportType' }, { header: 'UDF-Inspector', key: 'udfInspector' },
            { header: 'UDF-Approval Status', key: 'udfApprovalStatus' }, { header: 'UDF-Submitted inspection date', key: 'udfSubmittedInspectionDate' },
            { header: 'FindField_Product', key: 'findField_Product' },
        ];

        sizesSheet.columns = [
            { header: 'PurchaseOrder', key: 'purchaseOrder' }, { header: 'LineItem', key: 'lineItem' },
            { header: 'Range', key: 'range' }, { header: 'Product', key: 'product' },
            { header: 'SizeName', key: 'sizeName' }, { header: 'ProductSize', key: 'productSize' },
            { header: 'Quantity', key: 'quantity' }, { header: 'Colour', key: 'colour' },
            { header: 'Customer', key: 'customer' }, { header: 'Department', key: 'department' },
            { header: 'CustomAttribute1', key: 'customAttribute1' }, { header: 'CustomAttribute2', key: 'customAttribute2' },
            { header: 'CustomAttribute3', key: 'customAttribute3' }, { header: 'LineRatio', key: 'lineRatio' },
            { header: 'ColourExt', key: 'colourExt' }, { header: 'CustomerExt', key: 'customerExt' },
            { header: 'DepartmentExt', key: 'departmentExt' }, { header: 'CustomAttribute1Ext', key: 'customAttribute1Ext' },
            { header: 'CustomAttribute2Ext', key: 'customAttribute2Ext' }, { header: 'CustomAttribute3Ext', key: 'customAttribute3Ext' },
            { header: 'ProductExternalRef', key: 'productExternalRef' }, { header: 'ProductCustomerRef', key: 'productCustomerRef' },
            { header: 'FindField_Colour', key: 'findField_Colour' }, { header: 'FindField_Customer', key: 'findField_Customer' },
            { header: 'FindField_Department', key: 'findField_Department' }, { header: 'FindField_CustomAttribute1', key: 'findField_CustomAttribute1' },
            { header: 'FindField_CustomAttribute2', key: 'findField_CustomAttribute2' }, { header: 'FindField_CustomAttribute3', key: 'findField_CustomAttribute3' },
            { header: 'FindField_Product', key: 'findField_Product' },
        ];

        if (data && data.length > 0) {
            data.forEach(po => {
                const brandKey = (po.header.brandKey || '').trim().toLowerCase();
                const isLlBean = brandKey === 'll bean' || (po.header.customer || '').trim().toLowerCase() === 'll bean';
                const orderEntries = (po.orderKeys && po.orderKeys.length > 0)
                    ? po.orderKeys
                    : [{ purchaseOrder: po.header.purchaseOrder, customer: po.header.customer || '', customerName: po.header.customer, transportLocation: po.header.transportLocation, transportMethod: po.header.transportMethod, ordersTemplate: po.header.ordersTemplate }];
                orderEntries.forEach(entry => {
                    const hhTransportLocation = brandKey === 'hh' || brandKey === 'helly hansen'
                        ? (entry.transportLocation || this.extractCountryFromPurchaseOrder(entry.purchaseOrder) || po.header.transportLocation)
                        : (isLlBean
                            ? (entry.transportLocation || this.extractCountryFromPurchaseOrder(entry.purchaseOrder) || po.header.transportLocation)
                            : (entry.transportLocation || po.header.transportLocation));
                    ordersSheet.addRow({
                        purchaseOrder: entry.purchaseOrder || po.header.purchaseOrder, productSupplier: po.header.productSupplier,
                        status: 'Confirmed', customer: entry.customerName || entry.customer,
                        transportMethod: entry.transportMethod, transportLocation: hhTransportLocation,
                        paymentTerm: '', template: entry.ordersTemplate,
                        keyDate: po.header.keyDateFormat === 'manual' ? this.formatManualDateString(po.header.keyDate) : this.formatDateString(po.header.keyDate),
                        closedDate: '', defaultDeliveryDate: '', comments: po.header.comments, currency: 'USD',
                        keyUser1: po.header.keyUser1, keyUser2: po.header.keyUser2, keyUser3: po.header.keyUser3,
                        keyUser4: po.header.keyUser4, keyUser5: po.header.keyUser5, keyUser6: po.header.keyUser6,
                        keyUser7: po.header.keyUser7, keyUser8: po.header.keyUser8,
                        archiveDate: '', purchaseUOM: '', sellingUOM: '', productSupplierExt: '', findField_ProductSupplier: '',
                    });
                });
            });

            data.forEach(po => {
                const brandKey = (po.header.brandKey || '').trim().toLowerCase();
                const isLlBean = brandKey === 'll bean' || (po.header.customer || '').trim().toLowerCase() === 'll bean';
                const orderEntries = (po.orderKeys && po.orderKeys.length > 0)
                    ? po.orderKeys
                    : [{ purchaseOrder: po.header.purchaseOrder, customer: po.header.customer || '', customerName: po.header.customer, transportLocation: po.header.transportLocation, transportMethod: po.header.transportMethod, ordersTemplate: po.header.ordersTemplate }];
                orderEntries.forEach(entry => {
                    const hhTransportLocation = brandKey === 'hh' || brandKey === 'helly hansen'
                        ? (entry.transportLocation || this.extractCountryFromPurchaseOrder(entry.purchaseOrder) || po.header.transportLocation)
                        : (isLlBean
                            ? (entry.transportLocation || this.extractCountryFromPurchaseOrder(entry.purchaseOrder) || po.header.transportLocation)
                            : (entry.transportLocation || po.header.transportLocation));
                    const linesForEntry = isLlBean
                        ? po.lines.filter(line => {
                            const lineLocation = this.normalizeTransportLocation(line.transportLocation || po.header.transportLocation || '');
                            const entryLocation = this.normalizeTransportLocation(entry.transportLocation || po.header.transportLocation || '');
                            return !entryLocation || !lineLocation || lineLocation === entryLocation;
                        })
                        : po.lines;
                    const lineItemMap = new Map<number, number>();
                    linesForEntry.forEach((line, idx) => lineItemMap.set(line.lineItem, idx + 1));
                    linesForEntry.forEach(line => {
                        const isOnAg = (po.header.customer || '').trim().toLowerCase() === 'on ag';
                        const normalizedCustomer = (po.header.customer || '').trim().toLowerCase();
                        const isArcteryx = normalizedCustomer === 'arcteryx' || normalizedCustomer === "arc'teryx";
                        const isJackWolfskin = normalizedCustomer === 'jack wolfskin' || brandKey === 'jack wolfskin';
                        const isBurton = (po.header.brandKey || '').trim().toLowerCase() === 'burton';
                        const isHunter = (po.header.brandKey || '').trim().toLowerCase() === 'hunter';
                        const isHH = brandKey === 'hh' || brandKey === 'helly hansen' || normalizedCustomer === 'helly hansen';
                        const isLlBean = brandKey === 'll bean' || normalizedCustomer === 'll bean';
                        const isDynafit = brandKey === 'dynafit' || normalizedCustomer === 'dynafit';
                        const hhDeliveryDate = isHH
                            ? (line.hhConfirmedDeliveryDate || this.formatDateString(line.cancelDate) || this.formatDateString(line.startDate) || line.hhCancelDate || line.hhStartDate || '')
                            : '';
                        const llbDeliveryDate = isLlBean
                            ? (this.formatDateString(line.cancelDate) || this.formatDateString(line.startDate) || '')
                            : '';
                        const dynafitDeliveryDate = isDynafit
                            ? (this.formatDateString((line.startDate as any)) || this.formatDateString((line.cancelDate as any)) || this.formatDateString((line.dynafitLineKeyDate as any)))
                            : '';
                        const jwsDeliveryDate = isJackWolfskin
                            ? (this.formatDateString(line.startDate) || this.formatDateString(line.cancelDate) || '')
                            : '';
                        const dynafitLineKeyDate = isDynafit
                            ? (line.dynafitLineKeyDate || line.startDate || '')
                            : line.startDate;
                        const jwsLineKeyDate = isJackWolfskin
                            ? (() => {
                                const parsed = this.parseDate(line.startDate as any);
                                if (!parsed) return line.startDate;
                                const firstDay = new Date(parsed);
                                firstDay.setDate(1);
                                firstDay.setHours(0, 0, 0, 0);
                                return firstDay;
                            })()
                            : line.startDate;
                        const manualLineKeyDate = po.manualKeyDate || '';
                        const exportDeliveryDate = isOnAg
                            ? this.formatDateString(this.shiftDate(line.startDate, -1) || line.startDate)
                            : isHH
                                ? hhDeliveryDate
                                : isLlBean
                                    ? llbDeliveryDate
                                    : isJackWolfskin
                                        ? jwsDeliveryDate
                                            : isDynafit
                                                ? (this.formatDateString(line.cancelDate) || this.formatDateString(line.startDate) || dynafitDeliveryDate)
                                            : ((brandKey === '66 degrees north' || brandKey === 'cotopaxi' || brandKey === 'hunter' || brandKey === 'marmot')
                                                ? (this.formatDateString(line.cancelDate) || this.formatDateString(line.startDate) || '')
                                                : this.formatDateString(line.startDate));
                        const lineKeyDate = isOnAg
                            ? (this.shiftDate(line.startDate, -1) || line.startDate)
                            : (manualLineKeyDate || (isDynafit ? dynafitLineKeyDate : line.startDate));
                        linesSheet.addRow({
                            purchaseOrder: entry.purchaseOrder || po.header.purchaseOrder, lineItem: lineItemMap.get(line.lineItem) || line.lineItem, productRange: line.productRange,
                            product: line.styleNumber, customer: entry.customerName || entry.customer || po.header.customer,
                            deliveryDate: exportDeliveryDate,
                            transportMethod: po.header.transportMethod, transportLocation: isHunter ? (line.transportLocation || '') : hhTransportLocation,
                            status: 'Confirmed', purchasePrice: line.cost ?? '', sellingPrice: '',
                            template: po.header.linesTemplate, keyDate: this.formatDateString(manualLineKeyDate || jwsLineKeyDate || lineKeyDate),
                            supplierProfile: line.supplierProfile, closedDate: '', comments: '', currency: 'USD',
                            archiveDate: '', productExternalRef: (isArcteryx || isHunter || isBurton || isJackWolfskin) ? '' : (line.productExternalRef || ''), productCustomerRef: (isArcteryx || isJackWolfskin) ? '' : (line.productCustomerRef || ''),
                            purchaseUOM: '', sellingUOM: '', udfBuyerPoNumber: isDynafit
                                ? (line.buyerPoNumber?.toString?.() || '')
                                : (line.buyerPoNumber?.toString?.() || line.buyerPoNumber || ''),
                            udfStartDate: exportDeliveryDate || this.formatDateString(line.startDate) || this.formatDateString(line.cancelDate) || '',
                            udfCanelDate: exportDeliveryDate || this.formatDateString(line.startDate) || this.formatDateString(line.cancelDate) || '',
                            udfInspectionResult: '', udfReportType: '', udfInspector: '', udfApprovalStatus: '',
                            udfSubmittedInspectionDate: '', findField_Product: '',
                        }).commit();
                    });
                });
            });

        data.forEach(po => {
            const brandKey = (po.header.brandKey || '').trim().toLowerCase();
            const isLlBean = brandKey === 'll bean' || (po.header.customer || '').trim().toLowerCase() === 'll bean';
            if (isLlBean && Array.isArray(po.llBeanReferenceSizeRows) && po.llBeanReferenceSizeRows.length > 0) {
                po.llBeanReferenceSizeRows.forEach(refRow => {
                    sizesSheet.addRow({
                        purchaseOrder: '',
                        lineItem: refRow.lineItem,
                        range: refRow.range,
                        product: refRow.product,
                        sizeName: refRow.sizeName,
                        productSize: refRow.productSize,
                        quantity: refRow.quantity,
                        colour: refRow.colour,
                        customer: '',
                        department: '',
                        customAttribute1: '',
                        customAttribute2: '',
                        customAttribute3: '',
                        lineRatio: '',
                        colourExt: '',
                        customerExt: '',
                        departmentExt: '',
                        customAttribute1Ext: '',
                        customAttribute2Ext: '',
                        customAttribute3Ext: '',
                        productExternalRef: '',
                        productCustomerRef: '',
                        findField_Colour: '',
                        findField_Customer: '',
                        findField_Department: '',
                        findField_CustomAttribute1: '',
                        findField_CustomAttribute2: '',
                        findField_CustomAttribute3: '',
                        findField_Product: '',
                    });
                });
                return;
            }
            const orderEntries = (po.orderKeys && po.orderKeys.length > 0)
                ? po.orderKeys
                : [{ purchaseOrder: po.header.purchaseOrder, customer: po.header.customer || '', customerName: po.header.customer, transportLocation: po.header.transportLocation, transportMethod: po.header.transportMethod, ordersTemplate: po.header.ordersTemplate }];
            orderEntries.forEach(entry => {
                const hhTransportLocation = brandKey === 'hh' || brandKey === 'helly hansen'
                    ? (entry.transportLocation || this.extractCountryFromPurchaseOrder(entry.purchaseOrder) || po.header.transportLocation)
                    : (isLlBean
                        ? (entry.transportLocation || this.extractCountryFromPurchaseOrder(entry.purchaseOrder) || po.header.transportLocation)
                            : (entry.transportLocation || po.header.transportLocation));
                    const linesForEntry = isLlBean
                        ? po.lines.filter(line => {
                            const lineLocation = this.normalizeTransportLocation(line.transportLocation || po.header.transportLocation || '');
                            const entryLocation = this.normalizeTransportLocation(entry.transportLocation || po.header.transportLocation || '');
                            return !entryLocation || !lineLocation || lineLocation === entryLocation;
                        })
                        : po.lines;
                    const lineItemMap = new Map<number, number>();
                    linesForEntry.forEach((line, idx) => lineItemMap.set(line.lineItem, idx + 1));
                    linesForEntry.forEach(line => {
                        (po.sizes[line.lineItem] || []).forEach(sz => {
                            sizesSheet.addRow({
                                purchaseOrder: entry.purchaseOrder || po.header.purchaseOrder, lineItem: lineItemMap.get(line.lineItem) || line.lineItem, range: line.productRange,
                                product: line.styleNumber, sizeName: sz.productSize, productSize: sz.productSize,
                        quantity: sz.quantity, colour: (brandKey === 'dynafit' ? (line.colour || line.rawColour || line.styleColor) : line.colour), customer: '', department: '',
                                customAttribute1: '', customAttribute2: '', customAttribute3: '',
                                lineRatio: '', colourExt: '', customerExt: '', departmentExt: '',
                                customAttribute1Ext: '', customAttribute2Ext: '', customAttribute3Ext: '',
                                productExternalRef: '', productCustomerRef: '',
                                findField_Colour: '', findField_Customer: '', findField_Department: '',
                                findField_CustomAttribute1: '', findField_CustomAttribute2: '',
                                findField_CustomAttribute3: '', findField_Product: '',
                            });
                        });
                    });
                });
            });
        }

        if (this.runId && this.userId) {
            await logEvent({ eventName: 'OUTPUT_GENERATED', userId: this.userId, runId: this.runId, metadata: { orders_count: data.length } });
        }

        return {
            orders: await ordersWb.xlsx.writeBuffer(),
            lines: await linesWb.xlsx.writeBuffer(),
            sizes: await sizesWb.xlsx.writeBuffer(),
        };
    }
}
