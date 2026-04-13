export const FALLBACK_COLUMN_ALIASES: Record<string, string> = {
    'transportlocation': 'transportLocation',
    'transport location': 'transportLocation',
    'destination': 'transportLocation',
    'dest country': 'transportLocation',
    'plant': 'plant',
    'plant code': 'plant',
    'ult. destination': 'transportLocation',
    'purchaseorder': 'purchaseOrder',
    'po': 'purchaseOrder',
    'po #': 'purchaseOrder',
    'po#': 'buyerPoNumber',
    'pono': 'purchaseOrder',
    'purchase order': 'purchaseOrder',
    'purchase order no': 'purchaseOrder',
    'style code': 'product',
    'm88 reference': 'product',
    'tracking number': 'purchaseOrder',
    'extraction po #': 'buyerPoNumber',
    'extraction po#': 'buyerPoNumber',
    'style number': 'product',
    'style nr': 'productCustomerRef',
    'item': 'product',
    'line numb': 'lineItem',
    'line id': 'lineItem',
    'line number': 'lineItem',
    'style no': 'product',
    'style#': 'product',
    'style #': 'product',
    'material style': 'product',
    'style': 'product',
    'product': 'product',
    'article': 'product',
    'model': 'product',
    'product name': 'product',
    'buyer style number': 'productCustomerRef',
    'buyer style no': 'productCustomerRef',
    'buyer style #': 'productCustomerRef',
    'buyer style': 'productCustomerRef',
    'name': 'productExternalRef',
    'product external ref': 'productExternalRef',
    'product customer ref': 'productCustomerRef',
    'style name': 'productExternalRef',
    'material name': 'ignore',
    'stylecolor': 'inlineStyleColor',
    'qty jan buy size-split': 'quantity',
    'bp no': 'buyerPoNumber',
    'vendor confirmed etd': 'confirmedExFac',
    'vendor account': 'ignore',
    'vendor code': 'productSupplier',
    'vendor number': 'productSupplier',
    'vendor sku': 'ignore',
    'upc code': 'ignore',
    'vendorcode': 'productSupplier',
    'vendor': 'productSupplier',
    'supplier': 'productSupplier',
    'product supplier': 'productSupplier',
    'productsupplier': 'productSupplier',
    'vendorname': 'vendorName',
    'vendor name': 'vendorName',
    'supplier name': 'vendorName',
    'factory': 'vendorName',
    'updated fty': 'productSupplier',
    'size': 'sizeName',
    'size name': 'sizeName',
    'sizename': 'sizeName',
    'productsize': 'sizeName',
    'product size': 'sizeName',
    'size #': 'sizeName',
    'size#': 'sizeName',
    'size code': 'sizeName',
    'size_name': 'sizeName',
    'size-name': 'sizeName',
    'ordered qty': 'quantity',
    'open qty (pcs/prs)': 'quantity',
    'requested qty': 'quantity',
    'quantity': 'quantity',
    'quantity|n': 'quantity',
    'qty': 'quantity',
    'qty (lum)': 'quantity',
    'sum of qty (lum)': 'quantity',
    'revised qty (0 if cancel, new qty if top up or reduce)': 'quantity',
    'final qty': 'finalQty',
    'deliverydate': 'exFtyDate',
    'delivery date': 'exFtyDate',
    'requested delivery date': 'exFtyDate',
    'requested etd|n': 'exFtyDate',
    'xf': 'exFtyDate',
    'xf date': 'exFtyDate',
    'final xf date': 'exFtyDate',
    'final xf date 3.16': 'exFtyDate',
    'confirmed delivery date': 'confirmedExFac',
    'confirmed ex-factory date|n': 'confirmedExFac',
    'orig ex fac': 'exFtyDate',
    'negotiated ex fac date': 'exFtyDate',
    'ex fac': 'exFtyDate',
    'ex-factory': 'exFtyDate',
    'vendor confirmed crd': 'exFtyDate',
    'final crd (order date + lt1)': 'confirmedExFac',
    'confirmed fty ex fac': 'confirmedExFac',
    'confirmed ex fac': 'confirmedExFac',
    'confirmed ex-factory': 'confirmedExFac',
    'fty ex fac': 'confirmedExFac',
    'keydate': 'poIssuanceDate',
    'buy date': 'buyDate',
    'file date': 'buyDate',
    'creation date': 'buyDate',
    'created date': 'buyDate',
    'report_date': 'poIssuanceDate',
    'report date': 'poIssuanceDate',
    'cancel date': 'cancelDate',
    'canceldate': 'cancelDate',
    'cancel': 'cancelDate',
    'po issuance date': 'poIssuanceDate',
    'issue date': 'buyDate',
    'transportmethod': 'transportMethod',
    'transport method': 'transportMethod',
    'ship mode': 'transportMethod',
    'shipping instr.': 'transportMethod',
    'mode of delivery': 'transportMethod',
    'mode of delivery|n': 'transportMethod',
    'm3 delivery method description': 'transportMethod',
    'trans cond': 'transportMethod',
    'transport mode': 'transportMethod',
    'transportation mode': 'transportMethod',
    'shipment mode': 'transportMethod',
    'doc type': 'template',
    'purchase order type': 'template',
    'template': 'template',
    'supply planning team owner': 'ignore',
    'supplier number': 'ignore',
    'whs': 'ignore',
    'po company name': 'ignore',
    'po line status': 'status',
    'po cutting status': 'status',
    'agreement used head': 'ignore',
    'supply lead time': 'ignore',
    'purchase price (po m3)': 'ignore',
    'sp comments': 'ignore',
    'remarks': 'ignore',
    'remark': 'ignore',
    'your reference': 'ignore',
    'range': 'season',
    'productrange': 'season',
    'season': 'season',
    'season plan': 'season',
    'season plan model commercial name': 'ignore',
    'season indicator': 'season',
    'season cc': 'season',
    'brand': 'brand',
    'business unit description': 'brand',
    'customer': 'customerName',
    'customer name': 'customerName',
    'plm customer name': 'customerName',
    'status': 'status',
    'approval stat': 'status',
    'confirmation status': 'status',
    'gsc type': 'status',
    'colour': 'colour',
    'color': 'colour',
    'color name': 'colour',
    'article name': 'colour',
    'color description': 'colour',
    'warehouse name': 'transportLocation',
    'delivery terms|n': 'ignore',
    'coo line': 'ignore',
    'po ref 1': 'ignore',
    'po ref 2': 'ignore',
    'brand id tag code': 'ignore',
    'site|n': 'ignore',
    'warehouse|n': 'ignore',
    'product line': 'ignore',
    'product class': 'ignore',
    'product subclass': 'ignore',
    'load id': 'ignore',
    'remain delivery qty': 'ignore',
    'approval status': 'status',
    'purch line group num': 'ignore',
    'purchase type': 'ignore',
    'barcode': 'ignore',
    'material': 'colour',
    'longtext': 'colourDisplay',
    'sku': 'productExternalRef',
    'submit buy': 'buyRound',
    'buy round': 'buyRound',
    'order date': 'buyDate',
    'buyer style name': 'ignore',
    'jde style': 'jdeStyle',
    'm88 ref': 'jdeStyle',
    'udf-buyer_po_number': 'buyerPoNumber',
    'udf-start_date': 'exFtyDate',
    'udf-canel_date': 'cancelDate',
    'final po cut#': 'ignore',
    'master po#': 'purchaseOrder',
    'final factory': 'productSupplier',
    'final vendor name': 'vendorName',
    'plant name': 'plantName',
    'mhp capacity type': 'category',
    'purchasing document number': 'purchaseOrder',
    'purchasing document': 'purchaseOrder',
    'material description': 'colour',
    'grid value': 'sizeName',
    'order qty': 'quantity',
    'ex factory date': 'exFtyDate',
    'ex. factory': 'exFtyDate',
    'so order cancel date': 'cancelDate',
    'shipping instructions': 'transportMethod',
    'goods supplier name': 'vendorName',
    'po status': 'status',
    'purchasing document date': 'buyDate',
    'purchasing group': 'ignore',
    'purchasing group description': 'ignore',
    'item number of purchasing document': 'ignore',
    'item#': 'product',
    'item #': 'product',
    'color code': 'colour',
    'updated planned exit date': 'exFtyDate',
    'confirmed x-fty': 'confirmedExFac',
    'wh': 'plant',
    'style color': 'colour',
    'buyer item #': 'product',
    'buyer item#': 'product',
    'short text': 'productExternalRef',
    'short description': 'productExternalRef',
    'country/region': 'transportLocation',
    'create date': 'buyDate',
    'purchase order number': 'purchaseOrder',
    'requested exv date': 'exFtyDate',
    'expected exv date (per supplier)': 'exFtyDate',
    'expected exw date (per supplier)': 'exFtyDate',
    'expected receipt date (per supplier exv)': 'ignore',
    'expected receipt date (per supplier exw)': 'ignore',
    'rev ex fact': 'ignore',
    'stat. ex fact': 'ignore',
    'po size delivery date': 'ignore',
    'stat del date': 'ignore',
    'fcr date': 'ignore',
    'receipt date': 'ignore',
    'ticket type': 'ignore',
    'ean/upc': 'ignore',
    'country of origin': 'ignore',
    'port of origin': 'ignore',
    'address': 'ignore',
    'city': 'ignore',
    'state': 'ignore',
    'zip': 'ignore',
    'quantity fulfilled/received': 'ignore',
    'recvd qty': 'ignore',
    'open qty': 'ignore',
    'goods supplier': 'ignore',
    'quantity on shipments': 'ignore',
    'merch - style name': 'productExternalRef',
    'merch - gender': 'ignore',
    'merch - sub category': 'category',
    'lineplan seasons': 'season',
    'memo (main)': 'ignore',
    'internal id': 'ignore',
    'type': 'ignore',
    'account assignment': 'ignore',
    'sales order': 'ignore',
    'sales order item': 'ignore',
    'so requested delivery date': 'ignore',
    'net value(so)': 'ignore',
    'material group description': 'ignore',
    'po date': 'buyDate',
    'style description': 'productExternalRef',
    'style - color': 'ignore',
    'style status': 'ignore',
    'x factory': 'exFtyDate',
    'xfactory': 'exFtyDate',
    'delivery date confirmed': 'confirmedExFac',
    'article code [sap]': 'product',
    'model code [sap]': 'productAlt',
    'primary color peak pdm code': 'colour',
    'buy 1 - tracking no.': 'purchaseOrder',
    'buy 1 cfm crd': 'exFtyDate',
    'buy 1 cfm crd (request crd)': 'exFtyDate',
    'final po qty': 'quantity',
    'buy 1 agreed qty': 'quantity',
    'production supplier name': 'vendorName',
    'production supplier sap supplier code': 'productSupplier',
    'ship via': 'transportMethod',
    'ship mode description': 'transportMethod',
    'ex-factory date': 'exFtyDate',
    'ex factory': 'exFtyDate',
    'requested exw date': 'exFtyDate',
    'ship window end date': 'exFtyDate',
    'merch - color': 'colour',
    'merch - size': 'sizeName',
    'po number': 'purchaseOrder',
    'destination name': 'ignore',
    'total': 'ignore',
    'efd': 'exFtyDate',
    'item code': 'product',
    'item description': 'productExternalRef',
    'colour code': 'colour',
    'colour description': 'colourDisplay',
    'requisition no': 'season',
    'shipment method': 'transportMethod',
    'bulk qty': 'quantity',
    'merch - season': 'season',
    'style no.': 'product',
    'colour desc': 'colour',
    'color desc': 'colour',
    'cost': 'ignore',
    'production surcharge': 'ignore',
    'sell': 'ignore',
    'costing reference': 'ignore',
    'net price': 'ignore',
    'net value': 'ignore',
    'article full colors': 'ignore',
    'article product scope [cpf]': 'ignore',
    'production supplier country': 'ignore',
    'eb_all in 1': 'ignore',
    'base fob': 'ignore',
    'prod s/c': 'ignore',
    'fob with s/c': 'ignore',
    'total selling price': 'ignore',
    'purchase price': 'ignore',
    'total purchase price': 'ignore',
    'landed cost': 'ignore',
    'total landed cost': 'ignore',
    'total margin': 'ignore',
    'crd': 'ignore',
    'eta date': 'ignore',
    'ship date': 'exFtyDate',
    'ship window': 'exFtyDate',
    'planned ship date': 'exFtyDate',
    'in-dc date': 'exFtyDate',
    'in dc date': 'exFtyDate',
    'dc arrival date': 'exFtyDate',
    'cancellation date': 'cancelDate',
    'qty ordered': 'quantity',
    'total qty': 'quantity',
    'sum of order qty': 'quantity',
    'units': 'quantity',
    'supplier code': 'productSupplier',
    'mfr code': 'productSupplier',
    'ship to country': 'transportLocation',
    'mode': 'transportMethod',
    'freight mode': 'transportMethod',
    'shipping method': 'transportMethod',
    'division': 'category',
    'product division': 'category',
    'gender': 'category',
    'gender code': 'category',
    'season code': 'season',
    'buyer po': 'buyerPoNumber',
    'buyer po #': 'buyerPoNumber',
    'customer po': 'buyerPoNumber',
    // Haglofs
    'old q': 'ignore',
    'old qty': 'ignore',
    'shifting (qty)': 'ignore',
    'new qty': 'quantity',
    'delivery da': 'exFtyDate',
    'fob': 'cost',
    // Vans
    'purchase order #': 'purchaseOrder',
    'mdm material (base)': 'product',
    'sap size 1': 'sizeName',
    'dc plant': 'plant',
    'factory name': 'vendorName',
    'factory na': 'vendorName',
    'factory code': 'productSupplier',
    'requested et': 'exFtyDate',
    'product subcl': 'category',
    // Roscoe
    'product code': 'product',
    'tot qty': 'quantity',
    'shipping date': 'exFtyDate',
    // Dynafit
    'price': 'ignore',
    'total price': 'ignore',
    'sizes': 'ignore',
    'ng style name': 'ignore',
    'drop date': 'ignore',
    'etd': 'buyDate',
    // Burton
    'seller name': 'vendorName',
    'supplier party id': 'ignore',
    'final destination': 'plant',
    'po reference': 'ignore',
    'materialnumber': 'ignore',
    'extended sizing': 'ignore',
    'ship to location name': 'ignore',
    // ON AG additional
    'sku #': 'sizeName',
    // Dynafit
    'description': 'productExternalRef',
    // Hunter
    'country': 'transportLocation',
    'ship-to party name': 'plantName',
    'sum of order total qty': 'quantity',
    'sum of quantity': 'quantity',
    'sum of qty (lum)2': 'quantity',
    'brand requested crd': 'exFtyDate',
    'packing splits': 'ignore',
};

export function normalizeHeaderText(value: string): string {
    return value.trim().toLowerCase().replace(/\s+/g, ' ');
}

export function isLikelyPivotSizeHeader(headerText: string): boolean {
    const normalized = normalizeHeaderText(headerText);
    if (!normalized) return false;

    const summaryKeywords = [
        'purchase price', 'confirmed unit price', 'grand total', 'moq',
        'unit price', 'total confirmed unit price', 'confirmed delivery date',
        'final xf date', 'your reference', 'sp comments', 'comments',
        'supply lead time', 'supply planning team owner', 'supplier number',
        'po company name', 'po line status', 'agreement used head', 'whs',
    ];
    if (summaryKeywords.some(k => normalized.includes(k))) return false;

    const knownSizeTokens = new Set([
        'std', 'os', 'one size', 'xss', 'ss', 'ms', 'ls', 'xls', 'xs', 's', 'm', 'l', 'xl',
        '2xs', '3xs', '2x', '3x', '2xl', '3xl', '4xl', '5xl', '1x',
    ]);
    if (knownSizeTokens.has(normalized)) return true;

    if (/^y\d+\/\d+$/i.test(normalized)) return true;
    if (/^\d{1,3}$/.test(normalized)) return true;
    if (/^\d+(?:\.\d+)?[a-z]?$/i.test(normalized)) return true;
    if (/^\d+\s*[\/-]\s*\d+[a-z]?$/i.test(normalized)) return true;
    if (/^[a-z]\d+\s*[\/-]\s*\d+$/i.test(normalized)) return true;

    return false;
}

export function resolveColumnAlias(headerText: string, aliases: Record<string, string> = FALLBACK_COLUMN_ALIASES): string | undefined {
    return aliases[normalizeHeaderText(headerText)];
}

function isValidDateParts(year: number, monthIndex: number, day: number): boolean {
    const date = new Date(year, monthIndex, day);
    return !Number.isNaN(date.getTime())
        && date.getFullYear() === year
        && date.getMonth() === monthIndex
        && date.getDate() === day;
}

export function parseExcelEngineDate(raw: string | number | Date | undefined | null): Date | null {
    if (raw === undefined || raw === null || raw === '') return null;
    if (raw instanceof Date) return Number.isNaN(raw.getTime()) ? null : raw;
    if (typeof raw === 'number' && Number.isFinite(raw)) {
        const serial = Math.trunc(raw);
        if (serial >= 1 && serial <= 2958465) {
            const epoch = new Date(1899, 11, 30);
            epoch.setDate(epoch.getDate() + serial);
            return epoch;
        }
        return null;
    }

    const text = String(raw).trim();
    if (!text) return null;

    let match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (match) {
        const year = Number(match[1]);
        const monthIndex = Number(match[2]) - 1;
        const day = Number(match[3]);
        return isValidDateParts(year, monthIndex, day) ? new Date(year, monthIndex, day) : null;
    }

    match = text.match(/^(\d{4})\/(\d{2})\/(\d{2})$/);
    if (match) {
        const year = Number(match[1]);
        const monthIndex = Number(match[2]) - 1;
        const day = Number(match[3]);
        return isValidDateParts(year, monthIndex, day) ? new Date(year, monthIndex, day) : null;
    }

    match = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
    if (match) {
        const monthIndex = Number(match[1]) - 1;
        const day = Number(match[2]);
        const year = Number(match[3]);
        return isValidDateParts(year, monthIndex, day) ? new Date(year, monthIndex, day) : null;
    }

    const monthNames = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
    match = text.match(/^(\d{1,2})-([A-Za-z]+)-(\d{4})$/);
    if (match) {
        const m1 = match;
        const monthIndex = monthNames.findIndex(m => m1[2].toLowerCase().startsWith(m));
        const day = Number(match[1]);
        const year = Number(match[3]);
        return monthIndex >= 0 && isValidDateParts(year, monthIndex, day) ? new Date(year, monthIndex, day) : null;
    }

    match = text.match(/^([A-Za-z]+)\s+(\d{1,2})\s+(\d{4})$/);
    if (match) {
        const m2 = match;
        const monthIndex = monthNames.findIndex(m => m2[1].toLowerCase().startsWith(m));
        const day = Number(match[2]);
        const year = Number(match[3]);
        return monthIndex >= 0 && isValidDateParts(year, monthIndex, day) ? new Date(year, monthIndex, day) : null;
    }

    const lastMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (lastMatch) {
        const day = Number(lastMatch[1]);
        const monthIndex = Number(lastMatch[2]) - 1;
        const year = Number(lastMatch[3]);
        if (day > 12) {
            return isValidDateParts(year, monthIndex, day) ? new Date(year, monthIndex, day) : null;
        }
    }

    return null;
}

export function formatStandardDate(raw: string | number | Date | undefined | null): string {
    const date = parseExcelEngineDate(raw);
    if (!date) return '';
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const dd = String(date.getDate()).padStart(2, '0');
    return `${mm}/${dd}/${date.getFullYear()}`;
}

export function formatManualDate(raw: string | number | Date | undefined | null): string {
    const date = parseExcelEngineDate(raw);
    if (!date) return '';
    return `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
}

export interface PivotColumn {
    colNumber: number;
    headerText: string;
}

export interface PivotFormatDetection {
    isPivotFormat: boolean;
    fixedColumnNumbers: number[];
    pivotColumns: PivotColumn[];
}

export function detectPivotFormatFromHeaders(
    headers: PivotColumn[],
    aliases: Record<string, string>,
    shouldIgnoreHeader: (header: string) => boolean,
): PivotFormatDetection {
    const fixedColumnNumbers = headers
        .filter(({ headerText }) => {
            const alias = aliases[normalizeHeaderText(headerText)];
            return !!alias && alias !== 'ignore';
        })
        .map(({ colNumber }) => colNumber);

    const maxFixedColumn = fixedColumnNumbers.length > 0 ? Math.max(...fixedColumnNumbers) : 0;
    const pivotColumns = headers.filter(({ colNumber, headerText }) => {
        const normalized = normalizeHeaderText(headerText);
        if (!normalized || colNumber <= maxFixedColumn) return false;
        if (aliases[normalized]) return false;
        if (shouldIgnoreHeader(headerText)) return false;
        if (!isLikelyPivotSizeHeader(headerText)) return false;
        // Rossignol and similar sheets may have a single side-list product code header
        // (e.g. RLOMH02) after the real fixed columns; that is not a pivot column.
        if (/^[A-Z]{2,}\d[A-Z0-9-]*$/i.test(headerText.trim())) return false;
        return true;
    });

    return {
        isPivotFormat: pivotColumns.length > 0,
        fixedColumnNumbers,
        pivotColumns,
    };
}
