import { NextRequest, NextResponse } from "next/server";
import { createWriteClient, isWriteEnabled, NextGenInsertResult } from "@/lib/nextgen";
import { ProcessedPO } from "@/lib/excel-engine";
import {
    CURRENCY_MAP, STATUS_MAP, TEMPLATE_MAP, TRANSPORT_METHOD_MAP,
    LOCATION_MAP, resolveId,
} from "@/lib/nextgen-field-mapping";
import { updateProgress } from "../nextgen-upload-progress/route";

/**
 * POST /api/nextgen-upload
 *
 * Uploads processed PO data (header + lines + sizes) to the NextGen TEST
 * environment via the InsertFormData / InsertFormLinesGridRecord /
 * UpdateQuantityRecord endpoints.
 *
 * SAFETY: This route uses createWriteClient() which enforces that the
 * target URL is the test environment (port 8443). It will REFUSE to
 * run against production.
 *
 * Request body:
 * {
 *   poData: ProcessedPO[]    // the merged PO data from /api/upload
 * }
 *
 * Response:
 * {
 *   success: boolean,
 *   results: {
 *     po: NextGenInsertResult,
 *     lines: NextGenInsertResult[],
 *     sizes: NextGenInsertResult[],
 *   },
 *   summary: { posCreated, linesCreated, sizesUpdated, errors }
 * }
 */
export async function POST(req: NextRequest) {
    let progressId = '';
    try {
        if (!isWriteEnabled()) {
            return NextResponse.json({
                error: "NextGen write operations are not enabled. Set NEXTGEN_WRITE_BASE_URL (port 8443), NEXTGEN_WRITE_USERNAME, NEXTGEN_WRITE_PASSWORD, and NEXTGEN_WRITE_ENABLED=true in .env.local."
            }, { status: 503 });
        }

        const body = await req.json();
        const { poData } = body as { poData: ProcessedPO[] };
        progressId = body.progressId || `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;

        if (!poData || !Array.isArray(poData) || poData.length === 0) {
            return NextResponse.json({
                error: "poData is required and must be a non-empty array of ProcessedPO objects."
            }, { status: 400 });
        }

        // Normalize poData: accept both ProcessedPO[] (with header) and
        // ProductData[]/flat objects (without header) by synthesizing a header
        // from line-level fields when needed.
        const normalizedPoData = poData.map((po: any, idx: number) => {
            if (po.header) return po;
            // Flat ProductData or BuyFileItem — build a minimal header
            const firstLine = (po.lines || [])[0] || po;
            const poNumber = po.header?.purchaseOrder || po.poNumber || firstLine.poNumber || firstLine.po_number || `PO-${idx + 1}`;
            return {
                header: {
                    purchaseOrder: poNumber,
                    customer: firstLine.customer || po.customer || '',
                    currency: firstLine.currency || po.currency || 'USD',
                    productSupplier: firstLine.factory || po.factory || '',
                    status: 'Confirmed',
                    transportMethod: firstLine.transportMethod || '',
                    transportLocation: firstLine.transportLocation || '',
                    ordersTemplate: 'Major Brand Bulk',
                    linesTemplate: '',
                    keyDate: firstLine.deliveryDate || firstLine.exFtyDate || '',
                    comments: '',
                    keyUser1: '', keyUser2: '', keyUser3: '', keyUser4: '', keyUser5: '',
                    keyUser6: '', keyUser7: '', keyUser8: '',
                },
                lines: po.lines || [{
                    lineItem: 1,
                    styleNumber: po.style || firstLine.styleNumber || '',
                    productExternalRef: po.productExternalRef || '',
                    productCustomerRef: po.productCustomerRef || po.style || '',
                    colour: po.colorName || po.color || '',
                    color: po.colorName || po.color || '',
                    styleColor: po.colorCode || '',
                    season: po.season || '',
                    exFtyDate: po.deliveryDate || '',
                    factory: po.factory || '',
                    customer: po.customer || '',
                }],
                sizes: po.sizes || {},
                orderKeys: po.orderKeys || [],
            };
        });

        console.log('[nextgen-upload] starting upload to TEST environment', {
            poCount: normalizedPoData.length,
            totalLines: normalizedPoData.reduce((a, p) => a + (p.lines?.length || 0), 0),
            totalSizes: normalizedPoData.reduce((a, p) => {
                const sizes = p.sizes as any;
                if (Array.isArray(sizes)) return a + sizes.flat().length;
                return a + Object.values(sizes).reduce((b: number, s: any) => b + (s?.length || 0), 0);
            }, 0),
        });

        const client = createWriteClient();

        // Progress tracking
        const totalLines = normalizedPoData.reduce((a, p) => a + (p.lines?.length || 0), 0);
        const totalSizes = normalizedPoData.reduce((a, p) => {
            const sizes = p.sizes as any;
            if (Array.isArray(sizes)) return a + sizes.flat().length;
            return a + Object.values(sizes || {}).reduce((b: number, s: any) => b + (s?.length || 0), 0);
        }, 0);
        const totalSteps = normalizedPoData.length + totalLines + totalSizes;
        let currentStep = 0;
        updateProgress(progressId, { total: totalSteps, current: 0, status: 'uploading', message: 'Starting upload...' });

        const allResults = {
            pos: [] as Array<{ poNumber: string; result: NextGenInsertResult; lines: Array<{ lineItem: number; result: NextGenInsertResult }>; sizes: Array<{ lineItem: number; sizeName: string; result: NextGenInsertResult }> }>,
        };

        const errors: string[] = [];
        let posCreated = 0;
        let linesCreated = 0;
        let sizesUpdated = 0;

        for (const po of normalizedPoData) {
            // 1. Insert PO header
            const header = po.header;

            // Resolve dropdown values to numeric IDs
            const currencyId = resolveId(header.currency || 'USD', CURRENCY_MAP) || '1';
            const statusId = resolveId(header.status || 'Confirmed', STATUS_MAP) || '2';
            const templateId = resolveId(header.ordersTemplate, TEMPLATE_MAP);
            const transportMethodId = resolveId(header.transportMethod, TRANSPORT_METHOD_MAP);
            const locationId = resolveId(header.transportLocation, LOCATION_MAP);

            // Resolve Customer and Supplier names to IDs (virtualized dropdowns)
            let customerId = '';
            let supplierId = '';
            if (header.customer) {
                // Skip purely numeric SAP sold-to party IDs — they are not
                // NextGen customer names and will cause 500 errors.
                if (/^\d+$/.test(header.customer.trim())) {
                    console.log('[nextgen-upload] customer is numeric SAP ID, skipping CustomerName:', header.customer);
                } else {
                    customerId = await client.resolveEntityId(header.customer, 'customer');
                    console.log('[nextgen-upload] resolved customer:', header.customer, '->', customerId || '(not found)');
                }
            }
            if (header.productSupplier) {
                supplierId = await client.resolveEntityId(header.productSupplier, 'supplier');
                console.log('[nextgen-upload] resolved supplier:', header.productSupplier, '->', supplierId || '(not found)');
            }

            // NextGen requires a SupplierName on PO headers — if the buy file's
            // supplier doesn't exist in the test env, fall back to a known test
            // supplier ("PT. UWU JUMP INDONESIA", Id=13).
            if (!supplierId) {
                supplierId = '13';
                console.log('[nextgen-upload] supplier not found, using default test supplier Id=13');
            }

            const poFields: Record<string, string | number | null> = {
                Name: header.purchaseOrder,
                CurrencyName: currencyId,
                DefaultDeliveryDate: header.keyDate ? formatDate(header.keyDate) : '',
                KeyDate: header.keyDate ? formatDate(header.keyDate) : '',
                TemplateName: templateId,
                StatusName: statusId,
                Comments: header.comments || '',
            };

            // Only include optional dropdown fields if they have values.
            // Sending empty strings for virtualized dropdown fields (SupplierName,
            // CustomerName, TransportMethodName, LocationName) causes HTTP 500.
            if (customerId) poFields['CustomerName'] = customerId;
            if (supplierId) poFields['SupplierName'] = supplierId;
            if (transportMethodId) poFields['TransportMethodName'] = transportMethodId;
            if (locationId) poFields['LocationName'] = locationId;

            // Add KeyUsers if present
            if (header.keyUser1) poFields['KeyUser1'] = header.keyUser1;
            if (header.keyUser2) poFields['KeyUser2'] = header.keyUser2;
            if (header.keyUser3) poFields['KeyUser3'] = header.keyUser3;
            if (header.keyUser4) poFields['KeyUser4'] = header.keyUser4;
            if (header.keyUser5) poFields['KeyUser5'] = header.keyUser5;

            console.log('[nextgen-upload] inserting PO header:', header.purchaseOrder, {
                customerId, supplierId, currencyId, statusId, templateId, transportMethodId, locationId,
            });
            const poResult = await client.insertPO(poFields);

            const poEntry = {
                poNumber: header.purchaseOrder,
                result: poResult,
                lines: [] as Array<{ lineItem: number; result: NextGenInsertResult }>,
                sizes: [] as Array<{ lineItem: number; sizeName: string; result: NextGenInsertResult }>,
                missingProducts: [] as string[],
            };

            if (!poResult.success) {
                errors.push(`PO ${header.purchaseOrder}: ${poResult.error}`);
                allResults.pos.push(poEntry);
                continue;
            }

            posCreated++;
            currentStep++;
            updateProgress(progressId, { current: currentStep, poNumber: header.purchaseOrder, message: `PO ${header.purchaseOrder} created` });
            const orderId = poResult.createdId;
            if (!orderId) {
                errors.push(`PO ${header.purchaseOrder}: created but no ID returned — cannot insert lines.`);
                allResults.pos.push(poEntry);
                continue;
            }

            console.log('[nextgen-upload] PO created, orderId:', orderId);

            // Track missing products for logging and reporting
            const missingProducts = new Set<string>();

            // Resolve a fallback product once per PO for when the real product
            // isn't found in the test environment. We search for a known test
            // product (M88F25-029) and cache its Id/Name.
            let fallbackProduct: { id: number; name: string } | null = null;

            // 2. Insert lines for this PO
            for (let i = 0; i < po.lines.length; i++) {
                const line = po.lines[i];
                updateProgress(progressId, {
                    current: currentStep,
                    lineItem: `Line ${line.lineItem} (${i + 1}/${po.lines.length})`,
                    message: `PO ${header.purchaseOrder}: creating line ${i + 1}/${po.lines.length}`,
                });

                // Build line fields using the frGrid insert format:
                // {data:{rowValues:[{fieldValues:[{Field,Value}]}]}}
                const lineFields: Record<string, any> = {};

                // OrderId is required
                lineFields['OrderId'] = orderId;

                // Product/Commodity — search for the product by style number
                // The frGrid requires both CommodityName and CommodityId
                if (line.styleNumber) {
                    const product = await client.findProduct(line.styleNumber);
                    if (product) {
                        lineFields['CommodityName'] = product.name;
                        lineFields['CommodityId'] = product.id;
                    } else {
                        // Product not found in test env — log it and use fallback product
                        // so lines/sizes can still be created for testing.
                        missingProducts.add(line.styleNumber);
                        (poEntry as any).missingProducts.push(line.styleNumber);
                        console.warn(`[nextgen-upload] ⚠️ Product not found in NextGen: "${line.styleNumber}" (PO ${header.purchaseOrder} Line ${line.lineItem}) — using fallback`);
                        if (!fallbackProduct) {
                            fallbackProduct = await client.findProduct('M88F25-029');
                        }
                        if (fallbackProduct) {
                            lineFields['CommodityName'] = fallbackProduct.name;
                            lineFields['CommodityId'] = fallbackProduct.id;
                            console.log('[nextgen-upload] product not found, using fallback:',
                                line.styleNumber, '->', fallbackProduct.name, '(Id:', fallbackProduct.id + ')');
                            // Store original style + color in Comments since we're using a fallback product
                            const parts: string[] = [`Style: ${line.styleNumber}`];
                            if (line.colour || line.color) parts.push(`Color: ${line.colour || line.color}`);
                            if (line.styleColor) parts.push(`Code: ${line.styleColor}`);
                            lineFields['Comments'] = parts.join(' | ');
                        } else {
                            errors.push(`PO ${header.purchaseOrder} Line ${line.lineItem}: product "${line.styleNumber}" not found and no fallback available`);
                            continue;
                        }
                    }
                }

                // DeliveryDate is REQUIRED by NextGen — fall back to a sensible
                // default if neither line nor header has a date.
                const dateForLine = line.cancelDate || line.exFtyDate || header.keyDate || new Date().toISOString().split('T')[0];
                lineFields['DeliveryDate'] = formatDate(dateForLine);

                // Key date
                if (header.keyDate) lineFields['KeyDate'] = formatDate(header.keyDate);

                // Purchase price
                if (line.cost != null) lineFields['PurchasePrice'] = line.cost;

                // Comments: include color info since NextGen's color field is
                // read-only (derived from product master). Format:
                // "Color: TNF Black (E6Q) | Style: NF0A8CGZ"
                if (!lineFields['Comments']) {
                    const parts: string[] = [];
                    if (line.colour || line.color) parts.push(`Color: ${line.colour || line.color}`);
                    if (line.styleColor) parts.push(`Code: ${line.styleColor}`);
                    if (parts.length) lineFields['Comments'] = parts.join(' | ');
                }

                // UDF fields if present (using the actual field names from the grid)
                if (line.buyerPoNumber) lineFields['PrimaryUserDefinedFieldValuesTextUdf1'] = line.buyerPoNumber;
                if (line.startDate) lineFields['PrimaryUserDefinedFieldValuesDateUdf1'] = formatDate(line.startDate);
                if (line.cancelDate) lineFields['PrimaryUserDefinedFieldValuesDateUdf2'] = formatDate(line.cancelDate);

                console.log('[nextgen-upload] inserting line', line.lineItem, 'for PO', header.purchaseOrder, 'fields:', Object.keys(lineFields).join(','));
                const lineResult = await client.insertLine(orderId, lineFields);
                poEntry.lines.push({ lineItem: line.lineItem, result: lineResult });

                if (!lineResult.success) {
                    errors.push(`PO ${header.purchaseOrder} Line ${line.lineItem}: ${lineResult.error}`);
                    continue;
                }

                linesCreated++;
                currentStep++;
                updateProgress(progressId, { current: currentStep, message: `PO ${header.purchaseOrder}: line ${line.lineItem} created` });
                const lineId = lineResult.createdId;
                if (!lineId) {
                    errors.push(`PO ${header.purchaseOrder} Line ${line.lineItem}: created but no ID returned — cannot insert sizes.`);
                    continue;
                }

                // 3. Insert sizes for this line
                // sizes can be Record<number, POSize[]> or Array<Array<POSize>>
                let sizes: any[] = [];
                if (Array.isArray(po.sizes)) {
                    // Array format: po.sizes[i] is an array of sizes for line i
                    sizes = po.sizes[i] || po.sizes[line.lineItem - 1] || [];
                } else {
                    // Record format: po.sizes[lineItem] or po.sizes[i+1]
                    sizes = po.sizes[line.lineItem] || po.sizes[i + 1] || [];
                }
                for (const size of sizes) {
                    // The quantity grid expects SizeName and Quantity
                    // Normalize size names: "OS", "0OS", "One Size" -> "OS" is not
                    // always valid in NextGen. Map to "One Size" which is standard.
                    let rawSize = String(size.productSize || size.sizeName || size.size || '').trim();
                    if (!rawSize || rawSize === '0' || /^(0?OS|ONE\s*SIZE)$/i.test(rawSize)) {
                        rawSize = 'One Size';
                    }
                    const sizeFields: Record<string, any> = {
                        SizeName: rawSize,
                        Quantity: size.quantity,
                    };

                    // NOTE: NextGen's Colour/ColourExt fields in the size grid are
                    // read-only — they're derived from the Product master data, not
                    // settable via the API. Sending them doesn't cause errors but the
                    // values are silently ignored. Color info is preserved in the
                    // line's Comments field instead (set during line insert).

                    console.log('[nextgen-upload] updating quantity for line', line.lineItem, 'size', sizeFields.SizeName);
                    const sizeResult = await client.updateQuantity(orderId, lineId, sizeFields);
                    poEntry.sizes.push({
                        lineItem: line.lineItem,
                        sizeName: sizeFields.SizeName,
                        result: sizeResult,
                    });

                    if (!sizeResult.success) {
                        errors.push(`PO ${header.purchaseOrder} Line ${line.lineItem} Size ${sizeFields.SizeName}: ${sizeResult.error}`);
                    } else {
                        sizesUpdated++;
                        currentStep++;
                        updateProgress(progressId, { current: currentStep, message: `PO ${header.purchaseOrder}: size ${sizeFields.SizeName} updated` });
                    }
                }
            }

            allResults.pos.push(poEntry);
        }

        const success = errors.length === 0;
        console.log('[nextgen-upload] complete:', {
            posCreated, linesCreated, sizesUpdated, errorCount: errors.length,
        });

        // Collect all missing products across POs
        const allMissingProducts = [...new Set(
            allResults.pos.flatMap(po =>
                (po as any).missingProducts || []
            )
        )];
        if (allMissingProducts.length > 0) {
            console.warn(`[nextgen-upload] ⚠️ ${allMissingProducts.length} product(s) not found in NextGen:`,
                allMissingProducts.join(', '));
        }

        updateProgress(progressId, {
            current: totalSteps,
            status: success ? 'complete' : 'error',
            message: success
                ? `Upload complete: ${posCreated} POs, ${linesCreated} lines, ${sizesUpdated} sizes`
                : `Upload completed with ${errors.length} errors`,
            errors,
        });

        return NextResponse.json({
            success,
            progressId,
            summary: {
                posCreated,
                linesCreated,
                sizesUpdated,
                errors: errors.length,
                missingProducts: allMissingProducts,
            },
            errors,
            missingProducts: allMissingProducts,
            results: allResults.pos,
        });

    } catch (error) {
        const message = error instanceof Error ? error.message : "Unknown error";
        console.error("[nextgen-upload] error:", message);
        updateProgress(progressId, { status: 'error', message, errors: [message] });
        return NextResponse.json({ error: message }, { status: 500 });
    }
}

function formatDate(date: string | Date): string {
    if (!date) return '';
    const d = date instanceof Date ? date : new Date(date);
    if (isNaN(d.getTime())) return String(date);
    // NextGen expects M/d/yyyy format (from kendoDatePicker config)
    return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
}
