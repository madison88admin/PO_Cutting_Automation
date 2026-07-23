import { BuyFileItem, NextGenStyleInfo, ProductData, ProductSize } from '@/lib/types/buy-file';

export function mergeBuyFileWithNextGen(
    items: BuyFileItem[],
    nextgen: Record<string, NextGenStyleInfo | null>
): ProductData[] {
    const groups = new Map<string, { items: BuyFileItem[]; style: string }>();

    for (const item of items) {
        const style = String(item.style || '').trim();
        const color = String(item.color || item.colorCode || '').trim();
        const po = String(item.poNumber || '').trim();
        const key = `${style}|${color}|${po}`;

        if (!groups.has(key)) {
            groups.set(key, { items: [], style });
        }
        groups.get(key)!.items.push(item);
    }

    const results: ProductData[] = [];
    for (const { items: groupItems, style } of groups.values()) {
        const styleKey = style.toLowerCase();
        const colorKey = String(groupItems[0]?.colorCode || groupItems[0]?.color || '').toLowerCase();
        const ng = nextgen[`${styleKey}|${colorKey}`] || nextgen[styleKey];
        const first = groupItems[0];

        const sizes: ProductSize[] = groupItems
            .map((item) => ({
                size: String(item.size || 'One Size').trim() || 'One Size',
                quantity: item.quantity || 0,
            }))
            .filter((s) => s.quantity > 0);

        if (!sizes.length) continue;

        results.push({
            style,
            product: ng?.product || first.product || null,
            productRange: ng?.productRange || null,
            productExternalRef: ng?.productExternalRef || first.productExternalRef || first.sku || null,
            productCustomerRef: ng?.productCustomerRef || style,
            styleName: ng?.styleName || first.styleName || null,
            brand: ng?.brand || null,
            season: ng?.season || first.season || null,
            department: ng?.department || null,
            colorName: ng?.colorName || first.color || null,
            colorCode: ng?.colorCode || first.colorCode || null,
            colorExt: ng?.colorExt || null,
            supplierProfile: ng?.supplierProfile || null,
            customer: ng?.customer || first.customer || null,
            factory: ng?.factory || first.factory || null,
            currency: ng?.currency || first.currency || 'USD',
            purchaseUOM: ng?.purchaseUOM || 'PCS',
            sellingUOM: ng?.sellingUOM || 'PCS',
            poNumber: first.poNumber || null,
            deliveryDate: first.deliveryDate || null,
            unitCost: first.unitCost || null,
            costingReference: first.costingReference || null,
            matchStatus: first.matchStatus,
            matchScore: first.matchScore,
            matchReason: first.matchReason,
            sizes,
        });
    }

    return results;
}
