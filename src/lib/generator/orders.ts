import ExcelJS from 'exceljs';
import { ProductData } from '@/lib/types/buy-file';

export function generateOrdersSheet(products: ProductData[]): ExcelJS.Workbook {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('ORDERS');

    sheet.columns = [
        { header: 'PurchaseOrder', key: 'purchaseOrder' },
        { header: 'ProductSupplier', key: 'productSupplier' },
        { header: 'Status', key: 'status' },
        { header: 'Customer', key: 'customer' },
        { header: 'TransportMethod', key: 'transportMethod' },
        { header: 'TransportLocation', key: 'transportLocation' },
        { header: 'PaymentTerm', key: 'paymentTerm' },
        { header: 'Template', key: 'template' },
        { header: 'KeyDate', key: 'keyDate' },
        { header: 'ClosedDate', key: 'closedDate' },
        { header: 'DefaultDeliveryDate', key: 'defaultDeliveryDate' },
        { header: 'Comments', key: 'comments' },
        { header: 'Currency', key: 'currency' },
        { header: 'KeyUser1', key: 'keyUser1' },
        { header: 'KeyUser2', key: 'keyUser2' },
        { header: 'KeyUser3', key: 'keyUser3' },
        { header: 'KeyUser4', key: 'keyUser4' },
        { header: 'KeyUser5', key: 'keyUser5' },
        { header: 'KeyUser6', key: 'keyUser6' },
        { header: 'KeyUser7', key: 'keyUser7' },
        { header: 'KeyUser8', key: 'keyUser8' },
        { header: 'ArchiveDate', key: 'archiveDate' },
        { header: 'PurchaseUOM', key: 'purchaseUOM' },
        { header: 'SellingUOM', key: 'sellingUOM' },
        { header: 'ProductSupplierExt', key: 'productSupplierExt' },
        { header: 'FindField_ProductSupplier', key: 'findField_ProductSupplier' },
    ];

    const grouped = groupByPo(products);
    for (const poNumber of Object.keys(grouped)) {
        const groupProducts = grouped[poNumber];
        const first = groupProducts[0];
        sheet.addRow({
            purchaseOrder: poNumber || first?.poNumber || '',
            productSupplier: first?.supplierProfile || first?.factory || '',
            status: 'Confirmed',
            customer: first?.customer || '',
            transportMethod: '',
            transportLocation: '',
            paymentTerm: '',
            template: '',
            keyDate: first?.deliveryDate || new Date().toISOString().split('T')[0],
            closedDate: '',
            defaultDeliveryDate: first?.deliveryDate || '',
            comments: '',
            currency: first?.currency || 'USD',
            keyUser1: '',
            keyUser2: '',
            keyUser3: '',
            keyUser4: '',
            keyUser5: '',
            keyUser6: '',
            keyUser7: '',
            keyUser8: '',
            archiveDate: '',
            purchaseUOM: first?.purchaseUOM || 'PCS',
            sellingUOM: first?.sellingUOM || 'PCS',
            productSupplierExt: '',
            findField_ProductSupplier: first?.supplierProfile || '',
        });
    }

    return workbook;
}

function groupByPo(products: ProductData[]): Record<string, ProductData[]> {
    const groups: Record<string, ProductData[]> = {};
    for (const product of products) {
        const po = product.poNumber || 'UNKNOWN';
        if (!groups[po]) groups[po] = [];
        groups[po].push(product);
    }
    return groups;
}
