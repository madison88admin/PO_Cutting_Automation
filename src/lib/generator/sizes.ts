import ExcelJS from 'exceljs';
import { ProductData } from '@/lib/types/buy-file';

export function generateSizesSheet(products: ProductData[]): ExcelJS.Workbook {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('ORDER_SIZES');

    sheet.columns = [
        { header: 'PurchaseOrder', key: 'purchaseOrder' },
        { header: 'LineItem', key: 'lineItem' },
        { header: 'Range', key: 'range' },
        { header: 'Product', key: 'product' },
        { header: 'SizeName', key: 'sizeName' },
        { header: 'ProductSize', key: 'productSize' },
        { header: 'Quantity', key: 'quantity' },
        { header: 'Colour', key: 'colour' },
        { header: 'Customer', key: 'customer' },
        { header: 'Department', key: 'department' },
        { header: 'CustomAttribute1', key: 'customAttribute1' },
        { header: 'CustomAttribute2', key: 'customAttribute2' },
        { header: 'CustomAttribute3', key: 'customAttribute3' },
        { header: 'LineRatio', key: 'lineRatio' },
        { header: 'ColourExt', key: 'colourExt' },
        { header: 'CustomerExt', key: 'customerExt' },
        { header: 'DepartmentExt', key: 'departmentExt' },
        { header: 'CustomAttribute1Ext', key: 'customAttribute1Ext' },
        { header: 'CustomAttribute2Ext', key: 'customAttribute2Ext' },
        { header: 'CustomAttribute3Ext', key: 'customAttribute3Ext' },
        { header: 'ProductExternalRef', key: 'productExternalRef' },
        { header: 'ProductCustomerRef', key: 'productCustomerRef' },
        { header: 'FindField_Colour', key: 'findField_Colour' },
        { header: 'FindField_Customer', key: 'findField_Customer' },
        { header: 'FindField_Department', key: 'findField_Department' },
        { header: 'FindField_CustomAttribute1', key: 'findField_CustomAttribute1' },
        { header: 'FindField_CustomAttribute2', key: 'findField_CustomAttribute2' },
        { header: 'FindField_CustomAttribute3', key: 'findField_CustomAttribute3' },
        { header: 'FindField_Product', key: 'findField_Product' },
    ];

    const grouped: Record<string, ProductData[]> = {};
    for (const product of products) {
        const po = product.poNumber || 'UNKNOWN';
        if (!grouped[po]) grouped[po] = [];
        grouped[po].push(product);
    }

    for (const poNumber of Object.keys(grouped)) {
        const groupProducts = grouped[poNumber];
        groupProducts.forEach((product, lineIdx) => {
            const lineItem = lineIdx + 1;
            product.sizes.forEach((size) => {
                sheet.addRow({
                    purchaseOrder: poNumber,
                    lineItem,
                    range: product.productRange || '',
                    product: product.product || product.style || '',
                    sizeName: size.size,
                    productSize: size.size,
                    quantity: size.quantity,
                    colour: product.colorName || '',
                    customer: product.customer || '',
                    department: product.department || '',
                    customAttribute1: '',
                    customAttribute2: '',
                    customAttribute3: '',
                    lineRatio: '',
                    colourExt: product.colorExt || '',
                    customerExt: '',
                    departmentExt: '',
                    customAttribute1Ext: '',
                    customAttribute2Ext: '',
                    customAttribute3Ext: '',
                    productExternalRef: product.productExternalRef || '',
                    productCustomerRef: product.productCustomerRef || product.style || '',
                    findField_Colour: product.colorName || '',
                    findField_Customer: product.customer || '',
                    findField_Department: product.department || '',
                    findField_CustomAttribute1: '',
                    findField_CustomAttribute2: '',
                    findField_CustomAttribute3: '',
                    findField_Product: product.product || '',
                });
            });
        });
    }

    return workbook;
}
