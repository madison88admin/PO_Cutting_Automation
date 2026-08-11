import ExcelJS from 'exceljs';
import { ProductData } from '@/lib/types/buy-file';

export function generateLinesSheet(products: ProductData[]): ExcelJS.Workbook {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('LINES');

    sheet.columns = [
        { header: 'PurchaseOrder', key: 'purchaseOrder' },
        { header: 'LineItem', key: 'lineItem' },
        { header: 'ProductRange', key: 'productRange' },
        { header: 'Product', key: 'product' },
        { header: 'Customer', key: 'customer' },
        { header: 'DeliveryDate', key: 'deliveryDate' },
        { header: 'TransportMethod', key: 'transportMethod' },
        { header: 'TransportLocation', key: 'transportLocation' },
        { header: 'Status', key: 'status' },
        { header: 'PurchasePrice', key: 'purchasePrice' },
        { header: 'SellingPrice', key: 'sellingPrice' },
        { header: 'Template', key: 'template' },
        { header: 'KeyDate', key: 'keyDate' },
        { header: 'SupplierProfile', key: 'supplierProfile' },
        { header: 'ClosedDate', key: 'closedDate' },
        { header: 'Comments', key: 'comments' },
        { header: 'Currency', key: 'currency' },
        { header: 'ArchiveDate', key: 'archiveDate' },
        { header: 'ProductExternalRef', key: 'productExternalRef' },
        { header: 'ProductCustomerRef', key: 'productCustomerRef' },
        { header: 'PurchaseUOM', key: 'purchaseUOM' },
        { header: 'SellingUOM', key: 'sellingUOM' },
        { header: 'UDF-buyer_po_number', key: 'udfBuyerPoNumber' },
        { header: 'UDF-start_date', key: 'udfStartDate' },
        { header: 'UDF-canel_date', key: 'udfCanelDate' },
        { header: 'UDF-Inspection result', key: 'udfInspectionResult' },
        { header: 'UDF-Report Type', key: 'udfReportType' },
        { header: 'UDF-Inspector', key: 'udfInspector' },
        { header: 'UDF-Approval Status', key: 'udfApprovalStatus' },
        { header: 'UDF-Submitted inspection date', key: 'udfSubmittedInspectionDate' },
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
        groupProducts.forEach((product, idx) => {
            sheet.addRow({
                purchaseOrder: poNumber,
                lineItem: idx + 1,
                productRange: product.productRange || '',
                product: product.product || product.style || '',
                customer: product.customer || '',
                deliveryDate: product.deliveryDate || '',
                transportMethod: '',
                transportLocation: '',
                status: 'Confirmed',
                purchasePrice: product.unitCost || '',
                sellingPrice: '',
                template: '',
                keyDate: product.deliveryDate || new Date().toISOString().split('T')[0],
                supplierProfile: product.supplierProfile || '',
                closedDate: '',
                comments: '',
                currency: product.currency || 'USD',
                archiveDate: '',
                productExternalRef: product.productExternalRef || '',
                productCustomerRef: product.productCustomerRef || product.style || '',
                purchaseUOM: product.purchaseUOM || 'PCS',
                sellingUOM: product.sellingUOM || 'PCS',
                udfBuyerPoNumber: product.poNumber || '',
                udfStartDate: '',
                udfCanelDate: '',
                udfInspectionResult: '',
                udfReportType: '',
                udfInspector: '',
                udfApprovalStatus: '',
                udfSubmittedInspectionDate: '',
                findField_Product: product.product || '',
            });
        });
    }

    return workbook;
}
