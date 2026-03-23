# Madison88 TNF Mapping (MAR 2026)

Source buy file: `MAR - Buy File Indonesia MADISON 88, LTD.-2026-03-17.xlsx` (sheet `BUY FILE`)
Source product file: `Product Shi's Export_Mar.xlsx` (sheet `FRS.FRNG_M88.Product.Shi's Expo`)

## Buy File -> Internal Keys

| Internal key | Buy file header |
|---|---|
| po | Master PO# |
| buyer_po_number | Final PO Cut# |
| plant | Plant |
| product | Style# |
| product_alt | Style# |
| colour | Material |
| qty | Revised Qty (0 if cancel, new qty if top up or reduce) |
| season | Season Indicator |
| buy_date | Order date |
| orig_ex_fac | Final CRD (Order Date + LT1) |
| trans_cond | Shipment Mode |
| vendor_code | Final Factory |
| vendor_name | Final Vendor Name |
| category | MHP Capacity Type |
| status | Production surcharge confirmation status |

## Product Sheet -> PLM Keys

| PLM key | Product export header |
|---|---|
| buyer_style_number | Buyer Style Number |
| colour | Color Name |
| product_name | Product Name |
| customer_name | Customer Name |
| factory | Factory |
| cost | Cost |

## Lookup Logic (used for this mapping)

- Style lookup uses normalized TNF style key, for example `NF0A887V` -> `A887V`.
- Colour lookup supports TNF coded colour tokens from either side, for example:
  - PLM colour `TNF-VQ2-Dimmed Algae-Pacific Teal` -> `VQ2`
  - Buy material `NF0A887VVQ2` -> `VQ2`
- Transport normalization includes:
  - `V` -> `Sea`
  - `Private Parcel` -> `Courier`

## Run Command

```powershell
$env:PYTHONIOENCODING='utf-8'
python ng_excel_generator.py \
  --input "c:\Users\jcmad\Downloads\MAR - Buy File Indonesia MADISON 88, LTD.-2026-03-17.xlsx" \
  --product-sheet "c:\Users\jcmad\Downloads\Product Shi's Export_Mar.xlsx" \
  --output-dir "c:\Users\jcmad\Desktop\PO Line\_check_out\mapping_test"
```

Generated outputs:
- `_check_out/mapping_test/orders.xlsx`
- `_check_out/mapping_test/lines.xlsx`
- `_check_out/mapping_test/order_sizes.xlsx`
