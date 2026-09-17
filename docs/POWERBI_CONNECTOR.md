# Power BI Connector — NextGen PO Lines

Source endpoint: **`GET /api/nextgen-po-lines`** (PO Cutting app on the VPS or localhost)

## Endpoint contract (as implemented)

| Param | Type | Default | Notes |
|---|---|---|---|
| `poNumber` | string | — | Exact PO number filter (e.g. `M88PO2488`) |
| `style` | string | — | Substring match across product/style fields |
| `page` | int | 1 | 1-based page number |
| `pageSize` | int | 200 | Rows per page, max 1000 |

**Response:**
```json
{
  "poNumber": null,
  "style": null,
  "page": 2,
  "pageSize": 5,
  "count": 5,
  "totalRows": 5,
  "hasMore": true,
  "lines": [
    {
      "id": "261710",
      "poNumber": "VUOUS0925B",
      "style": "M88130210",
      "color": "VUO - BLACK",
      "size": "One Size",
      "quantity": 1,
      "factory": "PT. UWU JUMP INDONESIA",
      "customer": "Vuori",
      "season": "",
      "unitCost": 12.735,
      "subtotal": null
      // ... plus raw NextGen fields (AR dates, UDFs, prices)
    }
  ]
}
```

Pagination behavior: server passes `page` through to NextGen's Kendo-style
`PurchaseOrder/Read` (sort: newest PO first). `hasMore` is `true` when the page
came back full — loop while `hasMore = true`. Note NextGen can be slow; expect
2–20s per page.

---

## Power Query M — full connector (copy into Power BI)

In Power BI: **Get Data → Blank Query → Advanced Editor**, paste:

```powerquery
let
    // ==== CONFIG ====
    BaseUrl    = "https://po-cutting.5-223-78-194.sslip.io/api/nextgen-po-lines",
    PageSize   = 500,          // 200–1000 recommended; higher = fewer calls
    MaxPages   = 50,           // safety cap (50 × 500 = 25,000 rows)
    OptionalPo = "",           // e.g. "M88PO2488" or ""
    OptionalStyle = ""         // e.g. "M88130210" or ""
    // ===============

    FetchPage = (page as number) =>
        let
            Query = [
                poNumber = OptionalPo,
                style = OptionalStyle,
                page = Text.From(page),
                pageSize = Text.From(PageSize)
            ],
            Url = BaseUrl & "?" & Uri.BuildQueryString(Query),
            Raw = Web.Contents(Url, [Headers=[Accept="application/json"]]),
            Json = Json.Document(Raw)
        in
            Json,

    Pages = List.Generate(
        () => [page = 1, data = FetchPage(1), keepGoing = true],
        each [keepGoing] and [page] <= MaxPages,
        each [page = [page] + 1, data = FetchPage([page] + 1), keepGoing = true],
        each [data]
    ),

    // stop condition: page that is not full → no more data
    AllPages = List.Generate(
        () => [page = 1, data = FetchPage(1)],
        each [page] <= MaxPages and ([data][hasMore] = true or [page] = 1),
        each [page = [page] + 1, data = FetchPage([page] + 1)],
        each [data][lines]
    ),

    Lines = List.Combine(AllPages),
    ToTable = Table.FromList(Lines, Splitter.SplitByNothing(), {"Line"}),
    Expanded = Table.ExpandRecordColumn(ToTable, "Line",
        {"id", "poNumber", "style", "color", "size", "quantity",
         "factory", "customer", "season", "unitCost", "subtotal"},
        {"Line Id", "PO Number", "Style", "Color", "Size", "Quantity",
         "Factory", "Customer", "Season", "Unit Cost", "Subtotal"})
in
    Expanded
```

> Note: if you also want the raw NextGen fields (AR dates, UDFs), replace the
> column list in `Table.ExpandRecordColumn` — or expand dynamically via the
> Power BI UI (the connector returns every raw field on each line).

### Simpler variant (no pagination — first page only, up to 1000 rows)

```powerquery
let
    Source = Json.Document(Web.Contents(
        "https://po-cutting.5-223-78-194.sslip.io/api/nextgen-po-lines?pageSize=1000"
    )),
    Lines = Source[lines],
    ToTable = Table.FromList(Lines, Splitter.SplitByNothing(), {"Line"})
in
    ToTable
```

---

## Authentication & refresh behavior

- **No API auth today** — the endpoint only authenticates *itself* against
  NextGen (server-side session, auto re-login). Power BI needs no credentials
  for anonymous calls → choose **Anonymous** when Power BI asks.
- **Recommended hardening before exposing publicly**: put the VPS endpoint
  behind an API key (e.g. require `x-api-key` header checked in the route) —
  then in Power BI use **Web API** credential with the key.
- **Refresh**: scheduled refresh works as-is; each refresh re-paginates from
  page 1. For 28K+ rows expect ~30–60 page calls at pageSize 500 — consider
  pageSize 1000 to halve that.
- **Incremental-ish approach**: filter by `poNumber` per PO for small targeted
  queries; the dump-style full snapshot is better served by the
  `PO Line Data Dump.xlsx` upload into the diff/DB layer.

## Timeout tips

- NextGen occasionally exceeds 20s → the route returns HTTP 500 with
  `NextGen request timed out`. In Power Query, wrap `Web.Contents` in
  `Value.WaitFor`-style retry or simply re-refresh; Power BI's default retry
  usually succeeds on the second attempt.
- Raise the app timeout with `NEXTGEN_REQUEST_TIMEOUT_MS=45000` in the VPS
  `.env.local` if timeouts are frequent.
