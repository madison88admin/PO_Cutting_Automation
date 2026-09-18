/**
 * NextGen API Client
 * Read-only access for PO validation and lookup
 * Base URL: https://nextgen.madison88.com
 *
 * Simplified flow based on HAR analysis:
 * - Single endpoint: PurchaseOrder/Read (returns line-level data with PO header info)
 * - PO number field: PrimaryUserDefinedFieldValuesTextUdf3
 * - Latest PO: sort by PrimaryUserDefinedFieldValuesTextUdf3-desc
 * - PO lines: returned directly by PurchaseOrder/Read with filter
 */

interface NextGenConfig {
    baseUrl: string;
    username: string;
    password: string;
}

interface NextGenAuth {
    cookie: string;
    requestVerificationToken: string;
}

export interface NextGenPOLine {
    id: string;
    poNumber: string;
    style: string;
    color: string;
    size: string;
    quantity: number;
    factory: string;
    customer: string;
    season: string;
    [key: string]: unknown;
}

export interface NextGenValidationResult {
    poNumber: string;
    exists: boolean;
    lines: NextGenPOLine[];
    matched: NextGenPOLine[];
    missing: { style: string; color: string; size: string; quantity: number }[];
    extra: { style: string; color: string; size: string; quantity: number }[];
}

export interface NextGenInsertField {
    Field: string;
    Value: string | number | null;
}

export interface NextGenInsertResult {
    success: boolean;
    status: number;
    endpoint: string;
    requestBody: NextGenInsertField[];
    responseBody: string;
    createdId?: string | number | null;
    error?: string;
}

export interface NextGenEndpointHint {
    url: string;
    context: string;
    method?: string;
}

const PO_NUMBER_FIELD = 'PrimaryUserDefinedFieldValuesTextUdf3';

export class NextGenClient {
    private config: NextGenConfig;
    private auth: NextGenAuth | null = null;

    constructor(config?: Partial<NextGenConfig>) {
        this.config = {
            baseUrl: config?.baseUrl || process.env.NEXTGEN_BASE_URL || 'https://nextgen.madison88.com',
            username: config?.username || process.env.NEXTGEN_USERNAME || '',
            password: config?.password || process.env.NEXTGEN_PASSWORD || '',
        };
    }

    async fetchWithCookie(url: string, options: RequestInit = {}, allowRetry: boolean = true): Promise<Response> {
        const headers: Record<string, string> = {
            'Content-Type': 'application/x-www-form-urlencoded',
            'Accept': 'application/json, text/html, */*',
            ...((options.headers as Record<string, string>) || {}),
        };

        if (this.auth?.cookie) {
            headers['Cookie'] = this.auth.cookie;
        }

        if (this.auth?.requestVerificationToken) {
            headers['RequestVerificationToken'] = this.auth.requestVerificationToken;
        }

        const response = await fetch(url, {
            ...options,
            headers,
            redirect: 'manual',
        });

        // If redirected to login, force re-login and retry once
        if (allowRetry && (response.status === 302 || response.status === 401)) {
            const location = response.headers.get('location') || response.headers.get('Location') || '';
            if (location.includes('/Account/Login') || location.includes('Login')) {
                console.log('[nextgen] Session expired, forcing re-login');
                this.auth = null;
                await this.login();
                return this.fetchWithCookie(url, options, false);
            }
        }

        return response;
    }

    private extractCookie(response: Response): string {
        const setCookie = response.headers.get('set-cookie');
        if (!setCookie) return '';
        return setCookie
            .split(/,(?=[^;]*=)/)
            .map(c => c.split(';')[0].trim())
            .filter(c => c)
            .join('; ');
    }

    private extractRequestVerificationToken(html: string): string {
        const match = html.match(/name="__RequestVerificationToken"[^>]*value="([^"]*)"/);
        return match?.[1] || '';
    }

    async login(retries = 2): Promise<NextGenAuth> {
        if (this.auth) return this.auth;

        let lastError: Error | null = null;
        for (let attempt = 1; attempt <= retries; attempt++) {
            try {
                const loginUrl = `${this.config.baseUrl}/Account/Login`;
                let loginPage = await this.fetchWithCookie(loginUrl, { method: 'GET' }, false);

                if (loginPage.status === 302) {
                    const location = loginPage.headers.get('location') || '';
                    console.log('[nextgen] login page returned 302, location:', location);
                    if (location && !location.includes('/Account/Login')) {
                        loginPage = await this.fetchWithCookie(location, { method: 'GET' }, false);
                    } else {
                        throw new Error(`NextGen login page redirect loop: ${location}`);
                    }
                }

                if (!loginPage.ok) {
                    throw new Error(`Failed to load NextGen login page: ${loginPage.status}`);
                }

                const html = await loginPage.text();
                const requestVerificationToken = this.extractRequestVerificationToken(html);
                const cookie = this.extractCookie(loginPage);

                if (!requestVerificationToken) {
                    throw new Error('Could not extract NextGen anti-forgery token');
                }

                const params = new URLSearchParams();
                params.append('__RequestVerificationToken', requestVerificationToken);
                params.append('UserName', this.config.username);
                params.append('Password', this.config.password);
                params.append('RememberMe', 'false');

                const loginResponse = await this.fetchWithCookie(loginUrl, {
                    method: 'POST',
                    body: params.toString(),
                    headers: { 'Cookie': cookie },
                }, false);

                const loginResponseText = await loginResponse.text();
                console.log(`[nextgen] login POST attempt ${attempt} status:`, loginResponse.status);
                console.log('[nextgen] login POST set-cookie:', loginResponse.headers.get('set-cookie'));
                console.log('[nextgen] login POST body preview:', loginResponseText.slice(0, 500));

                const loginCookie = this.extractCookie(loginResponse);
                if (loginCookie) {
                    this.auth = { cookie: loginCookie, requestVerificationToken };
                    return this.auth;
                }

                if (loginResponse.status === 302 && loginResponseText.includes('/Account/Login')) {
                    throw new Error('NextGen login failed: credentials rejected (redirected back to login)');
                }

                lastError = new Error('NextGen login failed: no session cookie returned');
            } catch (err) {
                lastError = err instanceof Error ? err : new Error(String(err));
                console.warn(`[nextgen] login attempt ${attempt} failed:`, lastError.message);
                if (attempt < retries) {
                    await new Promise((resolve) => setTimeout(resolve, 1000 * attempt));
                }
            }
        }

        console.error('[nextgen] all login attempts failed. Check NEXTGEN_USERNAME/NEXTGEN_PASSWORD.');
        throw lastError || new Error('NextGen login failed');
    }

    getPONumberFromRecord(record: any): string {
        return String(record[PO_NUMBER_FIELD] || record.OrderName || '');
    }

    mapToPOLine(line: Record<string, unknown>): NextGenPOLine {
        return {
            id: String(line.Id || line.ID || line.id || ''),
            poNumber: this.getPONumberFromRecord(line),
            style: String(line.CommodityName || line.Style || line.style || line.StyleNumber || line.Product || ''),
            color: String(line.OptionColourName || line.Color || line.color || line.ColorName || line.Colour || ''),
            size: String(line.SizeName || line.Size || line.size || ''),
            quantity: Number(line.Quantity || line.quantity || line.Qty || 0),
            factory: String(line.OrderSupplierName || line.Factory || line.factory || line.Vendor || line.Supplier || ''),
            customer: String(line.CustomerName || line.Customer || line.customer || ''),
            season: String(line.RangeName || line.Range || line.Season || line.season || ''),
            ...line,
        };
    }

    private buildReadParams({ sort, filter, page = 1, pageSize = 100 }: { sort?: string; filter?: string; page?: number; pageSize?: number }) {
        const params = new URLSearchParams();
        params.append('sort', sort || '');
        params.append('group', '');
        params.append('filter', filter || '');
        params.append('page', String(page));
        params.append('pageSize', String(pageSize));
        params.append('aggregates', '');
        return params;
    }

    private async readPurchaseOrder(params: URLSearchParams): Promise<any[]> {
        await this.login();
        const url = `${this.config.baseUrl}/PurchaseOrder/Read`;
        const body = params.toString();
        console.log('[nextgen] PurchaseOrder/Read request body:', body);

        const response = await this.fetchWithCookie(url, { method: 'POST', body });

        const text = await response.text();
        console.log('[nextgen] PurchaseOrder/Read status:', response.status, 'content-type:', response.headers.get('content-type'));
        console.log('[nextgen] PurchaseOrder/Read response preview:', text.substring(0, 500));

        if (!response.ok) {
            throw new Error(`NextGen PurchaseOrder/Read failed: ${response.status} ${text.substring(0, 500)}`);
        }

        try {
            const data = text ? JSON.parse(text) : {};
            return data?.Data || data?.data || [];
        } catch (err) {
            throw new Error(`NextGen PurchaseOrder/Read returned invalid JSON: ${text.substring(0, 500)}`);
        }
    }

    async lookupColorNames(skus: string[]): Promise<Record<string, string | null>> {
        const normalize = (s: string) => s.toLowerCase().trim();
        const targets = skus.map(s => normalize(s)).filter(Boolean);
        if (!targets.length) return {};

        const params = this.buildReadParams({
            sort: `${PO_NUMBER_FIELD}-desc~OrderName-asc`,
            filter: '',
            page: 1,
            pageSize: 500,
        });
        const results = await this.readPurchaseOrder(params);
        console.log('[nextgen] lookupColorNames targets:', targets, 'total rows:', results.length);

        const found: Record<string, string | null> = {};
        for (const row of results) {
            for (const [key, value] of Object.entries(row)) {
                const valStr = normalize(String(value || ''));
                for (const target of targets) {
                    if (found[target] !== undefined) continue;
                    if (valStr === target || valStr.includes(target)) {
                        const colorName = String(row.OptionColourName || '');
                        found[target] = colorName || null;
                        console.log('[nextgen] lookupColorNames found target:', target, 'field:', key, 'colorName:', colorName);
                    }
                }
            }
        }

        for (const target of targets) {
            if (found[target] === undefined) found[target] = null;
        }
        console.log('[nextgen] lookupColorNames results:', found);
        return found;
    }

    async fetchRecentRecords(pageSize: number = 500): Promise<any[]> {
        const params = this.buildReadParams({
            sort: `${PO_NUMBER_FIELD}-desc~OrderName-asc`,
            filter: '',
            page: 1,
            pageSize,
        });
        return this.readPurchaseOrder(params);
    }

    async getLatestPO(): Promise<{ poNumber: string; id: string } | null> {
        const params = this.buildReadParams({
            sort: `${PO_NUMBER_FIELD}-desc~OrderName-asc`,
            page: 1,
            pageSize: 1,
        });
        const results = await this.readPurchaseOrder(params);
        const match = results[0];
        if (!match) return null;
        const poNumber = this.getPONumberFromRecord(match);
        console.log('[nextgen] Latest PO found:', poNumber, 'OrderId:', match.OrderId, 'OrderName:', match.OrderName);
        return {
            id: String(match.OrderId || match.Id || match.ID || match.id || ''),
            poNumber,
        };
    }

    async validatePO(poNumber: string, uploadLines: { style: string; color: string; size: string; quantity: number }[]): Promise<NextGenValidationResult> {
        if (!uploadLines.length) {
            return {
                poNumber,
                exists: false,
                lines: [],
                matched: [],
                missing: [],
                extra: [],
            };
        }

        // The auto-incremented PO number is not yet in NextGen.
        // Search by style from the first upload line (NextGen color names differ from uploaded color codes).
        const target = uploadLines[0];
        const normalize = (s: string) => s.toLowerCase().trim();
        const targetStyle = normalize(target.style);

        // Fetch recent records and filter in code for matching style
        const params = this.buildReadParams({
            sort: `${PO_NUMBER_FIELD}-desc~OrderName-asc`,
            filter: '',
            page: 1,
            pageSize: 500,
        });
        const results = await this.readPurchaseOrder(params);
        console.log('[nextgen] validatePO target style:', targetStyle, 'uploaded color:', normalize(target.color), 'total rows:', results.length);
        console.log('[nextgen] sample CommodityName values:', results.slice(0, 5).map((r: any) => r.CommodityName));

        // Try matching by CommodityName first, then fallback to any field containing the style
        let matchingRows = results.filter((row: any) =>
            normalize(String(row.CommodityName || '')) === targetStyle
        );

        if (!matchingRows.length && targetStyle) {
            const fieldHits: Record<string, number> = {};
            for (const row of results) {
                for (const [key, value] of Object.entries(row)) {
                    const valStr = normalize(String(value || ''));
                    if (valStr === targetStyle || valStr.includes(targetStyle)) {
                        fieldHits[key] = (fieldHits[key] || 0) + 1;
                    }
                }
            }
            console.log('[nextgen] field hits for target style:', fieldHits);

            const bestField = Object.entries(fieldHits).sort((a, b) => b[1] - a[1])[0]?.[0];
            if (bestField) {
                matchingRows = results.filter((row: any) => {
                    const valStr = normalize(String(row[bestField] || ''));
                    return valStr === targetStyle || valStr.includes(targetStyle);
                });
            }
        }

        console.log('[nextgen] matching rows found:', matchingRows.length);

        if (!matchingRows.length) {
            return {
                poNumber,
                exists: false,
                lines: [],
                matched: [],
                missing: [],
                extra: [],
            };
        }

        const nextgenLines = matchingRows.map((line) => this.mapToPOLine(line));

        const matchKey = (a: { style: string; color: string; size: string }, b: NextGenPOLine) => {
            return normalize(a.style) === normalize(b.style) &&
                normalize(a.color) === normalize(b.color) &&
                normalize(a.size) === normalize(b.size);
        };

        const matched: NextGenPOLine[] = [];
        const missing: { style: string; color: string; size: string; quantity: number }[] = [];
        const nextgenMatched = new Set<number>();

        for (const uploadLine of uploadLines) {
            const matchIndex = nextgenLines.findIndex((ng, idx) => !nextgenMatched.has(idx) && matchKey(uploadLine, ng));
            if (matchIndex >= 0) {
                matched.push(nextgenLines[matchIndex]);
                nextgenMatched.add(matchIndex);
            } else {
                missing.push(uploadLine);
            }
        }

        const extra = nextgenLines.filter((_, idx) => !nextgenMatched.has(idx));

        return {
            poNumber,
            exists: true,
            lines: nextgenLines,
            matched,
            missing,
            extra,
        };
    }

    // ─── Entity resolution (for virtualized dropdowns) ──────────────────

    /**
     * Resolve a Customer or Supplier name to its numeric ID.
     *
     * The NextGen Insert page uses virtualized Kendo dropdowns for Customer
     * (entity type 3) and Supplier (entity type 17). These call a valueMapper
     * API to resolve text names to IDs.
     *
     * This method searches the PurchaseOrder/Read data for a matching
     * CustomerName/OrderSupplierName and returns the corresponding ID.
     *
     * @param name - The customer or supplier name to resolve
     * @param type - "customer" or "supplier"
     * @returns The numeric ID as a string, or empty string if not found
     */
    async resolveEntityId(name: string, type: 'customer' | 'supplier'): Promise<string> {
        if (!name) return '';
        const trimmed = name.trim();

        await this.login();
        const params = this.buildReadParams({
            sort: `${PO_NUMBER_FIELD}-desc~OrderName-asc`,
            filter: '',
            page: 1,
            pageSize: 500,
        });
        const results = await this.readPurchaseOrder(params);

        const normalize = (s: string) => s.toLowerCase().trim();
        const target = normalize(trimmed);

        for (const row of results) {
            if (type === 'customer') {
                const rowName = normalize(String(row.CustomerName || ''));
                const rowId = String(row.CustomerId || row.CustomerID || row.customerId || '');
                // Match by name (includes/equals) OR by numeric ID if the input is numeric
                if (rowName === target || (rowName && target && (rowName.includes(target) || target.includes(rowName)))) {
                    return rowId;
                }
                if (/^\d+$/.test(trimmed) && rowId === trimmed) {
                    return rowId;
                }
            } else {
                const rowName = normalize(String(row.OrderSupplierName || row.SupplierName || ''));
                const rowId = String(row.OrderSupplierId || row.SupplierId || row.SupplierID || row.supplierId || '');
                if (rowName === target || (rowName && target && (rowName.includes(target) || target.includes(rowName)))) {
                    return rowId;
                }
                if (/^\d+$/.test(trimmed) && rowId === trimmed) {
                    return rowId;
                }
            }
        }
        console.warn(`[nextgen] resolveEntityId: could not find ${type} "${name}"`);
        return '';
    }

    // ─── Write operations (PO Insert) ───────────────────────────────────
    //
    // SAFETY: All write methods enforce that the configured base URL is the
    // TEST environment (must contain ":8443"). They will REFUSE to run
    // against the production site (https://nextgen.madison88.com on port 443).
    // This is a hard guard — it cannot be bypassed by config alone.

    private assertTestEnvironment(): void {
        const url = this.config.baseUrl;
        if (!url.includes(':8443')) {
            throw new Error(
                `SAFETY BLOCK: Write operations are refused against non-test environment.\n` +
                `Configured URL: ${url}\n` +
                `Expected test URL: https://nextgen.madison88.com:8443\n` +
                `Set NEXTGEN_WRITE_BASE_URL to the test environment before attempting writes.`
            );
        }
    }

    /**
     * Insert a Purchase Order header via /PurchaseOrder/InsertFormData.
     *
     * The endpoint expects a JSON body: { "data": [{ Field, Value }, ...] }
     * with Content-Type: application/json and X-Requested-With: XMLHttpRequest.
     *
     * Known fields (from HAR capture + HTML form analysis):
     *   Name (Order Reference), CustomerName, SupplierName, CurrencyName,
     *   TransportMethodName, LocationName (Deliver To), DefaultDeliveryDate,
     *   KeyDate (PO Issue Date), TemplateName, StatusName, Comments
     *
     * Dropdown values use IDs:
     *   Currency: USD=1, RMB=2, IDR=3
     *   Status: Confirmed=2
     *   Template: Major Brand Bulk=6, Specialty Bulk PO=5, SMS PO Header=29
     *   TransportMethod: Courier=1, Sea=2, Air=3, Truck=4, Air/Sea Difference=5,
     *     Boat-Truck=6, Sea-FCL=7
     *   Location: see /PurchaseOrder/Insert page for full list (e.g. USA=36,
     *     Canada=41, China=37, "Not Set"=43)
     *
     * On success, the response should contain the created PO ID, and the
     * browser would redirect to /PurchaseOrder/Edit/{id}.
     */
    async insertPO(fields: Record<string, string | number | null>): Promise<NextGenInsertResult> {
        this.assertTestEnvironment();
        await this.login();

        const data: NextGenInsertField[] = Object.entries(fields).map(([Field, Value]) => ({
            Field,
            Value: Value === null || Value === undefined ? '' : String(Value),
        }));

        const url = `${this.config.baseUrl}/PurchaseOrder/InsertFormData`;
        const body = JSON.stringify({ data });
        console.log('[nextgen] insertPO request:', url, 'fields:', data.map(f => `${f.Field}=${String(f.Value ?? '').slice(0, 50)}`).join(', '));

        const response = await this.fetchWithCookie(url, {
            method: 'POST',
            body,
            headers: {
                'Content-Type': 'application/json',
                'X-Requested-With': 'XMLHttpRequest',
                'Accept': 'application/json, text/html, */*',
            },
        });

        const text = await response.text();
        console.log('[nextgen] insertPO status:', response.status, 'body preview:', text.slice(0, 500));

        // Auto re-login on 401 (session expired) and retry once
        if (response.status === 401) {
            console.log('[nextgen] insertPO got 401, re-logging in and retrying');
            this.auth = null;
            await this.login();
            const retryResp = await this.fetchWithCookie(url, {
                method: 'POST',
                body,
                headers: {
                    'Content-Type': 'application/json',
                    'X-Requested-With': 'XMLHttpRequest',
                    'Accept': 'application/json, text/html, */*',
                },
            });
            const retryText = await retryResp.text();
            console.log('[nextgen] insertPO retry status:', retryResp.status, 'body preview:', retryText.slice(0, 500));
            if (!retryResp.ok) {
                return {
                    success: false,
                    status: retryResp.status,
                    endpoint: url,
                    requestBody: data,
                    responseBody: retryText,
                    error: `InsertPO failed: HTTP ${retryResp.status}`,
                };
            }
            return this.parseInsertPOResponse(retryText, retryResp.status, url, data);
        }

        if (!response.ok) {
            return {
                success: false,
                status: response.status,
                endpoint: url,
                requestBody: data,
                responseBody: text,
                error: `InsertPO failed: HTTP ${response.status}`,
            };
        }

        return this.parseInsertPOResponse(text, response.status, url, data);
    }

    private parseInsertPOResponse(text: string, status: number, url: string, data: NextGenInsertField[]): NextGenInsertResult {
        // Try to extract the created PO ID from the response
        let createdId: string | number | null = null;
        let apiSuccess = true;
        let apiError: string | undefined;
        try {
            const parsed = JSON.parse(text);
            createdId = parsed?.InsertedItemId ?? parsed?.Data?.Id ?? parsed?.Data?.id ?? parsed?.id ?? parsed?.Id ?? parsed?.Data ?? null;
            // NextGen returns HTTP 200 even on logical failures — check the Success field
            if (parsed?.Success === false) {
                apiSuccess = false;
                apiError = parsed?.Items?.map((i: any) => i.Description).filter(Boolean).join('; ')
                    || 'Insert failed (Success: false)';
            }
        } catch {
            // Response may be HTML or non-JSON; try to extract ID from text
            const idMatch = text.match(/"(?:InsertedItemId|Id)"\s*:\s*(\d+)/i);
            if (idMatch) createdId = Number(idMatch[1]);
        }

        return {
            success: apiSuccess,
            status,
            endpoint: url,
            requestBody: data,
            responseBody: text,
            createdId,
            error: apiError,
        };
    }

    /**
     * Insert a Purchase Order LINE via /PurchaseOrder/InsertFormLinesGridRecord.
     *
     * This is a Kendo UI Grid insert endpoint. Kendo grid endpoints expect
     * form-urlencoded data with a "models" parameter containing a JSON array
     * of the records to insert.
     *
     * The line must be associated with a parent PO ID (orderId).
     *
     * Known line fields (from LINES.xlsx mapping):
     *   PurchaseOrder, LineItem, ProductRange, Product, Customer,
     *   DeliveryDate, TransportMethod, TransportLocation, Status,
     *   PurchasePrice, SellingPrice, Template, KeyDate, SupplierProfile,
     *   Comments, Currency, ProductExternalRef, ProductCustomerRef,
     *   PurchaseUOM, SellingUOM, UDF-buyer_po_number, UDF-start_date,
     *   UDF-canel_date, etc.
     */
    async insertLine(orderId: string | number, lineFields: Record<string, any>): Promise<NextGenInsertResult> {
        this.assertTestEnvironment();
        await this.login();

        const url = `${this.config.baseUrl}/PurchaseOrder/InsertFormLinesGridRecord`;
        console.log('[nextgen] insertLine request:', url, 'orderId:', orderId, 'fields:', Object.keys(lineFields).join(','));

        // Visit the Edit page first to set up grid session state
        try {
            await this.fetchWithCookie(`${this.config.baseUrl}/PurchaseOrder/Edit/${orderId}`, { method: 'GET' }, true);
        } catch (err) {
            console.warn('[nextgen] insertLine: could not visit Edit page:', err);
        }

        // The frGrid plugin uses a custom insert format (discovered from grid-helpers.js):
        //   {data:{rowValues:[{fieldValues:[{Field,Value}],rowIdentifier:uid}]}}
        // Content-Type: application/json; charset=utf-8
        const fieldValues = Object.entries(lineFields).map(([Field, Value]) => ({
            Field,
            Value: Value === null || Value === undefined ? '' : Value,
        }));

        const rowIdentifier = `line-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
        const body = JSON.stringify({
            data: {
                rowValues: [{
                    fieldValues,
                    rowIdentifier,
                }],
            },
        });

        const response = await this.fetchWithCookie(url, {
            method: 'POST',
            body,
            headers: {
                'Content-Type': 'application/json; charset=utf-8',
                'X-Requested-With': 'XMLHttpRequest',
                'Accept': 'application/json, text/javascript, */*; q=0.01',
                'Referer': `${this.config.baseUrl}/PurchaseOrder/Edit/${orderId}`,
            },
        });

        const text = await response.text();
        console.log('[nextgen] insertLine status:', response.status, 'body preview:', text.slice(0, 500));

        // Auto re-login on 401 (session expired) and retry once
        if (response.status === 401) {
            console.log('[nextgen] insertLine got 401, re-logging in and retrying');
            this.auth = null;
            await this.login();
            // Re-visit Edit page to set up grid session
            try {
                await this.fetchWithCookie(`${this.config.baseUrl}/PurchaseOrder/Edit/${orderId}`, { method: 'GET' }, true);
            } catch {}
            const retryResp = await this.fetchWithCookie(url, {
                method: 'POST',
                body,
                headers: {
                    'Content-Type': 'application/json; charset=utf-8',
                    'X-Requested-With': 'XMLHttpRequest',
                    'Accept': 'application/json, text/javascript, */*; q=0.01',
                    'Referer': `${this.config.baseUrl}/PurchaseOrder/Edit/${orderId}`,
                },
            });
            const retryText = await retryResp.text();
            console.log('[nextgen] insertLine retry status:', retryResp.status);
            if (!retryResp.ok) {
                return {
                    success: false,
                    status: retryResp.status,
                    endpoint: url,
                    requestBody: [{ Field: 'body', Value: body.slice(0, 500) }],
                    responseBody: retryText,
                    error: `InsertLine failed: HTTP ${retryResp.status}`,
                };
            }
            return this.parseInsertLineResponse(retryText, retryResp.status, url, body);
        }

        if (!response.ok) {
            return {
                success: false,
                status: response.status,
                endpoint: url,
                requestBody: [{ Field: 'body', Value: body.slice(0, 500) }],
                responseBody: text,
                error: `InsertLine failed: HTTP ${response.status}`,
            };
        }

        return this.parseInsertLineResponse(text, response.status, url, body, orderId);
    }

    private async parseInsertLineResponse(
        text: string, status: number, url: string, body: string, orderId?: string | number
    ): Promise<NextGenInsertResult> {
        // Parse the response - check for Success field and extract createdId
        let success = true;
        let createdId: string | number | null = null;
        let errorMsg: string | undefined;
        try {
            const parsed = JSON.parse(text);
            success = parsed.Success !== false;

            // Check for row-level errors
            if (parsed.GridRowErrorDescriptors?.length > 0) {
                const desc = parsed.GridRowErrorDescriptors[0];
                if (desc.RowErrorDescriptor?.Success === false) {
                    success = false;
                    errorMsg = desc.RowErrorDescriptor.Items?.map((i: any) => i.Description).join('; ');
                }
                // Extract the created Id from LatestData if available
                if (desc.LatestData?.Id) {
                    createdId = desc.LatestData.Id;
                }
            }
        } catch {
            const idMatch = text.match(/"(?:Id|ItemId|InsertedItemId)"\s*:\s*(\d+)/i);
            if (idMatch) createdId = Number(idMatch[1]);
        }

        // If insert succeeded but no createdId, read the grid to get the line ID
        if (success && !createdId && orderId) {
            try {
                createdId = await this.getLatestLineId(orderId);
            } catch (err) {
                console.warn('[nextgen] insertLine: could not fetch line ID:', err);
            }
        }

        return {
            success,
            status,
            endpoint: url,
            requestBody: [{ Field: 'body', Value: body.slice(0, 500) }],
            responseBody: text,
            createdId,
            error: errorMsg,
        };
    }

    /**
     * Read the lines grid for a PO and return the latest (highest) line ID.
     * Used after insertLine when the insert response doesn't include the new ID.
     */
    async getLatestLineId(orderId: string | number): Promise<number | null> {
        const url = `${this.config.baseUrl}/PurchaseOrder/FormLinesGridRead`;
        const params = new URLSearchParams({
            sort: 'LineItem-desc',
            page: '1',
            pageSize: '1',
            group: '',
            filter: '',
            aggregates: '',
            OrderId: String(orderId),
        });

        const response = await this.fetchWithCookie(url, {
            method: 'POST',
            body: params.toString(),
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                'X-Requested-With': 'XMLHttpRequest',
            },
        });

        if (!response.ok) {
            console.warn('[nextgen] getLatestLineId: read failed:', response.status);
            return null;
        }

        const data = await response.json();
        const items = data?.Data || [];
        if (items.length === 0) {
            console.warn('[nextgen] getLatestLineId: no lines found for order', orderId);
            return null;
        }

        const lineId = items[0]?.Id ?? items[0]?.ItemId ?? null;
        console.log('[nextgen] getLatestLineId: orderId', orderId, '-> lineId', lineId);
        return lineId;
    }

    /**
     * Search for a product/commodity by name, style number, or external reference.
     * Returns the first matching product with Id and Name.
     */
    async findProduct(searchTerm: string): Promise<{ id: number; name: string } | null> {
        await this.login();

        const url = `${this.config.baseUrl}/Product/ProductsBrowserGridRead`;
        const safeTerm = searchTerm.replace(/'/g, "''");

        // Try multiple filter fields: Name, StyleNumber, ExternalReference
        // Kendo filter format: field~operator~value
        const filter = `Name~contains~'${safeTerm}'~or~StyleNumber~contains~'${safeTerm}'~or~ExternalReference~contains~'${safeTerm}'`;
        const params = new URLSearchParams({
            sort: 'Name-asc',
            page: '1',
            pageSize: '50',
            group: '',
            filter,
            aggregates: '',
        });

        let response = await this.fetchWithCookie(url, {
            method: 'POST',
            body: params.toString(),
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                'X-Requested-With': 'XMLHttpRequest',
            },
        });

        // If the combined filter fails, try just Name
        if (!response.ok) {
            const params2 = new URLSearchParams({
                sort: 'Name-asc',
                page: '1',
                pageSize: '50',
                group: '',
                filter: `Name~contains~'${safeTerm}'`,
                aggregates: '',
            });
            response = await this.fetchWithCookie(url, {
                method: 'POST',
                body: params2.toString(),
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                    'X-Requested-With': 'XMLHttpRequest',
                },
            });
        }

        if (!response.ok) {
            console.warn('[nextgen] findProduct: search failed:', response.status);
            return null;
        }

        const data = await response.json();
        const items = data?.Data || [];
        if (items.length === 0) {
            console.warn('[nextgen] findProduct: no results for', searchTerm);
            return null;
        }

        // Try exact match on any identifier, then first result
        const exact = items.find((p: any) =>
            p.Name === searchTerm || p.StyleNumber === searchTerm || p.ExternalReference === searchTerm
        );
        const match = exact || items[0];
        console.log('[nextgen] findProduct:', searchTerm, '->', match.Name, '(Id:', match.Id + ')');
        return { id: match.Id, name: match.Name };
    }


    /**
     * Update a quantity record (size-level data) via /PurchaseOrder/UpdateQuantityRecord.
     *
     * This endpoint manages the size/quantity breakdown for a specific order line.
     * It expects form-urlencoded data with the quantity details.
     *
     * Known size fields (from ORDER_SIZES.xlsx mapping):
     *   PurchaseOrder, LineItem, Range, Product, SizeName, ProductSize,
     *   Quantity, Colour, Customer, Department, ColourExt, etc.
     */
    async updateQuantity(orderId: string | number, lineId: string | number, sizeFields: Record<string, any>): Promise<NextGenInsertResult> {
        this.assertTestEnvironment();
        await this.login();

        const url = `${this.config.baseUrl}/PurchaseOrder/UpdateQuantityRecord`;
        console.log('[nextgen] updateQuantity request:', url, 'orderId:', orderId, 'lineId:', lineId);

        // Visit the Edit page first to set up session state
        try {
            await this.fetchWithCookie(`${this.config.baseUrl}/PurchaseOrder/Edit/${orderId}`, { method: 'GET' }, true);
        } catch {}

        // The frGrid uses the same format for updates as for inserts:
        // {data:{rowValues:[{fieldValues:[{Field,Value}],rowIdentifier:uid}]}}
        const fieldValues: Array<{ Field: string; Value: any }> = [
            { Field: 'OrderId', Value: orderId },
            { Field: 'Id', Value: lineId },
        ];

        // Add the size fields
        for (const [key, value] of Object.entries(sizeFields)) {
            fieldValues.push({ Field: key, Value: value === null || value === undefined ? '' : value });
        }

        const rowIdentifier = `size-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
        const body = JSON.stringify({
            data: {
                rowValues: [{
                    fieldValues,
                    rowIdentifier,
                }],
            },
        });

        const response = await this.fetchWithCookie(url, {
            method: 'POST',
            body,
            headers: {
                'Content-Type': 'application/json; charset=utf-8',
                'X-Requested-With': 'XMLHttpRequest',
                'Accept': 'application/json, text/javascript, */*; q=0.01',
                'Referer': `${this.config.baseUrl}/PurchaseOrder/Edit/${orderId}`,
            },
        });

        const text = await response.text();
        console.log('[nextgen] updateQuantity status:', response.status, 'body preview:', text.slice(0, 500));

        // Auto re-login on 401 (session expired) and retry once
        if (response.status === 401) {
            console.log('[nextgen] updateQuantity got 401, re-logging in and retrying');
            this.auth = null;
            await this.login();
            try {
                await this.fetchWithCookie(`${this.config.baseUrl}/PurchaseOrder/Edit/${orderId}`, { method: 'GET' }, true);
            } catch {}
            const retryResp = await this.fetchWithCookie(url, {
                method: 'POST',
                body,
                headers: {
                    'Content-Type': 'application/json; charset=utf-8',
                    'X-Requested-With': 'XMLHttpRequest',
                    'Accept': 'application/json, text/javascript, */*; q=0.01',
                    'Referer': `${this.config.baseUrl}/PurchaseOrder/Edit/${orderId}`,
                },
            });
            const retryText = await retryResp.text();
            console.log('[nextgen] updateQuantity retry status:', retryResp.status);
            if (!retryResp.ok) {
                return {
                    success: false,
                    status: retryResp.status,
                    endpoint: url,
                    requestBody: [{ Field: 'body', Value: body.slice(0, 500) }],
                    responseBody: retryText,
                    error: `UpdateQuantity failed: HTTP ${retryResp.status}`,
                };
            }
            return this.parseUpdateQuantityResponse(retryText, retryResp.status, url, body);
        }

        if (!response.ok) {
            return {
                success: false,
                status: response.status,
                endpoint: url,
                requestBody: [{ Field: 'body', Value: body.slice(0, 500) }],
                responseBody: text,
                error: `UpdateQuantity failed: HTTP ${response.status}`,
            };
        }

        return this.parseUpdateQuantityResponse(text, response.status, url, body);
    }

    private parseUpdateQuantityResponse(
        text: string, status: number, url: string, body: string
    ): NextGenInsertResult {
        // Parse the response - same format as insertLine
        let success = true;
        let errorMsg: string | undefined;
        try {
            const parsed = JSON.parse(text);
            success = parsed.Success !== false;
            if (parsed.GridRowErrorDescriptors?.length > 0) {
                const desc = parsed.GridRowErrorDescriptors[0];
                if (desc.RowErrorDescriptor?.Success === false) {
                    success = false;
                    errorMsg = desc.RowErrorDescriptor.Items?.map((i: any) => i.Description).join('; ');
                }
            }
        } catch {}

        return {
            success,
            status,
            endpoint: url,
            requestBody: [{ Field: 'body', Value: body.slice(0, 500) }],
            responseBody: text,
            error: errorMsg,
        };
    }

    // ─── Endpoint discovery (read-only, safe) ───────────────────────────

    /**
     * Discover NextGen insert/update endpoints by fetching UI HTML pages and
     * parsing JavaScript, form actions, and AJAX URLs.
     *
     * This is read-only (GET requests to HTML pages) and does not create any
     * data in NextGen.
     *
     * @param pages - NextGen page paths to fetch (e.g. "/PurchaseOrder/Create", "/PurchaseOrder/Edit/123")
     */
    async discoverEndpoints(pages: string[] = [
        '/PurchaseOrder/Create',
        '/PurchaseOrder',
        '/PurchaseOrder/Index',
    ]): Promise<{
        pages: Array<{ path: string; status: number; html: string; endpoints: NextGenEndpointHint[] }>;
        allEndpoints: NextGenEndpointHint[];
    }> {
        await this.login();

        const results: Array<{ path: string; status: number; html: string; endpoints: NextGenEndpointHint[] }> = [];
        const allEndpointsMap = new Map<string, NextGenEndpointHint>();

        for (const pagePath of pages) {
            const url = `${this.config.baseUrl}${pagePath}`;
            console.log('[nextgen] discoverEndpoints fetching:', url);

            let response: Response;
            try {
                response = await this.fetchWithCookie(url, { method: 'GET' }, true);
            } catch (err) {
                console.warn('[nextgen] discoverEndpoints failed for', pagePath, err);
                continue;
            }

            const html = await response.text();
            const endpoints = this.parseEndpointsFromHtml(html, pagePath);

            results.push({ path: pagePath, status: response.status, html: html.slice(0, 50000), endpoints });

            for (const ep of endpoints) {
                const key = `${ep.method || 'POST'} ${ep.url}`;
                if (!allEndpointsMap.has(key)) {
                    allEndpointsMap.set(key, ep);
                }
            }
        }

        return {
            pages: results,
            allEndpoints: Array.from(allEndpointsMap.values()),
        };
    }

    /**
     * Parse HTML for AJAX endpoint URLs, form actions, and JavaScript references.
     * Looks for patterns like:
     *   - url: "/PurchaseOrderLine/InsertFormData"
     *   - action="/PurchaseOrder/Create"
     *   - $.post("/PurchaseOrder/InsertFormData", ...)
     *   - fetch("/PurchaseOrderLine/Read", ...)
     *   - data-url="/PurchaseOrderSize/InsertFormData"
     */
    private parseEndpointsFromHtml(html: string, pagePath: string): NextGenEndpointHint[] {
        const endpoints: NextGenEndpointHint[] = [];
        const seen = new Set<string>();

        // Match URLs in JavaScript strings and HTML attributes that look like controller/action paths
        // Pattern: /Word/Word or /Word/Word/Word (PascalCase, typical ASP.NET MVC routing)
        const urlPattern = /["'`](\/[A-Z][a-zA-Z]+\/[A-Z][a-zA-Z]+(?:\/[a-zA-Z0-9_]+)?)["'`]/g;
        let match: RegExpExecArray | null;
        while ((match = urlPattern.exec(html)) !== null) {
            const url = match[1];
            const key = url;
            if (seen.has(key)) continue;
            seen.add(key);

            // Determine context from surrounding text
            const start = Math.max(0, match.index - 80);
            const end = Math.min(html.length, match.index + match[0].length + 80);
            const context = html.slice(start, end).replace(/\s+/g, ' ').trim();

            // Try to determine HTTP method from context
            let method: string | undefined;
            const lowerContext = context.toLowerCase();
            if (lowerContext.includes('$.post') || lowerContext.includes('fetch(') || lowerContext.includes('ajax') && lowerContext.includes('post')) {
                method = 'POST';
            } else if (lowerContext.includes('$.get') || lowerContext.includes('fetch(') && lowerContext.includes('get')) {
                method = 'GET';
            } else if (context.includes('action=')) {
                method = 'POST';
            }

            endpoints.push({ url, context, method });
        }

        // Also look for InsertFormData / UpdateFormData / DeleteFormData patterns specifically
        const insertPattern = /["'`](\/[A-Z][a-zA-Z]+\/Insert[A-Za-z]*)["'`]/g;
        while ((match = insertPattern.exec(html)) !== null) {
            const url = match[1];
            const key = `POST ${url}`;
            if (seen.has(key)) continue;
            seen.add(key);
            const start = Math.max(0, match.index - 80);
            const end = Math.min(html.length, match.index + match[0].length + 80);
            const context = html.slice(start, end).replace(/\s+/g, ' ').trim();
            endpoints.push({ url, context, method: 'POST' });
        }

        return endpoints;
    }
}

// ─── Write-safe client factory ───────────────────────────────────────

/**
 * Create a NextGenClient configured for the TEST/write environment.
 *
 * This client reads NEXTGEN_WRITE_BASE_URL, NEXTGEN_WRITE_USERNAME, and
 * NEXTGEN_WRITE_PASSWORD from environment variables. The base URL MUST
 * contain ":8443" (the test port) — all write methods enforce this at
 * runtime via assertTestEnvironment().
 *
 * If NEXTGEN_WRITE_ENABLED is "false", or if the write URL is missing or
 * points to production (no :8443), this function throws.
 *
 * SAFETY: This is the ONLY sanctioned way to create a client for write
 * operations. Never call new NextGenClient() with production credentials
 * for writes.
 */
export function createWriteClient(): NextGenClient {
    const enabled = process.env.NEXTGEN_WRITE_ENABLED !== 'false';
    if (!enabled) {
        throw new Error('NextGen write operations are disabled (NEXTGEN_WRITE_ENABLED=false).');
    }

    const baseUrl = process.env.NEXTGEN_WRITE_BASE_URL || '';
    const username = process.env.NEXTGEN_WRITE_USERNAME || '';
    const password = process.env.NEXTGEN_WRITE_PASSWORD || '';

    if (!baseUrl || !username || !password) {
        throw new Error(
            'NextGen write environment not configured. Set NEXTGEN_WRITE_BASE_URL, ' +
            'NEXTGEN_WRITE_USERNAME, and NEXTGEN_WRITE_PASSWORD in .env.local.'
        );
    }

    if (!baseUrl.includes(':8443')) {
        throw new Error(
            `SAFETY BLOCK: NEXTGEN_WRITE_BASE_URL must be the test environment (port 8443).\n` +
            `Got: ${baseUrl}\n` +
            `Production (https://nextgen.madison88.com) must NEVER be used for writes.`
        );
    }

    return new NextGenClient({ baseUrl, username, password });
}

/**
 * Check if write operations are available (test env configured + enabled).
 */
export function isWriteEnabled(): boolean {
    if (process.env.NEXTGEN_WRITE_ENABLED === 'false') return false;
    const baseUrl = process.env.NEXTGEN_WRITE_BASE_URL || '';
    return baseUrl.includes(':8443') &&
           Boolean(process.env.NEXTGEN_WRITE_USERNAME) &&
           Boolean(process.env.NEXTGEN_WRITE_PASSWORD);
}
