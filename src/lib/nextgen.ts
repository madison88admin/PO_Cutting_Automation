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

const PO_NUMBER_FIELD = 'PrimaryUserDefinedFieldValuesTextUdf3';

export class NextGenClient {
    private config: NextGenConfig;
    private auth: NextGenAuth | null = null;
    private loginPromise: Promise<NextGenAuth> | null = null;

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
        if (this.loginPromise) return this.loginPromise;

        this.loginPromise = this.performLogin(retries);
        try {
            return await this.loginPromise;
        } finally {
            this.loginPromise = null;
        }
    }

    private async performLogin(retries: number): Promise<NextGenAuth> {
        let lastError: Error | null = null;
        for (let attempt = 1; attempt <= retries; attempt++) {
            try {
                const loginUrl = `${this.config.baseUrl}/Account/Login`;
                let loginPage = await this.fetchWithCookie(loginUrl, { method: 'GET' }, false);

                if (loginPage.status === 302) {
                    const location = loginPage.headers.get('location') || '';
                    console.log('[nextgen] login page returned 302, location:', location);
                    if (location && !location.includes('/Account/Login')) {
                        loginPage = await this.fetchWithCookie(
                            new URL(location, this.config.baseUrl).toString(),
                            { method: 'GET' },
                            false,
                        );
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
                console.log('[nextgen] login POST session cookie received:', Boolean(loginResponse.headers.get('set-cookie')));
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
            season: String(line.Season || line.season || ''),
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
}
