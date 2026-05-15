import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';

const CACHE_KEY = 'nintex_tasks_token';
const CACHE_EXPIRY_KEY = 'nintex_tasks_token_expiry';
const CACHE_ITEM_ID_KEY = 'nintex_tasks_token_item_id';
const EXPIRY_BUFFER_MS = 5 * 60 * 1000; // 5 minutes buffer before JWT expiry

export class TokenService {
  private spHttpClient: SPHttpClient;
  private tokenListUrl: string;
  private titleFilter: string | undefined;
  private tokenColumnName: string;

  /**
   * @param spHttpClient SPFx SPHttpClient instance
   * @param tokenListUrl Full SP REST API URL or browser list URL for the token list
   * @param titleFilter Optional title to filter the token list by
   * @param tokenColumnName The internal name of the column containing the token (default: Token)
   */
  constructor(spHttpClient: SPHttpClient, tokenListUrl: string, titleFilter?: string, tokenColumnName: string = 'Token') {
    this.spHttpClient = spHttpClient;
    this.tokenListUrl = tokenListUrl.replace(/\/+$/, '');
    this.titleFilter = titleFilter;
    this.tokenColumnName = tokenColumnName || 'Token';
  }

  /**
   * Get the Nintex API token — checks localStorage cache first, refetches if expired or within 5 min of expiry
   */
  public async getToken(): Promise<string> {
    const cached = this._getCachedToken();
    if (cached) {
      return cached;
    }

    const { token, itemId } = await this._fetchTokenFromList();
    this._cacheToken(token, itemId);
    return token;
  }

  /**
   * Clear the cached token
   */
  public clearCache(): void {
    try {
      const cacheSuffix = this.titleFilter ? `_${this.titleFilter}` : '';
      localStorage.removeItem(`${CACHE_KEY}${cacheSuffix}`);
      localStorage.removeItem(`${CACHE_EXPIRY_KEY}${cacheSuffix}`);
      localStorage.removeItem(`${CACHE_ITEM_ID_KEY}${cacheSuffix}`);
    } catch {
      // localStorage may not be available
    }
  }

  /**
   * Parse the JWT token's `exp` claim without a library.
   * Returns the expiry timestamp in milliseconds, or undefined if parsing fails.
   */
  private _parseJwtExpiry(token: string): number | undefined {
    try {
      const parts = token.split('.');
      if (parts.length !== 3) return undefined;

      // Base64url decode the payload (2nd part)
      let payload = parts[1];
      // Pad with '=' to make length a multiple of 4
      payload = payload.replace(/-/g, '+').replace(/_/g, '/');
      while (payload.length % 4 !== 0) {
        payload += '=';
      }

      const decoded = atob(payload);
      const parsed = JSON.parse(decoded) as Record<string, unknown>;

      if (typeof parsed.exp === 'number') {
        // JWT exp is in seconds, convert to milliseconds
        return parsed.exp * 1000;
      }
    } catch {
      // Token is not a valid JWT or can't be parsed
    }
    return undefined;
  }

  private _getCachedToken(): string | undefined {
    try {
      const cacheSuffix = this.titleFilter ? `_${this.titleFilter}` : '';
      const token = localStorage.getItem(`${CACHE_KEY}${cacheSuffix}`);
      const expiryStr = localStorage.getItem(`${CACHE_EXPIRY_KEY}${cacheSuffix}`);

      if (token && expiryStr) {
        const expiry = parseInt(expiryStr, 10);
        if (Date.now() < expiry) {
          return token;
        }
        // Token expired or within buffer — clear cache
        this.clearCache();
      }
    } catch {
      // localStorage not available
    }
    return undefined;
  }

  /**
   * Cache the token with expiry derived from the JWT's exp claim minus 5 minutes buffer.
   * Also caches the SP list item ID for subsequent fetches.
   */
  private _cacheToken(token: string, itemId: number | undefined): void {
    try {
      const cacheSuffix = this.titleFilter ? `_${this.titleFilter}` : '';
      const jwtExpiry = this._parseJwtExpiry(token);
      let cacheExpiry: number;

      if (jwtExpiry) {
        // Cache until 5 minutes before JWT expiry
        cacheExpiry = jwtExpiry - EXPIRY_BUFFER_MS;
      } else {
        // Fallback: cache for 30 minutes if JWT can't be parsed
        cacheExpiry = Date.now() + 30 * 60 * 1000;
      }

      // Don't cache if already expired
      if (cacheExpiry <= Date.now()) {
        return;
      }

      localStorage.setItem(`${CACHE_KEY}${cacheSuffix}`, token);
      localStorage.setItem(`${CACHE_EXPIRY_KEY}${cacheSuffix}`, cacheExpiry.toString());

      if (itemId !== undefined) {
        localStorage.setItem(`${CACHE_ITEM_ID_KEY}${cacheSuffix}`, itemId.toString());
      }
    } catch {
      // localStorage not available
    }
  }

  /**
   * Convert a browser list URL to a REST API URL if needed.
   */
  private _buildApiUrl(): string {
    let url = this.tokenListUrl;

    if (url.indexOf('/_api/') >= 0) {
      if (url.indexOf('/items') < 0) {
        url = url.replace(/\/+$/, '') + '/items';
      }
      return url;
    }

    const listsMatch = /^(https?:\/\/[^/]+(?:\/sites\/[^/]+)?(?:\/[^/]+)*)\/Lists\/([^/?#]+)/i.exec(url);
    if (listsMatch) {
      const siteUrl = listsMatch[1];
      const listName = decodeURIComponent(listsMatch[2]);
      return `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
    }

    url = url.replace(/\/+$/, '');
    if (url.indexOf('/items') < 0) {
      url += '/items';
    }
    return url;
  }

  private async _fetchTokenFromList(): Promise<{ token: string; itemId: number | undefined }> {
    const baseApiUrl = this._buildApiUrl();
    const cacheSuffix = this.titleFilter ? `_${this.titleFilter}` : '';

    // Check if we have a cached item ID — fetch items with ID > cached ID (newer tokens)
    let cachedItemId: number | undefined;
    try {
      const storedId = localStorage.getItem(`${CACHE_ITEM_ID_KEY}${cacheSuffix}`);
      if (storedId) {
        cachedItemId = parseInt(storedId, 10);
      }
    } catch {
      // localStorage not available
    }

    const select = `Id,Title,${this.tokenColumnName}`;
    const filters: string[] = [];
    if (this.titleFilter) {
      filters.push(`Title eq '${this.titleFilter}'`);
    }

    let apiUrl: string;
    if (cachedItemId) {
      // Fetch items newer than the cached item ID
      const filterWithId = [...filters, `Id gt ${cachedItemId}`].join(' and ');
      apiUrl = `${baseApiUrl}?$select=${select}&$filter=${filterWithId}&$top=1&$orderby=Id desc`;
    } else {
      // First fetch — get the latest item
      const filterStr = filters.length > 0 ? `&$filter=${filters.join(' and ')}` : '';
      apiUrl = `${baseApiUrl}?$select=${select}${filterStr}&$top=1&$orderby=Id desc`;
    }

    let items = await this._fetchItems(apiUrl);

    // If no newer items found, fall back to fetching the latest item regardless of ID
    if (items.length === 0 && cachedItemId) {
      const filterStr = filters.length > 0 ? `&$filter=${filters.join(' and ')}` : '';
      const fallbackUrl = `${baseApiUrl}?$select=${select}${filterStr}&$top=1&$orderby=Id desc`;
      items = await this._fetchItems(fallbackUrl);
    }

    if (items.length === 0) {
      const filterMsg = this.titleFilter ? ` with Title '${this.titleFilter}'` : '';
      throw new Error(`No items found in the configured token list${filterMsg}. Please add a list item with the Nintex API token.`);
    }

    const tokenValue = (items[0][this.tokenColumnName] || items[0].Token || items[0].Title) as string;
    const itemId = items[0].Id as number | undefined;

    if (!tokenValue) {
      throw new Error('Token value is empty. Please ensure the list item has a value in the Token or Title column.');
    }

    return { token: tokenValue, itemId };
  }

  private async _fetchItems(apiUrl: string): Promise<Record<string, unknown>[]> {
    const requestHeaders: Headers = new Headers();
    requestHeaders.append('Accept', 'application/json;odata=nometadata');
    requestHeaders.append('odata-version', '');

    const httpClientOptions: ISPHttpClientOptions = {
      headers: requestHeaders
    };

    let response: SPHttpClientResponse;
    try {
      response = await this.spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1,
        httpClientOptions
      );
    } catch (fetchErr) {
      throw new Error(
        `Network error fetching token from SP list. URL: ${apiUrl}. ` +
        `Error: ${fetchErr instanceof Error ? fetchErr.message : String(fetchErr)}`
      );
    }

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to fetch token from SP list (HTTP ${response.status}): ${errorText}`);
    }

    const responseText = await response.text();

    if (responseText.indexOf('<!DOCTYPE') >= 0 || responseText.indexOf('<html') >= 0) {
      throw new Error(
        'Token list URL returned an HTML page instead of JSON. ' +
        'Make sure the URL includes /_api/web/lists/... and ends with /items. ' +
        "Example: https://tenant.sharepoint.com/sites/mysite/_api/web/lists/getbytitle('NintexTokens')/items"
      );
    }

    let data: Record<string, unknown>;
    try {
      data = JSON.parse(responseText);
    } catch {
      throw new Error(`Token list response is not valid JSON. Response starts with: ${responseText.substring(0, 200)}`);
    }

    return (data.value as Record<string, unknown>[]) || [];
  }
}
