import {
  IBatchRequest, IBatchPartResult, buildBatchBody, parseBatchResponse, nextBatchId, errorMessageFrom
} from './batch';
import { ThrottledError, parseRetryAfter } from '../WriteQueue';

/**
 * A thin REST layer over SharePoint, deliberately independent of SPFx types.
 *
 * Everything the app does with SharePoint is an HTTP request against `_api`. Keeping
 * the transport behind a two-method interface means the entire data layer — stores,
 * provisioning, sync — can be exercised against an in-memory fake, which is what makes
 * the test harness possible without a tenant.
 *
 * We use `odata=nometadata` throughout: it is the modern SPFx idiom, it keeps payloads
 * small, and it removes the need to look up ListItemEntityTypeFullName before every
 * write.
 */

export interface ISpRequestOptions {
  body?: string;
  etag?: string;
  headers?: { [name: string]: string };
  contentType?: string;
  accept?: string;
}

export interface ISpRawResponse {
  status: number;
  ok: boolean;
  text: string;
  json: unknown;
  etag: string | null;
  retryAfterMs: number;
}

export interface ISpTransport {
  readonly webUrl: string;
  request(method: string, absoluteUrl: string, options?: ISpRequestOptions): Promise<ISpRawResponse>;
}

export class SpRestError extends Error {
  public readonly status: number;
  public constructor(status: number, message: string) {
    super(message);
    this.name = 'SpRestError';
    this.status = status;
  }
}

/** Thrown when a write loses a race — the caller decides whether to merge or ask. */
export class ConflictError extends Error {
  public constructor(message?: string) {
    super(message || 'The item was changed by someone else since it was loaded.');
    this.name = 'ConflictError';
  }
}

const JSON_CT = 'application/json;odata=nometadata';

export class SpRest {
  public constructor(private readonly transport: ISpTransport) {}

  public get webUrl(): string { return this.transport.webUrl; }

  /** `web/lists/getbytitle('X')/items` -> absolute `_api` URL. */
  public api(path: string): string {
    const base = this.transport.webUrl.replace(/\/+$/, '');
    return `${base}/_api/${path.replace(/^\/+/, '')}`;
  }

  private raise(res: ISpRawResponse): never {
    if (res.status === 429 || res.status === 503) { throw new ThrottledError(res.retryAfterMs || 5000); }
    if (res.status === 412) { throw new ConflictError(errorMessageFrom(res.text, res.json)); }
    throw new SpRestError(res.status, errorMessageFrom(res.text, res.json));
  }

  public async get<T>(path: string): Promise<T> {
    const res = await this.transport.request('GET', this.api(path), { accept: JSON_CT });
    if (!res.ok) { this.raise(res); }
    return res.json as T;
  }

  public async post<T>(path: string, body?: unknown, headers?: { [k: string]: string }): Promise<T> {
    const res = await this.transport.request('POST', this.api(path), {
      body: body === undefined ? undefined : JSON.stringify(body),
      contentType: JSON_CT,
      accept: JSON_CT,
      headers
    });
    if (!res.ok) { this.raise(res); }
    return res.json as T;
  }

  /**
   * Update an item. Passing the etag from load is what turns "last writer wins" into
   * a detectable conflict; passing '*' is an explicit, deliberate overwrite.
   */
  public async merge(path: string, body: unknown, etag: string = '*'): Promise<void> {
    const res = await this.transport.request('POST', this.api(path), {
      body: JSON.stringify(body),
      contentType: JSON_CT,
      accept: JSON_CT,
      etag,
      headers: { 'X-HTTP-Method': 'MERGE' }
    });
    if (!res.ok) { this.raise(res); }
  }

  public async del(path: string, etag: string = '*'): Promise<void> {
    const res = await this.transport.request('POST', this.api(path), {
      accept: JSON_CT,
      etag,
      headers: { 'X-HTTP-Method': 'DELETE' }
    });
    if (!res.ok) { this.raise(res); }
  }

  /** Read an item and its etag together — the etag is what a later save needs. */
  public async getWithEtag<T>(path: string): Promise<{ data: T; etag: string | null }> {
    const res = await this.transport.request('GET', this.api(path), { accept: JSON_CT });
    if (!res.ok) { this.raise(res); }
    return { data: res.json as T, etag: res.etag };
  }

  /**
   * Follow @odata.nextLink until the result set is exhausted.
   *
   * A list view returns at most 5,000 rows per request. Reading a large graph is
   * therefore a handful of sequential GETs, not one — cheap, but only if nobody
   * forgets the paging, which is exactly the kind of thing that works in dev with
   * 200 rows and fails in production with 6,000.
   */
  public async getAll<T>(path: string): Promise<T[]> {
    const out: T[] = [];
    let url = this.api(path);
    while (url) {
      const res = await this.transport.request('GET', url, { accept: JSON_CT });
      if (!res.ok) { this.raise(res); }
      const page = res.json as { value?: T[]; 'odata.nextLink'?: string; '@odata.nextLink'?: string };
      if (Array.isArray(page.value)) { out.push(...page.value); }
      url = (page['odata.nextLink'] || page['@odata.nextLink'] || '') as string;
    }
    return out;
  }

  public async batch(requests: IBatchRequest[]): Promise<IBatchPartResult[]> {
    if (requests.length === 0) { return []; }
    const batchId = `batch_${nextBatchId()}`.replace('batch_', '');
    const changesetId = nextBatchId();
    const body = buildBatchBody(requests, batchId, changesetId);

    const res = await this.transport.request('POST', this.api('$batch'), {
      body,
      contentType: `multipart/mixed; boundary=batch_${batchId}`,
      accept: JSON_CT
    });
    if (!res.ok) { this.raise(res); }
    return parseBatchResponse(res.text);
  }

  /**
   * Send a batch and fail loudly if any part failed.
   *
   * SharePoint does not roll back a changeset, so a partial failure leaves real writes
   * behind. Surfacing it as one error with the failed parts named is the only honest
   * option — silently succeeding would let the caller believe the graph is saved.
   */
  public async batchOrThrow(requests: IBatchRequest[]): Promise<IBatchPartResult[]> {
    const results = await this.batch(requests);
    const failures = results.filter((r) => !r.ok);
    if (failures.length > 0) {
      const throttled = failures.filter((f) => f.status === 429 || f.status === 503)[0];
      if (throttled) { throw new ThrottledError(5000); }
      const first = failures[0];
      throw new SpRestError(
        first.status,
        `${failures.length} of ${results.length} writes failed. First error: ` +
        errorMessageFrom(first.body, first.json)
      );
    }
    return results;
  }
}

export { parseRetryAfter };
