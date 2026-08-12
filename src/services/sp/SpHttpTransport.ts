import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { ISpTransport, ISpRequestOptions, ISpRawResponse } from './SpRest';
import { parseRetryAfter } from '../WriteQueue';

/**
 * The real transport: SPFx's SPHttpClient, which carries the signed-in user's context
 * and the request digest automatically.
 *
 * This is the only file in the data layer that imports SPFx. Everything above it works
 * against ISpTransport, so the stores, the provisioner and the sync loop all run under
 * the test harness against an in-memory fake.
 */
export class SpHttpTransport implements ISpTransport {
  public constructor(
    private readonly client: SPHttpClient,
    public readonly webUrl: string
  ) {}

  public async request(
    method: string,
    absoluteUrl: string,
    options?: ISpRequestOptions
  ): Promise<ISpRawResponse> {
    const opts = options || {};
    const headers: { [name: string]: string } = {
      Accept: opts.accept || 'application/json;odata=nometadata',
      // Suppresses the OData version negotiation that otherwise rejects nometadata.
      'odata-version': ''
    };
    if (opts.contentType) { headers['Content-Type'] = opts.contentType; }
    if (opts.etag) { headers['IF-MATCH'] = opts.etag; }
    if (opts.headers) {
      for (const name of Object.keys(opts.headers)) { headers[name] = opts.headers[name]; }
    }

    const config: ISPHttpClientOptions = { headers, body: opts.body };

    const res: SPHttpClientResponse = method === 'GET'
      ? await this.client.get(absoluteUrl, SPHttpClient.configurations.v1, config)
      : await this.client.post(absoluteUrl, SPHttpClient.configurations.v1, config);

    const text = await res.text();
    let json: unknown = null;
    if (text && (text.charAt(0) === '{' || text.charAt(0) === '[')) {
      try { json = JSON.parse(text); } catch (_e) { json = null; }
    }

    return {
      status: res.status,
      ok: res.ok,
      text,
      json,
      etag: res.headers.get('ETag'),
      retryAfterMs: parseRetryAfter(res.headers.get('Retry-After'), Date.now())
    };
  }
}
