/**
 * OData $batch envelope construction and parsing.
 *
 * Batching does not reduce SharePoint's write METERING — 50 writes in one batch still
 * cost 50 operations — but it collapses 50 round trips into one, which is the
 * difference between a layout import feeling like a progress bar and feeling like a
 * hang. The queue above is what protects against throttling; this is what protects
 * against latency.
 *
 * A note on changesets: OData says a changeset is atomic. SharePoint does NOT roll one
 * back on partial failure — it reports per-operation status and leaves successful
 * writes in place. Callers must treat every response independently, which is why
 * `parseBatchResponse` returns a status per part rather than one verdict.
 *
 * Both functions here are pure strings-in/strings-out so the wire format is unit
 * testable without a tenant.
 */

export type BatchMethod = 'GET' | 'POST' | 'MERGE' | 'DELETE';

export interface IBatchRequest {
  method: BatchMethod;
  /** Absolute URL, e.g. https://site/_api/web/lists/getbytitle('X')/items(4) */
  url: string;
  body?: unknown;
  /** Concurrency guard for MERGE/DELETE. '*' means "last writer wins, deliberately". */
  etag?: string;
  headers?: { [name: string]: string };
}

export interface IBatchPartResult {
  status: number;
  ok: boolean;
  body: string;
  json: unknown;
  etag: string | null;
}

const JSON_CT = 'application/json;odata=nometadata';

export const buildBatchBody = (
  requests: IBatchRequest[],
  batchId: string,
  changesetId: string
): string => {
  const reads = requests.filter((r) => r.method === 'GET');
  const writes = requests.filter((r) => r.method !== 'GET');
  const lines: string[] = [];

  // Reads sit directly in the batch — they must not be inside a changeset.
  for (const r of reads) {
    lines.push(`--batch_${batchId}`);
    lines.push('Content-Type: application/http');
    lines.push('Content-Transfer-Encoding: binary');
    lines.push('');
    lines.push(`GET ${r.url} HTTP/1.1`);
    lines.push(`Accept: ${JSON_CT}`);
    lines.push('');
  }

  if (writes.length > 0) {
    lines.push(`--batch_${batchId}`);
    lines.push(`Content-Type: multipart/mixed; boundary=changeset_${changesetId}`);
    lines.push('');
    for (const r of writes) {
      lines.push(`--changeset_${changesetId}`);
      lines.push('Content-Type: application/http');
      lines.push('Content-Transfer-Encoding: binary');
      lines.push('');
      lines.push(`${r.method} ${r.url} HTTP/1.1`);
      lines.push(`Content-Type: ${JSON_CT}`);
      lines.push(`Accept: ${JSON_CT}`);
      if (r.method === 'MERGE' || r.method === 'DELETE') {
        lines.push(`IF-MATCH: ${r.etag || '*'}`);
      }
      if (r.headers) {
        for (const name of Object.keys(r.headers)) { lines.push(`${name}: ${r.headers[name]}`); }
      }
      lines.push('');
      lines.push(r.body === undefined ? '' : JSON.stringify(r.body));
    }
    lines.push(`--changeset_${changesetId}--`);
  }

  lines.push(`--batch_${batchId}--`);
  lines.push('');
  return lines.join('\r\n');
};

/**
 * Split a multipart batch response into one result per request, in order.
 *
 * The response nests changeset parts inside batch parts, but every part that matters
 * carries an `HTTP/1.1 <status>` status line, so scanning for those recovers the
 * sequence without a full MIME parser — and without a MIME dependency in the bundle.
 */
export const parseBatchResponse = (text: string): IBatchPartResult[] => {
  const results: IBatchPartResult[] = [];
  const normalized = text.replace(/\r\n/g, '\n');
  const statusRe = /^HTTP\/1\.1 (\d{3})/gm;

  const starts: { index: number; status: number }[] = [];
  let match = statusRe.exec(normalized);
  while (match !== null) {
    starts.push({ index: match.index, status: Number(match[1]) });
    match = statusRe.exec(normalized);
  }

  for (let i = 0; i < starts.length; i++) {
    const from = starts[i].index;
    const to = i + 1 < starts.length ? starts[i + 1].index : normalized.length;
    const part = normalized.slice(from, to);

    // Headers run to the first blank line; the body runs from there to the next
    // boundary marker. Truncating AT the boundary matters: slicing to the next status
    // line also swallows the following part's MIME headers, and appending
    // "Content-Type: application/http" to a JSON body makes it unparseable — which
    // silently loses the Id of every newly created item.
    const split = part.indexOf('\n\n');
    const headerBlock = split >= 0 ? part.slice(0, split) : part;
    let body = split >= 0 ? part.slice(split + 2) : '';
    const boundaryAt = body.search(/^--(batch|changeset)/m);
    if (boundaryAt >= 0) { body = body.slice(0, boundaryAt); }
    body = body.trim();

    const etagMatch = /^ETAG:\s*(.+)$/im.exec(headerBlock);
    let json: unknown = null;
    if (body && (body.charAt(0) === '{' || body.charAt(0) === '[')) {
      try { json = JSON.parse(body); } catch { json = null; }
    }

    const status = starts[i].status;
    results.push({
      status,
      ok: status >= 200 && status < 300,
      body,
      json,
      etag: etagMatch ? etagMatch[1].trim() : null
    });
  }

  return results;
};

/** Extract a human-usable message from a failed SharePoint REST response body. */
export const errorMessageFrom = (body: string, json: unknown): string => {
  const obj = json as { error?: { message?: { value?: string } | string } } | null;
  const msg = obj && obj.error ? obj.error.message : undefined;
  if (typeof msg === 'string') { return msg; }
  if (msg && typeof msg === 'object' && typeof msg.value === 'string') { return msg.value; }
  return (body || 'Unknown SharePoint error').slice(0, 300);
};

/**
 * Batch ids must be unique per request but need no cryptographic quality — they only
 * have to not appear in the payload. A counter plus a per-session salt is enough, and
 * avoids depending on crypto in a bundle that must run in older browsers.
 */
let batchCounter = 0;
const sessionSalt = Math.random().toString(36).slice(2, 10);
export const nextBatchId = (): string => `${sessionSalt}_${++batchCounter}`;
