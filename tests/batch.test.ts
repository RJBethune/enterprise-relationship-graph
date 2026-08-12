import { suite, test, assert } from './harness';
import { buildBatchBody, parseBatchResponse, IBatchRequest } from '../src/services/sp/batch';

const REQ: IBatchRequest[] = [
  { method: 'POST', url: 'https://site/_api/web/lists/getbytitle(\'ERG Nodes\')/items', body: { ErgNodeId: 'a' } },
  { method: 'MERGE', url: 'https://site/_api/web/lists/getbytitle(\'ERG Nodes\')/items(7)', body: { ErgNodeId: 'b' }, etag: '*' },
  { method: 'DELETE', url: 'https://site/_api/web/lists/getbytitle(\'ERG Edges\')/items(3)', etag: '*' }
];

suite('batch: envelope construction', () => {
  test('writes are wrapped in a changeset and the batch is closed', () => {
    const body = buildBatchBody(REQ, 'B1', 'C1');
    assert.includes(body, '--batch_B1');
    assert.includes(body, 'Content-Type: multipart/mixed; boundary=changeset_C1');
    assert.includes(body, '--changeset_C1--');
    assert.includes(body, '--batch_B1--');
  });

  test('each write carries its method, URL and payload', () => {
    const body = buildBatchBody(REQ, 'B1', 'C1');
    assert.includes(body, 'POST https://site/_api/web/lists/getbytitle(\'ERG Nodes\')/items HTTP/1.1');
    assert.includes(body, 'MERGE https://site/_api/web/lists/getbytitle(\'ERG Nodes\')/items(7) HTTP/1.1');
    assert.includes(body, 'DELETE https://site/_api/web/lists/getbytitle(\'ERG Edges\')/items(3) HTTP/1.1');
    assert.includes(body, '{"ErgNodeId":"a"}');
  });

  test('MERGE and DELETE carry IF-MATCH; POST does not', () => {
    const body = buildBatchBody(REQ, 'B1', 'C1');
    const ifMatches = body.split('IF-MATCH:').length - 1;
    assert.equal(ifMatches, 2, 'only the two updating verbs need a concurrency guard');
  });

  test('reads sit OUTSIDE the changeset — SharePoint rejects a GET inside one', () => {
    const body = buildBatchBody(
      [{ method: 'GET', url: 'https://site/_api/web/lists' }, REQ[0]],
      'B2', 'C2'
    );
    const getAt = body.indexOf('GET https://site/_api/web/lists HTTP/1.1');
    const changesetAt = body.indexOf('boundary=changeset_C2');
    assert.ok(getAt < changesetAt, 'the GET must appear before the changeset opens');
  });

  test('a read-only batch opens no changeset at all', () => {
    const body = buildBatchBody([{ method: 'GET', url: 'https://site/_api/web' }], 'B3', 'C3');
    assert.ok(body.indexOf('changeset_C3') < 0);
  });

  test('CRLF line endings, as the multipart spec requires', () => {
    assert.includes(buildBatchBody(REQ, 'B1', 'C1'), '\r\n');
  });
});

suite('batch: response parsing', () => {
  const RESPONSE = [
    '--batchresponse_x',
    'Content-Type: multipart/mixed; boundary=changesetresponse_y',
    '',
    '--changesetresponse_y',
    'Content-Type: application/http',
    '',
    'HTTP/1.1 201 Created',
    'CONTENT-TYPE: application/json;odata=nometadata',
    'ETAG: "3"',
    '',
    '{"Id":42,"ErgNodeId":"a"}',
    '--changesetresponse_y',
    'Content-Type: application/http',
    '',
    'HTTP/1.1 204 No Content',
    '',
    '',
    '--changesetresponse_y',
    'Content-Type: application/http',
    '',
    'HTTP/1.1 404 Not Found',
    '',
    '{"error":{"message":{"value":"Item does not exist."}}}',
    '--changesetresponse_y--',
    '--batchresponse_x--',
    ''
  ].join('\r\n');

  test('one result per operation, in order', () => {
    const results = parseBatchResponse(RESPONSE);
    assert.equal(results.length, 3);
    assert.deepEqual(results.map((r) => r.status), [201, 204, 404]);
  });

  test('a created item yields its new Id — without it the next edit would duplicate the row', () => {
    const results = parseBatchResponse(RESPONSE);
    assert.ok(results[0].ok);
    assert.equal((results[0].json as { Id: number }).Id, 42);
    assert.equal(results[0].etag, '"3"');
  });

  test('a failed operation inside a successful batch is reported as failed', () => {
    // SharePoint does not roll a changeset back, so per-part status is the only
    // honest signal — a 200 on the batch itself says nothing about the writes.
    const results = parseBatchResponse(RESPONSE);
    assert.ok(!results[2].ok);
    assert.includes(results[2].body, 'Item does not exist.');
  });

  test('an empty response parses to no results rather than throwing', () => {
    assert.equal(parseBatchResponse('').length, 0);
  });
});
