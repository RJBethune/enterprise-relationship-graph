import { suite, test, assert } from './harness';
import {
  stableStringify, sameEntity, splitPayload, joinPayload, parseBundle, normalizeGraph,
  PayloadTooLargeError, CHUNK_SIZE, MAX_CHUNKS, IGraph
} from '../src/model/bundle';
import { payloadPatch, payloadColumns } from '../src/services/sp/projectItems';

suite('bundle: stable serialization', () => {
  test('key order does not change the serialized form', () => {
    const a = { id: 'n1', label: 'Ops', type: 'Office' };
    const b = { type: 'Office', label: 'Ops', id: 'n1' };
    assert.equal(stableStringify(a), stableStringify(b));
    assert.ok(sameEntity(a, b), 'same node written in a different key order must compare equal');
  });

  test('nested objects and arrays are sorted too', () => {
    const a = { id: 'n1', meta: { b: 1, a: 2 }, tags: ['x', 'y'] };
    const b = { tags: ['x', 'y'], id: 'n1', meta: { a: 2, b: 1 } };
    assert.equal(stableStringify(a), stableStringify(b));
  });

  test('array ORDER still matters — it is data, not key order', () => {
    assert.ok(!sameEntity({ tags: ['x', 'y'] }, { tags: ['y', 'x'] }));
  });

  test('a real difference is still detected', () => {
    assert.ok(!sameEntity({ id: 'n1', label: 'Ops' }, { id: 'n1', label: 'Operations' }));
  });

  test('undefined members are omitted rather than serialized', () => {
    assert.equal(stableStringify({ a: 1, b: undefined }), '{"a":1}');
  });
});

suite('bundle: payload chunking', () => {
  test('a short payload round-trips through one chunk', () => {
    const text = JSON.stringify({ hello: 'world' });
    const chunks = splitPayload(text);
    assert.equal(chunks.length, 1);
    assert.equal(joinPayload(chunks), text);
  });

  test('a payload spanning several columns round-trips exactly', () => {
    const text = 'x'.repeat(CHUNK_SIZE * 2 + 17);
    const chunks = splitPayload(text);
    assert.equal(chunks.length, 3);
    assert.equal(chunks[0].length, CHUNK_SIZE);
    assert.equal(joinPayload(chunks), text);
  });

  test('an empty payload still yields one chunk', () => {
    assert.deepEqual(splitPayload(''), ['']);
  });

  test('oversize payloads fail loudly instead of truncating', () => {
    const text = 'x'.repeat(CHUNK_SIZE * MAX_CHUNKS + 1);
    assert.throws(() => splitPayload(text));
    try { splitPayload(text); } catch (e) {
      assert.ok(e instanceof PayloadTooLargeError);
      assert.includes((e as Error).message, 'per-item storage');
    }
  });

  test('joinPayload ignores unused trailing columns', () => {
    assert.equal(joinPayload(['a', 'b', '', null, undefined]), 'ab');
  });
});

suite('bundle: payload patch blanks stale columns', () => {
  test('shrinking a graph clears the columns it no longer needs', () => {
    // The bug this guards: a graph that used 3 columns then shrinks to 1 leaves the
    // old chunks 2 and 3 in place, and the next read concatenates live JSON with stale
    // JSON into something unparseable — a corruption that only appears when a graph
    // gets SMALLER, which is exactly when nobody is watching for it.
    const big = payloadPatch('y'.repeat(CHUNK_SIZE * 3));
    assert.equal(big.ErgPayload3.length, CHUNK_SIZE);

    const small = payloadPatch('{"tiny":true}');
    assert.equal(small.ErgPayload1, '{"tiny":true}');
    assert.equal(small.ErgPayload2, '', 'column 2 must be explicitly blanked');
    assert.equal(small.ErgPayload3, '', 'column 3 must be explicitly blanked');
  });

  test('every provisioned column appears in the patch', () => {
    const patch = payloadPatch('{}');
    assert.equal(Object.keys(patch).length, MAX_CHUNKS);
    assert.deepEqual(Object.keys(patch), payloadColumns());
  });
});

suite('bundle: parsing every shape the tool has written', () => {
  test('current bundle shape', () => {
    const bundle = parseBundle(JSON.stringify({
      version: 4,
      graph: { nodes: [{ id: 'a', label: 'A', type: 'Office' }], edges: [] },
      snapshots: [{ name: 'before reorg' }],
      lastModifiedBy: 'Ross'
    }));
    assert.equal(bundle.version, 4);
    assert.equal(bundle.graph.nodes.length, 1);
    assert.equal((bundle.snapshots as unknown[]).length, 1);
    assert.equal(bundle.lastModifiedBy, 'Ross');
  });

  test('legacy graph-only file', () => {
    const bundle = parseBundle(JSON.stringify({ nodes: [{ id: 'a', label: 'A', type: 'X' }], edges: [] }));
    assert.equal(bundle.graph.nodes.length, 1);
  });

  test('a brand-new project with no payload yet opens empty, not broken', () => {
    const bundle = parseBundle('');
    assert.equal(bundle.graph.nodes.length, 0);
    assert.deepEqual(bundle.snapshots, []);
  });

  test('unrecognized JSON is rejected rather than silently emptied', () => {
    assert.throws(() => parseBundle('{"something":"else"}'));
  });

  test('normalizeGraph fills the optional collections downstream code assumes', () => {
    const graph = normalizeGraph({ nodes: [], edges: [] } as unknown as IGraph);
    assert.deepEqual(graph.customNodeTypes, []);
    assert.deepEqual(graph.collapsedNodes, []);
    assert.deepEqual(graph.positions, {});
  });
});
