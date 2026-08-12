import { suite, test, assert } from './harness';
import { IGraph, normalizeGraph } from '../src/model/bundle';
import { diffGraphs, applyOps, repairOrphans, hasLayoutChange } from '../src/model/diff';
import { coalesceOps, orderOps, GraphOp } from '../src/model/ops';

const graph = (nodes: string[], edges: [string, string, string][] = []): IGraph => normalizeGraph({
  nodes: nodes.map((id) => ({ id, label: id.toUpperCase(), type: 'Office' })),
  edges: edges.map(([id, s, t]) => ({ id, source: s, target: t, type: 'CONTAINS' }))
} as unknown as IGraph);

suite('diff: whole-graph to per-item writes', () => {
  test('an added node produces one upsert', () => {
    const ops = diffGraphs(graph(['a']), graph(['a', 'b']));
    assert.equal(ops.length, 1);
    assert.equal(ops[0].kind, 'upsertNode');
  });

  test('an unchanged graph produces nothing', () => {
    assert.equal(diffGraphs(graph(['a', 'b']), graph(['a', 'b'])).length, 0);
  });

  test('a renamed node produces exactly one upsert', () => {
    const before = graph(['a', 'b']);
    const after = JSON.parse(JSON.stringify(before)) as IGraph;
    after.nodes[0].label = 'Renamed';
    const ops = diffGraphs(before, after);
    assert.equal(ops.length, 1);
    assert.equal((ops[0] as { node: { id: string } }).node.id, 'a');
  });

  test('deleting a node with an edge orders the edge delete FIRST', () => {
    // Ordering is the only referential-integrity tool available: SharePoint has no
    // transactions, so a flush that dies halfway must leave a missing edge (harmless,
    // re-derivable) rather than an edge pointing at a node that no longer exists.
    const ops = diffGraphs(graph(['a', 'b'], [['e1', 'a', 'b']]), graph(['a']));
    assert.equal(ops[0].kind, 'deleteEdge');
    assert.equal(ops[1].kind, 'deleteNode');
  });

  test('a new node is written BEFORE an edge that references it', () => {
    const ops = diffGraphs(graph(['a']), graph(['a', 'b'], [['e1', 'a', 'b']]));
    const nodeAt = ops.findIndex((o) => o.kind === 'upsertNode');
    const edgeAt = ops.findIndex((o) => o.kind === 'upsertEdge');
    assert.ok(nodeAt < edgeAt, 'the node must exist before the edge referencing it');
  });

  test('layout-only changes are not semantic changes', () => {
    const before = graph(['a', 'b']);
    const after = JSON.parse(JSON.stringify(before)) as IGraph;
    after.positions = { a: { x: 10, y: 20 } };
    assert.equal(diffGraphs(before, after).length, 0, 'moving a node must not write items');
    assert.ok(hasLayoutChange(before, after), 'but it must still be recognised as a layout change');
  });
});

suite('diff: applying ops back', () => {
  test('diff then apply reproduces the target graph', () => {
    const before = graph(['a', 'b'], [['e1', 'a', 'b']]);
    const after = graph(['a', 'c'], [['e2', 'a', 'c']]);
    const result = applyOps(before, diffGraphs(before, after));
    assert.deepEqual(result.nodes.map((n) => n.id).sort(), ['a', 'c']);
    assert.deepEqual(result.edges.map((e) => e.id), ['e2']);
  });

  test('deleting a node also drops edges that referenced it', () => {
    const start = graph(['a', 'b'], [['e1', 'a', 'b']]);
    const result = applyOps(start, [{ kind: 'deleteNode', id: 'b' }]);
    assert.equal(result.edges.length, 0);
  });
});

suite('diff: orphan repair', () => {
  test('edges with a missing endpoint are swept and counted', () => {
    const broken = normalizeGraph({
      nodes: [{ id: 'a', label: 'A', type: 'Office' }],
      edges: [
        { id: 'e1', source: 'a', target: 'ghost', type: 'CONTAINS' },
        { id: 'e2', source: 'a', target: 'a', type: 'OWNS' }
      ]
    } as unknown as IGraph);
    const repaired = repairOrphans(broken);
    assert.equal(repaired.removedEdges, 1);
    assert.equal(repaired.graph.edges.length, 1);
  });
});

suite('ops: coalescing and ordering', () => {
  test('repeated edits to one node collapse to the last write', () => {
    const ops: GraphOp[] = [
      { kind: 'upsertNode', node: { id: 'a', label: 'O', type: 'Office' } },
      { kind: 'upsertNode', node: { id: 'a', label: 'Op', type: 'Office' } },
      { kind: 'upsertNode', node: { id: 'a', label: 'Ops', type: 'Office' } }
    ];
    const out = coalesceOps(ops);
    assert.equal(out.length, 1);
    assert.equal((out[0] as { node: { label: string } }).node.label, 'Ops');
  });

  test('create-then-delete in one window leaves only the delete', () => {
    const out = coalesceOps([
      { kind: 'upsertNode', node: { id: 'a', label: 'A', type: 'Office' } },
      { kind: 'deleteNode', id: 'a' }
    ]);
    assert.equal(out.length, 1);
    assert.equal(out[0].kind, 'deleteNode');
  });

  test('ops for different entities all survive', () => {
    const out = coalesceOps([
      { kind: 'upsertNode', node: { id: 'a', label: 'A', type: 'Office' } },
      { kind: 'upsertNode', node: { id: 'b', label: 'B', type: 'Office' } }
    ]);
    assert.equal(out.length, 2);
  });

  test('ordering is stable within a rank', () => {
    const out = orderOps([
      { kind: 'upsertEdge', edge: { id: 'e1', source: 'a', target: 'b', type: 'X' } },
      { kind: 'deleteNode', id: 'z' },
      { kind: 'upsertNode', node: { id: 'a', label: 'A', type: 'Office' } },
      { kind: 'deleteEdge', id: 'e9' }
    ]);
    assert.deepEqual(out.map((o) => o.kind), ['deleteEdge', 'deleteNode', 'upsertNode', 'upsertEdge']);
  });
});
