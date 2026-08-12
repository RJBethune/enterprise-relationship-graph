import { suite, test, assert } from './harness';
import { IGraph, IGraphNode, normalizeGraph } from '../src/model/bundle';
import { threeWayMerge } from '../src/model/merge';

const g = (nodes: IGraphNode[], edges: IGraph['edges'] = []): IGraph =>
  normalizeGraph({ nodes, edges } as unknown as IGraph);

const node = (id: string, label: string): IGraphNode => ({ id, label, type: 'Office' });

suite('merge: the common case is not a conflict', () => {
  test('two people editing different nodes both keep their work', () => {
    const base = g([node('a', 'A'), node('b', 'B')]);
    const mine = g([node('a', 'Alpha'), node('b', 'B')]);
    const theirs = g([node('a', 'A'), node('b', 'Bravo')]);

    const result = threeWayMerge(base, mine, theirs);
    assert.equal(result.conflicts.length, 0, 'different nodes must not conflict');
    const byId = new Map(result.merged.nodes.map((n) => [n.id, n.label]));
    assert.equal(byId.get('a'), 'Alpha');
    assert.equal(byId.get('b'), 'Bravo');
  });

  test('a node they added arrives without a prompt', () => {
    const base = g([node('a', 'A')]);
    const mine = g([node('a', 'A')]);
    const theirs = g([node('a', 'A'), node('c', 'C')]);
    const result = threeWayMerge(base, mine, theirs);
    assert.equal(result.merged.nodes.length, 2);
    assert.equal(result.conflicts.length, 0);
    assert.ok(result.tookRemoteChanges);
  });

  test('identical edits on both sides are not a conflict', () => {
    const base = g([node('a', 'A')]);
    const same = g([node('a', 'Agreed')]);
    const result = threeWayMerge(base, same, g([node('a', 'Agreed')]));
    assert.equal(result.conflicts.length, 0);
    assert.equal(result.merged.nodes[0].label, 'Agreed');
  });

  test('a clean delete on their side is honoured', () => {
    const base = g([node('a', 'A'), node('b', 'B')]);
    const mine = g([node('a', 'A'), node('b', 'B')]);
    const theirs = g([node('a', 'A')]);
    const result = threeWayMerge(base, mine, theirs);
    assert.equal(result.merged.nodes.length, 1);
    assert.equal(result.conflicts.length, 0);
  });

  test('a clean delete on my side is honoured', () => {
    const base = g([node('a', 'A'), node('b', 'B')]);
    const result = threeWayMerge(base, g([node('a', 'A')]), g([node('a', 'A'), node('b', 'B')]));
    assert.equal(result.merged.nodes.length, 1);
    assert.equal(result.conflicts.length, 0);
  });
});

suite('merge: genuine collisions', () => {
  test('both editing the same node keeps mine and reports it', () => {
    const base = g([node('a', 'A')]);
    const result = threeWayMerge(base, g([node('a', 'Mine')]), g([node('a', 'Theirs')]));
    assert.equal(result.conflicts.length, 1);
    assert.equal(result.conflicts[0].reason, 'both-edited');
    assert.equal(result.merged.nodes[0].label, 'Mine');
  });

  test('my edit survives their delete', () => {
    // Resurrecting an edited node is recoverable — anyone can delete it again.
    // Silently discarding somebody's edit is not.
    const base = g([node('a', 'A'), node('b', 'B')]);
    const result = threeWayMerge(base, g([node('a', 'A'), node('b', 'Edited')]), g([node('a', 'A')]));
    assert.equal(result.conflicts.length, 1);
    assert.equal(result.conflicts[0].reason, 'edited-and-deleted');
    assert.equal(result.merged.nodes.length, 2);
  });

  test('their edit survives my delete', () => {
    const base = g([node('a', 'A'), node('b', 'B')]);
    const result = threeWayMerge(base, g([node('a', 'A')]), g([node('a', 'A'), node('b', 'Edited')]));
    assert.equal(result.conflicts.length, 1);
    assert.equal(result.merged.nodes.length, 2);
    assert.equal(result.merged.nodes.filter((n) => n.id === 'b')[0].label, 'Edited');
  });

  test('deleted on both sides stays deleted, silently', () => {
    const base = g([node('a', 'A'), node('b', 'B')]);
    const result = threeWayMerge(base, g([node('a', 'A')]), g([node('a', 'A')]));
    assert.equal(result.merged.nodes.length, 1);
    assert.equal(result.conflicts.length, 0);
  });
});

suite('merge: keeps the result coherent', () => {
  test('an edge left dangling by a merge is swept', () => {
    // The case the entity merge alone cannot catch: I ADD a relationship to a node
    // while they delete that node. My edge is new, so it survives the entity merge on
    // its own merits — and then points at nothing. Only the orphan sweep catches it.
    const base = g([node('a', 'A'), node('b', 'B')], []);
    const mine = g(
      [node('a', 'A'), node('b', 'B')],
      [{ id: 'e1', source: 'a', target: 'b', type: 'CONTAINS' }]
    );
    const theirs = g([node('a', 'A')], []);

    const result = threeWayMerge(base, mine, theirs);
    assert.equal(result.merged.nodes.length, 1, 'their clean delete stands');
    assert.equal(result.merged.edges.length, 0, 'my edge to the deleted node must not survive');
    assert.equal(result.removedOrphanEdges, 1);
  });

  test('a clean delete on both sides needs no sweep at all', () => {
    const base = g([node('a', 'A'), node('b', 'B')], [{ id: 'e1', source: 'a', target: 'b', type: 'CONTAINS' }]);
    const mine = g([node('a', 'A'), node('b', 'B')], [{ id: 'e1', source: 'a', target: 'b', type: 'CONTAINS' }]);
    const theirs = g([node('a', 'A')], []);
    const result = threeWayMerge(base, mine, theirs);
    assert.equal(result.merged.edges.length, 0);
    assert.equal(result.removedOrphanEdges, 0, 'the entity merge already dropped it');
  });

  test('my viewport and layout win — presentation is never a conflict', () => {
    const base = g([node('a', 'A')]);
    const mine = JSON.parse(JSON.stringify(base)) as IGraph;
    mine.positions = { a: { x: 5, y: 5 } };
    const theirs = JSON.parse(JSON.stringify(base)) as IGraph;
    theirs.positions = { a: { x: 900, y: 900 } };

    const result = threeWayMerge(base, mine, theirs);
    assert.deepEqual(result.merged.positions, { a: { x: 5, y: 5 } });
    assert.equal(result.conflicts.length, 0);
  });

  test('custom node types from both sides are unioned', () => {
    const base = g([node('a', 'A')]);
    const mine = JSON.parse(JSON.stringify(base)) as IGraph;
    mine.customNodeTypes = [{ name: 'Programme', color: '#f00' }];
    const theirs = JSON.parse(JSON.stringify(base)) as IGraph;
    theirs.customNodeTypes = [{ name: 'Initiative', color: '#0f0' }];

    const result = threeWayMerge(base, mine, theirs);
    assert.equal((result.merged.customNodeTypes as unknown[]).length, 2);
  });
});
