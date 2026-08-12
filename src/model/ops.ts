import { IGraphNode, IGraphEdge } from './bundle';

/**
 * The unit of per-item storage: one semantic change to the graph.
 *
 * Layout is deliberately NOT an op. Dragging a node or running a force layout
 * changes presentation, not meaning, and a layout run touches every node at once —
 * as ops that would be one write per node per run, which is exactly the write storm
 * that makes naive per-item storage unusable. Positions travel as a single blob on
 * the project item instead, so the highest-frequency mutation in the app costs one
 * write regardless of graph size.
 */
export type GraphOp =
  | { kind: 'upsertNode'; node: IGraphNode }
  | { kind: 'deleteNode'; id: string }
  | { kind: 'upsertEdge'; edge: IGraphEdge }
  | { kind: 'deleteEdge'; id: string };

/**
 * Identity for coalescing: two ops with the same key act on the same entity, so
 * only the last one matters. "60 keystrokes renaming a node" collapses to one write.
 */
export const opKey = (op: GraphOp): string => {
  switch (op.kind) {
    case 'upsertNode': return 'n:' + op.node.id;
    case 'deleteNode': return 'n:' + op.id;
    case 'upsertEdge': return 'e:' + op.edge.id;
    case 'deleteEdge': return 'e:' + op.id;
    default: return '';
  }
};

export const isEdgeOp = (op: GraphOp): boolean => op.kind === 'upsertEdge' || op.kind === 'deleteEdge';
export const isDeleteOp = (op: GraphOp): boolean => op.kind === 'deleteNode' || op.kind === 'deleteEdge';

/**
 * Collapse a stream of ops to one per entity, preserving first-seen order so a
 * flush stays readable in the network log.
 *
 * A delete always wins over an earlier upsert for the same entity: creating then
 * deleting a node within one debounce window must not leave a create behind.
 */
export const coalesceOps = (ops: GraphOp[]): GraphOp[] => {
  const byKey = new Map<string, GraphOp>();
  const order: string[] = [];
  for (const op of ops) {
    const key = opKey(op);
    if (!byKey.has(key)) { order.push(key); }
    byKey.set(key, op);
  }
  return order.map((k) => byKey.get(k) as GraphOp);
};

/**
 * Execution order for a flush: edge deletes, node deletes, node upserts, edge upserts.
 *
 * SharePoint has no transactions, so ordering is the only referential-integrity tool
 * available. Deleting an edge before its endpoint means a failure mid-flush can leave
 * a missing edge (harmless, re-derivable) rather than an edge pointing at a node that
 * no longer exists. Likewise a node must exist before an edge referencing it is written.
 */
export const orderOps = (ops: GraphOp[]): GraphOp[] => {
  const rank = (op: GraphOp): number => {
    if (op.kind === 'deleteEdge') { return 0; }
    if (op.kind === 'deleteNode') { return 1; }
    if (op.kind === 'upsertNode') { return 2; }
    return 3;
  };
  return ops
    .map((op, i) => ({ op, i }))
    .sort((a, b) => (rank(a.op) - rank(b.op)) || (a.i - b.i))
    .map((x) => x.op);
};
