import { IGraph, IGraphNode, IGraphEdge, sameEntity } from './bundle';
import { GraphOp, orderOps } from './ops';

/**
 * Turn "the graph used to look like this, now it looks like that" into the minimal
 * set of per-item writes.
 *
 * This is the bridge between the engine — which only ever hands over a whole graph —
 * and per-item storage, which needs to know precisely what changed. Keeping the diff
 * here rather than instrumenting the engine is what let the engine stay verbatim.
 */

const byId = <T extends { id: string }>(items: T[]): Map<string, T> => {
  const map = new Map<string, T>();
  for (const item of items || []) {
    if (item && typeof item.id === 'string') { map.set(item.id, item); }
  }
  return map;
};

export const diffGraphs = (base: IGraph, next: IGraph): GraphOp[] => {
  const ops: GraphOp[] = [];

  const baseNodes = byId<IGraphNode>(base.nodes);
  const nextNodes = byId<IGraphNode>(next.nodes);
  nextNodes.forEach((node, id) => {
    const before = baseNodes.get(id);
    if (!before || !sameEntity(before, node)) { ops.push({ kind: 'upsertNode', node: node }); }
  });
  baseNodes.forEach((_node, id) => {
    if (!nextNodes.has(id)) { ops.push({ kind: 'deleteNode', id: id }); }
  });

  const baseEdges = byId<IGraphEdge>(base.edges);
  const nextEdges = byId<IGraphEdge>(next.edges);
  nextEdges.forEach((edge, id) => {
    const before = baseEdges.get(id);
    if (!before || !sameEntity(before, edge)) { ops.push({ kind: 'upsertEdge', edge: edge }); }
  });
  baseEdges.forEach((_edge, id) => {
    if (!nextEdges.has(id)) { ops.push({ kind: 'deleteEdge', id: id }); }
  });

  return orderOps(ops);
};

/**
 * True when nothing but presentation changed. Lets the autosave path skip a write
 * entirely for a pure pan/zoom, and lets per-item mode send only the layout blob
 * after a drag.
 */
export const hasSemanticChange = (base: IGraph, next: IGraph): boolean =>
  diffGraphs(base, next).length > 0;

/** Positions differ — worth a layout-blob write, not worth an item write. */
export const hasLayoutChange = (base: IGraph, next: IGraph): boolean =>
  !sameEntity(base.positions || {}, next.positions || {}) ||
  !sameEntity(base.collapsedNodes || [], next.collapsedNodes || []);

/**
 * Apply ops to a graph. Used to replay remote changes into the local copy during
 * sync, and by the test harness to prove diff/apply round-trips.
 */
export const applyOps = (graph: IGraph, ops: GraphOp[]): IGraph => {
  const nodes = byId<IGraphNode>(graph.nodes);
  const edges = byId<IGraphEdge>(graph.edges);

  for (const op of ops) {
    switch (op.kind) {
      case 'upsertNode': nodes.set(op.node.id, op.node); break;
      case 'deleteNode':
        nodes.delete(op.id);
        // Referential integrity is this layer's job — SharePoint will not enforce it.
        edges.forEach((edge, id) => {
          if (edge.source === op.id || edge.target === op.id) { edges.delete(id); }
        });
        break;
      case 'upsertEdge': edges.set(op.edge.id, op.edge); break;
      case 'deleteEdge': edges.delete(op.id); break;
      default: break;
    }
  }

  return {
    ...graph,
    nodes: Array.from(nodes.values()),
    edges: Array.from(edges.values()).filter((e) => nodes.has(e.source) && nodes.has(e.target))
  };
};

/**
 * Drop edges with missing endpoints and report how many went.
 *
 * Per-item storage can produce these legitimately: a delete flush that fails partway,
 * or two people deleting a node and adding an edge to it at the same moment. Sweeping
 * on load is cheaper and more reliable than trying to make distributed writes atomic.
 */
export const repairOrphans = (graph: IGraph): { graph: IGraph; removedEdges: number } => {
  const ids = new Set((graph.nodes || []).map((n) => n.id));
  const kept = (graph.edges || []).filter((e) => ids.has(e.source) && ids.has(e.target));
  return { graph: { ...graph, edges: kept }, removedEdges: (graph.edges || []).length - kept.length };
};
