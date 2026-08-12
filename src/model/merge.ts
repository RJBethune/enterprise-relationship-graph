import { IGraph, IGraphNode, IGraphEdge, sameEntity, normalizeGraph } from './bundle';
import { repairOrphans } from './diff';

/**
 * Three-way merge for document storage — "git pull before push".
 *
 * When a save is rejected because someone else saved first, the naive options are
 * both bad: overwrite them, or throw away the local work. Since the bundle is
 * structured data with stable ids, there is a third option — merge at entity level.
 * Two people editing different nodes is not a conflict, and that is the overwhelmingly
 * common case for an office graph. Only a genuine same-entity collision reaches a human.
 *
 * Policy, in order:
 *  - Only one side changed an entity  -> take that side.
 *  - Both sides made the SAME change  -> no conflict.
 *  - Both changed it differently      -> keep MINE, report a conflict.
 *  - One edited, the other deleted    -> keep the EDIT, report a conflict. Resurrecting
 *    an edited node is recoverable; silently discarding someone's edit is not.
 *  - Layout and collapse state        -> mine wins outright. Presentation is per-viewer
 *    and not worth a conflict prompt.
 */

export type MergeReason = 'both-edited' | 'edited-and-deleted';

export interface IMergeConflict {
  kind: 'node' | 'edge';
  id: string;
  label: string;
  reason: MergeReason;
}

export interface IMergeResult {
  merged: IGraph;
  conflicts: IMergeConflict[];
  /** Remote brought changes the local copy did not have. */
  tookRemoteChanges: boolean;
  /** Edges dropped because the merge left them dangling. */
  removedOrphanEdges: number;
}

const byId = <T extends { id: string }>(items: T[]): Map<string, T> => {
  const map = new Map<string, T>();
  for (const item of items || []) {
    if (item && typeof item.id === 'string') { map.set(item.id, item); }
  }
  return map;
};

interface IEntityMerge<T> {
  result: T[];
  conflicts: IMergeConflict[];
  tookRemote: boolean;
}

const mergeEntities = <T extends { id: string }>(
  kind: 'node' | 'edge',
  base: T[],
  mine: T[],
  theirs: T[],
  labelOf: (item: T) => string
): IEntityMerge<T> => {
  const b = byId(base);
  const m = byId(mine);
  const t = byId(theirs);
  const conflicts: IMergeConflict[] = [];
  const out = new Map<string, T>();
  let tookRemote = false;

  const allIds = new Set<string>();
  m.forEach((_v, k) => allIds.add(k));
  t.forEach((_v, k) => allIds.add(k));
  b.forEach((_v, k) => allIds.add(k));

  allIds.forEach((id) => {
    const inBase = b.get(id);
    const inMine = m.get(id);
    const inTheirs = t.get(id);

    // Present on both sides.
    if (inMine && inTheirs) {
      if (sameEntity(inMine, inTheirs)) { out.set(id, inMine); return; }
      const iChanged = !inBase || !sameEntity(inBase, inMine);
      const theyChanged = !inBase || !sameEntity(inBase, inTheirs);
      if (iChanged && theyChanged) {
        out.set(id, inMine);
        conflicts.push({ kind, id, label: labelOf(inMine), reason: 'both-edited' });
      } else if (theyChanged) {
        out.set(id, inTheirs);
        tookRemote = true;
      } else {
        out.set(id, inMine);
      }
      return;
    }

    // I deleted it.
    if (!inMine && inTheirs) {
      const theyChanged = !inBase || !sameEntity(inBase, inTheirs);
      if (inBase && !theyChanged) { return; }          // clean delete on my side
      if (!inBase) { out.set(id, inTheirs); tookRemote = true; return; }  // they created it
      out.set(id, inTheirs);                            // edit beats delete
      tookRemote = true;
      conflicts.push({ kind, id, label: labelOf(inTheirs), reason: 'edited-and-deleted' });
      return;
    }

    // They deleted it.
    if (inMine && !inTheirs) {
      const iChanged = !inBase || !sameEntity(inBase, inMine);
      if (inBase && !iChanged) { tookRemote = true; return; }   // clean delete on their side
      if (!inBase) { out.set(id, inMine); return; }              // I created it
      out.set(id, inMine);
      conflicts.push({ kind, id, label: labelOf(inMine), reason: 'edited-and-deleted' });
      return;
    }

    // Deleted on both sides — nothing to carry forward.
  });

  return { result: Array.from(out.values()), conflicts, tookRemote };
};

/** Union custom node types by their identifying key, preferring the local definition. */
const mergeCustomTypes = (mine: unknown[], theirs: unknown[]): unknown[] => {
  const keyOf = (t: unknown): string => {
    const obj = t as { id?: string; key?: string; name?: string; label?: string };
    return String((obj && (obj.id || obj.key || obj.name || obj.label)) || JSON.stringify(t));
  };
  const out = new Map<string, unknown>();
  for (const t of theirs || []) { out.set(keyOf(t), t); }
  for (const t of mine || []) { out.set(keyOf(t), t); }
  return Array.from(out.values());
};

export const threeWayMerge = (base: IGraph, mine: IGraph, theirs: IGraph): IMergeResult => {
  normalizeGraph(base); normalizeGraph(mine); normalizeGraph(theirs);

  const nodes = mergeEntities<IGraphNode>(
    'node', base.nodes, mine.nodes, theirs.nodes, (n) => n.label || n.id
  );
  const edges = mergeEntities<IGraphEdge>(
    'edge', base.edges, mine.edges, theirs.edges, (e) => e.type ? `${e.type} (${e.source}→${e.target})` : e.id
  );

  const draft: IGraph = {
    ...theirs,
    ...mine,
    nodes: nodes.result,
    edges: edges.result,
    customNodeTypes: mergeCustomTypes(mine.customNodeTypes || [], theirs.customNodeTypes || []),
    // Presentation: local view wins, no prompt.
    positions: { ...(theirs.positions || {}), ...(mine.positions || {}) },
    collapsedNodes: mine.collapsedNodes || []
  };

  const repaired = repairOrphans(draft);

  return {
    merged: repaired.graph,
    conflicts: nodes.conflicts.concat(edges.conflicts),
    tookRemoteChanges: nodes.tookRemote || edges.tookRemote,
    removedOrphanEdges: repaired.removedEdges
  };
};

export const describeConflict = (c: IMergeConflict): string =>
  c.reason === 'both-edited'
    ? `${c.label} was changed by both of you — your version was kept.`
    : `${c.label} was edited by one of you and deleted by the other — the edit was kept.`;
