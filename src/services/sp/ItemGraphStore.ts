import { SpRest } from './SpRest';
import { IBatchRequest } from './batch';
import {
  IGraphStore, IOpenProject, IProjectSummary, ISaveOutcome, IRemoteUpdate, StorageMode
} from '../IGraphStore';
import {
  IGraph, IGraphNode, IGraphEdge, IBundle, BUNDLE_SCHEMA_VERSION, normalizeGraph, stableStringify
} from '../../model/bundle';
import { GraphOp, opKey, orderOps } from '../../model/ops';
import { diffGraphs, applyOps, hasLayoutChange, repairOrphans } from '../../model/diff';
import { threeWayMerge } from '../../model/merge';
import { WriteQueue } from '../WriteQueue';
import { NODES_LIST, EDGES_LIST } from '../../provisioning/schema';
import {
  projectsPath, listProjectItems, createProjectItem, readProjectItem, toSummary, IProjectItemDto
} from './projectItems';

/**
 * One list item per node and per relationship, with change-log sync.
 *
 * This is the model that earns its complexity only when several people edit the same
 * graph at once: two editors touching different nodes never collide, and each other's
 * changes appear within a poll interval. Everything expensive about it is contained by
 * three decisions:
 *
 *  1. LAYOUT IS NOT AN ITEM. Positions live in one blob on the project item. A force
 *     layout moves every node; as per-item writes that is one operation per node per
 *     run — the write storm that makes naive per-item storage unusable. As a blob it
 *     is one write, whatever the graph's size.
 *  2. ATTRIBUTES ARE NOT COLUMNS. The whole node travels as JSON in ErgData, so an
 *     office inventing a custom node type never triggers a schema change. The few
 *     queryable fields are projected alongside for list views.
 *  3. THE CHANGE LOG, NOT A SCAN. GetChanges returns only what moved since a token,
 *     including deletes — which a `Modified gt` query cannot see at all. Polling is
 *     therefore proportional to the change rate, not the graph size.
 *
 * What it cannot offer is atomicity: SharePoint has no transactions, so a failed flush
 * can leave a partial write. Ordering (edges deleted before nodes, nodes written before
 * edges) plus an orphan sweep on load is the mitigation, and it is a mitigation, not a
 * guarantee. That trade is the reason Document mode remains the default.
 */

interface INodeItemDto {
  Id: number; Title: string; ErgNodeId: string; ErgType: string; ErgData: string;
}
interface IEdgeItemDto {
  Id: number; Title: string; ErgEdgeId: string; ErgSource: string; ErgTarget: string;
  ErgType: string; ErgData: string;
}
interface IChangeDto {
  ChangeType: number;
  ItemId?: number;
  ChangeToken?: { StringValue: string };
}

/** SP.ChangeType */
const CHANGE_ADD = 1;
const CHANGE_UPDATE = 2;
const CHANGE_DELETE = 3;

const nodesPath = (): string => `web/lists/getbytitle('${encodeURIComponent(NODES_LIST)}')`;
const edgesPath = (): string => `web/lists/getbytitle('${encodeURIComponent(EDGES_LIST)}')`;

const parseData = <T>(json: string, fallback: T): T => {
  try { return JSON.parse(json) as T; } catch (_e) { return fallback; }
};

export class ItemGraphStore implements IGraphStore {
  public constructor(private readonly sp: SpRest, private readonly editorName: string) {}

  public listProjects(): Promise<IProjectSummary[]> { return listProjectItems(this.sp); }

  public createProject(title: string, mode: StorageMode = 'Items'): Promise<IProjectSummary> {
    return createProjectItem(this.sp, title, mode);
  }

  public async renameProject(id: number, title: string): Promise<void> {
    await this.sp.merge(`${projectsPath()}/items(${id})`, { Title: title });
  }

  public async archiveProject(id: number): Promise<void> {
    await this.sp.merge(`${projectsPath()}/items(${id})`, { ErgStatus: 'Archived' });
  }

  public async openProject(id: number): Promise<IOpenProject> {
    const { dto } = await readProjectItem(this.sp, id);

    const [nodeItems, edgeItems] = await Promise.all([
      this.sp.getAll<INodeItemDto>(
        `${nodesPath()}/items?$select=Id,Title,ErgNodeId,ErgType,ErgData` +
        `&$filter=ErgProjectId eq ${id}&$top=5000`
      ),
      this.sp.getAll<IEdgeItemDto>(
        `${edgesPath()}/items?$select=Id,Title,ErgEdgeId,ErgSource,ErgTarget,ErgType,ErgData` +
        `&$filter=ErgProjectId eq ${id}&$top=5000`
      )
    ]);

    const nodeItemIds = new Map<string, number>();
    const edgeItemIds = new Map<string, number>();
    const nodes: IGraphNode[] = [];
    const edges: IGraphEdge[] = [];

    for (const item of nodeItems) {
      const node = parseData<IGraphNode | null>(item.ErgData, null);
      if (!node || !node.id) { continue; }
      nodes.push(node);
      nodeItemIds.set(node.id, item.Id);
    }
    for (const item of edgeItems) {
      const edge = parseData<IGraphEdge | null>(item.ErgData, null);
      if (!edge || !edge.id) { continue; }
      edges.push(edge);
      edgeItemIds.set(edge.id, item.Id);
    }

    const assembled: IGraph = normalizeGraph({
      nodes,
      edges,
      customNodeTypes: parseData<unknown[]>(dto.ErgCustomTypes || '[]', []),
      collapsedNodes: parseData<string[]>(dto.ErgCollapsed || '[]', []),
      positions: parseData<{ [k: string]: { x: number; y: number } }>(dto.ErgLayout || '{}', {})
    });

    // A partial flush or a concurrent delete can leave an edge pointing at a node that
    // is gone. Sweeping on load is cheaper and far more reliable than pretending
    // distributed writes were atomic.
    const repaired = repairOrphans(assembled);

    const [nodeToken, edgeToken] = await Promise.all([
      this.currentToken(nodesPath()),
      this.currentToken(edgesPath())
    ]);

    const bundle: IBundle = {
      version: BUNDLE_SCHEMA_VERSION,
      graph: repaired.graph,
      snapshots: []
    };

    return new ItemProject(
      this.sp, this.editorName, toSummary(dto), bundle,
      nodeItemIds, edgeItemIds, nodeToken, edgeToken, repaired.removedEdges
    );
  }

  private async currentToken(listPath: string): Promise<string | null> {
    try {
      const res = await this.sp.get<{ CurrentChangeToken?: { StringValue?: string } }>(
        `${listPath}?$select=CurrentChangeToken`
      );
      return (res.CurrentChangeToken && res.CurrentChangeToken.StringValue) || null;
    } catch (_e) {
      return null;
    }
  }
}

class ItemProject implements IOpenProject {
  private base: IGraph;
  private readonly queue: WriteQueue<GraphOp>;
  /** graph id -> SharePoint item id, and the reverse, for interpreting delete changes. */
  private readonly nodeItems: Map<string, number>;
  private readonly edgeItems: Map<string, number>;
  private readonly nodeIdByItem: Map<number, string> = new Map();
  private readonly edgeIdByItem: Map<number, string> = new Map();

  public constructor(
    private readonly sp: SpRest,
    private readonly editorName: string,
    public readonly summary: IProjectSummary,
    public readonly bundle: IBundle,
    nodeItems: Map<string, number>,
    edgeItems: Map<string, number>,
    private nodeToken: string | null,
    private edgeToken: string | null,
    public readonly orphansRemovedOnLoad: number
  ) {
    this.base = JSON.parse(JSON.stringify(bundle.graph)) as IGraph;
    this.nodeItems = nodeItems;
    this.edgeItems = edgeItems;
    nodeItems.forEach((itemId, graphId) => this.nodeIdByItem.set(itemId, graphId));
    edgeItems.forEach((itemId, graphId) => this.edgeIdByItem.set(itemId, graphId));

    this.queue = new WriteQueue<GraphOp>({
      keyOf: opKey,
      order: orderOps,
      // The shell debounces save() already; this queue exists for coalescing,
      // batching and the retry ladder, so it should not add a second delay.
      debounceMs: 0,
      batchSize: 50,
      flush: (ops) => this.flushOps(ops)
    });
  }

  public async save(graph: IGraph, _snapshots: unknown[] | null): Promise<ISaveOutcome> {
    const next = normalizeGraph(graph);
    const ops = diffGraphs(this.base, next);
    const layoutMoved = hasLayoutChange(this.base, next);

    if (ops.length === 0 && !layoutMoved) { return { status: 'saved', writes: 0 }; }

    try {
      if (ops.length > 0) {
        this.queue.enqueue(ops);
        await this.queue.flushNow();
        // Ask the queue whether the work LANDED, rather than whether an error was ever
        // seen: the retry ladder exists precisely so that a recovered throttle is a
        // success, and treating the last error as fatal would report every survived
        // 429 as a failed save.
        if (this.queue.getPendingCount() > 0) {
          throw this.queue.getLastError() || new Error('Some changes could not be saved.');
        }
      }
      // Presentation travels as one blob no matter how many nodes moved.
      if (layoutMoved || ops.length > 0) { await this.writeProjectBlob(next); }
    } catch (e) {
      return { status: 'error', message: e instanceof Error ? e.message : String(e) };
    }

    this.base = JSON.parse(JSON.stringify(next)) as IGraph;
    return { status: 'saved', writes: ops.length + (layoutMoved ? 1 : 0) };
  }

  private async writeProjectBlob(graph: IGraph): Promise<void> {
    await this.sp.merge(`${projectsPath()}/items(${this.summary.id})`, {
      ErgLayout: JSON.stringify(graph.positions || {}),
      ErgCustomTypes: JSON.stringify(graph.customNodeTypes || []),
      ErgCollapsed: JSON.stringify(graph.collapsedNodes || []),
      ErgSchemaVersion: BUNDLE_SCHEMA_VERSION,
      ErgNodeCount: (graph.nodes || []).length,
      ErgEdgeCount: (graph.edges || []).length
    }, '*');
  }

  /** Translate ops into one batch, then record the item ids SharePoint assigned. */
  private async flushOps(ops: GraphOp[]): Promise<void> {
    const requests: IBatchRequest[] = [];
    const creates: { kind: 'node' | 'edge'; graphId: string }[] = [];

    for (const op of ops) {
      switch (op.kind) {
        case 'upsertNode': {
          const existing = this.nodeItems.get(op.node.id);
          const body = {
            Title: (op.node.label || op.node.id).slice(0, 255),
            ErgNodeId: op.node.id,
            ErgType: op.node.type || '',
            ErgData: JSON.stringify(op.node),
            ErgProjectId: this.summary.id
          };
          if (existing) {
            requests.push({ method: 'MERGE', url: this.sp.api(`${nodesPath()}/items(${existing})`), body, etag: '*' });
          } else {
            requests.push({ method: 'POST', url: this.sp.api(`${nodesPath()}/items`), body });
            creates.push({ kind: 'node', graphId: op.node.id });
          }
          break;
        }
        case 'upsertEdge': {
          const existing = this.edgeItems.get(op.edge.id);
          const body = {
            Title: `${op.edge.type || 'RELATES'}: ${op.edge.source} → ${op.edge.target}`.slice(0, 255),
            ErgEdgeId: op.edge.id,
            ErgSource: op.edge.source,
            ErgTarget: op.edge.target,
            ErgType: op.edge.type || '',
            ErgData: JSON.stringify(op.edge),
            ErgProjectId: this.summary.id
          };
          if (existing) {
            requests.push({ method: 'MERGE', url: this.sp.api(`${edgesPath()}/items(${existing})`), body, etag: '*' });
          } else {
            requests.push({ method: 'POST', url: this.sp.api(`${edgesPath()}/items`), body });
            creates.push({ kind: 'edge', graphId: op.edge.id });
          }
          break;
        }
        case 'deleteNode': {
          const existing = this.nodeItems.get(op.id);
          if (existing) {
            requests.push({ method: 'DELETE', url: this.sp.api(`${nodesPath()}/items(${existing})`), etag: '*' });
            this.nodeItems.delete(op.id);
            this.nodeIdByItem.delete(existing);
          }
          break;
        }
        case 'deleteEdge': {
          const existing = this.edgeItems.get(op.id);
          if (existing) {
            requests.push({ method: 'DELETE', url: this.sp.api(`${edgesPath()}/items(${existing})`), etag: '*' });
            this.edgeItems.delete(op.id);
            this.edgeIdByItem.delete(existing);
          }
          break;
        }
        default: break;
      }
    }

    if (requests.length === 0) { return; }
    const results = await this.sp.batchOrThrow(requests);

    // Creates come back with their new Id; without recording it, the next edit to the
    // same node would create a duplicate row instead of updating this one.
    let createIndex = 0;
    for (let i = 0; i < results.length && createIndex < creates.length; i++) {
      const created = results[i].json as { Id?: number } | null;
      if (results[i].status !== 201 || !created || typeof created.Id !== 'number') { continue; }
      const target = creates[createIndex++];
      if (target.kind === 'node') {
        this.nodeItems.set(target.graphId, created.Id);
        this.nodeIdByItem.set(created.Id, target.graphId);
      } else {
        this.edgeItems.set(target.graphId, created.Id);
        this.edgeIdByItem.set(created.Id, target.graphId);
      }
    }
  }

  /**
   * Ask both lists what changed since our tokens, translate that into ops, and merge
   * them against the local graph.
   *
   * Our own writes appear in the change log too. They need no special handling: an op
   * that reproduces what `base` already says produces no difference, so self-changes
   * collapse to nothing without any author filtering.
   */
  public async poll(localGraph: IGraph): Promise<IRemoteUpdate | null> {
    const [nodeChanges, edgeChanges] = await Promise.all([
      this.fetchChanges(nodesPath(), this.nodeToken),
      this.fetchChanges(edgesPath(), this.edgeToken)
    ]);

    if (nodeChanges.token) { this.nodeToken = nodeChanges.token; }
    if (edgeChanges.token) { this.edgeToken = edgeChanges.token; }
    if (nodeChanges.changes.length === 0 && edgeChanges.changes.length === 0) { return null; }

    const ops = await this.changesToOps(nodeChanges.changes, edgeChanges.changes);
    if (ops.length === 0) { return null; }

    const remoteGraph = applyOps(this.base, ops);
    // Our own writes come back through the change log too. Rather than filtering by
    // author — which misses the case of the same person in two tabs — compare the
    // result: if replaying the changes reproduces what we already had, nothing moved.
    if (stableStringify(remoteGraph) === stableStringify(this.base)) { return null; }

    const local = normalizeGraph(localGraph);

    if (stableStringify(this.base) === stableStringify(local)) {
      // Nothing local at risk — adopt the remote state directly.
      this.base = JSON.parse(JSON.stringify(remoteGraph)) as IGraph;
      return { graph: this.withLocalPresentation(remoteGraph, local), conflicts: [], by: null };
    }

    const merge = threeWayMerge(this.base, local, remoteGraph);
    this.base = JSON.parse(JSON.stringify(remoteGraph)) as IGraph;
    return { graph: this.withLocalPresentation(merge.merged, local), conflicts: merge.conflicts, by: null };
  }

  /** Remote structure, local view. Nobody wants their viewport yanked by a colleague. */
  private withLocalPresentation(graph: IGraph, local: IGraph): IGraph {
    return { ...graph, positions: { ...(graph.positions || {}), ...(local.positions || {}) } };
  }

  private async fetchChanges(
    listPath: string, token: string | null
  ): Promise<{ changes: IChangeDto[]; token: string | null }> {
    if (!token) { return { changes: [], token: null }; }
    try {
      const res = await this.sp.post<{ value?: IChangeDto[] }>(`${listPath}/getchanges`, {
        query: {
          Item: true, Add: true, Update: true, DeleteObject: true,
          ChangeTokenStart: { StringValue: token }
        }
      });
      const changes = (res && res.value) || [];
      const last = changes[changes.length - 1];
      return {
        changes,
        token: (last && last.ChangeToken && last.ChangeToken.StringValue) || token
      };
    } catch (_e) {
      // A failed poll must never break editing — the next tick tries again.
      return { changes: [], token };
    }
  }

  private async changesToOps(nodeChanges: IChangeDto[], edgeChanges: IChangeDto[]): Promise<GraphOp[]> {
    const ops: GraphOp[] = [];
    const nodeIdsToRead: number[] = [];
    const edgeIdsToRead: number[] = [];

    for (const c of nodeChanges) {
      if (!c.ItemId) { continue; }
      if (c.ChangeType === CHANGE_DELETE) {
        const graphId = this.nodeIdByItem.get(c.ItemId);
        if (graphId) {
          ops.push({ kind: 'deleteNode', id: graphId });
          this.nodeItems.delete(graphId);
          this.nodeIdByItem.delete(c.ItemId);
        }
      } else if (c.ChangeType === CHANGE_ADD || c.ChangeType === CHANGE_UPDATE) {
        nodeIdsToRead.push(c.ItemId);
      }
    }
    for (const c of edgeChanges) {
      if (!c.ItemId) { continue; }
      if (c.ChangeType === CHANGE_DELETE) {
        const graphId = this.edgeIdByItem.get(c.ItemId);
        if (graphId) {
          ops.push({ kind: 'deleteEdge', id: graphId });
          this.edgeItems.delete(graphId);
          this.edgeIdByItem.delete(c.ItemId);
        }
      } else if (c.ChangeType === CHANGE_ADD || c.ChangeType === CHANGE_UPDATE) {
        edgeIdsToRead.push(c.ItemId);
      }
    }

    if (nodeIdsToRead.length > 0) {
      const items = await this.readByIds<INodeItemDto>(
        nodesPath(), 'Id,Title,ErgNodeId,ErgType,ErgData', nodeIdsToRead
      );
      for (const item of items) {
        const node = parseData<IGraphNode | null>(item.ErgData, null);
        if (!node || !node.id) { continue; }
        ops.push({ kind: 'upsertNode', node });
        this.nodeItems.set(node.id, item.Id);
        this.nodeIdByItem.set(item.Id, node.id);
      }
    }
    if (edgeIdsToRead.length > 0) {
      const items = await this.readByIds<IEdgeItemDto>(
        edgesPath(), 'Id,Title,ErgEdgeId,ErgSource,ErgTarget,ErgType,ErgData', edgeIdsToRead
      );
      for (const item of items) {
        const edge = parseData<IGraphEdge | null>(item.ErgData, null);
        if (!edge || !edge.id) { continue; }
        ops.push({ kind: 'upsertEdge', edge });
        this.edgeItems.set(edge.id, item.Id);
        this.edgeIdByItem.set(item.Id, edge.id);
      }
    }

    return orderOps(ops);
  }

  /**
   * Read specific items by id, chunked.
   *
   * `Id eq 1 or Id eq 2 or …` hits SharePoint's filter-complexity ceiling somewhere
   * above a hundred clauses, and a busy minute can easily change more rows than that.
   * Chunking keeps a large remote change set working instead of failing the whole poll.
   */
  private async readByIds<T>(listPath: string, select: string, ids: number[]): Promise<T[]> {
    const CHUNK = 40;
    const out: T[] = [];
    for (let i = 0; i < ids.length; i += CHUNK) {
      const filter = ids.slice(i, i + CHUNK).map((id) => `Id eq ${id}`).join(' or ');
      const page = await this.sp.getAll<T>(
        `${listPath}/items?$select=${select}&$filter=${encodeURIComponent(filter)}&$top=${CHUNK}`
      );
      out.push(...page);
    }
    return out;
  }

  public dispose(): void { this.queue.dispose(); }
}

export type { IProjectItemDto };
