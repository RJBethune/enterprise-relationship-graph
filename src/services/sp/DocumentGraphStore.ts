import { SpRest, ConflictError } from './SpRest';
import {
  IGraphStore, IOpenProject, IProjectSummary, ISaveOutcome, IRemoteUpdate, StorageMode
} from '../IGraphStore';
import {
  IGraph, IBundle, BUNDLE_SCHEMA_VERSION, parseBundle, normalizeGraph, stableStringify,
  MAX_DOCUMENT_CHARS
} from '../../model/bundle';
import { threeWayMerge } from '../../model/merge';
import {
  projectsPath, listProjectItems, createProjectItem, readProjectItem, readProjectStamp,
  readPayload, payloadPatch, toSummary, IProjectItemDto
} from './projectItems';

/**
 * Whole graph in one list item.
 *
 * Every save is one atomic write, guarded by the item's ETag. That single property is
 * what makes this the default: a layout run, a Mermaid import and a rename all cost
 * exactly one operation, and a failure leaves the stored graph untouched rather than
 * half-updated. SharePoint's own version history becomes a free save ladder.
 *
 * Conflicts are resolved by three-way merge rather than by asking a human, because for
 * two people editing different parts of an org graph there is nothing to ask about.
 * Only a genuine same-entity collision produces a prompt.
 */
export class DocumentGraphStore implements IGraphStore {
  public constructor(private readonly sp: SpRest, private readonly editorName: string) {}

  public listProjects(): Promise<IProjectSummary[]> { return listProjectItems(this.sp); }

  public createProject(title: string, mode: StorageMode = 'Document'): Promise<IProjectSummary> {
    return createProjectItem(this.sp, title, mode);
  }

  public async renameProject(id: number, title: string): Promise<void> {
    await this.sp.merge(`${projectsPath()}/items(${id})`, { Title: title });
  }

  public async archiveProject(id: number): Promise<void> {
    await this.sp.merge(`${projectsPath()}/items(${id})`, { ErgStatus: 'Archived' });
  }

  public async openProject(id: number): Promise<IOpenProject> {
    const { dto, etag } = await readProjectItem(this.sp, id);
    const bundle = parseBundle(readPayload(dto));
    return new DocumentProject(this.sp, this.editorName, toSummary(dto), bundle, etag);
  }
}

class DocumentProject implements IOpenProject {
  /** The graph as it was last known to be stored — the merge base. */
  private base: IGraph;
  private lastModified: string | null;

  public constructor(
    private readonly sp: SpRest,
    private readonly editorName: string,
    public readonly summary: IProjectSummary,
    public readonly bundle: IBundle,
    private etag: string | null
  ) {
    this.base = JSON.parse(JSON.stringify(normalizeGraph(bundle.graph))) as IGraph;
    this.lastModified = summary.modified;
  }

  public async save(graph: IGraph, snapshots: unknown[] | null): Promise<ISaveOutcome> {
    try {
      const chars = await this.write(graph, snapshots, this.etag);
      this.base = JSON.parse(JSON.stringify(graph)) as IGraph;
      return { status: 'saved', writes: 1, warning: this.ceilingWarning(chars) };
    } catch (e) {
      if (!(e instanceof ConflictError)) {
        return { status: 'error', message: e instanceof Error ? e.message : String(e) };
      }
      return this.resolveAndRetry(graph, snapshots);
    }
  }

  /**
   * Somebody saved between our load and our write. Pull their version, merge ours on
   * top of the common ancestor, and write the result — the "git pull before push" the
   * document model can afford precisely because the whole graph arrives in one read.
   */
  private async resolveAndRetry(graph: IGraph, snapshots: unknown[] | null): Promise<ISaveOutcome> {
    const { dto, etag } = await readProjectItem(this.sp, this.summary.id);
    const remote = parseBundle(readPayload(dto));
    const merge = threeWayMerge(this.base, normalizeGraph(graph), normalizeGraph(remote.graph));

    try {
      await this.write(merge.merged, snapshots, etag);
    } catch {
      // Losing the race twice means somebody is saving continuously. Report rather
      // than loop: an unbounded retry here is how you build a write storm.
      return {
        status: 'conflict',
        graph: merge.merged,
        conflicts: merge.conflicts,
        message: 'Saved again while merging. Your changes are merged locally — press save once more.'
      };
    }

    this.base = JSON.parse(JSON.stringify(merge.merged)) as IGraph;
    return {
      status: 'merged',
      graph: merge.merged,
      conflicts: merge.conflicts,
      writes: 1,
      message: merge.conflicts.length === 0
        ? `Merged changes from ${dto.Editor && dto.Editor.Title ? dto.Editor.Title : 'another editor'}.`
        : undefined
    };
  }

  /**
   * Document storage has a hard ceiling, and hitting it means saves simply stop. That
   * is a bad moment to learn about it, so say something while there is still room to
   * act — the only remedy is moving the project to per-item storage.
   */
  private ceilingWarning(chars: number): string | undefined {
    const used = chars / MAX_DOCUMENT_CHARS;
    if (used < 0.75) { return undefined; }
    return `This graph is using ${Math.round(used * 100)}% of what a single list item can hold. ` +
      'Past 100% saves will start failing; move the project to per-item storage before then.';
  }

  private async write(graph: IGraph, snapshots: unknown[] | null, etag: string | null): Promise<number> {
    const now = new Date().toISOString();
    const bundle: IBundle = {
      version: BUNDLE_SCHEMA_VERSION,
      exportedAt: now,
      lastModifiedBy: this.editorName,
      lastModifiedAt: now,
      graph,
      snapshots: snapshots || []
    };
    const serialized = JSON.stringify(bundle);
    const patch: { [k: string]: unknown } = payloadPatch(serialized);
    patch.ErgSchemaVersion = BUNDLE_SCHEMA_VERSION;
    patch.ErgNodeCount = (graph.nodes || []).length;
    patch.ErgEdgeCount = (graph.edges || []).length;

    await this.sp.merge(`${projectsPath()}/items(${this.summary.id})`, patch, etag || '*');

    // The write invalidated our etag; re-read the cheap stamp so the next save has a
    // current one. Without this every save after the first would false-conflict.
    const stamp = await readProjectStamp(this.sp, this.summary.id);
    this.etag = stamp.etag;
    this.lastModified = stamp.modified;
    return serialized.length;
  }

  /**
   * Cheap liveness check: compare the item's Modified stamp. Only when it moves do we
   * pay for a full read. A poll that always pulled the whole payload would be a
   * bandwidth problem on a large graph with several viewers.
   */
  public async poll(localGraph: IGraph): Promise<IRemoteUpdate | null> {
    const stamp = await readProjectStamp(this.sp, this.summary.id);
    if (!stamp.modified || stamp.modified === this.lastModified) { return null; }

    const { dto, etag } = await readProjectItem(this.sp, this.summary.id);
    const remote = parseBundle(readPayload(dto));
    this.lastModified = stamp.modified;
    this.etag = etag;

    const local = normalizeGraph(localGraph);
    const remoteGraph = normalizeGraph(remote.graph);

    // Nothing local to protect: adopt theirs wholesale.
    if (stableStringify(this.base) === stableStringify(local)) {
      this.base = JSON.parse(JSON.stringify(remoteGraph)) as IGraph;
      return { graph: remoteGraph, conflicts: [], by: stamp.modifiedBy };
    }

    const merge = threeWayMerge(this.base, local, remoteGraph);
    this.base = JSON.parse(JSON.stringify(remoteGraph)) as IGraph;
    return { graph: merge.merged, conflicts: merge.conflicts, by: stamp.modifiedBy };
  }

  public dispose(): void { /* no timers or listeners of its own */ }
}

export type { IProjectItemDto };
