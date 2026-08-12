import { SpRest } from './SpRest';
import { SNAPSHOTS_LIST } from '../../provisioning/schema';
import { stableStringify } from '../../model/bundle';
import { payloadPatch, payloadColumns } from './projectItems';
import { joinPayload } from '../../model/bundle';

/**
 * Named restore points, one row per snapshot, shared by both storage models.
 *
 * They used to travel inside the document payload, which had two consequences worth
 * undoing: every snapshot is a FULL COPY of the graph, so a handful of them could
 * exhaust the ~480KB a single list item can hold; and per-item projects dropped them
 * altogether, so switching storage model silently destroyed every restore point.
 *
 * Giving them their own list fixes both, makes them visible in SharePoint, and means
 * changing a project's storage model does not touch them at all.
 */

export interface ISnapshotRecord {
  id: string;
  name?: string;
  ts?: number;
  [key: string]: unknown;
}

interface ISnapshotItemDto {
  Id: number;
  Title: string | null;
  ErgSnapshotId: string | null;
  [payloadColumn: string]: unknown;
}

const snapshotsPath = (): string => `web/lists/getbytitle('${encodeURIComponent(SNAPSHOTS_LIST)}')`;

/** Content-derived id, for a snapshot that arrived without one. Stable, not unique-ish. */
const derivedId = (snapshot: ISnapshotRecord): string => {
  const text = stableStringify(snapshot);
  let hash = 5381;
  for (let i = 0; i < text.length; i++) {
    hash = (((hash << 5) + hash) ^ text.charCodeAt(i)) >>> 0;
  }
  return hash.toString(36);
};

export class SnapshotStore {
  /** Set when the site has no snapshot list, so we stop asking for the session. */
  private unavailable: boolean = false;

  public constructor(private readonly sp: SpRest) {}

  public get isUnavailable(): boolean { return this.unavailable; }

  private missingList(e: unknown): boolean {
    const message = e instanceof Error ? e.message : String(e);
    return /list\s+'[^']*'\s+does not exist/i.test(message) || /does not exist at site/i.test(message);
  }

  /**
   * Every snapshot for a project. Never throws: a graph must stay fully usable on a
   * site whose snapshot list has not been provisioned yet.
   */
  public async list(projectId: number): Promise<ISnapshotRecord[]> {
    if (this.unavailable) { return []; }
    try {
      // No $select — the payload columns are numerous and a site one version behind
      // would reject a query naming a column it lacks, taking every snapshot with it.
      const rows = await this.sp.getAll<ISnapshotItemDto>(
        `${snapshotsPath()}/items?$filter=ErgProjectId eq ${projectId}&$top=500`
      );
      const out: ISnapshotRecord[] = [];
      for (const row of rows) {
        const text = joinPayload(payloadColumns().map((c) => row[c] as string | null));
        if (!text) { continue; }
        try {
          const snap = JSON.parse(text) as ISnapshotRecord;
          if (snap && snap.id) { out.push(snap); }
        } catch { /* a corrupt row must not lose the others */ }
      }
      return out.sort((a, b) => (Number(b.ts) || 0) - (Number(a.ts) || 0));
    } catch (e) {
      if (this.missingList(e)) { this.unavailable = true; }
      return [];
    }
  }

  /**
   * Make the list match `snapshots`, writing only what actually differs.
   *
   * A snapshot's graph is immutable once taken — the engine only ever adds, renames or
   * deletes them — so an existing row is rewritten only when its NAME changed. That
   * keeps a save proportional to what the user did rather than to how many restore
   * points they happen to have.
   */
  public async sync(projectId: number, snapshots: ISnapshotRecord[]): Promise<number> {
    if (this.unavailable) { return 0; }
    // An id is how a row is matched to a snapshot across saves, so one without an id
    // cannot be tracked — but dropping it would silently destroy a restore point
    // somebody deliberately took. Derive a stable id from the content instead: the same
    // snapshot always yields the same id, so it neither duplicates nor disappears.
    const wanted = (snapshots || [])
      .filter((s) => !!s)
      .map((s) => (s.id ? s : { ...s, id: `snap-${derivedId(s)}` }));

    try {
      const rows = await this.sp.getAll<ISnapshotItemDto>(
        `${snapshotsPath()}/items?$select=Id,Title,ErgSnapshotId&$filter=ErgProjectId eq ${projectId}&$top=500`
      );
      const existing = new Map<string, { itemId: number; title: string }>();
      for (const row of rows) {
        if (row.ErgSnapshotId) {
          existing.set(row.ErgSnapshotId, { itemId: row.Id, title: row.Title || '' });
        }
      }

      let writes = 0;
      const keep = new Set<string>();

      for (const snap of wanted) {
        keep.add(snap.id);
        const name = String(snap.name || 'Snapshot').slice(0, 255);
        const known = existing.get(snap.id);

        if (!known) {
          const patch: { [k: string]: unknown } = payloadPatch(stableStringify(snap));
          patch.Title = name;
          patch.ErgSnapshotId = snap.id;
          patch.ErgProjectId = projectId;
          await this.sp.post(`${snapshotsPath()}/items`, patch);
          writes++;
        } else if (known.title !== name) {
          await this.sp.merge(`${snapshotsPath()}/items(${known.itemId})`, { Title: name }, '*');
          writes++;
        }
      }

      existing.forEach((row, id) => {
        if (!keep.has(id)) {
          // Fire and forget order does not matter here; a failed delete leaves a row
          // the next sync will try again.
          writes++;
          void this.sp.del(`${snapshotsPath()}/items(${row.itemId})`).catch(() => undefined);
        }
      });

      return writes;
    } catch (e) {
      if (this.missingList(e)) { this.unavailable = true; }
      return 0;
    }
  }
}
