import { SpRest } from './SpRest';
import { PRESENCE_LIST } from '../../provisioning/schema';

/**
 * Who else has this graph open.
 *
 * SharePoint offers no push channel here, so presence is a heartbeat: each open graph
 * refreshes one row for its viewer every `HEARTBEAT_MS`, and a row is treated as live
 * only while it is fresher than `STALE_MS`. Two missed beats and you disappear, which
 * is the right failure direction — a closed laptop should stop showing as present
 * without needing a clean goodbye.
 *
 * Cost is one small write per open graph per 30s, and the row count is naturally
 * bounded: Title is `<projectId>|<login>` and unique, so a person with six tabs open
 * still owns exactly one row.
 */

export const HEARTBEAT_MS = 30000;
export const STALE_MS = 75000;

export type PresenceMode = 'Viewing' | 'Editing';

export interface IPresentUser {
  login: string;
  name: string;
  /** Account name for the profile-photo endpoint; blank when unknown. */
  email: string;
  mode: PresenceMode;
  lastSeen: string;
  isSelf: boolean;
}

interface IPresenceItemDto {
  Id: number;
  Title: string;
  ErgUser: string | null;
  ErgLogin: string | null;
  ErgEmail: string | null;
  ErgMode: string | null;
  ErgHeartbeat: string | null;
}

const presencePath = (): string => `web/lists/getbytitle('${encodeURIComponent(PRESENCE_LIST)}')`;

export class PresenceService {
  /** projectId -> our own row id, so a heartbeat is one write after the first. */
  private readonly ownRow: Map<number, number> = new Map();
  /** Set once presence is known to be unavailable, so we stop retrying every tick. */
  private disabled: boolean = false;

  public constructor(
    private readonly sp: SpRest,
    private readonly login: string,
    private readonly displayName: string,
    private readonly email: string = ''
  ) {}

  /** Set when the site's presence list predates the ErgEmail column. */
  private omitEmail: boolean = false;

  public get isDisabled(): boolean { return this.disabled; }

  /**
   * Give presence another chance after the schema has been repaired.
   *
   * Disabling is deliberately sticky so a missing list is not retried every 30s for
   * the rest of the session — which means provisioning the list mid-session would
   * otherwise leave presence dead until a page reload.
   */
  public reset(): void {
    this.disabled = false;
    this.omitEmail = false;
    this.ownRow.clear();
  }

  private key(projectId: number): string { return `${projectId}|${this.login}`; }

  /**
   * Only a MISSING LIST disables presence.
   *
   * SharePoint says "does not exist" for a missing column too, and treating that the
   * same way would switch presence off permanently on any site whose list is one
   * column behind — silently, and exactly when a re-deploy would have fixed it.
   */
  private isMissingList(message: string): boolean {
    return /list\s+'[^']*'\s+does not exist/i.test(message) ||
      /does not exist at site/i.test(message);
  }

  private isUnknownColumn(message: string): boolean {
    return /field or property/i.test(message) || /column\s+'[^']*'\s+does not exist/i.test(message);
  }

  private rowBody(projectId: number, mode: PresenceMode): { [k: string]: unknown } {
    const body: { [k: string]: unknown } = {
      Title: this.key(projectId),
      ErgProjectId: projectId,
      ErgUser: this.displayName,
      ErgLogin: this.login,
      ErgMode: mode,
      ErgHeartbeat: new Date().toISOString()
    };
    if (!this.omitEmail && this.email) { body.ErgEmail = this.email; }
    return body;
  }

  /**
   * Refresh our row. Never throws: presence is a nicety, and a graph must stay fully
   * usable on a site where the presence list was not provisioned or is not writable.
   */
  public async heartbeat(projectId: number, mode: PresenceMode): Promise<void> {
    if (this.disabled) { return; }
    try {
      await this.writeRow(projectId, mode);
    } catch (e) {
      const message = e instanceof Error ? e.message : String(e);

      if (this.isMissingList(message)) { this.disabled = true; return; }

      // The list exists but is a column behind — a site provisioned before the photo
      // support landed. Drop the optional column and try once more, permanently, so
      // presence works without waiting for anyone to re-run setup.
      if (this.isUnknownColumn(message) && !this.omitEmail) {
        this.omitEmail = true;
        try { await this.writeRow(projectId, mode); return; } catch { /* falls through */ }
      }
      // Anything else is transient; the next beat retries.
    }
  }

  private async writeRow(projectId: number, mode: PresenceMode): Promise<void> {
    const body = this.rowBody(projectId, mode);

    const known = this.ownRow.get(projectId);
    if (known) {
      await this.sp.merge(`${presencePath()}/items(${known})`, body, '*');
      return;
    }

    // First beat for this project: adopt an existing row if we already have one
    // (a previous session, or another tab), otherwise create it.
    const existing = await this.sp.getAll<IPresenceItemDto>(
      `${presencePath()}/items?$select=Id,Title&$filter=Title eq '${encodeURIComponent(this.key(projectId))}'&$top=1`
    );
    if (existing.length > 0) {
      this.ownRow.set(projectId, existing[0].Id);
      await this.sp.merge(`${presencePath()}/items(${existing[0].Id})`, body, '*');
      return;
    }
    const created = await this.sp.post<{ Id: number }>(`${presencePath()}/items`, body);
    if (created && created.Id) { this.ownRow.set(projectId, created.Id); }
  }

  /** Everyone currently on this graph, most recently seen first. */
  public async list(projectId: number): Promise<IPresentUser[]> {
    if (this.disabled) { return []; }
    try {
      // NO $select. Naming a column the list does not have yet — ErgEmail on a site
      // provisioned before photo support — makes SharePoint reject the WHOLE query
      // with 400, so presence would show nobody at all rather than degrade. The rows
      // are half a dozen short fields; asking for all of them costs nothing.
      const rows = await this.sp.getAll<IPresenceItemDto>(
        `${presencePath()}/items?$filter=ErgProjectId eq ${projectId}&$top=200`
      );
      const cutoff = Date.now() - STALE_MS;
      return rows
        // Staleness is filtered here rather than in $filter: OData datetime literals
        // are a reliable source of 400s across tenants, and the row count is small.
        .filter((r) => !!r.ErgHeartbeat && Date.parse(r.ErgHeartbeat) >= cutoff)
        .map((r) => ({
          login: r.ErgLogin || '',
          name: r.ErgUser || r.ErgLogin || 'Someone',
          email: r.ErgEmail || '',
          mode: (r.ErgMode === 'Editing' ? 'Editing' : 'Viewing') as PresenceMode,
          lastSeen: r.ErgHeartbeat as string,
          isSelf: (r.ErgLogin || '') === this.login
        }))
        .sort((a, b) => Date.parse(b.lastSeen) - Date.parse(a.lastSeen));
    } catch (e) {
      const message = e instanceof Error ? e.message : String(e);
      if (this.isMissingList(message)) { this.disabled = true; }
      return [];
    }
  }

  /** Best-effort goodbye on close or project switch; staleness covers the rest. */
  public async leave(projectId: number): Promise<void> {
    const known = this.ownRow.get(projectId);
    if (!known || this.disabled) { return; }
    this.ownRow.delete(projectId);
    try { await this.sp.del(`${presencePath()}/items(${known})`); } catch { /* stale-out covers it */ }
  }
}
