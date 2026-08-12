import { SpRest } from './SpRest';
import { IProjectSummary, StorageMode } from '../IGraphStore';
import { PROJECTS_LIST } from '../../provisioning/schema';
import { MAX_CHUNKS, splitPayload, joinPayload } from '../../model/bundle';

/** Shared reading/writing of the ERG Projects item, used by both storage models. */

export interface IProjectItemDto {
  Id: number;
  Title: string;
  ErgStorageMode: string | null;
  ErgSchemaVersion: number | null;
  ErgNodeCount: number | null;
  ErgEdgeCount: number | null;
  ErgStatus: string | null;
  ErgLayout: string | null;
  ErgCustomTypes: string | null;
  ErgCollapsed: string | null;
  Modified: string | null;
  Editor?: { Title?: string } | null;
  [payloadColumn: string]: unknown;
}

export const projectsPath = (): string => `web/lists/getbytitle('${encodeURIComponent(PROJECTS_LIST)}')`;

export const payloadColumns = (): string[] => {
  const cols: string[] = [];
  for (let i = 1; i <= MAX_CHUNKS; i++) { cols.push(`ErgPayload${i}`); }
  return cols;
};

const SUMMARY_SELECT =
  'Id,Title,ErgStorageMode,ErgSchemaVersion,ErgNodeCount,ErgEdgeCount,ErgStatus,Modified,Editor/Title';

export const toSummary = (dto: IProjectItemDto): IProjectSummary => ({
  id: dto.Id,
  title: dto.Title,
  storageMode: (dto.ErgStorageMode === 'Items' ? 'Items' : 'Document') as StorageMode,
  nodeCount: dto.ErgNodeCount || 0,
  edgeCount: dto.ErgEdgeCount || 0,
  status: dto.ErgStatus || 'Active',
  modified: dto.Modified || null,
  modifiedBy: (dto.Editor && dto.Editor.Title) || null
});

export const listProjectItems = async (sp: SpRest): Promise<IProjectSummary[]> => {
  const items = await sp.getAll<IProjectItemDto>(
    `${projectsPath()}/items?$select=${SUMMARY_SELECT}&$expand=Editor` +
    `&$filter=ErgStatus ne 'Archived'&$orderby=Title&$top=500`
  );
  return items.map(toSummary);
};

export const createProjectItem = async (
  sp: SpRest, title: string, mode: StorageMode
): Promise<IProjectSummary> => {
  const created = await sp.post<IProjectItemDto>(`${projectsPath()}/items`, {
    Title: title,
    ErgStorageMode: mode,
    ErgStatus: 'Active',
    ErgSchemaVersion: 4,
    ErgNodeCount: 0,
    ErgEdgeCount: 0
  });
  return toSummary(created);
};

/** Read one project item with every column both storage models might need. */
export const readProjectItem = async (
  sp: SpRest, id: number
): Promise<{ dto: IProjectItemDto; etag: string | null }> => {
  const select = [SUMMARY_SELECT, 'ErgLayout', 'ErgCustomTypes', 'ErgCollapsed']
    .concat(payloadColumns()).join(',');
  const res = await sp.getWithEtag<IProjectItemDto>(
    `${projectsPath()}/items(${id})?$select=${select}&$expand=Editor`
  );
  return { dto: res.data, etag: res.etag };
};

/** Just the concurrency-relevant bits — used by the poll loop, which runs often. */
export const readProjectStamp = async (
  sp: SpRest, id: number
): Promise<{ modified: string | null; modifiedBy: string | null; etag: string | null }> => {
  const res = await sp.getWithEtag<IProjectItemDto>(
    `${projectsPath()}/items(${id})?$select=Id,Modified,Editor/Title&$expand=Editor`
  );
  return {
    modified: res.data.Modified || null,
    modifiedBy: (res.data.Editor && res.data.Editor.Title) || null,
    etag: res.etag
  };
};

export const readPayload = (dto: IProjectItemDto): string =>
  joinPayload(payloadColumns().map((c) => dto[c] as string | null));

/**
 * Build the payload column patch.
 *
 * Unused trailing columns are explicitly blanked. Without that, shrinking a graph
 * leaves the tail of the previous, larger payload in ErgPayload5..8, and the next read
 * concatenates live data with stale data into unparseable JSON — a corruption that only
 * shows up after a graph gets smaller, which is exactly when nobody is looking for it.
 */
export const payloadPatch = (text: string): { [column: string]: string } => {
  const chunks = splitPayload(text);
  const patch: { [column: string]: string } = {};
  const cols = payloadColumns();
  for (let i = 0; i < cols.length; i++) {
    patch[cols[i]] = i < chunks.length ? chunks[i] : '';
  }
  return patch;
};
