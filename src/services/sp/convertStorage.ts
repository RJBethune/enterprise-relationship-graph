import { SpRest } from './SpRest';
import { StorageMode } from '../IGraphStore';
import { DocumentGraphStore } from './DocumentGraphStore';
import { ItemGraphStore } from './ItemGraphStore';
import { IGraph } from '../../model/bundle';
import { NODES_LIST, EDGES_LIST } from '../../provisioning/schema';
import { projectsPath, payloadPatch } from './projectItems';

/**
 * Move a project between storage models.
 *
 * Without this, choosing a storage model at creation was a decision nobody could revise:
 * a graph that outgrew a single list item had nowhere to go, and a graph that turned out
 * not to need per-item concurrency was stuck paying for it. The ceiling warning bought
 * time; this is the actual remedy.
 *
 * It works by READING through the model the project is in today and SAVING through the
 * one it is moving to — both of which are the ordinary, tested code paths, so conversion
 * has no bespoke write logic of its own to get subtly wrong.
 *
 * Snapshots are untouched: they are rows in ERG Snapshots under both models.
 *
 * SharePoint has no transactions, so this is deliberately ordered to be SAFE TO RETRY:
 * the new form is written and verified before the old form is cleared. An interruption
 * leaves a project whose data exists in both places, which the mode column resolves and
 * a second run tidies — never one whose data exists in neither.
 */

export interface IConversionResult {
  from: StorageMode;
  to: StorageMode;
  nodes: number;
  edges: number;
  snapshots: number;
}

/** Remove every node and relationship row belonging to a project. */
const deleteProjectRows = async (sp: SpRest, projectId: number): Promise<void> => {
  for (const listPath of [EDGES_LIST, NODES_LIST]) {
    // Edges first: an interruption then leaves nodes without their relationships, which
    // the orphan sweep tolerates, rather than relationships pointing at nothing.
    const path = `web/lists/getbytitle('${encodeURIComponent(listPath)}')`;
    const rows = await sp.getAll<{ Id: number }>(
      `${path}/items?$select=Id&$filter=ErgProjectId eq ${projectId}&$top=5000`
    );
    for (const row of rows) {
      await sp.del(`${path}/items(${row.Id})`);
    }
  }
};

export const convertProjectStorage = async (
  sp: SpRest,
  editorName: string,
  projectId: number,
  target: StorageMode,
  onProgress?: (message: string) => void
): Promise<IConversionResult> => {
  const say = (m: string): void => { if (onProgress) { onProgress(m); } };

  const documentStore = new DocumentGraphStore(sp, editorName);
  const itemStore = new ItemGraphStore(sp, editorName);

  const source = target === 'Items' ? documentStore : itemStore;
  const destination = target === 'Items' ? itemStore : documentStore;
  const from: StorageMode = target === 'Items' ? 'Document' : 'Items';

  say('Reading the current graph…');
  const before = await source.openProject(projectId);
  const graph: IGraph = before.bundle.graph;
  const snapshots = (before.bundle.snapshots || []) as unknown[];
  const nodes = (graph.nodes || []).length;
  const edges = (graph.edges || []).length;
  before.dispose();

  // Flip the mode FIRST so that if anything below fails, the project opens through the
  // model its data is about to live in rather than the one being emptied.
  say('Switching storage mode…');
  await sp.merge(`${projectsPath()}/items(${projectId})`, { ErgStorageMode: target }, '*');

  // VERIFY it landed before anything is cleared. A mode that silently fails to change
  // is the worst outcome available here: the rows get written, the old copy gets
  // cleared, and the project carries on reading the place the data just left. The data
  // is present in SharePoint and invisible in the app, which is indistinguishable from
  // data loss to everyone except the person reading the list.
  const confirmed = await sp.get<{ ErgStorageMode: string | null }>(
    `${projectsPath()}/items(${projectId})?$select=ErgStorageMode`
  );
  if (confirmed.ErgStorageMode !== target) {
    throw new Error(
      `SharePoint did not accept the storage mode change (it still reads ` +
      `"${confirmed.ErgStorageMode || 'empty'}"). Nothing has been changed. This usually ` +
      `means the ErgStorageMode column is missing the "${target}" choice — press Backend ` +
      `and Deploy, then try again.`
    );
  }

  say(`Writing ${nodes} node${nodes === 1 ? '' : 's'} and ${edges} relationship${edges === 1 ? '' : 's'}…`);
  const after = await destination.openProject(projectId);
  const outcome = await after.save(graph, snapshots);
  after.dispose();

  if (outcome.status === 'error') {
    // The mode is already flipped, but nothing was cleared, so the old data is intact
    // and a retry — or flipping the mode back — loses nothing.
    throw new Error(`Conversion could not write the graph: ${outcome.message || 'unknown error'}`);
  }

  say('Clearing the old copy…');
  if (target === 'Items') {
    // The payload columns are what Document mode read from; blanking them is what stops
    // two copies of the graph existing.
    await sp.merge(`${projectsPath()}/items(${projectId})`, payloadPatch(''), '*');
  } else {
    await deleteProjectRows(sp, projectId);
  }

  return { from, to: target, nodes, edges, snapshots: snapshots.length };
};

/**
 * Find a project whose data and storage mode disagree.
 *
 * An interrupted conversion — or one whose mode change was rejected — leaves rows in
 * the item lists while the project still claims to be a Document whose payload has been
 * cleared. The graph then opens EMPTY even though every node is sitting in SharePoint,
 * which reads as catastrophic data loss and is in fact a one-column fix.
 *
 * Returns the mode the data actually implies, or null when everything agrees.
 */
export const detectStorageMismatch = async (
  sp: SpRest, projectId: number
): Promise<StorageMode | null> => {
  try {
    const item = await sp.get<{ ErgStorageMode: string | null; [k: string]: unknown }>(
      `${projectsPath()}/items(${projectId})?$select=ErgStorageMode,ErgPayload1`
    );
    const mode: StorageMode = item.ErgStorageMode === 'Items' ? 'Items' : 'Document';
    const hasPayload = !!String(item.ErgPayload1 || '').trim();

    const nodeRows = await sp.getAll<{ Id: number }>(
      `web/lists/getbytitle('${encodeURIComponent(NODES_LIST)}')/items` +
      `?$select=Id&$filter=ErgProjectId eq ${projectId}&$top=1`
    );
    const hasRows = nodeRows.length > 0;

    // The only combination worth reporting: it says Document, its payload is gone, and
    // its nodes are sitting in the item list.
    if (mode === 'Document' && !hasPayload && hasRows) { return 'Items'; }
    return null;
  } catch {
    return null;
  }
};

/** Point a project at whichever storage its data is actually in. */
export const repairStorageMode = async (
  sp: SpRest, projectId: number, actual: StorageMode
): Promise<void> => {
  await sp.merge(`${projectsPath()}/items(${projectId})`, { ErgStorageMode: actual }, '*');
};
