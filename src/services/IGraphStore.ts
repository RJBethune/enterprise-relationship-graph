import { IGraph, IBundle } from '../model/bundle';
import { IMergeConflict } from '../model/merge';

/**
 * The storage contract.
 *
 * Two implementations satisfy it — the whole graph in one list item, and one item per
 * node and relationship — and the shell cannot tell them apart. That is the point:
 * the storage model is the part of this design most likely to change once real usage
 * arrives, so it sits behind an interface from day one. The engine never sees any of it.
 */

export type StorageMode = 'Document' | 'Items';

export interface IProjectSummary {
  id: number;
  title: string;
  storageMode: StorageMode;
  nodeCount: number;
  edgeCount: number;
  status: string;
  modified: string | null;
  modifiedBy: string | null;
}

export type SaveStatus = 'saved' | 'merged' | 'conflict' | 'error';

export interface ISaveOutcome {
  status: SaveStatus;
  /** Present when a remote change was merged in: the graph the engine should now show. */
  graph?: IGraph;
  conflicts?: IMergeConflict[];
  message?: string;
  /** Entities written on this save. Item mode reports real counts; document mode reports 1. */
  writes?: number;
}

export interface IRemoteUpdate {
  /** The graph as it now stands remotely, already merged with local edits where needed. */
  graph: IGraph;
  conflicts: IMergeConflict[];
  /** Who moved it, for the "updated by X" toast. */
  by: string | null;
}

/** A project opened for editing. Owns the session state a save needs: base graph, etag, cursor. */
export interface IOpenProject {
  summary: IProjectSummary;
  bundle: IBundle;
  /**
   * Persist the current graph. Implementations are responsible for conflict detection
   * and, where possible, resolving it by merge rather than by asking a human.
   */
  save(graph: IGraph, snapshots: unknown[] | null): Promise<ISaveOutcome>;
  /** Check for other people's changes. Returns null when nothing moved. */
  poll(localGraph: IGraph): Promise<IRemoteUpdate | null>;
  dispose(): void;
}

export interface IGraphStore {
  listProjects(): Promise<IProjectSummary[]>;
  createProject(title: string, mode: StorageMode): Promise<IProjectSummary>;
  renameProject(id: number, title: string): Promise<void>;
  archiveProject(id: number): Promise<void>;
  openProject(id: number): Promise<IOpenProject>;
}
