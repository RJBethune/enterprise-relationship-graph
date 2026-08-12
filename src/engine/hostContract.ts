import { IGraph, IBundle } from '../model/bundle';

/**
 * The contract between the extracted graph engine and whatever is hosting it.
 *
 * The engine is the v1.x application, unchanged. It knows nothing about SharePoint,
 * React, or projects — it only knows that when the graph changes it calls
 * `onGraphChanged`, and that at the end of boot it hands back an `IEngineApi`.
 * Everything the SPFx shell does (project switching, autosave, remote merges) it
 * does through that API. That boundary is what lets the storage model change
 * underneath — document blob today, per-item write-through tomorrow — without the
 * engine noticing.
 */
export interface IErgHost {
  /**
   * 'sharepoint' suppresses the file-oriented behaviours that make no sense when a
   * list is the source of truth: reopening the last local file, and the first-run
   * "open a JSON file to start" hint.
   */
  mode: 'sharepoint' | 'standalone';

  /** Stamped on saves so teammates see who edited last. The signed-in user under SPFx. */
  editorName: string;

  /** Graph to boot with. null falls back to the engine's built-in sample data. */
  initialGraph: IGraph | null;

  /** Snapshots to boot with. Mutated in place by the engine's snapshot commands. */
  initialSnapshots: unknown[] | null;

  /** Fires on every engine mutation (the engine's `persist()` seam). */
  onGraphChanged(graph: IGraph): void;

  /** Fires when the snapshot list changes. */
  onSnapshotsChanged(snapshots: unknown[]): void;

  /** Lets the shell own the header chip's text — "Saved", "Saving…", "Conflict". */
  renderStatusChip?(chip: HTMLElement, nameEl: HTMLElement, state: { dirty: boolean; fileName: string | null }): void;

  /** Called once at the end of boot with the control surface for the running engine. */
  onReady(api: IEngineApi): void;
}

/** The running engine, as driven by the shell. */
export interface IEngineApi {
  /** Replace the whole graph — project switch, snapshot restore, remote reload. */
  setBundle(bundle: IBundle, label?: string | null): void;
  getBundle(): IBundle;
  getGraph(): IGraph;
  setProjectLabel(label: string | null): void;
  isDirty(): boolean;
  markClean(): void;
  markDirty(): void;
  toast(message: string, kind?: 'ok' | 'err'): void;
  refresh(): void;
  fit(): void;
  /** Re-measure the canvas after the host resizes its container. */
  resize(): void;
  /** Stops the render loop and removes document/window listeners (web part dispose). */
  destroy(): void;
}
