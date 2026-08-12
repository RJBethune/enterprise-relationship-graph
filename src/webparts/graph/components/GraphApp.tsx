import * as React from 'react';
import styles from './GraphApp.module.scss';
import { IGraphAppProps } from './IGraphAppProps';
import { SetupPanel } from './SetupPanel';
import { mountEngine, unmountEngine, isMounted } from '../../../engine/mount';
import { IErgHost, IEngineApi } from '../../../engine/hostContract';
import { IGraph, IBundle, normalizeGraph } from '../../../model/bundle';
import { IProjectSummary, IOpenProject, StorageMode } from '../../../services/IGraphStore';
import { IProvisioningPlan, isHealthy, planSummary } from '../../../provisioning/planner';
import { IProvisioningStepResult } from '../../../services/sp/SpProvisioningService';
import { IPresentUser, PresenceMode, HEARTBEAT_MS } from '../../../services/sp/PresenceService';
import { CORE_LISTS, PROJECTS_LIST } from '../../../provisioning/schema';
import { readListRights } from '../../../services/sp/permissions';

/** Never collapse smaller than this, however cramped the page section is. */
const MIN_SHELL_HEIGHT = 520;
/** Breathing room below the web part so the page does not gain a scrollbar. */
const SHELL_BOTTOM_GAP = 12;
/** How long after an edit somebody still counts as "editing" rather than "viewing". */
const EDITING_WINDOW_MS = 120000;

type Phase = 'checking' | 'setup' | 'loading' | 'ready' | 'error';
type SyncState = 'idle' | 'saving' | 'saved' | 'merged' | 'conflict' | 'error' | 'readonly';

interface IGraphAppState {
  phase: Phase;
  error: string | null;
  plan: IProvisioningPlan | null;
  provisioning: boolean;
  provisionProgress: { done: number; total: number; label: string } | null;
  provisionResults: IProvisioningStepResult[] | null;
  projects: IProjectSummary[];
  currentProjectId: number | null;
  sync: SyncState;
  syncDetail: string | null;
  switching: boolean;
  present: IPresentUser[];
  /** The declared schema and the live site disagree; the graph still works. */
  schemaGap: boolean;
  /** Write rights on ERG Projects itself. null = not determined; fall back to the web. */
  canEditList: boolean | null;
  /** Native Fullscreen API is active. */
  fullscreen: boolean;
  /** Fallback "cover the page" mode, for when the browser refuses fullscreen. */
  expanded: boolean;
}

/**
 * Everything between the SharePoint lists and the graph engine.
 *
 * Its whole job is to make saving invisible. The engine calls `persist()` on every
 * mutation exactly as it always did; this component debounces those calls into whole-
 * graph saves, resolves the conflicts that arise when two people edit at once, and
 * polls for other people's changes. Nothing about the graph's behaviour, appearance or
 * interaction model changes — the user just stops pressing a save button.
 */
export default class GraphApp extends React.Component<IGraphAppProps, IGraphAppState> {
  private engineHost: HTMLDivElement | null = null;
  private shellEl: HTMLDivElement | null = null;
  private engine: IEngineApi | null = null;
  private session: IOpenProject | null = null;
  private saveTimer: ReturnType<typeof setTimeout> | null = null;
  private pollTimer: ReturnType<typeof setInterval> | null = null;
  private presenceTimer: ReturnType<typeof setInterval> | null = null;
  private resizeObserver: { disconnect(): void } | null = null;
  private lastEditAt: number = 0;
  private pendingGraph: IGraph | null = null;
  private snapshots: unknown[] = [];
  private savingNow: boolean = false;
  private disposed: boolean = false;
  private readonly onBeforeUnload: () => void;
  private readonly onVisibility: () => void;

  public constructor(props: IGraphAppProps) {
    super(props);
    this.state = {
      phase: 'checking',
      error: null,
      plan: null,
      provisioning: false,
      provisionProgress: null,
      provisionResults: null,
      projects: [],
      currentProjectId: null,
      sync: props.canEdit ? 'idle' : 'readonly',
      syncDetail: null,
      switching: false,
      present: [],
      schemaGap: false,
      canEditList: null,
      fullscreen: false,
      expanded: false
    };
    this.onBeforeUnload = (): void => { void this.flushPendingSave(); };
    this.onVisibility = (): void => {
      if (document.visibilityState === 'hidden') { void this.flushPendingSave(); }
    };
  }

  public componentDidMount(): void {
    window.addEventListener('beforeunload', this.onBeforeUnload);
    document.addEventListener('visibilitychange', this.onVisibility);
    window.addEventListener('resize', this.fitToViewport);
    document.addEventListener('fullscreenchange', this.onFullscreenChange);
    document.addEventListener('keydown', this.onKeyDown);
    this.watchContainerSize();
    this.fitToViewport();
    void this.initialize();
  }

  public componentWillUnmount(): void {
    this.disposed = true;
    window.removeEventListener('beforeunload', this.onBeforeUnload);
    document.removeEventListener('visibilitychange', this.onVisibility);
    window.removeEventListener('resize', this.fitToViewport);
    document.removeEventListener('fullscreenchange', this.onFullscreenChange);
    document.removeEventListener('keydown', this.onKeyDown);
    if (this.resizeObserver) { this.resizeObserver.disconnect(); }
    if (this.saveTimer) { clearTimeout(this.saveTimer); }
    if (this.pollTimer) { clearInterval(this.pollTimer); }
    if (this.presenceTimer) { clearInterval(this.presenceTimer); }
    if (this.state.currentProjectId) { void this.props.services.presence.leave(this.state.currentProjectId); }
    if (this.engine) { this.engine.destroy(); }
    if (this.session) { this.session.dispose(); }
    if (this.engineHost) { unmountEngine(this.engineHost); }
  }

  /* ------------------------------------------------------------------- sizing */

  /**
   * Give the graph the height that is actually left on the page.
   *
   * The engine was a whole-page application and sizes itself in viewport units; the
   * host stylesheet re-points it at its container, and this decides how tall that
   * container is. Measuring from the element's own top means it works on any page
   * layout — app page, section, or a page with a tall header — without hard-coding
   * what SharePoint puts above it.
   */
  private fitToViewport = (): void => {
    const el = this.shellEl;
    if (!el) { return; }

    // Filling the screen is the whole point of both expanded modes; there is no page
    // chrome left to measure around.
    if (this.state.fullscreen || this.state.expanded) {
      if (el.style.height !== '100vh') {
        el.style.height = '100vh';
        if (this.engine) { this.engine.resize(); }
      }
      return;
    }

    const top = el.getBoundingClientRect().top;
    const height = Math.max(
      MIN_SHELL_HEIGHT,
      Math.round(window.innerHeight - top - SHELL_BOTTOM_GAP)
    );
    if (el.style.height === `${height}px`) { return; }
    el.style.height = `${height}px`;
    // The engine only re-measures its canvas on WINDOW resize, which does not fire
    // when a web part resizes itself. Without this the canvas keeps stale dimensions
    // and the drawing is stretched.
    if (this.engine) { this.engine.resize(); }
  };

  /** SharePoint lays sections out after mount, so watch the container, not just the window. */
  private watchContainerSize(): void {
    if (this.resizeObserver) { this.resizeObserver.disconnect(); this.resizeObserver = null; }
    const Observer = (window as unknown as {
      ResizeObserver?: new (cb: () => void) => { observe(t: Element): void; disconnect(): void };
    }).ResizeObserver;
    if (!Observer || !this.shellEl || !this.shellEl.parentElement) { return; }
    const observer = new Observer(() => { this.fitToViewport(); });
    observer.observe(this.shellEl.parentElement);
    this.resizeObserver = observer;
  }

  /* --------------------------------------------------------------- fullscreen */

  /**
   * Native fullscreen where the browser allows it, a page-covering overlay where it
   * does not. SharePoint pages can sit in contexts that refuse the Fullscreen API, and
   * "expand the graph" is too useful to be at the mercy of that.
   */
  private toggleFullscreen = (): void => {
    const el = this.shellEl;
    if (!el) { return; }

    if (this.state.fullscreen) {
      void Promise.resolve(document.exitFullscreen && document.exitFullscreen()).catch(() => undefined);
      return;
    }
    if (this.state.expanded) {
      this.setState({ expanded: false }, this.fitToViewport);
      return;
    }

    if (!el.requestFullscreen) {
      this.setState({ expanded: true }, this.fitToViewport);
      return;
    }
    el.requestFullscreen().catch(() => {
      this.setState({ expanded: true }, this.fitToViewport);
    });
  };

  private onFullscreenChange = (): void => {
    const active = document.fullscreenElement === this.shellEl;
    this.setState({ fullscreen: active }, () => {
      // The browser needs a frame to apply the new box before it can be measured.
      window.setTimeout(this.fitToViewport, 50);
    });
  };

  /** Escape exits the overlay fallback, matching what Escape does in real fullscreen. */
  private onKeyDown = (event: KeyboardEvent): void => {
    if (event.key === 'Escape' && this.state.expanded) {
      this.setState({ expanded: false }, this.fitToViewport);
    }
  };

  private setShellRef = (el: HTMLDivElement | null): void => {
    if (el === this.shellEl) { return; }
    this.shellEl = el;
    if (!el) { return; }
    this.fitToViewport();
    this.watchContainerSize();
  };

  /* ------------------------------------------------------------------ startup */

  private async initialize(): Promise<void> {
    try {
      // A verified-healthy verdict for THIS build's schema on THIS site means the
      // fifteen reads have nothing to discover. The verdict is keyed by a fingerprint
      // of the schema in the bundle, so an upgraded .sppkg never matches an old one.
      if (this.props.services.provisioning.wasVerifiedHealthy()) {
        await this.loadProjects();
        return;
      }
      const plan = await this.props.services.provisioning.buildPlan();
      // Only a MISSING CORE LIST forces the setup screen. A site that is merely a
      // column behind should still open its graph — the setup panel stays one click
      // away rather than standing between people and their data.
      const missingCore = plan.actions.some(
        (a) => a.kind === 'createList' && CORE_LISTS.indexOf(a.list) >= 0
      );
      this.setState({ plan });
      if (missingCore) { this.setState({ phase: 'setup' }); return; }
      await this.loadProjects();

      // A gap that is not in a core list must not block the graph — but staying silent
      // about it is how a site ends up with features quietly missing and nobody
      // knowing why. Presence disappearing because ERG Presence was never created is
      // exactly that failure. Say it once, plainly, and keep the button marked.
      if (!isHealthy(plan)) {
        this.setState({ schemaGap: true });
        if (this.engine) {
          this.engine.toast(
            `SharePoint setup is incomplete — ${planSummary(plan)} Open Backend to finish it.`,
            'err'
          );
        }
      }
    } catch (e) {
      this.fail(e);
    }
  }

  /** The web-level answer, narrowed by the list's own permissions once we know them. */
  private get canEdit(): boolean {
    return this.state.canEditList === null ? this.props.canEdit : this.state.canEditList;
  }

  private async loadProjects(): Promise<void> {
    this.setState({ phase: 'loading' });
    const rights = await readListRights(this.props.services.sp, PROJECTS_LIST);
    if (rights) { this.setState({ canEditList: rights.canEdit }); }
    const store = this.props.services.storeFor('Document');
    const projects = await store.listProjects();

    if (projects.length === 0) {
      if (!this.canEdit) {
        this.setState({
          phase: 'error',
          error: 'No graphs have been created on this site yet, and you do not have permission to create one.'
        });
        return;
      }
      const created = await store.createProject('Enterprise Relationship Graph', 'Document');
      await this.openProject(created.id, [created]);
      return;
    }

    const requested = this.props.initialProjectId;
    const exists = projects.filter((p) => p.id === requested)[0];
    await this.openProject(exists ? exists.id : projects[0].id, projects);
  }

  private async openProject(projectId: number, projects: IProjectSummary[]): Promise<void> {
    const summary = projects.filter((p) => p.id === projectId)[0];
    if (!summary) { throw new Error(`Project ${projectId} is not in the list.`); }

    const store = this.props.services.storeFor(summary.storageMode);
    const session = await store.openProject(projectId);
    this.session = session;
    this.snapshots = (session.bundle.snapshots as unknown[]) || [];

    this.setState(
      { phase: 'ready', projects, currentProjectId: projectId, error: null },
      () => { this.afterProjectOpen(session); }
    );
    this.props.onProjectChanged(projectId);
  }

  private afterProjectOpen(session: IOpenProject): void {
    if (this.engine) {
      // Engine already running: swap the data, keep the canvas.
      this.engine.setBundle(session.bundle, session.summary.title);
      this.engine.fit();
    } else {
      this.mountEngineNow(session.bundle, session.summary.title);
    }
    this.fitToViewport();
    this.startPolling();
    this.startPresence();
  }

  /* ----------------------------------------------------------------- presence */

  private currentPresenceMode(): PresenceMode {
    return Date.now() - this.lastEditAt < EDITING_WINDOW_MS ? 'Editing' : 'Viewing';
  }

  private startPresence(): void {
    if (this.presenceTimer) { clearInterval(this.presenceTimer); this.presenceTimer = null; }
    void this.beatAndRefresh();
    this.presenceTimer = setInterval(() => { void this.beatAndRefresh(); }, HEARTBEAT_MS);
  }

  private async beatAndRefresh(): Promise<void> {
    const projectId = this.state.currentProjectId;
    if (!projectId || this.disposed) { return; }
    // A backgrounded tab keeps its row alive but stops polling for others — it is
    // still "present", just not worth spending reads on.
    const { presence } = this.props.services;
    await presence.heartbeat(projectId, this.currentPresenceMode());

    // Presence switching itself off means its list is not there. That is a schema gap
    // the operator can fix in one click, so surface it instead of just showing nobody.
    if (presence.isDisabled) {
      if (this.presenceTimer) { clearInterval(this.presenceTimer); this.presenceTimer = null; }
      if (!this.disposed && !this.state.schemaGap) {
        this.setState({ schemaGap: true });
        if (this.engine) {
          this.engine.toast(
            'Who-is-here needs the ERG Presence list, which this site does not have yet. ' +
            'Open Backend and press Deploy to add it.',
            'err'
          );
        }
      }
      return;
    }

    if (document.visibilityState === 'hidden') { return; }
    const present = await presence.list(projectId);
    if (!this.disposed) { this.setState({ present }); }
  }

  private mountEngineNow(bundle: IBundle, label: string): void {
    if (!this.engineHost) { return; }
    if (isMounted()) {
      this.setState({
        phase: 'error',
        error: 'The graph is already open elsewhere on this page. Add it to a single-part app page, ' +
          'or remove the duplicate web part — two copies would fight over the same page elements.'
      });
      return;
    }

    const host: IErgHost = {
      mode: 'sharepoint',
      editorName: this.props.editorName,
      initialGraph: normalizeGraph(bundle.graph),
      initialSnapshots: this.snapshots,
      onGraphChanged: (graph: IGraph): void => { this.scheduleSave(graph); },
      onSnapshotsChanged: (snaps: unknown[]): void => {
        this.snapshots = snaps;
        if (this.pendingGraph) { this.scheduleSave(this.pendingGraph); }
      },
      onGraphReplaced: (graph: IGraph, sourceName: string | null): void => {
        void this.importIntoProject(graph, sourceName);
      },
      renderStatusChip: (chip, nameEl): void => { this.paintChip(chip, nameEl); },
      onReady: (api: IEngineApi): void => {
        this.engine = api;
        api.setProjectLabel(label);
      }
    };

    mountEngine({ container: this.engineHost, host, assetBaseUrl: this.props.assetBaseUrl });
  }

  /* ------------------------------------------------------------------ import */

  /**
   * A file the user opened becomes this project's stored graph.
   *
   * The engine treats a file load as "now clean", which is true of the file and false
   * of SharePoint — so this has to push it, not wait for the next edit. It is also
   * destructive: it replaces whatever the project held. Version history makes that
   * recoverable, but recoverable is not the same as expected, so an import over a
   * populated project asks first, and declining restores what was there.
   */
  private async importIntoProject(graph: IGraph, sourceName: string | null): Promise<void> {
    if (!this.session || !this.engine) { return; }

    if (!this.canEdit) {
      this.engine.toast('Loaded for viewing only — you do not have permission to save to this graph.', 'err');
      return;
    }

    const existing = this.session.summary;
    const hadContent = (this.state.projects.filter((p) => p.id === existing.id)[0] || existing).nodeCount > 0;
    const incoming = (graph.nodes || []).length;
    const label = sourceName ? `"${sourceName}"` : 'that file';

    if (hadContent) {
      const ok = window.confirm(
        `Replace the graph "${existing.title}" with ${label}?\n\n` +
        `${incoming} node${incoming === 1 ? '' : 's'} will be saved to SharePoint, replacing what is ` +
        'stored now.\n\nThe previous version stays in this list item\'s version history, but the ' +
        'graph everyone else sees will change.'
      );
      if (!ok) {
        // Put back exactly what SharePoint holds rather than leaving the import on screen.
        const reopened = await this.props.services
          .storeFor(existing.storageMode).openProject(existing.id);
        this.session.dispose();
        this.session = reopened;
        this.engine.setBundle(reopened.bundle, existing.title);
        this.engine.toast('Import cancelled — the stored graph is back.', 'ok');
        return;
      }
    }

    this.engine.setProjectLabel(existing.title);
    this.engine.markDirty();
    this.pendingGraph = graph;
    this.lastEditAt = Date.now();
    this.setSync('saving', `Importing ${label}`);
    await this.flushPendingSave();

    if (!this.disposed && this.engine && this.state.sync !== 'error') {
      this.engine.toast(
        `Imported ${incoming} node${incoming === 1 ? '' : 's'} into "${existing.title}" and saved to SharePoint.`,
        'ok'
      );
    }
  }

  /* --------------------------------------------------------------- autosaving */

  private scheduleSave(graph: IGraph): void {
    if (!this.canEdit) { return; }
    // Any mutation marks us as editing rather than merely viewing, for presence.
    this.lastEditAt = Date.now();
    this.pendingGraph = graph;
    if (this.saveTimer) { clearTimeout(this.saveTimer); }
    this.setSync('saving', null);
    this.saveTimer = setTimeout(() => { void this.doSave(); }, this.props.autosaveMs);
  }

  private async flushPendingSave(): Promise<void> {
    if (this.saveTimer) { clearTimeout(this.saveTimer); this.saveTimer = null; }
    if (this.pendingGraph) { await this.doSave(); }
  }

  private async doSave(): Promise<void> {
    if (!this.session || !this.pendingGraph || this.savingNow) { return; }
    // A save already in flight will not include the newest edit; leave the pending
    // graph in place so the next tick writes it rather than dropping it.
    this.savingNow = true;
    const graph = this.pendingGraph;
    this.pendingGraph = null;

    try {
      const outcome = await this.session.save(graph, this.snapshots);
      if (this.disposed) { return; }

      switch (outcome.status) {
        case 'saved':
          this.setSync('saved', null);
          this.refreshCounts(graph);
          if (this.engine) { this.engine.markClean(); }
          break;
        case 'merged':
          if (this.engine && outcome.graph) {
            this.engine.setBundle({ version: 4, graph: outcome.graph, snapshots: this.snapshots });
            this.engine.markClean();
            this.engine.toast(
              outcome.conflicts && outcome.conflicts.length
                ? `Merged with another editor — ${outcome.conflicts.length} item(s) needed a decision, yours were kept.`
                : (outcome.message || 'Merged changes from another editor.'),
              'ok'
            );
          }
          this.setSync('merged', outcome.message || null);
          break;
        case 'conflict':
          this.setSync('conflict', outcome.message || 'Someone else is saving right now.');
          if (this.engine && outcome.graph) {
            this.engine.setBundle({ version: 4, graph: outcome.graph, snapshots: this.snapshots });
            this.engine.markDirty();
          }
          // Keep the work queued so the next quiet moment retries it.
          this.pendingGraph = outcome.graph || graph;
          break;
        default:
          this.setSync('error', outcome.message || 'Save failed.');
          this.pendingGraph = graph;
          break;
      }
    } catch (e) {
      if (this.disposed) { return; }
      this.setSync('error', e instanceof Error ? e.message : String(e));
      this.pendingGraph = graph;
    } finally {
      this.savingNow = false;
      if (this.engine) { this.engine.refresh(); }
    }
  }

  /**
   * Keep the switcher's counts honest.
   *
   * They are loaded from the project item, so without this the label still reads
   * "0 nodes" while the graph plainly shows several — which reads as "my work is not
   * being saved" precisely when it is.
   */
  private refreshCounts(graph: IGraph): void {
    const id = this.state.currentProjectId;
    if (!id) { return; }
    const nodeCount = (graph.nodes || []).length;
    const edgeCount = (graph.edges || []).length;
    this.setState((prev) => ({
      projects: prev.projects.map(
        (p) => (p.id === id ? { ...p, nodeCount, edgeCount } : p)
      )
    }));
  }

  /* ------------------------------------------------------------------ polling */

  private startPolling(): void {
    if (this.pollTimer) { clearInterval(this.pollTimer); this.pollTimer = null; }
    if (this.props.pollSeconds <= 0) { return; }
    this.pollTimer = setInterval(() => { void this.pollOnce(); }, this.props.pollSeconds * 1000);
  }

  private async pollOnce(): Promise<void> {
    if (!this.session || !this.engine || this.savingNow || this.state.switching) { return; }
    // Polling a tab nobody is looking at costs requests and buys nothing.
    if (document.visibilityState === 'hidden') { return; }

    try {
      const update = await this.session.poll(this.engine.getGraph());
      if (!update || this.disposed) { return; }

      this.engine.setBundle({ version: 4, graph: update.graph, snapshots: this.snapshots });
      const who = update.by ? ` by ${update.by}` : '';
      this.engine.toast(
        update.conflicts.length
          ? `Graph updated${who}. ${update.conflicts.length} item(s) you were editing changed — yours were kept.`
          : `Graph updated${who}.`,
        'ok'
      );
      if (update.conflicts.length) { this.setSync('merged', `${update.conflicts.length} merged edit(s)`); }
    } catch {
      // A failed poll is not worth interrupting anyone over; the next tick retries.
    }
  }

  /* ----------------------------------------------------------------- switching */

  private switchProject = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    const id = Number(event.target.value);
    if (!id || id === this.state.currentProjectId) { return; }
    void this.performSwitch(id);
  };

  private async performSwitch(id: number): Promise<void> {
    this.setState({ switching: true });
    try {
      // Never carry unsaved work across a switch.
      await this.flushPendingSave();
      const leaving = this.state.currentProjectId;
      if (leaving) { void this.props.services.presence.leave(leaving); }
      this.setState({ present: [] });
      if (this.session) { this.session.dispose(); this.session = null; }
      await this.openProject(id, this.state.projects);
    } catch (e) {
      this.fail(e);
    } finally {
      if (!this.disposed) { this.setState({ switching: false }); }
    }
  }

  private newProject = (): void => {
    const title = window.prompt('Name for the new graph:', 'New relationship graph');
    if (!title || !title.trim()) { return; }
    const mode: StorageMode = window.confirm(
      'Use per-item storage for this graph?\n\n' +
      'Cancel — Document (recommended). The whole graph lives in one list item, so every ' +
      'save is a single atomic write with full version history.\n\n' +
      'OK — Per-item (EXPERIMENTAL). One row per node and relationship: several people can ' +
      'edit at once and see each other\'s changes, and there is no size ceiling. This path ' +
      'has not yet been exercised against a production tenant, and there is no way to ' +
      'convert a project between the two afterwards.'
    ) ? 'Items' : 'Document';

    void (async (): Promise<void> => {
      try {
        await this.flushPendingSave();
        const store = this.props.services.storeFor(mode);
        const created = await store.createProject(title.trim(), mode);
        if (this.session) { this.session.dispose(); this.session = null; }
        await this.openProject(created.id, this.state.projects.concat([created]));
      } catch (e) {
        this.fail(e);
      }
    })();
  };

  /* -------------------------------------------------------------- provisioning */

  private deploy = (): void => {
    void (async (): Promise<void> => {
      const plan = this.state.plan;
      if (!plan) { return; }
      this.setState({ provisioning: true, provisionResults: null, provisionProgress: { done: 0, total: plan.actions.length, label: '' } });
      try {
        const results = await this.props.services.provisioning.execute(plan, (done, total, label) => {
          if (!this.disposed) { this.setState({ provisionProgress: { done, total, label } }); }
        });
        const fresh = await this.props.services.provisioning.buildPlan();
        this.setState({ provisioning: false, provisionResults: results, plan: fresh, schemaGap: !isHealthy(fresh) });
        // Anything that switched itself off because its list was missing gets another
        // go now, rather than staying dead until someone reloads the page.
        this.props.services.presence.reset();
        if (this.state.currentProjectId) { this.startPresence(); }
      } catch (e) {
        this.setState({ provisioning: false });
        this.fail(e);
      }
    })();
  };

  private recheck = (): void => {
    void (async (): Promise<void> => {
      try {
        const plan = await this.props.services.provisioning.buildPlan();
        this.setState({ plan, provisionResults: null, schemaGap: !isHealthy(plan) });
      } catch (e) { this.fail(e); }
    })();
  };

  private continueToGraph = (): void => { void this.loadProjects(); };

  private openSetup = (): void => {
    this.setState({ phase: 'setup', provisionResults: null });
    this.recheck();
  };

  /* --------------------------------------------------------------------- misc */

  private setSync(sync: SyncState, syncDetail: string | null): void {
    if (this.disposed) { return; }
    this.setState({ sync, syncDetail });
    if (this.engine) { this.engine.refresh(); }
  }

  private fail(e: unknown): void {
    if (this.disposed) { return; }
    this.setState({ phase: 'error', error: e instanceof Error ? e.message : String(e) });
  }

  /** Repaints the engine's own header chip so save state reads where the eye already looks. */
  private paintChip(chip: HTMLElement, nameEl: HTMLElement): void {
    const label: { [k in SyncState]: string } = {
      idle: 'All changes saved',
      saving: 'Saving…',
      saved: 'All changes saved',
      merged: 'Merged with another editor',
      conflict: 'Save queued — retrying',
      error: 'Save failed',
      readonly: 'Read-only'
    };
    const text = label[this.state.sync];
    nameEl.textContent = text;
    nameEl.classList.remove('no-file');
    chip.classList.toggle('dirty', this.state.sync === 'error' || this.state.sync === 'conflict');
    chip.setAttribute('aria-label', text + (this.state.syncDetail ? '. ' + this.state.syncDetail : ''));
    chip.title = this.state.syncDetail || text;
  }

  /**
   * Who else is on this graph.
   *
   * Deliberately shows YOU as well: seeing your own initials is how you know the
   * indicator is live, so an empty strip reads as "nobody else here" rather than
   * "presence is broken".
   */
  private renderPresence(): JSX.Element | null {
    const people = this.state.present;
    if (people.length === 0) { return null; }

    const initials = (name: string): string => name
      .split(/[\s,]+/).filter(Boolean).slice(0, 2)
      .map((part) => part.charAt(0).toUpperCase()).join('') || '?';

    const others = people.filter((p) => !p.isSelf).length;
    const shown = people.slice(0, 6);
    const summary = people
      .map((p) => `${p.name}${p.isSelf ? ' (you)' : ''} — ${p.mode.toLowerCase()}`)
      .join(', ');

    // SharePoint's own photo endpoint: same origin, already authenticated, and it is
    // not a library file, so it is unaffected by the download policy. A missing photo
    // 404s or returns a placeholder, which the onError handler turns back into initials.
    const photoUrl = (email: string): string =>
      `${this.props.services.sp.webUrl}/_layouts/15/userphoto.aspx` +
      `?size=M&accountname=${encodeURIComponent(email)}`;

    const ago = (iso: string): string => {
      const seconds = Math.max(0, Math.round((Date.now() - Date.parse(iso)) / 1000));
      if (seconds < 45) { return 'just now'; }
      if (seconds < 90) { return 'a minute ago'; }
      return `${Math.round(seconds / 60)} minutes ago`;
    };

    return (
      <span className={styles.presence}>
        {shown.map((p) => (
          <span
            key={p.login || p.name}
            className={`${styles.person} ${p.mode === 'Editing' ? styles.editing : ''} ${p.isSelf ? styles.self : ''}`}
            tabIndex={0}
            role="button"
            aria-label={`${p.name}${p.isSelf ? ', you' : ''}, ${p.mode.toLowerCase()}, last seen ${ago(p.lastSeen)}`}
          >
            <span className={styles.initials} aria-hidden="true">{initials(p.name)}</span>
            {p.email && (
              <img
                className={styles.photo}
                src={photoUrl(p.email)}
                alt=""
                aria-hidden="true"
                onError={(e): void => { (e.target as HTMLImageElement).style.display = 'none'; }}
              />
            )}
            <span className={styles.card} role="tooltip">
              <span className={styles.cardTop}>
                {p.email && (
                  <img
                    className={styles.cardPhoto}
                    src={photoUrl(p.email)}
                    alt=""
                    onError={(e): void => { (e.target as HTMLImageElement).style.display = 'none'; }}
                  />
                )}
                <span className={styles.cardWho}>
                  <strong>{p.name}{p.isSelf ? ' (you)' : ''}</strong>
                  {p.email && <span className={styles.cardMeta}>{p.email}</span>}
                </span>
              </span>
              <span className={styles.cardMeta}>
                {p.mode === 'Editing' ? 'Editing this graph' : 'Viewing this graph'} · seen {ago(p.lastSeen)}
              </span>
            </span>
          </span>
        ))}
        {people.length > shown.length && (
          <span className={styles.person} aria-hidden="true">
            <span className={styles.initials}>+{people.length - shown.length}</span>
          </span>
        )}
        <span className={styles.srOnly}>
          {others === 0
            ? 'No one else is on this graph.'
            : `${others} other ${others === 1 ? 'person' : 'people'} on this graph: ${summary}`}
        </span>
      </span>
    );
  }

  private renderBadge(): JSX.Element {
    const tone =
      this.state.sync === 'error' || this.state.sync === 'conflict' ? styles.bad
        : this.state.sync === 'saving' ? styles.busy
          : styles.ok;
    const text: { [k in SyncState]: string } = {
      idle: 'Saved', saving: 'Saving…', saved: 'Saved', merged: 'Merged',
      conflict: 'Retrying', error: 'Save failed', readonly: 'Read-only'
    };
    return (
      <span className={styles.badge} title={this.state.syncDetail || undefined}>
        <span className={`${styles.dot} ${tone}`} aria-hidden="true" />
        <span>{text[this.state.sync]}</span>
      </span>
    );
  }

  public render(): React.ReactElement {
    const { phase, plan, projects, currentProjectId } = this.state;

    if (phase === 'checking' || phase === 'loading') {
      return (
        <div
        className={`${styles.shell} ${this.state.expanded ? styles.expanded : ''}`}
        ref={this.setShellRef}
      >
          <div className={styles.panel} role="status" aria-live="polite">
            <h2 className={styles.panelTitle}>Loading the graph…</h2>
            <p className={styles.panelText}>
              {phase === 'checking' ? 'Checking the SharePoint backend.' : 'Reading the current graph.'}
            </p>
          </div>
        </div>
      );
    }

    if (phase === 'error') {
      return (
        <div
        className={`${styles.shell} ${this.state.expanded ? styles.expanded : ''}`}
        ref={this.setShellRef}
      >
          <div className={styles.panel} role="alert">
            <h2 className={styles.panelTitle}>The graph could not be opened</h2>
            <p className={styles.panelText}>{this.state.error}</p>
            <button type="button" className={styles.button} onClick={this.openSetup}>
              Check the SharePoint backend
            </button>
          </div>
        </div>
      );
    }

    if (phase === 'setup' && plan) {
      return (
        <div
        className={`${styles.shell} ${this.state.expanded ? styles.expanded : ''}`}
        ref={this.setShellRef}
      >
          <SetupPanel
            plan={plan}
            canEdit={this.props.canEdit}
            running={this.state.provisioning}
            progress={this.state.provisionProgress}
            results={this.state.provisionResults}
            onDeploy={this.deploy}
            onRecheck={this.recheck}
            onContinue={this.continueToGraph}
          />
        </div>
      );
    }

    const current = projects.filter((p) => p.id === currentProjectId)[0];

    return (
      <div
        className={`${styles.shell} ${this.state.expanded ? styles.expanded : ''}`}
        ref={this.setShellRef}
      >
        <div className={styles.bar}>
          <label className={styles.barLabel} htmlFor="erg-project-select">Graph</label>
          <select
            id="erg-project-select"
            className={styles.select}
            value={currentProjectId || ''}
            onChange={this.switchProject}
            disabled={this.state.switching}
          >
            {projects.map((p) => (
              <option key={p.id} value={p.id}>
                {p.title} — {p.nodeCount} nodes, {p.edgeCount} relations
              </option>
            ))}
          </select>

          {this.canEdit && (
            <button type="button" className={styles.button} onClick={this.newProject} disabled={this.state.switching}>
              New graph
            </button>
          )}

          <span className={styles.spacer} />

          {this.renderPresence()}

          {current && (
            <span
              className={styles.badge}
              title={current.storageMode === 'Items'
                ? 'Every node and relationship is its own row in the ERG Nodes and ERG Edges lists.'
                : 'The whole graph is stored in the payload columns of this project\'s ERG Projects ' +
                  'item, saved in one atomic write. The ERG Nodes and ERG Edges lists stay empty for ' +
                  'this graph — that is by design, not a failed save.'}
            >
              {current.storageMode === 'Items' ? 'Per-item · experimental' : 'Document · one list item'}
            </span>
          )}
          {this.renderBadge()}
          <button
            type="button"
            className={`${styles.button} ${this.state.schemaGap ? styles.warn : ''}`}
            onClick={this.openSetup}
            title={this.state.schemaGap && this.state.plan
              ? `SharePoint setup is incomplete. ${planSummary(this.state.plan)}`
              : 'Check or update the SharePoint lists behind this graph'}
          >
            {this.state.schemaGap ? '⚠ Backend' : 'Backend'}
          </button>
          <button
            type="button"
            className={styles.button}
            onClick={this.toggleFullscreen}
            aria-pressed={this.state.fullscreen || this.state.expanded}
            title={this.state.fullscreen || this.state.expanded
              ? 'Exit full screen (Esc)'
              : 'Work in full screen'}
          >
            {this.state.fullscreen || this.state.expanded ? '⤡ Exit full screen' : '⤢ Full screen'}
          </button>
        </div>

        <div className={styles.canvas} ref={(el): void => { this.engineHost = el; }} />

        <span className={styles.srOnly} role="status" aria-live="polite">
          {this.state.syncDetail || ''}
        </span>
      </div>
    );
  }
}
