/**
 * The graph document — the shape the engine reads and writes, and the shape that
 * round-trips through both storage models.
 *
 * Node and edge attributes are deliberately open (`[key: string]: unknown`). The
 * tool lets each office define custom node types at runtime, so the set of
 * attributes is not knowable at build time. Everything the storage layer needs to
 * query — id, label, type, endpoints — is named; the rest travels as data.
 */

export interface IGraphNode {
  id: string;
  label: string;
  type: string;
  [key: string]: unknown;
}

export interface IGraphEdge {
  id: string;
  source: string;
  target: string;
  type: string;
  [key: string]: unknown;
}

export interface IPosition {
  x: number;
  y: number;
}

export interface IGraph {
  nodes: IGraphNode[];
  edges: IGraphEdge[];
  customNodeTypes?: unknown[];
  collapsedNodes?: string[];
  /** Hand-arranged layout, keyed by node id. Presentation, not meaning. */
  positions?: { [nodeId: string]: IPosition };
  [key: string]: unknown;
}

export interface IBundle {
  version: number;
  exportedAt?: string;
  lastModifiedBy?: string | null;
  lastModifiedAt?: string | null;
  graph: IGraph;
  snapshots?: unknown[] | null;
}

/** Matches BUNDLE_SCHEMA_VERSION in the engine. Bump both together. */
export const BUNDLE_SCHEMA_VERSION: number = 4;

export const emptyGraph = (): IGraph => ({
  nodes: [],
  edges: [],
  customNodeTypes: [],
  collapsedNodes: [],
  positions: {}
});

/**
 * Deterministic JSON: object keys sorted at every level.
 *
 * Every change-detection path in this app compares serialized forms — is this node
 * different from the stored one, did the graph change since the last save. JSON.stringify
 * preserves insertion order, so a node rebuilt with its keys in a different order
 * serializes differently while meaning exactly the same thing. That produces phantom
 * writes on every save and phantom conflicts on every merge. Sorting removes the
 * whole class of bug.
 */
export const stableStringify = (value: unknown): string => {
  if (value === null || typeof value !== 'object') { return JSON.stringify(value) ?? 'null'; }
  if (Array.isArray(value)) {
    return '[' + value.map((v) => stableStringify(v)).join(',') + ']';
  }
  const obj = value as { [k: string]: unknown };
  const keys = Object.keys(obj).sort();
  const parts: string[] = [];
  for (const k of keys) {
    if (obj[k] === undefined) { continue; }
    parts.push(JSON.stringify(k) + ':' + stableStringify(obj[k]));
  }
  return '{' + parts.join(',') + '}';
};

/** True when two graph entities are meaningfully identical. */
export const sameEntity = (a: unknown, b: unknown): boolean => stableStringify(a) === stableStringify(b);

/**
 * SharePoint Online caps a plain-text multi-line column at 63,999 characters. We
 * chunk below that with headroom for the REST envelope, and the schema provisions
 * a fixed number of chunk columns.
 */
export const CHUNK_SIZE: number = 60000;
export const MAX_CHUNKS: number = 8;
export const MAX_DOCUMENT_CHARS: number = CHUNK_SIZE * MAX_CHUNKS;

export class PayloadTooLargeError extends Error {
  public readonly chars: number;
  public constructor(chars: number) {
    super(
      `Graph is ${Math.round(chars / 1024)} KB, over the ${Math.round(MAX_DOCUMENT_CHARS / 1024)} KB ` +
      `limit for document storage. Switch this project to per-item storage, which has no size ceiling.`
    );
    this.name = 'PayloadTooLargeError';
    this.chars = chars;
  }
}

/** Split a serialized bundle across the provisioned payload columns. */
export const splitPayload = (text: string): string[] => {
  if (text.length > MAX_DOCUMENT_CHARS) { throw new PayloadTooLargeError(text.length); }
  const chunks: string[] = [];
  for (let i = 0; i < text.length; i += CHUNK_SIZE) {
    chunks.push(text.slice(i, i + CHUNK_SIZE));
  }
  // An empty payload still occupies one (empty) chunk so callers can rely on chunks[0].
  return chunks.length > 0 ? chunks : [''];
};

/** Reassemble payload columns in order, ignoring unused trailing columns. */
export const joinPayload = (chunks: (string | null | undefined)[]): string => {
  let out = '';
  for (const c of chunks) {
    if (c === null || c === undefined || c === '') { continue; }
    out += c;
  }
  return out;
};

/** Guarantee the optional collections exist so downstream code can assume arrays. */
export const normalizeGraph = (graph: IGraph): IGraph => {
  if (!Array.isArray(graph.nodes)) { graph.nodes = []; }
  if (!Array.isArray(graph.edges)) { graph.edges = []; }
  if (!Array.isArray(graph.customNodeTypes)) { graph.customNodeTypes = []; }
  if (!Array.isArray(graph.collapsedNodes)) { graph.collapsedNodes = []; }
  if (!graph.positions || typeof graph.positions !== 'object') { graph.positions = {}; }
  return graph;
};

/**
 * Parse a stored payload into a bundle, accepting every shape the tool has ever
 * written: current bundle, legacy graph-only, and empty (a freshly created project).
 */
export const parseBundle = (text: string | null | undefined): IBundle => {
  if (!text || !text.trim()) {
    return { version: BUNDLE_SCHEMA_VERSION, graph: emptyGraph(), snapshots: [] };
  }
  const obj = JSON.parse(text) as Record<string, unknown>;
  const asGraph = obj as unknown as IGraph;
  const inner = obj.graph as IGraph | undefined;

  if (inner && Array.isArray(inner.nodes) && Array.isArray(inner.edges)) {
    return {
      version: typeof obj.version === 'number' ? obj.version : BUNDLE_SCHEMA_VERSION,
      graph: normalizeGraph(inner),
      snapshots: Array.isArray(obj.snapshots) ? obj.snapshots : [],
      lastModifiedBy: (obj.lastModifiedBy as string) || null,
      lastModifiedAt: (obj.lastModifiedAt as string) || null
    };
  }
  if (Array.isArray(asGraph.nodes) && Array.isArray(asGraph.edges)) {
    return { version: BUNDLE_SCHEMA_VERSION, graph: normalizeGraph(asGraph), snapshots: [] };
  }
  throw new Error('Stored payload is not a recognized graph bundle');
};

export const serializeBundle = (bundle: IBundle): string => JSON.stringify(bundle);
