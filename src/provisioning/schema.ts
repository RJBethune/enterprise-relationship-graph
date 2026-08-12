import { MAX_CHUNKS } from '../model/bundle';

/**
 * The executable spec for the SharePoint backend.
 *
 * This declaration is the single source of truth for both halves of the Deploy /
 * Update Schema button: the health check reads it to decide what is missing, and the
 * planner reads it to decide what to create. One declaration means check and fix
 * cannot drift — the thing that makes "Update Schema" trustworthy rather than hopeful.
 *
 * Design notes worth keeping in view:
 *  - Node and edge ATTRIBUTES are not columns. They live in a plain-text `ErgData`
 *    column as JSON, because offices define custom node types at runtime and a column
 *    per attribute would mean a schema migration every time somebody invents a type.
 *    The handful of attributes worth querying (label, type, endpoints) are projected
 *    into real columns so list views and diagnostics still work.
 *  - Every Note column is plain text, never enhanced rich text: rich text HTML-encodes
 *    content, and a payload that does not round-trip verbatim is a corrupted graph.
 *  - Lookups to ERG Projects cascade, so deleting a project takes its nodes, edges and
 *    snapshots with it instead of orphaning thousands of rows.
 */

export type SpFieldType = 'Text' | 'Note' | 'Number' | 'Boolean' | 'DateTime' | 'Choice' | 'Lookup';

export interface IExpectedField {
  internal: string;
  display: string;
  /** Acceptable live types; the FIRST is the type used to create the column. */
  types: SpFieldType[];
  required?: boolean;
  indexed?: boolean;
  unique?: boolean;
  choices?: string[];
  defaultValue?: string;
  lookup?: { list: string; behavior: 'Cascade' | 'None' };
  /** SharePoint built-in (Title): never created, but its index/uniqueness is still enforced. */
  builtIn?: boolean;
  description?: string;
}

export interface IExpectedList {
  title: string;
  description: string;
  versioning: boolean;
  fields: IExpectedField[];
}

export const PROJECTS_LIST = 'ERG Projects';
export const NODES_LIST = 'ERG Nodes';
export const EDGES_LIST = 'ERG Edges';
export const SNAPSHOTS_LIST = 'ERG Snapshots';

/** ErgPayload1..N — the document-mode payload columns. */
const payloadFields = (): IExpectedField[] => {
  const fields: IExpectedField[] = [];
  for (let i = 1; i <= MAX_CHUNKS; i++) {
    fields.push({
      internal: `ErgPayload${i}`,
      display: `Payload ${i}`,
      types: ['Note'],
      description: i === 1
        ? 'Graph bundle JSON. Split across the payload columns when it exceeds one column.'
        : `Continuation ${i} of the graph bundle JSON.`
    });
  }
  return fields;
};

const projectLookup = (): IExpectedField => ({
  internal: 'ErgProject',
  display: 'Project',
  types: ['Lookup'],
  required: true,
  indexed: true,
  lookup: { list: PROJECTS_LIST, behavior: 'Cascade' },
  description: 'Owning project. Cascade delete: removing a project removes its rows.'
});

export const EXPECTED_SCHEMA: IExpectedList[] = [
  {
    title: PROJECTS_LIST,
    description: 'One item per relationship graph. The switcher lists these.',
    versioning: true,
    fields: [
      {
        internal: 'Title', display: 'Project name', types: ['Text'],
        builtIn: true, indexed: true, unique: true,
        description: 'Shown in the project switcher. Unique so links by name stay unambiguous.'
      },
      {
        internal: 'ErgStorageMode', display: 'Storage mode', types: ['Choice'],
        choices: ['Document', 'Items'], defaultValue: 'Document',
        description: 'Document = whole graph in the payload columns. Items = one row per node/edge.'
      },
      ...payloadFields(),
      {
        internal: 'ErgLayout', display: 'Layout', types: ['Note'],
        description: 'Hand-arranged node positions (items mode). Presentation only; one write per layout change.'
      },
      {
        internal: 'ErgCustomTypes', display: 'Custom node types', types: ['Note'],
        description: 'Office-defined node types (items mode).'
      },
      {
        internal: 'ErgCollapsed', display: 'Collapsed nodes', types: ['Note'],
        description: 'Collapsed branch ids (items mode).'
      },
      {
        internal: 'ErgSchemaVersion', display: 'Schema version', types: ['Number'], defaultValue: '4',
        description: 'Bundle format version, so older graphs can be migrated on load.'
      },
      { internal: 'ErgNodeCount', display: 'Nodes', types: ['Number'], description: 'Switcher metadata; avoids parsing the payload.' },
      { internal: 'ErgEdgeCount', display: 'Relationships', types: ['Number'], description: 'Switcher metadata.' },
      {
        internal: 'ErgStatus', display: 'Status', types: ['Choice'],
        choices: ['Active', 'Archived'], defaultValue: 'Active', indexed: true
      }
    ]
  },
  {
    title: NODES_LIST,
    description: 'One item per node, for projects using per-item storage.',
    versioning: false,
    fields: [
      { internal: 'Title', display: 'Label', types: ['Text'], builtIn: true },
      projectLookup(),
      {
        internal: 'ErgNodeId', display: 'Node id', types: ['Text'], required: true, indexed: true,
        description: 'The graph-level id. Client-generated, so writes are idempotent on retry.'
      },
      { internal: 'ErgType', display: 'Node type', types: ['Text'], indexed: true },
      {
        internal: 'ErgData', display: 'Data', types: ['Note'], required: true,
        description: 'The complete node as JSON. Open attribute set — see the header note.'
      }
    ]
  },
  {
    title: EDGES_LIST,
    description: 'One item per relationship, for projects using per-item storage.',
    versioning: false,
    fields: [
      { internal: 'Title', display: 'Label', types: ['Text'], builtIn: true },
      projectLookup(),
      { internal: 'ErgEdgeId', display: 'Relationship id', types: ['Text'], required: true, indexed: true },
      { internal: 'ErgSource', display: 'Source node id', types: ['Text'], indexed: true },
      { internal: 'ErgTarget', display: 'Target node id', types: ['Text'], indexed: true },
      { internal: 'ErgType', display: 'Relationship type', types: ['Text'] },
      { internal: 'ErgData', display: 'Data', types: ['Note'], required: true }
    ]
  },
  {
    title: SNAPSHOTS_LIST,
    description: 'Named restore points. Independent of list version history, which captures every save.',
    versioning: false,
    fields: [
      { internal: 'Title', display: 'Snapshot name', types: ['Text'], builtIn: true },
      projectLookup(),
      ...payloadFields(),
      { internal: 'ErgNote', display: 'Note', types: ['Note'], description: 'Why this restore point matters.' }
    ]
  }
];

/** Lists needed before a project can be opened at all. */
export const CORE_LISTS: string[] = [PROJECTS_LIST];

/** Lists only per-item storage needs; a Document-mode-only site can skip them. */
export const ITEM_MODE_LISTS: string[] = [NODES_LIST, EDGES_LIST];

export const getExpectedList = (title: string): IExpectedList | undefined =>
  EXPECTED_SCHEMA.filter((l) => l.title === title)[0];
