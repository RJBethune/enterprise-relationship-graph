import { IExpectedField, IExpectedList } from './schema';

/**
 * Turn the gap between the declared schema and what a site actually has into a typed,
 * reviewable plan.
 *
 * Safety rules, in order of importance:
 *  - ADDITIVE ONLY. The plan creates lists and columns and turns settings on. It never
 *    deletes, renames, retypes, or removes anything. An "Update Schema" button that can
 *    destroy data is a button nobody will press.
 *  - An existing column with the WRONG TYPE is a conflict, never an action. SharePoint
 *    cannot safely change a column's type, so that repair stays human.
 *  - An UNVERIFIED list (the read failed) is a conflict. Provisioning blind against a
 *    list that might already exist is how you get duplicate columns.
 *
 * The same plan drives both the health display and the fix: when it is empty, the site
 * is healthy, by construction. Check and fix cannot drift because they are one function.
 */

export interface IFieldSnapshot {
  internal: string;
  /** SharePoint's TypeAsString, e.g. 'Text', 'Note', 'Lookup'. */
  typeAsString: string;
  indexed: boolean;
  unique: boolean;
  choices?: string[];
  /** SP.RelationshipDeleteBehaviorType: 0 None, 1 Cascade, 2 Restrict. */
  relationshipDeleteBehavior?: number;
}

export interface IListSnapshot {
  title: string;
  exists: boolean;
  /** False when the read failed (permissions, transient error) — never provision blind. */
  verified: boolean;
  versioning: boolean;
  fields: { [internal: string]: IFieldSnapshot };
  /** What actually went wrong when `verified` is false. Surfaced to the operator: a
   *  generic "could not be read" reads like a permissions problem even when it is a
   *  malformed query, which makes the real cause almost impossible to find from the UI. */
  readError?: string;
}

export type ProvisioningAction =
  | { kind: 'createList'; list: string; description: string }
  | { kind: 'createField'; list: string; field: IExpectedField }
  | { kind: 'addChoices'; list: string; internal: string; add: string[]; merged: string[] }
  | { kind: 'setIndexed'; list: string; internal: string }
  | { kind: 'setUnique'; list: string; internal: string }
  | { kind: 'setLookupBehavior'; list: string; internal: string; behavior: 'Cascade' | 'None' }
  | { kind: 'enableVersioning'; list: string };

export interface IProvisioningConflict {
  list: string;
  internal?: string;
  reason: string;
}

export interface IProvisioningPlan {
  /** Ordered for execution: every createList first, so lookup targets exist before the
   *  fields that point at them; then columns and settings in schema order. */
  actions: ProvisioningAction[];
  conflicts: IProvisioningConflict[];
}

const CASCADE = 1;

export const planProvisioning = (
  expected: IExpectedList[],
  snapshots: { [listTitle: string]: IListSnapshot | undefined }
): IProvisioningPlan => {
  const creates: ProvisioningAction[] = [];
  const rest: ProvisioningAction[] = [];
  const conflicts: IProvisioningConflict[] = [];

  for (const list of expected) {
    const snap = snapshots[list.title];

    if (snap && snap.exists && !snap.verified) {
      conflicts.push({
        list: list.title,
        reason: 'The list exists but its columns could not be read, so provisioning is skipped ' +
          'rather than run blind against it. SharePoint said: ' +
          (snap.readError || 'no error detail was captured') + '.'
      });
      continue;
    }

    const listExists = !!(snap && snap.exists);
    if (!listExists) {
      creates.push({ kind: 'createList', list: list.title, description: list.description });
    }

    if (list.versioning && (!listExists || !snap!.versioning)) {
      rest.push({ kind: 'enableVersioning', list: list.title });
    }

    for (const field of list.fields) {
      const live = listExists ? snap!.fields[field.internal] : undefined;

      if (!live) {
        if (field.builtIn) {
          // Title is never created: SharePoint puts it on every list. On a list this
          // plan is about to create, its settings still have to be applied — asking to
          // CREATE it would fail the step and leave the index or uniqueness unset.
          if (listExists) {
            conflicts.push({
              list: list.title, internal: field.internal,
              reason: `Built-in column ${field.internal} is missing from an existing list.`
            });
          } else if (field.unique) {
            rest.push({ kind: 'setUnique', list: list.title, internal: field.internal });
          } else if (field.indexed) {
            rest.push({ kind: 'setIndexed', list: list.title, internal: field.internal });
          }
          continue;
        }
        rest.push({ kind: 'createField', list: list.title, field: field });
        continue;
      }

      if (field.types.indexOf(live.typeAsString as IExpectedField['types'][0]) < 0) {
        conflicts.push({
          list: list.title, internal: field.internal,
          reason: `Column ${field.internal} is type ${live.typeAsString}, expected ` +
            `${field.types.join(' or ')}. SharePoint cannot change a column's type safely — ` +
            `rename the existing column and re-run to get a correct one.`
        });
        continue;
      }

      // Existing column of the right type: bring its settings up to spec, additively.
      if (field.unique && !live.unique) {
        rest.push({ kind: 'setUnique', list: list.title, internal: field.internal });
      } else if (field.indexed && !live.indexed) {
        rest.push({ kind: 'setIndexed', list: list.title, internal: field.internal });
      }

      if (field.choices && field.choices.length > 0) {
        const existing = live.choices || [];
        const missing = field.choices.filter((c) => existing.indexOf(c) < 0);
        if (missing.length > 0) {
          // Merge preserves every existing value verbatim and in place — a site may have
          // added its own, and reordering a Choice column rewrites nothing but confuses users.
          rest.push({
            kind: 'addChoices', list: list.title, internal: field.internal,
            add: missing, merged: existing.concat(missing)
          });
        }
      }

      if (field.lookup && field.lookup.behavior === 'Cascade' && live.relationshipDeleteBehavior !== CASCADE) {
        rest.push({
          kind: 'setLookupBehavior', list: list.title, internal: field.internal, behavior: 'Cascade'
        });
      }
    }
  }

  return { actions: creates.concat(rest), conflicts };
};

export const isHealthy = (plan: IProvisioningPlan): boolean =>
  plan.actions.length === 0 && plan.conflicts.length === 0;

export const describeAction = (a: ProvisioningAction): string => {
  switch (a.kind) {
    case 'createList': return `Create list "${a.list}"`;
    case 'createField': return `Add column ${a.field.display} (${a.field.internal}) to ${a.list}`;
    case 'addChoices': return `Add choice${a.add.length === 1 ? '' : 's'} ${a.add.join(', ')} to ${a.list}.${a.internal}`;
    case 'setIndexed': return `Index ${a.list}.${a.internal}`;
    case 'setUnique': return `Enforce unique values on ${a.list}.${a.internal}`;
    case 'setLookupBehavior': return `Set ${a.list}.${a.internal} delete behavior to ${a.behavior}`;
    case 'enableVersioning': return `Turn on version history for ${a.list}`;
    default: return 'Unknown action';
  }
};

/** One-line status for the setup panel. */
export const planSummary = (plan: IProvisioningPlan): string => {
  if (isHealthy(plan)) { return 'The SharePoint backend is up to date.'; }
  const lists = plan.actions.filter((a) => a.kind === 'createList').length;
  const cols = plan.actions.filter((a) => a.kind === 'createField').length;
  const other = plan.actions.length - lists - cols;
  const bits: string[] = [];
  if (lists) { bits.push(`${lists} list${lists === 1 ? '' : 's'}`); }
  if (cols) { bits.push(`${cols} column${cols === 1 ? '' : 's'}`); }
  if (other) { bits.push(`${other} setting${other === 1 ? '' : 's'}`); }
  const head = bits.length ? `${bits.join(', ')} to create or update.` : '';
  const tail = plan.conflicts.length
    ? ` ${plan.conflicts.length} item${plan.conflicts.length === 1 ? '' : 's'} need${plan.conflicts.length === 1 ? 's' : ''} a person.`
    : '';
  return (head + tail).trim();
};
