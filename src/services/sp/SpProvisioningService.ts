import { SpRest, SpRestError } from './SpRest';
import { fieldSchemaXml, lookupToken } from '../../provisioning/fieldXml';
import { EXPECTED_SCHEMA, IExpectedList, schemaFingerprint } from '../../provisioning/schema';
import {
  IListSnapshot, IFieldSnapshot, ProvisioningAction, IProvisioningPlan, planProvisioning,
  describeAction, isHealthy
} from '../../provisioning/planner';

/** SP.AddFieldOptions.AddFieldInternalNameHint — makes the Name attribute the internal
 *  name verbatim, which is what removes the display-name rename trap the by-hand
 *  provisioning guides spend so many warnings on. */
const ADD_FIELD_INTERNAL_NAME_HINT = 8;

const GENERIC_LIST = 100;
const CASCADE = 1;

export interface IProvisioningStepResult {
  action: ProvisioningAction;
  label: string;
  ok: boolean;
  error?: string;
}

export interface IProvisioningService {
  readSnapshots(expected?: IExpectedList[]): Promise<{ [list: string]: IListSnapshot }>;
  buildPlan(expected?: IExpectedList[]): Promise<IProvisioningPlan>;
  execute(
    plan: IProvisioningPlan,
    onProgress?: (done: number, total: number, label: string) => void
  ): Promise<IProvisioningStepResult[]>;
  /** This exact schema was confirmed healthy on this site recently — skip the check. */
  wasVerifiedHealthy(): boolean;
  forgetHealthy(): void;
}

interface ISpFieldDto {
  InternalName: string;
  TypeAsString: string;
  Indexed: boolean;
  EnforceUniqueValues: boolean;
  Choices?: string[];
  RelationshipDeleteBehavior?: number;
}

/**
 * Reads the live schema and applies plans. Deliberately thin: every decision about
 * what to create, skip or refuse was made by the pure planner, so this class only
 * translates actions into requests and reports what happened.
 *
 * It runs entirely in the browser as the signed-in user, which is why it works in a
 * tenant where PnP PowerShell cannot connect at all: this is the same SPHttpClient
 * path the web part already uses to read data, not a server-side provisioning channel.
 */
export class SpProvisioningService implements IProvisioningService {
  private readonly listIds: { [title: string]: string } = {};

  public constructor(private readonly sp: SpRest) {}

  /**
   * Read every list's live shape.
   *
   * Fanned out across lists rather than looped: the reads are independent, and doing
   * them one after another turned a schema check into fifteen sequential round trips
   * of dead time before the graph could open. The call COUNT is unchanged — that was
   * never the problem, SharePoint throttles on sustained volume, not on a handful of
   * reads — but the depth drops from fifteen to two.
   */
  public async readSnapshots(
    expected: IExpectedList[] = EXPECTED_SCHEMA
  ): Promise<{ [list: string]: IListSnapshot }> {
    const snapshots = await Promise.all(expected.map((list) => this.readOne(list.title)));
    const out: { [list: string]: IListSnapshot } = {};
    for (const snap of snapshots) { out[snap.title] = snap; }
    return out;
  }

  private async readOne(title: string): Promise<IListSnapshot> {
    const empty: IListSnapshot = { title, exists: false, verified: true, versioning: false, fields: {} };
    let listInfo: { EnableVersioning?: boolean; Id?: string };
    try {
      listInfo = await this.sp.get<{ EnableVersioning: boolean; Id: string }>(
        `web/lists/getbytitle('${encodeURIComponent(title)}')?$select=Id,EnableVersioning`
      );
    } catch (e) {
      // 404 is the ordinary "not deployed yet" answer. Anything else means we could
      // not see the list, and provisioning blind against it would double-create columns.
      if (e instanceof SpRestError && e.status === 404) { return empty; }
      // Some tenants answer a missing list with 400 + "List 'X' does not exist" rather
      // than 404. Treating that as unreadable would block provisioning on a site that
      // simply has nothing deployed yet.
      const message = e instanceof Error ? e.message : String(e);
      if (/does not exist/i.test(message)) { return empty; }
      return {
        title, exists: true, verified: false, versioning: false, fields: {},
        readError: message
      };
    }

    if (listInfo.Id) { this.listIds[title] = listInfo.Id; }

    // The columns and the view are independent reads; the list read above had to come
    // first only because it decides whether these are worth making at all.
    const viewFieldsPromise = this.readViewFields(title);

    try {
      // NO $select here, deliberately.
      //
      // /fields is a POLYMORPHIC collection whose declared type is SP.Field. Naming a
      // subtype-only property in $select — Choices (SP.FieldChoice),
      // RelationshipDeleteBehavior (SP.FieldLookup) — makes SharePoint reject the WHOLE
      // query with HTTP 400, not just drop that column. The failure then presents as
      // "the list exists but could not be read", which reads like a permissions problem
      // and is not one. Requesting the full entity returns each field serialized as its
      // ACTUAL type, so the subtype properties arrive for the fields that have them.
      const fields = await this.sp.getAll<ISpFieldDto>(
        `web/lists/getbytitle('${encodeURIComponent(title)}')/fields?$top=500`
      );
      const map: { [internal: string]: IFieldSnapshot } = {};
      for (const f of fields) {
        map[f.InternalName] = {
          internal: f.InternalName,
          typeAsString: f.TypeAsString,
          indexed: !!f.Indexed,
          unique: !!f.EnforceUniqueValues,
          choices: Array.isArray(f.Choices) ? f.Choices : undefined,
          relationshipDeleteBehavior: f.RelationshipDeleteBehavior
        };
      }
      return {
        title, exists: true, verified: true, versioning: !!listInfo.EnableVersioning, fields: map,
        viewFields: await viewFieldsPromise
      };
    } catch (e) {
      await viewFieldsPromise.catch(() => []);   // never leave it unhandled
      return {
        title, exists: true, verified: false, versioning: !!listInfo.EnableVersioning, fields: {},
        readError: e instanceof Error ? e.message : String(e)
      };
    }
  }

  /**
   * Columns currently on the list's default view.
   *
   * A failure here returns an empty list rather than throwing: a view we cannot read
   * is a cosmetic problem, and it must not turn a healthy, fully provisioned site into
   * an "unreadable" conflict.
   */
  private async readViewFields(title: string): Promise<string[]> {
    try {
      const res = await this.sp.get<{ Items?: string[] | { results?: string[] }; value?: string[] }>(
        `${this.listPath(title)}/defaultView/viewfields`
      );
      const items = res.Items;
      if (Array.isArray(items)) { return items; }
      if (items && Array.isArray(items.results)) { return items.results; }
      if (Array.isArray(res.value)) { return res.value; }
      return [];
    } catch {
      return [];
    }
  }

  public async buildPlan(expected: IExpectedList[] = EXPECTED_SCHEMA): Promise<IProvisioningPlan> {
    const snapshots = await this.readSnapshots(expected);
    const plan = planProvisioning(expected, snapshots);
    // Only a CLEAN result is remembered. Caching a gap would keep the warning on
    // screen after somebody fixed it, and the whole point of the warning is that it
    // goes away when the problem does.
    if (isHealthy(plan)) { this.rememberHealthy(); } else { this.forgetHealthy(); }
    return plan;
  }

  /* ----------------------------------------------------------- healthy cache */

  /**
   * Remembering a healthy verdict is what stops a fifteen-call schema check running
   * on every single page load, forever, to answer a question whose answer changes
   * about twice a year.
   *
   * It is keyed by the fingerprint of the schema shipped in THIS bundle, so a new
   * .sppkg with different expectations cannot match an old verdict — the check runs
   * again automatically on the first load after an upgrade, which is exactly when it
   * matters. Nobody has to remember to bump anything.
   */
  private static readonly HEALTHY_TTL_MS: number = 12 * 60 * 60 * 1000;

  private cacheKey(): string {
    return `erg.schemaOk.${this.sp.webUrl}.${schemaFingerprint()}`;
  }

  public wasVerifiedHealthy(): boolean {
    try {
      const at = Number(window.localStorage.getItem(this.cacheKey()));
      return !!at && (Date.now() - at) < SpProvisioningService.HEALTHY_TTL_MS;
    } catch {
      return false;
    }
  }

  private rememberHealthy(): void {
    try { window.localStorage.setItem(this.cacheKey(), String(Date.now())); } catch { /* private mode */ }
  }

  public forgetHealthy(): void {
    try { window.localStorage.removeItem(this.cacheKey()); } catch { /* private mode */ }
  }

  public async execute(
    plan: IProvisioningPlan,
    onProgress?: (done: number, total: number, label: string) => void
  ): Promise<IProvisioningStepResult[]> {
    const results: IProvisioningStepResult[] = [];
    let done = 0;
    for (const action of plan.actions) {
      const label = describeAction(action);
      if (onProgress) { onProgress(done, plan.actions.length, label); }
      try {
        await this.apply(action);
        results.push({ action, label, ok: true });
      } catch (e) {
        // One failed column must not abandon the rest of the plan: a half-provisioned
        // site is fixable by pressing the button again, but only if the button keeps going.
        results.push({
          action, label, ok: false,
          error: (e instanceof Error ? e.message : String(e)).slice(0, 300)
        });
      }
      done++;
      if (onProgress) { onProgress(done, plan.actions.length, label); }
    }
    return results;
  }

  private listPath(title: string): string {
    return `web/lists/getbytitle('${encodeURIComponent(title)}')`;
  }

  private fieldPath(list: string, internal: string): string {
    return `${this.listPath(list)}/fields/getbyinternalnameortitle('${encodeURIComponent(internal)}')`;
  }

  private async apply(action: ProvisioningAction): Promise<void> {
    switch (action.kind) {
      case 'createList': {
        const created = await this.sp.post<{ Id: string }>('web/lists', {
          Title: action.list,
          Description: action.description,
          BaseTemplate: GENERIC_LIST,
          AllowContentTypes: false,
          ContentTypesEnabled: false
        });
        if (created && created.Id) { this.listIds[action.list] = created.Id; }
        // Title is Required by default, but rows in the node/edge lists are identified
        // by their graph id, not their label — a blank label must not fail the write.
        try {
          await this.sp.merge(this.fieldPath(action.list, 'Title'), { Required: false });
        } catch { /* cosmetic; the list itself exists */ }
        return;
      }

      case 'createField': {
        let xml = fieldSchemaXml(action.field);
        if (action.field.lookup) {
          const id = await this.listId(action.field.lookup.list);
          xml = xml.split(lookupToken(action.field.lookup.list)).join(`{${id}}`);
        }
        await this.sp.post(`${this.listPath(action.list)}/fields/createfieldasxml`, {
          parameters: { SchemaXml: xml, Options: ADD_FIELD_INTERNAL_NAME_HINT }
        });
        return;
      }

      case 'addChoices':
        await this.sp.merge(this.fieldPath(action.list, action.internal), { Choices: action.merged });
        return;

      case 'setIndexed':
        await this.sp.merge(this.fieldPath(action.list, action.internal), { Indexed: true });
        return;

      case 'setUnique':
        // SharePoint requires the index before uniqueness; one update carries both.
        await this.sp.merge(this.fieldPath(action.list, action.internal), {
          Indexed: true, EnforceUniqueValues: true
        });
        return;

      case 'setLookupBehavior':
        await this.sp.merge(this.fieldPath(action.list, action.internal), {
          Indexed: true,
          RelationshipDeleteBehavior: action.behavior === 'Cascade' ? CASCADE : 0
        });
        return;

      case 'enableVersioning':
        await this.sp.merge(this.listPath(action.list), { EnableVersioning: true });
        return;

      case 'addViewFields':
        // Sequential, and each one tolerated individually: a built-in that this list
        // happens not to expose must not stop the rest of the columns appearing.
        for (const internal of action.add) {
          try {
            await this.sp.post(
              `${this.listPath(action.list)}/defaultView/viewfields/addviewfield('${encodeURIComponent(internal)}')`
            );
          } catch { /* cosmetic — the data is stored regardless of what the view shows */ }
        }
        return;

      default:
        return;
    }
  }

  /** Lookup columns need the target list's GUID, which only exists after creation. */
  private async listId(title: string): Promise<string> {
    if (this.listIds[title]) { return this.listIds[title]; }
    const info = await this.sp.get<{ Id: string }>(`${this.listPath(title)}?$select=Id`);
    this.listIds[title] = info.Id;
    return info.Id;
  }
}
