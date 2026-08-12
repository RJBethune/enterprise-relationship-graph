import { SpRest, SpRestError } from './SpRest';
import { fieldSchemaXml, lookupToken } from '../../provisioning/fieldXml';
import { EXPECTED_SCHEMA, IExpectedList } from '../../provisioning/schema';
import {
  IListSnapshot, IFieldSnapshot, ProvisioningAction, IProvisioningPlan, planProvisioning, describeAction
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

  public async readSnapshots(
    expected: IExpectedList[] = EXPECTED_SCHEMA
  ): Promise<{ [list: string]: IListSnapshot }> {
    const out: { [list: string]: IListSnapshot } = {};
    for (const list of expected) {
      out[list.title] = await this.readOne(list.title);
    }
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
      return { title, exists: true, verified: false, versioning: false, fields: {} };
    }

    if (listInfo.Id) { this.listIds[title] = listInfo.Id; }

    try {
      const fields = await this.sp.getAll<ISpFieldDto>(
        `web/lists/getbytitle('${encodeURIComponent(title)}')/fields` +
        '?$select=InternalName,TypeAsString,Indexed,EnforceUniqueValues,Choices,RelationshipDeleteBehavior' +
        '&$top=500'
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
      return { title, exists: true, verified: true, versioning: !!listInfo.EnableVersioning, fields: map };
    } catch (_e) {
      return { title, exists: true, verified: false, versioning: !!listInfo.EnableVersioning, fields: {} };
    }
  }

  public async buildPlan(expected: IExpectedList[] = EXPECTED_SCHEMA): Promise<IProvisioningPlan> {
    const snapshots = await this.readSnapshots(expected);
    return planProvisioning(expected, snapshots);
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
        } catch (_e) { /* cosmetic; the list itself exists */ }
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
