import { ISpTransport, ISpRequestOptions, ISpRawResponse } from '../../src/services/sp/SpRest';

/**
 * An in-memory SharePoint that speaks enough REST to run the real data layer.
 *
 * This exists because the interesting failures in this codebase are not unit-sized.
 * "Does a concurrent edit merge correctly", "does a per-item save write exactly the
 * rows that changed", "does provisioning leave a site the health check calls healthy" —
 * those only appear when the provisioner, the stores, the batch encoder and the merge
 * engine run together. Against this fake they run in milliseconds, with no tenant, no
 * credentials and no network.
 *
 * It implements the subset the app actually calls, and it enforces the parts that bite:
 * ETag concurrency, per-list change logs including deletes, and cascade delete.
 */

interface IFakeField {
  internal: string;
  typeAsString: string;
  indexed: boolean;
  unique: boolean;
  choices?: string[];
  relationshipDeleteBehavior?: number;
  lookupList?: string;
}

interface IFakeItem {
  Id: number;
  version: number;
  data: { [k: string]: unknown };
  editor: string;
  modified: string;
}

interface IFakeChange { changeType: number; itemId: number; seq: number; }

interface IFakeList {
  id: string;
  title: string;
  versioning: boolean;
  fields: Map<string, IFakeField>;
  items: Map<number, IFakeItem>;
  nextId: number;
  changes: IFakeChange[];
}

const CHANGE_ADD = 1;
const CHANGE_UPDATE = 2;
const CHANGE_DELETE = 3;

export class FakeSharePoint implements ISpTransport {
  public readonly webUrl: string = 'https://contoso.sharepoint.us/sites/exec';
  public readonly lists: Map<string, IFakeList> = new Map();
  /** Every request, for asserting on write volume — the whole point of the storage debate. */
  public readonly log: { method: string; url: string }[] = [];
  public currentUser: string = 'Test Editor';
  /** Set to fail the next N writes, to exercise the retry ladder. */
  public failWrites: number = 0;
  private clock: number = 0;
  private seq: number = 0;

  public writeCount(): number {
    return this.log.filter((r) => r.method !== 'GET' && r.url.indexOf('$batch') < 0).length;
  }

  public reset(): void { this.log.length = 0; }

  private now(): string {
    this.clock += 1000;
    return new Date(Date.parse('2026-01-01T00:00:00Z') + this.clock).toISOString();
  }

  private ok(json: unknown, etag?: string, status: number = 200): ISpRawResponse {
    const text = json === undefined ? '' : JSON.stringify(json);
    return { status, ok: true, text, json: json === undefined ? null : json, etag: etag || null, retryAfterMs: 0 };
  }

  private err(status: number, message: string): ISpRawResponse {
    const json = { error: { message: { value: message } } };
    return { status, ok: false, text: JSON.stringify(json), json, etag: null, retryAfterMs: status === 429 ? 3000 : 0 };
  }

  /* ------------------------------------------------------------------ routing */

  public async request(
    method: string, absoluteUrl: string, options?: ISpRequestOptions
  ): Promise<ISpRawResponse> {
    const opts = options || {};
    const override = (opts.headers && opts.headers['X-HTTP-Method']) || '';
    const verb = override || method;
    const path = absoluteUrl.replace(`${this.webUrl}/_api/`, '');
    this.log.push({ method: verb, url: path });

    if (verb !== 'GET' && this.failWrites > 0) {
      this.failWrites--;
      return this.err(429, 'Too many requests');
    }

    if (path.indexOf('$batch') === 0) { return this.handleBatch(opts.body || ''); }
    return this.route(verb, path, opts);
  }

  private async route(verb: string, path: string, opts: ISpRequestOptions): Promise<ISpRawResponse> {
    const body = opts.body ? JSON.parse(opts.body) as { [k: string]: unknown } : {};

    if (path === 'web/lists' && verb === 'POST') { return this.createList(body); }

    const listMatch = /^web\/lists\/getbytitle\('([^']+)'\)(.*)$/.exec(path);
    if (!listMatch) { return this.err(400, `Unrouted path: ${path}`); }

    const title = decodeURIComponent(listMatch[1]);
    const rest = listMatch[2];
    const list = this.lists.get(title);

    if (!list) { return this.err(404, `List '${title}' does not exist.`); }

    if (rest.indexOf('/fields/createfieldasxml') === 0 && verb === 'POST') {
      return this.createFieldAsXml(list, body);
    }
    const fieldMatch = /^\/fields\/getbyinternalnameortitle\('([^']+)'\)/.exec(rest);
    if (fieldMatch) { return this.updateField(list, decodeURIComponent(fieldMatch[1]), body, verb); }
    if (rest.indexOf('/fields') === 0 && verb === 'GET') { return this.readFields(list); }
    if (rest.indexOf('/getchanges') === 0 && verb === 'POST') { return this.getChanges(list, body); }

    const itemMatch = /^\/items\((\d+)\)/.exec(rest);
    if (itemMatch) { return this.itemById(list, Number(itemMatch[1]), verb, body, opts, rest); }
    if (rest.indexOf('/items') === 0) {
      if (verb === 'POST') { return this.createItem(list, body); }
      return this.queryItems(list, rest);
    }
    if (rest === '' || rest.charAt(0) === '?') { return this.readList(list, verb, body, rest); }

    return this.err(400, `Unrouted list path: ${rest}`);
  }

  /* -------------------------------------------------------------- list + fields */

  private createList(body: { [k: string]: unknown }): ISpRawResponse {
    const title = String(body.Title);
    if (this.lists.has(title)) { return this.err(400, `List '${title}' already exists.`); }
    const list: IFakeList = {
      id: `guid-${title.replace(/\W+/g, '-').toLowerCase()}`,
      title,
      versioning: false,
      // Every SharePoint list is born with Title.
      fields: new Map([['Title', { internal: 'Title', typeAsString: 'Text', indexed: false, unique: false }]]),
      items: new Map(),
      nextId: 1,
      changes: []
    };
    this.lists.set(title, list);
    return this.ok({ Id: list.id, Title: title }, undefined, 201);
  }

  private readList(list: IFakeList, verb: string, body: { [k: string]: unknown }, rest: string): ISpRawResponse {
    if (verb === 'MERGE') {
      if (body.EnableVersioning !== undefined) { list.versioning = !!body.EnableVersioning; }
      return this.ok(undefined, undefined, 204);
    }
    const payload: { [k: string]: unknown } = { Id: list.id, Title: list.title, EnableVersioning: list.versioning };
    if (rest.indexOf('CurrentChangeToken') >= 0) {
      payload.CurrentChangeToken = { StringValue: `1;3;${list.id};${this.seq}` };
    }
    return this.ok(payload);
  }

  private readFields(list: IFakeList): ISpRawResponse {
    return this.ok({
      value: Array.from(list.fields.values()).map((f) => ({
        InternalName: f.internal,
        TypeAsString: f.typeAsString,
        Indexed: f.indexed,
        EnforceUniqueValues: f.unique,
        Choices: f.choices,
        RelationshipDeleteBehavior: f.relationshipDeleteBehavior
      }))
    });
  }

  private createFieldAsXml(list: IFakeList, body: { [k: string]: unknown }): ISpRawResponse {
    const params = body.parameters as { SchemaXml: string } | undefined;
    if (!params || !params.SchemaXml) { return this.err(400, 'Missing SchemaXml'); }
    const xml = params.SchemaXml;

    // The attribute name must be anchored to whitespace: an unanchored /Name="/ also
    // matches DisplayName= and StaticName=, which would silently give every column the
    // wrong internal name — and SharePoint's real parser does not make that mistake.
    const attr = (name: string): string | undefined => {
      const m = new RegExp(`(?:^|\\s)${name}="([^"]*)"`).exec(xml);
      return m ? m[1] : undefined;
    };
    const internal = attr('Name');
    const type = attr('Type');
    if (!internal || !type) { return this.err(400, 'SchemaXml missing Name or Type'); }
    if (list.fields.has(internal)) { return this.err(400, `Column '${internal}' already exists.`); }

    const choices: string[] = [];
    const choiceRe = /<CHOICE>([^<]*)<\/CHOICE>/g;
    let cm = choiceRe.exec(xml);
    while (cm !== null) { choices.push(cm[1]); cm = choiceRe.exec(xml); }

    list.fields.set(internal, {
      internal,
      typeAsString: type,
      indexed: /Indexed="TRUE"/.test(xml),
      unique: /EnforceUniqueValues="TRUE"/.test(xml),
      choices: type === 'Choice' ? choices : undefined,
      relationshipDeleteBehavior: type === 'Lookup'
        ? (/RelationshipDeleteBehavior="Cascade"/.test(xml) ? 1 : 0)
        : undefined,
      lookupList: type === 'Lookup' ? (attr('List') || '').replace(/[{}]/g, '') : undefined
    });
    return this.ok({ InternalName: internal }, undefined, 201);
  }

  private updateField(
    list: IFakeList, internal: string, body: { [k: string]: unknown }, verb: string
  ): ISpRawResponse {
    const field = list.fields.get(internal);
    if (!field) { return this.err(404, `Column '${internal}' does not exist.`); }
    if (verb !== 'MERGE') { return this.ok(field); }
    if (body.Indexed !== undefined) { field.indexed = !!body.Indexed; }
    if (body.EnforceUniqueValues !== undefined) { field.unique = !!body.EnforceUniqueValues; }
    if (body.Choices !== undefined) { field.choices = body.Choices as string[]; }
    if (body.RelationshipDeleteBehavior !== undefined) {
      field.relationshipDeleteBehavior = body.RelationshipDeleteBehavior as number;
    }
    return this.ok(undefined, undefined, 204);
  }

  /* -------------------------------------------------------------------- items */

  private itemPayload(list: IFakeList, item: IFakeItem, select?: string): { [k: string]: unknown } {
    const out: { [k: string]: unknown } = { Id: item.Id, Modified: item.modified };
    list.fields.forEach((_f, name) => {
      if (name === 'Title') { out.Title = item.data.Title === undefined ? null : item.data.Title; }
    });
    for (const key of Object.keys(item.data)) { out[key] = item.data[key]; }
    if (!select || select.indexOf('Editor') >= 0) { out.Editor = { Title: item.editor }; }
    return out;
  }

  private createItem(list: IFakeList, body: { [k: string]: unknown }): ISpRawResponse {
    const unique = Array.from(list.fields.values()).filter((f) => f.unique);
    for (const f of unique) {
      const value = body[f.internal];
      if (value === undefined) { continue; }
      let clash = false;
      list.items.forEach((it) => { if (it.data[f.internal] === value) { clash = true; } });
      if (clash) {
        return this.err(400, `A value on column '${f.internal}' must be unique; '${String(value)}' is taken.`);
      }
    }

    const item: IFakeItem = {
      Id: list.nextId++,
      version: 1,
      data: { ...body },
      editor: this.currentUser,
      modified: this.now()
    };
    list.items.set(item.Id, item);
    list.changes.push({ changeType: CHANGE_ADD, itemId: item.Id, seq: ++this.seq });
    return this.ok(this.itemPayload(list, item), `"${item.version}"`, 201);
  }

  private itemById(
    list: IFakeList, id: number, verb: string,
    body: { [k: string]: unknown }, opts: ISpRequestOptions, rest: string
  ): ISpRawResponse {
    const item = list.items.get(id);
    if (!item) { return this.err(404, `Item ${id} does not exist in '${list.title}'.`); }

    if (verb === 'GET') {
      const select = /[?&]\$select=([^&]*)/.exec(rest);
      return this.ok(this.itemPayload(list, item, select ? select[1] : undefined), `"${item.version}"`);
    }

    // ETag concurrency: this is the mechanism the whole conflict story rests on, so
    // the fake enforces it exactly rather than waving it through.
    const ifMatch = opts.etag;
    if (ifMatch && ifMatch !== '*' && ifMatch !== `"${item.version}"`) {
      return this.err(412, 'The item was modified by another user.');
    }

    if (verb === 'DELETE') {
      list.items.delete(id);
      list.changes.push({ changeType: CHANGE_DELETE, itemId: id, seq: ++this.seq });
      this.cascadeDelete(list, id);
      return this.ok(undefined, undefined, 204);
    }

    if (verb === 'MERGE') {
      for (const key of Object.keys(body)) { item.data[key] = body[key]; }
      item.version++;
      item.editor = this.currentUser;
      item.modified = this.now();
      list.changes.push({ changeType: CHANGE_UPDATE, itemId: id, seq: ++this.seq });
      return this.ok(undefined, undefined, 204);
    }

    return this.err(400, `Unsupported verb ${verb}`);
  }

  /** Lookup columns declared Cascade take their dependent rows with them. */
  private cascadeDelete(parent: IFakeList, parentItemId: number): void {
    this.lists.forEach((list) => {
      list.fields.forEach((field) => {
        if (field.typeAsString !== 'Lookup' || field.relationshipDeleteBehavior !== 1) { return; }
        if ((field.lookupList || '') !== parent.id) { return; }
        const doomed: number[] = [];
        list.items.forEach((item, itemId) => {
          if (item.data[`${field.internal}Id`] === parentItemId) { doomed.push(itemId); }
        });
        for (const itemId of doomed) {
          list.items.delete(itemId);
          list.changes.push({ changeType: CHANGE_DELETE, itemId, seq: ++this.seq });
        }
      });
    });
  }

  private queryItems(list: IFakeList, rest: string): ISpRawResponse {
    const filterMatch = /[?&]\$filter=([^&]*)/.exec(rest);
    const orderMatch = /[?&]\$orderby=([^&]*)/.exec(rest);
    const selectMatch = /[?&]\$select=([^&]*)/.exec(rest);

    let items = Array.from(list.items.values());
    if (filterMatch) {
      const filter = decodeURIComponent(filterMatch[1]);
      items = items.filter((item) => this.matches(item, filter));
    }
    if (orderMatch) {
      const key = decodeURIComponent(orderMatch[1]).split(' ')[0];
      items = items.sort((a, b) => String(a.data[key] || '').localeCompare(String(b.data[key] || '')));
    }
    const select = selectMatch ? selectMatch[1] : undefined;
    return this.ok({ value: items.map((i) => this.itemPayload(list, i, select)) });
  }

  /** Supports exactly the three filter shapes the app builds. */
  private matches(item: IFakeItem, filter: string): boolean {
    if (/\bor\b/.test(filter)) {
      return filter.split(/\s+or\s+/).some((clause) => this.matches(item, clause.trim()));
    }
    let m = /^(\w+) eq (\d+)$/.exec(filter);
    if (m) {
      const value = m[1] === 'Id' ? item.Id : item.data[m[1]];
      return Number(value) === Number(m[2]);
    }
    m = /^(\w+) ne '(.*)'$/.exec(filter);
    if (m) { return String(item.data[m[1]] || '') !== m[2]; }
    m = /^(\w+) eq '(.*)'$/.exec(filter);
    if (m) { return String(item.data[m[1]] || '') === m[2]; }
    return true;
  }

  private getChanges(list: IFakeList, body: { [k: string]: unknown }): ISpRawResponse {
    const query = body.query as { ChangeTokenStart?: { StringValue?: string } } | undefined;
    const token = (query && query.ChangeTokenStart && query.ChangeTokenStart.StringValue) || '';
    const since = Number(token.split(';')[3] || 0);
    const changes = list.changes.filter((c) => c.seq > since);
    return this.ok({
      value: changes.map((c) => ({
        ChangeType: c.changeType,
        ItemId: c.itemId,
        ChangeToken: { StringValue: `1;3;${list.id};${c.seq}` }
      }))
    });
  }

  /* -------------------------------------------------------------------- batch */

  /**
   * Parse the envelope the app generates and dispatch each part.
   *
   * Round-tripping our own multipart format through a parser that did not write it is
   * itself a check: a malformed changeset shows up here as an unrouted path rather
   * than as a mysterious 400 from a real tenant six weeks later.
   */
  private async handleBatch(body: string): Promise<ISpRawResponse> {
    const parts = body.split(/\r\n--(?:batch|changeset)_[^\r\n]*\r\n/);
    const responses: string[] = [];

    for (const part of parts) {
      const m = /^(GET|POST|MERGE|DELETE) (\S+) HTTP\/1\.1\r\n([\s\S]*)$/m.exec(part.trim());
      if (!m) { continue; }
      const verb = m[1];
      const url = m[2];
      const tail = m[3];
      const blank = tail.indexOf('\r\n\r\n');
      const headerBlock = blank >= 0 ? tail.slice(0, blank) : tail;
      const partBody = blank >= 0 ? tail.slice(blank + 4).trim() : '';
      const etagMatch = /^IF-MATCH:\s*(.+)$/im.exec(headerBlock);

      const res = await this.request(
        verb === 'MERGE' || verb === 'DELETE' ? 'POST' : verb,
        url,
        {
          body: partBody || undefined,
          etag: etagMatch ? etagMatch[1].trim() : undefined,
          headers: verb === 'MERGE' || verb === 'DELETE' ? { 'X-HTTP-Method': verb } : undefined
        }
      );

      responses.push([
        '--changesetresponse_f',
        'Content-Type: application/http',
        '',
        `HTTP/1.1 ${res.status} ${res.ok ? 'OK' : 'Error'}`,
        'CONTENT-TYPE: application/json;odata=nometadata',
        res.etag ? `ETAG: ${res.etag}` : 'X-Fake: 1',
        '',
        res.text
      ].join('\r\n'));
    }

    const text = [
      '--batchresponse_f',
      'Content-Type: multipart/mixed; boundary=changesetresponse_f',
      '',
      responses.join('\r\n'),
      '--changesetresponse_f--',
      '--batchresponse_f--',
      ''
    ].join('\r\n');

    return { status: 200, ok: true, text, json: null, etag: null, retryAfterMs: 0 };
  }
}
