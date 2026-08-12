import { suite, test, assert } from './harness';
import { resetBrowserStorage } from './fakes/browserShim';
import { FakeSharePoint } from './fakes/FakeSharePoint';
import { SpRest } from '../src/services/sp/SpRest';
import { SpProvisioningService } from '../src/services/sp/SpProvisioningService';
import { DocumentGraphStore } from '../src/services/sp/DocumentGraphStore';
import { readListRights } from '../src/services/sp/permissions';
import {
  schemaFingerprint, EXPECTED_SCHEMA, PROJECTS_LIST, PRESENCE_LIST,
  SNAPSHOTS_LIST, NODES_LIST, EDGES_LIST
} from '../src/provisioning/schema';
import { ItemGraphStore } from '../src/services/sp/ItemGraphStore';
import { convertProjectStorage } from '../src/services/sp/convertStorage';
import { IGraph, IGraphNode, normalizeGraph, CHUNK_SIZE } from '../src/model/bundle';
import { ICON_CSS } from '../src/engine/iconAssets';

const node = (id: string, label: string): IGraphNode => ({ id, label, type: 'Office' });
const graph = (nodes: IGraphNode[]): IGraph =>
  normalizeGraph({ nodes, edges: [] } as unknown as IGraph);

const deployed = async (): Promise<{ fake: FakeSharePoint; sp: SpRest; prov: SpProvisioningService }> => {
  resetBrowserStorage();
  const fake = new FakeSharePoint();
  const sp = new SpRest(fake);
  const prov = new SpProvisioningService(sp);
  await prov.execute(await prov.buildPlan());
  return { fake, sp, prov };
};

suite('hardening: the schema check does not run on every load forever', () => {
  test('a healthy verdict is remembered, and skips the whole check', async () => {
    const { fake, prov } = await deployed();
    await prov.buildPlan();
    assert.ok(prov.wasVerifiedHealthy(), 'a clean plan should be worth remembering');

    fake.reset();
    // The caller checks this BEFORE spending anything.
    assert.ok(prov.wasVerifiedHealthy());
    assert.equal(fake.log.length, 0, 'answering from the cache must cost no requests');
  });

  test('an incomplete site is never cached, so the warning cannot get stuck', async () => {
    const { fake, prov } = await deployed();
    await prov.buildPlan();
    assert.ok(prov.wasVerifiedHealthy());

    fake.lists.delete(PRESENCE_LIST);
    await prov.buildPlan();
    assert.ok(!prov.wasVerifiedHealthy(), 'a gap must re-check every load until it is fixed');
  });

  test('the check still costs 15 calls when it runs — it is the DEPTH that changed', async () => {
    const { fake, prov } = await deployed();
    fake.reset();
    await prov.buildPlan();
    // 5 lists x (list + fields + view). Volume was never the problem; SharePoint
    // throttles sustained operations, not a handful of reads on page load. What made
    // it hurt was doing them one after another.
    assert.equal(fake.log.length, 15);
  });
});

suite('hardening: a new build re-checks automatically', () => {
  test('the fingerprint is stable for the same schema', () => {
    assert.equal(schemaFingerprint(EXPECTED_SCHEMA), schemaFingerprint(EXPECTED_SCHEMA));
  });

  test('adding a column changes it', () => {
    const changed = EXPECTED_SCHEMA.map((l) => (l.title !== PROJECTS_LIST ? l : {
      ...l, fields: l.fields.concat([{ internal: 'ErgSomethingNew', display: 'New', types: ['Text' as const] }])
    }));
    assert.ok(schemaFingerprint(changed) !== schemaFingerprint(EXPECTED_SCHEMA));
  });

  test('adding a whole list changes it', () => {
    const changed = EXPECTED_SCHEMA.concat([
      { title: 'ERG Future', description: '', versioning: false, fields: [] }
    ]);
    assert.ok(schemaFingerprint(changed) !== schemaFingerprint(EXPECTED_SCHEMA));
  });

  test('a cached verdict from a DIFFERENT schema does not count', async () => {
    // This is what makes an upgraded .sppkg re-check on its first load without anyone
    // remembering to bump a version: the key contains the schema's own fingerprint.
    const { prov } = await deployed();
    await prov.buildPlan();
    assert.ok(prov.wasVerifiedHealthy());

    // Simulate the next release by writing a verdict under a different fingerprint.
    resetBrowserStorage();
    window.localStorage.setItem(
      'erg.schemaOk.https://contoso.sharepoint.us/sites/exec.someoldhash', String(Date.now())
    );
    assert.ok(!prov.wasVerifiedHealthy(), 'a verdict for another schema must not be reused');
  });
});

suite('hardening: list-level permissions, not just site-level', () => {
  test('full control reads as editable', async () => {
    const { fake, sp } = await deployed();
    fake.listRights.delete(PROJECTS_LIST);
    const rights = await readListRights(sp, PROJECTS_LIST);
    assert.ok(rights);
    assert.ok(rights!.canEdit);
    assert.ok(rights!.canManage);
  });

  test('view-only on the list reads as NOT editable, whatever the site says', async () => {
    // ViewListItems only (bit 1). Someone who can edit the site but not this list would
    // otherwise be handed a fully editable graph that fails on their first save.
    const { fake, sp } = await deployed();
    fake.listRights.set(PROJECTS_LIST, { High: 0, Low: 1 << 1 });
    const rights = await readListRights(sp, PROJECTS_LIST);
    assert.ok(rights);
    assert.ok(!rights!.canEdit);
    assert.ok(!rights!.canManage);
  });

  test('add and edit without manage reads as editable but not an owner', async () => {
    const { fake, sp } = await deployed();
    fake.listRights.set(PROJECTS_LIST, { High: 0, Low: (1 << 1) | (1 << 2) | (1 << 3) });
    const rights = await readListRights(sp, PROJECTS_LIST);
    assert.ok(rights!.canEdit);
    assert.ok(!rights!.canManage, 'they can edit graphs but should not be offered Deploy');
  });

  test('an unreadable list returns null so the caller can fall back', async () => {
    const { sp } = await deployed();
    assert.equal(await readListRights(sp, 'No Such List'), null);
  });
});

suite('hardening: an idle graph writes nothing', () => {
  test('saving an unchanged graph performs no write at all', async () => {
    // The engine persists for SYNTHETIC mutations too — a project open, an undo, a
    // boot cleanup. Each of those used to become a real write, which bumped Modified,
    // which every other client read as "somebody edited this" and answered with a
    // write of its own: a loop between browsers with nobody editing. It showed up as
    // colleagues who appeared to be saving constantly while sitting still.
    const { fake, sp } = await deployed();
    const store = new DocumentGraphStore(sp, 'Ross');
    const project = await store.createProject('Quiet', 'Document');
    const session = await store.openProject(project.id);
    await session.save(graph([node('a', 'A')]), []);

    fake.reset();
    const outcome = await session.save(graph([node('a', 'A')]), []);
    assert.equal(outcome.writes, 0);
    assert.equal(fake.log.length, 0, 'an unchanged save must not touch the network');
  });

  test('re-saving what was just opened is a no-op', async () => {
    const { fake, sp } = await deployed();
    const store = new DocumentGraphStore(sp, 'Ross');
    const project = await store.createProject('Reopen', 'Document');
    const first = await store.openProject(project.id);
    await first.save(graph([node('a', 'A'), node('b', 'B')]), []);

    const reopened = await store.openProject(project.id);
    fake.reset();
    const outcome = await reopened.save(reopened.bundle.graph, reopened.bundle.snapshots || []);
    assert.equal(outcome.writes, 0, 'opening a project must not rewrite it');
    assert.equal(fake.log.length, 0);
  });

  test('a real change still writes, and key order alone is not a change', async () => {
    const { sp } = await deployed();
    const store = new DocumentGraphStore(sp, 'Ross');
    const project = await store.createProject('Real', 'Document');
    const session = await store.openProject(project.id);
    await session.save(graph([node('a', 'A')]), []);

    // Same node, different key insertion order.
    const reordered = normalizeGraph({
      nodes: [{ type: 'Office', label: 'A', id: 'a' }], edges: []
    } as unknown as IGraph);
    assert.equal((await session.save(reordered, [])).writes, 0, 'key order is not a change');
    assert.equal((await session.save(graph([node('a', 'Renamed')]), [])).writes, 1);
  });

  test('a snapshot-only change still counts as a change', async () => {
    const { sp } = await deployed();
    const store = new DocumentGraphStore(sp, 'Ross');
    const project = await store.createProject('Snaps', 'Document');
    const session = await store.openProject(project.id);
    await session.save(graph([node('a', 'A')]), []);
    const outcome = await session.save(graph([node('a', 'A')]), [{ id: 's1', name: 'Baseline' }]);
    assert.equal(outcome.writes, 1, 'the graph is unchanged but the snapshots are not');
  });
});

suite('hardening: snapshots survive a change of storage model', () => {
  const snap = (id: string, name: string): { id: string; name: string; ts: number; graph: unknown } =>
    ({ id, name, ts: 1000, graph: { nodes: [node('a', 'A')], edges: [] } });

  test('they are rows in their own list, not passengers in the payload', async () => {
    // Each snapshot is a FULL copy of the graph. Carried inside the document payload,
    // a handful of them could exhaust the ~480KB a list item holds — and per-item
    // projects dropped them entirely, so changing storage model destroyed them.
    const { fake, sp } = await deployed();
    const store = new DocumentGraphStore(sp, 'Ross');
    const project = await store.createProject('Snaps', 'Document');
    const session = await store.openProject(project.id);
    await session.save(graph([node('a', 'A')]), [snap('s1', 'Before reorg'), snap('s2', 'After')]);

    assert.equal(fake.lists.get(SNAPSHOTS_LIST)!.items.size, 2, 'each snapshot is its own row');

    const payload = String(
      Array.from(fake.lists.get(PROJECTS_LIST)!.items.values())[0].data.ErgPayload1 || ''
    );
    assert.ok(payload.indexOf('Before reorg') < 0, 'and no longer duplicated in the payload');

    const reopened = await store.openProject(project.id);
    assert.equal((reopened.bundle.snapshots as unknown[]).length, 2);
  });

  test('a snapshot with no id is kept rather than silently dropped', async () => {
    const { fake, sp } = await deployed();
    const store = new DocumentGraphStore(sp, 'Ross');
    const project = await store.createProject('NoId', 'Document');
    const session = await store.openProject(project.id);
    await session.save(graph([node('a', 'A')]), [{ name: 'baseline' }]);
    assert.equal(fake.lists.get(SNAPSHOTS_LIST)!.items.size, 1);
  });

  test('converting to per-item storage keeps the graph AND the snapshots', async () => {
    const { fake, sp } = await deployed();
    const store = new DocumentGraphStore(sp, 'Ross');
    const project = await store.createProject('Convert', 'Document');
    const session = await store.openProject(project.id);
    const original = graph([node('a', 'A'), node('b', 'B')]);
    original.edges = [{ id: 'e1', source: 'a', target: 'b', type: 'CONTAINS' }];
    original.positions = { a: { x: 5, y: 6 } };
    await session.save(original, [snap('s1', 'Before the move')]);

    const result = await convertProjectStorage(sp, 'Ross', project.id, 'Items');
    assert.equal(result.nodes, 2);
    assert.equal(result.edges, 1);

    assert.equal(fake.lists.get(NODES_LIST)!.items.size, 2, 'nodes are rows now');
    assert.equal(fake.lists.get(EDGES_LIST)!.items.size, 1);
    assert.equal(fake.lists.get(SNAPSHOTS_LIST)!.items.size, 1, 'restore points survive');

    // The old copy must not linger, or two sources of truth exist for one graph.
    const projectRow = Array.from(fake.lists.get(PROJECTS_LIST)!.items.values())[0];
    assert.equal(String(projectRow.data.ErgPayload1 || ''), '', 'the payload is cleared');
    assert.equal(projectRow.data.ErgStorageMode, 'Items');

    const items = new ItemGraphStore(sp, 'Ross');
    const reopened = await items.openProject(project.id);
    assert.deepEqual(reopened.bundle.graph.nodes.map((n) => n.id).sort(), ['a', 'b']);
    assert.deepEqual(reopened.bundle.graph.positions, { a: { x: 5, y: 6 } }, 'layout comes too');
    assert.equal((reopened.bundle.snapshots as unknown[]).length, 1);
  });

  test('converting back to a document leaves no orphan rows', async () => {
    const { fake, sp } = await deployed();
    const items = new ItemGraphStore(sp, 'Ross');
    const project = await items.createProject('Back', 'Items');
    const session = await items.openProject(project.id);
    const g = graph([node('a', 'A'), node('b', 'B')]);
    g.edges = [{ id: 'e1', source: 'a', target: 'b', type: 'CONTAINS' }];
    await session.save(g, [snap('s1', 'Keep me')]);

    await convertProjectStorage(sp, 'Ross', project.id, 'Document');

    assert.equal(fake.lists.get(NODES_LIST)!.items.size, 0, 'node rows are cleaned up');
    assert.equal(fake.lists.get(EDGES_LIST)!.items.size, 0);
    assert.equal(fake.lists.get(SNAPSHOTS_LIST)!.items.size, 1, 'but snapshots are not touched');

    const store = new DocumentGraphStore(sp, 'Ross');
    const reopened = await store.openProject(project.id);
    assert.deepEqual(reopened.bundle.graph.nodes.map((n) => n.id).sort(), ['a', 'b']);
    assert.equal(reopened.bundle.graph.edges.length, 1);
    assert.equal((reopened.bundle.snapshots as unknown[]).length, 1);
  });
});

suite('hardening: the document ceiling is announced before it is hit', () => {
  test('a small graph says nothing', async () => {
    const { sp } = await deployed();
    const store = new DocumentGraphStore(sp, 'Ross');
    const project = await store.createProject('Small', 'Document');
    const session = await store.openProject(project.id);
    const outcome = await session.save(graph([node('a', 'A')]), []);
    assert.equal(outcome.warning, undefined);
  });

  test('a graph past three quarters of the ceiling warns while there is still room', async () => {
    const { sp } = await deployed();
    const store = new DocumentGraphStore(sp, 'Ross');
    const project = await store.createProject('Big', 'Document');
    const session = await store.openProject(project.id);

    const filler = 'x'.repeat(500);
    const nodes: IGraphNode[] = [];
    // ~80% of 8 x 60,000 characters.
    for (let i = 0; i < Math.ceil((CHUNK_SIZE * 8 * 0.8) / 560); i++) {
      nodes.push({ ...node('n' + i, 'Node ' + i), notes: filler });
    }
    const outcome = await session.save(graph(nodes), []);

    assert.equal(outcome.status, 'saved', 'it must still save — this is advice, not a failure');
    assert.ok(outcome.warning, 'the operator should hear about it before saves start failing');
    assert.includes(outcome.warning as string, 'per-item storage');
  });
});

suite('hardening: DOM icons survive the build', () => {
  test('the class names are UNMANGLED — they are the icon library API', () => {
    // SPFx runs css-loader with CSS Modules over imported stylesheets, which hashes
    // class names. For a component's own styles that is right; for Font Awesome it
    // renamed `.fa-solid` to `.fa-solid_f1fd2f8f` and every <i class="fa-solid fa-user">
    // in the engine's markup matched nothing. The failure was asymmetric and therefore
    // confusing: @font-face is not a class, so canvas glyphs kept working while every
    // icon in the DOM became an empty box.
    assert.includes(ICON_CSS, '.fa-solid');
    assert.ok(!/\.fa-solid_[0-9a-f]{6}/.test(ICON_CSS), 'a hashed class name means modules ran over it');
  });

  test('both font families are assigned to their classes', () => {
    assert.includes(ICON_CSS, 'font-family:"Font Awesome 6 Free"');
    assert.includes(ICON_CSS, 'font-family:"Font Awesome 6 Brands"');
  });

  test('glyph mappings are present and singly escaped', () => {
    // CSS needs a LITERAL backslash; the CSS parser does the unescaping, not
    // JavaScript. Double-escaping would render every icon as literal text.
    const bs = String.fromCharCode(92);
    assert.includes(ICON_CSS, `.fa-user:before{content:"${bs}f007"}`);
    assert.ok(
      ICON_CSS.indexOf(bs + bs + 'f007') < 0,
      'double-escaped glyphs would print as text rather than draw an icon'
    );
  });

  test('no url() survives, because a string constant cannot resolve one', () => {
    // @font-face is deliberately stripped: the stylesheet IMPORT supplies it, with URLs
    // webpack rewrote to the emitted woff2 files.
    assert.ok(ICON_CSS.indexOf('url(') < 0);
    assert.ok(ICON_CSS.indexOf('@font-face') < 0);
  });
});
