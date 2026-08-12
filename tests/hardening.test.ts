import { suite, test, assert } from './harness';
import { resetBrowserStorage } from './fakes/browserShim';
import { FakeSharePoint } from './fakes/FakeSharePoint';
import { SpRest } from '../src/services/sp/SpRest';
import { SpProvisioningService } from '../src/services/sp/SpProvisioningService';
import { DocumentGraphStore } from '../src/services/sp/DocumentGraphStore';
import { readListRights } from '../src/services/sp/permissions';
import { schemaFingerprint, EXPECTED_SCHEMA, PROJECTS_LIST, PRESENCE_LIST } from '../src/provisioning/schema';
import { IGraph, IGraphNode, normalizeGraph, CHUNK_SIZE } from '../src/model/bundle';

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
