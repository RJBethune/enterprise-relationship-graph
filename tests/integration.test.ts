import { suite, test, assert } from './harness';
import { FakeSharePoint } from './fakes/FakeSharePoint';
import { SpRest } from '../src/services/sp/SpRest';
import { SpProvisioningService } from '../src/services/sp/SpProvisioningService';
import { DocumentGraphStore } from '../src/services/sp/DocumentGraphStore';
import { ItemGraphStore } from '../src/services/sp/ItemGraphStore';
import { isHealthy } from '../src/provisioning/planner';
import { IGraph, IGraphNode, normalizeGraph, CHUNK_SIZE } from '../src/model/bundle';
import { PresenceService } from '../src/services/sp/PresenceService';
import {
  NODES_LIST, EDGES_LIST, PROJECTS_LIST, PRESENCE_LIST, CORE_LISTS
} from '../src/provisioning/schema';

const node = (id: string, label: string): IGraphNode => ({ id, label, type: 'Office' });

const graph = (nodes: IGraphNode[], edges: IGraph['edges'] = []): IGraph =>
  normalizeGraph({ nodes, edges } as unknown as IGraph);

/** A site with the schema already deployed. */
const deployedSite = async (): Promise<{ fake: FakeSharePoint; sp: SpRest }> => {
  const fake = new FakeSharePoint();
  const sp = new SpRest(fake);
  const provisioning = new SpProvisioningService(sp);
  const plan = await provisioning.buildPlan();
  await provisioning.execute(plan);
  return { fake, sp };
};

suite('integration: reading a live schema', () => {
  test('an existing list can be READ — the health check must not 400 on its own query', async () => {
    // Regression: the fields query named Choices and RelationshipDeleteBehavior in
    // $select. /fields is polymorphic and declared as SP.Field, which has neither, so
    // SharePoint rejected the WHOLE query with 400. Every list then reported "exists but
    // could not be read", the planner refused to act, and setup was dead in the water
    // with four conflicts and zero actions — looking exactly like a permissions problem.
    const { sp } = await deployedSite();
    const provisioning = new SpProvisioningService(sp);

    const snapshots = await provisioning.readSnapshots();
    for (const title of [PROJECTS_LIST, NODES_LIST, EDGES_LIST]) {
      const snap = snapshots[title];
      assert.ok(snap.exists, `${title} should exist after provisioning`);
      assert.ok(snap.verified, `${title} must be READABLE — SharePoint said: ${snap.readError || 'n/a'}`);
      assert.ok(Object.keys(snap.fields).length > 1, `${title} should report its columns`);
    }
  });

  test('the exact query shape that shipped broken is rejected, so this cannot regress', async () => {
    // Proves the guard above is live rather than decorative: if the fake accepted this,
    // the regression test would pass no matter what the production query did.
    const { sp } = await deployedSite();
    let status = 0;
    try {
      await sp.get(
        `web/lists/getbytitle('${encodeURIComponent(PROJECTS_LIST)}')/fields` +
        '?$select=InternalName,TypeAsString,Indexed,EnforceUniqueValues,Choices,RelationshipDeleteBehavior' +
        '&$top=500'
      );
    } catch (e) {
      status = (e as { status?: number }).status || 0;
    }
    assert.equal(status, 400, 'a subtype-only property in $select must 400, exactly as SharePoint does');
  });

  test('subtype-only properties still arrive, so choices and cascade are checkable', async () => {
    const { sp } = await deployedSite();
    const snapshots = await new SpProvisioningService(sp).readSnapshots();

    const status = snapshots[PROJECTS_LIST].fields.ErgStatus;
    assert.ok(status, 'ErgStatus should be present');
    assert.deepEqual(status.choices, ['Active', 'Archived'], 'Choice values must be readable');

    const lookup = snapshots[NODES_LIST].fields.ErgProject;
    assert.ok(lookup, 'ErgProject lookup should be present');
    assert.equal(lookup.relationshipDeleteBehavior, 1, 'cascade behaviour must be readable');
  });

  test('an unreadable list reports what SharePoint actually said', async () => {
    // A generic "could not be read" is what made the original bug so hard to place.
    const fake = new FakeSharePoint();
    const sp = new SpRest(fake);
    const provisioning = new SpProvisioningService(sp);
    await provisioning.execute(await provisioning.buildPlan());

    const original = fake.request.bind(fake);
    fake.request = async (method, url, opts) => {
      if (url.indexOf('/fields') >= 0 && method === 'GET') {
        return { status: 403, ok: false, etag: null, retryAfterMs: 0,
          text: '{"odata.error":{"message":{"value":"Access denied."}}}',
          json: { 'odata.error': { message: { value: 'Access denied.' } } } };
      }
      return original(method, url, opts);
    };

    const plan = await provisioning.buildPlan();
    assert.ok(plan.conflicts.length > 0);
    assert.includes(plan.conflicts[0].reason, 'Access denied.');
  });
});

suite('integration: provisioning a real site', () => {
  test('an empty site is fully deployed by one press of the button', async () => {
    const fake = new FakeSharePoint();
    const sp = new SpRest(fake);
    const provisioning = new SpProvisioningService(sp);

    const plan = await provisioning.buildPlan();
    assert.ok(plan.actions.length > 0);
    assert.equal(plan.conflicts.length, 0);

    const results = await provisioning.execute(plan);
    const failed = results.filter((r) => !r.ok);
    assert.equal(failed.length, 0, failed.map((f) => `${f.label}: ${f.error}`).join('; '));

    // The health check and the fix read the same declaration, so a successful run must
    // leave the site with an empty plan. If it does not, they have drifted.
    const after = await provisioning.buildPlan();
    assert.ok(isHealthy(after), `expected a healthy site, still planned: ${after.actions.length} actions`);
  });

  test('pressing the button twice changes nothing the second time', async () => {
    const { sp, fake } = await deployedSite();
    const provisioning = new SpProvisioningService(sp);
    fake.reset();
    const plan = await provisioning.buildPlan();
    await provisioning.execute(plan);
    assert.equal(fake.writeCount(), 0, 'a healthy site must produce no writes');
  });

  test('all four lists exist with versioning where declared', async () => {
    const { fake } = await deployedSite();
    assert.ok(fake.lists.has(PROJECTS_LIST));
    assert.ok(fake.lists.has(NODES_LIST));
    assert.ok(fake.lists.has(EDGES_LIST));
    assert.equal(fake.lists.get(PROJECTS_LIST)!.versioning, true, 'projects need version history');
  });

  test('project names are enforced unique, so links by name stay unambiguous', async () => {
    const { sp } = await deployedSite();
    const store = new DocumentGraphStore(sp, 'Ross');
    await store.createProject('Executive Office', 'Document');
    let rejected = false;
    try { await store.createProject('Executive Office', 'Document'); } catch (_e) { rejected = true; }
    assert.ok(rejected, 'a duplicate project name must be rejected by the list, not just the UI');
  });
});

suite('integration: document storage', () => {
  test('a graph round-trips through SharePoint unchanged', async () => {
    const { sp } = await deployedSite();
    const store = new DocumentGraphStore(sp, 'Ross');
    const project = await store.createProject('Exec', 'Document');

    const session = await store.openProject(project.id);
    const g = graph(
      [node('a', 'Executive Office'), node('b', 'Data & Analytics')],
      [{ id: 'e1', source: 'a', target: 'b', type: 'CONTAINS' }]
    );
    g.positions = { a: { x: 10, y: 20 } };
    g.customNodeTypes = [{ name: 'Programme', color: '#6366F1' }];

    const outcome = await session.save(g, [{ name: 'baseline' }]);
    assert.equal(outcome.status, 'saved');

    const reopened = await store.openProject(project.id);
    assert.equal(reopened.bundle.graph.nodes.length, 2);
    assert.equal(reopened.bundle.graph.edges.length, 1);
    assert.deepEqual(reopened.bundle.graph.positions, { a: { x: 10, y: 20 } });
    assert.equal((reopened.bundle.graph.customNodeTypes as unknown[]).length, 1);
    assert.equal((reopened.bundle.snapshots as unknown[]).length, 1);
  });

  test('a whole-graph save is ONE write, whatever changed', async () => {
    const { sp, fake } = await deployedSite();
    const store = new DocumentGraphStore(sp, 'Ross');
    const project = await store.createProject('Exec', 'Document');
    const session = await store.openProject(project.id);

    const many: IGraphNode[] = [];
    for (let i = 0; i < 400; i++) { many.push(node('n' + i, 'Node ' + i)); }

    fake.reset();
    await session.save(graph(many), []);
    assert.equal(fake.writeCount(), 1, '400 nodes still cost exactly one operation');
  });

  test('a graph that shrinks does not leave stale payload behind', async () => {
    const { sp } = await deployedSite();
    const store = new DocumentGraphStore(sp, 'Ross');
    const project = await store.createProject('Exec', 'Document');
    const session = await store.openProject(project.id);

    // Big enough to span several payload columns, then tiny.
    const big: IGraphNode[] = [];
    const filler = 'x'.repeat(200);
    for (let i = 0; i < Math.ceil((CHUNK_SIZE * 2) / 240); i++) {
      big.push({ ...node('n' + i, 'Node ' + i), notes: filler });
    }
    await session.save(graph(big), []);
    await session.save(graph([node('only', 'Only one left')]), []);

    const reopened = await store.openProject(project.id);
    assert.equal(reopened.bundle.graph.nodes.length, 1, 'stale chunks would make this unparseable');
  });

  test('two people editing different nodes both keep their work', async () => {
    const { sp } = await deployedSite();
    const store = new DocumentGraphStore(sp, 'Ross');
    const project = await store.createProject('Exec', 'Document');

    const setup = await store.openProject(project.id);
    await setup.save(graph([node('a', 'A'), node('b', 'B')]), []);

    // Two browsers, both opened before either saved.
    const alice = await store.openProject(project.id);
    const bob = await store.openProject(project.id);

    await alice.save(graph([node('a', 'Alice edited'), node('b', 'B')]), []);
    const bobOutcome = await bob.save(graph([node('a', 'A'), node('b', 'Bob edited')]), []);

    assert.equal(bobOutcome.status, 'merged', 'a losing write must merge, not overwrite');
    assert.equal((bobOutcome.conflicts || []).length, 0, 'different nodes are not a conflict');

    const final = await store.openProject(project.id);
    const labels = new Map(final.bundle.graph.nodes.map((n) => [n.id, n.label]));
    assert.equal(labels.get('a'), 'Alice edited', "Alice's edit survived Bob's save");
    assert.equal(labels.get('b'), 'Bob edited');
  });

  test('editing the SAME node is reported rather than silently resolved', async () => {
    const { sp } = await deployedSite();
    const store = new DocumentGraphStore(sp, 'Ross');
    const project = await store.createProject('Exec', 'Document');
    const setup = await store.openProject(project.id);
    await setup.save(graph([node('a', 'A')]), []);

    const alice = await store.openProject(project.id);
    const bob = await store.openProject(project.id);
    await alice.save(graph([node('a', 'Alice')]), []);
    const bobOutcome = await bob.save(graph([node('a', 'Bob')]), []);

    assert.equal(bobOutcome.status, 'merged');
    assert.equal((bobOutcome.conflicts || []).length, 1);
    assert.equal((bobOutcome.conflicts || [])[0].reason, 'both-edited');
    const final = await store.openProject(project.id);
    assert.equal(final.bundle.graph.nodes[0].label, 'Bob', "the saver's version is kept, and they are told");
  });

  test("polling surfaces another editor's change", async () => {
    const { sp } = await deployedSite();
    const store = new DocumentGraphStore(sp, 'Ross');
    const project = await store.createProject('Exec', 'Document');
    const setup = await store.openProject(project.id);
    await setup.save(graph([node('a', 'A')]), []);

    const viewer = await store.openProject(project.id);
    assert.equal(await viewer.poll(viewer.bundle.graph), null, 'a quiet graph must not churn');

    const editor = await store.openProject(project.id);
    await editor.save(graph([node('a', 'A'), node('z', 'Added elsewhere')]), []);

    const update = await viewer.poll(viewer.bundle.graph);
    assert.ok(update, 'the viewer should see the change');
    assert.equal(update!.graph.nodes.length, 2);
    assert.equal(update!.conflicts.length, 0);
  });
});

suite('integration: per-item storage', () => {
  test('the first save creates one row per node and relationship', async () => {
    const { sp, fake } = await deployedSite();
    const store = new ItemGraphStore(sp, 'Ross');
    const project = await store.createProject('Live', 'Items');
    const session = await store.openProject(project.id);

    await session.save(graph(
      [node('a', 'A'), node('b', 'B')],
      [{ id: 'e1', source: 'a', target: 'b', type: 'CONTAINS' }]
    ), []);

    assert.equal(fake.lists.get(NODES_LIST)!.items.size, 2);
    assert.equal(fake.lists.get(EDGES_LIST)!.items.size, 1);
  });

  test('renaming one node writes exactly one row — not the whole graph', async () => {
    const { sp, fake } = await deployedSite();
    const store = new ItemGraphStore(sp, 'Ross');
    const project = await store.createProject('Live', 'Items');
    const session = await store.openProject(project.id);

    const nodes: IGraphNode[] = [];
    for (let i = 0; i < 50; i++) { nodes.push(node('n' + i, 'Node ' + i)); }
    await session.save(graph(nodes), []);

    fake.reset();
    const renamed = nodes.map((n) => (n.id === 'n7' ? node('n7', 'Renamed') : n));
    await session.save(graph(renamed), []);

    // One item MERGE, plus the project blob that carries counts and layout.
    assert.equal(fake.writeCount(), 2, 'a rename in a 50-node graph must not rewrite 50 rows');
  });

  test('moving every node writes the layout blob only — no item writes at all', async () => {
    // This is the decision that makes per-item storage viable. A force layout touches
    // every node; as item writes that is one operation per node per run.
    const { sp, fake } = await deployedSite();
    const store = new ItemGraphStore(sp, 'Ross');
    const project = await store.createProject('Live', 'Items');
    const session = await store.openProject(project.id);

    const nodes: IGraphNode[] = [];
    for (let i = 0; i < 50; i++) { nodes.push(node('n' + i, 'Node ' + i)); }
    await session.save(graph(nodes), []);

    fake.reset();
    const moved = graph(nodes);
    moved.positions = {};
    for (let i = 0; i < 50; i++) { moved.positions['n' + i] = { x: i * 3, y: i * 7 }; }
    const outcome = await session.save(moved, []);

    assert.equal(outcome.status, 'saved');
    assert.equal(fake.writeCount(), 1, 'a full re-layout of 50 nodes costs one write');
  });

  test('an unchanged save costs nothing', async () => {
    const { sp, fake } = await deployedSite();
    const store = new ItemGraphStore(sp, 'Ross');
    const project = await store.createProject('Live', 'Items');
    const session = await store.openProject(project.id);
    const g = graph([node('a', 'A')]);
    await session.save(g, []);

    fake.reset();
    const outcome = await session.save(graph([node('a', 'A')]), []);
    assert.equal(outcome.writes, 0);
    assert.equal(fake.writeCount(), 0);
  });

  test('a re-opened project sees exactly what was written', async () => {
    const { sp } = await deployedSite();
    const store = new ItemGraphStore(sp, 'Ross');
    const project = await store.createProject('Live', 'Items');
    const session = await store.openProject(project.id);

    const g = graph(
      [node('a', 'A'), node('b', 'B')],
      [{ id: 'e1', source: 'a', target: 'b', type: 'CONTAINS' }]
    );
    g.positions = { a: { x: 1, y: 2 } };
    await session.save(g, []);

    const reopened = await store.openProject(project.id);
    assert.deepEqual(reopened.bundle.graph.nodes.map((n) => n.id).sort(), ['a', 'b']);
    assert.equal(reopened.bundle.graph.edges.length, 1);
    assert.deepEqual(reopened.bundle.graph.positions, { a: { x: 1, y: 2 } });
  });

  test('editing a node twice updates its row instead of creating a duplicate', async () => {
    // Depends on capturing the Id SharePoint assigns during a batched create. This is
    // the exact failure the batch-response parser bug would have caused in production.
    const { sp, fake } = await deployedSite();
    const store = new ItemGraphStore(sp, 'Ross');
    const project = await store.createProject('Live', 'Items');
    const session = await store.openProject(project.id);

    await session.save(graph([node('a', 'First')]), []);
    await session.save(graph([node('a', 'Second')]), []);

    assert.equal(fake.lists.get(NODES_LIST)!.items.size, 1, 'a second edit must not create a second row');
    const only = Array.from(fake.lists.get(NODES_LIST)!.items.values())[0];
    assert.includes(String(only.data.ErgData), 'Second');
  });

  test('deleting a node removes its row and its relationship rows', async () => {
    const { sp, fake } = await deployedSite();
    const store = new ItemGraphStore(sp, 'Ross');
    const project = await store.createProject('Live', 'Items');
    const session = await store.openProject(project.id);

    await session.save(graph(
      [node('a', 'A'), node('b', 'B')],
      [{ id: 'e1', source: 'a', target: 'b', type: 'CONTAINS' }]
    ), []);
    await session.save(graph([node('a', 'A')]), []);

    assert.equal(fake.lists.get(NODES_LIST)!.items.size, 1);
    assert.equal(fake.lists.get(EDGES_LIST)!.items.size, 0, 'the relationship must go with the node');
  });

  test("the change log surfaces another session's add, including into a busy local graph", async () => {
    const { sp } = await deployedSite();
    const store = new ItemGraphStore(sp, 'Ross');
    const project = await store.createProject('Live', 'Items');

    const alice = await store.openProject(project.id);
    await alice.save(graph([node('a', 'A')]), []);

    const bob = await store.openProject(project.id);
    await bob.save(graph([node('a', 'A'), node('z', 'Bob added')]), []);

    // Alice has unsaved local work at the same time.
    const aliceLocal = graph([node('a', 'A'), node('local', 'Alice drafting')]);
    const update = await alice.poll(aliceLocal);

    assert.ok(update, 'the change log should report the new row');
    const ids = update!.graph.nodes.map((n) => n.id).sort();
    assert.deepEqual(ids, ['a', 'local', 'z'], "both editors' work survives");
    assert.equal(update!.conflicts.length, 0);
  });

  test("the change log surfaces another session's delete", async () => {
    // A `Modified gt` query cannot see deletes at all — this is why the sync uses
    // GetChanges rather than a timestamp scan.
    const { sp } = await deployedSite();
    const store = new ItemGraphStore(sp, 'Ross');
    const project = await store.createProject('Live', 'Items');

    const alice = await store.openProject(project.id);
    await alice.save(graph([node('a', 'A'), node('b', 'B')]), []);

    const bob = await store.openProject(project.id);
    await bob.save(graph([node('a', 'A')]), []);

    // What Alice currently has on screen — the shell passes the live graph, not the
    // bundle captured when the project was opened.
    const aliceOnScreen = graph([node('a', 'A'), node('b', 'B')]);
    const update = await alice.poll(aliceOnScreen);
    assert.ok(update, 'the delete must be reported');
    assert.deepEqual(update!.graph.nodes.map((n) => n.id), ['a']);
  });

  test('a quiet graph produces no update', async () => {
    const { sp } = await deployedSite();
    const store = new ItemGraphStore(sp, 'Ross');
    const project = await store.createProject('Live', 'Items');
    const session = await store.openProject(project.id);
    await session.save(graph([node('a', 'A')]), []);
    assert.equal(await session.poll(graph([node('a', 'A')])), null);
  });

  test('a throttled write is retried and still lands', async () => {
    const { sp, fake } = await deployedSite();
    const store = new ItemGraphStore(sp, 'Ross');
    const project = await store.createProject('Live', 'Items');
    const session = await store.openProject(project.id);

    fake.failWrites = 2;  // two 429s before the write succeeds
    const outcome = await session.save(graph([node('a', 'A')]), []);

    assert.equal(outcome.status, 'saved', 'throttling must be survivable, not fatal');
    assert.equal(fake.lists.get(NODES_LIST)!.items.size, 1);
  });
});

suite('integration: presence', () => {
  test('a heartbeat creates one row, and repeating it updates rather than duplicates', async () => {
    const { sp, fake } = await deployedSite();
    const presence = new PresenceService(sp, 'i:0#.f|m|ross@example.gov', 'Ross Bethune');
    const project = await new DocumentGraphStore(sp, 'Ross').createProject('Exec', 'Document');

    await presence.heartbeat(project.id, 'Viewing');
    await presence.heartbeat(project.id, 'Editing');
    await presence.heartbeat(project.id, 'Editing');

    assert.equal(fake.lists.get(PRESENCE_LIST)!.items.size, 1, 'one row per person per graph');
  });

  test('another session shows up, and you can see yourself', async () => {
    const { sp } = await deployedSite();
    const project = await new DocumentGraphStore(sp, 'Ross').createProject('Exec', 'Document');
    const alice = new PresenceService(sp, 'alice@example.gov', 'Alice Adams');
    const bob = new PresenceService(sp, 'bob@example.gov', 'Bob Brown');

    await alice.heartbeat(project.id, 'Editing');
    await bob.heartbeat(project.id, 'Viewing');

    const seenByAlice = await alice.list(project.id);
    assert.equal(seenByAlice.length, 2);
    const self = seenByAlice.filter((p) => p.isSelf)[0];
    assert.ok(self, 'you should see yourself, so an empty strip means nobody else is here');
    assert.equal(self.mode, 'Editing');
    assert.equal(seenByAlice.filter((p) => p.name === 'Bob Brown')[0].mode, 'Viewing');
  });

  test('a stale row is ignored — a closed laptop stops showing as present', async () => {
    const { sp } = await deployedSite();
    const project = await new DocumentGraphStore(sp, 'Ross').createProject('Exec', 'Document');
    const me = new PresenceService(sp, 'me@example.gov', 'Me');
    await me.heartbeat(project.id, 'Viewing');

    // A row from a session that vanished without saying goodbye.
    await sp.post(`web/lists/getbytitle('${encodeURIComponent(PRESENCE_LIST)}')/items`, {
      Title: `${project.id}|ghost@example.gov`,
      ErgProjectId: project.id,
      ErgUser: 'Ghost',
      ErgLogin: 'ghost@example.gov',
      ErgMode: 'Editing',
      ErgHeartbeat: new Date(Date.now() - 10 * 60 * 1000).toISOString()
    });

    const present = await me.list(project.id);
    assert.equal(present.length, 1);
    assert.equal(present[0].name, 'Me');
  });

  test('leaving removes the row immediately', async () => {
    const { sp, fake } = await deployedSite();
    const project = await new DocumentGraphStore(sp, 'Ross').createProject('Exec', 'Document');
    const me = new PresenceService(sp, 'me@example.gov', 'Me');
    await me.heartbeat(project.id, 'Viewing');
    await me.leave(project.id);
    assert.equal(fake.lists.get(PRESENCE_LIST)!.items.size, 0);
  });

  test('presence still works on a site whose list predates the Email column', async () => {
    // The reported symptom: no avatars at all. A site provisioned by an earlier build
    // has no ErgEmail, so BOTH the heartbeat write and the $select naming it were
    // rejected outright — presence showed nobody rather than degrading to initials.
    const { sp, fake } = await deployedSite();
    const project = await new DocumentGraphStore(sp, 'Ross').createProject('Exec', 'Document');

    // Roll the list back to the older schema.
    fake.lists.get(PRESENCE_LIST)!.fields.delete('ErgEmail');

    const me = new PresenceService(sp, 'me@example.gov', 'Me', 'me@example.gov');
    await me.heartbeat(project.id, 'Editing');

    assert.equal(fake.lists.get(PRESENCE_LIST)!.items.size, 1, 'the row must still be written');
    const present = await me.list(project.id);
    assert.equal(present.length, 1, 'and it must still be listed');
    assert.equal(present[0].name, 'Me');
    assert.equal(present[0].email, '', 'no photo, but the person is still shown');
    assert.ok(!me.isDisabled, 'a missing COLUMN must not switch presence off entirely');
  });

  test('a missing column and a missing list are told apart', async () => {
    const { sp, fake } = await deployedSite();
    const project = await new DocumentGraphStore(sp, 'Ross').createProject('Exec', 'Document');
    fake.lists.get(PRESENCE_LIST)!.fields.delete('ErgEmail');
    const me = new PresenceService(sp, 'me@example.gov', 'Me', 'me@example.gov');
    await me.heartbeat(project.id, 'Viewing');
    assert.ok(!me.isDisabled);

    // Now the list itself goes away.
    fake.lists.delete(PRESENCE_LIST);
    const other = new PresenceService(sp, 'me@example.gov', 'Me', 'me@example.gov');
    await other.heartbeat(project.id, 'Viewing');
    assert.ok(other.isDisabled, 'a missing LIST should stop the heartbeat');
  });

  test('presence never breaks the graph when its list is missing', async () => {
    // A site provisioned before presence existed must keep working, silently.
    const fake = new FakeSharePoint();
    const sp = new SpRest(fake);
    const presence = new PresenceService(sp, 'me@example.gov', 'Me');
    await presence.heartbeat(1, 'Viewing');
    assert.deepEqual(await presence.list(1), []);
    assert.ok(presence.isDisabled, 'it should stop asking rather than retry every beat');
  });
});

suite('integration: a site that is a schema version behind', () => {
  /** A site provisioned before the presence list existed. */
  const olderSite = async (): Promise<{ fake: FakeSharePoint; sp: SpRest }> => {
    const { fake, sp } = await deployedSite();
    fake.lists.delete(PRESENCE_LIST);
    return { fake, sp };
  };

  test('the gap is DETECTED rather than silently tolerated', async () => {
    // The reported failure: presence just showed nobody, with nothing anywhere saying
    // the site was missing a list. Silence is the bug; the missing list is fixable.
    const { sp } = await olderSite();
    const plan = await new SpProvisioningService(sp).buildPlan();

    assert.ok(!isHealthy(plan), 'an incomplete site must not report itself healthy');
    assert.ok(
      plan.actions.some((a) => a.kind === 'createList' && a.list === PRESENCE_LIST),
      'the plan must offer to create the missing list'
    );
    assert.equal(plan.conflicts.length, 0, 'a missing list needs no human — it just needs the button');
  });

  test('but the graph still opens — a missing extra must not block the core', async () => {
    const { sp } = await olderSite();
    const plan = await new SpProvisioningService(sp).buildPlan();
    const missingCore = plan.actions.some(
      (a) => a.kind === 'createList' && CORE_LISTS.indexOf(a.list) >= 0
    );
    assert.ok(!missingCore, 'nothing core is missing, so the app should warn rather than block');

    const store = new DocumentGraphStore(sp, 'Ross');
    const project = await store.createProject('Still works', 'Document');
    const session = await store.openProject(project.id);
    const outcome = await session.save(graph([node('a', 'A')]), []);
    assert.equal(outcome.status, 'saved');
  });

  test('pressing Deploy closes the gap and presence starts working', async () => {
    const { sp, fake } = await olderSite();
    const provisioning = new SpProvisioningService(sp);
    const presence = new PresenceService(sp, 'me@example.gov', 'Me', 'me@example.gov');
    const project = await new DocumentGraphStore(sp, 'Ross').createProject('Exec', 'Document');

    await presence.heartbeat(project.id, 'Viewing');
    assert.ok(presence.isDisabled, 'presence should switch off against a missing list');
    assert.deepEqual(await presence.list(project.id), []);

    const results = await provisioning.execute(await provisioning.buildPlan());
    assert.equal(results.filter((r) => !r.ok).length, 0);
    assert.ok(isHealthy(await provisioning.buildPlan()), 'the site should now be healthy');

    // Without reset(), presence would stay dead until a page reload.
    presence.reset();
    await presence.heartbeat(project.id, 'Editing');
    const present = await presence.list(project.id);
    assert.equal(present.length, 1, 'presence should come back without reloading the page');
    assert.equal(present[0].mode, 'Editing');
    assert.equal(fake.lists.get(PRESENCE_LIST)!.items.size, 1);
  });
});

suite('integration: the lists are actually readable by a human', () => {
  test('provisioned lists show their columns in the default view', async () => {
    // Columns created through the API land on NO view, so a correct save looks like a
    // broken one to anyone who opens the list to check.
    const { fake } = await deployedSite();
    const projects = fake.lists.get(PROJECTS_LIST)!;
    for (const column of ['ErgStorageMode', 'ErgNodeCount', 'ErgEdgeCount', 'ErgStatus']) {
      assert.ok(projects.viewFields.indexOf(column) >= 0, `${column} should be on the default view`);
    }
  });

  test('payload columns stay OFF the view — they are 60k of JSON each', async () => {
    const { fake } = await deployedSite();
    const projects = fake.lists.get(PROJECTS_LIST)!;
    assert.ok(projects.viewFields.indexOf('ErgPayload1') < 0);
  });
});

suite('integration: deleting a project', () => {
  test('cascade takes the node and relationship rows with it', async () => {
    const { sp, fake } = await deployedSite();
    const store = new ItemGraphStore(sp, 'Ross');
    const project = await store.createProject('Doomed', 'Items');
    const session = await store.openProject(project.id);
    await session.save(graph(
      [node('a', 'A'), node('b', 'B')],
      [{ id: 'e1', source: 'a', target: 'b', type: 'CONTAINS' }]
    ), []);

    assert.equal(fake.lists.get(NODES_LIST)!.items.size, 2);
    await sp.del(`web/lists/getbytitle('${encodeURIComponent(PROJECTS_LIST)}')/items(${project.id})`);

    assert.equal(fake.lists.get(NODES_LIST)!.items.size, 0, 'orphaned rows would accumulate forever');
    assert.equal(fake.lists.get(EDGES_LIST)!.items.size, 0);
  });
});
