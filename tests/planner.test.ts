import { suite, test, assert } from './harness';
import { planProvisioning, isHealthy, IListSnapshot } from '../src/provisioning/planner';
import { IExpectedList } from '../src/provisioning/schema';
import { EXPECTED_SCHEMA, PROJECTS_LIST } from '../src/provisioning/schema';
import { fieldSchemaXml } from '../src/provisioning/fieldXml';

const LIST: IExpectedList[] = [{
  title: 'Demo',
  description: 'demo list',
  versioning: true,
  fields: [
    { internal: 'Title', display: 'Name', types: ['Text'], builtIn: true, indexed: true, unique: true },
    { internal: 'DemoText', display: 'Text', types: ['Text'] },
    { internal: 'DemoStatus', display: 'Status', types: ['Choice'], choices: ['Active', 'Archived'] },
    { internal: 'DemoLink', display: 'Link', types: ['Lookup'], indexed: true, lookup: { list: 'Demo', behavior: 'Cascade' } }
  ]
}];

const snapshot = (over: Partial<IListSnapshot>): IListSnapshot => ({
  title: 'Demo', exists: true, verified: true, versioning: true, fields: {}, ...over
});

const field = (over: Partial<IListSnapshot['fields'][string]> & { internal: string; typeAsString: string }) => ({
  indexed: false, unique: false, ...over
});

suite('planner: a site with nothing', () => {
  test('plans the list first, so lookup targets exist before the lookups', () => {
    const plan = planProvisioning(LIST, {});
    assert.equal(plan.actions[0].kind, 'createList');
    const createIndex = plan.actions.findIndex((a) => a.kind === 'createList');
    const lookupIndex = plan.actions.findIndex(
      (a) => a.kind === 'createField' && a.field.internal === 'DemoLink'
    );
    assert.ok(createIndex < lookupIndex);
  });

  test('never tries to CREATE the built-in Title, but does apply its settings', () => {
    const plan = planProvisioning(LIST, {});
    const createsTitle = plan.actions.some((a) => a.kind === 'createField' && a.field.internal === 'Title');
    assert.ok(!createsTitle, 'Title exists on every list — creating it would fail the step');
    const setsUnique = plan.actions.some((a) => a.kind === 'setUnique' && a.internal === 'Title');
    assert.ok(setsUnique, 'Title still needs its uniqueness applied');
  });

  test('turns on version history', () => {
    const plan = planProvisioning(LIST, {});
    assert.ok(plan.actions.some((a) => a.kind === 'enableVersioning'));
  });
});

suite('planner: an up-to-date site', () => {
  const healthy = snapshot({
    fields: {
      Title: field({ internal: 'Title', typeAsString: 'Text', indexed: true, unique: true }),
      DemoText: field({ internal: 'DemoText', typeAsString: 'Text' }),
      DemoStatus: field({ internal: 'DemoStatus', typeAsString: 'Choice', choices: ['Active', 'Archived'] }),
      DemoLink: field({ internal: 'DemoLink', typeAsString: 'Lookup', indexed: true, relationshipDeleteBehavior: 1 })
    }
  });

  test('produces an empty plan — the health check and the fix are one function', () => {
    const plan = planProvisioning(LIST, { Demo: healthy });
    assert.equal(plan.actions.length, 0);
    assert.ok(isHealthy(plan));
  });

  test('re-running after provisioning is a no-op, so the button is safe to press twice', () => {
    const plan = planProvisioning(LIST, { Demo: healthy });
    const second = planProvisioning(LIST, { Demo: healthy });
    assert.equal(plan.actions.length, second.actions.length);
  });
});

suite('planner: partial and awkward sites', () => {
  test('only the missing column is planned', () => {
    const plan = planProvisioning(LIST, {
      Demo: snapshot({
        fields: {
          Title: field({ internal: 'Title', typeAsString: 'Text', indexed: true, unique: true }),
          DemoStatus: field({ internal: 'DemoStatus', typeAsString: 'Choice', choices: ['Active', 'Archived'] }),
          DemoLink: field({ internal: 'DemoLink', typeAsString: 'Lookup', indexed: true, relationshipDeleteBehavior: 1 })
        }
      })
    });
    const creates = plan.actions.filter((a) => a.kind === 'createField');
    assert.equal(creates.length, 1);
    assert.equal((creates[0] as { field: { internal: string } }).field.internal, 'DemoText');
  });

  test('a wrong column type is a conflict, never an action', () => {
    const plan = planProvisioning(LIST, {
      Demo: snapshot({
        fields: {
          Title: field({ internal: 'Title', typeAsString: 'Text', indexed: true, unique: true }),
          DemoText: field({ internal: 'DemoText', typeAsString: 'Note' }),
          DemoStatus: field({ internal: 'DemoStatus', typeAsString: 'Choice', choices: ['Active', 'Archived'] }),
          DemoLink: field({ internal: 'DemoLink', typeAsString: 'Lookup', indexed: true, relationshipDeleteBehavior: 1 })
        }
      })
    });
    assert.equal(plan.conflicts.length, 1);
    assert.includes(plan.conflicts[0].reason, 'cannot change a column');
    assert.ok(!plan.actions.some((a) => a.kind === 'createField' && a.field.internal === 'DemoText'));
  });

  test('an unreadable list is refused wholesale rather than provisioned blind', () => {
    const plan = planProvisioning(LIST, { Demo: snapshot({ verified: false }) });
    assert.equal(plan.actions.length, 0, 'nothing may be created against a list we cannot see');
    assert.equal(plan.conflicts.length, 1);
    assert.includes(plan.conflicts[0].reason, 'could not be read');
  });

  test('a site that added its own choice values keeps them, in place', () => {
    const plan = planProvisioning(LIST, {
      Demo: snapshot({
        fields: {
          Title: field({ internal: 'Title', typeAsString: 'Text', indexed: true, unique: true }),
          DemoText: field({ internal: 'DemoText', typeAsString: 'Text' }),
          DemoStatus: field({ internal: 'DemoStatus', typeAsString: 'Choice', choices: ['Active', 'Local value'] }),
          DemoLink: field({ internal: 'DemoLink', typeAsString: 'Lookup', indexed: true, relationshipDeleteBehavior: 1 })
        }
      })
    });
    const add = plan.actions.filter((a) => a.kind === 'addChoices')[0] as
      { merged: string[]; add: string[] } | undefined;
    assert.ok(add, 'the missing choice must be planned');
    assert.deepEqual(add!.add, ['Archived']);
    assert.deepEqual(add!.merged, ['Active', 'Local value', 'Archived']);
  });

  test('a lookup missing its cascade behaviour is corrected', () => {
    const plan = planProvisioning(LIST, {
      Demo: snapshot({
        fields: {
          Title: field({ internal: 'Title', typeAsString: 'Text', indexed: true, unique: true }),
          DemoText: field({ internal: 'DemoText', typeAsString: 'Text' }),
          DemoStatus: field({ internal: 'DemoStatus', typeAsString: 'Choice', choices: ['Active', 'Archived'] }),
          DemoLink: field({ internal: 'DemoLink', typeAsString: 'Lookup', indexed: true, relationshipDeleteBehavior: 0 })
        }
      })
    });
    assert.ok(plan.actions.some((a) => a.kind === 'setLookupBehavior'));
  });

  test('the plan is additive only — no action can delete or retype anything', () => {
    const plan = planProvisioning(LIST, { Demo: snapshot({ fields: {} }) });
    const destructive = plan.actions.filter(
      (a) => /delete|remove|drop|retype|rename/i.test(a.kind)
    );
    assert.equal(destructive.length, 0);
  });
});

suite('schema: the real declaration', () => {
  test('every declared field produces valid CAML', () => {
    for (const list of EXPECTED_SCHEMA) {
      for (const f of list.fields) {
        if (f.builtIn) { continue; }
        const xml = fieldSchemaXml(f);
        assert.includes(xml, `Name="${f.internal}"`);
        assert.includes(xml, `StaticName="${f.internal}"`);
      }
    }
  });

  test('payload columns are plain text — rich text would corrupt the graph', () => {
    const projects = EXPECTED_SCHEMA.filter((l) => l.title === PROJECTS_LIST)[0];
    const payload = projects.fields.filter((f) => f.internal === 'ErgPayload1')[0];
    assert.includes(fieldSchemaXml(payload), 'RichText="FALSE"');
  });

  test('project lookups cascade so deleting a project cannot orphan its rows', () => {
    for (const list of EXPECTED_SCHEMA) {
      const lookup = list.fields.filter((f) => f.internal === 'ErgProject')[0];
      if (!lookup) { continue; }
      assert.equal(lookup.lookup!.behavior, 'Cascade', `${list.title}.ErgProject must cascade`);
      assert.ok(lookup.indexed, 'a cascading lookup must be indexed');
      assert.includes(fieldSchemaXml(lookup), 'RelationshipDeleteBehavior="Cascade"');
    }
  });

  test('the whole declared schema plans cleanly against an empty site', () => {
    const plan = planProvisioning(EXPECTED_SCHEMA, {});
    assert.equal(plan.conflicts.length, 0);
    assert.ok(plan.actions.length > 0);
    // Lookup targets must be created before any lookup column referencing them.
    const lastCreateList = plan.actions.map((a) => a.kind).lastIndexOf('createList');
    const firstLookup = plan.actions.findIndex(
      (a) => a.kind === 'createField' && !!a.field.lookup
    );
    assert.ok(lastCreateList < firstLookup, 'all lists must be created before any lookup column');
  });
});
