#!/usr/bin/env node
/**
 * Enterprise Relationship Graph — smoke test
 * --------------------------------------------------------------------------
 * Loads the HTML file, extracts the embedded script, runs key features in a
 * Node VM sandbox with minimal browser stubs, and reports pass/fail.
 *
 * Purpose: catch regressions before pushing. Run as part of every major edit.
 *
 * Usage:    node scripts/smoke-test.js
 * Exit:     0 if all tests pass, 1 otherwise.
 *
 * Adding a new test:
 *   test('name of thing', () => {
 *     const sandbox = buildSandbox(['SAMPLE_DATA', 'computeLeadershipMetrics']);
 *     // ... assertions
 *   });
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const HTML_PATH = path.resolve(__dirname, '..', 'enterprise-relationship-graph.html');
const html = fs.readFileSync(HTML_PATH, 'utf8');
const scripts = [...html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/g)];
const code = scripts[scripts.length - 1][1];

// --- assertion helpers ---
function expect(actual) {
  return {
    toBe(expected, msg) { if (actual !== expected) throw new Error(`${msg || ''} expected ${actual} === ${expected}`); },
    toBeGt(n, msg)     { if (!(actual > n))    throw new Error(`${msg || ''} expected ${actual} > ${n}`); },
    toBeGte(n, msg)    { if (!(actual >= n))   throw new Error(`${msg || ''} expected ${actual} >= ${n}`); },
    toBeLt(n, msg)     { if (!(actual < n))    throw new Error(`${msg || ''} expected ${actual} < ${n}`); },
    toBeTruthy(msg)    { if (!actual)          throw new Error(`${msg || ''} expected truthy, got ${actual}`); },
    toContain(s, msg)  { if (!String(actual).includes(s)) throw new Error(`${msg || ''} expected to contain "${s}"`); },
  };
}

// --- helpers for extracting and running blocks of the embedded script ---
function findBlock(startNeedle, endNeedle) {
  const s = code.indexOf(startNeedle);
  if (s < 0) throw new Error(`Start marker not found: ${startNeedle}`);
  const e = code.indexOf(endNeedle, s);
  if (e < 0) throw new Error(`End marker not found (after start): ${endNeedle}`);
  return code.substring(s, e);
}

/**
 * Build a minimal sandbox with the browser stubs needed for the in-app code
 * to execute without throwing. We selectively expose top-level consts/functions
 * by replacing their declarations with `globalThis.X = ...` so the test can
 * call them. Any extra browser API the code touches gets a no-op stub.
 */
function buildSandbox() {
  const sandbox = {
    console, Map, Set, JSON, Math, Date, Array, Object, String, Number, Boolean, RegExp,
    setTimeout, clearTimeout, setInterval, clearInterval,
    escapeHtml: s => String(s || '').replace(/[<>&"]/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;'}[c])),
    state: {
      graph: null,
      filtersNode: new Set(),
      filtersEdge: new Set(),
      hiddenByCollapse: new Set(),
      search: '',
      showNeighborhoodOnly: false,
      selectedNode: null,
      recentNodes: [],
      fileName: 'test.json',
      lastModifiedAt: new Date().toISOString(),
      lastModifiedBy: 'smoke-test',
    },
    typeStyle: t => ({ color: '#888', icon: '', size: 14, shape: 'round' }),
    isVisibleNode: () => true,
    showToast: () => {},
    saveUIState: () => {},
    loadUIState: () => ({}),
    selectNode: () => {},
    centerOnNode: () => {},
    mEntityLink: (n, opts) => `<a data-node-id="${n.id}">${n.label}</a>`,
    mEntityLinkByName: (n, id) => `<a data-node-id="${id}">${n}</a>`,
    mKpi: (lbl, val, sub) => `[${lbl}:${val}]`,
    mHBars: data => `[bars:${data.length}]`,
    mSegmentedDonut: () => '[donut]',
    mSegLegend: () => '[legend]',
    mDonut: () => '[singleDonut]',
    mBarRow: () => '[barRow]',
    mStackedHBar: () => '[stacked]',
    statusBadge: s => s,
    navigator: { clipboard: null },
    document: {
      getElementById: () => null,
      createElement: () => ({ style: {}, classList: { add: () => {}, remove: () => {} }, addEventListener: () => {} }),
      addEventListener: () => {},
    },
  };
  sandbox.globalThis = sandbox;
  vm.createContext(sandbox);
  return sandbox;
}

/** Run SAMPLE_DATA + the collapse helpers + everything between two markers. */
function setupGraphSandbox() {
  const sandbox = buildSandbox();

  // SAMPLE_DATA
  const sampleStart = code.indexOf('const SAMPLE_DATA = (function(){');
  const sampleEnd = code.indexOf('})();', sampleStart) + 5;
  const sampleBlock = code.substring(sampleStart, sampleEnd).replace(/^const SAMPLE_DATA/m, 'globalThis.SAMPLE_DATA');
  vm.runInContext(sampleBlock, sandbox);

  // Collapse helpers (used by EX-scope computation)
  const colBlock = findBlock('const COLLAPSE_EDGE_TYPES', 'function recomputeHiddenByCollapse');
  vm.runInContext(
    colBlock
      .replace(/^const COLLAPSE_EDGE_TYPES/m, 'globalThis.COLLAPSE_EDGE_TYPES')
      .replace(/^function getCollapseChildren/m, 'globalThis.getCollapseChildren = function')
      .replace(/^function getCollapseDescendants/m, 'globalThis.getCollapseDescendants = function'),
    sandbox
  );

  sandbox.state.graph = sandbox.SAMPLE_DATA;
  sandbox.state.filtersNode = new Set(sandbox.SAMPLE_DATA.nodes.map(n => n.type));
  sandbox.state.filtersEdge = new Set(sandbox.SAMPLE_DATA.edges.map(e => e.type));
  return sandbox;
}

/** Expose the metrics + analysis functions into the existing sandbox. */
function loadMetricsAndAnalysis(sandbox) {
  const metricsBlock = findBlock('const METRICS_DEPENDENCY_EDGES', '/* ============================================================================\n   WIRE UP');
  vm.runInContext(
    metricsBlock
      .replace(/^const METRICS_DEPENDENCY_EDGES/m, 'globalThis.METRICS_DEPENDENCY_EDGES')
      .replace(/^const METRICS_ORG_TYPES/m,        'globalThis.METRICS_ORG_TYPES')
      .replace(/^const METRICS_PEOPLE_TYPES/m,     'globalThis.METRICS_PEOPLE_TYPES')
      .replace(/^const METRICS_TECH_TYPES/m,       'globalThis.METRICS_TECH_TYPES')
      .replace(/^const METRICS_COLLAB_TYPES/m,     'globalThis.METRICS_COLLAB_TYPES')
      .replace(/^const METRICS_PROCESS_TYPES/m,    'globalThis.METRICS_PROCESS_TYPES')
      .replace(/^const M_PALETTE/m,                'globalThis.M_PALETTE')
      .replace(/^const RECENT_MAX/m,               'globalThis.RECENT_MAX')
      .replace(/^const IMPACT_DEPENDENCY_EDGES/m,  'globalThis.IMPACT_DEPENDENCY_EDGES')
      .replace(/^const IMPACT_OWNERSHIP_EDGES/m,   'globalThis.IMPACT_OWNERSHIP_EDGES')
      .replace(/^function computeLeadershipMetrics/m, 'globalThis.computeLeadershipMetrics = function')
      .replace(/^function generateMermaid/m,       'globalThis.generateMermaid = function')
      .replace(/^function computeValidationViolations/m, 'globalThis.computeValidationViolations = function')
      .replace(/^function computeImpactTree/m,     'globalThis.computeImpactTree = function')
      .replace(/^function computeOrgIntelligence/m, 'globalThis.computeOrgIntelligence = function')
      .replace(/^function computeMetricsFor/m,     'globalThis.computeMetricsFor = function')
      .replace(/^function computeTopRisks/m,       'globalThis.computeTopRisks = function')
      .replace(/^function deltaChip/m,             'globalThis.deltaChip = function')
      // The rest of the functions don't need to be exposed for testing
      .replace(/^function /gm, 'function _'),
    sandbox
  );
}

// --- test registry ---
const tests = [];
function test(name, fn) { tests.push({ name, fn }); }

// --- TESTS ---

test('Script extracted and reasonable size', () => {
  expect(code.length).toBeGt(200000, 'script size');
});

test('Sample data loads with 80 nodes and 100+ edges', () => {
  const sb = setupGraphSandbox();
  expect(sb.SAMPLE_DATA.nodes.length).toBeGte(80, 'node count');
  expect(sb.SAMPLE_DATA.edges.length).toBeGte(100, 'edge count');
});

test('Leadership metrics compute without error', () => {
  const sb = setupGraphSandbox();
  loadMetricsAndAnalysis(sb);
  const m = sb.computeLeadershipMetrics();
  expect(m.N).toBeGte(80);
  expect(m.E).toBeGte(100);
  expect(m.risk.criticalHubs.length).toBeGt(0, 'should find at least one critical hub');
  expect(m.org.dirStats.length).toBeGt(0, 'should have directorate stats');
  expect(m.pipeline.activeWorkflows + m.pipeline.activeAutomations).toBeGt(0, 'should have active workflows');
});

test('Mermaid export produces valid syntax with classDef rules', () => {
  const sb = setupGraphSandbox();
  loadMetricsAndAnalysis(sb);
  const md = sb.generateMermaid();
  expect(md).toContain('flowchart TD');
  expect(md).toContain('classDef execOffice');
  expect(md).toContain('-->|');  // edge syntax
  expect(md.length).toBeGt(2000);
});

test('Impact analysis is pure-dependency (no ownership leakage)', () => {
  const sb = setupGraphSandbox();
  loadMetricsAndAnalysis(sb);
  // Power Automate platform should affect only the workflows that depend on it,
  // not "everything in the org"
  const impactPA = sb.computeImpactTree('plat-pa');
  expect(impactPA.totalAffected).toBeGt(0);
  expect(impactPA.totalAffected).toBeLt(15, 'PA impact should be small, not the whole graph');
});

test('Validation rules find expected gaps in sample data', () => {
  const sb = setupGraphSandbox();
  loadMetricsAndAnalysis(sb);
  const violations = sb.computeValidationViolations();
  // Sample data has Persons without email and without office field
  const personNoEmail = violations.find(v => v.rule.id === 'info-person-no-email');
  expect(personNoEmail).toBeTruthy('should detect Persons missing email');
  expect(personNoEmail.nodes.length).toBeGte(5);
});

test('Org Intelligence computes structural patterns', () => {
  const sb = setupGraphSandbox();
  loadMetricsAndAnalysis(sb);
  const oi = sb.computeOrgIntelligence();
  expect(oi).toBeTruthy();
  // The bridge analysis should at minimum find the M365 stack platforms
  expect(oi.bridges.length).toBeGte(2, 'should find bridge entities');
  expect(oi.topOwners.length).toBeGt(0, 'should find named owners');
});

test('Critical hubs include the M365 platforms', () => {
  const sb = setupGraphSandbox();
  loadMetricsAndAnalysis(sb);
  const m = sb.computeLeadershipMetrics();
  const hubLabels = m.risk.criticalHubs.map(h => h.n.label).join(',');
  expect(hubLabels).toContain('Power Automate');
});

test('Per-directorate footprint is non-empty (regression test)', () => {
  const sb = setupGraphSandbox();
  loadMetricsAndAnalysis(sb);
  const m = sb.computeLeadershipMetrics();
  // The descendant-based rollup should find people and tech in each directorate
  const hasContent = m.org.dirStats.some(d => d.people > 0 || d.systemsOwned > 0);
  expect(hasContent).toBeTruthy('directorates should have descendants counted');
});

test('typeStyle sanitizes malicious custom-type icon and color (XSS guard)', () => {
  const sb = buildSandbox();
  // TYPE_STYLE table + typeStyle/isBuiltInType
  const styleBlock = findBlock('const TYPE_STYLE = {', '// Type-style lookup');
  vm.runInContext(styleBlock.replace(/^const TYPE_STYLE/m, 'globalThis.TYPE_STYLE'), sb);
  const fnBlock = findBlock('function typeStyle', 'function isBuiltInType');
  vm.runInContext(fnBlock.replace(/^function typeStyle/m, 'globalThis.typeStyle = function'), sb);
  sb.state.graph = { nodes: [], edges: [], customNodeTypes: [
    { name: 'Evil', color: 'red;} body{display:none', icon: '<img src=x onerror=alert(1)>' },
    { name: 'Good', color: '#AB34CD', icon: '' },
  ]};
  const evil = sb.typeStyle('Evil');
  expect(evil.icon).toBe('', 'markup icon must be rejected');
  expect(/^#[0-9a-fA-F]{3,8}$/.test(evil.color)).toBe(true, 'non-hex color must fall back to a safe hex');
  const good = sb.typeStyle('Good');
  expect(good.color).toBe('#AB34CD', 'valid hex color preserved');
  expect(good.icon).toBe('', 'glyph icon preserved');
});

test('coerceLoadedGraph repairs/drops malformed nodes and dangling edges', () => {
  const sb = buildSandbox();
  let n = 0; sb.uid = p => p + '-test-' + (++n);
  const block = findBlock('function coerceLoadedGraph', 'function applyLoadedBundle');
  vm.runInContext(block.replace(/^function coerceLoadedGraph/m, 'globalThis.coerceLoadedGraph = function'), sb);
  const graph = {
    nodes: [
      { id: 'a', label: 'Alpha', type: 'Office' },
      { id: 'b', type: 'Person', tags: 'one, two' },     // missing label, string tags
      { id: 'a', label: 'Dupe', type: 'Office' },        // duplicate id
      null,                                               // garbage
      { label: 'No identity at all' },                    // no id
    ],
    edges: [
      { id: 'e1', source: 'a', target: 'b', type: 'CONTAINS' },
      { source: 'a', target: 'b', type: 'USES' },         // missing id -> generated
      { id: 'e3', source: 'a', target: 'ghost', type: 'USES' }, // dangling
      { id: 'e4', source: 'a', target: 'b' },              // missing type
    ],
  };
  const report = sb.coerceLoadedGraph(graph);
  expect(graph.nodes.length).toBe(2, 'two valid nodes survive');
  expect(graph.nodes[1].label).toBe('b', 'label defaults to id');
  expect(Array.isArray(graph.nodes[1].tags)).toBe(true, 'string tags coerced to array');
  expect(graph.nodes[1].tags.length).toBe(2);
  expect(graph.edges.length).toBe(2, 'dangling/typeless edges dropped');
  expect(graph.edges[1].id).toBeTruthy('missing edge id generated');
  expect(report.droppedNodes).toBe(3);
  expect(report.droppedEdges).toBe(2);
});

test('allEdgeTypes includes unrecognized imported edge types', () => {
  const sb = buildSandbox();
  const typesBlock = findBlock('const EDGE_TYPES', 'const EDGE_CATEGORY');
  vm.runInContext(typesBlock.replace(/^const EDGE_TYPES/m, 'globalThis.EDGE_TYPES'), sb);
  const fnBlock = findBlock('function allEdgeTypes', '/* ---------- Group/branch collapse');
  vm.runInContext(fnBlock.replace(/^function allEdgeTypes/m, 'globalThis.allEdgeTypes = function'), sb);
  sb.state.graph = { nodes: [], edges: [
    { id: 'e1', source: 'a', target: 'b', type: 'FUNDED_BY' },
    { id: 'e2', source: 'a', target: 'b', type: 'USES' },
  ]};
  const all = sb.allEdgeTypes();
  expect(all.includes('FUNDED_BY')).toBe(true, 'custom type seeded into filter list');
  expect(all.includes('USES')).toBe(true, 'built-ins retained');
  expect(all.filter(t => t === 'USES').length).toBe(1, 'no duplicates');
});

test('Edge hit-testing measures distance to the drawn curve, not the chord', () => {
  const sb = buildSandbox();
  const curveBlock = findBlock('function pointToEdgeCurveDist', '/* ----- Mouse interaction');
  vm.runInContext(
    curveBlock
      .replace(/^function pointToEdgeCurveDist/m, 'globalThis.pointToEdgeCurveDist = function')
      .replace(/^function pointToSegmentDist/m, 'globalThis.pointToSegmentDist = function'),
    sb
  );
  const a = { x: 0, y: 0 }, b = { x: 400, y: 0 }; // long edge -> max bow of 20
  // Curve midpoint bows 10 world px off the chord (quadratic midpoint = offset/2... actually
  // B(0.5) = 0.25*a + 0.5*cp + 0.25*b, cp offset 20 -> bow 10). Click ON the bow:
  const distAtBow = sb.pointToEdgeCurveDist(200, 10, a, b);
  expect(distAtBow).toBeLt(1.5, 'click on the visible curve should register');
  // A click on the chord midpoint is ~10px from the curve, not 0:
  const distAtChord = sb.pointToEdgeCurveDist(200, 0, a, b);
  expect(distAtChord).toBeGt(8, 'chord midpoint is measurably off the drawn curve');
});

test('Structured search parses operators, quotes, and free text', () => {
  const sb = buildSandbox();
  const block = findBlock('const SEARCH_FIELDS', 'function isVisibleNode');
  vm.runInContext(
    block
      .replace(/^const SEARCH_FIELDS/m, 'globalThis.SEARCH_FIELDS')
      .replace(/^function parseSearchQuery/m, 'globalThis.parseSearchQuery = function')
      .replace(/^function searchClauses/m, 'globalThis.searchClauses = function')
      .replace(/^function nodeMatchesSearch/m, 'globalThis.nodeMatchesSearch = function'),
    sb
  );
  const clauses = sb.parseSearchQuery('type:Person owner:"jane smith" misc:x sharepoint');
  expect(clauses.length).toBe(4);
  expect(clauses[0].field).toBe('type');
  expect(clauses[1].value).toBe('jane smith', 'quoted value keeps spaces');
  expect(clauses[2].field).toBe(null, 'unknown field treated as free text');
  expect(clauses[3].field).toBe(null);

  const person = { id:'p1', label:'Jane Smith', type:'Person', owner:'', tags:['hr'], status:'active' };
  const site   = { id:'s1', label:'HR SharePoint', type:'SharePoint Site', owner:'Jane Smith', tags:[], status:'retired' };
  sb.state.search = 'type:Person';
  expect(sb.nodeMatchesSearch(person)).toBe(true);
  expect(sb.nodeMatchesSearch(site)).toBe(false);
  sb.state.search = 'owner:"jane smith" status:retired';
  expect(sb.nodeMatchesSearch(site)).toBe(true, 'AND of two field clauses');
  expect(sb.nodeMatchesSearch(person)).toBe(false);
  sb.state.search = 'sharepoint';
  expect(sb.nodeMatchesSearch(site)).toBe(true, 'free text matches label');
  sb.state.search = 'tag:hr';
  expect(sb.nodeMatchesSearch(person)).toBe(true, 'tag: operator');
});

test('Node positions round-trip through graph.positions', () => {
  const sb = buildSandbox();
  const block = findBlock('function syncPositionsToGraph', '// Debounced localStorage persist');
  vm.runInContext(
    block
      .replace(/^function syncPositionsToGraph/m, 'globalThis.syncPositionsToGraph = function')
      .replace(/^function restorePositionsFromGraph/m, 'globalThis.restorePositionsFromGraph = function'),
    sb
  );
  sb.state.graph = { nodes: [{ id:'a' }, { id:'b' }, { id:'c' }], edges: [] };
  sb.state.positions = new Map([
    ['a', { x: 100.26, y: -50, vx: 3, vy: 1, fixed: true, pinned: true }],
    ['b', { x: 0, y: 0, vx: 0, vy: 0, fixed: false }],
    ['ghost', { x: 1, y: 2 }], // deleted node — must not serialize
  ]);
  sb.syncPositionsToGraph();
  expect(Object.keys(sb.state.graph.positions).length).toBe(2, 'only live nodes serialized');
  expect(sb.state.graph.positions.a.x).toBe(100.3, 'coords rounded to 0.1');
  expect(sb.state.graph.positions.a.pinned).toBe(1, 'pin flag survives');
  expect(sb.state.graph.positions.b.pinned).toBe(undefined);

  sb.state.positions = new Map();
  const restored = sb.restorePositionsFromGraph();
  expect(restored).toBe(2);
  const a = sb.state.positions.get('a');
  expect(a.x).toBe(100.3);
  expect(a.pinned).toBe(true);
  expect(a.fixed).toBe(true, 'pinned implies fixed (sim only checks fixed)');
  expect(sb.state.positions.get('b').fixed).toBe(false, 'unpinned restores unfixed');
  expect(a.vx).toBe(0, 'velocities reset on restore');
  expect(sb.state.positions.has('c')).toBe(false, 'node without saved position left for ensurePositions');
  // Garbage positions object must not throw
  sb.state.graph.positions = "corrupt";
  expect(sb.restorePositionsFromGraph()).toBe(0);
});

test('Top Risks watchlist + baseline metrics helpers', () => {
  const sb = setupGraphSandbox();
  loadMetricsAndAnalysis(sb);
  // computeTopRisks calls these from outside the metrics block — stub them
  sb.computeArticulationPoints = () => [];
  sb.findNode = id => sb.state.graph.nodes.find(n => n.id === id);

  const risks = sb.computeTopRisks();
  expect(risks.length).toBeGt(0, 'sample data should produce a watchlist');
  expect(risks[0].score >= risks[risks.length - 1].score).toBe(true, 'sorted by score desc');
  expect(risks[0].reasons.length).toBeGt(0, 'top risk carries human-readable reasons');

  // computeMetricsFor: runs against a different graph, restores state after
  const base = JSON.parse(JSON.stringify(sb.SAMPLE_DATA));
  base.nodes = base.nodes.slice(0, base.nodes.length - 5);
  const keep = new Set(base.nodes.map(n => n.id));
  base.edges = base.edges.filter(e => keep.has(e.source) && keep.has(e.target));
  const pm = sb.computeMetricsFor(base);
  expect(pm.N).toBe(sb.SAMPLE_DATA.nodes.length - 5, 'metrics reflect the swapped-in graph');
  expect(sb.state.graph).toBe(sb.SAMPLE_DATA, 'state.graph restored after the swap');

  // deltaChip: direction, polarity, no-baseline, flat
  expect(sb.deltaChip(10, 7, 'up')).toContain('+3');
  expect(sb.deltaChip(10, 7, 'up')).toContain('good');
  expect(sb.deltaChip(10, 7, 'down')).toContain('bad');
  expect(sb.deltaChip(5, 5, 'up')).toContain('flat');
  expect(sb.deltaChip(5, null, 'up')).toBe('');
});

test('XSS regression guards present in source', () => {
  // Cheap static assertions that the known injection points stay escaped.
  expect(code).toContain('escapeHtml(cx.byDeg[0].n.label)');           // metrics top-hub KPI
  expect(code).toContain('const safeIcon');                            // typeStyle sanitization
  expect(html).toContain('id="toast" role="status" aria-live="polite"'); // announced toasts
  expect(code).toContain('/^https?:\\/\\//i.test(node.url');           // javascript: URL guard
  expect(html).toContain('http-equiv="Content-Security-Policy"');      // CSP meta present
  expect(code).toContain('BUNDLE_SCHEMA_VERSION = 4');                 // schema version pinned
});

// --- runner ---
let passed = 0, failed = 0;
const startTime = Date.now();

console.log(`\nRunning ${tests.length} smoke tests…\n`);
for (const t of tests) {
  try {
    t.fn();
    console.log(`  PASS  ${t.name}`);
    passed++;
  } catch (e) {
    console.log(`  FAIL  ${t.name}`);
    console.log(`        ${e.message}`);
    failed++;
  }
}
const elapsed = Date.now() - startTime;
console.log(`\n${passed} passed, ${failed} failed (${elapsed}ms)`);
process.exit(failed === 0 ? 0 : 1);
