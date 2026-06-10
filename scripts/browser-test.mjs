#!/usr/bin/env node
/**
 * Enterprise Relationship Graph — browser test (Playwright)
 * ---------------------------------------------------------------------------
 * Covers the DOM-level behavior the Node smoke test can't reach: real boot,
 * canvas interaction, localStorage persistence, dialogs, and modal keyboard
 * handling. These are exactly the areas where past regressions lived
 * (e.g. "loading a snapshot destroys saved views" was invisible to the VM
 * harness).
 *
 * Setup:    npm install            (installs playwright)
 *           npx playwright install chromium
 * Usage:    npm run test:browser   (or: node scripts/browser-test.mjs)
 * Exit:     0 if all tests pass, 1 otherwise.
 */
import { chromium } from 'playwright';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const APP_URL = pathToFileURL(path.resolve(__dirname, '..', 'enterprise-relationship-graph.html')).href;

let passed = 0, failed = 0;
function check(name, cond, detail) {
  if (cond) { console.log(`  PASS  ${name}`); passed++; }
  else      { console.log(`  FAIL  ${name}${detail ? `\n        ${detail}` : ''}`); failed++; }
}

const browser = await chromium.launch();
const context = await browser.newContext({ viewport: { width: 1440, height: 900 } });
const page = await context.newPage();

// prompt() -> supply a name; confirm()/beforeunload -> accept
page.on('dialog', d => d.accept(d.type() === 'prompt' ? 'PW Baseline' : undefined).catch(() => {}));
const pageErrors = [];
page.on('pageerror', e => pageErrors.push(String(e)));

async function waitForBoot() {
  await page.waitForFunction(
    () => typeof state !== 'undefined' && state.graph && state.graph.nodes.length > 0,
    null, { timeout: 15000 }
  );
}

console.log(`\nRunning browser tests against ${APP_URL}\n`);
await page.goto(APP_URL);
await waitForBoot();

// ---- 1. Boot: sample data, stats wired, no uncaught errors ----
const boot = await page.evaluate(() => ({
  nodes: state.graph.nodes.length,
  statText: document.getElementById('stat-nodes').textContent,
}));
check('boots with sample data', boot.nodes >= 80, `${boot.nodes} nodes`);
check('header stats render', boot.statText === String(boot.nodes), `stat shows "${boot.statText}"`);

// ---- 2. Structured search filters the canvas ----
await page.fill('#search-input', 'type:Person');
const search = await page.evaluate(() => ({
  visible: visibleNodes().length,
  people: state.graph.nodes.filter(n => n.type === 'Person').length,
}));
check('structured search narrows visible nodes to matches', search.visible === search.people,
  `visible=${search.visible}, Person nodes=${search.people}`);
await page.click('#btn-clear-search');

// ---- 3. Edge hit-testing: click the drawn curve of the longest edge ----
await page.evaluate(() => fitGraph());
await page.waitForTimeout(400); // allow fit + render to settle
const edgeTarget = await page.evaluate(() => {
  // pick the longest visible edge so its midpoint has clearance from nodes
  let best = null, bestLen = -1;
  for (const e of state.graph.edges) {
    if (!isVisibleEdge(e)) continue;
    const a = state.positions.get(e.source), b = state.positions.get(e.target);
    if (!a || !b) continue;
    const len = Math.hypot(b.x - a.x, b.y - a.y);
    if (len > bestLen) { bestLen = len; best = { e, a, b }; }
  }
  if (!best) return null;
  const { e, a, b } = best;
  const dx = b.x - a.x, dy = b.y - a.y, len = Math.hypot(dx, dy) || 1;
  const off = Math.min(20, len * 0.08), nx = -dy / len, ny = dx / len;
  const cpx = (a.x + b.x) / 2 + nx * off, cpy = (a.y + b.y) / 2 + ny * off;
  const wx = 0.25 * a.x + 0.5 * cpx + 0.25 * b.x;   // quadratic at t=0.5
  const wy = 0.25 * a.y + 0.5 * cpy + 0.25 * b.y;
  const t = state.transform;
  const rect = document.getElementById('graph-canvas').getBoundingClientRect();
  return { id: e.id, x: rect.left + wx * t.scale + t.x, y: rect.top + wy * t.scale + t.y };
});
if (edgeTarget) {
  await page.mouse.click(edgeTarget.x, edgeTarget.y);
  const selEdge = await page.evaluate(() => state.selectedEdge);
  check('clicking the drawn edge curve selects the edge', selEdge === edgeTarget.id,
    `selectedEdge=${selEdge}, expected=${edgeTarget.id} (a node may overlap the midpoint)`);
} else {
  check('clicking the drawn edge curve selects the edge', false, 'no visible edge found');
}
await page.keyboard.press('Escape');

// ---- 4. Views survive snapshot loads (regression: data loss) ----
await page.evaluate(() => saveView());                    // prompt -> "PW Baseline"
await page.evaluate(() => saveCurrentSnapshot());          // prompt -> "PW Baseline"
const snapId = await page.evaluate(() => loadSnapshots()[0].id);
await page.evaluate(id => loadSnapshot(id), snapId);       // confirm -> accepted
const afterLoad = await page.evaluate(() => ({
  views: (state.graph.views || []).length,
  nodes: state.graph.nodes.length,
}));
check('saved views survive loading a snapshot', afterLoad.views >= 1, `views=${afterLoad.views}`);
check('snapshot load restores the graph', afterLoad.nodes >= 80, `nodes=${afterLoad.nodes}`);

// ---- 5. Undo restores a deleted node ----
const delTarget = await page.evaluate(() => state.graph.nodes[0].id);
await page.evaluate(id => deleteNode(id), delTarget);
const goneAfterDelete = await page.evaluate(id => !findNode(id), delTarget);
await page.evaluate(() => undo());
const backAfterUndo = await page.evaluate(id => !!findNode(id), delTarget);
check('delete + undo round-trips a node', goneAfterDelete && backAfterUndo,
  `gone=${goneAfterDelete}, restored=${backAfterUndo}`);

// ---- 6. Hand-arranged positions survive a reload ----
const pinned = await page.evaluate(() => {
  const id = state.graph.nodes[0].id;
  const p = state.positions.get(id);
  p.x = 12345; p.y = 678; p.pinned = true; p.fixed = true;
  persist();
  return id;
});
await page.reload();
await waitForBoot();
const restored = await page.evaluate(id => {
  const p = state.positions.get(id);
  return p ? { x: p.x, y: p.y, pinned: !!p.pinned, fixed: !!p.fixed } : null;
}, pinned);
check('pinned position survives reload exactly',
  restored && Math.abs(restored.x - 12345) < 0.2 && Math.abs(restored.y - 678) < 0.2 && restored.pinned && restored.fixed,
  JSON.stringify(restored));

// ---- 7. Modal keyboard: open with N, close with Escape, shortcuts muted while open ----
await page.click('#graph-canvas', { position: { x: 30, y: 30 } });
await page.keyboard.press('Escape'); // clear any selection from the click
await page.keyboard.press('n');
const modalOpen = await page.evaluate(() => document.getElementById('modal-backdrop').classList.contains('open'));
const nodesBeforeKeys = await page.evaluate(() => state.graph.nodes.length);
await page.keyboard.press('e'); // must NOT replace the modal / act on the graph
await page.keyboard.press('Escape');
const modalClosed = await page.evaluate(() => !document.getElementById('modal-backdrop').classList.contains('open'));
const nodesAfterKeys = await page.evaluate(() => state.graph.nodes.length);
check('N opens the node modal', modalOpen);
check('Escape closes the modal', modalClosed);
check('global shortcuts are muted under an open modal', nodesBeforeKeys === nodesAfterKeys);

// ---- 8. Metrics modal renders with Top Risks and baseline picker ----
await page.evaluate(() => openMetricsModal());
const metrics = await page.evaluate(() => ({
  open: document.getElementById('modal-backdrop').classList.contains('open'),
  kpis: document.querySelectorAll('#modal-body .kpi').length,
  baseline: !!document.getElementById('m-baseline'),
  topRisks: document.body.textContent.includes('Top risks'),
}));
check('metrics modal opens with KPIs', metrics.open && metrics.kpis >= 8, `kpis=${metrics.kpis}`);
check('baseline picker present', metrics.baseline);
check('Top Risks section present', metrics.topRisks);
await page.keyboard.press('Escape');

// ---- 9. Canvas is not a keyboard trap ----
await page.click('#graph-canvas', { position: { x: 30, y: 30 } });
await page.evaluate(() => document.getElementById('graph-canvas').focus());
await page.keyboard.press('Tab');
const focusLeftCanvas = await page.evaluate(() => document.activeElement?.id !== 'graph-canvas');
check('Tab moves focus off the canvas (no keyboard trap)', focusLeftCanvas);

// ---- no uncaught page errors during the whole run ----
check('no uncaught page errors', pageErrors.length === 0, pageErrors.slice(0, 3).join(' | '));

await browser.close();
console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed === 0 ? 0 : 1);
