# Code Review — Enterprise Relationship Graph
**Date:** 2026-06-09 · **Scope:** `enterprise-relationship-graph.html` (~10,000 lines), `index.html`, `scripts/smoke-test.js`
**Coverage:** full read of the codebase by two independent review passes — one for bugs/logic, one for Section 508 / WCAG 2.1 AA.

Verification after fixes: both embedded script blocks parse cleanly, and all 9 smoke tests pass (`node scripts/smoke-test.js` → 9/9).

---

## Part 1 — Fixes applied in this review

### Security (critical)

**1. Stored XSS via imported custom node types.** `typeStyle()` returned `icon` and `color` verbatim from `customNodeTypes` in any opened JSON file, and those values are interpolated into `innerHTML`/`style` in the legend, filters, profile panel, tooltip, and Manage Types list. A shared file with `icon: "<img src=x onerror=...>"` executed script on open. *Fix:* `typeStyle()` now whitelists `color` to `#hex` and rejects icons containing markup characters; the custom-type `<option>` names in the node form are now escaped (matching what `openBulkEditModal` already did — the two paths had drifted).

**2. Stored XSS via node label in Leadership Metrics.** The "Max connections (top hub)" KPI passed the hub node's raw label into `mKpi()`, which builds `innerHTML` without escaping. A malicious label executed when the metrics modal opened. *Fix:* label is escaped at the call site. (`mKpi` itself can't blanket-escape because other call sites intentionally pass entities like `&ge;`.)

**3. `javascript:` URLs were clickable in the node profile.** `escapeHtml` prevents attribute breakout but not the scheme. *Fix:* only `http(s)` URLs render as links; anything else displays as text.

### Data-loss and state bugs (major)

**4. Loading a snapshot silently destroyed all Saved Views and Walkthroughs.** `snapshot()` captures only nodes/edges/types/collapsed, but `loadSnapshot` replaced `state.graph` wholesale and persisted the loss. *Fix:* views/walkthroughs are preserved across snapshot loads.

**5. Stale path-finder / multi-select state survived file loads.** Opening a file, sample, or snapshot while a found path was showing left `pathFinding.pathNodes` populated — the entire new graph rendered dimmed, with a dead path panel and stale multi-select IDs. *Fix:* new `clearLoadedViewState()` called from all three loaders.

**6. "Add node here" ignored the click position.** The context-menu action stored coordinates in `pendingNodePos`, which was never read — nodes always spawned at canvas center. *Fix:* the saved position is now used (and cleared on modal cancel so it can't leak into the next add).

**7. Keyboard shortcuts fired underneath open modals.** With focus on a modal button (e.g., a color swatch), `n`/`e`/`c` replaced the modal and `Delete` deleted the selected node behind it. *Fix:* single-key shortcuts are swallowed while a modal is open; Escape now also closes the modal when focus is inside a form field.

### Smaller bug fixes

- `deleteNode` left the deleted ID in `graph.collapsedNodes` (persisted into the file forever) and in the multi-select set. Both now cleaned.
- Manage Types: the confirmation toast always said "added" even on edit (`editingName` was nulled before the ternary read it). Now reports "Type updated" correctly.
- `selectEdge` didn't clear multi-select (inconsistent with `selectNode`), leaving stale selection rings and an armed Bulk Edit button. Now consistent.

### Section 508 / WCAG 2.1 AA fixes

- **Keyboard trap removed (2.1.2 — the most serious 508 blocker).** The canvas intercepted Tab/Shift+Tab to cycle nodes, so focus could never leave the canvas. Node cycling moved to `[` / `]` and PageUp/PageDown; arrows still navigate spatially; the canvas `aria-label` and Help → Shortcuts were updated.
- **Status messages announced (4.1.3, 3.3.1).** The toast had no live region — every save confirmation *and every form validation error* was invisible to screen readers. Toast now has `role="status"`/`aria-live` (assertive `role="alert"` for errors); the relationship-guardrail warning box is `role="alert"`.
- **Relationship rows keyboard-operable (2.1.1).** The `.rel-item` rows (the primary way to traverse the graph from the profile panel) were click-only divs. Now `role="button" tabindex="0"` with Enter/Space activation and a visible focus outline — in the node profile, edge endpoints, and path results.
- **Collapse/Expand branch reachable without a mouse (2.1.1).** This action existed only in the right-click context menu. A "Collapse branch / Expand branch" button now appears in the node profile Actions for nodes with containment children.
- **400% zoom no longer blocked (1.4.10).** The "Best viewed on desktop" overlay triggered on width alone, so a desktop browser at 400% zoom (320px viewport) lost the entire app. It now requires a coarse pointer (actual touch device) as well.
- **Reduced motion honored (2.3.3, 2.2.2).** New `prefers-reduced-motion` CSS block; the selected-edge "marching ants" dashes freeze (still dashed, not animated); camera tweens become jump-cuts.
- **Toggle state exposed (4.1.2).** `setBtnActive()` and the edge-label/layout-mode sync helpers now set `aria-pressed` (heatmap, dim-edges, what-if, SPOF, etc. — previously only Find Path did this).
- **Contrast (1.4.3).** Dark-theme primary buttons were 4.07:1 (`#F5F3FF` on `#6366F1`); a dark-theme-only override (`#5156E5`, ≈5:1) now passes. Light theme already passed and is unchanged.
- **Form names (1.3.1, 3.3.2).** `for=` added to the Layout / Cluster-by / Scope selects; `aria-label` on the search input; per-step labels on the walkthrough-builder selects, narration inputs, and reorder buttons.
- **Keyboard drill-down parity (2.1.1).** The Impact modal installed only a click delegate before setting the shared `_metricsDrillDownAttached` flag — if Impact opened first in a session, Enter/Space on entity links broke in every later metrics/data-quality modal too. Keydown delegate added.
- **Recently Viewed popover (4.1.2).** Trigger now has `aria-haspopup`/`aria-expanded` (kept in sync), Escape closes it and returns focus, and focused items show a visible outline (they were focusable but the focus state was ~1.1:1, i.e., invisible).
- **Names and small items.** `aria-label` on the icon-only header/toolbar buttons (data quality, metrics, help, zoom, fit, reset, undo, redo) with `aria-hidden` icons; legend show/hide button gets a real name + `aria-expanded`; the file chip's accessible name now includes "unsaved changes" (was color/title only — 1.4.1); the screen-reader graph mirror notes "…and N more nodes" when it truncates at 500.

---

## Part 2 — Recommendations
> **STATUS UPDATE (same day):** every item below has since been implemented in a
> second pass (see CHANGELOG 1.3.0), with two scope notes: heading-structure
> work was limited to labeling the two side-panel landmarks (sr-only h2
> restructuring wasn't warranted), and the smoke-test additions cover
> sanitization, file coercion, edge-type seeding, and curve hit-testing rather
> than a full DOM round-trip (snapshot/views persistence needs a browser, not
> the Node VM harness). Verified: both script blocks parse; 14/14 smoke tests pass.

### Bugs worth fixing next

1. **Edge hit-testing doesn't match edge rendering.** Edges draw as quadratic curves (offset up to 20 world px) but clicks test against the straight chord with a fixed 6-world-unit tolerance. Zoomed out, edges are nearly unclickable; zoomed in, the slop zone is huge; and clicks on the visible curve often miss while clicks on empty space along the invisible chord hit. Fix: sample the same quadratic and scale tolerance by `1/transform.scale`.
2. **Imported edges with unrecognized types are permanently invisible.** The edge filter set and filter UI are seeded only from built-in `EDGE_TYPES`, so a file with `"FUNDED_BY"` edges counts them in stats but never draws them, with no way to enable. (Node types handle this correctly via `allNodeTypes()` — the edge path drifted.) Fix: seed/render the edge filter from built-ins ∪ types present in the graph, or normalize unknown types on load.
3. **Inline edge type change / reverse can silently delete the edited edge.** If the change makes it a duplicate, `removeDuplicateEdges()` in `afterMutate` keeps the *first* copy and drops the one being edited; the profile then edits a dead object. Fix: after `afterMutate()`, re-find the edge and reselect the survivor, or block like the creation-time guardrail does.
4. **Manually edited auto-edges keep `auto:true`.** The edge modal's save merges over the existing object, so `auto`/`autoField` survive a manual repoint — the next save of the source node lets `syncStructuralEdges` silently revert the user's edit. Fix: clear the flags on manual edit.
5. **`parseBundle` does no shape validation.** Nodes missing `label`, or `tags` as a string, load fine and crash later in `label.localeCompare` sorts. Fix: coerce to canonical shape in `applyLoadedBundle` and drop edges with unresolvable endpoints.
6. **Pan is disabled mid-mode.** In what-if, path-finder, and quick-connect modes, empty-canvas drags return early instead of falling through to pan, so you can't reach distant nodes except by wheel-zoom.
7. **`applyFoundPath` mutates `collapsedNodes` without `persist()`/history**, so the expansion is lost on reload and invisible to undo.
8. **Org-chart `leadParents` rescue lacks a type check** — an orphaned Office that `MANAGES` an Application gets parented *under the Application* in the org tree. Restrict to Person/Role sources.
9. **"Knowledge concentration (top owners)" counts org units as people.** The fallback adds any owner string that isn't a Person label, so offices dominate a table titled "People named as Owner" — the bus-factor code nearby gets this right; mirror it.
10. **Cosmetic/doc:** `tryReopenLast` shows a contradictory double toast ("Opened…" then "Reopened…"); the Impact modal description says ownership edges are walked when they deliberately aren't; `exportPng` has two dead variables (`savedTransform`, `savedCtxBackup`); `CANVAS_COLORS.edgeDefault` is defined in both palettes but never used.

### Remaining accessibility items

11. **Canvas edge contrast (1.4.11).** Default dark-theme edge strokes composite to ≈1.8:1 against the canvas (≈1.3:1 in dim-edges mode) — below the 3:1 minimum for meaningful graphics. Consider raising the default alpha or offering a high-contrast-edges toggle. Related: `EDGE_STYLE` colors aren't theme-aware, and two hardcoded dark hover colors (`.btn:hover`, `.rel-item:hover`) render nearly black-on-black-text in light mode.
12. **Staleness heatmap is color-only (1.4.1).** Add "Last edited" to the node profile and tooltip so age has a non-color representation.
13. **Context menu semantics.** The menu is still mouse-only as a widget (`role="menu"`/`menuitem`, arrow keys, Shift+F10 to open). All its actions now have keyboard equivalents elsewhere, so this is polish rather than a blocker.
14. **Tab widgets** (Help, Metrics, Legend tabs) are plain buttons — add `role="tablist"/"tab"/"tabpanel"` with `aria-selected`, or minimally `aria-current`.
15. **Walkthrough presentation overlay** should move focus into itself on start (arrow keys already work globally, so impact is low). The data-quality relationship links use `<a>` without `href`/`role` — give them `role="link"` like `mEntityLink()` does.
16. **Heading structure** skips levels (h1 → h3/h4); consider sr-only h2s for the three main regions and `aria-label`s on the two unlabeled `<aside>`s.

### Test coverage

The smoke test covers pure computation only (metrics, mermaid, impact, validation, org-intel). Cheap high-value additions: a save→load round-trip asserting views/walkthroughs survive, and an escaping probe (a node labeled `<img onerror=…>` should never reach `innerHTML` unescaped). Items 4–5 in Part 1 and both XSS fixes would have been caught by those.

---

*Note: `index.html` (redirect page) audited and found compliant — `lang`, title, instant redirect with manual fallback link.*
