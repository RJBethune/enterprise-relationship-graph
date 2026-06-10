# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [1.5.0] - 2026-06-10

Leadership metrics: trends and a composite risk watchlist, plus a browser test.

### Added
- **Trend deltas vs. a baseline snapshot** — the Executive Summary now has a
  "Compare to" picker (defaults to the most recent snapshot, choice persists).
  Headline KPIs show ▲/▼ change chips colored by whether the direction is
  good (more ownership coverage = green; more critical hubs = red), and a
  "Trend across snapshots" table tracks entities, relations, owner %, critical
  hubs, and deprecated-in-use across the last 8 snapshots. Powered by
  `computeMetricsFor(graph)`, which runs the full metrics engine against any
  graph object (e.g. a saved snapshot).
- **Top Risks watchlist** — `computeTopRisks()` combines every risk signal
  (dependency load, missing backups, articulation points / SPOFs,
  deprecated-but-connected, ownership gaps, staleness compounding) into one
  ranked, owner-attributed list. Top 3 on the Executive Summary, full top 10
  with scores and reasons on the Risk & SPOFs tab, top 5 in the Copy-summary
  markdown.
- **Playwright browser test** (`npm run test:browser`) covering what the Node
  harness can't: real boot, structured-search canvas filtering, clicking the
  drawn edge curve, views surviving snapshot loads, delete+undo, pinned
  positions surviving reload, modal keyboard behavior, metrics modal render,
  and the no-keyboard-trap guarantee. Setup: `npm install`,
  `npx playwright install chromium`.

### Fixed
- Restored pinned positions now also restore `fixed`, matching the force
  sim's contract — without it, pins drifted after reload.

## [1.4.0] - 2026-06-09

### Added
- **Persisted layouts** — node positions (and pin flags) are saved into the
  JSON bundle (`graph.positions`, schema v4) and the localStorage cache, and
  restored exactly on open instead of re-running the layout algorithm. Drag
  re-arrangements persist via a debounced cache write; picking a layout from
  the dropdown still recomputes as usual.
- **Structured search** — the search box accepts field operators with quoted
  values (`type:Person owner:"jane smith" status:retired`); supported fields:
  type, tag, owner, lead, status, platform, office, directorate, id. Plain
  terms keep matching label/type/tags; all clauses AND. Shared by the canvas
  highlight, visibility filter, and "Jump to".
- **Schema versioning** — bundles are stamped `version: 4`
  (`BUNDLE_SCHEMA_VERSION`); opening a file written by a newer build warns
  that saving may drop newer fields.
- **Content-Security-Policy meta** — only the three CDNs already in use are
  allowed (Font Awesome, Google Fonts, SheetJS); all other remote loads and
  all fetch/XHR are blocked, so any future injection cannot exfiltrate data.
- Two new smoke tests (structured search parsing/matching, position
  round-trip) — 16 total.

## [1.3.0] - 2026-06-09

Full code review pass: two stored-XSS fixes, data-loss and state bugs, and a
Section 508 / WCAG 2.1 AA remediation. See CODE-REVIEW-2026-06-09.md for the
complete findings list.

### Security
- Custom node types from imported JSON are sanitized in `typeStyle()` (hex-only
  colors, markup rejected in icons); custom type names escaped in the node form.
- Node labels escaped in the Metrics "top hub" KPI (was an innerHTML injection).
- `javascript:` URLs in node profiles render as text instead of links.

### Fixed
- Loading a snapshot no longer destroys saved Views and Walkthroughs.
- Stale path-finder / multi-select state is cleared when a file, sample, or
  snapshot loads (previously dimmed the entire new graph).
- "Add node here" places the node at the click position (was always center).
- Edge hit-testing follows the drawn quadratic curve and scales with zoom —
  edges are now clickable when zoomed out and clicks on the visible curve land.
- Pan works on empty canvas during what-if, path-finder, and quick-connect modes.
- Inline edge retype/reverse that collides with an existing identical edge now
  selects the surviving edge instead of leaving a dead editor; manual edits to
  auto-created edges clear the `auto` flag so node saves don't revert them.
- Edges with unrecognized types from imported files are seeded into the edge
  filter (new `allEdgeTypes()`), appear in the filter UI, and stay selectable
  in the edge editors.
- Malformed imported files are coerced on load (`coerceLoadedGraph`): label-less
  nodes repaired, duplicate-ID/garbage nodes and dangling edges dropped, with a
  cleanup toast.
- Org-chart lead clustering only rescues Person/Role nodes (an orphaned Office
  that MANAGES an Application no longer becomes the Application's child).
- "People named as Owner" no longer counts org units named in owner/lead fields.
- Path-finder branch expansion persists through afterMutate (survives reload,
  participates in undo).
- Deleting a node cleans it from collapsedNodes and multi-select; Manage Types
  edit toast says "updated"; redundant double "Opened/Reopened" toast removed;
  Impact modal description matches the algorithm; dead exportPng vars removed.

### Accessibility (Section 508 / WCAG 2.1 AA)
- Canvas keyboard trap removed: Tab moves focus normally; `[`/`]` and
  PageUp/PageDown cycle nodes; arrows navigate spatially.
- Toasts are announced via a live region (assertive for errors); inline
  relationship warnings use role=alert.
- Relationship rows, context menu (role=menu, arrow keys, Shift+F10 to open),
  and per-node Collapse/Expand branch are fully keyboard-operable.
- Edge strokes meet the 3:1 non-text contrast minimum in dark theme and have
  dedicated saturated light-theme colors (legend swatches follow the theme).
- "Last edited" shown in node profile and tooltip (staleness no longer
  color-only); dark-theme primary buttons meet 4.5:1; prefers-reduced-motion
  honored (static dashes, jump-cut camera); desktop-blocking overlay no longer
  triggers at 400% browser zoom.
- aria-pressed on all toggle buttons; tablist/tab semantics on Help, Metrics,
  and Legend tabs; labels/names for sidebar selects, search, walkthrough
  builder rows, icon-only buttons, and the legend toggle; landmark labels on
  both side panels; presentation overlay manages focus; recently-viewed
  popover exposes state, closes on Escape, and shows a visible focus ring.

### Tests
- Five new smoke tests: typeStyle XSS sanitization, malformed-file coercion,
  unknown edge-type filter seeding, curve-aware edge hit-testing, and static
  XSS regression guards (14 total).

## [1.2.0] - 2026-05-31

Performance, scale, and a round of analysis/UX refinements.

### Performance & scale
- Force simulation rewritten with a Barnes-Hut quadtree: O(n log n) repulsion
  instead of the previous O(n²) all-pairs loop. Validated against brute force
  (exact as the opening angle approaches 0; ~5% average force error at the
  shipped theta=0.75, imperceptible in an organic layout). ~1.8x faster at
  1,000 nodes, ~7x at 2,000, scaling from there.
- Freeze-when-settled: the simulation halts once movement stays below threshold
  for 60 consecutive frames and wakes on drag, edit, layout change, or resize.
  A settled graph of any size now costs zero CPU.
- Hit-testing now uses a uniform spatial-hash grid (cursor cell + 8 neighbors)
  instead of scanning every node. Validated identical to the prior linear scan
  across 60,000 probes, including topmost-wins ordering.
- Net effect: interactive graphs into the low thousands of nodes, up from a few
  hundred. Node sizes are computed into a parallel array, so nothing leaks into
  the saved JSON.

### Analysis
- What-if scenario mode now models organizational orphaning in addition to
  dependency breakage: disabling a node that contains/owns others (e.g. a
  directorate over its offices) cascades down CONTAINS / OWNS / MANAGES / HAS_*
  and up REPORTS_TO, marking orphaned descendants amber. Broken dependents stay
  red; the two impact kinds are shown distinctly.

### UI / UX
- Dim relationships toggle: shade edges darker so a busy graph reads calmly;
  hovering or selecting a node still lights its connections. Toolbar button,
  persisted in localStorage.
- Fixed the legend collapse control overlapping the "Legend" title when the
  legend box is collapsed.

### Content & icons
- Added Jira as a built-in sample node (System), consistent with how ServiceNow,
  Salesforce, and ArcGIS are modeled; the Jira brand icon was already available
  in the icon registry.
- Distribution List nodes now use the envelopes-bulk icon (a centered envelope
  stack) instead of the off-center paper-plane, better distinguishing them from
  single-envelope Mailbox and people-icon O365 Group nodes.

## [1.1.0] - 2026-05-30

Adds analysis, bulk curation, and briefing tools on top of the 1.0 feature set.

### Analysis
- What-if scenario mode: click nodes to non-destructively disable them and
  see the transitive dependency cascade highlighted in red. Nothing is saved;
  exit restores the normal view.
- Single-point-of-failure detection: computes articulation points (undirected
  DFS, Tarjan-style low-link) and rings them in amber, with a clickable list.
  Verified against the sample graph: 16 cut-points, each confirmed to truly
  disconnect the graph, zero false positives.
- Staleness heatmap (`H`): tints nodes by time since last edit, with a bucketed
  legend (under a week through 6 months+). Nodes with no edit date yet show
  neutral until first touched.

### Bulk curation
- Multi-select via Ctrl/Cmd+click or Shift+drag lasso box.
- Bulk edit: change type, status, owner, lead, directorate, office, platform,
  or append tags across the whole selection at once. Ownership/lead/structural
  fields re-sync their auto-edges per node.

### Views & briefing
- Saved views capture filters + layout + zoom + selection as named perspectives,
  stored in the graph bundle so they travel with the file. Distinct from
  snapshots, which capture data.
- Presentation walkthroughs: build an ordered list of nodes each with a line of
  narration, then play fullscreen and advance with the arrow keys. Also stored
  in the bundle.

### Data model
- Nodes now carry an optional `updatedAt` ISO timestamp, stamped on every edit
  and bulk edit going forward (powers the staleness heatmap).
- Graph bundle gains optional `views` and `walkthroughs` arrays.

## [1.0.0] - 2026-05-24

Initial public release. Stable feature set for enterprise architecture review,
operational awareness, leadership briefings, and onboarding.

### Graph & rendering
- Custom HTML5 Canvas renderer (no external graph library required).
- Four layout algorithms: force-directed, hierarchical, concentric, and grid.
- Force layout with size-aware repulsion and adjustable spacing.
- Pan, zoom, drag, and click interaction; smooth animations.
- Node shapes (circle / square / diamond / triangle / hexagon) and colors
  encode entity type as two independent visual channels.
- Font Awesome icons rendered inside nodes for instant type recognition,
  with brand-icon support for known products (Microsoft, Salesforce).
- Edges styled by relationship category: hierarchy, ownership, composition,
  dependency, integration, hosting, docs, usage. Each gets a distinct
  line weight and dash pattern.

### Editing
- Add, edit, delete nodes and relationships from the UI.
- Quick-connect mode: select a source, press `C`, click any target to wire
  them with a sensible default relationship type.
- Right-click context menus on nodes, edges, and empty canvas.
- Bounded undo / redo history (30 steps) via `Ctrl+Z` / `Ctrl+Y`.
- Drag-without-select: hold Shift or use Layout mode to pin nodes for
  print-ready arrangements without changing the current selection.
- User-defined custom node types with name, color, shape, icon, and size.

### Data
- File-as-source-of-truth model with the File System Access API in Chrome and
  Edge for true in-place save. Graceful download fallback in Firefox / Safari.
- JSON bundle export includes the graph, all custom types, and named
  snapshots in one shareable file.
- CSV export of nodes and edges with resolved source / target labels.
- Excel (.xlsx) export via SheetJS, lazily loaded from CDN.
- High-resolution PNG export of the visible graph for slides.
- Named snapshots persisted in localStorage as crash-recovery cache.
- Drag-and-drop a `.json` file onto the window to open it.
- Unsaved changes warning before tab close.

### Discovery
- Search by name, type, or tag with pulse-glow highlighting.
- Path-finding (BFS, undirected traversal) between any two nodes with step-by-step
  path display and toggle button with clear on/off states.
- Neighborhood-only view: hide everything except a selected node and its
  direct connections.
- Filters for node types and relationship types with All / None toggles.
- Profile panel auto-organizes a node's relationships into sections
  (People & Roles, Systems, Platforms, Documents, Dependencies, etc.).

### Auto-sync
- Setting `Parent`, `Reports to`, `Directorate`, `Office`, `Platform`, `Owner`,
  or `Lead` on a node automatically creates or updates the corresponding
  CONTAINS, REPORTS_TO, HOSTED_ON, OWNS, or RESPONSIBLE_FOR edge.
- Auto-created edges are tagged so renaming or clearing the field updates
  them without disturbing manually authored relationships.

### Accessibility
- Modal dialogs with `role="dialog"`, focus trap, and focus restoration.
- Section headers are real buttons with `aria-expanded`.
- All form labels paired with their inputs via `for`.
- `:focus-visible` styles on every interactive element.
- Status badges combine color, icon, and text (not color alone).
- Canvas is keyboard-navigable: Tab cycles through nodes, arrow keys move
  in cardinal directions.
- Touch targets hit 44×44 px (visual size + invisible hit-area padding).
- Mobile / tablet breakpoint at 900 px stacks panels vertically.

### Keyboard shortcuts
- `Ctrl+S` save, `Ctrl+O` open, `Ctrl+Z` / `Y` undo / redo
- `F` fit, `R` reset view, `?` help
- `C` quick-connect, `N` new node, `E` edit selected
- `L` toggle always-on edge labels, `M` toggle Layout mode
- `Ctrl+F` focus search, `Esc` cancels any mode or clears selection

### Distribution
- Single self-contained HTML file (~4,200 lines).
- Two CDN dependencies (Font Awesome, Inter) loaded with CORS.
- Works offline once cached; degrades gracefully without internet
  (icons hidden, system font substituted).
