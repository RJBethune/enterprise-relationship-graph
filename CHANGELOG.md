# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

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
