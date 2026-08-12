# CLAUDE.md — enterprise-relationship-graph

Project rules. Supplements `C:\Repos\CLAUDE.md` and `C:\Repos\ExecutiveOffice\CLAUDE.md`
(the workspace rules for SPFx builds, the CDN model, and the no-PnP provisioning rule).

## What this is

An SPFx web part hosting the v1.x single-file relationship graph, with SharePoint lists
as the single source of truth. The graph engine is **unchanged** from the standalone
application — same look, same iconography, same canvas animations. Only persistence moved.

| Thing | Value |
|---|---|
| Solution id | `e7c41d92-3b6a-4f18-9a52-8d3c07b41e6f` |
| Feature id | `5a2f8c11-96d4-4e73-b8a1-c2f409d7e5b3` |
| Web part component id | `9d1e4b77-2c85-4a36-91f0-6e8b5d3a2c14` |
| CDN folder | `https://irm.azureedge.us/M/enterprise-relationship-graph/` |
| Lane | Heft, SPFx 1.23.2 (no Fluent React — see below) |

## The engine is generated. Do not hand-edit it.

`src/engine/engine.ts` and `src/engine/engineAssets.ts` are produced by
`scripts/extract-engine.py` from `enterprise-relationship-graph.html`, which remains the
source of truth for engine behaviour.

```bash
py -3 scripts/extract-engine.py enterprise-relationship-graph.html src/engine
```

The extractor applies a short, **asserted** list of host seams — each anchor must match
exactly once or the script fails. That is deliberate: a silently-skipped seam ships an
app that looks fine and never saves. If you change the HTML around a seam, the extractor
will tell you.

To change engine behaviour: edit the HTML, re-run the extractor, re-run
`node scripts/build-preview.mjs`, and look at the result.

## No Fluent React, on purpose

The web part uses native semantic HTML plus the engine's own CSS custom properties. This
is a permitted choice under SPFX-ARCHITECTURE.md §2a, and it means the **§2b
Fluent/Tabster pin contract does not apply to this project** — there is no Fluent
dependency closure to hold in place. If Fluent controls are ever added, the entire §2b
contract applies from that moment, so weigh it against writing 30 lines of CSS.

The shell chrome reads `--panel`, `--border`, `--text`, `--accent` from the injected
engine stylesheet, so it follows the engine's own light/dark toggle with no extra wiring.

## One instance per page

The engine addresses its DOM by element id, exactly as the standalone page did — that is
what let it stay byte-identical. Two instances on one page would fight over those ids, so
`mountEngine` refuses the second and says why. Use a **single-part app page**.

## Storage: two models behind one interface

Chosen per project by the `ErgStorageMode` column; nothing above `IGraphStore` knows which.

- **Document** (default): the whole graph in the `ErgPayload1..8` columns of one
  `ERG Projects` item. Every save is one atomic, ETag-guarded write; SharePoint version
  history is a free save ladder. Conflicts resolve by three-way merge.
- **Items**: one row per node and relationship in `ERG Nodes` / `ERG Edges`, with
  change-log sync for multi-editor use. **Layout is never an item** — positions travel as
  one blob on the project item, because a force layout moves every node and as item
  writes that is one operation per node per run.

SharePoint has no transactions. Ordering (edges deleted before nodes, nodes written
before edges) plus an orphan sweep on load is the mitigation, and it is a mitigation, not
a guarantee. That is why Document mode is the default.

### Payload columns must be blanked, not just written

`payloadPatch` explicitly writes `''` to unused trailing columns. Without it, a graph that
shrinks leaves stale chunks behind and the next read concatenates live JSON with dead
JSON. Covered by a test; do not "optimise" it away.

## Provisioning runs in the browser, as the signed-in user

`SpProvisioningService` creates lists and columns through `SPHttpClient`. This is not PnP
and does not need `Connect-PnPOnline`, so it works in this tenant where PnP PowerShell
cannot connect at all. The workspace no-PnP rule is about server-side provisioning; this
is the supported in-app path that rule points to.

The plan is **additive only**. Wrong column types and unreadable lists are reported as
conflicts for a human, never auto-fixed. `planProvisioning` is pure, so the health check
and the fix are literally the same function and cannot drift.

## Tests

```bash
npm run test:safety     # pure logic + full integration against an in-memory SharePoint
npm test                # test:safety, then the Heft/jest lane
node scripts/build-preview.mjs   # reassemble the extracted engine into a page you can open
```

`tests/fakes/FakeSharePoint.ts` implements enough REST — ETags, per-list change logs
including deletes, cascade delete, unique columns — to run the real stores end to end.
Prefer adding to the integration suite over mocking: the failures that matter here are
between components, not inside them.

The safety suite compiles with the project's own TypeScript and runs on plain Node, with
no external test runner (SPFX-ARCHITECTURE.md §4). It is chained into `check`, `test` and
`ship`, so a release cannot go out with this logic broken.

## Fonts

Node icons are Font Awesome glyphs drawn onto the canvas, and labels are Inter — they are
not decoration. They load from `<assetBaseUrl>/fonts/...`, defaulting to the project's CDN
folder, so **the CDN handoff must include the font files**. The engine races the load
against a 3s timeout and renders with fallbacks if they do not arrive.

## Before a production release

Follow the workspace CLAUDE.md release procedure. Project-specific reminders:

1. Bump `package.json` version **and** solution + feature versions in `package-solution.json`.
2. `includeClientSideAssets: false` and confirm `cdnBasePath` is this project's folder.
3. `npm run release` (runs the safety suite, ships, then `eo-spfx preflight`).
4. Font files go in the CDN handoff alongside the bundle.
