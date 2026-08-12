/**
 * Reassemble the EXTRACTED engine modules into a standalone page.
 *
 * This is the check that the extraction is faithful. It deliberately reads
 * src/engine/*.ts — not the original HTML — so that if the extractor ever drops the
 * tail of the stylesheet, mangles the template, or breaks the wrapper, the preview
 * breaks visibly instead of the web part breaking on a production site.
 *
 * It also gives a way to click through the real UI with no SPFx build and no tenant:
 * the standalone host stores the graph in localStorage, exactly as v1.x did.
 *
 *   node scripts/build-preview.mjs && open temp/preview/index.html
 */
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const read = (p) => readFileSync(join(root, p), 'utf8');

/** Pull a template-literal constant back out of the generated assets module. */
function extractLiteral(source, name) {
  const marker = `export const ${name}: string = \``;
  const start = source.indexOf(marker);
  if (start < 0) { throw new Error(`${name} not found in engineAssets.ts`); }
  const from = start + marker.length;

  let i = from;
  for (;;) {
    const tick = source.indexOf('`', i);
    if (tick < 0) { throw new Error(`Unterminated template literal for ${name}`); }
    // Count preceding backslashes; an odd number means this backtick is escaped.
    let slashes = 0;
    while (source[tick - 1 - slashes] === '\\') { slashes++; }
    if (slashes % 2 === 0) {
      return source.slice(from, tick)
        .replace(/\\\$\{/g, '${')
        .replace(/\\`/g, '`')
        .replace(/\\\\/g, '\\');
    }
    i = tick + 1;
  }
}

const assets = read('src/engine/engineAssets.ts');
const css = extractLiteral(assets, 'ENGINE_CSS');
const markup = extractLiteral(assets, 'ENGINE_HTML');

const engineTs = read('src/engine/engine.ts');
const openMarker = 'export function startEngine(host: IErgHost): void {';
const openAt = engineTs.indexOf(openMarker);
if (openAt < 0) { throw new Error('startEngine wrapper not found in engine.ts'); }
const engineBody = engineTs
  .slice(openAt + openMarker.length)
  .replace(/\}\s*$/, '');   // drop the wrapper's closing brace

// Sanity gates. A truncated extraction is the failure mode worth catching here.
const expect = (cond, msg) => { if (!cond) { throw new Error(`Preview build failed: ${msg}`); } };
expect(css.length > 40000, `stylesheet looks truncated (${css.length} chars)`);
expect(markup.includes('</div>'), 'template looks truncated');
expect(engineBody.includes('boot();'), 'engine body is missing its boot() call');
expect(engineBody.includes('function buildEngineApi()'), 'host bridge missing from engine body');

const html = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Enterprise Relationship Graph — extraction preview</title>
<!-- Preview only. In the web part these load from the project's CDN folder. -->
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css" crossorigin="anonymous" />
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" />
<script>(function(){try{var t=localStorage.getItem("erg.theme");if(!t)t=(window.matchMedia&&matchMedia("(prefers-color-scheme: light)").matches)?"light":"dark";if(t==="light")document.documentElement.setAttribute("data-theme","light");}catch(e){}})();</script>
<style>
${css}
</style>
</head>
<body>
${markup}
<script>
function startEngine(host) {
${engineBody}
}

/* Standalone host: the same contract the SPFx shell implements, backed by
   localStorage instead of SharePoint. Proves the seams work without a tenant. */
var SAVE_KEY = "erg.preview.graph";
var SNAP_KEY_PREVIEW = "erg.preview.snapshots";
function readJson(key, fallback) {
  try { var raw = localStorage.getItem(key); return raw ? JSON.parse(raw) : fallback; }
  catch (e) { return fallback; }
}

startEngine({
  mode: "standalone",
  editorName: "Preview user",
  initialGraph: readJson(SAVE_KEY, null),
  initialSnapshots: readJson(SNAP_KEY_PREVIEW, []),
  onGraphChanged: function (graph) {
    try { localStorage.setItem(SAVE_KEY, JSON.stringify(graph)); } catch (e) {}
  },
  onSnapshotsChanged: function (snaps) {
    try { localStorage.setItem(SNAP_KEY_PREVIEW, JSON.stringify(snaps)); } catch (e) {}
  },
  onReady: function (api) {
    window.__ergApi = api;   // handy for poking at the running engine from the console
    console.log("ERG preview ready:", api.getBundle().graph.nodes.length, "nodes");
  }
});
</script>
</body>
</html>
`;

mkdirSync(join(root, 'temp/preview'), { recursive: true });
writeFileSync(join(root, 'temp/preview/index.html'), html, 'utf8');
console.log(
  `Preview written: temp/preview/index.html ` +
  `(css ${css.length}B, markup ${markup.length}B, engine ${engineBody.length}B)`
);
