#!/usr/bin/env python3
"""
Extract the v1.x single-file Enterprise Relationship Graph into SPFx-consumable
modules, applying ONLY the host seams needed to run under a SharePoint web part.

Design rule: the engine body is copied VERBATIM apart from a short, asserted list
of replacements. Every anchor must match an exact expected count or this script
fails -- a silently-skipped seam would ship a subtly broken app.

Outputs (written by the caller via redirect):
  engine.ts          -> export function startEngine(host)
  engineAssets.ts    -> ENGINE_CSS + ENGINE_HTML string constants
"""
import re
import sys
import os

SRC = sys.argv[1]
OUTDIR = sys.argv[2]

with open(SRC, 'r', encoding='utf-8', newline='') as f:
    html = f.read()

# Normalise CRLF -> LF for stable anchor matching (the repo file is CRLF on Windows).
html = html.replace('\r\n', '\n')

# ---------------------------------------------------------------- slice it up
head = html[:html.index('</head>')]
styles = re.findall(r'<style>(.*?)</style>', head, re.S)
assert len(styles) == 2, f'expected 2 <style> blocks in head, found {len(styles)}'
css = '\n'.join(s.strip('\n') for s in styles)

b0 = html.index('<body>') + len('<body>')
b1 = html.index('</body>')
body = html[b0:b1]

s0 = body.index('<script>')
s1 = body.rindex('</script>')
markup = body[:s0].strip('\n')
engine = body[s0 + len('<script>'):s1]

print(f'[slice] css={len(css)}B markup={len(markup)}B engine={len(engine)}B', file=sys.stderr)

# ------------------------------------------------------------- host seams
REPLACEMENTS = []


def seam(name, old, new, count=1):
    REPLACEMENTS.append((name, old, new, count))


# 1. Initial graph comes from the host (a SharePoint project) when one is supplied.
seam(
    'loadInitial-read',
    '  try { const s=localStorage.getItem(STORAGE_KEY); if(s) loaded=JSON.parse(s); } catch(_){}',
    '  if (host && host.initialGraph) { loaded = host.initialGraph; }\n'
    '  else { try { const s=localStorage.getItem(STORAGE_KEY); if(s) loaded=JSON.parse(s); } catch(_){} }'
)

# 2. persist() is the single write seam the whole app funnels through. Under a host
#    it becomes "tell the host the graph changed"; the host owns debounce + transport.
seam(
    'persist',
    'function persist(){ try { syncPositionsToGraph(); localStorage.setItem(STORAGE_KEY, JSON.stringify(state.graph)); } catch(_){} }',
    'function persist(){\n'
    '  try { syncPositionsToGraph(); } catch(_){}\n'
    '  if (host && typeof host.onGraphChanged === "function"){\n'
    '    try { host.onGraphChanged(state.graph); } catch(_){}\n'
    '    return;\n'
    '  }\n'
    '  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state.graph)); } catch(_){}\n'
    '}'
)

# 3. The dedup cleanup inside loadInitial wrote straight to localStorage; route it
#    through persist() so a host-backed load persists its cleanup too.
seam(
    'loadInitial-cleanup-write',
    '    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state.graph)); } catch(_){}',
    '    persist();'
)

# 4. Snapshots live beside the graph in the host store.
seam(
    'snapshots-io',
    'function loadSnapshots(){\n'
    '  try { const s = localStorage.getItem(SNAP_KEY); return s ? JSON.parse(s) : []; } catch(_){ return []; }\n'
    '}\n'
    'function saveSnapshots(snaps){\n'
    '  try { localStorage.setItem(SNAP_KEY, JSON.stringify(snaps)); } catch(_){}\n'
    '}',
    'function loadSnapshots(){\n'
    '  if (host && Array.isArray(host.initialSnapshots)) return host.initialSnapshots;\n'
    '  try { const s = localStorage.getItem(SNAP_KEY); return s ? JSON.parse(s) : []; } catch(_){ return []; }\n'
    '}\n'
    'function saveSnapshots(snaps){\n'
    '  if (host && typeof host.onSnapshotsChanged === "function"){\n'
    '    if (Array.isArray(host.initialSnapshots)) host.initialSnapshots = snaps;\n'
    '    try { host.onSnapshotsChanged(snaps); } catch(_){}\n'
    '    return;\n'
    '  }\n'
    '  try { localStorage.setItem(SNAP_KEY, JSON.stringify(snaps)); } catch(_){}\n'
    '}'
)

# 5. The signed-in SharePoint user replaces the name prompt.
seam(
    'editor-name',
    'function getEditorName(forcePrompt){\n'
    '  let name = localStorage.getItem(EDITOR_NAME_KEY) || "";',
    'function getEditorName(forcePrompt){\n'
    '  if (host && host.editorName) return host.editorName;\n'
    '  let name = localStorage.getItem(EDITOR_NAME_KEY) || "";'
)

# 6. The header chip means "saved to SharePoint", not "saved to a file", under a host.
seam(
    'file-chip',
    'function updateFileChip(){\n'
    '  const chip = document.getElementById("file-chip");\n'
    '  const name = document.getElementById("fc-name");\n'
    '  if (!chip || !name) return;',
    'function updateFileChip(){\n'
    '  const chip = document.getElementById("file-chip");\n'
    '  const name = document.getElementById("fc-name");\n'
    '  if (!chip || !name) return;\n'
    '  if (host && typeof host.renderStatusChip === "function"){\n'
    '    try { host.renderStatusChip(chip, name, state); return; } catch(_){}\n'
    '  }'
)

# 7. Local-file reopen and the first-run "open a JSON file" hint are file-mode only.
seam(
    'boot-tail',
    '  restoreUIState();\n'
    '  tryReopenLast();\n'
    '}',
    '  restoreUIState();\n'
    '  if (!host || host.mode !== "sharepoint") tryReopenLast();\n'
    '  if (host && typeof host.onReady === "function"){\n'
    '    try { host.onReady(buildEngineApi()); } catch(e){ console.error("ERG host.onReady failed", e); }\n'
    '  }\n'
    '}'
)

seam(
    'tour-toast',
    '  if (!localStorage.getItem("erg.tour-seen")){',
    '  if ((!host || host.mode !== "sharepoint") && !localStorage.getItem("erg.tour-seen")){'
)

# 8. The trailing boot() call becomes the exported entry point's body, preceded by
#    the API object the host drives the running engine through.
ENGINE_API = '''
/* ==========================================================================
   HOST BRIDGE — the control surface the SPFx shell drives the running engine
   through. Everything here delegates to existing engine functions; no new
   behaviour, so the graph keeps behaving exactly as the v1.x file did.
   ========================================================================== */
function buildEngineApi(){
  return {
    /** Replace the whole graph (project switch, remote change, snapshot restore). */
    setBundle: function(bundle, label){
      applyLoadedBundle(bundle, label || null);
      state.fileHandle = null;
      state.dirty = false;
      updateFileChip();
    },
    /** The current graph + snapshots, in the on-disk bundle shape. */
    getBundle: function(){ return currentBundle(); },
    /** Graph only, with hand-arranged positions folded in. */
    getGraph: function(){ syncPositionsToGraph(); return state.graph; },
    setProjectLabel: function(label){ state.fileName = label || null; updateFileChip(); },
    isDirty: function(){ return !!state.dirty; },
    markClean: function(){ markClean(); },
    markDirty: function(){ markDirty(); },
    toast: function(msg, kind){ showToast(msg, kind); },
    refresh: function(){ requestRedraw(); },
    fit: function(){ fitGraph(); },
    /** Stop the render loop and drop document-level listeners (web part dispose). */
    destroy: function(){
      __ergDestroyed = true;
      rafActive = false;
      for (var i = 0; i < __ergListeners.length; i++){
        var L = __ergListeners[i];
        try { L.t.removeEventListener(L.e, L.f, L.o); } catch(_){}
      }
      __ergListeners.length = 0;
    }
  };
}

boot();'''

seam('boot-call', '\nboot();', ENGINE_API)

# ------------------------------------------------------------- apply seams
for name, old, new, count in REPLACEMENTS:
    found = engine.count(old)
    if found != count:
        raise SystemExit(
            f'FATAL seam "{name}": expected {count} occurrence(s), found {found}.\n'
            f'--- anchor ---\n{old[:400]}'
        )
    engine = engine.replace(old, new)
    print(f'[seam] {name}: ok ({count})', file=sys.stderr)

# ----------------------------------------------- listener capture for destroy()
# Route every document/window listener through a recorder so the web part can
# unregister them on dispose. Mechanical, and it is what makes destroy() honest.
doc_n = engine.count('document.addEventListener(')
win_n = engine.count('window.addEventListener(')
engine = engine.replace('document.addEventListener(', '__ergOn(document, ')
engine = engine.replace('window.addEventListener(', '__ergOn(window, ')
print(f'[listeners] captured document={doc_n} window={win_n}', file=sys.stderr)
assert doc_n > 0 and win_n > 0, 'expected document and window listeners to exist'

PRELUDE = '''// @ts-nocheck
/* eslint-disable */
/* ==========================================================================
   Enterprise Relationship Graph — GRAPH ENGINE
   ==========================================================================
   Extracted VERBATIM from the v1.7.0 single-file application
   (enterprise-relationship-graph.html) by scripts/extract-engine.py.

   DO NOT hand-edit this file. Change the source HTML and re-run the extractor,
   or the two copies drift and the extractor's seam assertions stop protecting
   anything.

   The only modifications the extractor makes are the host seams: persistence
   (persist / snapshots / editor name / status chip) routes through the `host`
   object instead of localStorage, and boot() hands back a control API. Rendering,
   layout, interaction, iconography and animation are untouched.

   Typing is deliberately suppressed (@ts-nocheck): this is ~9,400 lines of
   working, shipped, browser-tested JavaScript whose value is that it is
   IDENTICAL to the original. Retyping it would be a rewrite wearing a port's
   clothes.
   ========================================================================== */
import { IErgHost, IEngineApi } from './hostContract';

export function startEngine(host: IErgHost): void {
  // --- extractor-injected scaffolding -------------------------------------
  // Frame scheduling is shadowed so destroy() can stop the render loop within
  // one frame: every internal requestAnimationFrame call routes through here.
  var __ergDestroyed = false;
  var __ergRaf = window.requestAnimationFrame.bind(window);
  function requestAnimationFrame(cb){ return __ergDestroyed ? 0 : __ergRaf(cb); }
  // Listener recorder — see destroy() in buildEngineApi().
  var __ergListeners = [];
  function __ergOn(target, evt, fn, opts){
    __ergListeners.push({ t: target, e: evt, f: fn, o: opts });
    return target.addEventListener(evt, fn, opts);
  }
  // ------------------------------------------------------------------------
'''

EPILOGUE = '''
}
'''

engine_ts = PRELUDE + engine + EPILOGUE

# ------------------------------------------------------------- assets module
def as_template_literal(text):
    return text.replace('\\', '\\\\').replace('`', '\\`').replace('${', '\\${')


assets_ts = (
    '/* eslint-disable */\n'
    '/* ==========================================================================\n'
    '   Enterprise Relationship Graph — ENGINE ASSETS (generated; do not hand-edit)\n'
    '   ==========================================================================\n'
    '   The stylesheet and DOM template of the v1.7.0 single-file application,\n'
    '   extracted verbatim by scripts/extract-engine.py.\n'
    '\n'
    '   They are emitted as string constants rather than .css/.html files so no\n'
    '   webpack raw-loader configuration is required: the SPFx rig bundles this\n'
    '   like any other TypeScript module, and the web part injects both at mount.\n'
    '   ========================================================================== */\n\n'
    'export const ENGINE_CSS: string = `' + as_template_literal(css) + '`;\n\n'
    'export const ENGINE_HTML: string = `' + as_template_literal(markup) + '`;\n'
)

os.makedirs(OUTDIR, exist_ok=True)
with open(os.path.join(OUTDIR, 'engine.ts'), 'w', encoding='utf-8', newline='\n') as f:
    f.write(engine_ts)
with open(os.path.join(OUTDIR, 'engineAssets.ts'), 'w', encoding='utf-8', newline='\n') as f:
    f.write(assets_ts)

print(f'[write] engine.ts={len(engine_ts)}B engineAssets.ts={len(assets_ts)}B', file=sys.stderr)
print('OK')
