/* eslint-disable */
/* ==========================================================================
   Enterprise Relationship Graph — ENGINE ASSETS (generated; do not hand-edit)
   ==========================================================================
   The stylesheet and DOM template of the v1.7.0 single-file application,
   extracted verbatim by scripts/extract-engine.py.

   They are emitted as string constants rather than .css/.html files so no
   webpack raw-loader configuration is required: the SPFx rig bundles this
   like any other TypeScript module, and the web part injects both at mount.
   ========================================================================== */

export const ENGINE_CSS: string = `  :root {
    --bg-0:#0D0E11; --bg-1:#131419; --panel:#181A1F; --panel-2:#1C1E24;
    --card:#20232B; --card-2:#272A33;
    --border:rgba(170,166,160,0.20); --border-strong:rgba(170,166,160,0.40);
    --text:#F6F5F3; --text-2:#CFCCC6; --muted:#A6A29A;
    --accent:#6366F1; --accent-2:#A5B4FC;
    --warn:#FBBF24; --ok:#34D399; --crit:#F87171;
    --violet:#A78BFA;
    --shadow-2:0 10px 24px rgba(0,0,0,0.45);
    --radius:10px; --radius-sm:6px;
  }
  :root[data-theme="light"]{
    --bg-0:#F1F0EC; --bg-1:#FAF9F6; --panel:#FEFDFB; --panel-2:#F3F1EB;
    --card:#FEFDFB; --card-2:#F4F2EC;
    --border:rgba(40,36,30,0.14); --border-strong:rgba(40,36,30,0.30);
    --text:#1A1916; --text-2:#46443E; --muted:#6E6B62;
    --accent:#4338CA; --accent-2:#3730A3;
    --warn:#B45309; --ok:#047857; --crit:#C0392B;
    --violet:#6D28D9;
    --shadow-2:0 10px 24px rgba(40,36,30,0.16);
  }
  :root[data-theme="light"] .ctx-item.danger{ color:var(--crit); }
  :root[data-theme="light"] .ctx-item.danger:hover{ background:rgba(192,57,43,0.12); }
  :root[data-theme="light"] .toast.err{ color:var(--crit); }
  :root[data-theme="light"] .toast.ok{ color:var(--ok); }
  :root[data-theme="light"] .rel-warn{ color:#7A4E0A; }
  :root[data-theme="light"] .rel-warn .rel-suggest{ color:#92560B; }
  :root[data-theme="light"] .btn.primary,
  :root[data-theme="light"] .header-btn:hover,
  :root[data-theme="light"] .tool-btn:hover,
  :root[data-theme="light"] .tool-btn.active{ color:#FDF6EA; }
  :root[data-theme="light"] .canvas-toolbar{ background:rgba(252,251,248,0.90); }
  /* The base .btn:hover / .rel-item:hover use hardcoded dark navies that
     rendered almost-black-on-dark-text in light mode. */
  :root[data-theme="light"] .btn:hover{ background:rgba(99,102,241,0.10); }
  :root[data-theme="light"] .rel-item:hover{ background:rgba(99,102,241,0.08); }
  *{box-sizing:border-box}
  html,body{height:100%}
  body{
    margin:0;
    background: radial-gradient(1200px 800px at 80% -10%, rgba(99,102,241,0.08), transparent 60%),
                radial-gradient(900px 700px at -10% 110%, rgba(167,139,250,0.07), transparent 55%),
                var(--bg-0);
    color:var(--text);
    font-family: "Inter", system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    font-size:14px; line-height:1.45; overflow:hidden; -webkit-font-smoothing:antialiased;
  }
  ::selection{background:rgba(99,102,241,0.35)}
  button,input,select,textarea{font-family:inherit;color:inherit}
  a{color:var(--accent-2);text-decoration:none}
  a:hover{text-decoration:underline}

  .app{
    display:grid;
    grid-template-columns:280px 1fr 360px;
    grid-template-rows:64px 1fr 32px;
    grid-template-areas:"header header header" "sidebar canvas details" "footer footer footer";
    height:100vh; width:100vw;
  }

  header.app-header{
    grid-area:header;
    display:flex; align-items:center; justify-content:space-between;
    padding:0 18px;
    background: linear-gradient(180deg, var(--card-2), var(--panel));
    border-bottom:1px solid var(--border); backdrop-filter:blur(6px);
  }
  .brand{display:flex; align-items:center; gap:12px}
  .brand-mark{
    width:34px; height:34px; border-radius:9px;
    background: linear-gradient(135deg, var(--accent) 0%, var(--violet) 100%);
    display:grid; place-items:center; color:#04101F; font-size:15px;
    box-shadow: 0 0 0 1px rgba(255,255,255,0.06), 0 6px 16px rgba(99,102,241,0.25);
  }
  .brand-text h1{
    font-size:16px; margin:0; font-weight:700; letter-spacing:0.2px;
    color:var(--text);
  }
  .brand-text p{font-size:11.5px; margin:0; color:var(--muted)}
  .header-right{display:flex; align-items:center; gap:10px}
  .stats{
    display:flex; align-items:center; gap:0;
    background:var(--card); border:1px solid var(--border);
    border-radius:9px; padding:4px 4px; font-size:12px;
  }
  .stat{
    display:flex; align-items:baseline; gap:5px;
    padding:4px 10px;
    border-right:1px solid var(--border);
  }
  .stat:last-child{border-right:none}
  .stat .val{font-weight:600; color:var(--text); font-variant-numeric:tabular-nums}
  .stat .lbl{font-size:10.5px; color:var(--muted); text-transform:lowercase; letter-spacing:0.2px}
  .stat.accent .val{color:var(--accent)}
  .file-chip{
    display:inline-flex; align-items:center; gap:7px;
    background:var(--card); border:1px solid var(--border);
    border-radius:8px; padding:5px 11px;
    color:var(--text-2); font-size:12px;
    cursor:pointer; transition:background 0.15s, border-color 0.15s;
    max-width:280px;
  }
  .file-chip:hover{background:var(--card-2); border-color:var(--border-strong); color:var(--text)}
  .file-chip .fc-icon{color:var(--accent-2); font-size:13px}
  .file-chip .fc-name{
    font-weight:500; color:var(--text);
    white-space:nowrap; overflow:hidden; text-overflow:ellipsis; max-width:180px;
  }
  .file-chip .fc-name.no-file{color:var(--muted); font-style:italic; font-weight:400}
  .file-chip .fc-dirty{
    width:7px; height:7px; border-radius:50%;
    background:var(--warn); box-shadow:0 0 6px var(--warn);
    display:none;
  }
  .file-chip.dirty .fc-dirty{display:inline-block}
  .file-chip.dirty .fc-name{color:var(--warn)}
  .modified-chip{
    display:inline-flex; align-items:center; gap:6px;
    background:transparent; border:1px solid var(--border);
    border-radius:8px; padding:5px 10px;
    color:var(--text-2); font-size:11.5px;
    cursor:pointer; transition:background 0.15s, border-color 0.15s;
    max-width:260px;
  }
  .modified-chip:hover{background:var(--card); border-color:var(--border-strong); color:var(--text)}
  .modified-chip .mc-icon{color:var(--muted); font-size:11px}
  .modified-chip .mc-text{white-space:nowrap; overflow:hidden; text-overflow:ellipsis}
  /* drag-and-drop overlay */
  .drop-overlay{
    position:fixed; inset:0; z-index:300;
    display:none; align-items:center; justify-content:center;
    background:rgba(7,17,31,0.85); backdrop-filter:blur(4px);
    pointer-events:none;
  }
  .drop-overlay.active{display:flex}
  .drop-overlay .drop-card{
    border:2px dashed var(--accent);
    background:var(--panel);
    padding:32px 48px; border-radius:14px;
    text-align:center;
  }
  .drop-overlay .drop-card i{font-size:32px; color:var(--accent); margin-bottom:10px; display:block}
  .drop-overlay .drop-card h3{margin:0 0 6px; color:var(--text); font-size:16px}
  .drop-overlay .drop-card p{margin:0; color:var(--muted); font-size:13px}
  .header-btn{
    width:40px; height:40px; border-radius:9px;
    background:var(--card); border:1px solid var(--border);
    color:var(--text-2); cursor:pointer; display:grid; place-items:center;
    font-size:14px; transition:background-color 0.15s, border-color 0.15s, color 0.15s;
    position:relative;
  }
  .header-btn::before{content:""; position:absolute; inset:-2px; border-radius:11px}
  .header-btn:hover{background:var(--accent); color:#F5F3FF; border-color:var(--accent)}

  aside.sidebar{
    grid-area:sidebar; background:var(--panel);
    border-right:1px solid var(--border); overflow-y:auto; padding:14px;
  }
  .section{
    margin-bottom:14px; background:var(--card);
    border:1px solid var(--border); border-radius:var(--radius); overflow:hidden;
  }
  .section-head{
    padding:10px 12px; font-size:11px; font-weight:700;
    text-transform:uppercase; letter-spacing:0.8px; color:var(--text-2);
    border-bottom:1px solid var(--border);
    display:flex; align-items:center; justify-content:space-between;
    user-select:none; cursor:pointer;
    width:100%; background:transparent; border-left:none; border-right:none; border-top:none;
    font-family:inherit; text-align:left;
  }
  .section-head:hover{color:var(--text)}
  .section-head:focus-visible{
    outline:2px solid var(--accent); outline-offset:-2px;
  }
  .section-head .icon{color:var(--accent); font-size:12px}
  .section-head .chev{color:var(--muted); font-size:11px; transition:transform 0.28s cubic-bezier(0.16,1,0.3,1)}
  .section.collapsed .chev{transform:rotate(-90deg)}
  .section-body{
    padding:10px 12px;
    overflow:hidden;
    max-height:1200px;
    opacity:1;
    transition:max-height 0.28s cubic-bezier(0.16,1,0.3,1),
               opacity 0.18s ease-out,
               padding-top 0.28s cubic-bezier(0.16,1,0.3,1),
               padding-bottom 0.28s cubic-bezier(0.16,1,0.3,1);
  }
  .section.collapsed .section-body{
    max-height:0;
    opacity:0;
    padding-top:0;
    padding-bottom:0;
  }
  .section-body.scroll{max-height:180px; overflow-y:auto}

  .field{display:flex; flex-direction:column; gap:4px; margin-bottom:8px}
  .field label{font-size:11px; color:var(--muted)}
  .field-hint{font-size:10.5px; color:var(--accent-2); font-style:italic; margin-left:4px}
  .form-callout{
    background:rgba(99,102,241,0.08); border:1px solid rgba(99,102,241,0.25);
    border-radius:7px; padding:8px 10px; margin-bottom:10px;
    font-size:11.5px; color:var(--text-2); display:flex; gap:8px; align-items:flex-start;
  }
  .form-callout i{color:var(--accent); margin-top:1px}
  .form-callout.warn{
    background:rgba(251,191,36,0.08); border-color:rgba(251,191,36,0.30);
  }
  .form-callout.warn i{color:var(--warn)}
  input[type="text"], input[type="url"], input[type="email"],
  input[type="tel"], input[type="search"], input[type="number"],
  input[type="password"], input[type="date"], input[type="time"],
  select, textarea{
    background:var(--bg-1); border:1px solid var(--border); color:var(--text);
    padding:7px 9px; border-radius:var(--radius-sm); font-size:13px;
    outline:none; transition:border-color 0.15s, box-shadow 0.15s;
    font-family:inherit;
  }
  input:focus, select:focus, textarea:focus{
    border-color:var(--accent); box-shadow:0 0 0 2px rgba(99,102,241,0.25);
    outline:none;
  }
  input:focus-visible, select:focus-visible, textarea:focus-visible{
    box-shadow:0 0 0 2px rgba(99,102,241,0.35);
  }
  textarea{resize:vertical; min-height:60px}
  /* Native select dropdown popup — Chrome/Edge respect background/color on
     options; Safari and some Firefox versions partially apply them. The
     closed select itself stays styled by the rule above. */
  select{
    /* Strip default browser appearance so background/font are respected */
    appearance:none; -webkit-appearance:none; -moz-appearance:none;
    background-image:linear-gradient(45deg, transparent 50%, var(--muted) 50%),
                     linear-gradient(135deg, var(--muted) 50%, transparent 50%);
    background-position:calc(100% - 14px) center, calc(100% - 9px) center;
    background-size:5px 5px, 5px 5px;
    background-repeat:no-repeat;
    padding-right:28px;
  }
  select::-ms-expand{display:none}
  option, optgroup{
    background:var(--panel); color:var(--text);
    padding:4px 8px; font-family:inherit;
  }
  optgroup{
    color:var(--accent-2); font-weight:600; font-style:normal;
    text-transform:uppercase; letter-spacing:0.5px; font-size:11px;
  }
  option{font-style:normal}
  option:checked{
    background:var(--card-2); color:var(--text);
  }
  /* Datalist popups are browser-rendered and unstyleable beyond the input
     itself, so nothing extra needed here. */
  .field.with-icon{position:relative}
  .field.with-icon input{padding-left:30px}
  .field.with-icon .field-icon{
    position:absolute; left:9px; top:50%; transform:translateY(-50%);
    color:var(--muted); font-size:12px;
  }

  .btn{
    display:inline-flex; align-items:center; justify-content:center; gap:6px;
    background:var(--card-2); border:1px solid var(--border); color:var(--text);
    padding:7px 10px; border-radius:var(--radius-sm); cursor:pointer;
    font-size:12.5px; font-family:inherit;
    transition:background 0.15s, border-color 0.15s, transform 0.05s; user-select:none;
  }
  .btn:hover{background:#1F2B47; border-color:var(--border-strong)}
  .btn:active{transform:translateY(1px)}
  .btn:focus-visible, .tool-btn:focus-visible, .header-btn:focus-visible,
  .legend-tab:focus-visible, .legend-toggle:focus-visible,
  .modal-close:focus-visible, .icon-btn:focus-visible, .help-tab:focus-visible {
    outline:2px solid var(--accent); outline-offset:2px;
  }
  .tool-btn:focus-visible, .header-btn:focus-visible{ outline-offset:1px }
  .btn:disabled{opacity:0.4; cursor:not-allowed}
  .btn.active{
    background:rgba(99,102,241,0.18); color:var(--accent-2);
    border-color:rgba(99,102,241,0.55);
  }
  .btn.active:hover{background:rgba(99,102,241,0.25)}
  .btn.primary{background:#6366F1; border-color:#5458E0; color:#F5F3FF}
  .btn.primary:hover{background:var(--accent-2)}
  .btn.warn{background:#3a2a10; border-color:#6b4a17; color:#FCD34D}
  .btn.danger{background:#3a1414; border-color:#6b1d1d; color:#FCA5A5}
  .btn.ghost{background:transparent}
  .btn-row{display:grid; grid-template-columns:1fr 1fr; gap:6px}
  .btn-row.three{grid-template-columns:1fr 1fr 1fr}
  .btn.block{width:100%}

  .checkbox-row{
    display:flex; align-items:center; gap:8px;
    font-size:12px; color:var(--text-2); padding:4px 0; cursor:pointer;
  }
  .checkbox-row input{accent-color:var(--accent)}
  .checkbox-row .swatch{
    width:10px; height:10px; border-radius:50%;
    box-shadow:0 0 0 1px rgba(255,255,255,0.15);
  }
  .checkbox-row .type-icon{
    width:14px; text-align:center; color:var(--text-2); font-size:11px;
  }
  .filter-list{display:flex; flex-direction:column; gap:2px}
  .filter-controls{display:flex; gap:6px; margin-bottom:8px; padding-bottom:8px; border-bottom:1px solid var(--border)}
  .filter-controls .btn{padding:4px 8px; font-size:11px; flex:1}

  main.canvas-area{
    grid-area:canvas; position:relative; overflow:hidden;
    background:
      radial-gradient(800px 600px at 30% 20%, rgba(99,102,241,0.05), transparent 60%),
      radial-gradient(700px 500px at 80% 90%, rgba(167,139,250,0.05), transparent 65%),
      var(--bg-1);
  }
  #graph-canvas{position:absolute; inset:0; width:100%; height:100%; cursor:grab; outline:none}
  #graph-canvas.dragging{cursor:grabbing}
  #graph-canvas.connecting{cursor:crosshair}
  #graph-canvas.layout-mode{cursor:move}
  #graph-canvas:focus-visible{outline:2px solid var(--accent); outline-offset:-2px}

  .canvas-toolbar{
    position:absolute; top:14px; left:14px;
    display:flex; gap:6px; align-items:center;
    background:rgba(24,26,31,0.88); border:1px solid var(--border);
    padding:6px; border-radius:10px; backdrop-filter:blur(6px); z-index:5;
  }
  .canvas-toolbar .layout-select{
    padding:5px 8px; font-size:12px;
    background:var(--bg-1); border:1px solid var(--border); border-radius:6px;
  }
  .tool-btn{
    width:36px; height:36px;
    background:var(--card-2); border:1px solid var(--border);
    border-radius:6px; cursor:pointer; color:var(--text-2);
    display:grid; place-items:center; font-size:13px;
    transition:background 0.15s, color 0.15s; font-family:inherit;
    position:relative;
  }
  /* Expand pointer/touch hit area to 44x44 without changing visual size */
  .tool-btn::before{content:""; position:absolute; inset:-4px; border-radius:8px}
  .tool-btn:hover{background:var(--accent); color:#F5F3FF}
  .tool-btn.active{background:var(--accent); color:#F5F3FF}
  .tool-btn:disabled{opacity:0.35; cursor:not-allowed}
  .tool-btn:disabled:hover{background:var(--card-2); color:var(--text-2)}
  .tool-divider{width:1px; height:18px; background:var(--border); margin:0 2px}

  .canvas-banner{
    position:absolute; top:14px; left:50%; transform:translateX(-50%);
    background:var(--panel);
    border:1px solid rgba(99,102,241,0.5);
    color:var(--text); padding:7px 14px; border-radius:8px;
    font-size:12.5px; z-index:6;
    display:flex; align-items:center; gap:10px;
    box-shadow:var(--shadow-2);
    animation:slideDown 0.2s ease-out;
  }
  @keyframes slideDown{from{transform:translateX(-50%) translateY(-12px); opacity:0} to{transform:translateX(-50%) translateY(0); opacity:1}}
  .canvas-banner .ban-close{
    background:none; border:none; color:var(--accent-2); cursor:pointer; font-size:13px;
  }

  .legend{
    position:absolute; bottom:14px; left:14px;
    background:var(--panel); border:1px solid var(--border);
    border-radius:10px; padding:10px 12px; max-width:320px;
    z-index:5; box-shadow:var(--shadow-2);
  }
  .legend h4{
    margin:0 0 8px 0; font-size:11px; padding-right:42px;
    text-transform:uppercase; letter-spacing:0.8px; color:var(--text-2);
  }
  /* When collapsed, only the title shows; keep its bottom margin from adding
     dead space below the lone heading. */
  .legend.collapsed h4{ margin-bottom:0; }
  .legend-tabs{display:flex; gap:4px; margin-bottom:8px}
  .legend-tab{
    background:transparent; border:1px solid var(--border); color:var(--muted);
    padding:3px 8px; border-radius:5px; font-size:10.5px; cursor:pointer;
  }
  .legend-tab.active{background:rgba(99,102,241,0.15); color:var(--accent-2); border-color:rgba(99,102,241,0.4)}
  .legend-grid{display:grid; grid-template-columns:repeat(2,1fr); gap:4px 12px}
  .legend-item{display:flex; align-items:center; gap:6px; font-size:11px; color:var(--muted)}
  .legend-icon{
    width:16px; height:16px; border-radius:4px; display:grid; place-items:center;
    font-size:9px; color:#04101F; box-shadow:0 0 0 1px rgba(255,255,255,0.12);
  }
  .legend-line{width:24px; height:0; border-bottom-width:2px; border-bottom-style:solid}
  .legend-toggle{
    position:absolute; top:6px; right:8px; cursor:pointer;
    color:var(--muted); font-size:10px; background:none; border:none;
  }
  .legend-toggle:hover{color:var(--accent)}
  .legend.collapsed .legend-grid, .legend.collapsed .legend-tabs{display:none}

  .tooltip{
    position:fixed; background:var(--panel-2); border:1px solid var(--border-strong);
    color:var(--text); padding:8px 11px; border-radius:7px;
    font-size:12px; pointer-events:none; z-index:50;
    max-width:240px; box-shadow:var(--shadow-2);
    opacity:0; transition:opacity 0.12s; transform:translate(-50%, -100%);
  }
  .tooltip.show{opacity:1}
  .tooltip-title{font-weight:600; font-size:12.5px}
  .tooltip-type{color:var(--accent-2); font-size:10.5px; text-transform:uppercase; letter-spacing:0.5px; margin-top:2px}
  .tooltip-meta{color:var(--muted); font-size:11px; margin-top:4px}

  aside.details{
    grid-area:details; background:var(--panel);
    border-left:1px solid var(--border); overflow-y:auto; padding:16px;
  }
  .empty-state{
    display:flex; flex-direction:column; align-items:center; justify-content:center;
    height:70%; color:var(--muted); text-align:center; gap:10px;
  }
  .empty-state .glyph{
    width:56px; height:56px; border-radius:14px;
    background:linear-gradient(135deg, rgba(99,102,241,0.15), rgba(167,139,250,0.15));
    border:1px solid var(--border); display:grid; place-items:center;
    font-size:22px; color:var(--accent);
  }
  .empty-state h3{margin:0; font-size:14px; color:var(--text)}
  .empty-state p{margin:0; font-size:12px; max-width:260px}
  .empty-state .hint{
    margin-top:6px; padding:6px 10px; background:var(--card);
    border:1px solid var(--border); border-radius:7px;
    font-size:11px; color:var(--text-2);
  }
  .empty-state .hint kbd{
    background:var(--card-2); border:1px solid var(--border-strong);
    border-radius:4px; padding:1px 5px; font-size:10.5px; color:var(--text);
    font-family:ui-monospace,monospace;
  }

  .profile-header{margin-bottom:14px}
  .type-pill{
    display:inline-flex; align-items:center; gap:6px;
    padding:3px 9px; border-radius:999px;
    font-size:10.5px; font-weight:600;
    text-transform:uppercase; letter-spacing:0.6px;
    background:rgba(99,102,241,0.15); color:var(--accent-2);
    border:1px solid rgba(99,102,241,0.3);
  }
  .type-pill .ti{font-size:10px}
  .profile-name{font-size:20px; font-weight:700; margin:8px 0 4px; color:var(--text); line-height:1.25}
  .status-badge{
    display:inline-flex; align-items:center; gap:5px;
    padding:2px 8px; font-size:11px; border-radius:6px;
    background:rgba(52,211,153,0.13); color:var(--ok);
    border:1px solid rgba(52,211,153,0.3);
  }
  .status-badge.inactive{background:rgba(248,113,113,0.13); color:var(--crit); border-color:rgba(248,113,113,0.3)}
  .status-badge.planned{background:rgba(251,191,36,0.13); color:var(--warn); border-color:rgba(251,191,36,0.3)}
  .profile-desc{color:var(--text-2); font-size:13px; margin:10px 0 4px}

  .profile-section{
    margin-top:14px; background:var(--card); border:1px solid var(--border);
    border-radius:10px; padding:10px 12px;
  }
  .profile-section h4{
    margin:0 0 8px; font-size:11px;
    text-transform:uppercase; letter-spacing:0.7px; color:var(--accent-2);
    display:flex; align-items:center; gap:6px;
  }
  .profile-section h4 .count{
    background:rgba(99,102,241,0.18); color:var(--accent-2);
    border-radius:999px; font-size:10px; padding:1px 6px;
    border:1px solid rgba(99,102,241,0.25);
  }
  .kv-grid{display:grid; grid-template-columns:110px 1fr; gap:6px 10px; font-size:12.5px}
  .kv-grid .k{color:var(--muted)}
  .kv-grid .v{color:var(--text); word-break:break-word}
  .tag-list{display:flex; flex-wrap:wrap; gap:5px}
  .tag{
    background:rgba(165,180,252,0.12); color:var(--accent-2);
    font-size:11px; padding:2px 8px; border-radius:999px;
    border:1px solid rgba(165,180,252,0.3);
  }
  .rel-list{display:flex; flex-direction:column; gap:5px}
  .rel-item{
    display:flex; align-items:center; justify-content:space-between; gap:8px;
    padding:6px 8px; background:var(--bg-1); border:1px solid var(--border);
    border-radius:7px; cursor:pointer;
    transition:background 0.12s, border-color 0.12s;
  }
  .rel-item:hover{background:#19243B; border-color:var(--border-strong)}
  .rel-item:focus-visible{outline:2px solid var(--accent); outline-offset:-1px}
  .rel-item .rel-left{display:flex; align-items:center; gap:8px; min-width:0; flex:1}
  .rel-item .ti{
    width:18px; height:18px; border-radius:5px; display:grid; place-items:center;
    font-size:9px; color:#04101F;
  }
  .rel-item .rel-node-name{
    font-size:12.5px; color:var(--text);
    white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
  }
  .rel-item .rel-node-type{font-size:10px; color:var(--muted); text-transform:uppercase; letter-spacing:0.4px}
  .rel-item .rel-type{
    font-size:10px; background:rgba(167,139,250,0.13); color:#C4B5FD;
    border:1px solid rgba(167,139,250,0.3);
    padding:1px 6px; border-radius:4px;
    text-transform:uppercase; letter-spacing:0.4px; white-space:nowrap;
  }
  .rel-item .arrow{color:var(--muted); font-size:11px}
  .notes-block{
    background:var(--bg-1); border:1px solid var(--border); border-radius:7px;
    padding:8px 10px; font-size:12.5px; color:var(--text-2); white-space:pre-wrap;
  }
  .empty-line{color:var(--muted); font-style:italic; font-size:12px}

  footer.status-bar{
    grid-area:footer; background:var(--panel-2); border-top:1px solid var(--border);
    display:flex; align-items:center; justify-content:space-between;
    padding:0 14px; font-size:11.5px; color:var(--muted);
  }
  .status-bar .left, .status-bar .center, .status-bar .right{display:flex; gap:14px; align-items:center}
  .status-bar .center{color:var(--muted); flex:0 1 auto}
  .status-bar .center strong{color:var(--text-2); font-weight:500}
  .status-bar .center a{color:var(--accent-2); display:inline-flex; align-items:center; gap:5px}
  .status-bar .center a:hover{color:var(--accent)}
  .status-bar .center a i{font-size:13px}
  .status-dot{
    display:inline-block; width:7px; height:7px; border-radius:50%;
    background:var(--ok); box-shadow:0 0 8px var(--ok); margin-right:5px; vertical-align:middle;
  }

  .modal-backdrop{
    position:fixed; inset:0; background:rgba(7,17,31,0.7);
    backdrop-filter:blur(3px); display:none;
    align-items:center; justify-content:center; z-index:100;
  }
  .modal-backdrop.open{display:flex}
  .modal{
    background:var(--panel); border:1px solid var(--border);
    border-radius:14px; width:480px; max-width:92vw; max-height:86vh;
    overflow:hidden; box-shadow:var(--shadow-2); display:flex; flex-direction:column;
  }
  .modal.wide{width:720px}
  .modal-head{
    padding:14px 16px; border-bottom:1px solid var(--border);
    display:flex; justify-content:space-between; align-items:center;
  }
  .modal-head h3{margin:0; font-size:14px; color:var(--text); display:flex; align-items:center; gap:8px}
  .modal-close{background:none; border:none; color:var(--muted); cursor:pointer; font-size:18px}
  .modal-close:hover{color:var(--crit)}
  .modal-body{padding:14px 16px; overflow-y:auto}
  .modal-foot{padding:12px 16px; border-top:1px solid var(--border);
              display:flex; justify-content:flex-end; gap:8px}

  /* Help drawer */
  .help-tabs{
    display:flex; gap:2px; border-bottom:1px solid var(--border);
    margin:-14px -16px 14px -16px; padding:0 12px;
    overflow-x:auto;
  }
  .help-tab{
    background:transparent; border:none; color:var(--muted);
    padding:10px 12px; cursor:pointer; font-size:12px;
    border-bottom:2px solid transparent; font-family:inherit;
    white-space:nowrap;
  }
  .help-tab:hover{color:var(--text-2)}
  .help-tab.active{color:var(--accent); border-bottom-color:var(--accent)}
  .help-pane{display:none}
  .help-pane.active{display:block}
  .help-pane h4{
    margin:14px 0 6px; color:var(--accent-2);
    font-size:12.5px; text-transform:uppercase; letter-spacing:0.6px;
  }
  .help-pane h4:first-child{margin-top:0}
  .help-pane p{margin:0 0 10px; color:var(--text-2); font-size:13px; line-height:1.55}
  .help-pane ul{margin:0 0 10px 18px; padding:0; color:var(--text-2); font-size:13px}
  .help-pane li{margin-bottom:5px; line-height:1.5}
  .help-pane kbd{
    background:var(--card-2); border:1px solid var(--border-strong);
    border-radius:4px; padding:1px 6px; font-size:11px; color:var(--text);
    font-family:ui-monospace,monospace;
  }
  .help-pane .step-grid{
    display:grid; grid-template-columns:32px 1fr; gap:8px 12px;
    margin:8px 0 12px;
  }
  .help-pane .step-num{
    width:26px; height:26px; border-radius:50%;
    background:rgba(99,102,241,0.15); color:var(--accent);
    display:grid; place-items:center; font-weight:600; font-size:12px;
    border:1px solid rgba(99,102,241,0.35);
  }
  .help-pane .step-text{font-size:13px; color:var(--text-2); align-self:center}
  .ref-grid{display:grid; grid-template-columns:160px 1fr; gap:6px 12px; font-size:12.5px}
  .ref-grid .rk{color:var(--text); font-weight:500}
  .ref-grid .rv{color:var(--muted)}
  .help-pane pre{
    background:var(--card); border:1px solid var(--border);
    border-radius:6px; padding:10px 12px; margin:0 0 10px;
    font-family:ui-monospace, "SF Mono", Menlo, Consolas, monospace;
    font-size:12px; line-height:1.5; color:var(--text-2);
    overflow-x:auto;
  }
  .help-pane code{
    background:rgba(99,102,241,0.08); border:1px solid rgba(99,102,241,0.18);
    border-radius:4px; padding:1px 5px;
    font-family:ui-monospace, "SF Mono", Menlo, Consolas, monospace;
    font-size:11.5px; color:var(--text);
  }
  .kbd-grid{display:grid; grid-template-columns:90px 1fr; gap:6px 14px; font-size:12.5px}
  .kbd-grid .rv{color:var(--text-2)}

  /* Manage Types modal */
  .mt-section{margin-bottom:14px}
  .mt-section h4{
    margin:0 0 8px; font-size:11px; text-transform:uppercase; letter-spacing:0.7px;
    color:var(--accent-2); display:flex; align-items:center; justify-content:space-between; gap:6px;
  }
  .mt-list{display:flex; flex-direction:column; gap:5px}
  .mt-row{
    display:flex; align-items:center; gap:10px;
    padding:7px 10px; background:var(--bg-1);
    border:1px solid var(--border); border-radius:7px;
  }
  /* Universal swatch box: CSS Grid place-items:center is more reliable than
     flex+text-align for centering FA icons, which have varying internal
     baselines and glyph positions. The previous width:1em + text-align:center
     trick was causing visible off-centering for asymmetric glyphs like cubes,
     handshake, and chart icons. */
  .mt-swatch{
    width:26px; height:26px; border-radius:6px;
    display:inline-grid; place-items:center;
    color:#04101F; box-shadow:0 0 0 1px rgba(255,255,255,0.12);
    flex-shrink:0; line-height:1; font-size:14px;
  }
  .mt-swatch i{
    line-height:1; display:block;
  }
  .mt-row .mt-name{flex:1; min-width:0; font-size:12.5px; color:var(--text)}
  .mt-row .mt-meta{font-size:10.5px; color:var(--muted); white-space:nowrap}
  .mt-row .mt-actions{display:flex; gap:4px}
  .mt-builtin-grid{
    display:grid; grid-template-columns:repeat(auto-fill, minmax(150px,1fr));
    gap:5px;
  }
  .mt-builtin{
    display:flex; align-items:center; gap:8px;
    padding:6px 9px; background:var(--bg-1);
    border:1px solid var(--border); border-radius:6px;
    font-size:11.5px; color:var(--text-2);
  }
  .mt-builtin .mt-swatch{width:24px; height:24px; font-size:12px; border-radius:5px}
  .mt-reset-btn{
    background:transparent; border:1px solid var(--border);
    color:var(--accent-2); cursor:pointer;
    padding:2px 6px; border-radius:4px; font-size:10px;
    margin-left:auto; flex-shrink:0;
    display:inline-flex; align-items:center; gap:3px;
    transition:background 0.15s, color 0.15s, border-color 0.15s;
  }
  .mt-reset-btn:hover{
    background:rgba(99,102,241,0.12); color:var(--accent);
    border-color:var(--accent);
  }
  .mt-editor{
    background:var(--card-2); border:1px solid var(--accent);
    border-radius:9px; padding:12px; margin-bottom:12px;
  }
  .mt-editor h5{
    margin:0 0 10px; font-size:12px; color:var(--accent-2);
    text-transform:uppercase; letter-spacing:0.6px;
  }
  .color-palette, .icon-palette{
    display:flex; flex-wrap:wrap; gap:5px;
    padding:6px; background:var(--bg-1);
    border:1px solid var(--border); border-radius:6px;
  }
  .color-swatch{
    width:24px; height:24px; border-radius:5px; cursor:pointer;
    border:2px solid transparent; transition:transform 0.1s, border-color 0.1s;
    box-shadow:0 0 0 1px rgba(255,255,255,0.12);
  }
  .color-swatch:hover{transform:scale(1.15)}
  .color-swatch.selected{border-color:var(--text); transform:scale(1.1)}
  .icon-swatch{
    width:30px; height:30px; border-radius:5px; cursor:pointer;
    border:2px solid transparent; transition:background 0.1s, border-color 0.1s;
    background:var(--card); display:grid; place-items:center;
    color:var(--text-2); font-size:13px;
  }
  .icon-swatch:hover{background:var(--card-2); color:var(--accent)}
  .icon-swatch.selected{border-color:var(--accent); background:rgba(99,102,241,0.18); color:var(--accent)}
  .mt-preview-row{
    display:flex; align-items:center; gap:14px; padding:8px 10px;
    background:var(--bg-0); border:1px solid var(--border);
    border-radius:7px; margin:8px 0;
  }
  .mt-preview-row .preview-label{font-size:11px; color:var(--muted)}
  /* Context menu */
  .ctx-menu{
    position:fixed; background:var(--panel); border:1px solid var(--border-strong);
    border-radius:9px; box-shadow:var(--shadow-2);
    padding:5px; min-width:200px; z-index:90; display:none;
  }
  .ctx-menu.open{display:block}
  .ctx-item{
    display:flex; align-items:center; gap:10px;
    padding:7px 12px; font-size:12.5px; color:var(--text-2);
    border-radius:6px; cursor:pointer;
  }
  .ctx-item:hover{background:var(--card-2); color:var(--text)}
  .ctx-item:focus-visible{background:var(--card-2); color:var(--text); outline:2px solid var(--accent); outline-offset:-2px}
  .ctx-item.danger{color:#FCA5A5}
  .ctx-item.danger:hover{background:#3a1414}
  .ctx-item .ci{width:14px; text-align:center; font-size:11px}
  .ctx-sep{height:1px; background:var(--border); margin:4px 6px}

  .toast{
    position:fixed; bottom:50px; left:50%; transform:translateX(-50%);
    background:var(--card-2); color:var(--text);
    border:1px solid var(--border-strong);
    padding:8px 14px; border-radius:8px; font-size:12.5px;
    box-shadow:var(--shadow-2);
    opacity:0; pointer-events:none;
    transition:opacity 0.2s, transform 0.2s; z-index:200;
    display:flex; align-items:center; gap:8px;
  }
  .toast.show{opacity:1; transform:translateX(-50%) translateY(-6px)}
  .toast.err{border-color:rgba(248,113,113,0.45); color:#FCA5A5}
  .toast.ok{border-color:rgba(52,211,153,0.45); color:#A7F3D0}

  /* Snapshot picker */
  .snap-list{
    max-height:180px; overflow-y:auto;
    background:var(--bg-1); border:1px solid var(--border);
    border-radius:7px; padding:4px;
  }
  .snap-item{
    display:flex; align-items:center; justify-content:space-between;
    padding:7px 9px; border-radius:6px; cursor:pointer; gap:8px;
  }
  .snap-item:hover{background:var(--card-2)}
  .snap-item .snap-name{font-size:12.5px; color:var(--text); flex:1; min-width:0}
  .snap-item .snap-meta{font-size:10.5px; color:var(--muted); white-space:nowrap}
  .snap-item .snap-actions{display:flex; gap:4px}
  .snap-item.wt-item{flex-direction:column; align-items:stretch; gap:7px}
  .snap-item.wt-item .snap-actions{justify-content:flex-end}
  .icon-btn{
    width:32px; height:32px; border:none; background:transparent;
    color:var(--muted); cursor:pointer; border-radius:4px; font-size:12px;
    position:relative;
  }
  .icon-btn::before{content:""; position:absolute; inset:-6px; border-radius:8px}
  .icon-btn:hover{background:var(--card); color:var(--text)}
  .icon-btn.danger:hover{color:var(--crit)}
  .snap-empty{
    padding:14px; text-align:center; color:var(--muted); font-size:12px;
  }

  ::-webkit-scrollbar{width:9px; height:9px}
  ::-webkit-scrollbar-track{background:transparent}
  ::-webkit-scrollbar-thumb{background:rgba(148,163,184,0.22); border-radius:6px}
  ::-webkit-scrollbar-thumb:hover{background:rgba(148,163,184,0.45)}

  @media (max-width:1100px){.app{grid-template-columns:240px 1fr 320px}}

  /* ============================================================
     PRINT STYLES
     Hides chrome, maximizes canvas, forces light surrounding background.
     Canvas itself still renders with the dark palette; this is intentional
     so the visual identity carries to PDF. For ink-saving prints, use
     PNG export and edit the image instead.
     ============================================================ */
  @media print {
    @page { size: landscape; margin: 8mm; }
    body{
      background:#FFFFFF !important;
      color:#000 !important;
      overflow:visible !important;
    }
    aside.sidebar, aside.details, footer.status-bar,
    .canvas-toolbar, .canvas-banner, .legend-toggle,
    .header-right, .file-chip,
    .toast, .ctx-menu, .tooltip, .drop-overlay, .modal-backdrop:not(.metrics-print){
      display:none !important;
    }
    /* When the metrics modal is open for printing, show its contents and
       hide the graph behind. Triggered by .metrics-print class on the
       backdrop (set in openMetricsModal). */
    .modal-backdrop.metrics-print{
      display:block !important; position:static !important;
      background:#FFFFFF !important; backdrop-filter:none !important;
    }
    .modal-backdrop.metrics-print .modal{
      width:100% !important; max-width:none !important;
      max-height:none !important; border:none !important;
      box-shadow:none !important; background:#FFFFFF !important;
      color:#000 !important; display:block !important;
    }
    .modal-backdrop.metrics-print .modal-head,
    .modal-backdrop.metrics-print .modal-foot{
      border-color:#CCC !important;
    }
    .modal-backdrop.metrics-print .modal-close,
    .modal-backdrop.metrics-print #btn-metrics-copy,
    .modal-backdrop.metrics-print #btn-metrics-print,
    .modal-backdrop.metrics-print #modal-cancel,
    .modal-backdrop.metrics-print .m-tabs{
      display:none !important;
    }
    /* Show every pane stacked when printing, with section titles */
    .modal-backdrop.metrics-print .m-pane{
      display:block !important;
      page-break-inside:avoid;
      border-top:1px solid #DDD; margin-top:14px; padding-top:10px;
    }
    .modal-backdrop.metrics-print .m-pane:first-of-type{
      border-top:none; margin-top:0; padding-top:0;
    }
    .modal-backdrop.metrics-print .m-pane::before{
      content:attr(data-pane-title);
      display:block; font-weight:700; font-size:14px;
      color:#000; margin-bottom:8px; text-transform:capitalize;
    }
    .modal-backdrop.metrics-print .kpi,
    .modal-backdrop.metrics-print .m-table{
      background:#FFF !important; color:#000 !important;
      border:1px solid #CCC !important;
    }
    .modal-backdrop.metrics-print .kpi .lbl,
    .modal-backdrop.metrics-print .m-table th{
      color:#666 !important;
    }
    .modal-backdrop.metrics-print .kpi .val,
    .modal-backdrop.metrics-print .m-table td,
    .modal-backdrop.metrics-print .m-table td:first-child,
    .modal-backdrop.metrics-print .m-bar-row .bar-label,
    .modal-backdrop.metrics-print .m-bar-row .bar-val,
    .modal-backdrop.metrics-print .m-section h4,
    .modal-backdrop.metrics-print .m-section p,
    .modal-backdrop.metrics-print .m-lede{
      color:#000 !important;
    }
    .app{
      grid-template-columns:1fr !important;
      grid-template-rows:auto 1fr !important;
      grid-template-areas:"header" "canvas" !important;
      height:auto !important;
      width:100% !important;
      overflow:visible !important;
    }
    header.app-header{
      padding:6px 12px !important;
      background:#fff !important;
      border-bottom:1px solid #ccc !important;
      backdrop-filter:none !important;
    }
    .brand-text h1{color:#0D0E11 !important; background:none !important}
    .brand-text p{color:#555 !important}
    .brand-mark{
      background:linear-gradient(135deg,#0EA5E9,#7C3AED) !important;
      box-shadow:none !important;
    }
    main.canvas-area{
      height:calc(100vh - 60px) !important;
      background:#0D0E11 !important; /* dark canvas surface stays */
      border-radius:6px;
    }
    /* Legend prints with a light card for context */
    .legend{
      background:#fff !important;
      color:#000 !important;
      border:1px solid #999 !important;
      box-shadow:none !important;
      position:absolute !important;
      bottom:10mm !important; left:10mm !important;
      max-width:60mm !important;
      page-break-inside:avoid;
    }
    .legend h4{color:#333 !important}
    .legend-item{color:#222 !important}
    .legend-tabs{display:none !important} /* show whichever tab is open */
  }
  /* Mid-width: header buttons multiplying make the right side cramped.
     Hide the filename label below 1100px (keep file-chip icon for save action). */
  @media (max-width:1100px){
    .file-chip .fc-name{display:none}
    .stats .stat .lbl{display:none}
    .status-bar .center{display:none}
  }
  /* Narrower still: hide modified chip and brand subtitle */
  @media (max-width:1000px){
    .modified-chip{display:none !important}
    .brand-text p{display:none}
  }
  /* Tablet/phone: stack panels vertically, hide rarely-used controls */
  @media (max-width:900px){
    .app{
      grid-template-columns:1fr;
      grid-template-rows:auto auto 1fr auto auto;
      grid-template-areas:"header" "sidebar" "canvas" "details" "footer";
    }
    aside.sidebar{
      max-height:42vh; border-right:none; border-bottom:1px solid var(--border);
    }
    aside.details{
      max-height:36vh; border-left:none; border-top:1px solid var(--border);
    }
    .stats .stat{min-width:auto; padding:4px 7px}
    .stats .stat .lbl{display:none}
    .brand-text p{display:none}
  }

  /* ---------- Leadership Metrics modal ---------- */
  .m-tabs{
    display:flex; gap:2px; border-bottom:1px solid var(--border);
    margin:-14px -16px 14px -16px; padding:0 12px; overflow-x:auto;
  }
  .m-tab{
    background:transparent; border:none; color:var(--muted);
    padding:10px 12px; cursor:pointer; font-size:12px;
    border-bottom:2px solid transparent; font-family:inherit; white-space:nowrap;
  }
  .m-tab:hover{color:var(--text-2)}
  .m-tab.active{color:var(--accent); border-bottom-color:var(--accent)}
  .m-pane{display:none}
  .m-pane.active{display:block}

  .kpi-grid{
    display:grid; grid-template-columns:repeat(auto-fit, minmax(140px, 1fr));
    gap:10px; margin-bottom:14px;
  }
  .kpi{
    background:var(--card); border:1px solid var(--border);
    border-radius:10px; padding:12px 14px;
  }
  .kpi .lbl{
    font-size:10.5px; color:var(--muted); text-transform:uppercase;
    letter-spacing:0.5px;
  }
  .kpi .val{
    font-size:22px; font-weight:700; color:var(--text);
    font-variant-numeric:tabular-nums; margin-top:2px; line-height:1.1;
  }
  .kpi .sub{font-size:11.5px; color:var(--text-2); margin-top:3px;}
  .kpi.accent .val{color:var(--accent)}
  .kpi.warn .val{color:var(--warn)}
  .kpi.crit .val{color:var(--crit)}
  .kpi.ok .val{color:var(--ok)}

  .m-section{margin-bottom:14px;}
  .m-section h4{
    margin:14px 0 8px; color:var(--accent-2);
    font-size:11.5px; text-transform:uppercase; letter-spacing:0.6px;
  }
  .m-section h4:first-child{margin-top:0}
  .m-section p.m-lede{
    color:var(--text-2); font-size:12.5px; margin:0 0 10px; line-height:1.5;
  }

  .m-table{width:100%; border-collapse:collapse; font-size:12.5px;}
  .m-table th{
    text-align:left; color:var(--muted); font-weight:500; font-size:10.5px;
    text-transform:uppercase; letter-spacing:0.5px;
    padding:6px 8px; border-bottom:1px solid var(--border);
  }
  .m-table td{
    padding:7px 8px; border-bottom:1px solid var(--border);
    color:var(--text-2); vertical-align:top;
  }
  .m-table td:first-child{color:var(--text)}
  .m-table tr:last-child td{border-bottom:none}
  .m-table .num{text-align:right; font-variant-numeric:tabular-nums; color:var(--text)}
  .m-table .type-cell{color:var(--muted); font-size:11px;}

  .m-bar-row{
    display:grid; grid-template-columns:140px 1fr 120px;
    gap:10px; align-items:center; margin-bottom:8px; font-size:12.5px;
  }
  .m-bar-row .bar-label{color:var(--text)}
  .m-bar-row .bar-val{
    color:var(--text); font-variant-numeric:tabular-nums;
    text-align:right; font-weight:600;
  }
  .m-bar{
    height:7px; background:var(--card-2); border-radius:4px; overflow:hidden;
  }
  .m-bar-fill{
    height:100%; border-radius:4px;
    background:linear-gradient(90deg, var(--accent), var(--violet));
  }
  .m-bar-fill.low{background:var(--crit)}
  .m-bar-fill.mid{background:var(--warn)}
  .m-bar-fill.high{background:var(--ok)}

  .m-empty{
    color:var(--muted); font-size:12.5px; padding:8px 0; font-style:italic;
  }
  .m-pill{
    display:inline-block; background:var(--card-2); border:1px solid var(--border);
    border-radius:10px; padding:1px 7px; font-size:10.5px;
    color:var(--text-2); margin-left:6px;
  }
  .m-pill.ok{color:var(--ok); border-color:rgba(52,211,153,0.35)}
  .m-pill.warn{color:var(--warn); border-color:rgba(251,191,36,0.35)}
  .m-pill.crit{color:var(--crit); border-color:rgba(248,113,113,0.35)}

  /* Org chart view toggle — visual emphasis when active */
  .btn.primary{
    background:var(--accent); color:#F5F3FF;
    border-color:var(--accent);
  }
  /* Dark theme only: #6366F1 gives 4.07:1 against #F5F3FF text — just under
     the WCAG 1.4.3 minimum of 4.5:1. #5156E5 reaches ~5:1. Light theme keeps
     its own (passing) accent. */
  :root:not([data-theme="light"]) .btn.primary{
    background:#5156E5; border-color:#5156E5;
  }
  .btn.primary:hover{
    background:var(--accent-2); border-color:var(--accent-2);
  }

  /* Recently Viewed popover (anchored under the header recent button) */
  #recent-popover{
    position:absolute; top:54px; right:130px; z-index:120;
    width:280px; background:var(--panel); border:1px solid var(--border);
    border-radius:10px; box-shadow:var(--shadow-2);
    padding:10px; display:none;
  }
  #recent-popover.open{display:block}
  .recent-head{
    font-size:11px; text-transform:uppercase; letter-spacing:0.6px;
    color:var(--accent-2); font-weight:600; margin-bottom:8px;
  }
  .recent-empty{
    color:var(--muted); font-size:12px; font-style:italic;
    padding:14px 6px; text-align:center;
  }
  .recent-list{
    max-height:340px; overflow-y:auto;
    display:flex; flex-direction:column; gap:2px;
  }
  .recent-item{
    display:flex; align-items:center; gap:8px;
    padding:6px 8px; border-radius:6px; cursor:pointer;
    transition:background 0.12s;
  }
  .recent-item:hover{
    background:var(--card);
  }
  .recent-item:focus-visible{
    background:var(--card); outline:2px solid var(--accent); outline-offset:-2px;
  }
  .recent-dot{
    width:8px; height:8px; border-radius:50%; flex-shrink:0;
  }
  .recent-meta{min-width:0; flex:1}
  .recent-name{
    font-size:13px; color:var(--text); font-weight:500;
    overflow:hidden; text-overflow:ellipsis; white-space:nowrap;
  }
  .recent-type{
    font-size:11px; color:var(--muted);
    overflow:hidden; text-overflow:ellipsis; white-space:nowrap;
  }

  /* Edge inline editing: smaller select + textarea inside the profile panel */
  .edge-inline-row{
    display:flex; flex-direction:column; gap:4px; margin-bottom:8px;
  }
  .edge-inline-row label{
    font-size:11px; color:var(--muted); text-transform:uppercase;
    letter-spacing:0.5px;
  }
  .edge-inline-row select, .edge-inline-row textarea{
    width:100%;
  }
  .edge-inline-row textarea{
    min-height:48px; font-size:12.5px;
  }
  .edge-inline-saved{
    color:var(--ok); font-size:11px; opacity:0;
    transition:opacity 0.2s;
  }
  .edge-inline-saved.show{opacity:1}

  /* Drill-down: clickable entity name in a metric row jumps to the node */
  .m-table .m-entity-link, .m-entity-link{
    color:var(--accent-2); cursor:pointer;
    border-bottom:1px dotted transparent;
    transition:border-color 0.15s, color 0.15s;
  }
  .m-table .m-entity-link:hover, .m-entity-link:hover{
    color:var(--accent); border-bottom-color:var(--accent);
  }
  .m-table .m-entity-link:focus-visible, .m-entity-link:focus-visible{
    outline:2px solid var(--accent); outline-offset:1px; border-radius:2px;
  }

  /* ---------- Extra-large modal (used by metrics) ---------- */
  .modal.xl{
    width:1100px; max-width:94vw; max-height:90vh;
  }
  .modal.xl .modal-body{
    padding:20px 24px;
  }
  .modal.xl .m-tabs{
    margin:-20px -24px 18px -24px; padding:0 16px;
  }
  .modal.xl .m-tab{
    padding:12px 14px; font-size:12.5px;
  }

  /* ---------- Tailwind-ish polish on existing primitives ---------- */
  .modal.xl .kpi{
    background:var(--card);
    border:1px solid rgba(148,163,184,0.14);
    border-radius:12px;
    padding:14px 16px;
    transition:border-color 0.15s, background 0.15s;
  }
  .modal.xl .kpi:hover{
    border-color:rgba(148,163,184,0.28);
    background:var(--card-2);
  }
  .modal.xl .kpi .lbl{
    font-size:10.5px; letter-spacing:0.6px;
  }
  .modal.xl .kpi .val{
    font-size:26px; font-weight:700;
    margin-top:4px; line-height:1.05;
  }
  .modal.xl .kpi .sub{
    font-size:11.5px; margin-top:4px; color:var(--muted);
  }
  .modal.xl .kpi-grid{
    gap:12px; margin-bottom:18px;
    grid-template-columns:repeat(auto-fit, minmax(160px, 1fr));
  }
  .modal.xl .m-section{margin-bottom:22px;}
  .modal.xl .m-section h4{
    margin:18px 0 10px;
    font-size:11px; letter-spacing:0.7px;
    color:var(--text-2);
    font-weight:600;
  }
  .modal.xl .m-section h4:first-child{margin-top:0}
  .modal.xl .m-section p.m-lede{
    font-size:13px; color:var(--text-2); margin:0 0 14px;
  }
  .modal.xl .m-table{
    background:var(--card);
    border:1px solid rgba(148,163,184,0.14);
    border-radius:10px;
    overflow:hidden;
  }
  .modal.xl .m-table th{
    background:rgba(255,255,255,0.02);
    padding:9px 12px;
    font-size:10.5px;
  }
  .modal.xl .m-table td{
    padding:10px 12px;
    border-bottom-color:rgba(148,163,184,0.10);
  }
  .modal.xl .m-table tr:hover td{
    background:rgba(99,102,241,0.04);
  }

  /* ---------- Donut (radial gauge) ---------- */
  .m-donut-grid{
    display:grid; grid-template-columns:repeat(auto-fit, minmax(180px, 1fr));
    gap:14px; margin-bottom:18px;
  }
  .m-donut{
    background:var(--card);
    border:1px solid rgba(148,163,184,0.14);
    border-radius:12px;
    padding:18px 16px 16px;
    text-align:center;
    transition:border-color 0.15s, background 0.15s;
  }
  .m-donut:hover{
    border-color:rgba(148,163,184,0.28);
    background:var(--card-2);
  }
  .m-donut-svg-wrap{
    position:relative; width:120px; height:120px; margin:0 auto;
  }
  .m-donut-svg{width:100%; height:100%; display:block}
  .m-donut-center{
    position:absolute; inset:0;
    display:flex; flex-direction:column; align-items:center; justify-content:center;
    pointer-events:none;
  }
  .m-donut-val{
    font-size:24px; font-weight:700; color:var(--text);
    font-variant-numeric:tabular-nums; line-height:1;
  }
  .m-donut-pct{
    font-size:13px; color:var(--muted); margin-left:1px; font-weight:500;
  }
  .m-donut-center-lbl{
    font-size:10px; color:var(--muted); text-transform:uppercase;
    letter-spacing:0.5px; margin-top:3px;
  }
  .m-donut-lbl{
    margin-top:12px; font-size:13px; color:var(--text); font-weight:500;
  }
  .m-donut-sub{
    margin-top:3px; font-size:11.5px; color:var(--muted);
  }

  /* ---------- Segmented donut (distributions) ---------- */
  .m-seg-donut-wrap{
    display:grid; grid-template-columns:200px 1fr;
    gap:24px; align-items:center;
    background:var(--card);
    border:1px solid rgba(148,163,184,0.14);
    border-radius:12px;
    padding:18px 20px;
  }
  .m-seg-donut-svg-wrap{
    position:relative; width:180px; height:180px;
  }
  .m-seg-donut-svg{width:100%; height:100%; display:block}
  .m-seg-donut .m-donut-center .m-donut-val{font-size:30px}
  .m-seg-legend{
    display:grid; grid-template-columns:repeat(auto-fit, minmax(160px, 1fr));
    gap:8px 16px;
  }
  .m-seg-legend-row{
    display:flex; align-items:center; gap:8px;
    font-size:12.5px; color:var(--text-2);
  }
  .m-seg-dot{
    width:10px; height:10px; border-radius:3px; flex-shrink:0;
  }
  .m-seg-lbl{flex:1; min-width:0; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}
  .m-seg-val{
    color:var(--text); font-variant-numeric:tabular-nums; font-weight:600;
  }
  .m-seg-pct{
    color:var(--muted); margin-left:4px; font-weight:400; font-size:11.5px;
  }

  /* ---------- Horizontal bar charts ---------- */
  .m-hbars{
    background:var(--card);
    border:1px solid rgba(148,163,184,0.14);
    border-radius:12px;
    padding:14px 16px;
    display:flex; flex-direction:column; gap:8px;
  }
  .m-hbar-row{
    display:grid; grid-template-columns:minmax(140px, 230px) 1fr 50px;
    gap:12px; align-items:center;
    font-size:12.5px;
  }
  .m-hbar-label{
    color:var(--text); min-width:0;
    overflow:hidden; text-overflow:ellipsis; white-space:nowrap;
  }
  .m-hbar-sub{
    color:var(--muted); margin-left:6px; font-size:11px;
  }
  .m-hbar-track{
    height:10px; background:var(--card-2);
    border-radius:5px; overflow:hidden;
  }
  .m-hbar-fill{
    height:100%; border-radius:5px;
    transition:width 0.4s ease;
  }
  .m-hbar-val{
    text-align:right; color:var(--text); font-weight:600;
    font-variant-numeric:tabular-nums;
  }

  /* ---------- Stacked horizontal bars ---------- */
  .m-stacked-hbars{
    background:var(--card);
    border:1px solid rgba(148,163,184,0.14);
    border-radius:12px;
    padding:14px 16px;
    display:flex; flex-direction:column; gap:10px;
  }
  .m-stacked-row{
    display:grid; grid-template-columns:minmax(160px, 240px) 1fr 50px;
    gap:12px; align-items:center;
    font-size:12.5px;
  }
  .m-stacked-label{color:var(--text); overflow:hidden; text-overflow:ellipsis; white-space:nowrap;}
  .m-stacked-track-wrap{position:relative; height:12px;}
  .m-stacked-track{
    height:100%; display:flex;
    border-radius:6px; overflow:hidden;
    background:var(--card-2);
    transition:width 0.4s ease;
  }
  .m-stack-seg{
    height:100%;
    transition:width 0.4s ease;
  }
  .m-stack-seg:first-child{border-radius:6px 0 0 6px}
  .m-stack-seg:last-child{border-radius:0 6px 6px 0}
  .m-stack-seg:only-child{border-radius:6px}
  .m-stacked-total{
    text-align:right; color:var(--text); font-weight:600;
    font-variant-numeric:tabular-nums;
  }
  .m-chart-legend{
    display:flex; flex-wrap:wrap; gap:14px;
    margin-top:10px; padding-top:10px;
    border-top:1px solid rgba(148,163,184,0.10);
    font-size:11.5px; color:var(--text-2);
  }
  .m-chart-legend-item{display:flex; align-items:center; gap:6px}

  /* ---------- Print rules for charts ---------- */
  @media print {
    .modal-backdrop.metrics-print .m-donut,
    .modal-backdrop.metrics-print .m-seg-donut-wrap,
    .modal-backdrop.metrics-print .m-hbars,
    .modal-backdrop.metrics-print .m-stacked-hbars{
      background:#FFF !important;
      border:1px solid #CCC !important;
    }
    .modal-backdrop.metrics-print .m-donut-val,
    .modal-backdrop.metrics-print .m-donut-lbl,
    .modal-backdrop.metrics-print .m-seg-lbl,
    .modal-backdrop.metrics-print .m-seg-val,
    .modal-backdrop.metrics-print .m-hbar-label,
    .modal-backdrop.metrics-print .m-hbar-val,
    .modal-backdrop.metrics-print .m-stacked-label,
    .modal-backdrop.metrics-print .m-stacked-total{
      color:#000 !important;
    }
    .modal-backdrop.metrics-print .m-hbar-track,
    .modal-backdrop.metrics-print .m-stacked-track{
      background:#EEE !important;
    }
  }
  /* ===== v1.1 feature styles ===== */
  #present-overlay{position:fixed; inset:0; background:transparent; z-index:60; display:none; align-items:flex-end; justify-content:center;}
  #present-overlay.active{display:flex;}
  #present-stage{width:100%; max-width:900px; margin:0 auto 5vh; background:var(--card); border:1px solid var(--border); border-radius:14px; padding:20px 24px; box-shadow:0 20px 60px rgba(0,0,0,0.5);}
  #present-note .pn-title{font-size:20px; font-weight:700; color:var(--text); margin-bottom:6px;}
  #present-note .pn-body{font-size:15px; line-height:1.55; color:var(--text-2); white-space:pre-wrap;}
  #present-note .pn-empty{font-size:14px; color:var(--muted); font-style:italic;}
  #present-controls{display:flex; align-items:center; gap:10px; margin-top:16px;}
  #present-controls #present-progress{font-size:12px; color:var(--muted); margin:0 auto;}
  .be-row{display:flex; align-items:flex-start; gap:8px; margin-bottom:9px;}
  .be-row input[type=checkbox]{flex:none; margin-top:7px;}
  .be-row .be-field{flex:1; display:flex; flex-direction:column; gap:2px;}
  .be-row label.be-name{font-size:11px; color:var(--muted);}
  .step-row{display:flex; align-items:center; gap:6px; margin-bottom:6px;}
  .step-row .step-note{flex:1;}
  .step-row .step-num{width:22px; text-align:center; color:var(--muted); font-size:12px; flex:none;}
  .step-row select.step-node{flex:1; min-width:0;}
  .heat-note{font-size:11px; color:var(--muted); margin:2px 0 6px;}
.rel-warn{ margin-top:10px; padding:8px 10px; border-radius:8px; background:rgba(245,158,11,0.12); border:1px solid rgba(245,158,11,0.5); color:#FCD34D; font-size:12px; line-height:1.45; }
.rel-warn .rel-suggest{ color:#FDE68A; opacity:0.92; }
.rel-force{ display:flex; align-items:center; gap:6px; margin-top:8px; font-size:12px; color:var(--muted); cursor:pointer; }
.rel-force input{ cursor:pointer; }
.sr-only{ position:absolute; width:1px; height:1px; padding:0; margin:-1px; overflow:hidden; clip:rect(0 0 0 0); white-space:nowrap; border:0; }
/* Metric trend chips (change vs. baseline snapshot) */
.kpi .delta{ font-size:11px; font-weight:600; vertical-align:middle; margin-left:4px; padding:1px 5px; border-radius:5px; white-space:nowrap; }
.delta.good{ color:var(--ok); background:rgba(52,211,153,0.12); }
.delta.bad{ color:var(--crit); background:rgba(248,113,113,0.12); }
.delta.flat{ color:var(--muted); background:rgba(148,163,184,0.12); }
.m-baseline-row{ display:flex; align-items:center; gap:8px; margin:0 0 12px; font-size:11.5px; color:var(--text-2); flex-wrap:wrap; }
.m-baseline-row select{ font-size:11.5px; padding:3px 6px; max-width:280px; }
.m-baseline-hint{ color:var(--muted); font-size:11px; }
@media (pointer: coarse){
  .btn, select, .legend-tab{ min-height:44px; }
}
#small-viewport-notice{ display:none; }
/* Only block on genuinely small TOUCH devices. Width alone also matches a
   desktop browser zoomed to 400% (320px CSS viewport), which WCAG 1.4.10
   requires to stay usable — a low-vision user must not lose the whole app. */
@media (max-width:720px) and (pointer:coarse){
  #small-viewport-notice{ display:flex; position:fixed; inset:0; z-index:9999; align-items:center; justify-content:center; background:var(--bg-0); padding:24px; box-sizing:border-box; text-align:center; }
  #small-viewport-notice strong{ color:var(--text); font-size:16px; }
  #small-viewport-notice p{ color:var(--muted); font-size:13px; margin-top:8px; max-width:40ch; line-height:1.5; }
}
/* WCAG 2.3.3 / 2.2.2: respect the OS-level reduced-motion preference */
@media (prefers-reduced-motion: reduce){
  *, *::before, *::after{
    transition-duration:0.01ms !important;
    animation-duration:0.01ms !important;
    animation-iteration-count:1 !important;
  }
}`;

export const ENGINE_HTML: string = `<div id="a11y-live" class="sr-only" aria-live="polite" aria-atomic="true"></div>
<nav id="a11y-graph" class="sr-only" aria-label="Graph structure, text alternative"></nav>
<div id="small-viewport-notice" role="note"><div><strong>Best viewed on desktop</strong><p>This relationship graph is built for a larger screen. Open it on a desktop or laptop for full pan, zoom, and editing.</p></div></div>
<div class="app" id="app">
  <header class="app-header">
    <div class="brand">
      <div class="brand-mark"><i class="fa-solid fa-diagram-project"></i></div>
      <div class="brand-text">
        <h1>Enterprise Relationship Graph</h1>
        <p>Discover offices, systems, platforms, sites, teams, groups, and operational dependencies.</p>
      </div>
    </div>
    <div class="header-right">
      <button class="file-chip" id="file-chip" title="Click to save (or open if no file is loaded)" aria-label="Current file">
        <i class="fa-solid fa-file-code fc-icon" aria-hidden="true"></i>
        <span class="fc-name no-file" id="fc-name">no file open</span>
        <span class="fc-dirty" title="Unsaved changes"></span>
      </button>
      <button class="modified-chip" id="modified-chip" title="Click to update your name" aria-label="Last save info" style="display:none">
        <i class="fa-solid fa-user-clock mc-icon" aria-hidden="true"></i>
        <span class="mc-text"></span>
      </button>
      <div class="stats" id="stats" aria-label="Graph statistics">
        <div class="stat"><span class="val" id="stat-nodes">0</span><span class="lbl">nodes</span></div>
        <div class="stat"><span class="val" id="stat-edges">0</span><span class="lbl">relations</span></div>
        <div class="stat"><span class="val" id="stat-vnodes">0</span><span class="lbl">visible</span></div>
        <div class="stat accent"><span class="val" id="stat-sel">·</span><span class="lbl">selected</span></div>
      </div>
      <button class="header-btn" id="btn-recent" title="Recently viewed entities" aria-label="Recently viewed entities" aria-haspopup="dialog" aria-expanded="false"><i class="fa-solid fa-clock-rotate-left" aria-hidden="true"></i></button>
      <button class="header-btn" id="btn-data-quality" title="Data quality checks" aria-label="Data quality checks"><i class="fa-solid fa-clipboard-check" aria-hidden="true"></i></button>
      <button class="header-btn" id="btn-metrics" title="Leadership metrics" aria-label="Leadership metrics"><i class="fa-solid fa-chart-line" aria-hidden="true"></i></button>
      <button class="header-btn" id="btn-value-brief" title="Value brief: why this tool" aria-label="Value brief: why this tool"><i class="fa-solid fa-award"></i></button>
      <button class="header-btn" id="btn-theme" title="Switch to light mode" aria-label="Switch to light mode"><i class="fa-solid fa-sun"></i></button>
      <button class="header-btn" id="btn-help" title="Help (?)" aria-label="Help"><i class="fa-solid fa-circle-question" aria-hidden="true"></i></button>
    </div>
  </header>
  <div id="recent-popover" role="dialog" aria-label="Recently viewed entities"></div>

  <aside class="sidebar" aria-label="Graph controls and data tools">
    <div class="section" data-section="search">
      <button class="section-head" type="button" aria-expanded="true"><span><i class="fa-solid fa-magnifying-glass icon" aria-hidden="true"></i>&nbsp; Search</span><i class="fa-solid fa-chevron-down chev" aria-hidden="true"></i></button>
      <div class="section-body">
        <div class="field with-icon">
          <i class="fa-solid fa-magnifying-glass field-icon"></i>
          <input type="text" id="search-input" aria-label="Search by name, type, or tag. Supports field operators like type:Person or owner:name" placeholder="Search… (e.g. type:Person owner:smith)" title="Free text matches name/type/tags. Operators: type: tag: owner: lead: status: platform: office: directorate: id: — quote multi-word values" />
        </div>
        <div class="btn-row" style="margin-top:6px">
          <button class="btn ghost" id="btn-clear-search"><i class="fa-solid fa-xmark"></i> Clear</button>
          <button class="btn" id="btn-focus-search" title="Jump to and select the first matching node"><i class="fa-solid fa-crosshairs"></i> Jump to</button>
        </div>
      </div>
    </div>

    <div class="section" data-section="view-controls">
      <button class="section-head" type="button" aria-expanded="true"><span><i class="fa-solid fa-sliders icon" aria-hidden="true"></i>&nbsp; View Controls</span><i class="fa-solid fa-chevron-down chev" aria-hidden="true"></i></button>
      <div class="section-body">
        <div class="field">
          <label for="layout-select">Layout</label>
          <select id="layout-select">
            <option value="force">Force-directed</option>
            <option value="hierarchical">Hierarchical (top-down)</option>
            <option value="hierarchical-lr">Hierarchical (left-right)</option>
            <option value="org-chart">Org Chart (people + reports)</option>
            <option value="radial">Radial tree</option>
            <option value="cluster">Group by attribute</option>
            <option value="concentric">Concentric (by type)</option>
            <option value="grid">Grid</option>
          </select>
        </div>
        <div class="field" id="cluster-attr-field" style="display:none">
          <label for="cluster-attr">Cluster by</label>
          <select id="cluster-attr">
            <option value="type">Type</option>
            <option value="directorate">Directorate</option>
            <option value="office">Office</option>
            <option value="parent">Parent</option>
            <option value="status">Status</option>
            <option value="platform">Platform</option>
          </select>
        </div>
        <div class="field">
          <label for="scope-select">Scope to branch</label>
          <select id="scope-select"><option value="">Whole graph</option></select>
        </div>
        <div class="btn-row">
          <button class="btn" id="btn-fit"><i class="fa-solid fa-expand"></i> Fit</button>
          <button class="btn" id="btn-reset-view"><i class="fa-solid fa-rotate-left"></i> Reset</button>
        </div>
        <div class="btn-row" style="margin-top:6px">
          <button class="btn" id="btn-neighborhood"><i class="fa-solid fa-circle-nodes"></i> Neighborhood</button>
          <button class="btn" id="btn-clear-sel"><i class="fa-solid fa-ban"></i> Clear</button>
        </div>
        <button class="btn block" id="btn-find-path" style="margin-top:6px"><i class="fa-solid fa-route"></i> Find path between two nodes</button>
        <div class="btn-row" style="margin-top:6px">
          <button class="btn" id="btn-collapse-all" title="Collapse every parent node into a clean overview"><i class="fa-solid fa-square-minus"></i> Collapse</button>
          <button class="btn" id="btn-expand-all" title="Expand every collapsed branch"><i class="fa-solid fa-square-plus"></i> Expand</button>
        </div>
      </div>
    </div>

    <div class="section" data-section="node-types">
      <button class="section-head" type="button" aria-expanded="true"><span><i class="fa-solid fa-circle icon" aria-hidden="true"></i>&nbsp; Node Types</span><i class="fa-solid fa-chevron-down chev" aria-hidden="true"></i></button>
      <div class="section-body">
        <div class="filter-controls">
          <button class="btn" id="ntf-all">All</button>
          <button class="btn" id="ntf-none">None</button>
        </div>
        <div class="scroll filter-list" id="node-type-filters"></div>
      </div>
    </div>

    <div class="section" data-section="relationship-types">
      <button class="section-head" type="button" aria-expanded="true"><span><i class="fa-solid fa-arrow-right-long icon" aria-hidden="true"></i>&nbsp; Relationship Types</span><i class="fa-solid fa-chevron-down chev" aria-hidden="true"></i></button>
      <div class="section-body">
        <div class="filter-controls">
          <button class="btn" id="etf-all">All</button>
          <button class="btn" id="etf-none">None</button>
        </div>
        <div class="scroll filter-list" id="edge-type-filters"></div>
      </div>
    </div>

    <div class="section" data-section="edit-graph">
      <button class="section-head" type="button" aria-expanded="true"><span><i class="fa-solid fa-plus icon" aria-hidden="true"></i>&nbsp; Edit Graph</span><i class="fa-solid fa-chevron-down chev" aria-hidden="true"></i></button>
      <div class="section-body">
        <div class="btn-row">
          <button class="btn primary" id="btn-add-node"><i class="fa-solid fa-plus"></i> Node</button>
          <button class="btn primary" id="btn-add-edge"><i class="fa-solid fa-link"></i> Relation</button>
        </div>
        <div class="btn-row" style="margin-top:6px">
          <button class="btn" id="btn-connect-mode"><i class="fa-solid fa-arrows-turn-to-dots"></i> Quick connect</button>
          <button class="btn" id="btn-edit-sel"><i class="fa-solid fa-pen"></i> Edit</button>
        </div>
        <div class="btn-row" style="margin-top:6px">
          <button class="btn" id="btn-undo"><i class="fa-solid fa-rotate-left"></i> Undo</button>
          <button class="btn" id="btn-redo"><i class="fa-solid fa-rotate-right"></i> Redo</button>
        </div>
        <button class="btn danger block" id="btn-delete-sel" style="margin-top:6px"><i class="fa-solid fa-trash"></i> Delete selected</button>
        <button class="btn block" id="btn-bulk-edit" style="margin-top:6px" disabled><i class="fa-solid fa-layer-group"></i> Bulk edit selection</button>
        <button class="btn block" id="btn-manage-types" style="margin-top:6px"><i class="fa-solid fa-palette"></i> Manage node types</button>
      </div>
    </div>

    <div class="section" data-section="analysis">
      <button class="section-head" type="button" aria-expanded="true"><span><i class="fa-solid fa-flask icon" aria-hidden="true"></i>&nbsp; Analysis</span><i class="fa-solid fa-chevron-down chev" aria-hidden="true"></i></button>
      <div class="section-body">
        <button class="btn block" id="btn-whatif"><i class="fa-solid fa-flask-vial"></i> What-if scenario mode</button>
        <button class="btn block" id="btn-spof" style="margin-top:6px"><i class="fa-solid fa-triangle-exclamation"></i> Find single points of failure</button>
        <button class="btn block" id="btn-heatmap" style="margin-top:6px"><i class="fa-solid fa-temperature-half"></i> Staleness heatmap</button>
      </div>
    </div>

    <div class="section" data-section="views">
      <button class="section-head" type="button" aria-expanded="true"><span><i class="fa-solid fa-bookmark icon" aria-hidden="true"></i>&nbsp; Saved Views</span><i class="fa-solid fa-chevron-down chev" aria-hidden="true"></i></button>
      <div class="section-body">
        <button class="btn primary block" id="btn-view-save" style="margin-bottom:6px"><i class="fa-solid fa-bookmark"></i> Save current view</button>
        <div id="view-list-container"></div>
      </div>
    </div>

    <div class="section" data-section="presentation">
      <button class="section-head" type="button" aria-expanded="true"><span><i class="fa-solid fa-chalkboard-user icon" aria-hidden="true"></i>&nbsp; Presentation</span><i class="fa-solid fa-chevron-down chev" aria-hidden="true"></i></button>
      <div class="section-body">
        <button class="btn primary block" id="btn-present-manage" style="margin-bottom:6px"><i class="fa-solid fa-chalkboard-user"></i> Build / play walkthrough</button>
        <div id="walkthrough-list-container"></div>
      </div>
    </div>

    <div class="section" data-section="snapshots">
      <button class="section-head" type="button" aria-expanded="true"><span><i class="fa-solid fa-camera icon" aria-hidden="true"></i>&nbsp; Snapshots</span><i class="fa-solid fa-chevron-down chev" aria-hidden="true"></i></button>
      <div class="section-body">
        <button class="btn primary block" id="btn-snap-save" style="margin-bottom:6px"><i class="fa-solid fa-floppy-disk"></i> Save current as snapshot</button>
        <div id="snap-list-container"></div>
      </div>
    </div>

    <div class="section" data-section="data">
      <button class="section-head" type="button" aria-expanded="true"><span><i class="fa-solid fa-database icon" aria-hidden="true"></i>&nbsp; Data</span><i class="fa-solid fa-chevron-down chev" aria-hidden="true"></i></button>
      <div class="section-body">
        <div style="font-size:10.5px; color:var(--muted); margin-bottom:4px; letter-spacing:0.3px">Working file</div>
        <div class="btn-row">
          <button class="btn primary" id="btn-file-open"><i class="fa-solid fa-folder-open" aria-hidden="true"></i> Open</button>
          <button class="btn primary" id="btn-file-save"><i class="fa-solid fa-floppy-disk" aria-hidden="true"></i> Save</button>
        </div>
        <button class="btn block" id="btn-file-save-as" style="margin-top:6px"><i class="fa-solid fa-file-arrow-down" aria-hidden="true"></i> Save As...</button>
        <button class="btn block" id="btn-reopen-last" style="display:none; margin-top:6px"><i class="fa-solid fa-clock-rotate-left" aria-hidden="true"></i> Reopen <span id="reopen-last-name"></span></button>
        <button class="btn block" id="btn-import-mermaid" style="margin-top:6px"><i class="fa-solid fa-file-import" aria-hidden="true"></i> Import Mermaid&hellip;</button>
        <div style="font-size:10.5px; color:var(--muted); margin:10px 0 4px; letter-spacing:0.3px">Export a copy as</div>
        <div class="btn-row" style="grid-template-columns:repeat(4,1fr); gap:4px">
          <button class="btn" id="btn-export-csv" title="CSV (nodes.csv + edges.csv)" aria-label="Export CSV"><i class="fa-solid fa-file-csv" aria-hidden="true"></i></button>
          <button class="btn" id="btn-export-xlsx" title="Excel workbook (.xlsx)" aria-label="Export Excel"><i class="fa-solid fa-file-excel" aria-hidden="true"></i></button>
          <button class="btn" id="btn-export-png" title="PNG image of the visible graph" aria-label="Export PNG"><i class="fa-solid fa-image" aria-hidden="true"></i></button>
          <button class="btn" id="btn-export-mermaid" title="Mermaid diagram (paste into any markdown editor that supports Mermaid)" aria-label="Export Mermaid"><i class="fa-solid fa-diagram-project" aria-hidden="true"></i></button>
        </div>
        <button class="btn warn block" id="btn-sample" style="margin-top:8px"><i class="fa-solid fa-flask" aria-hidden="true"></i> Load sample data</button>
        <input type="file" id="file-input" accept="application/json" style="display:none" />
      </div>
    </div>
  </aside>

  <main class="canvas-area">
    <canvas id="graph-canvas" tabindex="0" role="application" aria-label="Relationship graph canvas. Use arrow keys to navigate nodes spatially, left and right bracket keys to cycle through nodes, Enter to view details. Tab moves to the next control."></canvas>
    <div class="canvas-toolbar">
      <button class="tool-btn" id="tb-zoom-in" title="Zoom in" aria-label="Zoom in"><i class="fa-solid fa-plus" aria-hidden="true"></i></button>
      <button class="tool-btn" id="tb-zoom-out" title="Zoom out" aria-label="Zoom out"><i class="fa-solid fa-minus" aria-hidden="true"></i></button>
      <button class="tool-btn" id="tb-fit" title="Fit" aria-label="Fit graph to view"><i class="fa-solid fa-expand" aria-hidden="true"></i></button>
      <button class="tool-btn" id="tb-reset" title="Reset view" aria-label="Reset view"><i class="fa-solid fa-rotate-left" aria-hidden="true"></i></button>
      <div class="tool-divider"></div>
      <button class="tool-btn" id="tb-undo" title="Undo (Ctrl+Z)" aria-label="Undo"><i class="fa-solid fa-rotate-left" style="font-size:11px" aria-hidden="true"></i><i class="fa-solid fa-arrow-left" style="font-size:8px; margin-left:-2px" aria-hidden="true"></i></button>
      <button class="tool-btn" id="tb-redo" title="Redo (Ctrl+Y)" aria-label="Redo"><i class="fa-solid fa-arrow-right" style="font-size:8px; margin-right:-2px" aria-hidden="true"></i><i class="fa-solid fa-rotate-right" style="font-size:11px" aria-hidden="true"></i></button>
      <div class="tool-divider"></div>
      <button class="tool-btn" id="tb-edge-labels" title="Toggle edge labels (L)" aria-label="Toggle edge labels"><i class="fa-solid fa-tag" aria-hidden="true"></i></button>
      <button class="tool-btn" id="tb-layout-mode" title="Layout mode: drag nodes without selecting (M)" aria-label="Toggle layout mode"><i class="fa-solid fa-up-down-left-right" aria-hidden="true"></i></button>
      <button class="tool-btn" id="tb-unpin" title="Release all pinned positions" aria-label="Release all pinned nodes"><i class="fa-solid fa-thumbtack" aria-hidden="true" style="transform:rotate(45deg)"></i></button>
      <div class="tool-divider"></div>
      <button class="tool-btn" id="tb-heatmap" title="Staleness heatmap (H): tint nodes by time since last edit" aria-label="Toggle staleness heatmap"><i class="fa-solid fa-temperature-half" aria-hidden="true"></i></button>
      <button class="tool-btn" id="tb-dim-edges" title="Dim relationships: shade edges darker so the base view is less busy" aria-label="Toggle dimmed relationships"><i class="fa-solid fa-circle-half-stroke" aria-hidden="true"></i></button>
    </div>
    <div class="canvas-banner" id="connect-banner" style="display:none">
      <i class="fa-solid fa-arrows-turn-to-dots" aria-hidden="true"></i>
      <span id="connect-banner-text">Quick-connect mode: click a target node</span>
      <button class="ban-close" id="connect-banner-close"><i class="fa-solid fa-xmark" aria-hidden="true"></i> Cancel (Esc)</button>
    </div>
    <div class="canvas-banner" id="org-chart-banner" style="display:none">
      <i class="fa-solid fa-people-arrows" aria-hidden="true"></i>
      <span>Org Chart view &mdash; showing People + Roles + org units only</span>
      <button class="ban-close" id="org-chart-banner-close"><i class="fa-solid fa-xmark" aria-hidden="true"></i> Exit (back to previous layout)</button>
    </div>
    <div class="canvas-banner" id="path-banner" style="display:none">
      <i class="fa-solid fa-route" aria-hidden="true"></i>
      <span id="path-banner-text">Path-finder: click source node, then click target</span>
      <button class="ban-close" id="path-banner-close"><i class="fa-solid fa-xmark" aria-hidden="true"></i> Cancel (Esc)</button>
    </div>
    <div class="canvas-banner" id="whatif-banner" style="display:none">
      <i class="fa-solid fa-flask-vial" aria-hidden="true"></i>
      <span id="whatif-banner-text">What-if mode: click nodes to disable them and see the impact cascade</span>
      <button class="ban-close" id="whatif-banner-close"><i class="fa-solid fa-xmark" aria-hidden="true"></i> Exit (Esc)</button>
    </div>
    <div class="canvas-banner" id="spof-banner" style="display:none">
      <i class="fa-solid fa-triangle-exclamation" aria-hidden="true"></i>
      <span id="spof-banner-text">Single points of failure highlighted</span>
      <button class="ban-close" id="spof-banner-close"><i class="fa-solid fa-xmark" aria-hidden="true"></i> Clear (Esc)</button>
    </div>
    <div class="canvas-banner" id="scope-banner" style="display:none">
      <i class="fa-solid fa-filter" aria-hidden="true"></i>
      <span id="scope-banner-text">Scoped to a branch</span>
      <button class="ban-close" id="scope-banner-close"><i class="fa-solid fa-xmark" aria-hidden="true"></i> Show whole graph</button>
    </div>
    <div class="legend" id="legend">
      <button class="legend-toggle" id="legend-toggle" aria-label="Hide legend" aria-expanded="true">hide</button>
      <h4>Legend</h4>
      <div class="legend-tabs">
        <button class="legend-tab active" data-tab="nodes">Node Types</button>
        <button class="legend-tab" data-tab="edges">Edge Styles</button>
      </div>
      <div class="legend-grid" id="legend-grid-nodes"></div>
      <div class="legend-grid" id="legend-grid-edges" style="display:none"></div>
    </div>
    <div class="tooltip" id="tooltip" role="tooltip" aria-hidden="true"></div>
  </main>

  <aside class="details" id="details-panel" aria-label="Selection details"></aside>

  <footer class="status-bar">
    <div class="left">
      <span><span class="status-dot"></span>Offline-ready · localStorage</span>
      <span id="footer-layout">Layout: force</span>
      <span id="footer-mode"></span>
    </div>
    <div class="center">
      <span>Created by <strong>Ross Bethune</strong></span>
      <span aria-hidden="true">·</span>
      <a href="https://github.com/RJBethune/enterprise-relationship-graph" target="_blank" rel="noopener" title="View source on GitHub">
        <i class="fa-brands fa-github" aria-hidden="true"></i>
        <span>View on GitHub</span>
      </a>
    </div>
    <div class="right">
      <span id="footer-hover">Hover a node to inspect</span>
      <span id="footer-zoom">Zoom: 100%</span>
    </div>
  </footer>
</div>

<div id="present-overlay" aria-hidden="true">
  <div id="present-stage">
    <div id="present-note"></div>
    <div id="present-controls">
      <button class="btn" id="present-prev"><i class="fa-solid fa-arrow-left"></i> Prev</button>
      <span id="present-progress">1 / 1</span>
      <button class="btn" id="present-next">Next <i class="fa-solid fa-arrow-right"></i></button>
      <button class="btn" id="present-play"><i class="fa-solid fa-play"></i> Auto</button>
      <button class="btn" id="present-isolate" title="Show only the nodes in this walkthrough"><i class="fa-solid fa-eye"></i> Isolate</button>
      <button class="btn ghost" id="present-exit"><i class="fa-solid fa-xmark"></i> Exit</button>
    </div>
  </div>
</div>
<div class="modal-backdrop" id="modal-backdrop">
  <div class="modal" id="modal" role="dialog" aria-modal="true" aria-labelledby="modal-title">
    <div class="modal-head">
      <h3 id="modal-title">Add node</h3>
      <button class="modal-close" id="modal-close" aria-label="Close dialog"><i class="fa-solid fa-xmark" aria-hidden="true"></i></button>
    </div>
    <div class="modal-body" id="modal-body"></div>
    <div class="modal-foot" id="modal-foot">
      <button class="btn ghost" id="modal-cancel">Cancel</button>
      <button class="btn primary" id="modal-save"><i class="fa-solid fa-check" aria-hidden="true"></i> Save</button>
    </div>
  </div>
</div>

<div class="ctx-menu" id="ctx-menu"></div>
<div class="toast" id="toast" role="status" aria-live="polite" aria-atomic="true"></div>
<div class="drop-overlay" id="drop-overlay">
  <div class="drop-card">
    <i class="fa-solid fa-file-arrow-down" aria-hidden="true"></i>
    <h3>Drop JSON to open</h3>
    <p>Release to load the graph from this file</p>
  </div>
</div>`;
