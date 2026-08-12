import { SPComponentLoader } from '@microsoft/sp-loader';
import { ENGINE_CSS, ENGINE_HTML } from './engineAssets';
import { startEngine } from './engine';
import { IErgHost } from './hostContract';

/**
 * Fonts ship INSIDE the web part.
 *
 * Node icons are Font Awesome glyphs drawn onto the canvas and labels are Inter, so
 * these are data, not decoration — a missing font is a graph of blank squares.
 *
 * Importing them makes webpack emit the woff2 files alongside the bundle with
 * content-hashed names, and SPFx points `__webpack_public_path__` at wherever the
 * bundle is served. That resolves the deployment problem in both directions at once:
 * on a dev site they come from the package's own assets, and in production they come
 * from the project's CDN folder because that is where the bundle came from. No public
 * CDN to be blocked by the gov network, and no separate font handoff to be forgotten —
 * the fonts cannot arrive without the code, because they travel with it.
 */
import '@fortawesome/fontawesome-free/css/all.min.css';
// Latin subsets only. The full package ships Cyrillic, Greek and Vietnamese as well —
// 708KB of fonts for an English-language org chart, most of it for glyphs no node label
// here will ever contain. Latin plus latin-ext keeps accented European names working.
import '@fontsource/inter/latin-400.css';
import '@fontsource/inter/latin-500.css';
import '@fontsource/inter/latin-600.css';
import '@fontsource/inter/latin-700.css';

/**
 * Puts the graph engine on the page inside the web part's element.
 *
 * The engine addresses its own DOM by element id, exactly as it did as a standalone
 * page. That keeps it byte-identical to the shipped original — no rewrite of ~9,400
 * lines of working code to thread a root element through every lookup — at the cost of
 * one documented constraint: ONE INSTANCE PER PAGE. A single-part app page, which is
 * how this web part is meant to be used, satisfies that by construction. `isMounted()`
 * exists so the second instance can say so plainly instead of fighting the first over
 * shared ids.
 */

const STYLE_ID = 'erg-engine-styles';
const MOUNT_FLAG = 'data-erg-mounted';


export const isMounted = (): boolean => !!document.querySelector(`[${MOUNT_FLAG}]`);

export interface IMountOptions {
  container: HTMLElement;
  host: IErgHost;
  /** Optional extra stylesheet base. Blank is correct — the fonts are bundled. */
  assetBaseUrl?: string;
}

/**
 * Optional extra stylesheet base.
 *
 * The fonts themselves are bundled, so this is now only an escape hatch: a tenant that
 * must serve them from somewhere specific can point at it, and those rules land after
 * the bundled ones and win. Leaving it blank — the default — is correct everywhere.
 */
export const loadEngineFonts = (assetBaseUrl?: string): void => {
  const base = (assetBaseUrl || '').trim();
  if (!base) { return; }
  try {
    const root = base.replace(/\/*$/, '/');
    SPComponentLoader.loadCss(`${root}fonts/font-awesome/all.min.css`);
    SPComponentLoader.loadCss(`${root}fonts/inter/inter.css`);
  } catch {
    // Non-fatal: the bundled faces are already registered.
  }
};

/**
 * Host overrides, appended AFTER the engine stylesheet so they win on order.
 *
 * The engine stylesheet is generated from the single-file app and stays verbatim, so
 * everything the web part needs to change about its presentation lives here.
 */
const HOST_OVERRIDES = `
/* The engine was a whole-page application: .app is sized in VIEWPORT units
   (width:100vw; height:100vh). Inside a SharePoint page that is wider and taller than
   the space the web part actually occupies, so the details panel is pushed off the
   right edge and the page scrolls sideways. Size it to its CONTAINER instead; the
   shell decides how tall that container is. */
.erg-root { position: relative; width: 100%; height: 100%; overflow: hidden; }
.erg-root .app { width: 100%; height: 100%; }

/* The status bar advertised "Offline-ready · localStorage", a GitHub link and a
   byline — none of which are true once SharePoint is the source of truth. It is
   HIDDEN rather than removed from the template because the engine still writes to
   #footer-layout and #footer-hover; deleting the elements would throw.

   Reclaiming its 32px row is scoped above the engine's own stacked breakpoint
   (max-width:1000px), where the rows are auto-sized and a hidden footer already
   collapses to nothing. */
.erg-root footer.status-bar { display: none; }
@media (min-width: 1001px) {
  .erg-root .app { grid-template-rows: 64px 1fr 0; }
}

/* The header is a fixed 64px row holding a title and a one-line strapline. At full
   viewport width the strapline fits; in a SharePoint column it wraps to two lines and
   spills out of the row, overlapping the graph below. Clamp both lines to the width
   actually available and let them ellipsize.

   min-width:0 on every flex ancestor is the load-bearing part: a flex item defaults to
   min-width:auto, which refuses to shrink below its content, so text-overflow never
   engages no matter what is set on the text itself. */
.erg-root header.app-header { overflow: hidden; }
.erg-root header.app-header .brand,
.erg-root header.app-header .brand-text { min-width: 0; }
.erg-root header.app-header .brand-text h1,
.erg-root header.app-header .brand-text p {
  max-width: 100%;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
/* Below this the strapline has no room to say anything useful; the title carries it. */
@media (max-width: 900px) {
  .erg-root header.app-header .brand-text p { display: none; }
}

/* The brand claimed a fixed share of a 64px row that also carries the save chip, the
   stats and six tool buttons. At SharePoint widths that squeezed the stats until
   "<name> selected" wrapped onto a second line and spilled out of the row. Shrinking
   the brand hands that space back, and the stats are pinned to one line so they
   truncate rather than wrap. */
.erg-root header.app-header .brand { flex: 0 1 auto; }
.erg-root header.app-header .brand-mark {
  width: 30px; height: 30px; min-width: 30px; font-size: 14px;
}
.erg-root header.app-header .brand-text h1 { font-size: 15px; line-height: 1.25; }
.erg-root header.app-header .brand-text p { font-size: 11px; line-height: 1.3; }
.erg-root header.app-header .header-right { flex: 0 0 auto; min-width: 0; flex-wrap: nowrap; }
.erg-root header.app-header .stats,
.erg-root header.app-header .stats * { white-space: nowrap; flex-wrap: nowrap; }
`;

export const injectEngineStyles = (): void => {
  if (document.getElementById(STYLE_ID)) { return; }
  const style = document.createElement('style');
  style.id = STYLE_ID;
  style.appendChild(document.createTextNode(ENGINE_CSS + HOST_OVERRIDES));
  document.head.appendChild(style);
};

export const mountEngine = (options: IMountOptions): void => {
  injectEngineStyles();
  loadEngineFonts(options.assetBaseUrl);

  options.container.setAttribute(MOUNT_FLAG, 'true');
  options.container.classList.add('erg-root');
  options.container.innerHTML = ENGINE_HTML;

  // The engine reads the theme attribute from <html>, as it did standalone. Honour the
  // saved preference here because the head-level bootstrap script that used to do it is
  // not part of a web part.
  try {
    const saved = window.localStorage.getItem('erg.theme');
    const prefersLight = window.matchMedia && window.matchMedia('(prefers-color-scheme: light)').matches;
    if ((saved || (prefersLight ? 'light' : 'dark')) === 'light') {
      document.documentElement.setAttribute('data-theme', 'light');
    }
  } catch { /* private mode */ }

  startEngine(options.host);
};

export const unmountEngine = (container: HTMLElement): void => {
  container.removeAttribute(MOUNT_FLAG);
  container.innerHTML = '';
};
