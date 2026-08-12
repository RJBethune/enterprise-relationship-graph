import { SPComponentLoader } from '@microsoft/sp-loader';
import { ENGINE_CSS, ENGINE_HTML } from './engineAssets';
import { startEngine } from './engine';
import { IErgHost } from './hostContract';

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

/**
 * Font Awesome glyphs and Inter are drawn ONTO THE CANVAS — node icons are text
 * rendered in those faces, not decoration. Where they load from differs by environment,
 * and getting that backwards is a silent failure (missing icons, working app), so both
 * paths are explicit:
 *
 *  - PRODUCTION sites must set `assetBaseUrl` to the project's CDN folder, the only
 *    remote origin the download policy permits. The CDN handoff has to include the font
 *    files under `fonts/`.
 *  - DEV and test sites leave it blank and get the public CDNs, exactly as the v1.x
 *    single-file app did — so a test-catalog deployment works with no CDN work at all.
 *
 * Either way the engine tolerates them not arriving: it races the load against a 3s
 * timeout and renders with fallbacks.
 */
export const PRODUCTION_ASSET_BASE = 'https://irm.azureedge.us/M/enterprise-relationship-graph/';

const PUBLIC_FONT_AWESOME = 'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css';
const PUBLIC_INTER = 'https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap';

export const isMounted = (): boolean => !!document.querySelector(`[${MOUNT_FLAG}]`);

export interface IMountOptions {
  container: HTMLElement;
  host: IErgHost;
  /** Base URL for font assets. Defaults to the project's CDN folder. */
  assetBaseUrl?: string;
}

export const loadEngineFonts = (assetBaseUrl?: string): void => {
  const base = (assetBaseUrl || '').trim();
  try {
    if (!base) {
      SPComponentLoader.loadCss(PUBLIC_FONT_AWESOME);
      SPComponentLoader.loadCss(PUBLIC_INTER);
      return;
    }
    const root = base.replace(/\/*$/, '/');
    SPComponentLoader.loadCss(`${root}fonts/font-awesome/all.min.css`);
    SPComponentLoader.loadCss(`${root}fonts/inter/inter.css`);
  } catch {
    // Non-fatal by design — see the note above.
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
