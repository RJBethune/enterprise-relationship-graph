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

export const injectEngineStyles = (): void => {
  if (document.getElementById(STYLE_ID)) { return; }
  const style = document.createElement('style');
  style.id = STYLE_ID;
  style.appendChild(document.createTextNode(ENGINE_CSS));
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
