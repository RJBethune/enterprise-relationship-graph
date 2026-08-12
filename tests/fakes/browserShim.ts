/**
 * The few browser globals the data layer touches, so the safety suite can exercise
 * code paths that would otherwise be skipped on Node.
 *
 * The schema-health cache is the reason this exists: it lives in localStorage and is
 * wrapped in try/catch, so without a shim every cache test would silently pass by
 * doing nothing — the worst kind of green.
 */
const store: { [key: string]: string } = {};

export const resetBrowserStorage = (): void => {
  for (const key of Object.keys(store)) { delete store[key]; }
};

const g = globalThis as { window?: unknown };
if (!g.window) {
  g.window = {
    localStorage: {
      getItem: (k: string): string | null => (k in store ? store[k] : null),
      setItem: (k: string, v: string): void => { store[k] = String(v); },
      removeItem: (k: string): void => { delete store[k]; }
    }
  };
}
