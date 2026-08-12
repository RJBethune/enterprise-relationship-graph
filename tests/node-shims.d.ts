/**
 * The few Node globals the safety suite needs, declared locally.
 *
 * The suite deliberately depends on NO external @types packages so it can run with
 * nothing installed but the project's own TypeScript — that is what makes it usable as
 * a release gate on a machine or CI runner that has not done a full SPFx install.
 */
declare const process: {
  stdout: { write(s: string): void };
  exitCode: number | undefined;
  exit(code?: number): void;
};

declare const globalThis: { window?: unknown };

declare const window: {
  localStorage: {
    getItem(k: string): string | null;
    setItem(k: string, v: string): void;
    removeItem(k: string): void;
  };
};

declare function setImmediate(fn: () => void): unknown;
declare function clearImmediate(handle: unknown): void;
