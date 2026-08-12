/**
 * A ~60-line test harness with no dependencies.
 *
 * Per SPFX-ARCHITECTURE.md §4, safety-critical pure logic is compiled with the
 * project's OWN TypeScript and run on plain Node, so the suite that gates a release
 * cannot be broken by a missing or mismatched external runner. Everything tested here
 * is deliberately free of SPFx and DOM dependencies.
 */

export interface ITestCase { name: string; fn: () => void | Promise<void>; }

const suites: { name: string; cases: ITestCase[] }[] = [];
let current: { name: string; cases: ITestCase[] } | null = null;

export const suite = (name: string, body: () => void): void => {
  current = { name, cases: [] };
  suites.push(current);
  body();
  current = null;
};

export const test = (name: string, fn: () => void | Promise<void>): void => {
  if (!current) { throw new Error('test() called outside suite()'); }
  current.cases.push({ name, fn });
};

export class AssertionError extends Error {}

const fail = (message: string): never => { throw new AssertionError(message); };

export const assert = {
  ok(value: unknown, message?: string): void {
    if (!value) { fail(message || `Expected truthy, got ${JSON.stringify(value)}`); }
  },
  equal(actual: unknown, expected: unknown, message?: string): void {
    if (actual !== expected) {
      fail(message || `Expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
    }
  },
  deepEqual(actual: unknown, expected: unknown, message?: string): void {
    const a = JSON.stringify(actual);
    const b = JSON.stringify(expected);
    if (a !== b) { fail(message || `Expected ${b}\n            got ${a}`); }
  },
  includes(haystack: string, needle: string, message?: string): void {
    if (haystack.indexOf(needle) < 0) {
      fail(message || `Expected to find "${needle}" in "${haystack.slice(0, 200)}"`);
    }
  },
  throws(fn: () => unknown, message?: string): void {
    try { fn(); } catch (_e) { return; }
    fail(message || 'Expected the call to throw, but it returned normally');
  }
};

export const runAll = async (): Promise<number> => {
  let passed = 0;
  const failures: string[] = [];

  for (const s of suites) {
    process.stdout.write(`\n${s.name}\n`);
    for (const c of s.cases) {
      try {
        await c.fn();
        passed++;
        process.stdout.write(`  ok   ${c.name}\n`);
      } catch (e) {
        const msg = e instanceof Error ? e.message : String(e);
        failures.push(`${s.name} > ${c.name}\n     ${msg}`);
        process.stdout.write(`  FAIL ${c.name}\n     ${msg}\n`);
      }
    }
  }

  process.stdout.write(`\n${passed} passed, ${failures.length} failed\n`);
  if (failures.length > 0) {
    process.stdout.write(`\nFailures:\n${failures.map((f) => '  - ' + f).join('\n')}\n`);
  }
  return failures.length;
};
