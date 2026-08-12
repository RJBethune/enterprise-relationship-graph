/**
 * Safety-suite entry point. Compiled with the project's own TypeScript and run on
 * plain Node — no jest, no foreign node_modules. Chained into `check`, `test` and
 * `ship` so a release cannot go out with this logic broken.
 */
import './bundle.test';
import './diff.test';
import './merge.test';
import './planner.test';
import './batch.test';
import './writeQueue.test';
import './integration.test';
import { runAll } from './harness';

void (async (): Promise<void> => {
  const failures = await runAll();
  process.exitCode = failures > 0 ? 1 : 0;
})();
