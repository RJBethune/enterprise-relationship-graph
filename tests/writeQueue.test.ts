import { suite, test, assert } from './harness';
import { WriteQueue, ThrottledError, parseRetryAfter, IQueueClock } from '../src/services/WriteQueue';
import { GraphOp, opKey, orderOps } from '../src/model/ops';

/**
 * Time is injected, so the retry ladder runs instantly and every delay the queue chose
 * is observable. This is the code you cannot afford to test only in production: a bug
 * in the backoff is invisible until the day SharePoint throttles, and then it makes the
 * throttling worse.
 */
class FakeClock implements IQueueClock {
  public readonly delays: number[] = [];
  private t: number = 0;
  public now(): number { return this.t; }
  public setTimeout(fn: () => void, ms: number): unknown {
    this.delays.push(ms);
    return setImmediate(() => { this.t += ms; fn(); });
  }
  public clearTimeout(handle: unknown): void { clearImmediate(handle); }
}

const upsert = (id: string, label: string): GraphOp =>
  ({ kind: 'upsertNode', node: { id, label, type: 'Office' } });

suite('write queue: fewer writes, not just fewer round trips', () => {
  test('a burst of edits to one node becomes a single write', async () => {
    const clock = new FakeClock();
    const flushed: GraphOp[][] = [];
    const queue = new WriteQueue<GraphOp>({
      keyOf: opKey, clock, debounceMs: 0,
      flush: async (ops) => { flushed.push(ops); }
    });

    queue.enqueue([upsert('a', 'O'), upsert('a', 'Op'), upsert('a', 'Ops')]);
    await queue.flushNow();

    assert.equal(flushed.length, 1);
    assert.equal(flushed[0].length, 1, 'three keystrokes on one node cost one write');
    assert.equal((flushed[0][0] as { node: { label: string } }).node.label, 'Ops');
  });

  test('ops for different entities are all sent', async () => {
    const clock = new FakeClock();
    const flushed: GraphOp[][] = [];
    const queue = new WriteQueue<GraphOp>({
      keyOf: opKey, clock, debounceMs: 0,
      flush: async (ops) => { flushed.push(ops); }
    });
    queue.enqueue([upsert('a', 'A'), upsert('b', 'B'), upsert('c', 'C')]);
    await queue.flushNow();
    assert.equal(flushed[0].length, 3);
  });

  test('a large change set is split into bounded batches', async () => {
    const clock = new FakeClock();
    const sizes: number[] = [];
    const queue = new WriteQueue<GraphOp>({
      keyOf: opKey, clock, debounceMs: 0, batchSize: 10,
      flush: async (ops) => { sizes.push(ops.length); }
    });
    const ops: GraphOp[] = [];
    for (let i = 0; i < 25; i++) { ops.push(upsert('n' + i, 'N' + i)); }
    queue.enqueue(ops);
    await queue.flushNow();
    assert.deepEqual(sizes, [10, 10, 5]);
  });

  test('the ordering function is applied before sending', async () => {
    const clock = new FakeClock();
    let sent: GraphOp[] = [];
    const queue = new WriteQueue<GraphOp>({
      keyOf: opKey, order: orderOps, clock, debounceMs: 0,
      flush: async (ops) => { sent = ops; }
    });
    queue.enqueue([
      { kind: 'upsertEdge', edge: { id: 'e1', source: 'a', target: 'b', type: 'X' } },
      { kind: 'deleteEdge', id: 'e0' },
      upsert('a', 'A')
    ]);
    await queue.flushNow();
    assert.deepEqual(sent.map((o) => o.kind), ['deleteEdge', 'upsertNode', 'upsertEdge']);
  });
});

suite('write queue: backing off the way the server asked', () => {
  test('a throttle response is retried after exactly the delay the server named', async () => {
    const clock = new FakeClock();
    let attempts = 0;
    const queue = new WriteQueue<GraphOp>({
      keyOf: opKey, clock, debounceMs: 0,
      flush: async () => {
        attempts++;
        if (attempts === 1) { throw new ThrottledError(7000); }
      }
    });
    queue.enqueue([upsert('a', 'A')]);
    await queue.flushNow();

    assert.equal(attempts, 2);
    assert.ok(clock.delays.indexOf(7000) >= 0, 'Retry-After must be honoured exactly, not guessed at');
    assert.equal(queue.getState(), 'idle');
  });

  test('other failures back off exponentially, and are capped', async () => {
    const clock = new FakeClock();
    let attempts = 0;
    const queue = new WriteQueue<GraphOp>({
      keyOf: opKey, clock, debounceMs: 0, baseBackoffMs: 100, maxBackoffMs: 250, maxRetries: 4,
      flush: async () => { attempts++; throw new Error('boom'); }
    });
    queue.enqueue([upsert('a', 'A')]);
    await queue.flushNow();

    assert.equal(attempts, 5, 'the first try plus four retries');
    // delays also records the 0ms debounce arm from enqueue(); the backoff ladder is
    // the non-zero waits.
    assert.deepEqual(
      clock.delays.filter((d) => d > 0), [100, 200, 250, 250], 'doubling, then capped'
    );
    assert.equal(queue.getState(), 'failed');
  });

  test('work is retained after a permanent failure so nothing is silently lost', async () => {
    const clock = new FakeClock();
    const queue = new WriteQueue<GraphOp>({
      keyOf: opKey, clock, debounceMs: 0, maxRetries: 0,
      flush: async () => { throw new Error('offline'); }
    });
    queue.enqueue([upsert('a', 'A')]);
    await queue.flushNow();
    assert.equal(queue.getPendingCount(), 1);
    assert.equal(queue.getState(), 'failed');
    assert.equal((queue.getLastError() as Error).message, 'offline');
  });

  test('a recovered flush clears the queue and the error', async () => {
    const clock = new FakeClock();
    let fail = true;
    const queue = new WriteQueue<GraphOp>({
      keyOf: opKey, clock, debounceMs: 0, maxRetries: 0,
      flush: async () => { if (fail) { throw new Error('offline'); } }
    });
    queue.enqueue([upsert('a', 'A')]);
    await queue.flushNow();
    fail = false;
    await queue.flushNow();
    assert.equal(queue.getPendingCount(), 0);
    assert.equal(queue.getState(), 'idle');
  });

  test('only one flush runs at a time', async () => {
    const clock = new FakeClock();
    let inFlight = 0;
    let maxConcurrent = 0;
    const queue = new WriteQueue<GraphOp>({
      keyOf: opKey, clock, debounceMs: 0, batchSize: 1,
      flush: async () => {
        inFlight++;
        maxConcurrent = Math.max(maxConcurrent, inFlight);
        await new Promise<void>((r) => setImmediate(() => r()));
        inFlight--;
      }
    });
    queue.enqueue([upsert('a', 'A'), upsert('b', 'B'), upsert('c', 'C')]);
    await Promise.all([queue.flushNow(), queue.flushNow(), queue.flushNow()]);
    assert.equal(maxConcurrent, 1, 'concurrent flushes would reorder writes');
  });
});

suite('write queue: Retry-After parsing', () => {
  test('a seconds value', () => { assert.equal(parseRetryAfter('12', 0), 12000); });
  test('an HTTP date', () => {
    const now = Date.parse('2026-01-01T00:00:00Z');
    assert.equal(parseRetryAfter(new Date(now + 5000).toUTCString(), now), 5000);
  });
  test('a missing header means no imposed delay', () => { assert.equal(parseRetryAfter(null, 0), 0); });
  test('garbage is treated as no delay rather than NaN', () => { assert.equal(parseRetryAfter('soon', 0), 0); });
});
