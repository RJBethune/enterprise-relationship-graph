/**
 * The throttle-proofing layer for per-item storage.
 *
 * SharePoint Online meters WRITE OPERATIONS, not HTTP connections. A $batch of 100
 * writes still costs 100 operations, so batching alone does not prevent a 429 — it
 * only saves round trips. What actually prevents throttling is doing fewer writes:
 *
 *  1. COALESCE. Renaming a node fires an op per keystroke; only the last one matters.
 *     Ops are keyed by entity, so a burst collapses to one write per entity touched.
 *  2. DEBOUNCE. Nothing is sent until edits stop for `debounceMs`, so a rapid sequence
 *     of edits becomes one flush.
 *  3. BOUNDED CONCURRENCY. One flush in flight at a time. A second flush queues behind
 *     it rather than racing it, which also guarantees writes land in order.
 *  4. BACK OFF AND OBEY. On 429/503 the server's Retry-After is honoured exactly;
 *     without one, exponential backoff with a cap. Retrying faster than the server
 *     asked is what turns a brief throttle into a long one.
 *
 * Clock and transport are injected so all of this is unit-testable with no timers and
 * no network — the retry ladder is exactly the code you cannot afford to test only in
 * production.
 */

export interface IQueueClock {
  now(): number;
  setTimeout(fn: () => void, ms: number): unknown;
  clearTimeout(handle: unknown): void;
}

export const systemClock: IQueueClock = {
  now: () => Date.now(),
  setTimeout: (fn, ms) => setTimeout(fn, ms),
  clearTimeout: (h) => clearTimeout(h as ReturnType<typeof setTimeout>)
};

export class ThrottledError extends Error {
  public readonly retryAfterMs: number;
  public constructor(retryAfterMs: number) {
    super(`Throttled; retry after ${retryAfterMs}ms`);
    this.name = 'ThrottledError';
    this.retryAfterMs = retryAfterMs;
  }
}

export interface IWriteQueueOptions<TItem> {
  /** Identity used for coalescing; same key means only the last item survives. */
  keyOf(item: TItem): string;
  /** Send one batch. Throwing ThrottledError triggers the honoured-delay path. */
  flush(items: TItem[]): Promise<void>;
  /** Optional ordering applied to each batch immediately before flushing. */
  order?(items: TItem[]): TItem[];
  debounceMs?: number;
  /** Max items per flush call. Batches beyond this are sent sequentially. */
  batchSize?: number;
  maxRetries?: number;
  baseBackoffMs?: number;
  maxBackoffMs?: number;
  clock?: IQueueClock;
  onStateChange?(state: WriteQueueState): void;
  onError?(error: Error, attempt: number): void;
}

export type WriteQueueState = 'idle' | 'pending' | 'writing' | 'retrying' | 'failed';

export class WriteQueue<TItem> {
  private readonly opts: Required<Omit<IWriteQueueOptions<TItem>, 'onStateChange' | 'onError' | 'order'>> &
    Pick<IWriteQueueOptions<TItem>, 'onStateChange' | 'onError' | 'order'>;
  private readonly clock: IQueueClock;
  private pending: Map<string, TItem> = new Map();
  private timer: unknown = null;
  private inFlight: boolean = false;
  private state: WriteQueueState = 'idle';
  private lastError: Error | null = null;

  public constructor(options: IWriteQueueOptions<TItem>) {
    this.opts = {
      keyOf: options.keyOf,
      flush: options.flush,
      order: options.order,
      debounceMs: options.debounceMs === undefined ? 800 : options.debounceMs,
      batchSize: options.batchSize === undefined ? 50 : options.batchSize,
      maxRetries: options.maxRetries === undefined ? 5 : options.maxRetries,
      baseBackoffMs: options.baseBackoffMs === undefined ? 1000 : options.baseBackoffMs,
      maxBackoffMs: options.maxBackoffMs === undefined ? 30000 : options.maxBackoffMs,
      clock: options.clock || systemClock,
      onStateChange: options.onStateChange,
      onError: options.onError
    };
    this.clock = this.opts.clock;
  }

  public getState(): WriteQueueState { return this.state; }
  public getPendingCount(): number { return this.pending.size; }
  public getLastError(): Error | null { return this.lastError; }

  public enqueue(items: TItem[]): void {
    for (const item of items) { this.pending.set(this.opts.keyOf(item), item); }
    if (this.pending.size === 0) { return; }
    this.setState('pending');
    this.arm();
  }

  /** Send everything now — used on tab hide/close and before a project switch. */
  public async flushNow(): Promise<void> {
    this.disarm();
    await this.drain();
  }

  public dispose(): void {
    this.disarm();
    this.pending.clear();
  }

  private setState(next: WriteQueueState): void {
    if (this.state === next) { return; }
    this.state = next;
    if (this.opts.onStateChange) { this.opts.onStateChange(next); }
  }

  private arm(): void {
    this.disarm();
    this.timer = this.clock.setTimeout(() => {
      this.timer = null;
      void this.drain();
    }, this.opts.debounceMs);
  }

  private disarm(): void {
    if (this.timer !== null) { this.clock.clearTimeout(this.timer); this.timer = null; }
  }

  private async drain(): Promise<void> {
    // One flush at a time: the in-flight run picks up anything queued behind it when
    // it loops, so writes stay ordered and concurrency stays bounded at one.
    if (this.inFlight) { return; }
    this.inFlight = true;
    try {
      while (this.pending.size > 0) {
        const all = Array.from(this.pending.values());
        const ordered = this.opts.order ? this.opts.order(all) : all;
        const batch = ordered.slice(0, this.opts.batchSize);

        this.setState('writing');
        const sent = await this.sendWithRetry(batch);
        if (!sent) { this.setState('failed'); return; }

        for (const item of batch) { this.pending.delete(this.opts.keyOf(item)); }
      }
      this.lastError = null;
      this.setState('idle');
    } finally {
      this.inFlight = false;
    }
  }

  private async sendWithRetry(batch: TItem[]): Promise<boolean> {
    for (let attempt = 0; attempt <= this.opts.maxRetries; attempt++) {
      try {
        await this.opts.flush(batch);
        return true;
      } catch (e) {
        const err = e instanceof Error ? e : new Error(String(e));
        this.lastError = err;
        if (this.opts.onError) { this.opts.onError(err, attempt); }
        if (attempt === this.opts.maxRetries) { return false; }

        const delay = err instanceof ThrottledError
          // The server told us exactly how long to wait. Guessing shorter is what
          // escalates a soft throttle into a hard one.
          ? err.retryAfterMs
          : Math.min(this.opts.baseBackoffMs * Math.pow(2, attempt), this.opts.maxBackoffMs);

        this.setState('retrying');
        await this.sleep(delay);
      }
    }
    return false;
  }

  private sleep(ms: number): Promise<void> {
    return new Promise<void>((resolve) => { this.clock.setTimeout(resolve, ms); });
  }
}

/** Parse a Retry-After header (seconds, or an HTTP date) into milliseconds. */
export const parseRetryAfter = (headerValue: string | null, nowMs: number): number => {
  if (!headerValue) { return 0; }
  const seconds = Number(headerValue);
  if (!isNaN(seconds) && seconds >= 0) { return Math.round(seconds * 1000); }
  const when = Date.parse(headerValue);
  if (!isNaN(when)) { return Math.max(0, when - nowMs); }
  return 0;
};
