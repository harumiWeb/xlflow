import { randomUUID } from "crypto";

export interface StartupTelemetryRecord {
  operation: "lsp/startup";
  side: "client";
  startup_id: string;
  event: string;
  elapsed_ms: number;
  wall_time_unix_ms: number;
  outcome: "ok" | "error";
  result_count?: number;
}

export type StartupClock = () => number;

export class StartupTelemetry {
  private readonly seen = new Set<string>();
  private readonly idFactory: () => string;
  private readonly started: number;

  public readonly id: string;

  public constructor(
    private readonly enabled: boolean,
    private readonly sink: (record: StartupTelemetryRecord) => void,
    private readonly clock: StartupClock = () => performance.now(),
    idFactory: () => string = randomUUID,
    started?: number,
  ) {
    this.idFactory = idFactory;
    this.id = enabled ? idFactory() : "";
    this.started = started ?? (enabled ? clock() : 0);
  }

  public newAttempt(started?: number): StartupTelemetry {
    const attemptStarted = started ?? (this.enabled ? this.clock() : undefined);
    return new StartupTelemetry(
      this.enabled,
      this.sink,
      this.clock,
      this.idFactory,
      attemptStarted,
    );
  }

  public withEnabled(enabled: boolean, started?: number): StartupTelemetry {
    const attemptStarted = started ?? (enabled ? this.clock() : undefined);
    return new StartupTelemetry(enabled, this.sink, this.clock, this.idFactory, attemptStarted);
  }

  public get isEnabled(): boolean {
    return this.enabled;
  }

  public mark(
    event: string,
    options: { outcome?: "ok" | "error"; resultCount?: number; once?: boolean } = {},
  ): void {
    if (!this.enabled || event.length === 0) {
      return;
    }
    const once = options.once !== false;
    if (once && this.seen.has(event)) {
      return;
    }
    if (once) {
      this.seen.add(event);
    }
    const now = this.clock();
    const record: StartupTelemetryRecord = {
      operation: "lsp/startup",
      side: "client",
      startup_id: this.id,
      event,
      elapsed_ms: Math.max(0, now - this.started),
      wall_time_unix_ms: Date.now(),
      outcome: options.outcome ?? "ok",
    };
    if (options.resultCount !== undefined) {
      record.result_count = options.resultCount;
    }
    this.sink(record);
  }
}

export function hasHoverResult(value: unknown): boolean {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  const contents = (value as { contents?: unknown }).contents;
  if (Array.isArray(contents)) {
    return contents.some(hasHoverContent);
  }
  return hasHoverContent(contents);
}

function hasHoverContent(value: unknown): boolean {
  if (typeof value === "string") {
    return value.trim().length > 0;
  }
  if (typeof value !== "object" || value === null) {
    return false;
  }
  const record = value as { value?: unknown };
  if (record.value !== undefined) {
    return typeof record.value === "string" ? record.value.trim().length > 0 : false;
  }
  return Object.keys(value).length > 0;
}

export function hasDefinitionResult(value: unknown): boolean {
  return Array.isArray(value) ? value.length > 0 : value !== undefined && value !== null;
}

export function hasCompletionResult(value: unknown): boolean {
  if (Array.isArray(value)) {
    return value.length > 0;
  }
  if (typeof value !== "object" || value === null) {
    return false;
  }
  const items = (value as { items?: unknown }).items;
  return Array.isArray(items) && items.length > 0;
}

export function completionResultCount(value: unknown): number {
  if (Array.isArray(value)) {
    return value.length;
  }
  if (typeof value !== "object" || value === null) {
    return 0;
  }
  const items = (value as { items?: unknown }).items;
  return Array.isArray(items) ? items.length : 0;
}
