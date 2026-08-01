import * as childProcess from "child_process";
import * as fs from "fs";
import * as path from "path";
import { readConfig } from "./config";

export interface RuleMetadata {
  id: string;
  inlineSuppressible: boolean;
}

export interface RulesCommandResult {
  exitCode: number;
  stdout: string;
  stderr: string;
}

export type RulesCommandRunner = (
  executable: string,
  args: readonly string[],
) => Promise<RulesCommandResult>;

export interface ResolvedExecutable {
  executable: string;
  identity: string;
}

export type ExecutableResolver = (configured: string) => ResolvedExecutable;

const rulesTimeoutMs = 5000;
const resolutionTtlMs = 1000;

export class XlflowRulesRegistryService {
  private cache: { identity: string; rules: ReadonlyMap<string, RuleMetadata> } | undefined;
  private pending:
    | { identity: string; promise: Promise<ReadonlyMap<string, RuleMetadata>> }
    | undefined;
  private resolvedCache:
    | { configured: string; resolved: ResolvedExecutable; at: number }
    | undefined;
  private generation = 0;

  constructor(
    private readonly executablePath: () => string = () => readConfig().path,
    private readonly run: RulesCommandRunner = runRulesCommand,
    private readonly resolveExecutable: ExecutableResolver = resolveExecutableIdentity,
    private readonly now: () => number = Date.now,
  ) {}

  async load(): Promise<ReadonlyMap<string, RuleMetadata>> {
    const resolved = this.currentResolution();
    if (this.cache?.identity === resolved.identity) {
      return this.cache.rules;
    }
    if (this.pending?.identity === resolved.identity) {
      return this.pending.promise;
    }
    if (
      (this.cache !== undefined && this.cache.identity !== resolved.identity) ||
      (this.pending !== undefined && this.pending.identity !== resolved.identity)
    ) {
      this.invalidate();
    }

    const generation = this.generation;
    const promise = this.loadForExecutable(resolved.executable)
      .then((rules) => {
        if (
          this.generation !== generation ||
          this.currentResolution().identity !== resolved.identity
        ) {
          throw new Error("Discarded stale xlflow rules metadata.");
        }
        this.cache = { identity: resolved.identity, rules };
        return rules;
      })
      .catch((error: unknown) => {
        if (this.generation === generation) {
          this.invalidate();
        }
        throw error;
      })
      .finally(() => {
        if (this.pending?.promise === promise) {
          this.pending = undefined;
        }
      });
    this.pending = { identity: resolved.identity, promise };
    return promise;
  }

  async isInlineSuppressible(code: string): Promise<boolean> {
    try {
      return (await this.load()).get(code.toUpperCase())?.inlineSuppressible === true;
    } catch {
      return false;
    }
  }

  invalidate(): void {
    this.generation++;
    this.cache = undefined;
    this.pending = undefined;
    this.resolvedCache = undefined;
  }

  private currentResolution(): ResolvedExecutable {
    const configured = this.executablePath();
    const now = this.now();
    if (
      this.resolvedCache !== undefined &&
      this.resolvedCache.configured === configured &&
      now - this.resolvedCache.at < resolutionTtlMs
    ) {
      return this.resolvedCache.resolved;
    }
    const resolved = this.resolveExecutable(configured);
    this.resolvedCache = { configured, resolved, at: now };
    return resolved;
  }

  private async loadForExecutable(executable: string): Promise<ReadonlyMap<string, RuleMetadata>> {
    const result = await this.run(executable, ["--json", "rules"]);
    if (result.exitCode !== 0) {
      throw new Error(result.stderr.trim() || `xlflow rules exited with code ${result.exitCode}.`);
    }
    let parsed: unknown;
    try {
      parsed = JSON.parse(result.stdout) as unknown;
    } catch {
      throw new Error("xlflow rules returned invalid JSON.");
    }
    return parseRulesEnvelope(parsed);
  }
}

export function resolveExecutableIdentity(configured: string): ResolvedExecutable {
  const command = configured.trim();
  const candidates = executableCandidates(command);
  for (const candidate of candidates) {
    try {
      const stat = fs.statSync(candidate);
      if (!stat.isFile()) {
        continue;
      }
      const executable = fs.realpathSync.native(candidate);
      const normalized = process.platform === "win32" ? executable.toLowerCase() : executable;
      return {
        executable,
        identity: `${normalized}\0${stat.size}\0${stat.mtimeMs}`,
      };
    } catch {
      // Keep searching PATH. An unresolved command is handled by spawn below.
    }
  }
  return {
    executable: command,
    identity: `unresolved\0${command}\0${process.env.PATH ?? ""}\0${process.env.PATHEXT ?? ""}`,
  };
}

function executableCandidates(command: string): string[] {
  if (command === "") {
    return [];
  }
  if (path.isAbsolute(command) || /[\\/]/.test(command)) {
    return [path.resolve(command)];
  }
  const extensions =
    process.platform === "win32" ? executableExtensions(command, process.env.PATHEXT) : [""];
  const candidates: string[] = [];
  for (const directory of (process.env.PATH ?? "").split(path.delimiter)) {
    if (directory.trim() === "") {
      continue;
    }
    for (const extension of extensions) {
      candidates.push(path.join(directory, command + extension));
    }
  }
  return candidates;
}

function executableExtensions(command: string, pathExt: string | undefined): string[] {
  if (path.extname(command) !== "") {
    return [""];
  }
  return (pathExt ?? ".COM;.EXE;.BAT;.CMD")
    .split(";")
    .filter((extension) => extension !== "")
    .map((extension) => extension.toLowerCase());
}

export function parseRulesEnvelope(value: unknown): ReadonlyMap<string, RuleMetadata> {
  if (!isObject(value)) {
    throw new Error("xlflow rules returned an invalid response envelope.");
  }
  const rules = value.rules;
  if (
    value.status !== "ok" ||
    value.command !== "rules" ||
    !isObject(rules) ||
    rules.schema_version !== 1 ||
    !Array.isArray(rules.items)
  ) {
    throw new Error("xlflow rules returned an unsupported response schema.");
  }

  const items = new Map<string, RuleMetadata>();
  for (const item of rules.items) {
    if (!isObject(item)) {
      throw new Error("xlflow rules returned invalid rule metadata.");
    }
    const id = typeof item.id === "string" ? item.id.trim().toUpperCase() : "";
    if (!/^(?:VB|VBA)\d{3}$/.test(id) || typeof item.inline_suppressible !== "boolean") {
      throw new Error("xlflow rules returned invalid rule metadata.");
    }
    if (items.has(id)) {
      throw new Error(`xlflow rules returned duplicate metadata for ${id}.`);
    }
    items.set(id, { id, inlineSuppressible: item.inline_suppressible });
  }
  return items;
}

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function runRulesCommand(executable: string, args: readonly string[]): Promise<RulesCommandResult> {
  return new Promise((resolve, reject) => {
    let settled = false;
    const stdoutChunks: Buffer[] = [];
    const stderrChunks: Buffer[] = [];
    const child = childProcess.spawn(executable, [...args], { windowsHide: true });
    const timer = setTimeout(() => {
      if (settled) {
        return;
      }
      settled = true;
      child.kill();
      reject(new Error("xlflow rules timed out."));
    }, rulesTimeoutMs);
    const finish = (): void => clearTimeout(timer);
    child.stdout.on("data", (data: Buffer) => stdoutChunks.push(data));
    child.stderr.on("data", (data: Buffer) => stderrChunks.push(data));
    child.on("error", (error) => {
      if (settled) {
        return;
      }
      settled = true;
      finish();
      reject(error);
    });
    child.on("close", (code) => {
      if (settled) {
        return;
      }
      settled = true;
      finish();
      resolve({
        exitCode: code ?? -1,
        stdout: Buffer.concat(stdoutChunks).toString("utf8"),
        stderr: Buffer.concat(stderrChunks).toString("utf8"),
      });
    });
  });
}
