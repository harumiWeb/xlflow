import * as path from "path";
import type { XlflowStatusEnvelope } from "./session";

export interface FormBuildSource {
  name: string;
  specPath: string;
}

export interface FormBuildCommandResult {
  exitCode: number;
  json?: XlflowStatusEnvelope;
  stderr?: string;
}

export type FormBuildRunner = (args: string[]) => Promise<FormBuildCommandResult>;

export interface FormBuildOutcome {
  succeeded: string[];
  failed?: string;
  pending: string[];
  phase: "status" | "start" | "build" | "save" | "stop" | "complete";
  error?: string;
  createdSession: boolean;
}

/** Canonical specs are direct children of specs/, keyed by VBA's case-insensitive name. */
export function collectFormBuildSources(files: readonly string[]): {
  sources: FormBuildSource[];
  duplicates: string[][];
} {
  const groups = new Map<string, FormBuildSource[]>();
  for (const specPath of files) {
    if (!/^[^/\\]+\.(yaml|yml|json)$/i.test(specPath)) continue;
    const name = specPath.replace(/\.(yaml|yml|json)$/i, "");
    const key = name.toLowerCase();
    const group = groups.get(key) ?? [];
    group.push({ name, specPath });
    groups.set(key, group);
  }
  return {
    sources: [...groups.values()]
      .filter((group) => group.length === 1)
      .map((group) => group[0])
      .sort((a, b) => a.name.localeCompare(b.name)),
    duplicates: [...groups.values()]
      .filter((group) => group.length > 1)
      .map((group) => group.map((source) => source.specPath)),
  };
}

function normalizedWorkbook(value: string, projectRoot: string): string {
  const resolved = path.resolve(projectRoot, value.replace(/\\/g, "/"));
  return process.platform === "win32" ? resolved.toLowerCase() : resolved;
}

/** The runner is always the extension's coordinated CLI runner, fixed to one project. */
export async function executeFormBuild(
  sources: readonly FormBuildSource[],
  workbook: string,
  projectRoot: string,
  run: FormBuildRunner,
): Promise<FormBuildOutcome> {
  const outcome: FormBuildOutcome = {
    succeeded: [],
    pending: sources.map((source) => source.name),
    phase: "status",
    createdSession: false,
  };
  const invoke = async (args: string[]): Promise<FormBuildCommandResult | undefined> => {
    try {
      const result = await run(["--json", ...args]);
      if (result.exitCode !== 0 || result.json?.status !== "ok") {
        outcome.error =
          result.json?.error?.message ?? result.stderr?.trim() ?? "CLI command failed.";
        return undefined;
      }
      return result;
    } catch (error) {
      outcome.error = String(error);
      return undefined;
    }
  };
  if (sources.length === 0) {
    outcome.phase = "complete";
    return outcome;
  }
  const status = await invoke(["session", "status"]);
  if (!status) return outcome;
  const session = status.json?.session;
  if (!session || typeof session.active !== "boolean") {
    outcome.error = "Session status did not return a valid session payload.";
    return outcome;
  }
  if (status.json?.coordination?.recovery_required || status.json?.coordination?.busy) {
    outcome.error = "The workbook is busy or requires recovery.";
    return outcome;
  }
  const matches = (value: string | undefined): boolean =>
    value !== undefined &&
    normalizedWorkbook(value, projectRoot) === normalizedWorkbook(workbook, projectRoot);
  if (session.active) {
    if (!matches(session.workbook_path ?? session.metadata?.workbook_path)) {
      outcome.error = "The active session belongs to another workbook.";
      return outcome;
    }
  } else {
    // session start closes existing metadata's Excel process. Never use it to replace a live session.
    if (session.running || session.workbook_open) {
      outcome.error = "An existing Excel session must be resolved before building forms.";
      return outcome;
    }
    outcome.phase = "start";
    const started = await invoke(["session", "start"]);
    if (!started) return outcome;
    outcome.createdSession = true;
    if (
      started.json?.session?.active !== true ||
      !matches(started.json.session.workbook_path ?? started.json.session.metadata?.workbook_path)
    ) {
      outcome.error = "The new session could not be confirmed for the target workbook.";
      return outcome;
    }
  }
  for (const source of sources) {
    outcome.phase = "build";
    outcome.failed = source.name;
    outcome.pending.shift();
    if (
      !(await invoke(["form", "build", source.specPath, "--overwrite", "--session", "--no-save"]))
    ) {
      return outcome;
    }
    outcome.succeeded.push(source.name);
    outcome.failed = undefined;
  }
  outcome.phase = "save";
  if (!(await invoke(["save", "--session"]))) return outcome;
  if (outcome.createdSession) {
    outcome.phase = "stop";
    if (!(await invoke(["session", "stop"]))) return outcome;
  }
  outcome.phase = "complete";
  return outcome;
}
