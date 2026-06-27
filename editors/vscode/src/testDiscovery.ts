import * as path from "path";
import * as vscode from "vscode";
import { XlflowChannels } from "./logging";
import { runXlflowJsonCommand } from "./xlflow";

export interface XlflowEnvelope {
  status?: string;
  error?: {
    code?: string;
    message?: string;
  };
  tests?: XlflowTestListPayload | XlflowTestRunPayload | XlflowTestRunItem[];
  logs?: string[];
}

export interface XlflowTestListPayload {
  root?: string;
  summary?: {
    files?: number;
    tests?: number;
  };
  items?: XlflowDiscoveredTest[];
}

export interface XlflowDiscoveredTest {
  module?: string;
  name?: string;
  qualified_name?: string;
  source_path?: string;
  line?: number;
  tags?: string[];
}

export interface XlflowTestRunPayload {
  summary?: {
    total?: number;
    passed?: number;
    failed?: number;
    inconclusive?: number;
  };
  items?: XlflowTestRunItem[];
}

export interface XlflowTestRunItem {
  module?: string;
  name?: string;
  status?: string;
  duration_ms?: number;
  error?: {
    code?: string;
    message?: string;
    source?: string;
    number?: number;
    line?: number;
  };
}

export async function discoverTests(
  folder: vscode.WorkspaceFolder,
  channels: XlflowChannels,
): Promise<{
  exitCode: number;
  stderr: string;
  json?: XlflowEnvelope;
  tests: XlflowDiscoveredTest[];
}> {
  const result = await runXlflowJsonCommand<XlflowEnvelope>(
    ["--json", "test", "list"],
    "xlflow test list",
    channels.output,
    { requireWorkspace: true, showCliUnavailable: false, workspaceFolder: folder },
  );
  return {
    exitCode: result.exitCode,
    stderr: result.stderr,
    json: result.json,
    tests: listDiscoveredTests(result.json),
  };
}

export function listDiscoveredTests(env: XlflowEnvelope | undefined): XlflowDiscoveredTest[] {
  const tests = env?.tests;
  if (isTestListPayload(tests) && Array.isArray(tests.items)) {
    return tests.items;
  }
  return [];
}

export function sourceUri(
  folder: vscode.WorkspaceFolder,
  sourcePath: string | undefined,
): vscode.Uri | undefined {
  const source = readNonEmpty(sourcePath);
  if (source === undefined) {
    return undefined;
  }
  if (path.isAbsolute(source)) {
    return vscode.Uri.file(source);
  }
  return vscode.Uri.joinPath(folder.uri, ...source.replace(/\\/g, "/").split("/"));
}

export function readNonEmpty(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }
  const trimmed = value.trim();
  return trimmed.length === 0 ? undefined : trimmed;
}

export function isTestRunPayload(value: unknown): value is XlflowTestRunPayload {
  return typeof value === "object" && value !== null && "items" in value;
}

function isTestListPayload(value: unknown): value is XlflowTestListPayload {
  return typeof value === "object" && value !== null && "items" in value;
}
