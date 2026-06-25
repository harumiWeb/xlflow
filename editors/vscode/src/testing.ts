import * as vscode from "vscode";
import { readConfig } from "./config";
import { XlflowChannels } from "./logging";
import { runXlflowJsonCommand } from "./xlflow";

interface XlflowEnvelope {
  status?: string;
  error?: {
    code?: string;
    message?: string;
  };
  tests?: XlflowTestListPayload | XlflowTestRunPayload | XlflowTestRunItem[];
  logs?: string[];
}

interface XlflowTestListPayload {
  root?: string;
  summary?: {
    files?: number;
    tests?: number;
  };
  items?: XlflowDiscoveredTest[];
}

interface XlflowDiscoveredTest {
  module?: string;
  name?: string;
  qualified_name?: string;
  source_path?: string;
  line?: number;
  tags?: string[];
}

interface XlflowTestRunPayload {
  summary?: {
    total?: number;
    passed?: number;
    failed?: number;
    inconclusive?: number;
  };
  items?: XlflowTestRunItem[];
}

interface XlflowTestRunItem {
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

interface TestMetadata {
  workspaceFolder: vscode.WorkspaceFolder;
  module: string;
  name: string;
  qualifiedName: string;
}

export class XlflowTestController implements vscode.Disposable {
  private readonly controller: vscode.TestController;
  private readonly metadata = new Map<string, TestMetadata>();
  private readonly runProfile: vscode.TestRunProfile;

  constructor(private readonly channels: XlflowChannels) {
    this.controller = vscode.tests.createTestController("xlflow-vba-tests", "xlflow VBA Tests");
    this.controller.refreshHandler = async () => {
      await this.discoverAll();
    };
    this.runProfile = this.controller.createRunProfile(
      "Run Tests",
      vscode.TestRunProfileKind.Run,
      async (request, token) => {
        await this.runTests(request, token);
      },
      true,
    );
  }

  dispose(): void {
    this.runProfile.dispose();
    this.controller.dispose();
  }

  async refresh(): Promise<void> {
    await this.discoverAll();
  }

  async refreshAuto(): Promise<void> {
    if (!readConfig().testingAutoDiscover) {
      return;
    }
    await this.discoverAll({ xlflowWorkspacesOnly: true });
  }

  private async discoverAll(options: { xlflowWorkspacesOnly?: boolean } = {}): Promise<void> {
    this.metadata.clear();
    const folders = await discoverableWorkspaceFolders(options.xlflowWorkspacesOnly === true);
    if (folders.length === 0) {
      this.controller.items.replace([]);
      return;
    }

    const roots: vscode.TestItem[] = [];
    for (const folder of folders) {
      const moduleItems = await this.discoverFolder(folder);
      if (folders.length === 1) {
        roots.push(...moduleItems);
        continue;
      }
      const folderItem = this.controller.createTestItem(
        workspaceItemId(folder),
        folder.name,
        folder.uri,
      );
      folderItem.children.replace(moduleItems);
      roots.push(folderItem);
    }
    this.controller.items.replace(roots);
  }

  private async discoverFolder(folder: vscode.WorkspaceFolder): Promise<vscode.TestItem[]> {
    const result = await runXlflowJsonCommand<XlflowEnvelope>(
      ["--json", "test", "list"],
      "xlflow test list",
      this.channels.output,
      { requireWorkspace: true, workspaceFolder: folder },
    );
    const tests = listDiscoveredTests(result.json);
    if (result.exitCode !== 0 || tests.length === 0) {
      if (result.exitCode !== 0) {
        this.channels.output.appendLine(`[error] xlflow test list failed for ${folder.uri.fsPath}`);
      }
      return [];
    }

    const modules = new Map<string, vscode.TestItem>();
    for (const test of tests) {
      const module = readNonEmpty(test.module);
      const name = readNonEmpty(test.name);
      if (module === undefined || name === undefined) {
        continue;
      }
      const qualifiedName = readNonEmpty(test.qualified_name) ?? `${module}.${name}`;
      const moduleId = `${folder.uri.toString()}::module::${module}`;
      let moduleItem = modules.get(module);
      if (moduleItem === undefined) {
        moduleItem = this.controller.createTestItem(moduleId, module);
        modules.set(module, moduleItem);
      }

      const uri = sourceUri(folder, test.source_path);
      const item = this.controller.createTestItem(
        `${folder.uri.toString()}::test::${qualifiedName}`,
        name,
        uri,
      );
      const line = typeof test.line === "number" && test.line > 0 ? test.line - 1 : 0;
      item.range = new vscode.Range(line, 0, line, 0);
      item.description =
        Array.isArray(test.tags) && test.tags.length > 0 ? test.tags.join(", ") : undefined;
      this.metadata.set(item.id, { workspaceFolder: folder, module, name, qualifiedName });
      moduleItem.children.add(item);
    }
    return [...modules.values()].sort((a, b) => a.label.localeCompare(b.label));
  }

  private async runTests(
    request: vscode.TestRunRequest,
    token: vscode.CancellationToken,
  ): Promise<void> {
    if (this.metadata.size === 0) {
      await this.discoverAll();
    }

    const run = this.controller.createTestRun(request);
    try {
      const queue = this.collectRequestedTests(request);
      for (const item of queue) {
        if (token.isCancellationRequested) {
          run.skipped(item);
          continue;
        }
        const metadata = this.metadata.get(item.id);
        if (metadata === undefined) {
          run.skipped(item);
          continue;
        }
        run.enqueued(item);
        await this.runSingleTest(run, item, metadata);
      }
    } finally {
      run.end();
    }
  }

  private collectRequestedTests(request: vscode.TestRunRequest): vscode.TestItem[] {
    const tests: vscode.TestItem[] = [];
    const excluded = new Set((request.exclude ?? []).map((item) => item.id));
    const visit = (item: vscode.TestItem): void => {
      if (excluded.has(item.id)) {
        return;
      }
      if (this.metadata.has(item.id)) {
        tests.push(item);
        return;
      }
      item.children.forEach(visit);
    };
    const included = request.include ?? collectionItems(this.controller.items);
    for (const item of included) {
      visit(item);
    }
    return tests;
  }

  private async runSingleTest(
    run: vscode.TestRun,
    item: vscode.TestItem,
    metadata: TestMetadata,
  ): Promise<void> {
    run.started(item);
    const result = await runXlflowJsonCommand<XlflowEnvelope>(
      ["--json", "test", "--module", metadata.module, "--filter", metadata.name],
      `xlflow test ${metadata.qualifiedName}`,
      this.channels.output,
      { requireWorkspace: true, workspaceFolder: metadata.workspaceFolder },
    );
    const runItem = findRunResult(result.json, metadata);
    if (runItem?.status === "passed") {
      run.passed(item, runItem.duration_ms);
      return;
    }

    const message = testFailureMessage(result, runItem);
    if (runItem?.status === "inconclusive") {
      run.skipped(item);
      run.appendOutput(`${message.message}\r\n`, undefined, item);
      return;
    }
    run.failed(item, message, runItem?.duration_ms);
  }
}

async function discoverableWorkspaceFolders(
  xlflowWorkspacesOnly: boolean,
): Promise<readonly vscode.WorkspaceFolder[]> {
  const folders = vscode.workspace.workspaceFolders ?? [];
  if (!xlflowWorkspacesOnly) {
    return folders;
  }
  const result: vscode.WorkspaceFolder[] = [];
  for (const folder of folders) {
    if (await hasXlflowConfig(folder)) {
      result.push(folder);
    }
  }
  return result;
}

async function hasXlflowConfig(folder: vscode.WorkspaceFolder): Promise<boolean> {
  try {
    await vscode.workspace.fs.stat(vscode.Uri.joinPath(folder.uri, "xlflow.toml"));
    return true;
  } catch {
    return false;
  }
}

function listDiscoveredTests(env: XlflowEnvelope | undefined): XlflowDiscoveredTest[] {
  const tests = env?.tests;
  if (isTestListPayload(tests) && Array.isArray(tests.items)) {
    return tests.items;
  }
  return [];
}

function findRunResult(
  env: XlflowEnvelope | undefined,
  metadata: TestMetadata,
): XlflowTestRunItem | undefined {
  const tests = env?.tests;
  const items = Array.isArray(tests) ? tests : isTestRunPayload(tests) ? tests.items : undefined;
  return items?.find(
    (item) =>
      readNonEmpty(item.module) === metadata.module && readNonEmpty(item.name) === metadata.name,
  );
}

function testFailureMessage(
  result: { exitCode: number; stderr: string; json?: XlflowEnvelope },
  item: XlflowTestRunItem | undefined,
): vscode.TestMessage {
  const error = item?.error ?? result.json?.error;
  const code = readNonEmpty(error?.code);
  const detail =
    readNonEmpty(error?.message) ?? readNonEmpty(result.stderr) ?? "xlflow test failed";
  return new vscode.TestMessage(code === undefined ? detail : `${code}: ${detail}`);
}

function collectionItems(collection: vscode.TestItemCollection): vscode.TestItem[] {
  const items: vscode.TestItem[] = [];
  collection.forEach((item) => items.push(item));
  return items;
}

function workspaceItemId(folder: vscode.WorkspaceFolder): string {
  return `${folder.uri.toString()}::workspace`;
}

function sourceUri(
  folder: vscode.WorkspaceFolder,
  sourcePath: string | undefined,
): vscode.Uri | undefined {
  const path = readNonEmpty(sourcePath);
  if (path === undefined) {
    return undefined;
  }
  return vscode.Uri.joinPath(folder.uri, ...path.replace(/\\/g, "/").split("/"));
}

function readNonEmpty(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }
  const trimmed = value.trim();
  return trimmed.length === 0 ? undefined : trimmed;
}

function isTestListPayload(value: unknown): value is XlflowTestListPayload {
  return typeof value === "object" && value !== null && "items" in value;
}

function isTestRunPayload(value: unknown): value is XlflowTestRunPayload {
  return typeof value === "object" && value !== null && "items" in value;
}
