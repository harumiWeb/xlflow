import * as path from "path";
import * as vscode from "vscode";
import { XlflowChannels } from "./logging";
import {
  readyWorkspaceFolder,
  XlflowProjectState,
  XlflowProjectStateService,
} from "./projectState";
import { SessionManager, SessionSnapshot, SessionState, XlflowSessionPayload } from "./session";
import { discoverTests, readNonEmpty, sourceUri, XlflowDiscoveredTest } from "./testDiscovery";
import { runXlflowJsonCommand } from "./xlflow";

type TreeNode =
  | SetupNode
  | ProjectNode
  | ModuleGroupNode
  | ModuleNode
  | ProcedureNode
  | TestCountNode
  | TestNode;

interface SetupNode {
  kind: "setup";
  label: string;
  description?: string;
  icon: vscode.ThemeIcon;
  command?: vscode.Command;
}

interface ProjectNode {
  kind: "project";
  label: string;
  description?: string;
  tooltip?: string;
  icon: vscode.ThemeIcon;
  command?: vscode.Command;
}

interface ModuleGroupNode {
  kind: "moduleGroup";
  label: string;
  children: ModuleNode[];
}

interface ModuleNode {
  kind: "module";
  name: string;
  moduleKind: string;
  uri: vscode.Uri;
  procedures: ProcedureNode[];
}

interface ProcedureNode {
  kind: "procedure";
  name: string;
  procedureKind: string;
  moduleName: string;
  qualifiedName: string;
  uri: vscode.Uri;
  line: number;
  runnable: boolean;
  test: boolean;
}

interface TestCountNode {
  kind: "testCount";
  label: string;
}

interface TestNode {
  kind: "test";
  test: XlflowDiscoveredTest;
  uri?: vscode.Uri;
}

interface InspectSymbolsEnvelope {
  status?: string;
  inspect?: {
    files?: InspectSymbolFile[];
  };
}

interface InspectSymbolFile {
  path?: string;
  moduleName?: string;
  moduleKind?: string;
  symbols?: InspectSymbol[];
}

interface InspectSymbol {
  name?: string;
  kind?: string;
  module?: string;
  file?: string;
  startLine?: number;
  parameters?: unknown[];
}

export class XlflowSidebar implements vscode.Disposable {
  private readonly setupProvider: SetupTreeProvider;
  private readonly projectProvider: ProjectTreeProvider;
  private readonly modulesProvider: ModulesTreeProvider;
  private readonly testsProvider: TestsTreeProvider;
  private readonly disposables: vscode.Disposable[] = [];

  constructor(
    private readonly projectState: XlflowProjectStateService,
    private readonly sessionManager: SessionManager,
    channels: XlflowChannels,
  ) {
    this.setupProvider = new SetupTreeProvider(projectState);
    this.projectProvider = new ProjectTreeProvider(projectState, sessionManager);
    this.modulesProvider = new ModulesTreeProvider(projectState, channels);
    this.testsProvider = new TestsTreeProvider(projectState, channels);

    this.disposables.push(
      vscode.window.registerTreeDataProvider("xlflow.setup", this.setupProvider),
      vscode.window.registerTreeDataProvider("xlflow.project", this.projectProvider),
      vscode.window.registerTreeDataProvider("xlflow.modules", this.modulesProvider),
      vscode.window.registerTreeDataProvider("xlflow.tests", this.testsProvider),
      projectState.onDidChangeState((state) => {
        sessionManager.setProjectKind(state.kind);
        this.refreshProjectViews();
        if (state.kind === "ready") {
          void this.refreshModules();
          void this.refreshTests();
        }
      }),
      sessionManager.onDidChangeSnapshot(() => this.projectProvider.refresh()),
    );
  }

  dispose(): void {
    for (const disposable of this.disposables) {
      disposable.dispose();
    }
  }

  refreshProjectViews(): void {
    this.setupProvider.refresh();
    this.projectProvider.refresh();
  }

  async refreshModules(): Promise<void> {
    await this.modulesProvider.refresh();
  }

  async refreshTests(): Promise<void> {
    await this.testsProvider.refresh();
  }

  async refreshAll(): Promise<void> {
    await this.projectState.refresh();
    this.refreshProjectViews();
    await Promise.all([this.refreshModules(), this.refreshTests()]);
  }
}

class SetupTreeProvider implements vscode.TreeDataProvider<SetupNode> {
  private readonly emitter = new vscode.EventEmitter<SetupNode | undefined>();
  readonly onDidChangeTreeData = this.emitter.event;

  constructor(private readonly projectState: XlflowProjectStateService) {}

  refresh(): void {
    this.emitter.fire(undefined);
  }

  getTreeItem(element: SetupNode): vscode.TreeItem {
    const item = new vscode.TreeItem(element.label, vscode.TreeItemCollapsibleState.None);
    item.description = element.description;
    item.iconPath = element.icon;
    item.command = element.command;
    item.contextValue = "xlflow.setup";
    return item;
  }

  getChildren(): SetupNode[] {
    const state = this.projectState.current();
    if (state.kind === "ready") {
      return [];
    }
    if (state.kind === "invalid") {
      return [
        {
          kind: "setup",
          label: "Project configuration error",
          description: state.error,
          icon: new vscode.ThemeIcon("warning"),
        },
        setupAction("Run Doctor", "stethoscope", "xlflow.runDoctor"),
        setupAction("Open Documentation", "book", "xlflow.openDocumentation"),
      ];
    }
    const status =
      state.kind === "noWorkspace"
        ? "Open a workspace folder to use xlflow."
        : "No xlflow project detected";
    return [
      { kind: "setup", label: status, icon: new vscode.ThemeIcon("info") },
      setupAction("New Project", "new-file", "xlflow.newProject"),
      setupAction("Init Existing Workbook", "file-add", "xlflow.initProject"),
      setupAction("Run Doctor", "stethoscope", "xlflow.runDoctor"),
      setupAction("Open Documentation", "book", "xlflow.openDocumentation"),
    ];
  }
}

class ProjectTreeProvider implements vscode.TreeDataProvider<ProjectNode> {
  private readonly emitter = new vscode.EventEmitter<ProjectNode | undefined>();
  readonly onDidChangeTreeData = this.emitter.event;

  constructor(
    private readonly projectState: XlflowProjectStateService,
    private readonly sessionManager: SessionManager,
  ) {}

  refresh(): void {
    this.emitter.fire(undefined);
  }

  getTreeItem(element: ProjectNode): vscode.TreeItem {
    const item = new vscode.TreeItem(element.label, vscode.TreeItemCollapsibleState.None);
    item.description = element.description;
    item.tooltip = element.tooltip;
    item.iconPath = element.icon;
    item.command = element.command;
    item.contextValue = "xlflow.projectItem";
    return item;
  }

  async getChildren(): Promise<ProjectNode[]> {
    const state = this.projectState.current();
    if (state.kind !== "ready") {
      return [];
    }
    const snapshot = this.sessionManager.currentSnapshot();
    const workbook = workbookDisplayName(snapshot.session) ?? (await configuredWorkbook(state));
    const nodes: ProjectNode[] = [
      {
        kind: "project",
        label: state.workspaceFolder.name,
        tooltip: state.workspaceFolder.uri.fsPath,
        icon: new vscode.ThemeIcon("folder"),
        command: {
          command: "revealFileInOS",
          title: "Reveal Workspace",
          arguments: [state.workspaceFolder.uri],
        },
      },
      {
        kind: "project",
        label: "Workbook",
        description: workbook ?? "Unknown",
        icon: new vscode.ThemeIcon("file-binary"),
        command:
          workbook === undefined
            ? undefined
            : {
                command: "vscode.open",
                title: "Open Workbook",
                arguments: [workbookUri(state.workspaceFolder, workbook)],
              },
      },
      {
        kind: "project",
        label: "Config",
        description: "xlflow.toml",
        icon: new vscode.ThemeIcon("settings-gear"),
        command: { command: "vscode.open", title: "Open Config", arguments: [state.configPath] },
      },
      {
        kind: "project",
        label: "Session",
        description: sessionDescription(snapshot.state),
        tooltip: snapshot.lastError,
        icon: sessionIcon(snapshot.state),
        command: { command: "xlflow.sessionActions", title: "Session Actions" },
      },
      {
        kind: "project",
        label: "Save required",
        description: String(snapshot.session?.save_required === true),
        tooltip: "Whether the managed session workbook has unsaved changes that should be saved.",
        icon: new vscode.ThemeIcon(snapshot.session?.save_required === true ? "warning" : "check"),
      },
    ];
    return nodes;
  }
}

class ModulesTreeProvider implements vscode.TreeDataProvider<TreeNode> {
  private readonly emitter = new vscode.EventEmitter<TreeNode | undefined>();
  readonly onDidChangeTreeData = this.emitter.event;
  private groups: ModuleGroupNode[] = [];

  constructor(
    private readonly projectState: XlflowProjectStateService,
    private readonly channels: XlflowChannels,
  ) {}

  async refresh(): Promise<void> {
    const folder = readyWorkspaceFolder(this.projectState.current());
    if (folder === undefined) {
      this.groups = [];
      this.emitter.fire(undefined);
      return;
    }
    const result = await runXlflowJsonCommand<InspectSymbolsEnvelope>(
      ["--json", "inspect", "symbols"],
      "xlflow inspect symbols",
      this.channels.output,
      { requireWorkspace: true, workspaceFolder: folder },
    );
    this.groups = result.exitCode === 0 ? moduleGroups(folder, result.json) : [];
    this.emitter.fire(undefined);
  }

  getTreeItem(element: TreeNode): vscode.TreeItem {
    switch (element.kind) {
      case "moduleGroup":
        return new vscode.TreeItem(element.label, vscode.TreeItemCollapsibleState.Expanded);
      case "module": {
        const item = new vscode.TreeItem(element.name, vscode.TreeItemCollapsibleState.Collapsed);
        item.iconPath = new vscode.ThemeIcon(moduleIcon(element.moduleKind));
        item.resourceUri = element.uri;
        item.contextValue = "xlflow.module";
        item.command = {
          command: "xlflow.openModule",
          title: "Open Module",
          arguments: [element],
        };
        return item;
      }
      case "procedure": {
        const item = new vscode.TreeItem(element.name, vscode.TreeItemCollapsibleState.None);
        item.description = element.procedureKind;
        item.iconPath = new vscode.ThemeIcon(element.test ? "beaker" : "symbol-method");
        item.contextValue = element.runnable
          ? element.test
            ? "xlflow.testProcedure"
            : "xlflow.procedure"
          : "xlflow.procedureStatic";
        item.command = {
          command: "xlflow.openProcedure",
          title: "Open Procedure",
          arguments: [element],
        };
        return item;
      }
      default:
        return new vscode.TreeItem("");
    }
  }

  getChildren(element?: TreeNode): TreeNode[] {
    if (element === undefined) {
      return this.groups;
    }
    if (element.kind === "moduleGroup") {
      return element.children;
    }
    if (element.kind === "module") {
      return element.procedures;
    }
    return [];
  }
}

class TestsTreeProvider implements vscode.TreeDataProvider<TreeNode> {
  private readonly emitter = new vscode.EventEmitter<TreeNode | undefined>();
  readonly onDidChangeTreeData = this.emitter.event;
  private nodes: TreeNode[] = [];

  constructor(
    private readonly projectState: XlflowProjectStateService,
    private readonly channels: XlflowChannels,
  ) {}

  async refresh(): Promise<void> {
    const folder = readyWorkspaceFolder(this.projectState.current());
    if (folder === undefined) {
      this.nodes = [];
      this.emitter.fire(undefined);
      return;
    }
    const result = await discoverTests(folder, this.channels);
    if (result.exitCode !== 0) {
      this.nodes = [];
      this.emitter.fire(undefined);
      return;
    }
    const tests = result.tests
      .filter(
        (test) => readNonEmpty(test.module) !== undefined && readNonEmpty(test.name) !== undefined,
      )
      .sort((a, b) => `${a.module}.${a.name}`.localeCompare(`${b.module}.${b.name}`));
    this.nodes = [
      { kind: "testCount", label: `${tests.length} tests` },
      ...tests.map((test) => ({
        kind: "test" as const,
        test,
        uri: sourceUri(folder, test.source_path),
      })),
    ];
    this.emitter.fire(undefined);
  }

  getTreeItem(element: TreeNode): vscode.TreeItem {
    if (element.kind === "testCount") {
      const item = new vscode.TreeItem(element.label, vscode.TreeItemCollapsibleState.None);
      item.iconPath = new vscode.ThemeIcon("beaker");
      return item;
    }
    if (element.kind === "test") {
      const name = readNonEmpty(element.test.name) ?? "Unknown test";
      const item = new vscode.TreeItem(name, vscode.TreeItemCollapsibleState.None);
      item.description = readNonEmpty(element.test.module);
      item.iconPath = new vscode.ThemeIcon("beaker");
      item.contextValue = element.uri === undefined ? "xlflow.test" : "xlflow.testWithSource";
      if (element.uri !== undefined) {
        item.command = {
          command: "xlflow.openProcedure",
          title: "Open Test",
          arguments: [testProcedureNode(element)],
        };
      }
      return item;
    }
    return new vscode.TreeItem("");
  }

  getChildren(element?: TreeNode): TreeNode[] {
    return element === undefined ? this.nodes : [];
  }
}

export function readExcelPathFromToml(text: string): string | undefined {
  let inExcel = false;
  for (const rawLine of text.split(/\r?\n/)) {
    const line = rawLine.replace(/#.*/, "").trim();
    if (line === "") {
      continue;
    }
    const section = line.match(/^\[([^\]]+)\]$/);
    if (section !== null) {
      inExcel = section[1].trim() === "excel";
      continue;
    }
    if (!inExcel) {
      continue;
    }
    const match = line.match(/^path\s*=\s*"([^"]+)"\s*$/);
    if (match !== null && match[1].trim() !== "") {
      return match[1].trim();
    }
  }
  return undefined;
}

function setupAction(label: string, icon: string, command: string): SetupNode {
  return {
    kind: "setup",
    label,
    icon: new vscode.ThemeIcon(icon),
    command: { command, title: label },
  };
}

async function configuredWorkbook(state: Extract<XlflowProjectState, { kind: "ready" }>) {
  try {
    const bytes = await vscode.workspace.fs.readFile(state.configPath);
    return readExcelPathFromToml(Buffer.from(bytes).toString("utf8"));
  } catch {
    return undefined;
  }
}

function workbookDisplayName(session: XlflowSessionPayload | undefined): string | undefined {
  const explicitName = readNonEmpty(session?.workbook_name);
  if (explicitName !== undefined) {
    return explicitName;
  }
  const workbookPath = readNonEmpty(session?.workbook_path ?? session?.metadata?.workbook_path);
  return workbookPath === undefined ? undefined : path.basename(workbookPath);
}

function workbookUri(folder: vscode.WorkspaceFolder, workbook: string): vscode.Uri {
  if (path.isAbsolute(workbook)) {
    return vscode.Uri.file(workbook);
  }
  return vscode.Uri.joinPath(folder.uri, ...workbook.replace(/\\/g, "/").split("/"));
}

function sessionDescription(state: SessionState): string {
  switch (state) {
    case "active":
      return "Active";
    case "inactive":
      return "Inactive";
    case "error":
      return "Error";
    case "starting":
      return "Starting";
    case "stopping":
      return "Stopping";
    case "unknown":
      return "Unknown";
  }
}

function sessionIcon(state: SessionState): vscode.ThemeIcon {
  switch (state) {
    case "active":
      return new vscode.ThemeIcon("circle-filled", new vscode.ThemeColor("testing.iconPassed"));
    case "error":
      return new vscode.ThemeIcon("warning");
    default:
      return new vscode.ThemeIcon("circle-outline");
  }
}

export function moduleGroups(
  folder: vscode.WorkspaceFolder,
  envelope: InspectSymbolsEnvelope | undefined,
): ModuleGroupNode[] {
  const files = envelope?.inspect?.files ?? [];
  const byKind = new Map<string, ModuleNode[]>();
  for (const file of files) {
    const name = readNonEmpty(file.moduleName);
    const sourcePath = readNonEmpty(file.path);
    if (name === undefined || sourcePath === undefined) {
      continue;
    }
    const moduleKind = normalizeModuleKind(file.moduleKind);
    const uri = sourceUri(folder, sourcePath) ?? vscode.Uri.joinPath(folder.uri, sourcePath);
    const module: ModuleNode = {
      kind: "module",
      name,
      moduleKind,
      uri,
      procedures: procedureNodes(file, name, uri),
    };
    const modules = byKind.get(moduleKind) ?? [];
    modules.push(module);
    byKind.set(moduleKind, modules);
  }

  return ["standard", "class", "document", "form"]
    .map((kind) => {
      const modules = (byKind.get(kind) ?? []).sort((a, b) => a.name.localeCompare(b.name));
      return { kind: "moduleGroup" as const, label: moduleGroupLabel(kind), children: modules };
    })
    .filter((group) => group.children.length > 0);
}

function procedureNodes(
  file: InspectSymbolFile,
  moduleName: string,
  uri: vscode.Uri,
): ProcedureNode[] {
  return (file.symbols ?? [])
    .filter((symbol) => isProcedureKind(symbol.kind))
    .map((symbol) => {
      const name = readNonEmpty(symbol.name) ?? "Unknown";
      const line =
        typeof symbol.startLine === "number" && symbol.startLine > 0 ? symbol.startLine : 1;
      const parameters = Array.isArray(symbol.parameters) ? symbol.parameters : [];
      const test = /^test/i.test(name) || /_test$/i.test(name);
      return {
        kind: "procedure" as const,
        name,
        procedureKind: readNonEmpty(symbol.kind) ?? "procedure",
        moduleName,
        qualifiedName: `${moduleName}.${name}`,
        uri,
        line,
        runnable: readNonEmpty(symbol.kind) === "sub" && parameters.length === 0,
        test,
      };
    })
    .sort((a, b) => a.line - b.line || a.name.localeCompare(b.name));
}

function testProcedureNode(node: TestNode): ProcedureNode {
  const moduleName = readNonEmpty(node.test.module) ?? "";
  const name = readNonEmpty(node.test.name) ?? "";
  const line = typeof node.test.line === "number" && node.test.line > 0 ? node.test.line : 1;
  return {
    kind: "procedure",
    name,
    procedureKind: "test",
    moduleName,
    qualifiedName: readNonEmpty(node.test.qualified_name) ?? `${moduleName}.${name}`,
    uri: node.uri ?? vscode.Uri.file(""),
    line,
    runnable: true,
    test: true,
  };
}

function normalizeModuleKind(value: unknown): string {
  const kind = readNonEmpty(value);
  switch (kind) {
    case "class":
    case "document":
    case "form":
    case "standard":
      return kind;
    default:
      return "standard";
  }
}

function moduleGroupLabel(kind: string): string {
  switch (kind) {
    case "class":
      return "Class Modules";
    case "document":
      return "Document Modules";
    case "form":
      return "UserForm Modules";
    default:
      return "Standard Modules";
  }
}

function moduleIcon(kind: string): string {
  switch (kind) {
    case "class":
      return "symbol-class";
    case "document":
      return "file-code";
    case "form":
      return "window";
    default:
      return "symbol-module";
  }
}

function isProcedureKind(kind: unknown): boolean {
  switch (kind) {
    case "sub":
    case "function":
    case "property":
    case "property_get":
    case "property_let":
    case "property_set":
    case "declare":
    case "declare_sub":
    case "declare_function":
      return true;
    default:
      return false;
  }
}
