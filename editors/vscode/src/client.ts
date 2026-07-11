import * as vscode from "vscode";
import {
  LanguageClient,
  LanguageClientOptions,
  ServerOptions,
  Trace,
  TransportKind,
} from "vscode-languageclient/node";
import { XlflowCliAvailabilityService } from "./cliAvailability";
import { readConfig, TraceServer, XlflowConfig } from "./config";
import { XlflowChannels } from "./logging";
import { resolveWorkspaceRoot } from "./xlflow";

export class XlflowLanguageClientManager implements vscode.Disposable {
  private client: LanguageClient | undefined;
  private workspaceFolderKey: string | undefined;
  private suggestTimer: NodeJS.Timeout | undefined;

  public constructor(
    private readonly channels: XlflowChannels,
    private readonly cliAvailability: XlflowCliAvailabilityService,
  ) {}

  public async start(): Promise<void> {
    const config = readConfig();
    if (!config.lspEnabled) {
      this.channels.output.info("xlflow LSP is disabled by xlflow.lsp.enabled.");
      return;
    }
    if (this.client !== undefined) {
      return;
    }
    const availability = await this.cliAvailability.refresh();
    if (!availability.ok) {
      this.channels.output.info(
        `Skipping xlflow lsp --stdio startup because ${availability.executable} is unavailable: ${availability.message}`,
      );
      return;
    }

    const folder = await resolveWorkspaceRoot({ prompt: false, fallbackToFirst: true });
    const cwd = folder?.uri.fsPath;
    const workspaceFolderKey = folder?.uri.toString();
    const xlflowProject = await hasXlflowConfig(folder);
    const args = lspServerArgsForProject(config, xlflowProject);
    const codeLens = lspCodeLensOptions(config, xlflowProject);
    const serverOptions: ServerOptions = {
      command: config.path,
      args,
      transport: TransportKind.stdio,
      options: cwd === undefined ? undefined : { cwd },
    };
    const clientOptions: LanguageClientOptions = {
      documentSelector: [
        { scheme: "file", language: "vba" },
        { scheme: "file", pattern: "**/*.{bas,cls,frm}" },
      ],
      synchronize: {
        fileEvents: vscode.workspace.createFileSystemWatcher("**/*.{bas,cls,frm}"),
      },
      outputChannel: this.channels.output,
      traceOutputChannel: this.channels.trace,
      initializationOptions: {
        codeLens,
      },
    };

    const client = new LanguageClient("xlflow-vscode", "xlflow", serverOptions, clientOptions);
    this.client = client;

    try {
      await client.start();
      this.workspaceFolderKey = workspaceFolderKey;
      await client.setTrace(toProtocolTrace(config.lspTraceServer));
      const logDescription = args.includes("--log-file")
        ? ` with log file ${config.lspLogFile}`
        : " without workspace log file";
      this.channels.output.info(
        `Started xlflow lsp --stdio${cwd === undefined ? "" : ` in ${cwd}`}${logDescription}`,
      );
    } catch (error) {
      this.client = undefined;
      this.workspaceFolderKey = undefined;
      this.channels.output.error(`Failed to start xlflow lsp --stdio: ${String(error)}`);
      throw error;
    }
  }

  public async stop(): Promise<void> {
    const client = this.client;
    this.client = undefined;
    this.workspaceFolderKey = undefined;
    this.clearPendingSuggest();
    await client?.stop();
  }

  public async restart(): Promise<void> {
    this.channels.output.info("Restarting xlflow language server.");
    await this.stop();
    await this.start();
  }

  public async restartIfWorkspaceChanged(): Promise<void> {
    if (this.client === undefined) {
      return;
    }
    const folder = await resolveWorkspaceRoot({ prompt: false, fallbackToFirst: true });
    const nextWorkspaceFolderKey = folder?.uri.toString();
    if (nextWorkspaceFolderKey === this.workspaceFolderKey) {
      return;
    }
    this.channels.output.info("Restarting xlflow language server for selected workspace change.");
    await this.stop();
    await this.start();
  }

  public scheduleSuggest(document: vscode.TextDocument): void {
    const config = readConfig();
    if (!config.completionTriggerSuggestInStatements && !config.completionProgIdsInStrings) {
      return;
    }

    const editor = vscode.window.activeTextEditor;
    if (editor === undefined || editor.document !== document || document.languageId !== "vba") {
      return;
    }

    const position = editor.selection.active;
    const linePrefix = document.lineAt(position.line).text.slice(0, position.character);
    if (config.completionTriggerSuggestInStatements && isDocCommentSnippetPrefix(linePrefix)) {
      this.clearPendingSuggest();
      this.suggestTimer = setTimeout(() => {
        this.suggestTimer = undefined;
        void vscode.commands.executeCommand("editor.action.quickFix");
      }, 75);
      return;
    }

    if (
      (config.completionTriggerSuggestInStatements && isStatementPrefix(linePrefix)) ||
      (config.completionProgIdsInStrings && isProgIdStringPrefix(linePrefix))
    ) {
      this.clearPendingSuggest();
      this.suggestTimer = setTimeout(() => {
        this.suggestTimer = undefined;
        void vscode.commands.executeCommand("editor.action.triggerSuggest");
      }, 75);
    }
  }

  public dispose(): void {
    this.clearPendingSuggest();
    void this.stop();
  }

  private clearPendingSuggest(): void {
    if (this.suggestTimer !== undefined) {
      clearTimeout(this.suggestTimer);
      this.suggestTimer = undefined;
    }
  }
}

export async function lspServerArgs(
  config: Pick<XlflowConfig, "lspLogFile" | "lspLogFileConfigured">,
  folder: vscode.WorkspaceFolder | undefined,
): Promise<string[]> {
  return lspServerArgsForProject(config, await hasXlflowConfig(folder));
}

export function lspServerArgsForProject(
  config: Pick<XlflowConfig, "lspLogFile" | "lspLogFileConfigured">,
  xlflowProject: boolean,
): string[] {
  const args = ["lsp", "--stdio"];
  if (config.lspLogFileConfigured || xlflowProject) {
    args.push("--log-file", config.lspLogFile);
  }
  return args;
}

export interface LSPCodeLensOptions {
  enabled: boolean;
  runProcedure: boolean;
  runTests: boolean;
  userFormEvents: boolean;
}

export function lspCodeLensOptions(
  config: Pick<
    XlflowConfig,
    "codeLensEnabled" | "codeLensRunProcedure" | "codeLensRunTests" | "codeLensUserFormEvents"
  >,
  xlflowProject: boolean,
): LSPCodeLensOptions {
  return {
    enabled: xlflowProject && config.codeLensEnabled,
    runProcedure: config.codeLensRunProcedure,
    runTests: config.codeLensRunTests,
    userFormEvents: config.codeLensUserFormEvents,
  };
}

async function hasXlflowConfig(folder: vscode.WorkspaceFolder | undefined): Promise<boolean> {
  if (folder === undefined) {
    return false;
  }
  try {
    const stat = await vscode.workspace.fs.stat(vscode.Uri.joinPath(folder.uri, "xlflow.toml"));
    return (stat.type & vscode.FileType.File) !== 0;
  } catch {
    return false;
  }
}

function toProtocolTrace(trace: TraceServer): Trace {
  switch (trace) {
    case "off":
      return Trace.Off;
    case "verbose":
      return Trace.Verbose;
    case "messages":
      return Trace.Messages;
  }
}

export function isStatementPrefix(linePrefix: string): boolean {
  const typed = linePrefix.trimStart();
  if (typed.length === 0 || /[."'():=]/.test(typed)) {
    return false;
  }
  return /^(o|op|opt|opti|optio|option|option\s+\w*|p|pu|pub|publ|publi|public|public\s+\w*|pr|pri|priv|priva|privat|private|private\s+\w*|f|fr|fri|frie|frien|friend|friend\s+\w*|s|su|sub|fu|fun|func|funct|functi|functio|function|d|di|dim|c|co|con|cons|const|t|ty|typ|type|e|en|enu|enum|declare|declare\s+\w*)$/i.test(
    typed,
  );
}

export function isDocCommentSnippetPrefix(linePrefix: string): boolean {
  return /^\s*'''$/.test(linePrefix);
}

export function isProgIdStringPrefix(linePrefix: string): boolean {
  return /\b(CreateObject|GetObject)\s*\(\s*"[^"]*$/i.test(linePrefix);
}
