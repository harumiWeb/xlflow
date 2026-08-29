import * as vscode from "vscode";
import {
  LanguageClient,
  LanguageClientOptions,
  ServerOptions,
  Trace,
  TransportKind,
} from "vscode-languageclient/node";
import type { XlflowCliAvailabilityService } from "./cliAvailability";
import { readConfig, TraceServer, XlflowConfig } from "./config";
import { XlflowChannels } from "./logging";
import { readFormsRootFromToml } from "./sidebar";
import {
  hasCompletionResult,
  hasDefinitionResult,
  hasHoverResult,
  completionResultCount,
  StartupTelemetry,
} from "./startupTelemetry";
import { resolveWorkspaceRoot } from "./xlflow";

class StartupLanguageClient extends LanguageClient {
  public constructor(
    id: string,
    name: string,
    serverOptions: ServerOptions,
    clientOptions: LanguageClientOptions,
    private readonly onInitialized: () => void,
  ) {
    super(id, name, serverOptions, clientOptions);
  }

  public override async start(): Promise<void> {
    await super.start();
    this.onInitialized();
  }
}

export function userFormSpecLSPGlob(formsRoot = "src/forms"): string {
  const normalizedRoot = formsRoot.replace(/\\/g, "/").replace(/^\/+|\/+$/g, "");
  return `**/${normalizedRoot || "src/forms"}/specs/*.{yaml,yml,json}`;
}

export class XlflowLanguageClientManager implements vscode.Disposable {
  private client: LanguageClient | undefined;
  private workspaceFolderKey: string | undefined;
  private suggestTimer: NodeJS.Timeout | undefined;
  private stateSubscription: vscode.Disposable | undefined;
  private hasStarted = false;
  private startup: StartupTelemetry | undefined;

  public constructor(
    private readonly channels: XlflowChannels,
    cliAvailabilityOrStartup?: XlflowCliAvailabilityService | StartupTelemetry,
    startup?: StartupTelemetry,
  ) {
    // Keep the pre-#757 constructor shape source-compatible for callers that
    // still pass the availability service. The service is no longer a startup
    // dependency because process launch is the availability check.
    this.startup =
      startup ??
      (cliAvailabilityOrStartup instanceof StartupTelemetry ? cliAvailabilityOrStartup : undefined);
  }

  public async start(): Promise<void> {
    const config = readConfig();
    if (!config.lspEnabled) {
      this.hasStarted = true;
      this.channels.output.info("xlflow LSP is disabled by xlflow.lsp.enabled.");
      return;
    }
    if (this.client !== undefined) {
      return;
    }
    if (this.startup !== undefined) {
      if (this.startup.isEnabled !== config.lspPerformanceLogging) {
        this.startup = this.startup.withEnabled(config.lspPerformanceLogging);
      } else if (this.hasStarted) {
        this.startup = this.startup.newAttempt();
      }
    }
    this.hasStarted = true;

    this.startup?.mark("workspaceResolutionStart");

    const folder = await resolveWorkspaceRoot({ prompt: false, fallbackToFirst: true });
    const cwd = folder?.uri.fsPath;
    const workspaceFolderKey = folder?.uri.toString();
    this.startup?.mark("workspaceResolutionComplete");
    this.startup?.mark("projectConfigDiscoveryStart");
    const [xlflowProject, formSpecGlob] = await Promise.all([
      hasXlflowConfig(folder),
      userFormSpecGlobForWorkspace(folder),
    ]);
    this.startup?.mark("projectConfigDiscoveryComplete");
    const args = lspServerArgsForProject(config, xlflowProject);
    const codeLens = lspCodeLensOptions(config, xlflowProject);
    const startupEnv =
      this.startup?.isEnabled === true
        ? { ...process.env, XLFLOW_LSP_STARTUP_ID: this.startup.id }
        : undefined;
    const serverOptions: ServerOptions = {
      command: config.path,
      args,
      transport: TransportKind.stdio,
      options:
        cwd === undefined && startupEnv === undefined
          ? undefined
          : { ...(cwd === undefined ? {} : { cwd }), env: startupEnv },
    };
    let telemetry = this.startup;
    const clientOptions: LanguageClientOptions = {
      documentSelector: [
        { scheme: "file", language: "vba" },
        { scheme: "file", pattern: "**/*.{bas,cls,frm}" },
        { scheme: "file", pattern: formSpecGlob },
      ],
      synchronize: {
        fileEvents: [
          vscode.workspace.createFileSystemWatcher("**/*.{bas,cls,frm}"),
          vscode.workspace.createFileSystemWatcher(formSpecGlob),
        ],
      },
      outputChannel: this.channels.output,
      traceOutputChannel: this.channels.trace,
      initializationOptions: {
        codeLens,
        declarationPriority: declarationPriorityInitializationOptions(),
      },
      middleware: {
        didOpen: (document, next) => {
          const requestTelemetry = telemetry;
          return next(document).then(() => {
            requestTelemetry?.mark("firstDidOpenSent");
          });
        },
        provideHover: async (document, position, token, next) => {
          const requestTelemetry = telemetry;
          const result = await next(document, position, token);
          if (!token.isCancellationRequested && hasHoverResult(result)) {
            requestTelemetry?.mark("firstHoverHandled", { resultCount: 1 });
          }
          return result;
        },
        provideDefinition: async (document, position, token, next) => {
          const requestTelemetry = telemetry;
          const result = await next(document, position, token);
          if (!token.isCancellationRequested && hasDefinitionResult(result)) {
            const resultCount = Array.isArray(result) ? result.length : 1;
            requestTelemetry?.mark("firstDefinitionHandled", { resultCount });
          }
          return result;
        },
        provideCompletionItem: async (document, position, context, token, next) => {
          const requestTelemetry = telemetry;
          const result = await next(document, position, context, token);
          if (!token.isCancellationRequested && hasCompletionResult(result)) {
            requestTelemetry?.mark("firstCompletionHandled", {
              resultCount: completionResultCount(result),
            });
          }
          return result;
        },
      },
    };

    const client = new StartupLanguageClient(
      "xlflow-vscode",
      "xlflow",
      serverOptions,
      clientOptions,
      () => telemetry?.mark("initializedSent"),
    );
    this.client = client;
    let processStartObserved = false;
    const stateSubscription = client.onDidChangeState((event) => {
      // State.Starting is the closest public boundary to child-process spawn;
      // the language-client package does not expose the enum from its node entrypoint.
      if (event.newState === 3) {
        const restarting = processStartObserved;
        if (restarting && telemetry?.isEnabled === true) {
          telemetry = telemetry.newAttempt();
          this.startup = telemetry;
          if (startupEnv !== undefined) {
            startupEnv.XLFLOW_LSP_STARTUP_ID = telemetry.id;
          }
          telemetry.mark("languageClientStart");
        }
        processStartObserved = true;
        telemetry?.mark("serverProcessSpawned");
        if (restarting) {
          telemetry?.mark("initializeRequestSent");
        }
      } else if (event.newState === 2) {
        telemetry?.mark("initializeResponseReceived");
      } else if (event.newState === 4) {
        telemetry?.mark("languageClientStartFailed", { outcome: "error" });
      }
    });
    this.stateSubscription = stateSubscription;

    const startAttemptTelemetry = telemetry;
    try {
      startAttemptTelemetry?.mark("languageClientStart");
      // start() publishes State.Starting before its first asynchronous
      // connection step, so the process boundary is observed before the
      // initialize request is recorded below.
      const startPromise = client.start();
      startAttemptTelemetry?.mark("initializeRequestSent");
      await startPromise;
      this.workspaceFolderKey = workspaceFolderKey;
      await client.setTrace(toProtocolTrace(config.lspTraceServer));
      this.notifyActiveDocument(vscode.window.activeTextEditor?.document);
      const logDescription = args.includes("--log-file")
        ? ` with log file ${config.lspLogFile}`
        : " without workspace log file";
      this.channels.output.info(
        `Started xlflow lsp --stdio${cwd === undefined ? "" : ` in ${cwd}`}${logDescription}`,
      );
    } catch (error) {
      if (this.client === client) {
        this.client = undefined;
        this.workspaceFolderKey = undefined;
      }
      stateSubscription.dispose();
      if (this.stateSubscription === stateSubscription) {
        this.stateSubscription = undefined;
      }
      startAttemptTelemetry?.mark("languageClientStartFailed", { outcome: "error" });
      this.channels.output.error(`Failed to start xlflow lsp --stdio: ${String(error)}`);
      throw error;
    }
  }

  public async stop(): Promise<void> {
    const client = this.client;
    this.client = undefined;
    this.workspaceFolderKey = undefined;
    this.stateSubscription?.dispose();
    this.stateSubscription = undefined;
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

  public notifyActiveDocument(document: vscode.TextDocument | undefined): void {
    if (this.client === undefined) {
      return;
    }
    void this.client
      .sendNotification("xlflow/didChangeActiveDocument", {
        uri: document?.languageId === "vba" ? document.uri.toString() : null,
      })
      .catch(() => undefined);
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
      (config.completionTriggerSuggestInStatements && isAnnotationCommentPrefix(linePrefix)) ||
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

export function declarationPriorityInitializationOptions(): {
  activeDocumentUri: string | null;
  openDocumentUris: string[];
} {
  const active = vscode.window.activeTextEditor?.document;
  const openDocumentUris = vscode.workspace.textDocuments
    .filter((document) => document.languageId === "vba")
    .map((document) => document.uri.toString());
  return {
    activeDocumentUri: active?.languageId === "vba" ? active.uri.toString() : null,
    openDocumentUris: [...new Set(openDocumentUris)],
  };
}

export async function lspServerArgs(
  config: Pick<XlflowConfig, "lspLogFile" | "lspLogFileConfigured" | "lspPerformanceLogging">,
  folder: vscode.WorkspaceFolder | undefined,
): Promise<string[]> {
  return lspServerArgsForProject(config, await hasXlflowConfig(folder));
}

export function lspServerArgsForProject(
  config: Pick<XlflowConfig, "lspLogFile" | "lspLogFileConfigured" | "lspPerformanceLogging">,
  xlflowProject: boolean,
): string[] {
  const args = ["lsp", "--stdio"];
  if (config.lspLogFileConfigured || xlflowProject) {
    args.push("--log-file", config.lspLogFile);
  }
  if (config.lspPerformanceLogging) {
    args.push("--performance-log");
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

async function userFormSpecGlobForWorkspace(
  folder: vscode.WorkspaceFolder | undefined,
): Promise<string> {
  if (folder === undefined) {
    return userFormSpecLSPGlob();
  }
  try {
    const configUri = vscode.Uri.joinPath(folder.uri, "xlflow.toml");
    const bytes = await vscode.workspace.fs.readFile(configUri);
    return userFormSpecLSPGlob(readFormsRootFromToml(Buffer.from(bytes).toString("utf8")));
  } catch {
    return userFormSpecLSPGlob();
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
  return /^(o|op|opt|opti|optio|option|option\s+\w*|p|pu|pub|publ|publi|public|public\s+\w*|pr|pri|priv|priva|privat|private|private\s+\w*|f|fr|fri|frie|frien|friend|friend\s+\w*|s|su|sub|fu|fun|func|funct|functi|functio|function|d|di|dim|dim\s+\w*|c|co|con|cons|const|t|ty|typ|type|e|en|enu|enum|declare|declare\s+\w*)$/i.test(
    typed,
  );
}

export function isDocCommentSnippetPrefix(linePrefix: string): boolean {
  return /^\s*'''$/.test(linePrefix);
}

export function isAnnotationCommentPrefix(linePrefix: string): boolean {
  return /^\s*'\s*@\w*$/.test(linePrefix);
}

export function isProgIdStringPrefix(linePrefix: string): boolean {
  return /\b(CreateObject|GetObject)\s*\(\s*"[^"]*$/i.test(linePrefix);
}
