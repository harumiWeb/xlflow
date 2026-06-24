import * as vscode from "vscode";
import {
  LanguageClient,
  LanguageClientOptions,
  ServerOptions,
  TransportKind,
} from "vscode-languageclient/node";
import { readConfig } from "./config";
import { XlflowChannels } from "./logging";
import { resolveWorkspaceRoot } from "./xlflow";

export class XlflowLanguageClientManager implements vscode.Disposable {
  private client: LanguageClient | undefined;
  private suggestTimer: NodeJS.Timeout | undefined;

  public constructor(private readonly channels: XlflowChannels) {}

  public async start(): Promise<void> {
    const config = readConfig();
    if (!config.lspEnabled) {
      this.channels.output.info("xlflow LSP is disabled by xlflow.lsp.enabled.");
      return;
    }
    if (this.client !== undefined) {
      return;
    }

    const folder = await resolveWorkspaceRoot({ prompt: false, fallbackToFirst: true });
    const cwd = folder?.uri.fsPath;
    const serverOptions: ServerOptions = {
      command: config.path,
      args: ["lsp", "--stdio", "--log-file", config.lspLogFile],
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
    };

    const client = new LanguageClient("xlflow-vscode", "xlflow", serverOptions, clientOptions);
    this.client = client;

    try {
      await client.start();
      this.channels.output.info(
        `Started xlflow lsp --stdio${cwd === undefined ? "" : ` in ${cwd}`} with log file ${config.lspLogFile}`,
      );
    } catch (error) {
      this.client = undefined;
      this.channels.output.error(`Failed to start xlflow lsp --stdio: ${String(error)}`);
      throw error;
    }
  }

  public async stop(): Promise<void> {
    const client = this.client;
    this.client = undefined;
    this.clearPendingSuggest();
    await client?.stop();
  }

  public async restart(): Promise<void> {
    this.channels.output.info("Restarting xlflow language server.");
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

function isStatementPrefix(linePrefix: string): boolean {
  const typed = linePrefix.trimStart();
  if (typed.length === 0 || /[."'():=]/.test(typed)) {
    return false;
  }
  return /^(o|op|opt|opti|optio|option|option\s+\w*|p|pu|pub|publ|publi|public|public\s+\w*|pr|pri|priv|priva|privat|private|private\s+\w*|f|fr|fri|frie|frien|friend|friend\s+\w*|s|su|sub|fu|fun|func|funct|functi|functio|function|d|di|dim|c|co|con|cons|const|t|ty|typ|type|e|en|enu|enum|declare|declare\s+\w*)$/i.test(
    typed,
  );
}

function isProgIdStringPrefix(linePrefix: string): boolean {
  return /\b(CreateObject|GetObject)\s*\(\s*"[^"]*$/i.test(linePrefix);
}
