import * as vscode from "vscode";

export type TraceServer = "off" | "messages" | "verbose";

export interface XlflowConfig {
  path: string;
  lspEnabled: boolean;
  lspLogFile: string;
  lspTraceServer: TraceServer;
  completionTriggerSuggestInStatements: boolean;
  completionProgIdsInStrings: boolean;
  testingAutoDiscover: boolean;
}

export function readConfig(): XlflowConfig {
  const config = vscode.workspace.getConfiguration("xlflow");
  return {
    path: readString(config, "path", "xlflow"),
    lspEnabled: config.get<boolean>("lsp.enabled", true),
    lspLogFile: readString(config, "lsp.logFile", ".xlflow/lsp.log"),
    lspTraceServer: readTraceServer(config),
    completionTriggerSuggestInStatements: config.get<boolean>(
      "completion.triggerSuggestInStatements",
      true,
    ),
    completionProgIdsInStrings: config.get<boolean>("completion.progIdsInStrings", true),
    testingAutoDiscover: config.get<boolean>("testing.autoDiscover", true),
  };
}

function readString(config: vscode.WorkspaceConfiguration, key: string, fallback: string): string {
  const value = config.get<string>(key, fallback).trim();
  return value.length === 0 ? fallback : value;
}

function readTraceServer(config: vscode.WorkspaceConfiguration): TraceServer {
  const value = config.get<string>("lsp.trace.server", "messages");
  if (value === "off" || value === "messages" || value === "verbose") {
    return value;
  }
  return "messages";
}
