import * as vscode from "vscode";

export type TraceServer = "off" | "messages" | "verbose";

export interface XlflowConfig {
  path: string;
  lspEnabled: boolean;
  lspLogFile: string;
  lspLogFileConfigured: boolean;
  lspPerformanceLogging: boolean;
  lspTraceServer: TraceServer;
  codeLensEnabled: boolean;
  codeLensRunProcedure: boolean;
  codeLensRunTests: boolean;
  codeLensUserFormEvents: boolean;
  runSaveBeforeRun: boolean;
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
    lspLogFileConfigured: hasConfiguredValue(config, "lsp.logFile"),
    lspPerformanceLogging: config.get<boolean>("lsp.performanceLogging", false),
    lspTraceServer: readTraceServer(config),
    codeLensEnabled: config.get<boolean>("codeLens.enabled", true),
    codeLensRunProcedure: config.get<boolean>("codeLens.runProcedure", true),
    codeLensRunTests: config.get<boolean>("codeLens.runTests", true),
    codeLensUserFormEvents: config.get<boolean>("codeLens.userFormEvents", false),
    runSaveBeforeRun: config.get<boolean>("run.saveBeforeRun", true),
    completionTriggerSuggestInStatements: config.get<boolean>(
      "completion.triggerSuggestInStatements",
      true,
    ),
    completionProgIdsInStrings: config.get<boolean>("completion.progIdsInStrings", true),
    testingAutoDiscover: config.get<boolean>("testing.autoDiscover", true),
  };
}

function hasConfiguredValue(config: vscode.WorkspaceConfiguration, key: string): boolean {
  const inspected = config.inspect(key);
  return (
    inspected?.globalValue !== undefined ||
    inspected?.workspaceValue !== undefined ||
    inspected?.workspaceFolderValue !== undefined
  );
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
