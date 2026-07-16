import * as vscode from "vscode";
import type { XlflowChannels } from "./logging";
import { runXlflowJsonCommand } from "./xlflow";

export const supportedCapabilityVersion = 1;

export interface XlflowCommandCapability {
  cli_paths: string[];
  resource_scope: string;
  operation_kind: string;
  parallel_safe: boolean;
  retryable_when_busy: boolean;
  default_wait_policy: string;
  recovery_behavior: string;
}

export interface XlflowCapabilities {
  capability_version: number;
  commands: Record<string, XlflowCommandCapability>;
}

interface XlflowCapabilitiesEnvelope {
  status?: string;
  capabilities?: unknown;
}

export interface XlflowBusyOperation {
  busy: boolean;
  command?: string;
  operationKind?: string;
  resourceScope?: string;
  pid?: number;
  startedAt?: string;
}

export interface XlflowCapabilityOperation {
  commandID: string;
  capability: XlflowCommandCapability;
}

export interface XlflowCapabilityHooks {
  currentBusyOperation(): XlflowBusyOperation | undefined;
  operationStarted(operation: XlflowCapabilityOperation): void;
  operationFinished(): void;
  refreshStatus(): Promise<void>;
}

export class XlflowCapabilitiesService {
  private loadPromise: Promise<XlflowCapabilities | undefined> | undefined;

  constructor(
    private readonly channels: XlflowChannels,
    private readonly hooks: XlflowCapabilityHooks,
  ) {}

  async load(): Promise<XlflowCapabilities | undefined> {
    if (this.loadPromise === undefined) {
      this.loadPromise = this.loadOnce();
    }
    return this.loadPromise;
  }

  async beforeManagedCommand(
    args: string[],
  ): Promise<XlflowCapabilityOperation | undefined | "blocked"> {
    const operation = await this.operationForArgs(args);
    if (operation === undefined || !isWorkbookExclusive(operation.capability)) {
      return operation;
    }
    if (!args.includes("--wait")) {
      const busy = this.hooks.currentBusyOperation();
      if (busy?.busy === true) {
        showBusyAdvisory(operation, busy);
        return "blocked";
      }
    }
    this.hooks.operationStarted(operation);
    return operation;
  }

  async afterManagedCommand(operation: XlflowCapabilityOperation | undefined): Promise<void> {
    if (operation === undefined || !isWorkbookExclusive(operation.capability)) {
      return;
    }
    this.hooks.operationFinished();
    await this.hooks.refreshStatus();
  }

  async beforeTerminalCommand(args: string[]): Promise<boolean> {
    const operation = await this.operationForArgs(args);
    if (
      operation === undefined ||
      !isWorkbookExclusive(operation.capability) ||
      args.includes("--wait")
    ) {
      return true;
    }
    const busy = this.hooks.currentBusyOperation();
    if (busy?.busy !== true) {
      return true;
    }
    showBusyAdvisory(operation, busy);
    return false;
  }

  async capabilityForArgs(args: string[]): Promise<XlflowCommandCapability | undefined> {
    return (await this.operationForArgs(args))?.capability;
  }

  private async loadOnce(): Promise<XlflowCapabilities | undefined> {
    const result = await runXlflowJsonCommand<XlflowCapabilitiesEnvelope>(
      ["--json", "capabilities"],
      "xlflow capabilities",
      this.channels.output,
      { requireWorkspace: false, showCliUnavailable: false, skipCoordination: true },
    );
    return result.exitCode === 0 ? parseCapabilitiesEnvelope(result.json) : undefined;
  }

  private async operationForArgs(args: string[]): Promise<XlflowCapabilityOperation | undefined> {
    const capabilities = await this.load();
    if (capabilities === undefined) {
      return undefined;
    }
    return capabilityOperationForArgs(capabilities, args);
  }
}

export function parseCapabilitiesEnvelope(value: unknown): XlflowCapabilities | undefined {
  if (!isRecord(value) || value.status !== "ok" || !isRecord(value.capabilities)) {
    return undefined;
  }
  const raw = value.capabilities;
  if (raw.capability_version !== supportedCapabilityVersion || !isRecord(raw.commands)) {
    return undefined;
  }
  const commands: Record<string, XlflowCommandCapability> = {};
  for (const [id, descriptor] of Object.entries(raw.commands)) {
    const parsed = parseCommandCapability(descriptor);
    if (parsed !== undefined) {
      commands[id] = parsed;
    }
  }
  return { capability_version: supportedCapabilityVersion, commands };
}

export function capabilityOperationForArgs(
  capabilities: XlflowCapabilities,
  args: string[],
): XlflowCapabilityOperation | undefined {
  let best: XlflowCapabilityOperation | undefined;
  for (const [commandID, capability] of Object.entries(capabilities.commands)) {
    if (!containsCliPath(args, capability.cli_paths)) {
      continue;
    }
    if (
      best === undefined ||
      longestCliPathLength(capability.cli_paths) > longestCliPathLength(best.capability.cli_paths)
    ) {
      best = { commandID, capability };
    }
  }
  return best;
}

export function isWorkbookExclusive(capability: XlflowCommandCapability): boolean {
  return capability.resource_scope === "workbook" && capability.parallel_safe === false;
}

function parseCommandCapability(value: unknown): XlflowCommandCapability | undefined {
  if (
    !isRecord(value) ||
    !Array.isArray(value.cli_paths) ||
    !value.cli_paths.every(isNonEmptyString)
  ) {
    return undefined;
  }
  const resourceScope = value.resource_scope;
  const operationKind = value.operation_kind;
  const defaultWaitPolicy = value.default_wait_policy;
  const recoveryBehavior = value.recovery_behavior;
  if (
    !isNonEmptyString(resourceScope) ||
    !isNonEmptyString(operationKind) ||
    !isNonEmptyString(defaultWaitPolicy) ||
    !isNonEmptyString(recoveryBehavior)
  ) {
    return undefined;
  }
  if (typeof value.parallel_safe !== "boolean" || typeof value.retryable_when_busy !== "boolean") {
    return undefined;
  }
  return {
    cli_paths: value.cli_paths,
    resource_scope: resourceScope,
    operation_kind: operationKind,
    parallel_safe: value.parallel_safe,
    retryable_when_busy: value.retryable_when_busy,
    default_wait_policy: defaultWaitPolicy,
    recovery_behavior: recoveryBehavior,
  };
}

function containsCliPath(args: string[], cliPaths: string[]): boolean {
  for (const path of cliPaths) {
    const cliPath = cliPathParts(path);
    for (let offset = 0; offset <= args.length - cliPath.length; offset += 1) {
      if (cliPath.every((part, index) => args[offset + index] === part)) {
        return true;
      }
    }
  }
  return false;
}

function cliPathParts(path: string): string[] {
  // The public contract keeps selectors as complete CLI-path strings, e.g. "form build".
  return path.trim().split(/\s+/).filter(Boolean);
}

function longestCliPathLength(paths: string[]): number {
  return Math.max(0, ...paths.map((path) => cliPathParts(path).length));
}

function showBusyAdvisory(operation: XlflowCapabilityOperation, busy: XlflowBusyOperation): void {
  const owner = busy.command ?? vscode.l10n.t("another xlflow operation");
  void vscode.window.showWarningMessage(
    vscode.l10n.t(
      "{command} is currently using this workbook. {requested} was not started; retry after it completes or use an explicit wait action.",
      { command: owner, requested: operation.commandID },
    ),
  );
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === "string" && value.trim().length > 0;
}
