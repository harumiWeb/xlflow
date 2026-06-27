# Contributing to xlflow VS Code Extension

Thank you for your interest in contributing to the xlflow VS Code extension.

This extension is the official VS Code frontend for [xlflow](https://github.com/harumiWeb/xlflow). Its primary role is to provide a convenient GUI wrapper around the xlflow CLI and LSP.

## Core Principle

The VS Code extension should remain a thin UI layer over xlflow.

The source of truth for project behavior must live in the xlflow core, CLI, or LSP, not in the extension.

In general:

- Use xlflow CLI commands for project operations.
- Use xlflow JSON output for structured data.
- Use the xlflow LSP for language intelligence.
- Avoid duplicating xlflow logic in TypeScript.
- If the CLI or LSP does not expose the required capability, prefer adding or improving an xlflow command/API instead of implementing custom logic in the extension.

## What Belongs in the Extension

The extension may implement:

- VS Code views and TreeViews
- Command registration
- Status bar items
- Context menus
- Test Explorer integration
- LSP client wiring
- User prompts and confirmation dialogs
- Output channels and progress reporting
- Configuration handling
- Localization
- Mapping xlflow JSON output to VS Code UI objects

## What Should Not Belong in the Extension

The extension should not implement its own version of xlflow core behavior.

Avoid adding:

- VBA parsing logic
- VBA symbol analysis
- Call graph analysis
- Test discovery logic independent from `xlflow test list`
- Project structure inference independent from `xlflow.toml` and xlflow CLI output
- Workbook operation logic
- Module import/export logic
- xlflow-specific business rules that should be enforced by the CLI
- Alternative implementations of lint, format, analyze, inspect, push, pull, run, or test behavior

If a feature requires xlflow-specific knowledge, it probably belongs in the xlflow core first.

## Development Setup

Install dependencies:

```bash
pnpm install
```

Compile the extension:

```bash
pnpm compile
```

Run tests:

```bash
pnpm test
```

Package a local VSIX:

```bash
pnpm vsce package
```

The extension expects the `xlflow` CLI to be available either from `PATH` or from the configured executable path.

## Working with xlflow CLI

Prefer structured CLI output whenever possible.

Good:

```bash
xlflow inspect symbols --json
xlflow test list --json
xlflow test --json
```

Avoid parsing human-readable output unless there is no alternative.

When new structured data is needed, add or improve a JSON-capable command in xlflow rather than scraping text output in the extension.

## UX Guidelines

The extension should make xlflow easier to use without hiding important behavior.

- Destructive actions must ask for confirmation.
- Errors should be actionable.
- Missing `xlflow` CLI should guide the user to installation or configuration.
- Long-running commands should use progress notifications or output channels.
- Commands should preserve xlflow's CLI semantics as much as possible.
- Do not silently ignore failed xlflow commands.

## Cross-Platform Notes

xlflow supports workflows across Windows, WSL, and other environments, but some operations may require Windows and Excel.

The extension should avoid hard-coding Windows-only assumptions unless the underlying xlflow command itself requires Windows.

Use VS Code APIs for paths, URIs, workspace folders, and environment handling where possible.

## Localization

User-facing strings should be localizable.

Prefer `vscode.l10n.t(...)` for messages shown in the UI.

## Pull Request Guidelines

Before opening a pull request, please make sure:

- The extension compiles.
- Tests pass.
- New UI strings are localizable.
- New xlflow behavior is implemented in xlflow core/CLI/LSP when appropriate.
- The extension does not duplicate xlflow core logic.
- Destructive operations have confirmation prompts.
- Error messages are clear and actionable.

## Design Rule of Thumb

When in doubt, ask:

> Is this VS Code UI glue, or is this xlflow behavior?

If it is VS Code UI glue, it belongs in this extension.

If it is xlflow behavior, it should belong in xlflow core, CLI, or LSP.
