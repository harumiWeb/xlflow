# xlflow generate

Generate project artifacts such as test modules.

## xlflow generate test

Create a new VBA test module under the configured module source directory.

### Usage

```bash
xlflow generate test <module-name> [--json]
```

### Arguments

| Argument      | Description                                     |
| ------------- | ----------------------------------------------- |
| `module-name` | Name of the test module (without `.bas` suffix) |

### Description

`generate test` scaffolds a new standard module with:

- `Attribute VB_Name` and `Option Explicit`
- Lifecycle hook stubs: `BeforeAll`, `AfterAll`, `BeforeEach`, `AfterEach`
- A sample test sub to copy and adapt

The file is written to the configured `[src].modules` directory. The command refuses to overwrite an existing file.

### Example

```bash
xlflow generate test OrderServiceTests
```

This creates `src/modules/OrderServiceTests.bas` with the following content:

```vb
Attribute VB_Name = "OrderServiceTests"
Option Explicit

Public Sub BeforeAll()
End Sub

Public Sub AfterAll()
End Sub

Public Sub BeforeEach()
End Sub

Public Sub AfterEach()
End Sub

Public Sub Test_Sample()
    XlflowAssert.AssertTrue True, "replace with real assertions"
End Sub
```

## Related

- [test](./test)
- [module install](./module)

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow generate` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.

## Prerequisites

Check the project configuration and run `xlflow doctor --json` before workbook-backed operations. Source-only commands can run without Excel; commands that read or mutate a workbook require Windows Excel and VBIDE access.

## What this command reads and changes

The command reads the inputs and configuration described in its syntax and examples. Treat source files, the saved workbook, and a live session as separate states; add `--session` when the live workbook is authoritative. Any mutation is reversible only when a backup or explicit session save boundary exists.

## Effect on source-of-truth state

Use `xlflow status --json` before and after the command. A source edit normally requires `push`; a workbook edit normally requires `pull`; a dirty live session requires `save --session` or an intentional discard.

## Common workflows

Combine this command with the relevant [source/workbook/session workflow](../concepts/workbook-session-source), and use `--json` in scripts and agent loops.

## Common failures

Read the structured `error.code`, exit code, and recovery metadata instead of scraping terminal text. The [symptom-oriented troubleshooting guide](../help/troubleshooting) maps installation, execution, session, VS Code, and WSL failures to recovery steps.
