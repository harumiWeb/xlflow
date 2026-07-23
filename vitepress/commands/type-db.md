# xlflow type db

Manage generated TypeLib databases used by the VBA LSP.

## Usage

```bash
xlflow type db status
xlflow type db init
xlflow type db refresh
xlflow type db refresh --library all
xlflow type db clean
```

## Options

| Option      | Description                                                                                                      | Default             |
| ----------- | ---------------------------------------------------------------------------------------------------------------- | ------------------- |
| `--dir`     | Override the generated type DB directory.                                                                        | `~/.xlflow/typelib` |
| `--library` | TypeLib library to import. Repeat, comma-separate, or use `all` for every known library present on this machine. | `excel`             |
| `--force`   | Deprecated compatibility flag; `refresh` always regenerates.                                                     | `false`             |

`status` reports manifest presence, generated files, library LIBID/version metadata, stale state, and the LSP database search order.

`init` generates the database only when it does not already exist. `refresh` always regenerates the generated database, so it is the one-command equivalent of `clean` followed by `init`. `clean` deletes the generated database directory.

The default importer target is Excel. Use `--library all` to generate databases for every known library present on the machine, including Office, MSForms, Scripting, ADODB, and VBIDE when available. Generated entries include TypeLib-derived ProgID mappings when registry metadata can be matched to TypeLib CoClass GUIDs, so the LSP can infer common late-bound expressions such as `CreateObject("Excel.Application")`, `CreateObject("Scripting.FileSystemObject")`, and `CreateObject("ADODB.Connection")`.

Generated entries are loaded by `xlflow lsp` when present; if no generated DB exists, the LSP continues with the embedded built-in database.

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow type db` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.

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
