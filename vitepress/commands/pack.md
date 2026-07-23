# xlflow pack

Build a release `.xlsm` artifact from source and a workbook template.

```bash
xlflow pack --out build/Release.xlsm --experimental
```

`pack` is intended for controlled release automation. Run source checks first and open the resulting artifact in real Excel to compile/run a sentinel macro before publishing. See the repository's [pack specification](https://github.com/harumiWeb/xlflow/blob/main/docs/specs/pack-command.md) for release-gate details.

## Common failures

Unsupported extensions, missing templates, source preflight failures, or an existing output file return structured errors. Keep the source and template under version control and never treat a generated artifact as the source of truth.

<!-- xlflow-command-guidance -->

## When to use this command

Use `xlflow pack` when the task matches the command description above. For a goal-oriented workflow, start with the [How-to guides](../guides/) and return here for exact options.

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
