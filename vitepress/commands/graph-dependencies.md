# xlflow graph dependencies

Inspect confirmed project-local dependencies in exported VBA source.

## Usage

```bash
xlflow graph dependencies
xlflow graph dependencies --module Billing
xlflow graph dependencies --path src/modules --json
```

## Options and arguments

| Option                 | Description                                                                     | Default                 |
| ---------------------- | ------------------------------------------------------------------------------- | ----------------------- |
| `--module <name>`      | Case-insensitive exact source-module filter; retains its outbound dependencies. | all modules             |
| `--path <dir-or-file>` | Restrict the configured source scan.                                            | configured source roots |
| `--json`               | Return nodes, typed edges, evidence, and uncertainty in the xlflow envelope.    | false                   |

## JSON output

The top-level `graph` payload has `target: "dependencies"`, `nodes`, `edges`, and `uncertain_edges`. Confirmed edge kinds are `calls`, `uses_type`, `constructs`, and `implements`. Module call edges aggregate all contributing procedure call locations under `evidence`.

## Related

- [impact](./impact)
- [inspect](./inspect)

<!-- xlflow-command-guidance -->

## When to use this command

Use `graph dependencies` to understand static project-local coupling before a refactor, architecture review, or targeted testing pass.

## Prerequisites

Run it from an xlflow project with exported VBA source. Excel and a workbook session are not required.

## What this command reads and changes

The command reads configured VBA source files, parses them once, and does not modify source files, workbooks, or session state.

## Effect on source-of-truth state

This is read-only. Use returned evidence locations to decide which source files to inspect or change, then follow the normal source-to-workbook workflow if needed.

## Common workflows

Start with `xlflow graph dependencies --json`, filter a high-interest module with `--module`, then use `xlflow impact Module.Procedure` for a focused caller/callee view.

## Common failures

Calls and project type references that cannot resolve to exactly one visible declaration remain uncertain; do not treat them as confirmed dependencies.
